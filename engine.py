from __future__ import annotations

def suggest_fix(issue: str, field: str = "", context: str = "") -> str:
    t = (issue or "").lower()
    f = (field or "").lower()
    if "email" in t or "email" in f:
        return "Add a valid email address for the primary contact in CRM (Contacts export)."
    if "address" in t or "address" in f or "location" in t:
        return "Add the meeting address / HQ address in CRM (Accounts export)."
    if "owner" in t or "owner" in f:
        return "Assign an account owner in CRM so meetings can be distributed."
    if "description" in t or "description" in f:
        return "Add a short company description in CRM (optional)."
    if "duplicate" in t:
        return "De-duplicate the record in CRM or confirm which record should be used."
    if "missing" in t:
        return "Fill the missing value in CRM and re-export."
    return "Review the CRM export and correct the value."

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Iterable, Optional, Tuple, Dict, List
import random
import re

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


TEMPLATE_HEADERS = [
    "Customer Account Name",
    "Meeting Date",
    "Meeting Time",
    "Meeting City",
    "Meeting Address",
    "East40 Meeting Owner",
    "Primary Contact Name",
    "Primary Contact Email",
    "Meeting Status",
    "Company Description",
]

STATUS_OPTIONS = ["Proposed", "Tentative", "Confirmed", "Rescheduled", "Cancelled", "Done"]


@dataclass(frozen=True)
class TripConfig:
    trip_name: str
    start_date: date
    end_date: date
    meetings: int = 12
    city: str = "Riyadh"
    owners: Tuple[str, ...] = ("Jason", "Meshari")
    seed: int = 42


@dataclass
class Issue:
    severity: str   # "BLOCKER" or "WARNING"
    entity: str     # "account" / "contact" / "meeting"
    entity_id: str  # e.g. Company ID or Person ID
    field: str
    message: str


def _safe_str(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def load_accounts_contacts(accounts_path: Path, contacts_path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    accounts = pd.read_excel(accounts_path)
    contacts = pd.read_excel(contacts_path)
    return accounts, contacts


def pick_accounts(accounts: pd.DataFrame, n: int, seed: int, city: Optional[str] = None) -> pd.DataFrame:
    rnd = random.Random(seed)

    df = accounts.copy()

    if city:
        # Prefer HQ City match if available
        hq_city = df.get("HQ City")
        if hq_city is not None:
            df_city = df[df["HQ City"].fillna("").astype(str).str.lower().str.contains(city.lower())]
            if len(df_city) >= n:
                df = df_city

    if len(df) <= n:
        return df.sample(n=min(n, len(df)), random_state=seed)

    idx = rnd.sample(list(df.index), k=n)
    return df.loc[idx].reset_index(drop=True)


def select_primary_contact_for_account(
    account_row: pd.Series,
    contacts: pd.DataFrame,
    issues: list[Issue],
) -> tuple[str, str]:
    """Return (name, email). Uses account primary contact if present, else picks from contacts export.

    Matching strategy:
    1) Use account's Primary Contact Email if present
    2) Match contacts by Primary Company Website == account Website (case-insensitive)
    3) Match contacts by Primary Company == account Companies (case-insensitive)
    Choose first contact with email; prefer exact name match to account Primary Contact if present.
    """
    company_id = _safe_str(account_row.get("Company ID", ""))
    acct_primary_name = _safe_str(account_row.get("Primary Contact", ""))
    acct_primary_email = _safe_str(account_row.get("Primary Contact Email", ""))
    acct_company = _safe_str(account_row.get("Companies", ""))
    acct_website = _safe_str(account_row.get("Website", ""))

    if acct_primary_email:
        return (acct_primary_name or "", acct_primary_email)

    # Build candidate set
    cands = contacts.copy()

    if acct_website:
        cands = cands[cands["Primary Company Website"].fillna("").astype(str).str.lower() == acct_website.lower()]
    if cands.empty and acct_company:
        cands = contacts[contacts["Primary Company"].fillna("").astype(str).str.lower() == acct_company.lower()]

    if cands.empty:
        issues.append(Issue(
            severity="WARNING",
            entity="account",
            entity_id=company_id or acct_company or "(unknown)",
            field="Primary Contact Email",
            message="No matching contacts found; primary contact left blank."
        ))
        return ("", "")

    # Prefer contacts with email
    cands = cands.copy()
    cands["_email"] = cands["Email"].apply(_safe_str)
    cands = cands[cands["_email"].astype(bool)]
    if cands.empty:
        issues.append(Issue(
            severity="WARNING",
            entity="account",
            entity_id=company_id or acct_company or "(unknown)",
            field="Primary Contact Email",
            message="Matching contacts exist but none have an email; primary contact left blank."
        ))
        return ("", "")

    if acct_primary_name:
        norm_target = _norm(acct_primary_name)
        cands["_name"] = cands["People"].apply(_safe_str)
        exact = cands[cands["_name"].apply(_norm) == norm_target]
        if not exact.empty:
            row = exact.iloc[0]
            return (_safe_str(row.get("People", "")), _safe_str(row.get("Email", "")))

    row = cands.iloc[0]
    return (_safe_str(row.get("People", "")), _safe_str(row.get("Email", "")))


def generate_schedule(cfg: TripConfig, n: int) -> list[tuple[date, str]]:
    """Returns list of (meeting_date, meeting_time). Distributes across days, avoids same-day time conflicts."""
    rnd = random.Random(cfg.seed)

    days = (cfg.end_date - cfg.start_date).days + 1
    if days <= 0:
        raise ValueError("end_date must be on or after start_date")

    # Simple day distribution
    per_day = [0] * days
    for i in range(n):
        per_day[i % days] += 1

    slots = []
    for d, count in enumerate(per_day):
        day = cfg.start_date + timedelta(days=d)
        # pick times in 30-min increments, business hours
        possible = [f"{h:02d}:{m:02d}" for h in range(9, 18) for m in (0, 30)]
        rnd.shuffle(possible)
        chosen = sorted(possible[:count])
        for t in chosen:
            slots.append((day, t))
    return slots


def build_meetings_df(
    accounts: pd.DataFrame,
    contacts: pd.DataFrame,
    cfg: TripConfig,
) -> tuple[pd.DataFrame, list[Issue], dict]:
    issues: list[Issue] = []

    picked = pick_accounts(accounts, cfg.meetings, cfg.seed, city=cfg.city)
    schedule = generate_schedule(cfg, len(picked))

    rows = []
    owners = list(cfg.owners) if cfg.owners else ["Owner"]
    rnd = random.Random(cfg.seed)

    for i, acct in picked.iterrows():
        meeting_date, meeting_time = schedule[i]
        owner = owners[i % len(owners)]
        status = rnd.choice(STATUS_OPTIONS)

        acct_name = _safe_str(acct.get("Companies", ""))
        desc = _safe_str(acct.get("Description", ""))
        addr1 = _safe_str(acct.get("HQ Address Line 1", ""))
        addr2 = _safe_str(acct.get("HQ Address Line 2", ""))
        city = _safe_str(acct.get("HQ City", "")) or cfg.city
        address = ", ".join([p for p in [addr1, addr2, city] if p])

        contact_name, contact_email = select_primary_contact_for_account(acct, contacts, issues)

        if not acct_name:
            issues.append(Issue("BLOCKER", "account", _safe_str(acct.get("Company ID", "(unknown)")), "Companies", "Missing account name."))

        if not address:
            issues.append(Issue("WARNING", "account", _safe_str(acct.get("Company ID", "(unknown)")), "HQ Address", "Missing HQ address; meeting address left blank."))

        if not contact_email:
            # treat as warning; downstream can fill in
            issues.append(Issue("WARNING", "contact", _safe_str(acct.get("Company ID", "(unknown)")), "Email", "Primary contact email missing."))

        rows.append({
            "Customer Account Name": acct_name,
            "Meeting Date": meeting_date.strftime("%b %d, %Y"),
            "Meeting Time": meeting_time,
            "Meeting City": cfg.city,
            "Meeting Address": address,
            "East40 Meeting Owner": owner,
            "Primary Contact Name": contact_name,
            "Primary Contact Email": contact_email,
            "Meeting Status": status,
            "Company Description": desc,
        })

    meetings_df = pd.DataFrame(rows, columns=TEMPLATE_HEADERS)

    # Stats for overview
    industries = picked.get("Primary Industry Group")
    if industries is None:
        ind_counts = {}
    else:
        ind_counts = industries.fillna("(blank)").astype(str).value_counts().to_dict()

    stats = {
        "accounts_loaded": int(len(accounts)),
        "contacts_loaded": int(len(contacts)),
        "meetings": int(len(meetings_df)),
        "days": int((cfg.end_date - cfg.start_date).days + 1),
        "owner_counts": meetings_df["East40 Meeting Owner"].value_counts().to_dict(),
        "status_counts": meetings_df["Meeting Status"].value_counts().to_dict(),
        "industry_counts": ind_counts,
    }
    return meetings_df, issues, stats


def _auto_fit_columns(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is None:
                continue
            max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), 60)


def export_excel(
    template_path: Path,
    output_path: Path,
    cfg: TripConfig,
    meetings_df: pd.DataFrame,
    contacts_df: pd.DataFrame,
    issues: list[Issue],
    stats: dict,
    run_log_lines: list[str],
) -> Path:
    wb = load_workbook(template_path)

    # Rename main sheet to Meetings
    main = wb[wb.sheetnames[0]]
    main.title = "Meetings"

    # Clear existing rows except header
    if main.max_row > 1:
        main.delete_rows(2, main.max_row - 1)

    # Write meetings
    for r_idx, row in enumerate(meetings_df.itertuples(index=False), start=2):
        for c_idx, value in enumerate(row, start=1):
            main.cell(row=r_idx, column=c_idx, value=value)

    main.freeze_panes = "A2"
    main.auto_filter.ref = f"A1:{get_column_letter(len(TEMPLATE_HEADERS))}{main.max_row}"

    # Trip Overview
    if "Trip Overview" in wb.sheetnames:
        del wb["Trip Overview"]
    ov = wb.create_sheet("Trip Overview", 0)
    ov["A1"] = "Trip"
    ov["B1"] = cfg.trip_name
    ov["A2"] = "Dates"
    ov["B2"] = f"{cfg.start_date.isoformat()} to {cfg.end_date.isoformat()}"
    ov["A3"] = "City"
    ov["B3"] = cfg.city
    ov["A4"] = "Meetings"
    ov["B4"] = stats.get("meetings", 0)
    ov["A6"] = "Run Log"
    ov["A6"].font = Font(bold=True)
    for i, line in enumerate(run_log_lines, start=7):
        ov[f"A{i}"] = line

    # Summary tab
    if "Summary" in wb.sheetnames:
        del wb["Summary"]
    sm = wb.create_sheet("Summary", 1)
    sm["A1"] = "Summary"
    sm["A1"].font = Font(bold=True)

    def _write_counts(start_row: int, title: str, counts: Dict[str, int]) -> int:
        sm[f"A{start_row}"] = title
        sm[f"A{start_row}"].font = Font(bold=True)
        r = start_row + 1
        sm[f"A{r}"] = "Category"
        sm[f"B{r}"] = "Count"
        sm[f"A{r}"].font = Font(bold=True)
        sm[f"B{r}"].font = Font(bold=True)
        r += 1
        for k, v in counts.items():
            sm[f"A{r}"] = k
            sm[f"B{r}"] = int(v)
            r += 1
        return r + 1

    r = 3
    r = _write_counts(r, "Meetings by Status", stats.get("status_counts", {}))
    r = _write_counts(r, "Meetings by Owner", stats.get("owner_counts", {}))
    r = _write_counts(r, "Accounts by Industry Group", stats.get("industry_counts", {}))

    # Contacts Directory tab (for selected accounts)
    if "Contacts Directory" in wb.sheetnames:
        del wb["Contacts Directory"]
    cd = wb.create_sheet("Contacts Directory")
    contact_cols = [c for c in ["People", "Email", "Phone", "Primary Position", "Primary Company", "Primary Company Type", "City", "Country/Territory/Region", "LinkedIn URL"] if c in contacts_df.columns]
    for c_idx, col in enumerate(contact_cols, start=1):
        cd.cell(row=1, column=c_idx, value=col).font = Font(bold=True)
    for r_idx, row in enumerate(contacts_df[contact_cols].fillna("").itertuples(index=False), start=2):
        for c_idx, v in enumerate(row, start=1):
            cd.cell(row=r_idx, column=c_idx, value=v)
    cd.freeze_panes = "A2"
    if contact_cols:
        cd.auto_filter.ref = f"A1:{get_column_letter(len(contact_cols))}{cd.max_row}"

    # Data Issues tab
    if "Data Issues" in wb.sheetnames:
        del wb["Data Issues"]
    di = wb.create_sheet("Data Issues")
    di_headers = ["Severity", "Entity", "Entity ID", "Field", "Message"]
    for c_idx, h in enumerate(di_headers, start=1):
        di.cell(row=1, column=c_idx, value=h).font = Font(bold=True)
    for r_idx, iss in enumerate(issues, start=2):
        di.cell(row=r_idx, column=1, value=iss.severity)
        di.cell(row=r_idx, column=2, value=iss.entity)
        di.cell(row=r_idx, column=3, value=iss.entity_id)
        di.cell(row=r_idx, column=4, value=iss.field)
        di.cell(row=r_idx, column=5, value=iss.message)
    di.freeze_panes = "A2"
    di.auto_filter.ref = f"A1:{get_column_letter(len(di_headers))}{di.max_row}"

    for ws in [ov, sm, cd, di, main]:
        _auto_fit_columns(ws)

    wb.save(output_path)
    return output_path
