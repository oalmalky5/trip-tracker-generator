from __future__ import annotations

from datetime import date
from pathlib import Path
import time

import pandas as pd
import streamlit as st

from engine import TripConfig, load_accounts_contacts, build_meetings_df, export_excel


st.set_page_config(page_title="Trip Tracker Generator", layout="wide")

st.title("Trip Tracker Generator")
st.caption("Push-button tracker creation from CRM exports (Accounts + Contacts).")


# ----------------------------
# Defaults (bundled sample data)
# ----------------------------
DEFAULT_ACCOUNTS = Path("sample_data/Case Study - Customer Account Data Set.xlsx")
DEFAULT_CONTACTS = Path("sample_data/Case Study - Associated Contacts Data Set.xlsx")
DEFAULT_TEMPLATE = Path("sample_data/Saudi Trip Tracker- Case Study (1).xlsx")


# ----------------------------
# Validation rules (minimal + safe)
# Keep these minimal so the tool stays flexible.
# ----------------------------
REQUIRED_ACCOUNTS_COLS = ["Companies"]  # absolute minimum to build a tracker
OPTIONAL_ACCOUNTS_COLS = ["Company ID", "Website", "Primary Contact", "Primary Contact Email", "HQ City"]

# Contacts are optional, but if provided we validate minimally:
REQUIRED_CONTACTS_COLS = ["People"]  # email is nice-to-have; we still work without it
OPTIONAL_CONTACTS_COLS = ["Email", "Primary Company", "Primary Company Website"]

# Template validation: we just require it's a readable .xlsx (since engine writes by column order)
# If you want stricter template checks, you can validate header row contents too.


# ----------------------------
# Helpers
# ----------------------------
def _status_line(ok: bool, msg: str) -> str:
    return f"{'✓' if ok else '⚠️'} {msg}"


def _save_uploaded(upload, default_path: Path) -> Path:
    """
    Save uploaded file into .cache so engine can read from a Path.
    If upload is None, return the default sample path.
    """
    if upload is None:
        return default_path

    out_dir = Path(".cache")
    out_dir.mkdir(exist_ok=True)

    p = out_dir / upload.name
    p.write_bytes(upload.getvalue())
    return p


def _human_file_error(label: str) -> None:
    st.error(
        f"{label}: I couldn’t read this as an Excel (.xlsx) file.\n\n"
        "Fix:\n"
        "- Upload the original CRM export as .xlsx\n"
        "- Don’t upload CSV, PDF, or a random spreadsheet that isn’t the export\n"
        "- If the file is open in Excel, close it and try again"
    )
    st.stop()


def _validate_required_columns(df: pd.DataFrame, required: list[str], label: str) -> None:
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(
            f"{label}: This file doesn’t match the expected export format.\n\n"
            f"Missing required columns: {', '.join(missing)}\n\n"
            "Fix:\n"
            "- Re-export from the CRM using the standard export\n"
            "- Or upload the provided sample data to see the expected structure"
        )
        st.stop()


def _read_excel_for_validation(path: Path, label: str) -> pd.DataFrame:
    """
    Lightweight read for validation. We still call engine.load_accounts_contacts later,
    but this lets us catch and explain errors before the engine runs.
    """
    try:
        return pd.read_excel(path)
    except Exception:
        _human_file_error(label)


def _safe_contacts_subset(accounts: pd.DataFrame, contacts: pd.DataFrame, meetings_df: pd.DataFrame) -> pd.DataFrame:
    """
    Build the Contacts Directory subset without assuming optional columns exist.
    If we can't safely subset, return full contacts as a fallback.
    """
    if contacts is None or contacts.empty:
        return pd.DataFrame()

    # If we don’t have company matching fields, just return all contacts
    if "Primary Company" not in contacts.columns and "Primary Company Website" not in contacts.columns:
        return contacts.reset_index(drop=True)

    picked_names = set(meetings_df["Customer Account Name"].fillna("").astype(str).tolist())

    picked_websites = set()
    if "Companies" in accounts.columns and "Website" in accounts.columns:
        try:
            picked_websites = set(
                accounts[accounts["Companies"].fillna("").astype(str).isin(picked_names)]["Website"]
                .fillna("")
                .astype(str)
                .str.lower()
                .tolist()
            )
        except Exception:
            picked_websites = set()

    c = contacts.copy()

    mask = pd.Series([False] * len(c))
    if "Primary Company" in c.columns:
        mask = mask | c["Primary Company"].fillna("").astype(str).isin(picked_names)
    if "Primary Company Website" in c.columns and picked_websites:
        mask = mask | c["Primary Company Website"].fillna("").astype(str).str.lower().isin(picked_websites)

    subset = c[mask].reset_index(drop=True)
    return subset


def _safe_output_filename(trip_name: str) -> str:
    safe = "".join(ch if ch.isalnum() or ch in (" ", "-", "_") else "_" for ch in (trip_name or "Trip Tracker"))
    safe = safe.strip() or "Trip Tracker"
    return f"{safe}.xlsx"


# ----------------------------
# Sidebar inputs
# ----------------------------
with st.sidebar:
    st.header("Inputs")
    st.write("Upload exports or use the bundled sample files.")

    acc_file = st.file_uploader("Accounts export (.xlsx)", type=["xlsx"], key="acc")
    con_file = st.file_uploader("Contacts export (.xlsx)", type=["xlsx"], key="con")
    template_file = st.file_uploader("Template tracker (.xlsx)", type=["xlsx"], key="tmpl")

    st.divider()
    st.header("Trip settings")

    trip_name = st.text_input("Trip name", value="Riyadh Feb 2026 - Trip Tracker")
    city = st.text_input("City", value="Riyadh")
    start = st.date_input("Start date", value=date(2026, 2, 24))
    end = st.date_input("End date", value=date(2026, 2, 26))
    meetings = st.slider("Number of meetings", min_value=1, max_value=40, value=12, step=1)
    owners_raw = st.text_input("Owners (comma-separated)", value="Jason, Meshari")
    seed = st.number_input("Random seed", min_value=1, max_value=10_000, value=42, step=1)

    st.divider()
    generate = st.button("Generate tracker", type="primary", use_container_width=True)


# ----------------------------
# Layout
# ----------------------------
col1, col2 = st.columns([1.2, 1])

with col1:
    st.subheader("Progress")
    status_box = st.empty()
    log_box = st.empty()

with col2:
    st.subheader("Output")
    out_box = st.empty()


# ----------------------------
# Main action
# ----------------------------
if generate:
    owners = tuple([o.strip() for o in owners_raw.split(",") if o.strip()]) or ("Owner",)

    if end < start:
        st.error("End date must be the same as or after the start date.")
        st.stop()

    cfg = TripConfig(
        trip_name=(trip_name or "").strip() or "Trip Tracker",
        start_date=start,
        end_date=end,
        meetings=int(meetings),
        city=(city or "").strip() or "Riyadh",
        owners=owners,
        seed=int(seed),
    )

    accounts_path = _save_uploaded(acc_file, DEFAULT_ACCOUNTS)
    contacts_path = _save_uploaded(con_file, DEFAULT_CONTACTS) if con_file else None
    template_path = _save_uploaded(template_file, DEFAULT_TEMPLATE)

    # Pre-validate files with friendly errors (prevents scary stack traces)
    # Accounts
    accounts_preview = _read_excel_for_validation(accounts_path, "Accounts export")
    _validate_required_columns(accounts_preview, REQUIRED_ACCOUNTS_COLS, "Accounts export")

    # Contacts (optional)
    if contacts_path is not None:
        contacts_preview = _read_excel_for_validation(contacts_path, "Contacts export")
        _validate_required_columns(contacts_preview, REQUIRED_CONTACTS_COLS, "Contacts export")

    # Template (just make sure it is readable xlsx)
    _ = _read_excel_for_validation(template_path, "Template tracker")

    run_log: list[str] = []
    t0 = time.time()

    with status_box.status("Loading data...", expanded=True) as s:
        try:
            accounts, contacts = load_accounts_contacts(accounts_path, contacts_path)
            run_log.append(_status_line(True, f"{len(accounts)} accounts loaded"))

            if contacts is None or contacts.empty:
                run_log.append(_status_line(False, "No contacts provided (contacts directory will be blank)"))
            else:
                run_log.append(_status_line(True, f"{len(contacts)} contacts loaded"))

            s.update(label="Loading data... done", state="complete")
        except Exception:
            run_log.append(_status_line(False, "Failed to load files (format issue)."))
            s.update(label="Loading data... failed", state="error")
            log_box.code("\n".join(run_log))
            st.error(
                "I couldn’t load one of the uploaded files.\n\n"
                "Make sure you uploaded the Accounts/Contacts exports and the template as .xlsx files."
            )
            st.stop()

    with status_box.status("Building meetings list...", expanded=True) as s:
        try:
            meetings_df, issues, stats = build_meetings_df(accounts, contacts, cfg)
            st.session_state["meetings_df"] = meetings_df
            run_log.append(_status_line(True, f"{len(meetings_df)} meetings generated"))
            s.update(label="Building meetings list... done", state="complete")
        except ValueError as e:
            run_log.append(_status_line(False, str(e)))
            s.update(label="Building meetings list... failed", state="error")
            log_box.code("\n".join(run_log))
            st.error(str(e))
            st.stop()
        except Exception:
            run_log.append(_status_line(False, "Unexpected error while building meetings list"))
            s.update(label="Building meetings list... failed", state="error")
            log_box.code("\n".join(run_log))
            st.error(
                "Something unexpected happened while generating meetings.\n\n"
                "Try uploading the standard exports (or use the sample files)."
            )
            st.stop()

    with status_box.status("Preparing contacts directory...", expanded=True) as s:
        contacts_subset = _safe_contacts_subset(accounts, contacts, meetings_df)
        if contacts_subset.empty:
            run_log.append(_status_line(False, "Contacts directory: no matching contacts (or none provided)"))
        else:
            run_log.append(_status_line(True, f"Contacts directory: {len(contacts_subset)} contacts included"))
        s.update(label="Preparing contacts directory... done", state="complete")

    with status_box.status("Validating schedule...", expanded=True) as s:
        conflicts = int(meetings_df.duplicated(subset=["East40 Meeting Owner", "Meeting Date", "Meeting Time"]).sum())
        run_log.append(_status_line(True, f"Trip dates: {cfg.start_date.isoformat()} to {cfg.end_date.isoformat()}"))
        if conflicts == 0:
            run_log.append(_status_line(True, "No owner/time conflicts detected"))
        else:
            run_log.append(_status_line(False, f"{conflicts} owner/time conflicts detected (review recommended)"))
        s.update(label="Validating schedule... done", state="complete")

    with status_box.status("Creating tracker file...", expanded=True) as s:
        try:
            out_dir = Path(".outputs")
            out_dir.mkdir(exist_ok=True)

            out_path = out_dir / _safe_output_filename(cfg.trip_name)

            export_excel(
                template_path=template_path,
                output_path=out_path,
                cfg=cfg,
                meetings_df=meetings_df,
                contacts_df=contacts_subset,
                issues=issues,
                stats=stats,
                run_log_lines=run_log,
            )

            run_log.append(_status_line(True, f"Excel created: {out_path.name}"))
            s.update(label="Creating tracker file... done", state="complete")

        except Exception:
            run_log.append(_status_line(False, "Failed to generate the Excel file"))
            s.update(label="Creating tracker file... failed", state="error")
            log_box.code("\n".join(run_log))
            st.error(
                "I couldn’t generate the tracker using the uploaded template.\n\n"
                "Fix:\n"
                "- Upload the provided tracker template (.xlsx)\n"
                "- Or re-download the template and try again"
            )
            st.stop()

    elapsed = time.time() - t0
    run_log.append("")
    run_log.append(f"Done in {elapsed:.1f} seconds ✨")

    log_box.code("\n".join(run_log))

    with open(out_path, "rb") as f:
        out_box.download_button(
            label="Download tracker",
            data=f,
            file_name=out_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# ----------------------------
# Preview
# ----------------------------
st.subheader("Preview")

if "meetings_df" not in st.session_state:
    st.info("Generate a tracker to see a preview here.")
else:
    meetings_df = st.session_state["meetings_df"]

    # Streamlit's dataframe renderer uses PyArrow sometimes; keep fallback
    try:
        st.dataframe(meetings_df, use_container_width=True, height=420)
    except Exception as e:
        st.warning(f"Preview fallback (dataframe renderer unavailable): {e}")
        st.markdown(meetings_df.head(50).to_html(index=False), unsafe_allow_html=True)
