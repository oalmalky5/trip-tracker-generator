from __future__ import annotations
from datetime import datetime

from datetime import date
from pathlib import Path
import time

import pandas as pd
import streamlit as st

from engine import TripConfig, load_accounts_contacts, build_meetings_df, export_excel


st.set_page_config(page_title="Trip Tracker Generator", layout="wide")

st.title("Trip Tracker Generator")
st.caption("Push-button tracker creation from CRM exports (Accounts + Contacts).")

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

# Default bundled paths (for local use). In your submission, you can keep these as sample data.
DEFAULT_ACCOUNTS = Path("sample_data/Case Study - Customer Account Data Set.xlsx")
DEFAULT_CONTACTS = Path("sample_data/Case Study - Associated Contacts Data Set.xlsx")
DEFAULT_TEMPLATE = Path("sample_data/Saudi Trip Tracker- Case Study (1).xlsx")

def _save_uploaded(upload, default_path: Path) -> Path:
    if upload is None:
        return default_path
    out = Path(".cache")
    out.mkdir(exist_ok=True)
    p = out / upload.name
    p.write_bytes(upload.getvalue())
    return p

def _status_line(ok: bool, msg: str) -> str:
    return f"{'✓' if ok else '⚠️'} {msg}"

col1, col2 = st.columns([1.2, 1])

with col1:
    st.subheader("Progress")
    status_box = st.empty()
    log_box = st.empty()

with col2:
    st.subheader("Output")
    out_box = st.empty()

if generate:
    owners = tuple([o.strip() for o in owners_raw.split(",") if o.strip()]) or ("Owner",)
    cfg = TripConfig(
        trip_name=trip_name.strip() or "Trip Tracker",
        start_date=start,
        end_date=end,
        meetings=int(meetings),
        city=city.strip() or "Riyadh",
        owners=owners,
        seed=int(seed),
    )

    accounts_path = _save_uploaded(acc_file, DEFAULT_ACCOUNTS)
    contacts_path = _save_uploaded(con_file, DEFAULT_CONTACTS)
    template_path = _save_uploaded(template_file, DEFAULT_TEMPLATE)

    run_log: list[str] = []
    t0 = time.time()

    with status_box.status("Loading data...", expanded=True) as s:
        try:
            accounts, contacts = load_accounts_contacts(accounts_path, contacts_path)
            run_log.append(_status_line(True, f"{len(accounts)} accounts loaded"))
            run_log.append(_status_line(True, f"{len(contacts)} contacts loaded"))
            s.update(label="Loading data... done", state="complete")
        except Exception as e:
            run_log.append(_status_line(False, f"Failed to load files: {e}"))
            s.update(label="Loading data... failed", state="error")
            log_box.code("\n".join(run_log))
            st.stop()

    with status_box.status("Selecting companies...", expanded=True) as s:
        meetings_df, issues, stats = build_meetings_df(accounts, contacts, cfg)
        st.session_state['meetings_df'] = meetings_df

        # Contacts directory limited to picked companies (match by website or name)
        picked_names = set(meetings_df["Customer Account Name"].fillna("").astype(str).tolist())
        picked_websites = set(
            accounts[accounts["Companies"].isin(picked_names)]["Website"].fillna("").astype(str).str.lower().tolist()
        )

        c = contacts.copy()
        mask = (
            c["Primary Company"].fillna("").astype(str).isin(picked_names)
            | c["Primary Company Website"].fillna("").astype(str).str.lower().isin(picked_websites)
        )
        contacts_subset = c[mask].reset_index(drop=True)

        run_log.append(_status_line(True, f"{len(meetings_df)} companies selected for meetings"))
        warn_emails = sum(1 for i in issues if i.field.lower().find("email") >= 0)
        if warn_emails:
            run_log.append(_status_line(False, f"{warn_emails} email-related issues flagged"))
        s.update(label="Selecting companies... done", state="complete")

    with status_box.status("Generating schedule...", expanded=True) as s:
        # Conflicts check: owner+date+time duplicates
        conflicts = int(meetings_df.duplicated(subset=["East40 Meeting Owner", "Meeting Date", "Meeting Time"]).sum())
        run_log.append(_status_line(True, f"{len(meetings_df)} meetings: {cfg.start_date.isoformat()} to {cfg.end_date.isoformat()}"))
        if conflicts == 0:
            run_log.append(_status_line(True, "No time conflicts detected"))
            s.update(label="Generating schedule... done", state="complete")
        else:
            run_log.append(_status_line(False, f"{conflicts} time conflicts detected (flagged)"))
            s.update(label="Generating schedule... done with warnings", state="complete")

    with status_box.status("Creating tracker...", expanded=True) as s:
        out_dir = Path(".outputs")
        out_dir.mkdir(exist_ok=True)
        safe_name = "".join(ch if ch.isalnum() or ch in (" ", "-", "_") else "_" for ch in cfg.trip_name).strip()
        out_path = out_dir / f"{safe_name}.xlsx"

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
        s.update(label="Creating tracker... done", state="complete")

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

st.subheader("Preview")

# Persist the last generated dataframe so the app can render on reruns
if "meetings_df" not in st.session_state:
    st.info("Generate a tracker to see a preview here.")
else:
    meetings_df = st.session_state["meetings_df"]

    # Streamlit's dataframe renderer uses PyArrow. On some macOS + Python combos,
    # PyArrow can fail to import if binary wheels mismatch. The tracker generation
    # still works, so we gracefully fall back to an HTML preview.
    try:
        st.dataframe(meetings_df, use_container_width=True, height=420)
    except Exception as e:
        st.warning(f"Preview fallback (dataframe renderer unavailable): {e}")
        st.markdown(meetings_df.head(50).to_html(index=False), unsafe_allow_html=True)