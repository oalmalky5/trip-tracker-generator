# Trip Tracker Generator (Streamlit)

A push-button tool that generates a trip coordination tracker from CRM exports (Accounts + Contacts).
Built to be **repeatable, automated, and maintainable**, with a clear seam to later swap the data source to **HubSpot**.

## Why this exists

We coordinate multi-day trips where portfolio companies meet Saudi stakeholders. External parties often do not have CRM access.
This tool produces a **shareable, editable tracker** (Excel) in one action.

## What it does (end-to-end)

1. User uploads **Accounts** + **Contacts** exports (or uses the included sample data)
2. User clicks **Generate Tracker**
3. App creates:

- a randomized list of meetings (within the trip date range)
- relevant account + contact context
- a filled Excel tracker based on the provided template

4. User downloads the output and shares it with the trip team

No manual copy/paste. No Excel formulas. All logic is implemented in Python.

## Output

Generates an Excel tracker with these tabs:

- **Trip Overview**: trip metadata, run log, quick stats
- **Meetings**: main tracker sheet (matches the provided template columns)
- **Contacts Directory**: contacts for the selected accounts
- **Summary**: basic breakdowns (status, owner, industry)
- **Data Issues**: flagged missing/invalid fields (blockers vs warnings) + suggested fixes

## Data assumptions + constraints

- **Accounts and Contacts** are exported with a stable header row.
- Accounts must have a unique identifier (e.g., Account Name or Account ID) that Contacts can reference.
- If multiple contacts exist for an account, the tool selects a “primary” contact using a deterministic rule
  (documented in code) and includes the rest in **Contacts Directory**.
- Meeting fields (date/time/address/status) may be **randomly generated** as permitted by the case study.
- Missing required fields are **not guessed**. They are left blank and flagged in **Data Issues**.

## Random seed (why it exists)

The tool uses randomness to generate meeting schedules. A **seed** makes that randomness reproducible:

- Same inputs + same seed = same generated schedule
- Useful for debugging, iteration, and reviewer consistency

From a user perspective: they normally ignore it. It’s there so the system can be tested and re-run reliably.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip setuptools wheel
pip install -r requirements.txt

streamlit run app.py
```
