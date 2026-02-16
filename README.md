# Trip Tracker Generator (Streamlit)

A push-button internal tool that generates a trip coordination tracker from CRM exports (Accounts + Contacts),
designed so the data source can later be swapped for HubSpot.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

streamlit run app.py
```

## What it generates
An Excel tracker with these tabs:
- **Trip Overview**: trip metadata + run log + quick stats
- **Meetings**: main tracker sheet (matches the provided template columns)
- **Contacts Directory**: contacts for the selected accounts
- **Summary**: simple breakdowns (status, owner, industry)
- **Data Issues**: flagged missing/invalid fields (blockers vs warnings)

## Notes
- This sample writes an Excel file for maximum portability.
- A Google Sheets exporter can be added as a drop-in `Exporter` implementation (adapter pattern).

## Troubleshooting (macOS)
If installation fails while building `pyarrow` and complains about missing `cmake`, do this:

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip setuptools wheel
pip install -r requirements.txt
```

This project pins `pyarrow` to a version that ships prebuilt wheels for most macOS + Python combinations.
If you still hit issues, try Python 3.11 (recommended) and repeat the steps above.

## If you see: `ImportError: numpy.core.multiarray failed to import`
This is a binary mismatch between `numpy` and `pyarrow` on some macOS + Python setups.

Fix by doing a clean reinstall inside the virtualenv:

```bash
deactivate  # if needed
rm -rf .venv
python -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip setuptools wheel
pip install -r requirements.txt
streamlit run app.py
```

This project pins `numpy==1.26.4` + `pyarrow==14.0.2` for compatibility with Python 3.10.

## Preview behavior
The preview area shows the most recent generated meeting table. Before the first generation,
it displays a friendly prompt instead of throwing an error.

## Recent improvements
- Clear data source indicator (sample vs uploaded files)
- Timestamped output filenames to avoid overwrites
- Data Issues sheet includes a "Suggested Fix" column for actionable remediation
- Random seed explained in the UI (reproducible generation)
