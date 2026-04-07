# Supplier and Sourcing Analysis Tool

Streamlit app for supplier spend analysis, sourcing risk review, Kraljic and Pareto views, supplier scenario evaluation, and executive-ready exports.

## What the app does

- Analyzes supplier and component spend concentration
- Scores sourcing risk using supplier depth, concentration, quality, lead time, and criticality
- Builds Kraljic, Pareto, and strategic-priority views
- Lets users test supplier consolidation and mitigation scenarios
- Exports analysis outputs as a CSV bundle and PowerPoint deck

## Run locally

1. Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

2. Start the app:

```powershell
python -m streamlit run app.py
```

## Default data behavior

On startup, the app uses data in this order:

1. `sample_data.xlsx` in the project folder
2. `sample_data.xls` in the project folder
3. built-in sample data

If you want a hosted deployment to open with your own dataset by default, add it to the repo as `sample_data.xlsx`.

## Deploy to Streamlit Community Cloud

1. Push this project to GitHub.
2. Go to [share.streamlit.io](https://share.streamlit.io/).
3. Create a new app from `joshua910m/supplier-sourcing-analysis-tool`.
4. Set the main file path to `app.py`.
5. Deploy or reboot the app after dependency changes.

## Files

- `app.py` - main Streamlit app
- `requirements.txt` - Python dependencies for local and cloud deployment
- `.gitignore` - excludes local cache, state, and log files
