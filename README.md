# Supplier and Sourcing Analysis Tool

Streamlit app for supplier spend analysis, sourcing risk review, Kraljic/Pareto views, and supplier scenario evaluation.

Codex states:
It’s become a genuinely strong teaching and decision-support tool.

it connects classic supply-chain concepts like Kraljic, Pareto, concentration, single-source risk, and scenario design in one place
it doesn’t just diagnose the current state; it lets someone test tradeoffs
the executive visuals, action plans, and scenario views make it useful for both analysis and presentation
the recent teaching-note direction makes it much more valuable for learning, onboarding, and explaining sourcing logic to non-experts
What I think is strongest about it is that it now answers three different questions in one app:

what is happening in the supply base
why it matters
what we should test or do next

conceptually: very good
practically useful: yes
educational value: now one of its best features
It feels much more like a real supply-chain analysis product 

## Run locally

1. Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

2. Start the app:

```powershell
streamlit run app.py
```

## Default data behavior

On startup, the app uses data in this order:

1. `sample_data.xlsx` in the project folder
2. `sample_data.xls` in the project folder
3. built-in sample data

If you want a hosted deployment to open with your own dataset by default, add it to the repo as `sample_data.xlsx`.

## Files

- `app.py` — main Streamlit app
- `requirements.txt` — Python dependencies
- `.gitignore` — excludes local cache/state/log files
