# Generic OEM Competitor Version

A Streamlit Python application for OEM-neutral competitor reporting across five markets.

## What this app does

- Provides a generic OEM competitor dashboard.
- Uses Product 1 numbering.
- Shows Start Here, Insight Report, Use Cases, Market Summary, Scorecard, and Data Dictionary pages.
- Allows users to select any OEM and any reporting period.
- Generates executive takeaways, market cards, rankings, and cross-market visualisations.
- Avoids locked OEMs, product-specific presets, and brand-specific benchmark logic.

## App structure

```text
generic_oem_competitor_app/
├── app.py
├── requirements.txt
├── README.md
├── .gitignore
└── data/
    └── sample_oem_data.csv
```

## Required data columns

Your CSV must contain these columns:

```text
market
period
oem
oem_grouping
visitors
contracts
website_to_contract_conversion_rate
```

Column names are normalised by the app, so spaces are converted to underscores and case is ignored.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate  # macOS/Linux
# .venv\Scripts\activate   # Windows
pip install -r requirements.txt
streamlit run app.py
```

## Upload to GitHub

```bash
git init
git add .
git commit -m "Initial Generic OEM competitor app"
git branch -M main
git remote add origin <your-repository-url>
git push -u origin main
```

## Deploy on Streamlit Community Cloud

1. Push this folder to GitHub.
2. Go to Streamlit Community Cloud.
3. Create a new app from the GitHub repo.
4. Set the main file path to `app.py`.
5. Deploy.

## Notes for production use

The bundled dataset is sample data only. Replace `data/sample_oem_data.csv` or upload your real CSV in the sidebar.

For production reporting, validate the exact metric definitions before releasing the app. A dashboard with unclear definitions is worse than no dashboard because it creates false confidence.
