# Mr Vikram Holdings Dashboard

This dashboard was built from an NSDL DP holdings report for Mr Vikram, client ID `220424013`.

## What this file contains

- Current holdings and quantities
- Current market value
- Day movement and day P&L calculated from the NSDL LTP change
- Theme exposure
- Review priority and discussion queue
- Technical placeholders for RS, RSI, moving averages, and Point & Figure

The NSDL report does not contain purchase price or cost basis, so total portfolio P&L is intentionally not shown.

## Local refresh

From this folder, run:

```bash
python3 scripts/refresh_nsdl_prices.py --input data/holdings.csv --output data/holdings.csv
python3 scripts/refresh_technical_indicators.py --input data/holdings.csv --output data/holdings.csv
python3 scripts/build_nsdl_dashboard.py --input-csv data/holdings.csv --output-dir .
```

Open `index.html` after the build completes.

## Online refresh

The workflow in `.github/workflows/refresh-dashboard.yml` refreshes prices,
refreshes technicals, rebuilds the dashboard, and commits the updated static
page.

## Hosting

Live site: https://vikram-portfolio-dashboard.netlify.app/ (Netlify, auto-deploys
every push to `main`; see `netlify.toml`). The GitHub Pages URL
https://hvj7p7shkv-hub.github.io/vikram-portfolio-dashboard/ is kept as a
temporary fallback.
