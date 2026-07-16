# Finmodgen — Excel financial model generator from publicly available data


Requirements
- Python 3.8+
- See `requirements.txt` for Python packages.

Install

Run in PowerShell (from repo root):

```powershell
python -m pip install -r Finmodgen/requirements.txt
```

Usage

Run the script with a ticker symbol (example uses AAPL):

```powershell
python -m Finmodgen.main AAPL
```

You can pass options:

- --output, -o : output path. Use `{ticker}` in path to inject the ticker (default: `output/{ticker}_model.xlsx`).
- --history, -H : history period for prices (yfinance period string, default `5y`).
- --forecast-years, -f : forecast horizon in years (default 5).

Example with options:

```powershell
python -m Finmodgen.main AAPL -o "output/models/{ticker}_v1.xlsx" -H 1y -f 3
```

This creates `output/{TICKER}_model.xlsx` (or the path you supplied) with sheets: Overview, Income Statement, Balance Sheet, Cash Flow, Historical Prices, Ratios, Forecast.

Notes
- Uses `yfinance` public data. Data availability depends on Yahoo's API and the ticker.
- This is a minimal starter; feel free to extend the Forecast sheet and formatting.
