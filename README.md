# Stock Prices

A small toolkit for looking up and tracking historical stock prices.

## Contents

| File | What it does |
|------|--------------|
| `index.html` | A standalone web page to look up a ticker's price history and download it as CSV. Open it in any browser. |
| `update_stocks.py` | Fetches daily prices from Yahoo Finance and writes/updates a local Excel workbook. |
| `stock_prices.xlsx` | The generated workbook (a `Summary` snapshot sheet + an append-only `Prices` log). |
| `.github/workflows/update-stocks.yml` | Scheduled GitHub Action that refreshes the workbook every weekday and commits the changes. |

## Updating the Excel sheet

```bash
pip install -r requirements.txt

python3 update_stocks.py                      # default tickers (AAPL MSFT GOOGL TSLA AMZN)
python3 update_stocks.py NVDA AMD INTC        # specific tickers
TICKERS="AAPL,TSLA" python3 update_stocks.py  # tickers via env var
python3 update_stocks.py --output prices.xlsx --range 1mo
```

Options:

- **Tickers** — positional args, or the `TICKERS` env var (comma/space separated), else the built-in default list.
- `--output` / `STOCK_XLSX` — workbook path (default `stock_prices.xlsx`).
- `--range` / `PRICE_RANGE` — how much history to fetch each run: `5d`, `1mo`, `6mo`, `1y`, `5y`, etc. (default `5d`).

The script is **idempotent**: rows are de-duplicated by Date + Ticker, so re-running
(or running on a schedule) never creates duplicate entries — it only appends genuinely
new trading days and rebuilds the `Summary` sheet.

## Automated (scheduled) updates

`.github/workflows/update-stocks.yml` runs the script at **22:30 UTC on weekdays**
(shortly after the US market close), then commits any change to `stock_prices.xlsx`
back to the repo. You can also trigger it manually from the **Actions** tab
("Run workflow"), optionally overriding the tickers and range.

The workflow needs no secrets — Yahoo Finance is public — but it does require the
repository's Actions setting **"Read and write permissions"** (Settings → Actions →
General → Workflow permissions) so it can push the updated workbook.
