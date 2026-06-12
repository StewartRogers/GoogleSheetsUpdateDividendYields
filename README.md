# GoogleSheetsUpdateDividendYields

A Google Apps Script that fetches current **dividend yields**, **dividend payable dates**, and **share prices** for TSX-listed stocks from the TMX API and writes them into your Google Sheet.

---

## Spreadsheet Layout

The script expects your data to start on **row 2** (row 1 is the header). The relevant columns are:

| Column | Default | Contents |
|--------|---------|----------|
| D | 4 | Ticker symbol (e.g. `ENB.TO`, `GRT-UN.TO`, `CASH`) |
| E | 5 | Number of shares held |
| G | 7 | Share price — written by the script (selected tickers only) |
| N | 14 | Dividend payable date — written by the script |
| T | 20 | Dividend yield — written by the script |

> These columns are configurable at the top of the script (see [Configuration](#configuration) below).

---

## Installation

### 1. Open the Apps Script editor

1. Open your Google Sheet.
2. Click **Extensions** → **Apps Script**.

### 2. Paste the script

1. Delete any placeholder code in the editor.
2. Copy the full contents of [code.gs](code.gs) and paste it into the editor.
3. Click the **Save** icon (or press `Ctrl+S` / `Cmd+S`).

### 3. Configure the script

At the top of the script, update the configuration constants to match your spreadsheet:

```js
const DIVIDEND_YIELD_SHEET_NAME  = "Portfolio"; // Name of your sheet tab
const DIVIDEND_YIELD_TICKER_COL  = 4;           // Column containing ticker symbols (D = 4)
const DIVIDEND_YIELD_SHARES_COL  = 5;           // Column containing share counts (E = 5)
const DIVIDEND_YIELD_OUTPUT_COL  = 20;          // Column to write dividend yields (T = 20)
const DIVIDEND_YIELD_HEADER_ROWS = 1;           // Number of header rows to skip

const PAYABLE_DATE_COL           = 14;                           // Column to write dividend payable dates (N = 14)
const SHARE_PRICE_OUTPUT_COL     = 7;                            // Column to write share prices (G = 7)
const SHARE_PRICE_TARGET_TICKERS = ["GRT-UN.TO", "REI-UN.TO"];  // Only these tickers get a share price update
```

Set `DIVIDEND_YIELD_SHEET_NAME` to exactly match the tab name at the bottom of your spreadsheet. Add or remove entries from `SHARE_PRICE_TARGET_TICKERS` to control which stocks have their share price updated.

---

## Running the Script

Three functions are available in the Apps Script editor's function dropdown:

| Function | What it does |
|----------|-------------|
| `runUpdatePortfolioData` | Runs both updates in a single pass (recommended) |
| `runUpdateDividendYields` | Updates dividend yield and payable date columns only |
| `runUpdateSelectedSharePrices` | Updates share price column only |

1. Select the desired function from the dropdown and click the **Run** button (▶).
2. The first time you run it, Google will ask you to grant permissions:
   - **View and manage your spreadsheets** — to read tickers and write output values.
   - **Connect to an external service** — to fetch data from the TMX API.
   - Click **Review permissions** → choose your Google account → **Allow**.

---

## Ticker Format

Tickers should be entered as they appear on the TSX, including the `.TO` suffix. The script normalises them automatically before querying TMX:

| You enter | Queried as |
|-----------|------------|
| `ENB.TO` | `ENB` |
| `GRT-UN.TO` | `GRT.UN` |
| `CASH` | written as `0%`, skipped |

---

## Output

### Dividend yield (Column T)

| Value | Meaning |
|-------|---------|
| `0.000%` formatted value | Yield fetched successfully |
| `0.000%` (zero) | Stock pays no dividend, or ticker is `CASH` |
| `NOT FOUND` | Ticker not recognised by TMX |
| `ERROR` | Network or parsing error — check the Apps Script logs |

### Dividend payable date (Column N)

| Value | Meaning |
|-------|---------|
| Date formatted `dd-mmm-yy` | Next upcoming payable date |
| `01-Dec-99` | TMX returned no payable date for this ticker |
| *(unchanged)* | Payable date is in the past — existing cell value preserved |

### Share price (Column G)

Only rows whose ticker appears in `SHARE_PRICE_TARGET_TICKERS` are updated. On any API failure the existing cell value is **left unchanged** — the script never overwrites a good price with an error marker. Check the Apps Script logs to see if any failures occurred.

To view logs: **View** → **Logs** in the Apps Script editor.

---

## Notes

- This script targets **TSX-listed stocks** via the unofficial TMX GraphQL API (`app-money.tmx.com`). It will not work for US or other non-TSX tickers.
- Rows where **shares = 0 or blank** are skipped entirely.
- A 300 ms delay is applied after each API call to avoid rate-limiting.
- To schedule the script, go to **Triggers** (clock icon in the Apps Script editor) and add a time-driven trigger for `runUpdatePortfolioData`.
