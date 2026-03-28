# GoogleSheetsUpdateDividendYields

A Google Apps Script that fetches current dividend yields for TSX-listed stocks from the TMX API and writes them into your Google Sheet.

---

## Spreadsheet Layout

The script expects your data to start on **row 2** (row 1 is the header). The relevant columns are:

| Column | Default | Contents |
|--------|---------|----------|
| D | 4 | Ticker symbol (e.g. `ENB.TO`, `GRT-UN.TO`, `CASH`) |
| E | 5 | Number of shares held |
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

At the top of the function, update the configuration constants to match your spreadsheet:

```js
const SHEET_NAME        = "Portfolio"; // The name of your sheet tab
const TICKER_COL        = 4;           // Column containing ticker symbols (D = 4)
const SHARES_COL        = 5;           // Column containing share counts (E = 5)
const DIVIDEND_YIELD_COL = 20;         // Column to write dividend yields into (T = 20)
const HEADER_ROWS       = 1;           // Number of header rows to skip
```

Set `SHEET_NAME` to exactly match the tab name at the bottom of your spreadsheet.

---

## Running the Script

1. In the Apps Script editor, make sure `updateDividendYields` is selected in the function dropdown.
2. Click the **Run** button (▶).
3. The first time you run it, Google will ask you to grant permissions:
   - **View and manage your spreadsheets** — to read tickers and write yields.
   - **Connect to an external service** — to fetch data from the TMX API.
   - Click **Review permissions** → choose your Google account → **Allow**.
4. The script will process each row and write the dividend yield into the configured output column.

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

| Value | Meaning |
|-------|---------|
| `0.000%` formatted value | Dividend yield fetched successfully |
| `0.000%` (zero) | Stock pays no dividend |
| `NOT FOUND` | Ticker not recognised by TMX |
| `ERROR` | Network or parsing error — check the Apps Script logs |

To view logs: **View** → **Logs** in the Apps Script editor.

---

## Notes

- This script targets **TSX-listed stocks** via the unofficial TMX GraphQL API (`app-money.tmx.com`). It will not work for US or other non-TSX tickers.
- Rows where **shares = 0 or blank** are skipped entirely.
- A 300ms delay between rows is included to avoid rate-limiting.
- The script runs manually. To schedule it, go to **Triggers** (clock icon in the Apps Script editor) and add a time-driven trigger for `updateDividendYields`.
