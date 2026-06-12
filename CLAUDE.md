# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project is

A single-file Google Apps Script (`code.gs`) that fetches TSX stock data from the TMX GraphQL API and writes it into a Google Sheet. There is no local build toolchain — the script is pasted directly into the Google Apps Script editor and executed there.

## Deployment

There are no build, lint, or test commands. To deploy a change:
1. Copy the updated `code.gs` contents.
2. Open the Google Sheet → **Extensions** → **Apps Script**.
3. Replace the editor contents and save (`Ctrl+S`).
4. Run the desired function from the function dropdown.

## Architecture

All code lives in `code.gs`. The structure is:

- **Global config constants** at the top — column numbers, sheet name, sentinel values, and target tickers for share price updates. These are the only values users should need to change.
- **Three public entry-point functions** called from the Apps Script editor or a trigger:
  - `runUpdatePortfolioData()` — calls `helperProcessRows` with both flags true (single pass)
  - `runUpdateDividendYields()` — calls `helperProcessRows(sheet, true, false)`
  - `runUpdateSelectedSharePrices()` — calls `helperProcessRows(sheet, false, true)`
- **`helperProcessRows(sheet, fetchDividends, fetchPrices)`** — the core loop. Reads input columns D–N in one `getValues()` call, reads the existing output buffers (yield column and/or price column) in up to two further reads, processes each row, accumulates results, then batch-writes all output columns with `setValues()`. When both flags are true, dividend and price fields are combined into a single TMX API call per row.
- **Three shared helpers**:
  - `helperGetConfiguredSheet(sheetName)` — looks up the sheet by name; logs and attempts a UI alert on failure (the alert is silently skipped when running from a trigger with no browser context)
  - `helperNormalizeTicker(rawTicker)` — strips `.TO` suffix and converts `-` to `.` (e.g. `GRT-UN.TO` → `GRT.UN`) to match the TMX symbol format
  - `helperFetchQuoteBySymbol(ticker, fields)` — POSTs a GraphQL query to `https://app-money.tmx.com/graphql`, returns `null` on non-200 HTTP responses or GraphQL errors (both are logged), otherwise returns the `getQuoteBySymbol` object

## Key behaviours

- Rows where shares ≤ 0 or blank are skipped entirely.
- `CASH` tickers are written as `0%` yield without an API call.
- The TMX API returns `dividendYield` as a percentage integer (e.g. `5.2` means 5.2%), so the script divides by 100 before writing.
- Payable dates are only updated if the new date is in the future — past dates are left unchanged. If the API returns no date, the sentinel value `NO_PAY_DATE` (1 Dec 1999) is written.
- `targetTickerSet` is built by uppercasing and normalizing `SHARE_PRICE_TARGET_TICKERS` so that case variations in the config array don't cause silent mismatches.
- `dividendPayDate` from the API is a `YYYY-MM-DD` string; the script parses it and only overwrites the cell if the resulting date is in the future. Unexpected formats are logged and the cell is left unchanged.
- Only TSX-listed tickers work — the TMX API does not cover US or other exchanges.
