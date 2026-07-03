# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project is

A single-file Google Apps Script (`code.gs`) that fetches TSX stock data from the TMX GraphQL API — and data for a configurable set of US tickers from Yahoo Finance — and writes it into a Google Sheet. There is no local build toolchain — the script is pasted directly into the Google Apps Script editor and executed there.

## Deployment

There are no build, lint, or test commands. To deploy a change:
1. Copy the updated `code.gs` contents.
2. Open the Google Sheet → **Extensions** → **Apps Script**.
3. Replace the editor contents and save (`Ctrl+S`).
4. Run the desired function from the function dropdown.

## Architecture

All code lives in `code.gs`. The structure is:

- **Global config constants** at the top — column numbers, sheet name, and the ticker sets for share price updates (`SHARE_PRICE_TARGET_TICKERS`) and Yahoo-routed US tickers (`US_TICKERS`). These are the only values users should need to change.
- **Three public entry-point functions** called from the Apps Script editor or a trigger:
  - `runUpdatePortfolioData()` — calls `helperProcessRows` with both flags true (single pass)
  - `runUpdateDividendYields()` — calls `helperProcessRows(sheet, true, false)`
  - `runUpdateSelectedSharePrices()` — calls `helperProcessRows(sheet, false, true)`
- **`helperProcessRows(sheet, fetchDividends, fetchPrices)`** — the core loop. Reads input columns D–N in one `getValues()` call, reads the existing yield output buffer, processes each row, then batch-writes yield and payable date columns with `setValues()`. When both flags are true, dividend and price fields are combined into a single API call per row. Tickers in `US_TICKERS` are routed to `helperFetchYahooQuote`; all others go to `helperFetchQuoteBySymbol` (TMX). Share prices are written with a direct per-cell `setValue()` rather than a batch write — this preserves any formulas the user has in column G that would be destroyed by a column-wide `setValues()`.
- **Shared helpers**:
  - `helperGetConfiguredSheet(sheetName)` — looks up the sheet by name; logs and attempts a UI alert on failure (the alert is silently skipped when running from a trigger with no browser context)
  - `helperNormalizeTicker(rawTicker)` — strips `.TO` suffix and converts `-` to `.` (e.g. `GRT-UN.TO` → `GRT.UN`) to match the TMX symbol format; not applied to `US_TICKERS` entries
  - `helperFetchQuoteBySymbol(ticker, fields)` — POSTs a GraphQL query to `https://app-money.tmx.com/graphql`, returns `null` on non-200 HTTP responses or GraphQL errors (both are logged), otherwise returns the `getQuoteBySymbol` object
  - `helperGetYahooSession()` — obtains an unauthenticated cookie from `fc.yahoo.com` and exchanges it for a crumb from `query1.finance.yahoo.com/v1/test/getcrumb`; the result is memoized in the module-level `_yahooSessionCache` so it's fetched once per script execution rather than once per ticker
  - `helperFetchYahooQuote(ticker, fields)` — fetches `summaryDetail`/`calendarEvents`/`price` modules from `query1.finance.yahoo.com/v10/finance/quoteSummary`, authenticated with the cached session; returns `null` on session failure, non-200 responses, or a missing result

## Key behaviours

- Rows where shares ≤ 0 or blank are skipped entirely.
- `CASH` tickers are written as `0%` yield without an API call.
- The TMX API returns `dividendYield` as a percentage integer (e.g. `5.2` means 5.2%); Yahoo returns it as a decimal fraction (e.g. `0.052`). Both are normalized to the TMX percentage-integer convention before the script divides by 100 for the final write.
- Payable dates are only updated if the new date is in the future — past dates are left unchanged. If the API returns no date, the cell is left unchanged (no sentinel is written).
- `targetTickerSet` is built by uppercasing and normalizing `SHARE_PRICE_TARGET_TICKERS` so that case variations in the config array don't cause silent mismatches. `US_TICKERS` entries are matched as-is (uppercase, no TSX normalization).
- On any failure — no data found, a missing field, or an exception — dividend yield and payable date cells are left unchanged rather than overwritten with an error marker; check the Apps Script logs for details. Share price cells follow the same rule: only a successful price fetch writes to column G.
- `dividendPayDate` from the API is a `YYYY-MM-DD` string; the script parses it and only overwrites the cell if the resulting date is in the future. Unexpected formats are logged and the cell is left unchanged.
- TSX-listed tickers are queried against TMX. US-listed tickers must be explicitly added to `US_TICKERS` to be queried against Yahoo Finance; any other non-TSX ticker will not resolve.
- Yahoo Finance requires session auth (cookie + crumb) that TMX does not; `helperGetYahooSession()` fetches this once per script execution and caches it, so a run touching multiple `US_TICKERS` rows doesn't re-authenticate per row.
