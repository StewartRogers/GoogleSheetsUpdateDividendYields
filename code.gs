// --- Global Configuration ---
const DIVIDEND_YIELD_SHEET_NAME = "Portfolio";
const DIVIDEND_YIELD_TICKER_COL = 4;               // Column D: ticker symbol
const DIVIDEND_YIELD_SHARES_COL = 5;               // Column E: number of shares
const DIVIDEND_YIELD_OUTPUT_COL = 20;              // Column T: dividend yield (output)
const DIVIDEND_YIELD_HEADER_ROWS = 1;

const PAYABLE_DATE_COL = 14;                       // Column N: dividend payable date (output)
const SHARE_PRICE_OUTPUT_COL = 7;                  // Column G: share price (output)
const SHARE_PRICE_TARGET_TICKERS = ["GRT-UN.TO", "REI-UN.TO"];

// Tickers treated as non-equity placeholders: written as 0% yield, no API call made
const NON_EQUITY_TICKERS = new Set(["CASH"]);

// US-listed tickers (Nasdaq/NYSE) fetched from Yahoo Finance instead of TMX
const US_TICKERS = new Set(["SPCX"]);

// Written to payable date column when the API returns no date for a ticker
const NO_PAY_DATE = new Date(1999, 11, 1);
// ---------------------------

function runUpdatePortfolioData() {
  Logger.log("Portfolio update started");
  const sheet = helperGetConfiguredSheet(DIVIDEND_YIELD_SHEET_NAME);
  if (!sheet) return;
  helperProcessRows(sheet, true, true);
  Logger.log("Portfolio update ended");
}

function runUpdateDividendYields() {
  Logger.log("Yield update started");
  const sheet = helperGetConfiguredSheet(DIVIDEND_YIELD_SHEET_NAME);
  if (!sheet) return;
  helperProcessRows(sheet, true, false);
  Logger.log("Yield update ended");
}

function runUpdateSelectedSharePrices() {
  Logger.log("Share price update started");
  const sheet = helperGetConfiguredSheet(DIVIDEND_YIELD_SHEET_NAME);
  if (!sheet) return;
  helperProcessRows(sheet, false, true);
  Logger.log("Share price update ended");
}

function helperProcessRows(sheet, fetchDividends, fetchPrices) {
  const lastRow = sheet.getLastRow();
  const firstDataRow = DIVIDEND_YIELD_HEADER_ROWS + 1;
  const numRows = lastRow - DIVIDEND_YIELD_HEADER_ROWS;
  if (numRows <= 0) return;

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const targetTickerSet = new Set(SHARE_PRICE_TARGET_TICKERS.map(t => helperNormalizeTicker(t.toUpperCase())));

  // ── Batch reads ──────────────────────────────────────────────────────────
  // Read columns D–N in one call: ticker (D=4), shares (E=5), ..., payable date (N=14)
  const inputCols = PAYABLE_DATE_COL - DIVIDEND_YIELD_TICKER_COL + 1;
  const inputData = sheet.getRange(firstDataRow, DIVIDEND_YIELD_TICKER_COL, numRows, inputCols).getValues();
  const tickerIdx  = 0;
  const sharesIdx  = DIVIDEND_YIELD_SHARES_COL - DIVIDEND_YIELD_TICKER_COL;
  const payDateIdx = PAYABLE_DATE_COL          - DIVIDEND_YIELD_TICKER_COL;

  // Read existing yield output so unchanged rows are preserved in the batch write
  const yieldBuf = fetchDividends
    ? sheet.getRange(firstDataRow, DIVIDEND_YIELD_OUTPUT_COL, numRows, 1).getValues()
    : null;
  // Note: share prices are written per-cell (not batched) to avoid overwriting
  // formulas in column G for rows that are not target tickers.
  // payDateBuf defaults to existing payable dates (already in inputData)
  const payDateBuf = inputData.map(r => [r[payDateIdx]]);

  // ── Process each row ─────────────────────────────────────────────────────
  for (let i = 0; i < numRows; i++) {
    const rawTicker = (inputData[i][tickerIdx] ?? "").toString().trim().toUpperCase();
    if (!rawTicker) continue;

    const rawShares = inputData[i][sharesIdx];
    const shares = parseFloat(rawShares) || 0;
    if (shares <= 0) {
      if (rawShares !== "" && rawShares != null && isNaN(Number(rawShares)))
        Logger.log(`Row ${i + firstDataRow}: non-numeric shares value "${rawShares}" — row skipped`);
      continue;
    }

    if (NON_EQUITY_TICKERS.has(rawTicker)) {
      if (fetchDividends) yieldBuf[i][0] = 0;
      continue;
    }

    const isUS = US_TICKERS.has(rawTicker);
    const ticker = isUS ? rawTicker : helperNormalizeTicker(rawTicker);
    const needsPrice = fetchPrices && targetTickerSet.has(ticker);
    if (!fetchDividends && !needsPrice) continue;

    const fields = [];
    if (fetchDividends) fields.push("dividendYield", "dividendPayDate");
    if (needsPrice)     fields.push("price");

    try {
      const quote = isUS
        ? helperFetchYahooQuote(ticker, fields)
        : helperFetchQuoteBySymbol(ticker, fields);
      Logger.log(`${ticker}: ${JSON.stringify(quote)}`);

      if (!quote) {
        Logger.log(`No data for ${ticker} — yield and pay date left unchanged${needsPrice ? ", share price cell left unchanged" : ""}`);
      } else {
        if (fetchDividends) {
          const yieldVal = quote.dividendYield;
          if (yieldVal != null) yieldBuf[i][0] = yieldVal / 100;
          else Logger.log(`${ticker}: no yield returned — yield left unchanged`);

          const rawPayDate = quote.dividendPayDate;
          if (!rawPayDate) {
            Logger.log(`${ticker}: no payable date — pay date left unchanged`);
          } else {
            const parts = rawPayDate.split("-");
            if (parts.length === 3) {
              const year  = parseInt(parts[0], 10);
              const month = parseInt(parts[1], 10) - 1;
              const day   = parseInt(parts[2], 10);
              if (isNaN(year) || month < 0 || month > 11 || isNaN(day)) {
                Logger.log(`${ticker}: invalid date values in "${rawPayDate}" — payable date unchanged`);
              } else {
                const newDate = new Date(year, month, day);
                if (newDate > today) payDateBuf[i][0] = newDate;
              }
            } else {
              Logger.log(`${ticker}: unexpected date format "${rawPayDate}" — payable date unchanged`);
            }
          }
        }

        if (needsPrice) {
          if (quote.price == null) {
            Logger.log(`${ticker}: no price returned — share price cell left unchanged`);
          } else {
            sheet.getRange(firstDataRow + i, SHARE_PRICE_OUTPUT_COL).setValue(quote.price);
          }
        }
      }
    } catch (e) {
      Logger.log(`Error for ${ticker}: ${e.message}${needsPrice ? " — share price cell left unchanged" : ""} — yield left unchanged`);
    }

    Utilities.sleep(300);
  }

  // ── Batch writes ─────────────────────────────────────────────────────────
  if (fetchDividends) {
    const yieldRange = sheet.getRange(firstDataRow, DIVIDEND_YIELD_OUTPUT_COL, numRows, 1);
    yieldRange.setValues(yieldBuf);
    // Apply percent format only to numeric cells; set error cells to plain text
    // so formulas referencing this column receive a number, not a string.
    const yieldFormats = yieldBuf.map(r => [typeof r[0] === "number" ? "0.000%" : "@"]);
    yieldRange.setNumberFormats(yieldFormats);
    sheet.getRange(firstDataRow, PAYABLE_DATE_COL, numRows, 1).setValues(payDateBuf).setNumberFormat("dd-mmm-yy");
  }
}

function helperGetConfiguredSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    const msg = `Sheet "${sheetName}" not found. Please check the SHEET_NAME setting.`;
    Logger.log(msg);
    try { SpreadsheetApp.getUi().alert(msg); } catch (_) {}
  }
  return sheet;
}

function helperNormalizeTicker(rawTicker) {
  return rawTicker.replace(/\.TO$/i, "").replace(/-/g, ".");
}

function helperGetYahooSession() {
  // Step 1: obtain a session cookie from Yahoo
  const cookieResp = UrlFetchApp.fetch("https://fc.yahoo.com", {
    muteHttpExceptions: true,
    followRedirects: false
  });
  const rawCookie = cookieResp.getAllHeaders()["Set-Cookie"];
  const cookie = Array.isArray(rawCookie) ? rawCookie.join("; ") : (rawCookie || "");

  // Step 2: exchange the cookie for a crumb
  const crumbResp = UrlFetchApp.fetch("https://query1.finance.yahoo.com/v1/test/getcrumb", {
    muteHttpExceptions: true,
    headers: { "Cookie": cookie, "User-Agent": "Mozilla/5.0" }
  });
  if (crumbResp.getResponseCode() !== 200) {
    Logger.log(`Yahoo crumb fetch failed (HTTP ${crumbResp.getResponseCode()})`);
    return null;
  }
  return { cookie, crumb: crumbResp.getContentText().trim() };
}

function helperFetchYahooQuote(ticker, fields) {
  const needsDividends = fields.includes("dividendYield") || fields.includes("dividendPayDate");
  const needsPrice     = fields.includes("price");
  const modules = [];
  if (needsDividends) modules.push("summaryDetail", "calendarEvents");
  if (needsPrice)     modules.push("price");

  const session = helperGetYahooSession();
  if (!session) return null;

  const url = `https://query1.finance.yahoo.com/v10/finance/quoteSummary/${ticker}?modules=${modules.join(",")}&crumb=${encodeURIComponent(session.crumb)}`;
  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: { "Cookie": session.cookie, "User-Agent": "Mozilla/5.0" }
  });

  if (response.getResponseCode() !== 200) {
    Logger.log(`Yahoo HTTP ${response.getResponseCode()} for ${ticker}: ${response.getContentText()}`);
    return null;
  }

  const parsed = JSON.parse(response.getContentText());
  const result = parsed?.quoteSummary?.result?.[0];
  if (!result) {
    Logger.log(`Yahoo: no result for ${ticker}`);
    return null;
  }

  const out = {};
  if (needsDividends) {
    // Yahoo returns yield as a decimal (0.05 = 5%); multiply by 100 to match TMX format
    const raw = result.summaryDetail?.dividendYield?.raw;
    out.dividendYield = raw != null ? raw * 100 : null;

    const ts = result.calendarEvents?.dividendDate?.raw;
    if (ts) {
      const d = new Date(ts * 1000);
      out.dividendPayDate = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    }
  }
  if (needsPrice) {
    out.price = result.price?.regularMarketPrice?.raw ?? null;
  }
  return out;
}

function helperFetchQuoteBySymbol(ticker, fields) {
  const query = `{
    getQuoteBySymbol(symbol: "${ticker}", locale: "en") {
      ${fields.join("\n      ")}
    }
  }`;

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ query }),
    muteHttpExceptions: true,
    headers: {
      "Origin": "https://money.tmx.com",
      "Referer": "https://money.tmx.com/"
    }
  };

  const response = UrlFetchApp.fetch("https://app-money.tmx.com/graphql", options);

  if (response.getResponseCode() !== 200) {
    Logger.log(`HTTP ${response.getResponseCode()} for ${ticker}: ${response.getContentText()}`);
    return null;
  }

  const parsed = JSON.parse(response.getContentText());
  if (parsed.errors) {
    Logger.log(`GraphQL errors for ${ticker}: ${JSON.stringify(parsed.errors)}`);
    return null;
  }

  return parsed?.data?.getQuoteBySymbol || null;
}
