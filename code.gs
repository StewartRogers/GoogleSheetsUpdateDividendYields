// --- Global Configuration ---
const DIVIDEND_YIELD_SHEET_NAME = "Portfolio";
const DIVIDEND_YIELD_TICKER_COL = 4;               // Column D: ticker symbol
const DIVIDEND_YIELD_SHARES_COL = 5;               // Column E: number of shares
const DIVIDEND_YIELD_OUTPUT_COL = 20;              // Column T: dividend yield (output)
const DIVIDEND_YIELD_HEADER_ROWS = 1;

const PAYABLE_DATE_COL = 14;                       // Column N: dividend payable date (output)
const SHARE_PRICE_OUTPUT_COL = 7;                  // Column G: share price (output)
const SHARE_PRICE_TARGET_TICKERS = ["GRT-UN.TO", "REI-UN.TO"];

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

  // Read existing output columns so unchanged rows are preserved in the batch write
  const yieldBuf = fetchDividends
    ? sheet.getRange(firstDataRow, DIVIDEND_YIELD_OUTPUT_COL, numRows, 1).getValues()
    : null;
  const priceBuf = fetchPrices
    ? sheet.getRange(firstDataRow, SHARE_PRICE_OUTPUT_COL, numRows, 1).getValues()
    : null;
  // payDateBuf defaults to existing payable dates (already in inputData)
  const payDateBuf = inputData.map(r => [r[payDateIdx]]);

  // ── Process each row ─────────────────────────────────────────────────────
  for (let i = 0; i < numRows; i++) {
    const rawTicker = (inputData[i][tickerIdx] ?? "").toString().trim().toUpperCase();
    if (!rawTicker) continue;

    // Rows with no current position are skipped intentionally for both yield and price
    const shares = parseFloat(inputData[i][sharesIdx]) || 0;
    if (shares <= 0) continue;

    if (rawTicker === "CASH") {
      if (fetchDividends) yieldBuf[i][0] = 0;
      continue;
    }

    const ticker = helperNormalizeTicker(rawTicker);
    const needsPrice = fetchPrices && targetTickerSet.has(ticker);
    if (!fetchDividends && !needsPrice) continue;

    // Combine all needed fields into one TMX call to minimise API requests
    const fields = [];
    if (fetchDividends) fields.push("dividendYield", "dividendPayDate");
    if (needsPrice)     fields.push("price");

    try {
      const quote = helperFetchQuoteBySymbol(ticker, fields);
      Logger.log(`${ticker}: ${JSON.stringify(quote)}`);

      if (!quote) {
        if (fetchDividends) yieldBuf[i][0] = "NOT FOUND";
        if (needsPrice)     priceBuf[i][0] = "NOT FOUND";
        Logger.log(`No data for ${ticker}`);
      } else {
        if (fetchDividends) {
          const yieldVal = quote.dividendYield;
          yieldBuf[i][0] = yieldVal == null ? 0 : yieldVal / 100;

          const rawPayDate = quote.dividendPayDate;
          if (!rawPayDate) {
            Logger.log(`${ticker}: no payable date — writing sentinel`);
            payDateBuf[i][0] = NO_PAY_DATE;
          } else {
            const parts = rawPayDate.split("-");
            if (parts.length === 3) {
              const newDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
              if (newDate > today) payDateBuf[i][0] = newDate;
            } else {
              Logger.log(`${ticker}: unexpected date format "${rawPayDate}" — payable date unchanged`);
            }
          }
        }

        if (needsPrice) {
          priceBuf[i][0] = quote.price == null ? "NOT FOUND" : quote.price;
        }
      }
    } catch (e) {
      if (fetchDividends) yieldBuf[i][0] = "ERROR";
      if (needsPrice)     priceBuf[i][0] = "ERROR";
      Logger.log(`Error for ${ticker}: ${e.message}`);
    }

    Utilities.sleep(300);
  }

  // ── Batch writes ─────────────────────────────────────────────────────────
  if (fetchDividends) {
    sheet.getRange(firstDataRow, DIVIDEND_YIELD_OUTPUT_COL, numRows, 1).setValues(yieldBuf).setNumberFormat("0.000%");
    sheet.getRange(firstDataRow, PAYABLE_DATE_COL,          numRows, 1).setValues(payDateBuf).setNumberFormat("dd-mmm-yy");
  }
  if (fetchPrices) {
    sheet.getRange(firstDataRow, SHARE_PRICE_OUTPUT_COL, numRows, 1).setValues(priceBuf);
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
