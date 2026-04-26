// --- Global Configuration ---
const DIVIDEND_YIELD_SHEET_NAME = "Portfolio";     // Name of the worksheet to update
const DIVIDEND_YIELD_TICKER_COL = 4;               // Column D: ticker symbol
const DIVIDEND_YIELD_SHARES_COL = 5;               // Column E: number of shares
const DIVIDEND_YIELD_OUTPUT_COL = 20;              // Column T: dividend yield (output)
const DIVIDEND_YIELD_HEADER_ROWS = 1;              // Number of header rows to skip

const PAYABLE_DATE_COL = 14;                       // Column N: dividend payable date (output)

const SHARE_PRICE_OUTPUT_COL = 7;                  // Column G: share price (output)
const SHARE_PRICE_TARGET_TICKERS = ["GRT-UN.TO", "REI-UN.TO"]; // Only these tickers will be updated
// ---------------------------

function runUpdatePortfolioData() {
  Logger.log("Portfolio update started");
  runUpdateDividendYields();
  runUpdateSelectedSharePrices();
  Logger.log("Portfolio update ended");
}

function runUpdateDividendYields() {
  Logger.log("Yield update started");
  const sheet = helperGetConfiguredSheet(DIVIDEND_YIELD_SHEET_NAME);
  if (!sheet) return;

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const lastRow = sheet.getLastRow();

  for (let row = DIVIDEND_YIELD_HEADER_ROWS + 1; row <= lastRow; row++) {
    const rawTicker = sheet.getRange(row, DIVIDEND_YIELD_TICKER_COL).getValue().toString().trim().toUpperCase();
    if (!rawTicker) continue;

    // Skip if shares is zero or empty
    const shares = parseFloat(sheet.getRange(row, DIVIDEND_YIELD_SHARES_COL).getValue()) || 0;
    if (shares <= 0) continue;

    if (rawTicker === "CASH") {
      sheet.getRange(row, DIVIDEND_YIELD_OUTPUT_COL).setValue(0).setNumberFormat("0.000%");
      continue;
    }

    const ticker = helperNormalizeTicker(rawTicker);

    try {
      const quote = helperFetchQuoteBySymbol(ticker, ["dividendYield", "dividendPayDate"]);
      Logger.log(`${ticker}: ${JSON.stringify(quote)}`);

      if (!quote) {
        sheet.getRange(row, DIVIDEND_YIELD_OUTPUT_COL).setValue("NOT FOUND");
        Logger.log(`No data for ${ticker}`);
        continue;
      }

      const yieldVal = quote.dividendYield;
      const yieldOut = yieldVal == null ? 0 : yieldVal / 100;
      sheet.getRange(row, DIVIDEND_YIELD_OUTPUT_COL).setValue(yieldOut).setNumberFormat("0.000%");

      const rawPayDate = quote.dividendPayDate;
      if (!rawPayDate) {
        Logger.log(`${ticker}: no payable date — writing 01-Dec-99`);
        sheet.getRange(row, PAYABLE_DATE_COL).setValue(new Date(1999, 11, 1)).setNumberFormat("dd-mmm-yy");
      } else {
        const parts = rawPayDate.split("-");
        const newDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
        if (newDate > today) {
          sheet.getRange(row, PAYABLE_DATE_COL).setValue(newDate).setNumberFormat("dd-mmm-yy");
        }
      }

    } catch (e) {
      sheet.getRange(row, DIVIDEND_YIELD_OUTPUT_COL).setValue("ERROR");
      Logger.log(`Error for ${ticker}: ${e.message}`);
    }

    Utilities.sleep(300);
  }

  Logger.log("Yield update ended");
}

function runUpdateSelectedSharePrices() {
  Logger.log("Share price update started");
  const sheet = helperGetConfiguredSheet(DIVIDEND_YIELD_SHEET_NAME);
  if (!sheet) return;

  const targetTickerSet = new Set(SHARE_PRICE_TARGET_TICKERS.map(helperNormalizeTicker));
  const lastRow = sheet.getLastRow();

  for (let row = DIVIDEND_YIELD_HEADER_ROWS + 1; row <= lastRow; row++) {
    const rawTicker = sheet.getRange(row, DIVIDEND_YIELD_TICKER_COL).getValue().toString().trim().toUpperCase();
    if (!rawTicker || rawTicker === "CASH") continue;

    const shares = parseFloat(sheet.getRange(row, DIVIDEND_YIELD_SHARES_COL).getValue()) || 0;
    if (shares <= 0) continue;

    const ticker = helperNormalizeTicker(rawTicker);
    if (!targetTickerSet.has(ticker)) continue;

    Logger.log(`Updating share price for ${ticker} on row ${row}`);

    try {
      const quote = helperFetchQuoteBySymbol(ticker, ["price"]);

      if (!quote) {
        sheet.getRange(row, SHARE_PRICE_OUTPUT_COL).setValue("NOT FOUND");
        Logger.log(`No price data for ${ticker}`);
        continue;
      }

      const sharePrice = quote.price;
      if (sharePrice == null) {
        sheet.getRange(row, SHARE_PRICE_OUTPUT_COL).setValue("NOT FOUND");
        Logger.log(`No price returned for ${ticker}`);
        continue;
      }

      sheet.getRange(row, SHARE_PRICE_OUTPUT_COL).setValue(sharePrice);

    } catch (e) {
      sheet.getRange(row, SHARE_PRICE_OUTPUT_COL).setValue("ERROR");
      Logger.log(`Error for ${ticker}: ${e.message}`);
    }

    Utilities.sleep(300);
  }

  Logger.log("Share price update ended");
}

function helperGetConfiguredSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet "${sheetName}" not found. Please check the SHEET_NAME setting.`);
    return null;
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
    payload: JSON.stringify({ query: query }),
    muteHttpExceptions: true,
    headers: {
      "Origin": "https://money.tmx.com",
      "Referer": "https://money.tmx.com/"
    }
  };

  const response = UrlFetchApp.fetch("https://app-money.tmx.com/graphql", options);
  const data = JSON.parse(response.getContentText());
  return data?.data?.getQuoteBySymbol || null;
}
