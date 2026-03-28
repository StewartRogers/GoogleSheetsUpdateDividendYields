function updateDividendYields() {
  // --- Configuration ---
  const SHEET_NAME        = "Portfolio"; // Name of the worksheet to update
  const TICKER_COL        = 4;           // Column D: ticker symbol
  const SHARES_COL        = 5;           // Column E: number of shares
  const DIVIDEND_YIELD_COL = 20;         // Column T: dividend yield (output)
  const HEADER_ROWS       = 1;           // Number of header rows to skip
  // ---------------------

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet "${SHEET_NAME}" not found. Please check the SHEET_NAME setting.`);
    return;
  }

  const lastRow = sheet.getLastRow();

  for (let row = HEADER_ROWS + 1; row <= lastRow; row++) {
    const rawTicker = sheet.getRange(row, TICKER_COL).getValue().toString().trim().toUpperCase();
    if (!rawTicker) continue;

    // Skip if shares is zero or empty
    const shares = parseFloat(sheet.getRange(row, SHARES_COL).getValue()) || 0;
    if (shares <= 0) continue;

    if (rawTicker === "CASH") {
      sheet.getRange(row, DIVIDEND_YIELD_COL).setValue(0).setNumberFormat("0.000%");
      continue;
    }

    // Strip .TO, then convert hyphens to dots (e.g. GRT-UN.TO → GRT.UN)
    const ticker = rawTicker.replace(/\.TO$/i, "").replace(/-/g, ".");

    try {
      const query = `{
        getQuoteBySymbol(symbol: "${ticker}", locale: "en") {
          dividendYield
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
      const quote = data?.data?.getQuoteBySymbol;

      if (!quote) {
        sheet.getRange(row, DIVIDEND_YIELD_COL).setValue("NOT FOUND");
        Logger.log(`No data for ${ticker}`);
        continue;
      }

      const yieldVal = quote.dividendYield;
      if (yieldVal == null) {
        sheet.getRange(row, DIVIDEND_YIELD_COL).setValue(0).setNumberFormat("0.000%");
        continue;
      }

      sheet.getRange(row, DIVIDEND_YIELD_COL).setValue(yieldVal / 100).setNumberFormat("0.000%");

    } catch (e) {
      sheet.getRange(row, DIVIDEND_YIELD_COL).setValue("ERROR");
      Logger.log(`Error for ${ticker}: ${e.message}`);
    }

    Utilities.sleep(300);
  }
}
