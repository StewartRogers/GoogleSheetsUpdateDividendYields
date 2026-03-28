function updateDividendYields() {                                                                                                                   
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();                                                                             
    const lastRow = sheet.getLastRow();                                                                                                               
                  
    for (let row = 2; row <= lastRow; row++) {                                                                                                        
      const rawTicker = sheet.getRange(row, 4).getValue().toString().trim().toUpperCase();
      if (!rawTicker) continue;

      // Skip if shares (column E) is zero or empty
      const shares = parseFloat(sheet.getRange(row, 5).getValue()) || 0;
      if (shares <= 0) continue;

      if (rawTicker === "CASH") {
        sheet.getRange(row, 20).setValue(0).setNumberFormat("0.000%");
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
          sheet.getRange(row, 20).setValue("NOT FOUND");
          Logger.log(`No data for ${ticker}`);
          continue;
        }

        const yieldVal = quote.dividendYield;
        if (yieldVal == null) {
          sheet.getRange(row, 20).setValue(0).setNumberFormat("0.000%");
          continue;
        }

        sheet.getRange(row, 20).setValue(yieldVal / 100).setNumberFormat("0.000%");

      } catch (e) {
        sheet.getRange(row, 20).setValue("ERROR");
        Logger.log(`Error for ${ticker}: ${e.message}`);
      }

      Utilities.sleep(300);
    }
  }
