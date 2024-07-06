/**
 * Include the projected dividend and include them in the cash dividend data sheet.
 */
function addProjectedDividendToCashDividend() {
  let colSettings = loadColumnSettings(SHEET_CASHDIV);
  let sheetCashDividend = SpreadsheetApp.getActive().getSheetByName(SHEET_CASHDIV);
  let [rows, columns] = [sheetCashDividend.getLastRow(), sheetCashDividend.getLastColumn()];

  let sheetProjectedDividend = SpreadsheetApp.getActive().getSheetByName(SHEET_PROJDIV);
  [rows, columns] = [sheetProjectedDividend.getLastRow(), sheetProjectedDividend.getLastColumn()];
  let dataProjectedDividend = sheetProjectedDividend.getRange(3, 1, rows, columns).getValues();

  let sheetCurrentPrice = SpreadsheetApp.getActive().getSheetByName(SHEET_CURRENTPRICE);
  [rows, columns] = [sheetCurrentPrice.getLastRow(), sheetCurrentPrice.getLastColumn()];
  let dataCurrentPrice = sheetCurrentPrice.getRange(4, 1, rows, columns).getValues();
  
  let yesterdayDate = new Date();
  yesterdayDate.setDate(yesterdayDate.getDate() - 1);

  for (let iRow = dataProjectedDividend.length - 1; iRow >= 0; iRow--) {
    let paymentDate = new Date(dataProjectedDividend[iRow][COLINDEX_PROJDIV_PAYMENTDATE]);
    let stockCode = dataProjectedDividend[iRow][COLINDEX_PROJDIV_STOCKCODE];
    console.log('ding ' + paymentDate + ' ' + stockCode);

    if (stockCode == '') {
      continue;
    }
    // Skip old date
    if (paymentDate.getTime() <= yesterdayDate.getTime()) {
      continue;
    }

    // Since this is adding projected cash dividend, we use the last average price as assumed price
    let currentPriceAvgPricePerShare = null;
    for (let iCurrPriceRow = 0; iCurrPriceRow < dataCurrentPrice.length; iCurrPriceRow++) {
      let currentPriceStockCode = dataCurrentPrice[iCurrPriceRow][COLINDEX_CURRENTPRICE_STOCKCODE];
      if (stockCode == currentPriceStockCode) {
        currentPriceAvgPricePerShare = dataCurrentPrice[iCurrPriceRow][COLINDEX_CURRENTPRICE_AVGPRICEPERSHARE];
        break;
      }
    }

    let aNoticeInfo = {
        "TransactionInfo": {
            "MarkerType": -1,
            "PaymentDate": dataProjectedDividend[iRow][COLINDEX_PROJDIV_PAYMENTDATE],
            "StockCode": dataProjectedDividend[iRow][COLINDEX_PROJDIV_STOCKCODE],
            "DividendPerShare": dataProjectedDividend[iRow][COLINDEX_PROJDIV_DIVPERSHARE],
            "ExDividendDate": dataProjectedDividend[iRow][COLINDEX_PROJDIV_EXDIVDATE],
            "NoOfShare": dataProjectedDividend[iRow][COLINDEX_PROJDIV_QUANTITY],
            "GrossAmount": dataProjectedDividend[iRow][COLINDEX_PROJDIV_GROSSAMT],
            "WithholdingTax": 0,
            "NetAmount": dataProjectedDividend[iRow][COLINDEX_PROJDIV_NETAMT],
            "PricePerShare": currentPriceAvgPricePerShare
        },
        "AverageBuyPricePerShare": currentPriceAvgPricePerShare
      };

    logCashDividendNoticeToSheets_(colSettings, 3, aNoticeInfo, sheetCurrentPrice.getLastRow());
  }
}
