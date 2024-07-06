function fillupBoughtSharesMissing() {
  let colSettings = loadColumnSettings(SHEET_STOCKTRANS);

  let sheetStockTrans = SpreadsheetApp.getActive().getSheetByName(SHEET_STOCKTRANS);
  let [rows, columns] = [sheetStockTrans.getLastRow(), sheetStockTrans.getLastColumn()];
  let dataStockTrans = sheetStockTrans.getRange(3, 1, rows, columns).getValues();

  for (let iRow = 0; iRow < dataStockTrans.length; iRow++) {
    let stockCode = dataStockTrans[iRow][COLINDEX_STOCKTRANS_STOCKCODE];
    let transDateText = dataStockTrans[iRow][COLINDEX_STOCKTRANS_TRANSDATE];
    let formattedTransactDate = Utilities.formatDate(new Date(transDateText), 'Asia/Tokyo', 'yyyy-MM-dd');
    let quantity = dataStockTrans[iRow][COLINDEX_STOCKTRANS_QUANTITY];
    let transactType = dataStockTrans[iRow][COLINDEX_STOCKTRANS_TRANSTYPE];
    let buyPerPrice = dataStockTrans[iRow][COLINDEX_STOCKTRANS_BUYPRICEPERSHARE];

    // Stop if transact type missing
    if (transactType != '') {
      console.log('break at line ' + iRow);
      break;
    }

    if (stockCode == '' || buyPerPrice == '') {
      continue;
    }
    let rowIndex = iRow + 3;

    console.log('stock ' + stockCode);
    console.log('test' + transDateText);
    if (transDateText == '') {
      cellRange = colSettings.TransactionDate + rowIndex;
      sheetStockTrans.getRange(cellRange).setFormula('=NOW()');
    }
    if (quantity == '') {
      cellRange = colSettings.Quantity + rowIndex;
      sheetStockTrans.getRange(cellRange).setValue('1');
    }
    cellRange = colSettings.TransactionType + rowIndex;
    sheetStockTrans.getRange(cellRange).setValue(TRANSTYPE_BOUGHTSHARES);
    cellRange = colSettings.GrossBuyAmount + rowIndex;
    sheetStockTrans.getRange(cellRange).setFormula('=' + colSettings.Quantity + rowIndex + '*' + colSettings.BuyPricePerShare + rowIndex);
    cellRange = colSettings.NetBuyAmount + rowIndex;
    sheetStockTrans.getRange(cellRange).setFormula('=' + colSettings.GrossBuyAmount + rowIndex + '*1.00295');

    // Sector
    let sheetCurrentPrice = SpreadsheetApp.getActive().getSheetByName(SHEET_CURRENTPRICE);
    let luLastRow = sheetCurrentPrice.getLastRow();

    cellRange = colSettings.Sector + rowIndex;
    formulaVal = "=VLOOKUP(" + colSettings.StockCode + rowIndex + ",CurrentPrice!$A$3:$C$" + luLastRow + ",3,FALSE)";
    sheetStockTrans.getRange(cellRange).setFormula(formulaVal); 

    cellRange = colSettings.Broker + rowIndex;
    sheetStockTrans.getRange(cellRange).setValue('COL');

  }
}
