/**
 * Compute teh average buy price of all the currently owned stocks
 */
var computeAverageBuyPrice = function() {
  
  let colSettings = loadColumnSettings(SHEET_CURRENTPRICE);

  let sheetTransTotal = SpreadsheetApp.getActive().getSheetByName(SHEET_TRANSTOTAL);
  let [rows, columns] = [sheetTransTotal.getLastRow(), sheetTransTotal.getLastColumn()];
  let dataTransTotal = sheetTransTotal.getRange(2, 1, rows, columns).getValues();

  let sheetStockTrans = SpreadsheetApp.getActive().getSheetByName(SHEET_STOCKTRANS);
  [rows, columns] = [sheetStockTrans.getLastRow(), sheetStockTrans.getLastColumn()];
  let datasheetStockTrans = sheetStockTrans.getRange(3, 1, rows, columns).getValues();

  let currentPriceSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_CURRENTPRICE);
  for (var row = ROWINDEX_CURRENTPRICE_FIRSTROW; row <= currentPriceSheet.getLastRow(); row++) {
    let stockCode1 = currentPriceSheet.getRange(colSettings.StockCode + row).getValue();
    if (stockCode1 == '') {
      break;
    }
    //if (stockCode1 != 'PSB') {
    //  continue;
    //}
    cellRange = colSettings.Quantity + row;
    currentPriceSheet.getRange(cellRange).setBackground("yellow");
    cellRange = colSettings.AveragePricePerShare + row;
    currentPriceSheet.getRange(cellRange).setBackground("yellow");
    SpreadsheetApp.flush();
    
    let computedAveragePrice = computeAverageBuyPriceByStockCodeInternal_(stockCode1, null, dataTransTotal, datasheetStockTrans);
    console.log(computedAveragePrice);
    if (computedAveragePrice.AverageBuyPricePerShare != null) {
      cellRange = colSettings.Quantity + row;
      // Do not change the quantity
      //currentPriceSheet.getRange(cellRange).setValue(computedAveragePrice.OwnedStockQuantity);
      cellRange = colSettings.AveragePricePerShare + row;
      currentPriceSheet.getRange(cellRange).setValue(computedAveragePrice.AverageBuyPricePerShare);
    }

    cellRange = colSettings.Quantity + row;
    currentPriceSheet.getRange(cellRange).setBackground(null);
    cellRange = colSettings.AveragePricePerShare + row;
    currentPriceSheet.getRange(cellRange).setBackground(null);
  }
}

/**
 * Compute the average buy price of an individual stock within a said date
 */
computeAveragePriceByStockCode = function(stockCode, cutOffDate) {

  let sheetTransTotal = SpreadsheetApp.getActive().getSheetByName(SHEET_TRANSTOTAL);
  let [rows, columns] = [sheetTransTotal.getLastRow(), sheetTransTotal.getLastColumn()];
  let dataTransTotal = sheetTransTotal.getRange(2, 1, rows, columns).getValues();

  let sheetStockTrans = SpreadsheetApp.getActive().getSheetByName(SHEET_STOCKTRANS);
  [rows, columns] = [sheetStockTrans.getLastRow(), sheetStockTrans.getLastColumn()];
  let datasheetStockTrans = sheetStockTrans.getRange(3, 1, rows, columns).getValues();

  return computeAverageBuyPriceByStockCodeInternal_(stockCode, cutOffDate, dataTransTotal, datasheetStockTrans);
}

/**
 * Internal function for computing the average buy price for a specified stocks with cutoff date
 * The date sheet is passed for performance purpose when calling this method is required for long check
 */
computeAverageBuyPriceByStockCodeInternal_ = function(stockCode, cutOffDate, dataTransTotal, datasheetStockTrans) {

  let stockTotalBuyAmount = 0;
  let stockBuyQuantityCounter = 0;
  let stockAveragePricePerShare = 0;
  let cutOffDateTime = new Date(cutOffDate);

  // Get shares owned from total
  let sharesOwned = 0;
  for (let iRow = 0; iRow < dataTransTotal.length; iRow++) {
    let transTotalStockCode = dataTransTotal[iRow][COLINDEX_TRANSTOTAL_STOCKCODE];
    if (stockCode != transTotalStockCode) {
      continue;
    }
    sharesOwned = parseInt(dataTransTotal[iRow][COLINDEX_TRANSTOTAL_QUANTITY], 10);
    break;
  }

  // Loop through all stock transaction and compute average price until all is found.
  for (let iRow = 0; iRow < datasheetStockTrans.length; iRow++) {
    let stockTransStockCode = datasheetStockTrans[iRow][COLINDEX_STOCKTRANS_STOCKCODE];
    let transactType = datasheetStockTrans[iRow][COLINDEX_STOCKTRANS_TRANSTYPE];
    let transactQuantity = datasheetStockTrans[iRow][COLINDEX_STOCKTRANS_QUANTITY];
    let transactDate = datasheetStockTrans[iRow][COLINDEX_STOCKTRANS_TRANSDATE];
    let netBuyAmt = datasheetStockTrans[iRow][COLINDEX_STOCKTRANS_NETBUYAMT];

    if (stockCode != stockTransStockCode) {
      continue;
    }

    if (cutOffDate != null) {
      let transactDateTime = new Date(transactDate);
      if (cutOffDateTime.getTime() < transactDateTime.getTime()) {
        continue;
      }
    }

    if (transactType == TRANSTYPE_BOUGHTSHARES ||
      transactType == TRANSTYPE_STOCKRIGHTS || 
      transactType == TRANSTYPE_STOCKDIVIDEND || 
      transactType == TRANSTYPE_IPOBUYSHARES
    ) {
      // If there are partially bought, compute
      let allocatedQuantity = 
        (stockBuyQuantityCounter + transactQuantity <= sharesOwned || sharesOwned <= 0) ? 
          transactQuantity : 
          (sharesOwned - stockBuyQuantityCounter);
      stockTotalBuyAmount += (netBuyAmt * allocatedQuantity / transactQuantity);
      stockBuyQuantityCounter += allocatedQuantity;
      console.log('index ' + (iRow + 3) + ' allocqty: ' + allocatedQuantity + 
        ' total: ' + stockTotalBuyAmount + ' count: ' + stockBuyQuantityCounter);
    }

    // If quantity owned is reached
    // Forget about past buy price since you can sell in between, only compute last n bought shares
    if (stockBuyQuantityCounter >= sharesOwned && sharesOwned > 0) {
      break;
    }
  }
  if (stockBuyQuantityCounter > 0) {
    stockAveragePricePerShare = stockTotalBuyAmount / stockBuyQuantityCounter;
  }
  
  if (stockAveragePricePerShare > 0) {
    return {
      'StockCode': stockCode,
      'CutOffDate': cutOffDate,
      'OwnedStockQuantity': sharesOwned,
      'AverageBuyPricePerShare': stockAveragePricePerShare,
    };
  } else {
    return {};
  }
}
