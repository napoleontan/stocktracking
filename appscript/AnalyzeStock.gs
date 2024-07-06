/**
 * Analyze the stock transaction and cash dividend price for the selected stock code
 */
function analyzeStock() {
  let currentSelectedSheetRange = SpreadsheetApp.getActiveRange();
  let originalStockCode = currentSelectedSheetRange.getValue();
  if (originalStockCode != originalStockCode.toUpperCase()) {
    currentSelectedSheetRange.setValue(originalStockCode.toUpperCase());
  }
  analyzeStockCode(originalStockCode.toUpperCase());

  SpreadsheetApp.flush();
  //Browser.msgBox(stockCode);
}

function analyzeStock2() {
  analyzeStockCode('PSB');
}

/**
 * Analyze the stock transaction and cash dividend price for the passed stock code
 */
function analyzeStockCode(stockCode) {
  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let activeSheet = activeSpreadsheet.getActiveSheet();
  let currentPriceSheet = activeSpreadsheet.getSheetByName(SHEET_CURRENTPRICE);

  let stockSheet = activeSpreadsheet.getSheetByName(stockCode + '-' + SHEET_STOCKANALYZE);
  if (!stockSheet) {
    stockSheet = activeSpreadsheet.insertSheet();
    stockSheet.setName(stockCode + '-' + SHEET_STOCKANALYZE);
    activeSpreadsheet.moveActiveSheet(activeSheet.getIndex() + 1);
    activeSpreadsheet.setActiveSheet(stockSheet);
  }

  // Know percentage of divident give per year
  let stockShareEquallyPerYearRate = getStockShareEquallyPerYear(stockCode);
  let colSettings = loadColumnSettings(SHEET_STOCKANALYZE);

  initializeStockSheet_(stockSheet, colSettings, 2);
  let loadResult = loadAllTransactions_(stockCode, stockSheet, colSettings);
  loadResult.StockShareEquallyPerYearRate = stockShareEquallyPerYearRate;
  computeCurrentQuantiy_(stockSheet, colSettings, loadResult);
  computeCurrentAveragePricePerShare_(stockCode, stockSheet, colSettings, loadResult);

  // write if the buy, sell, dividend diff and rate
  writeAnalysisRowsFormula_(stockSheet, colSettings, loadResult);
  formatStockSheet_(stockSheet, colSettings, loadResult.StartRow, loadResult.EndRow);

  // Write current price
  // Get shares owned from total
  loadResult.LatestPrice = 0;
  let [rows, columns] = [currentPriceSheet.getLastRow(), currentPriceSheet.getLastColumn()];
  let datasheetCurrentPrice = currentPriceSheet.getRange(ROWINDEX_CURRENTPRICE_FIRSTROW, 1, rows, columns).getValues();
  for (let iRow = 0; iRow < datasheetCurrentPrice.length; iRow++) {
    let currentPriceStockCode = datasheetCurrentPrice[iRow][COLINDEX_CURRENTPRICE_STOCKCODE];
    if (stockCode != currentPriceStockCode) {
      continue;
    }
    console.log('Price Row: ' + iRow);
    loadResult.LatestPrice = datasheetCurrentPrice[iRow][COLINDEX_CURRENTPRICE_LATESTPRICE];
    break;
  }
  console.log('Latest Price: ' + loadResult.LatestPrice);

  // Add aggregate values on top
  writeAnalysisTotalFormula_(stockSheet, colSettings, loadResult);

  // Hide selling details
  //if (loadResult.StockSoldCount == 0) {
    stockSheet.hideColumns(COLINDEX_STOCKANALYZE_SELLQTY + 1);
    stockSheet.hideColumns(COLINDEX_STOCKANALYZE_SELLPRICEPERSHARE + 1);
    stockSheet.hideColumns(COLINDEX_STOCKANALYZE_SELLBUYAMOUNT + 1);
  //}

  // Auto adjust columns
  stockSheet.autoResizeColumns(COLINDEX_STOCKANALYZE_TRANSDATE + 1, COLINDEX_STOCKANALYZE_ANNUALACTUALDIVRATE + 1);
}

/**
 * Initialize the stock analyze sheet
 * - add all the column headers
 */
function initializeStockSheet_(stockSheet, colSettings, rowIndex) {
  stockSheet.getRange(colSettings.TransactionDate + rowIndex).setValue(COLTITLE_STOCKANALYZE_TRANSDATE);
  stockSheet.getRange(colSettings.TransactionType + rowIndex).setValue(COLTITLE_STOCKANALYZE_TRANSTYPE);

  stockSheet.getRange(colSettings.BuyQuantity + rowIndex).setValue(COLTITLE_STOCKANALYZE_BUYQTY);
  stockSheet.getRange(colSettings.BuyPricePerShare + rowIndex).setValue(COLTITLE_STOCKANALYZE_BUYPRICEPERSHARE);
  stockSheet.getRange(colSettings.NetBuyAmount + rowIndex).setValue(COLTITLE_STOCKANALYZE_NETBUYAMOUNT);

  stockSheet.getRange(colSettings.SellQuantity + rowIndex).setValue(COLTITLE_STOCKANALYZE_SELLQTY);
  stockSheet.getRange(colSettings.SellPricePerShare + rowIndex).setValue(COLTITLE_STOCKANALYZE_SELLPRICEPERSHARE);
  stockSheet.getRange(colSettings.NetSellAmount + rowIndex).setValue(COLTITLE_STOCKANALYZE_NETSELLAMOUNT);

  stockSheet.getRange(colSettings.DividendPerShare + rowIndex).setValue(COLTITLE_STOCKANALYZE_DIVPERSHARE);
  stockSheet.getRange(colSettings.NetDividendAmount + rowIndex).setValue(COLTITLE_STOCKANALYZE_NETDIVAMOUNT);
  stockSheet.getRange(colSettings.ClosingPricePerShare + rowIndex).setValue(COLTITLE_STOCKANALYZE_CLOSINGPRICEPERSHARE);
  stockSheet.getRange(colSettings.DividendRate + rowIndex).setValue(COLTITLE_STOCKANALYZE_DIVRATE);

  stockSheet.getRange(colSettings.CurrentQuantity + rowIndex).setValue(COLTITLE_STOCKANALYZE_CURRQTY);
  stockSheet.getRange(colSettings.AveragePricePerShare + rowIndex).setValue(COLTITLE_STOCKANALYZE_AVGPRICEPERSHARE);

  stockSheet.getRange(colSettings.BuyVsAveragePrice + rowIndex).setValue(COLTITLE_STOCKANALYZE_BUYVSAVEPRICE);
  stockSheet.getRange(colSettings.BuyVsAveragePriceDiff + rowIndex).setValue(COLTITLE_STOCKANALYZE_BUYVSAVEPRICEDIFF);
  stockSheet.getRange(colSettings.SellVsAveragePrice + rowIndex).setValue(COLTITLE_STOCKANALYZE_SELLVSAVEPRICE);
  stockSheet.getRange(colSettings.SellVsAveragePriceDiff + rowIndex).setValue(COLTITLE_STOCKANALYZE_SELLVSAVEPRICEDIFF);
  stockSheet.getRange(colSettings.ActualDividendRate + rowIndex).setValue(COLTITLE_STOCKANALYZE_ACTUALDIVRATE);
  stockSheet.getRange(colSettings.DividendRateDiff + rowIndex).setValue(COLTITLE_STOCKANALYZE_DIVRATEDIFF);
  stockSheet.getRange(colSettings.AnnualActualDivRate + rowIndex).setValue(COLTITLE_STOCKANALYZE_ANNUALACTUALDIVRATE);
}

/**
 * Format the analyze sheet after wrigint all the value
 * - set the style of all title
 * - change the font
 * - change style of each column
 * - set a border for the whole table
 * - resize the whole table column to minimize usage
 */
function formatStockSheet_(stockSheet, colSettings, startRow, endRow) {
  let cellRange = '';

  // title
  cellRange = colSettings.TransactionDate + (startRow - 1) + ':' + colSettings.AnnualActualDivRate + (startRow - 1);
  stockSheet.getRange(cellRange).setFontWeight('bold').setFontSize(8);

  // All font
  cellRange = colSettings.TransactionDate + (startRow - 2) + ':' + colSettings.AnnualActualDivRate + endRow;
  stockSheet.getRange(cellRange).setFontFamily('Arial Narrow');

  cellRange = colSettings.TransactionDate + startRow + ':' + colSettings.TransactionDate + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_TRANSDATE);
  cellRange = colSettings.TransactionType + startRow + ':' + colSettings.TransactionType + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_TRANSTYPE);

  cellRange = colSettings.BuyQuantity + startRow + ':' + colSettings.BuyQuantity + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_BUYQTY);
  cellRange = colSettings.BuyPricePerShare + startRow + ':' + colSettings.BuyPricePerShare + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_BUYPRICEPERSHARE);
  cellRange = colSettings.NetBuyAmount + startRow + ':' + colSettings.NetBuyAmount + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_NETBUYAMOUNT);

  cellRange = colSettings.SellQuantity + startRow + ':' + colSettings.SellQuantity + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_SELLQTY);
  cellRange = colSettings.SellPricePerShare + startRow + ':' + colSettings.SellPricePerShare + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_SELLPRICEPERSHARE);
  cellRange = colSettings.NetSellAmount + startRow + ':' + colSettings.NetSellAmount + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_NETSELLAMOUNT);

  cellRange = colSettings.DividendPerShare + startRow + ':' + colSettings.DividendPerShare + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_DIVPERSHARE);
  cellRange = colSettings.NetDividendAmount + startRow + ':' + colSettings.NetDividendAmount + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_NETDIVAMOUNT);
  cellRange = colSettings.ClosingPricePerShare + startRow + ':' + colSettings.ClosingPricePerShare + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_CLOSINGPRICEPERSHARE);
  cellRange = colSettings.DividendRate + startRow + ':' + colSettings.DividendRate + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_DIVRATE);

  cellRange = colSettings.CurrentQuantity + startRow + ':' + colSettings.CurrentQuantity + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_CURRQTY);
  cellRange = colSettings.AveragePricePerShare + startRow + ':' + colSettings.AveragePricePerShare + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_AVGPRICEPERSHARE);

  cellRange = colSettings.BuyVsAveragePrice + startRow + ':' + colSettings.BuyVsAveragePrice + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_BUYVSAVEPRICE);
  cellRange = colSettings.BuyVsAveragePriceDiff + startRow + ':' + colSettings.BuyVsAveragePriceDiff + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_BUYVSAVEPRICEDIFF);
  cellRange = colSettings.SellVsAveragePrice + startRow + ':' + colSettings.SellVsAveragePrice + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_SELLVSAVEPRICE);
  cellRange = colSettings.SellVsAveragePriceDiff + startRow + ':' + colSettings.SellVsAveragePriceDiff + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_SELLVSAVEPRICEDIFF);
  cellRange = colSettings.ActualDividendRate + startRow + ':' + colSettings.ActualDividendRate + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_ACTUALDIVRATE);
  cellRange = colSettings.DividendRateDiff + startRow + ':' + colSettings.DividendRateDiff + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_DIVRATEDIFF);
  cellRange = colSettings.AnnualActualDivRate + (startRow - 2) + ':' + colSettings.AnnualActualDivRate + endRow;
  stockSheet.getRange(cellRange).setNumberFormat(COLFORMAT_STOCKANALYZE_ANNUALACTUALDIVRATE);

  // Border
  cellRange = colSettings.TransactionDate + (startRow - 1) + ':' + colSettings.AnnualActualDivRate + endRow;
  stockSheet.getRange(cellRange).setBorder(true, true, true, true, true, true);
}

/**
 * Load all stock transaction (buy or sell) and cash dividend for a certain stock
 */
function loadAllTransactions_(stockCode, stockSheet, colSettings) {

  let sheetStockTrans = SpreadsheetApp.getActive().getSheetByName(SHEET_STOCKTRANS);
  let [rows, columns] = [sheetStockTrans.getLastRow(), sheetStockTrans.getLastColumn()];
  let dataStockTrans = sheetStockTrans.getRange(2, 1, rows, columns).getValues();

  let sheetCashDividend = SpreadsheetApp.getActive().getSheetByName(SHEET_CASHDIV);
  [rows, columns] = [sheetCashDividend.getLastRow(), sheetCashDividend.getLastColumn()];
  let dataCashDividend = sheetCashDividend.getRange(3, 1, rows, columns).getValues();

  let sheetProjDividend = SpreadsheetApp.getActive().getSheetByName(SHEET_PROJDIV);
  [rows, columns] = [sheetProjDividend.getLastRow(), sheetProjDividend.getLastColumn()];
  let dataProjDividend = sheetProjDividend.getRange(3, 1, rows, columns).getValues();

  let foundProjDividendIndex = -1;
  let stockAnalyzeRow = 3;
  let currCashDividendIndex = 0;
  let cashDividendCount = 0;
  let stockSoldCount = 0;

  let loadResult = {};
  loadResult['LatestDivPerShare'] = null;

  // Get the projected cash dividend of the stock
  for (let projDividendIndex = 0; projDividendIndex < dataProjDividend.length; projDividendIndex++) {
    let stockCode3 = dataProjDividend[projDividendIndex][COLINDEX_PROJDIV_STOCKCODE];
    if (stockCode3 != stockCode) {
      continue;
    }
    foundProjDividendIndex = projDividendIndex;
    break;
  }

  for (let stockTransIndex = 0; stockTransIndex < dataStockTrans.length; stockTransIndex++) {
    let stockCode1 = dataStockTrans[stockTransIndex][COLINDEX_STOCKTRANS_STOCKCODE];
    if (stockCode1 != stockCode) {
      continue;
    }

    // Before adding the stock trans, check first if there are cash div before it, not yet added
    let currTransDate = new Date(dataStockTrans[stockTransIndex][COLINDEX_STOCKTRANS_TRANSDATE]);
    for (let cashDividendIndex = currCashDividendIndex; cashDividendIndex < dataCashDividend.length; cashDividendIndex++) {
      let stockCode2 = dataCashDividend[cashDividendIndex][COLINDEX_CASHDIV_STOCKCODE];
      if (stockCode2 != stockCode) {
        currCashDividendIndex++;
        continue;
      }

      let exDividendDate = new Date(dataCashDividend[cashDividendIndex][COLINDEX_CASHDIV_EXDIVDATE]);
      console.log('CDIndex: ' + cashDividendIndex + ' TransactDate: ' + currTransDate + ' < ExDivDate:' + exDividendDate);
      if (currTransDate.getTime() < exDividendDate.getTime()) {

        // Before adding cash dividend, check if there is a projected cash dividend before it
        if (cashDividendCount == 0 && foundProjDividendIndex > -1) {
          if (loadResult['LatestDivPerShare'] == null) {
            loadResult['LatestDivPerShare'] = dataProjDividend[foundProjDividendIndex][COLINDEX_PROJDIV_DIVPERSHARE];
          }
          let projExDividendDate = new Date(dataProjDividend[foundProjDividendIndex][COLINDEX_PROJDIV_EXDIVDATE]);
          console.log('Proj Div, Cash Div ExDate: ' + currTransDate + ' < Proj Ex Date: ' + projExDividendDate);
          if (exDividendDate.getTime() < projExDividendDate.getTime()) {
            addProjectedCashDividendRow_(dataProjDividend, foundProjDividendIndex, stockSheet, colSettings, stockAnalyzeRow);
            stockAnalyzeRow++;
            foundProjDividendIndex = -1;
          }
        }

        if (loadResult['LatestDivPerShare'] == null) {
          loadResult['LatestDivPerShare'] = dataCashDividend[cashDividendIndex][COLINDEX_CASHDIV_DIVPERSHARE];
        }
        console.log('Adding Div row');
        addCashDividendRow_(dataCashDividend, cashDividendIndex, stockSheet, colSettings, stockAnalyzeRow);
        stockAnalyzeRow++;
        cashDividendCount++;
        currCashDividendIndex++;
        continue;
      } else {
        break;
      }
    }

    // Next matched cash dividend for the code is for next loop, after stock trans have been added
    // Check projected dividend
    if (cashDividendCount == 0 && foundProjDividendIndex > -1) {
      if (loadResult['LatestDivPerShare'] == null) {
        loadResult['LatestDivPerShare'] = dataProjDividend[foundProjDividendIndex][COLINDEX_PROJDIV_DIVPERSHARE];
      }
      let projExDividendDate = new Date(dataProjDividend[foundProjDividendIndex][COLINDEX_PROJDIV_EXDIVDATE]);
      console.log('Proj Div, Transact Date: ' + currTransDate + ' < Proj Ex Date: ' + projExDividendDate);
      if (currTransDate.getTime() < projExDividendDate.getTime()) {
        addProjectedCashDividendRow_(dataProjDividend, foundProjDividendIndex, stockSheet, colSettings, stockAnalyzeRow);
        stockAnalyzeRow++;
        foundProjDividendIndex = -1;
      }
    }

    let addedTransactRow = addStockTransactionRow_(dataStockTrans, stockTransIndex, stockSheet, colSettings, stockAnalyzeRow);
    currTransDate = new Date(addedTransactRow.TransactionDate);
    let stockTransType = dataStockTrans[stockTransIndex][COLINDEX_STOCKTRANS_TRANSTYPE];
    if (stockTransType == TRANSTYPE_SOLDSHARES) {
      stockSoldCount++;
    }
    stockAnalyzeRow++;
  }

  [rows, columns] = [stockSheet.getLastRow(), COLINDEX_STOCKANALYZE_DIVRATEDIFF + 1];
  let dataStock = stockSheet.getRange(1, 1, rows, columns).getValues();

  loadResult['StartRow'] = 3;
  loadResult['EndRow'] = stockAnalyzeRow - 1;
  loadResult['StockData'] = dataStock;
  loadResult['StockSoldCount'] = stockSoldCount;
  return loadResult;
}

/**
 * add a stock transaction row data
 */
function addStockTransactionRow_(dataStockTrans, stockTransIndex, stockSheet, colSettings, stockAnalyzeRow) {

  let cellRange = '';
  let stockTransDate = dataStockTrans[stockTransIndex][COLINDEX_STOCKTRANS_TRANSDATE];
  cellRange = colSettings.TransactionDate + stockAnalyzeRow;
  stockSheet.getRange(cellRange).setValue(stockTransDate);

  let stockTransType = dataStockTrans[stockTransIndex][COLINDEX_STOCKTRANS_TRANSTYPE];
  cellRange = colSettings.TransactionType + stockAnalyzeRow;
  stockSheet.getRange(cellRange).setValue(stockTransType);

  let qtyAmount = Math.abs(dataStockTrans[stockTransIndex][COLINDEX_STOCKTRANS_QUANTITY]);
  if (isBuyTransactionType(stockTransType) || stockTransType == TRANSTYPE_STOCKDIVIDEND) {
    cellRange = colSettings.BuyQuantity + stockAnalyzeRow;
    stockSheet.getRange(cellRange).setValue(qtyAmount);

    let buyPricePerShare = dataStockTrans[stockTransIndex][COLINDEX_STOCKTRANS_BUYPRICEPERSHARE];
    cellRange = colSettings.BuyPricePerShare + stockAnalyzeRow;
    stockSheet.getRange(cellRange).setValue(buyPricePerShare);

    let netBuyAmt = dataStockTrans[stockTransIndex][COLINDEX_STOCKTRANS_NETBUYAMT];
    cellRange = colSettings.NetBuyAmount + stockAnalyzeRow;
    stockSheet.getRange(cellRange).setValue(netBuyAmt);
 } else if (stockTransType == TRANSTYPE_SOLDSHARES) {
    cellRange = colSettings.SellQuantity + stockAnalyzeRow;
    stockSheet.getRange(cellRange).setValue(qtyAmount);

    let sellPricePerShare = dataStockTrans[stockTransIndex][COLINDEX_STOCKTRANS_SELLPRICEPERSHARE];
    cellRange = colSettings.SellPricePerShare + stockAnalyzeRow;
    stockSheet.getRange(cellRange).setValue(sellPricePerShare);

    let netSellAmt = Math.abs(dataStockTrans[stockTransIndex][COLINDEX_STOCKTRANS_NETSELLAMT]);
    cellRange = colSettings.NetSellAmount + stockAnalyzeRow;
    stockSheet.getRange(cellRange).setValue(netSellAmt);
  }
 
  return {
    'TransactionDate': stockTransDate,
    'TransactionType': stockTransType,
  }
}

/**
 * Add a cash dividend data
 */
function addCashDividendRow_(dataCashDividend, cashDividendIndex, stockSheet, colSettings, stockAnalyzeRow) {
  let cellRange = '';
  let stockExDividendDate = dataCashDividend[cashDividendIndex][COLINDEX_CASHDIV_EXDIVDATE];
  cellRange = colSettings.TransactionDate + stockAnalyzeRow;
  stockSheet.getRange(cellRange).setValue(stockExDividendDate);

  cellRange = colSettings.TransactionType + stockAnalyzeRow;
  stockSheet.getRange(cellRange).setValue(TRANSTYPE_CASHDIVIDEND);

  let dividendPerShare = dataCashDividend[cashDividendIndex][COLINDEX_CASHDIV_DIVPERSHARE];
  cellRange = colSettings.DividendPerShare + stockAnalyzeRow;
  stockSheet.getRange(cellRange).setValue(dividendPerShare);

  let netAmount = dataCashDividend[cashDividendIndex][COLINDEX_CASHDIV_NETAMT];
  cellRange = colSettings.NetDividendAmount + stockAnalyzeRow;
  stockSheet.getRange(cellRange).setValue(netAmount);

  let closingPrice = dataCashDividend[cashDividendIndex][COLINDEX_CASHDIV_PRICEPERSHARE];
  cellRange = colSettings.ClosingPricePerShare + stockAnalyzeRow;
  stockSheet.getRange(cellRange).setValue(closingPrice);

  let divRate = dataCashDividend[cashDividendIndex][COLINDEX_CASHDIV_DIVRATE];
  cellRange = colSettings.DividendRate + stockAnalyzeRow;
  stockSheet.getRange(cellRange).setValue(divRate);
}

/**
 * Add a projected cash dividend data
 */
function addProjectedCashDividendRow_(dataProjDividend, foundProjDividendIndex, stockSheet, colSettings, stockAnalyzeRow) {
  let cellRange = '';
  let stockExDividendDate = dataProjDividend[foundProjDividendIndex][COLINDEX_PROJDIV_EXDIVDATE];
  cellRange = colSettings.TransactionDate + stockAnalyzeRow;
  stockSheet.getRange(cellRange).setValue(stockExDividendDate);

  cellRange = colSettings.TransactionType + stockAnalyzeRow;
  stockSheet.getRange(cellRange).setValue(TRANSTYPE_CASHDIVIDEND);

  let dividendPerShare = dataProjDividend[foundProjDividendIndex][COLINDEX_PROJDIV_DIVPERSHARE];
  cellRange = colSettings.DividendPerShare + stockAnalyzeRow;
  stockSheet.getRange(cellRange).setValue(dividendPerShare);

  let netAmount = dataProjDividend[foundProjDividendIndex][COLINDEX_PROJDIV_NETAMT];
  cellRange = colSettings.NetDividendAmount + stockAnalyzeRow;
  stockSheet.getRange(cellRange).setValue(netAmount);

  let closingPrice = dataProjDividend[foundProjDividendIndex][COLINDEX_PROJDIV_PRICEPERSHARE];
  cellRange = colSettings.ClosingPricePerShare + stockAnalyzeRow;
  stockSheet.getRange(cellRange).setValue(closingPrice);

  let divRate = dataProjDividend[foundProjDividendIndex][COLINDEX_PROJDIV_DIVRATE];
  cellRange = colSettings.DividendRate + stockAnalyzeRow;
  stockSheet.getRange(cellRange).setValue(divRate);
}

/**
 * Loop all the stock transaction and cash dividend written and compute the stock quantity owned
 */
function computeCurrentQuantiy_(stockSheet, colSettings, loadResult) {
  let currentQuantity = 0;
  for (let rowIndex = loadResult.StockData.length - 1; rowIndex >= loadResult.StartRow - 1; rowIndex--) {
    let transactType = loadResult.StockData[rowIndex][COLINDEX_STOCKANALYZE_TRANSTYPE];
    
    if (isBuyTransactionType(transactType) || transactType == TRANSTYPE_STOCKDIVIDEND) {
      currentQuantity += loadResult.StockData[rowIndex][COLINDEX_STOCKANALYZE_BUYQTY];
    } else if (transactType == TRANSTYPE_SOLDSHARES) {
      currentQuantity -= loadResult.StockData[rowIndex][COLINDEX_STOCKANALYZE_SELLQTY];
    }

    loadResult.StockData[rowIndex][COLINDEX_STOCKANALYZE_CURRQTY] = currentQuantity;
    stockSheet.getRange(colSettings.CurrentQuantity + (rowIndex + 1)).setValue(currentQuantity);
  }
}

/**
 * Compute the average buy price per share of the stock for each row
 */
function computeCurrentAveragePricePerShare_(stockCode, stockSheet, colSettings, loadResult) {
  for (let rowIndex1 = loadResult.StockData.length - 1; rowIndex1 >= loadResult.StartRow - 1; rowIndex1--) {
    let currentQuantity = loadResult.StockData[rowIndex1][COLINDEX_STOCKANALYZE_CURRQTY];
    if (currentQuantity == 0) {
      // If sold all, get previous qty
      currentQuantity = loadResult.StockData[rowIndex1 + 1][COLINDEX_STOCKANALYZE_CURRQTY];
    }
    let transactionDate = loadResult.StockData[rowIndex1][COLINDEX_STOCKANALYZE_TRANSDATE];

    let averagePricePerShare = 0;
    let totalBuyQty = 0;
    let totalNetBuyAmount = 0;
    for (let rowIndex2 = rowIndex1; rowIndex2 < loadResult.StockData.length; rowIndex2++) {

      let transactType = loadResult.StockData[rowIndex2][COLINDEX_STOCKANALYZE_TRANSTYPE];
      if (isBuyTransactionType(transactType) || transactType == TRANSTYPE_STOCKDIVIDEND) {
        let rowBuyQuantity = loadResult.StockData[rowIndex2][COLINDEX_STOCKANALYZE_BUYQTY];
        let rowNetBuyAmount = loadResult.StockData[rowIndex2][COLINDEX_STOCKANALYZE_NETBUYAMOUNT];
        if (isNaN(rowNetBuyAmount) || rowNetBuyAmount == '') {
          rowNetBuyAmount = 0;
        }

        if (totalBuyQty + rowBuyQuantity <= currentQuantity) {
          totalBuyQty += rowBuyQuantity;
          totalNetBuyAmount += rowNetBuyAmount;
        } else {
          let partRowBuyQuantity = currentQuantity - totalBuyQty;
          totalBuyQty += partRowBuyQuantity;
          totalNetBuyAmount += (partRowBuyQuantity / rowBuyQuantity) * rowNetBuyAmount;
        }

        if (totalBuyQty == currentQuantity) {
          averagePricePerShare = totalNetBuyAmount / totalBuyQty;
          break;
        }
      }
    } // end for
    
    loadResult.StockData[rowIndex1][COLINDEX_STOCKANALYZE_AVGPRICEPERSHARE] = averagePricePerShare;
    stockSheet.getRange(colSettings.AveragePricePerShare + (rowIndex1 + 1)).setValue(averagePricePerShare);

  } // end for
}

/**
 * Write the analysis formula for each stock
 */
function writeAnalysisRowsFormula_(stockSheet, colSettings, loadResult) {
  let lastDivPerShare = loadResult.LatestDivPerShare;
  for (let rowIndex = 0; rowIndex < loadResult.StockData.length; rowIndex++) {
    let transactType = loadResult.StockData[rowIndex][COLINDEX_STOCKANALYZE_TRANSTYPE];
    
    if (isBuyTransactionType(transactType) || transactType == TRANSTYPE_STOCKDIVIDEND) {
      let baseFormula = colSettings.BuyPricePerShare + (rowIndex + 1) + '-' + colSettings.AveragePricePerShare + (rowIndex + 2);
      let buyVsAveragePriceFormula = '=' + baseFormula;
      let diffRate = '=(' + baseFormula + ')/' + colSettings.AveragePricePerShare + (rowIndex + 2);

      if (rowIndex < loadResult.StockData.length - 1) {
        stockSheet.getRange(colSettings.BuyVsAveragePrice + (rowIndex + 1)).setFormula(buyVsAveragePriceFormula);
        stockSheet.getRange(colSettings.BuyVsAveragePriceDiff + (rowIndex + 1)).setFormula(diffRate);
        let rawDiffValue = 
          loadResult.StockData[rowIndex][COLINDEX_STOCKANALYZE_BUYPRICEPERSHARE] - 
            loadResult.StockData[rowIndex + 1][COLINDEX_STOCKANALYZE_AVGPRICEPERSHARE];
        if (rawDiffValue < 0) {
          stockSheet.getRange(colSettings.BuyVsAveragePriceDiff + (rowIndex + 1)).setFontColor('Green');
        } else {
          stockSheet.getRange(colSettings.BuyVsAveragePriceDiff + (rowIndex + 1)).setFontColor('Red');
        }
      }
    } else if (transactType == TRANSTYPE_SOLDSHARES) {
      let baseFormula = colSettings.SellPricePerShare + (rowIndex + 1) + '-' + colSettings.AveragePricePerShare + (rowIndex + 2);
      let buyVsAveragePriceFormula = '=' + baseFormula;
      let diffRate = '=(' + baseFormula + ')/' + colSettings.AveragePricePerShare + (rowIndex + 2);

      stockSheet.getRange(colSettings.SellVsAveragePrice + (rowIndex + 1)).setFormula(buyVsAveragePriceFormula);
      stockSheet.getRange(colSettings.SellVsAveragePriceDiff + (rowIndex + 1)).setFormula(diffRate);

      let rawDiffValue = 
        loadResult.StockData[rowIndex][COLINDEX_STOCKANALYZE_SELLPRICEPERSHARE] - 
          loadResult.StockData[rowIndex + 1][COLINDEX_STOCKANALYZE_AVGPRICEPERSHARE];
      if (rawDiffValue < 0) {
        stockSheet.getRange(colSettings.SellVsAveragePriceDiff + (rowIndex + 1)).setFontColor('Red');
      } else {
        stockSheet.getRange(colSettings.SellVsAveragePriceDiff + (rowIndex + 1)).setFontColor('Green');
      }
    } else if (transactType == TRANSTYPE_CASHDIVIDEND) {
      let divRateFormula = '=' + colSettings.DividendPerShare + (rowIndex + 1) + '/' + colSettings.AveragePricePerShare + (rowIndex + 1);
      let diffRate = '=' + colSettings.ActualDividendRate + (rowIndex + 1) + '-' + colSettings.DividendRate + (rowIndex + 1);

      stockSheet.getRange(colSettings.ActualDividendRate + (rowIndex + 1)).setFormula(divRateFormula);
      stockSheet.getRange(colSettings.DividendRateDiff + (rowIndex + 1)).setFormula(diffRate);

      lastDivPerShare = parseFloat(loadResult.StockData[rowIndex][COLINDEX_STOCKANALYZE_DIVPERSHARE]);
      let actualDivRate = parseFloat(loadResult.StockData[rowIndex][COLINDEX_STOCKANALYZE_DIVPERSHARE]) /
        parseFloat(loadResult.StockData[rowIndex][COLINDEX_STOCKANALYZE_AVGPRICEPERSHARE]);
      //console.log(' = ' + actualDivRate + ' - ' + loadResult.StockData[rowIndex][COLINDEX_STOCKANALYZE_DIVRATE]);
      let rawDiffValue = 
        actualDivRate - 
          parseFloat(loadResult.StockData[rowIndex][COLINDEX_STOCKANALYZE_DIVRATE]);
      //console.log('dd' + rawDiffValue);
      if (rawDiffValue < 0.0) {
        stockSheet.getRange(colSettings.DividendRateDiff + (rowIndex + 1)).setFontColor('Red');
      } else {
        stockSheet.getRange(colSettings.DividendRateDiff + (rowIndex + 1)).setFontColor('Green');
      }
    }

    if (transactType != TRANSTYPE_CASHDIVIDEND && transactType != COLTITLE_STOCKANALYZE_TRANSTYPE &&
      lastDivPerShare != null && 
      loadResult.StockShareEquallyPerYearRate != null) {
      // Add running actual div rate
      let divRateFormula = '=' + lastDivPerShare + '/' + colSettings.AveragePricePerShare + (rowIndex + 1);
      stockSheet.getRange(colSettings.ActualDividendRate + (rowIndex + 1)).setFormula(divRateFormula);
    }

    if (loadResult.StockShareEquallyPerYearRate != null && transactType != COLTITLE_STOCKANALYZE_TRANSTYPE) {
      let annualDivRateFormula = '=' + colSettings.ActualDividendRate + (rowIndex + 1) + '/' + loadResult.StockShareEquallyPerYearRate;
      stockSheet.getRange(colSettings.AnnualActualDivRate + (rowIndex + 1)).setFormula(annualDivRateFormula);
    }

  } // end for
}

/**
 * Write the total row formula
 */
function writeAnalysisTotalFormula_(stockSheet, colSettings, loadResult) {
  // Current Price
  stockSheet.getRange(colSettings.AveragePricePerShare + '1').setValue(loadResult.LatestPrice);
  stockSheet.getRange(colSettings.AveragePricePerShare + '1').setNumberFormat(COLFORMAT_STOCKANALYZE_AVGPRICEPERSHARE);

  // Amount diff of current price to the average buy price if we buy
  let latestAvgBuyPrice = stockSheet.getRange(colSettings.AveragePricePerShare + '3').getValue();
  let baseFormula = colSettings.AveragePricePerShare + '1-' + colSettings.AveragePricePerShare + '3';
  let buyVsAveragePriceFormula = '=' + baseFormula;
  let diffRate = '=(' + baseFormula + ')/' + colSettings.AveragePricePerShare + '3';
  stockSheet.getRange(colSettings.BuyVsAveragePrice + '1').setFormula(buyVsAveragePriceFormula);
  stockSheet.getRange(colSettings.BuyVsAveragePriceDiff + '1').setFormula(diffRate);
  let rawDiffValue = loadResult.LatestPrice - latestAvgBuyPrice;
  if (rawDiffValue < 0) {
    stockSheet.getRange(colSettings.BuyVsAveragePriceDiff + '1').setFontColor('Green');
  } else {
    stockSheet.getRange(colSettings.BuyVsAveragePriceDiff + '1').setFontColor('Red');
  }
  stockSheet.getRange(colSettings.BuyVsAveragePriceDiff + '1').setNumberFormat(COLFORMAT_STOCKANALYZE_BUYVSAVEPRICEDIFF);

  // Amount diff of current price to the average buy price if we sell
  let sellBaseFormula = colSettings.AveragePricePerShare + '1-' + colSettings.AveragePricePerShare + '3';
  let sellVsAveragePriceFormula = '=' + sellBaseFormula;
  let sellDiffRate = '=(' + sellBaseFormula + ')/' + colSettings.AveragePricePerShare + '3';

  stockSheet.getRange(colSettings.SellVsAveragePrice + '1').setFormula(sellVsAveragePriceFormula);
  stockSheet.getRange(colSettings.SellVsAveragePriceDiff + '1').setFormula(sellDiffRate);
  if (rawDiffValue < 0) {
    stockSheet.getRange(colSettings.SellVsAveragePriceDiff + '1').setFontColor('Red');
  } else {
    stockSheet.getRange(colSettings.SellVsAveragePriceDiff + '1').setFontColor('Green');
  }
  stockSheet.getRange(colSettings.SellVsAveragePriceDiff + '1').setNumberFormat(COLFORMAT_STOCKANALYZE_BUYVSAVEPRICEDIFF);

  // The latest div per share rate using current price
  if (loadResult.LatestDivPerShare != null) {
    stockSheet.getRange(colSettings.DividendPerShare + '1').setValue(loadResult.LatestDivPerShare);
    stockSheet.getRange(colSettings.DividendPerShare + '1').setNumberFormat(COLFORMAT_STOCKANALYZE_DIVPERSHARE);

    let dividendNetAmountFormula = '=' + colSettings.DividendPerShare + '1*' + colSettings.CurrentQuantity + '3*0.9';
    stockSheet.getRange(colSettings.NetDividendAmount + '1').setFormula(dividendNetAmountFormula);
    stockSheet.getRange(colSettings.NetDividendAmount + '1').setNumberFormat(COLFORMAT_STOCKANALYZE_NETDIVAMOUNT);

    let dividendRateFormula = '=' + colSettings.DividendPerShare + '1/' + colSettings.AveragePricePerShare + '1';
    stockSheet.getRange(colSettings.DividendRate + '1').setFormula(dividendRateFormula);
    stockSheet.getRange(colSettings.DividendRate + '1').setNumberFormat(COLFORMAT_STOCKANALYZE_DIVRATE);

    let actualDividendRateFormula = '=' + colSettings.DividendPerShare + '1/' + colSettings.AveragePricePerShare + '3';
    stockSheet.getRange(colSettings.ActualDividendRate + '1').setFormula(actualDividendRateFormula);
    stockSheet.getRange(colSettings.ActualDividendRate + '1').setNumberFormat(COLFORMAT_STOCKANALYZE_ACTUALDIVRATE);

    let diffRate = '=' + colSettings.ActualDividendRate + '1-' + colSettings.DividendRate + '1';
    stockSheet.getRange(colSettings.DividendRateDiff + '1').setFormula(diffRate);

    let actualDivRate = parseFloat(loadResult.LatestDivPerShare) / parseFloat(latestAvgBuyPrice);
    let currPriceDivRate = parseFloat(loadResult.LatestDivPerShare) / loadResult.LatestPrice;
    rawDiffValue = actualDivRate - currPriceDivRate;
    if (rawDiffValue < 0.0) {
      stockSheet.getRange(colSettings.DividendRateDiff + '1').setFontColor('Red');
    } else {
      stockSheet.getRange(colSettings.DividendRateDiff + '1').setFontColor('Green');
    }
  }
}

