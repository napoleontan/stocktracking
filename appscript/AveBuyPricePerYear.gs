function computeAveBuyPricePerYear() {
  let stockYearlyTotalInfo = getOwnedStockYearlyTotalInfo_();
  stockYearlyTotalInfo = computeStockYearlyAverageBuyPrice_(stockYearlyTotalInfo);
  //console.log(stockYearlyTotalInfo);
  writeSharesYearlyAveragePrice_(stockYearlyTotalInfo);
}

function computeStockYearlyAverageBuyPrice_(stockYearlyTotalInfo) {

  let firstYear = stockYearlyTotalInfo.FirstYear;
  let lastYear = firstYear;
  for (let keyStockCode in stockYearlyTotalInfo) {
    if (keyStockCode == 'FirstYear') {
      continue;
    }

    for (let keyStockField in stockYearlyTotalInfo[keyStockCode]) {
      let year = parseInt(keyStockField);
      if (isNaN(year)) {
        continue;
      }
      if (year > lastYear) { 
        lastYear = year;
      }

      let yearNoOfStock = stockYearlyTotalInfo[keyStockCode][year].NoOfStock;
      // Even if there is sell transaction for the year if it is all sell no point writing it
      if (parseInt(yearNoOfStock, 10) <= 0) {
        continue;
      }

      let lastTransactionDate = stockYearlyTotalInfo[keyStockCode][year].LastTransactDate;
      if (lastTransactionDate != null && lastTransactionDate != undefined) {
        let nextDateCutOff = year + '-12-31';

        let averagePriceVal = computeAveragePriceByStockCode(keyStockCode, nextDateCutOff, yearNoOfStock);
        console.log(averagePriceVal);
        stockYearlyTotalInfo[keyStockCode][year]['AverageBuyPricePerShare'] = averagePriceVal.AverageBuyPricePerShare;
      }
    } // for stock year
  } // for stock Code
  
  return stockYearlyTotalInfo;
}

function writeSharesYearlyAveragePrice_(stockYearlyTotalInfo) {

  let quantityFormatter = new Intl.NumberFormat('en-US',
    { style:'decimal', minimumFractionDigits: 0 }
  );
  let sharePriceFormatter = new Intl.NumberFormat('en-US',
    { style:'currency',currency:'PHP', minimumFractionDigits: 4 }
  );
  let amountFormatter = new Intl.NumberFormat('en-US',
    { style:'currency',currency:'PHP', minimumFractionDigits: 2 }
  );

  let sheetDivYearlyPerformance = SpreadsheetApp.getActive().getSheetByName(SHEET_YEARLYBUYAMOUNT);
  let [rows, columns] = [sheetDivYearlyPerformance.getLastRow(), sheetDivYearlyPerformance.getLastColumn()];
  if (rows.length > 0) {
    sheetDivYearlyPerformance.getRange(1, 1, rows, columns).clearContent();
    SpreadsheetApp.flush();
  }
  
  let rowIndex = 2;
  let firstYear = stockYearlyTotalInfo.FirstYear;
  let lastYear = firstYear;
  for (let keyStockCode in stockYearlyTotalInfo) {
    if (keyStockCode == 'FirstYear') {
      continue;
    }
    let aStockYearlyTotalInfo = stockYearlyTotalInfo[keyStockCode];

    sheetDivYearlyPerformance.getRange(columnToLetter(COLINDEX_YEARLYBUYAMOUNT_STOCKCODE + 1) + rowIndex).setValue(keyStockCode);
    
    let lastYearQuantity = 0;
    let lastYearAverageBuyPricePerShare = 0;
    for (let keyStockField in stockYearlyTotalInfo[keyStockCode]) {
      let year = parseInt(keyStockField);
      if (!isNaN(year)) {
        let colIndex = ((year - firstYear) * WriteDividendYearlyPerformanceYearColNum) + 2;
        if (year > lastYear) { 
          lastYear = year;
        }

        sheetDivYearlyPerformance.getRange(columnToLetter(colIndex) + '1').setValue(year);
        let yearNoOfStock = stockYearlyTotalInfo[keyStockCode][year].NoOfStock;
        // Even if there is sell transaction for the year if it is all sell no point writing it
        if (parseInt(yearNoOfStock, 10) <= 0) {
          continue;
        }

        let averageBuyPricePerShare = stockYearlyTotalInfo[keyStockCode][year].AverageBuyPricePerShare;
        if (averageBuyPricePerShare === undefined) {
          if (lastYearQuantity == yearNoOfStock) {
            averageBuyPricePerShare = lastYearAverageBuyPricePerShare;
          }
        } else {
          lastYearQuantity = yearNoOfStock;
          lastYearAverageBuyPricePerShare = averageBuyPricePerShare;
        }

        if (averageBuyPricePerShare !== undefined) {
          let marketValueRange = sheetDivYearlyPerformance.getRange(columnToLetter(colIndex) + rowIndex);
          let marketPrice = yearNoOfStock * averageBuyPricePerShare;
          marketValueRange.setValue(marketPrice);

          let comment = '\Number of Share: ' + quantityFormatter.format(yearNoOfStock);
          comment += '\nAve Buy Price Per Share: ' + sharePriceFormatter.format(averageBuyPricePerShare); 
          comment += '\nNet Amount: ' + amountFormatter.format(marketPrice);
          
          marketValueRange.setComment(comment);
        }
      }
    }
    rowIndex++;
  }
  
  // Format the column cells
  for (let year = firstYear; year <= lastYear; year++) {
    let colIndex = ((year - firstYear) * WriteDividendYearlyPerformanceYearColNum) + 2;
    let marketPriceColLetter = columnToLetter(colIndex);

    // Year title
    let yearCellRange = sheetDivYearlyPerformance.getRange(marketPriceColLetter + '1');
    yearCellRange.setNumberFormat("###00");

    // Market Price
    let pricePerShareColRange = sheetDivYearlyPerformance.getRange(marketPriceColLetter + '2:' + marketPriceColLetter);
    pricePerShareColRange.setNumberFormat("#,##00.00");

    // Add summation value
    let formulaSummation = "=SUM(" + marketPriceColLetter + "2:" + marketPriceColLetter + (rowIndex - 1) + ")";
    sheetDivYearlyPerformance.getRange(marketPriceColLetter + rowIndex).setFormula(formulaSummation);
  }
}
