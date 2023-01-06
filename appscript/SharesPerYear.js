/**
 * Computes the market price of each owned stock for every year since bought
 * Steps:
 * 1. Using StockTransaction ledger, read from oldest record to newest
 * 2. Retrieve all the year end price for all owned stock in year owned
 * - If the year has not ended yet, get latest price
 * 3. Write the shares in a pivot like table to see market value of owned stock yearly
 * 4. Write the stock year ending owned quantity to be written in special column of StockTransaction.
 * 
 * If getAllStockYearlyEndingPrice_ param is true, only get the latest year price for performance.
 * To get all year price again change param to false
 */
function computeSharesPerYear() {

  let stockYearlyTotalInfo = getOwnedStockYearlyTotalInfo_();
  getAllStockYearlyEndingPrice_(stockYearlyTotalInfo, true);
  //console.log(stockYearlyTotalInfo);
  writeSharesYearlyPerformance_(stockYearlyTotalInfo);

  writeYearEndStockQuantity_(stockYearlyTotalInfo);
}

/**
 * Get the owned stock for each year using the StockTransaction sheet info
 * - The records are read from oldest (last) record to the newest (first)
 * - The first year where stock is bought is marked here for use later
 * - Counts the stock owned by end of each year and overall final stock owned is set also
 * A secondary loop to each stock owned is done to fill up years were we owned a stock but had no transaction (BUY or SELL or Share Div)
 */
function getOwnedStockYearlyTotalInfo_() {
  let sheetStockTrans = SpreadsheetApp.getActive().getSheetByName(SHEET_STOCKTRANS);
  let [rows, columns] = [sheetStockTrans.getLastRow(), sheetStockTrans.getLastColumn()];
  let datasheetStockTrans = sheetStockTrans.getRange(3, 1, rows, columns).getValues();
  //console.log(datasheetStockTrans);

  let stockTotal = {};

  // Loop the stock gather data
  let thisYear = (new Date()).getFullYear();
  let firstYear = thisYear;
  for (let iRow = datasheetStockTrans.length - 1; iRow >= 0; iRow--) {
    let stockCode = datasheetStockTrans[iRow][COLINDEX_STOCKTRANS_STOCKCODE];
    let transactDate = new Date(datasheetStockTrans[iRow][COLINDEX_STOCKTRANS_TRANSDATE]);
    let quantity = datasheetStockTrans[iRow][COLINDEX_STOCKTRANS_QUANTITY];

    if (stockCode == '') {
      continue;
    }

    if (stockTotal['FirstYear'] === undefined) {
      stockTotal['FirstYear'] = transactDate.getFullYear();
      firstYear = transactDate.getFullYear();
    }

    if (!(stockCode in stockTotal)) {
      stockTotal[stockCode] = {};
      stockTotal[stockCode]['NoOfStock'] = 0;
      stockTotal[stockCode]['CurrentYear'] = null;
    }

    if (stockTotal[stockCode]['CurrentYear'] == null) {
      // If first time seeing stock
      stockTotal[stockCode][transactDate.getFullYear()] = {};
      stockTotal[stockCode][transactDate.getFullYear()]['NoOfStock'] = quantity;
      stockTotal[stockCode]['CurrentYear'] = transactDate.getFullYear();
    } else {

      if (stockTotal[stockCode]['CurrentYear'] == transactDate.getFullYear()) {
        // Year did not change
        stockTotal[stockCode][transactDate.getFullYear()]['NoOfStock'] += quantity;
      } else {
        // Year changed
        stockTotal[stockCode][transactDate.getFullYear()] = {};
        stockTotal[stockCode][transactDate.getFullYear()]['NoOfStock'] = stockTotal[stockCode]['NoOfStock'] + quantity;
        stockTotal[stockCode]['CurrentYear'] = transactDate.getFullYear();
      }
    }
    stockTotal[stockCode][transactDate.getFullYear()]['LastTransactDate'] = Utilities.formatDate(transactDate, 'Asia/Manila', 'yyyy-MM-dd');

    // Do tally of total as last step, since previous value is needed in computing the change of year
    stockTotal[stockCode]['NoOfStock'] += quantity;
  }

  // Fillup the middle part of year where no stock was bought
  for (let keyStockCode in stockTotal) {
    if (keyStockCode == 'FirstYear') {
      continue;
    }

    let currentStockTotal = 0;
    for (let year = firstYear; year <= thisYear; year++) {
      if (year in stockTotal[keyStockCode]) {
        currentStockTotal = stockTotal[keyStockCode][year].NoOfStock;
      } else if (currentStockTotal > 0) {
        // If missing, check last stockTotal if > 0
        // Else just ignore, but do not end loop since can have buy of stock on succeeding year
        stockTotal[keyStockCode][year] = {};
        stockTotal[keyStockCode][year]['NoOfStock'] = stockTotal[keyStockCode][year - 1]['NoOfStock'];
      }
    }
  }

  //console.log(stockTotal);
  return stockTotal;
}

/**
 * Retrieve end price of stock owned for each year the stock is owned.
 * If the current year is still not ended the last price is retrieved
 * - Used the REST api of the PSE Edge
 * - There is a limitation for security that have no id
 * - If onlyForThisYear set, only stock price for this year is retrieved for performance purpose
 *   - change this flag to false, if we want to retrieve all owned years
 */
function getAllStockYearlyEndingPrice_(stockYearlyTotalInfo, onlyForThisYear) {

  let sheetCurrentPrice = SpreadsheetApp.getActive().getSheetByName(SHEET_CURRENTPRICE);
  let [rows, columns] = [sheetCurrentPrice.getLastRow(), sheetCurrentPrice.getLastColumn()];
  let dataCurrentPrice = sheetCurrentPrice.getRange(ROWINDEX_CURRENTPRICE_FIRSTROW, 1, rows, columns).getValues();
  //console.log(dataCurrentPrice);

  let thisYear = (new Date()).getFullYear();

  // Get company id, security id and latest price info
  for (let iRow = 0; iRow < dataCurrentPrice.length; iRow++) {
    let stockCode = dataCurrentPrice[iRow][COLINDEX_CURRENTPRICE_STOCKCODE];
    if (stockCode in stockYearlyTotalInfo) {
      stockYearlyTotalInfo[stockCode]['CompanyId'] = dataCurrentPrice[iRow][COLINDEX_CURRENTPRICE_COMPANYID];
      stockYearlyTotalInfo[stockCode]['SecurityId'] = dataCurrentPrice[iRow][COLINDEX_CURRENTPRICE_SECURITYID];

      stockYearlyTotalInfo[stockCode]['LatestPrice'] = dataCurrentPrice[iRow][COLINDEX_CURRENTPRICE_LATESTPRICE];
      stockYearlyTotalInfo[stockCode]['LatestPriceDate'] = 
        Utilities.formatDate(new Date(dataCurrentPrice[iRow][COLINDEX_CURRENTPRICE_LATESTPRICEDATE]), 'Asia/Manila', 'yyyy-MM-dd');
    }
  }

  // Get the stock price per year where there is value
  for (let keyStockCode in stockYearlyTotalInfo) {
    if (keyStockCode == 'FirstYear') {
      continue;
    }

    for (let keyStockField in stockYearlyTotalInfo[keyStockCode]) {
      let year = parseInt(keyStockField);
      if (!isNaN(year)) {
        let yearNoOfStock = stockYearlyTotalInfo[keyStockCode][year].NoOfStock;
        // If stock was sold out during that year, no point getting the price
        if (yearNoOfStock > 0 && (onlyForThisYear == false || thisYear == year)) {

          // For performance purpose, use cached value for current year
          if (thisYear == year) {
            //console.log(keyStockCode + ' ' + year + ' ' + stockYearlyTotalInfo[keyStockCode].LatestPrice);
            //console.log(stockYearlyTotalInfo[keyStockCode]);
            stockYearlyTotalInfo[keyStockCode][year]['SharePrice'] = stockYearlyTotalInfo[keyStockCode].LatestPrice;
            stockYearlyTotalInfo[keyStockCode][year]['SharePriceDate'] = stockYearlyTotalInfo[keyStockCode].LatestPriceDate;
          } else {
            let latestPrice = getStockEndingPrice_(stockYearlyTotalInfo, keyStockCode, year);
            stockYearlyTotalInfo[keyStockCode][year]['SharePrice'] = latestPrice['CLOSE'];
            let chartDate = new Date(latestPrice['CHART_DATE']);
            stockYearlyTotalInfo[keyStockCode][year]['SharePriceDate'] = 
              Utilities.formatDate(chartDate, 'Asia/Manila', 'yyyy-MM-dd');
          }

        }
      }
    }
  }
}

/**
 * Get Individual Stock Year End (if this year get latest) market price
 * - Prices from Dec 25 to Dec 31 of the year is passed as parameter as stock's last price changes from year to year per stock
 */
function getStockEndingPrice_(stockYearlyTotalInfo, stockCode, year) {
  //console.log(stockCode + ' + ' + year);

  let thisYear = (new Date()).getFullYear();

  // If past year, last price of the year
  let companyId = stockYearlyTotalInfo[stockCode].CompanyId;
  let securityId = stockYearlyTotalInfo[stockCode].SecurityId;
  let fromDateText = '12-25-' + year;
  let toDateText = '12-31-' + year;
  if (year == thisYear) {
    let fromDate = new Date((new Date()).getTime()-3*(24*3600*1000));
    fromDateText = Utilities.formatDate(fromDate, "GMT+9", "MM-dd-yyyy");
    toDateText = Utilities.formatDate(new Date(), "GMT+9", "MM-dd-yyyy");
  }

  let formData = {
    cmpy_id: companyId,
    security_id: securityId,
    startDate: fromDateText,
    endDate: toDateText
  };
  //console.log(JSON.stringify(formData));
  if (securityId == '') {
    return {};
  }

  let options = {
    "method" : "POST",
    "contentType":"application/json",
    // Either stringfying a json or hardcoded string does not work
    "payload" :  JSON.stringify(formData),
    "muteHttpExceptions": true
  };

  let response = UrlFetchApp.fetch(disclousureUrl, options);
  let responseCode = response.getResponseCode();
  if (responseCode == 200) {
    let json_response = JSON.parse(response);
    let chartData = json_response['chartData'];
    if (chartData.length > 0) {
      let lastChartData = chartData[chartData.length - 1];
      return lastChartData;
    } else {
      console.log('Empty Chart Data: ' + stockCode + ' + ' + year);
      console.log(JSON.stringify(formData));
      console.log(response.getContentText());
    }
  } else {
    console.log('Response failed: ' + stockCode + ' + ' + year);
    console.log(response.getResponseCode());
    console.log(response.getContentText());
  }
  return {};
}

const WriteDividendYearlyPerformanceYearColNum = 1;

/**
 * Write down owned stock yearly market price
 * - If no price was retrieved, the cell will not be updated and must be done manually
 * - Each year total is being set, and column formatted 
 */
function writeSharesYearlyPerformance_(stockYearlyTotalInfo) {

  let quantityFormatter = new Intl.NumberFormat('en-US',
    { style:'decimal', minimumFractionDigits: 0 }
  );
  let sharePriceFormatter = new Intl.NumberFormat('en-US',
    { style:'currency',currency:'PHP', minimumFractionDigits: 4 }
  );
  let amountFormatter = new Intl.NumberFormat('en-US',
    { style:'currency',currency:'PHP', minimumFractionDigits: 2 }
  );

  let sheetDivYearlyPerformance = SpreadsheetApp.getActive().getSheetByName(SHEET_YEARLYMARKETVALUE);
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

    sheetDivYearlyPerformance.getRange(columnToLetter(COLINDEX_YEARLYMARKETVALUE_STOCKCODE + 1) + rowIndex).setValue(keyStockCode);
    
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

        let sharePrice = stockYearlyTotalInfo[keyStockCode][year].SharePrice;
        let sharePriceDate = stockYearlyTotalInfo[keyStockCode][year].SharePriceDate;

        let marketValueRange = sheetDivYearlyPerformance.getRange(columnToLetter(colIndex) + rowIndex);
        //console.log(keyStockCode + ' - ' + year + ' - ' + sharePrice);
        if (sharePrice !== undefined) {
          //let marketPrice = Math.round(((yearNoOfStock * sharePrice * COL_MARKETPRICE_NETRATE) * 100) / 100);
          let marketPrice = yearNoOfStock * sharePrice * COL_MARKETPRICE_NETRATE;
          let roundedMarketPrice = Math.round(marketPrice*100)/100;
          marketValueRange.setValue(marketPrice);

          let comment = '\Number of Share: ' + quantityFormatter.format(yearNoOfStock);
          comment += '\nPrice Per Share: ' + sharePriceFormatter.format(sharePrice); 
          comment += '\nPrice Date: ' + sharePriceDate;
          comment += '\nGross Amount: ' + amountFormatter.format(yearNoOfStock * sharePrice);
          
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

function writeYearEndStockQuantity_(stockYearlyTotalInfo) {

  let colSettings = loadColumnSettings(SHEET_STOCKTRANS);

  // loop the transaction table and set the last total number of stock per year
  let sheetStockTrans = SpreadsheetApp.getActive().getSheetByName(SHEET_STOCKTRANS);
  let [rows, columns] = [sheetStockTrans.getLastRow(), sheetStockTrans.getLastColumn()];
  let datasheetStockTrans = sheetStockTrans.getRange(3, 1, rows, columns).getValues();
  //console.log(datasheetStockTrans);

  for (let iRow = 0; iRow < datasheetStockTrans.length; iRow++) {
    let stockCode = datasheetStockTrans[iRow][0];
    let transactDate = new Date(datasheetStockTrans[iRow][1]);
    let formattedTransactDate = Utilities.formatDate(transactDate, 'Asia/Manila', 'yyyy-MM-dd');

    if (stockCode == '') {
      continue;
    }

    let rowIndex = iRow + 3;
    if (stockYearlyTotalInfo[stockCode][transactDate.getFullYear()]['LastTransactDate'] == formattedTransactDate) {
      sheetStockTrans.getRange(colSettings.YearEndQuantity + rowIndex).setValue(
        stockYearlyTotalInfo[stockCode][transactDate.getFullYear()]['NoOfStock']);
    } else {
      sheetStockTrans.getRange(colSettings.YearEndQuantity + rowIndex).setValue('');
    }
  }
}