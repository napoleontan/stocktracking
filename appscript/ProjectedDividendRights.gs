/**
 * Get all future (Unpaid) Dividend and get projected rate, with my owned share get projected gross and net amount
 * - It loops through the PSE Edge dividend pages from page 1 to N
 * - Retrieves the CASH dividend rights for stocks that I own right now (exclude previously owned too)
 * - If the latest page retrieve has no target dividend, then stops the loops, otherwise iterate to next page
 * - Limit upto 10 pages as limit too
 * - Writes all the up coming dividend in the page
 */
function getFutureDividend() {
  let ownedStockInfo = getOwnedStockInfo();
  let dividendPaidInfo = getDividendPaid();

  let colSettings = loadColumnSettings(SHEET_PROJDIV);

  let pageNo = 1;
  let allUnpaidCount = 0;
  let oldUnpaidCount = 0;
  let pageUnpaidCount = 0;
  let allUnpaidDividend = {};
  do {
    let dividendPageInfo = getDividendRightInfoPage_(pageNo, ownedStockInfo, allUnpaidCount);
    oldUnpaidCount = allUnpaidCount;
    allUnpaidCount = filterFutureDividendUnpaidOnly_(dividendPageInfo, allUnpaidDividend, ownedStockInfo, allUnpaidCount);
    pageUnpaidCount = allUnpaidCount - oldUnpaidCount;

    pageNo++;
    //if (pageNo > 2) break;
  } while (/*pageUnpaidCount > 0 &&*/ pageNo <= 10);

  addFutureDividendExpectedToSheet_(colSettings, allUnpaidDividend, ownedStockInfo);
}

function getFutureDividendFromBackup() {
  let ownedStockInfo = getOwnedStockInfo();
  let dividendPaidInfo = getDividendPaid();

  let colSettings = loadColumnSettings(SHEET_PROJDIV);

  let allUnpaidCount = 0;
  let allUnpaidDividend = {};
  let dividendPageInfo = getDividendRightInfoFromBackup_(ownedStockInfo);
  allUnpaidCount = filterFutureDividendUnpaidOnly_(dividendPageInfo, allUnpaidDividend, ownedStockInfo, allUnpaidCount);
  addFutureDividendExpectedToSheet_(colSettings, allUnpaidDividend, ownedStockInfo);
}

/**
 * Using the following sheets, create a list of owned shares info
 * The following data mapping is used.
+--------------+-----------------------+-------------------------------------------------------------------+
| Source Sheet |     Target Value      |                              Remarks                              |
+--------------+-----------------------+-------------------------------------------------------------------+
| Trans(Total) | StockCode             | Used as key                                                       |
| Trans(Total) | ShareOwned            |                                                                   |
| Trans(Total) | AvgPricePerShare      |                                                                   |
| CurrentPrice | CompanyId             | Used to distinguish which dividend rights info is for which stock |
| CashDividend | LastPaidExDividedDate | Used as signal to know which dividend rights info to ignore       |
| N/A          | ToBePaidExDividedDate | dummy field defaulted to null                                     |
+--------------+-----------------------+-------------------------------------------------------------------+
 */
function getOwnedStockInfo() {
  let sheetTransTotal = SpreadsheetApp.getActive().getSheetByName(SHEET_TRANSTOTAL);
  let [rows, columns] = [sheetTransTotal.getLastRow(), sheetTransTotal.getLastColumn()];
  let dataTransTotal = sheetTransTotal.getRange(2, 1, rows, columns).getValues();
  //console.log(dataTransTotal);

  let sheetCurrentPrice = SpreadsheetApp.getActive().getSheetByName(SHEET_CURRENTPRICE);
  [rows, columns] = [sheetCurrentPrice.getLastRow(), sheetCurrentPrice.getLastColumn()];
  let dataCurrentPrice = sheetCurrentPrice.getRange(ROWINDEX_CURRENTPRICE_FIRSTROW, 1, rows, columns).getValues();
  //console.log(dataCurrentPrice);

  let sheetCashDividend = SpreadsheetApp.getActive().getSheetByName(SHEET_CASHDIV);
  [rows, columns] = [sheetCashDividend.getLastRow(), sheetCashDividend.getLastColumn()];
  let dataCashDividend = sheetCashDividend.getRange(3, 1, rows, columns).getValues();
  //console.log(dataCashDividend);

  let dict_data = {};
  for (let iRow = 0; iRow < dataTransTotal.length; iRow++) {
    let stockCode1 = dataTransTotal[iRow][COLINDEX_TRANSTOTAL_STOCKCODE];
    if (stockCode1 == '') {
      continue;
    }
    let sharedOwned = parseInt(dataTransTotal[iRow][COLINDEX_TRANSTOTAL_QUANTITY], 10);
    if (sharedOwned <= 0) {
      continue;
    }

    let rowDictData = {};
    rowDictData['StockCode'] = stockCode1;
    rowDictData['ShareOwned'] = dataTransTotal[iRow][COLINDEX_TRANSTOTAL_QUANTITY];
    rowDictData['LastPaidExDividedDate'] = null;
    rowDictData['ToBePaidExDividedDate'] = null;

    // Copy the CompanyId info
    for (let iRow = 0; iRow < dataCurrentPrice.length; iRow++) {
      let stockCode2 = dataCurrentPrice[iRow][COLINDEX_CURRENTPRICE_STOCKCODE];
      if (stockCode1 == stockCode2) {
        rowDictData['CompanyId'] = dataCurrentPrice[iRow][COLINDEX_CURRENTPRICE_COMPANYID];
        rowDictData['AvgPricePerShare'] = dataCurrentPrice[iRow][COLINDEX_CURRENTPRICE_AVGPRICEPERSHARE];
        break;
      }
    }

    // Copy the last paid dividend date
    for (let iRow = 0; iRow < dataCashDividend.length; iRow++) {
      let stockCode3 = dataCashDividend[iRow][COLINDEX_CASHDIV_STOCKCODE];
      if (stockCode1 == stockCode3) {
        if (rowDictData['LastPaidExDividedDate'] == null) {
          rowDictData['LastPaidExDividedDate'] = Utilities.formatDate(new Date(dataCashDividend[iRow][COLINDEX_CASHDIV_EXDIVDATE]), 'Asia/Tokyo', 'yyyy-MM-dd');
          break;
        }
      }
    }

    dict_data[stockCode1] = rowDictData;
  }
  //console.log(dict_data);

  return dict_data;
}

/**
 * Get all paid dividend info
 */
function getDividendPaid() {
  let sheetCashDividend = SpreadsheetApp.getActive().getSheetByName(SHEET_CASHDIV);
  let [rows, columns] = [sheetCashDividend.getLastRow(), sheetCashDividend.getLastColumn()];
  let dataCashDividend = sheetCashDividend.getRange(3, 1, rows, columns).getValues();
  //console.log(dataCashDividend);

  let dividendIndex = 0;
  let dict_data = {};
  for (let iRow = 0; iRow < dataCashDividend.length; iRow++) {
    let stockCode = dataCashDividend[iRow][COLINDEX_CASHDIV_STOCKCODE];
    if (stockCode == '') {
      continue;
    }

    let rowDictData = {};
    rowDictData['StockCode'] = stockCode;
    rowDictData['ExDividendDate'] = Utilities.formatDate(new Date(dataCashDividend[iRow][COLINDEX_CASHDIV_EXDIVDATE]), 'Asia/Tokyo', 'yyyy-MM-dd');
    rowDictData['DividendPerShare'] = dataCashDividend[iRow][COLINDEX_CASHDIV_DIVPERSHARE];

    dict_data[dividendIndex + 1] = rowDictData;
    dividendIndex++;
  }
  //console.log(dict_data);

  return dict_data;
}

const dividendRightsInfoUrl = 'https://edge.pse.com.ph/disclosureData/dividends_and_rights_info_list.ax?DividendsOrRights=Dividends';

const dividendRightsInfoRowStartText = '<tr>';
// Company Name and Link 
const dividendRightsInfoCol1StartText = '<td><a href="/companyInformation/form.do?cmpy_id=';
const dividendRightsInfoCol1EndText = '</a></td>';
// Company Code
const dividendRightsInfoCol2StartText = '<td class="alignC">';
const dividendRightsInfoCol2EndText = '</td>';
// Dividend Type Cash or Stock
const dividendRightsInfoCol3StartText = '<td class="alignC">';
const dividendRightsInfoCol3EndText = '</td>';
// Dividend Rate
const dividendRightsInfoCol4StartText = '<td class="alignR">';
const dividendRightsInfoCol4EndText = '</td>';
// Ex Dividend Date
const dividendRightsInfoCol5StartText = '<td class="alignC">';
const dividendRightsInfoCol5EndText = '</td>';
// Record Date
const dividendRightsInfoCol6StartText = '<td class="alignC">';
const dividendRightsInfoCol6EndText = '</td>';
// Payment Date
const dividendRightsInfoCol7StartText = '<td class="alignC">';
const dividendRightsInfoCol7EndText = '</td>';

/**
 * Get dividend info for stock owned in a specific page
 * - Uses startIndex to know what key to place in returned info
 * - This methods uses table marker to parse the values.
 * - A dividend rights info has the following info:
 *   - PageNo, StockCode, DividendPerShare, ExDividendDate
 * - The DividendPerShare has lots of textual info and lots of parsing done to remove those
 *   - A limitation is for PIZZA stocks which writes values as "Three Centavos", did not handle this edge case
 */
function getDividendRightInfoPage_(pageNo, ownedStockInfo, startIndex) {

  let formData = {
    'pageNum': '' + pageNo,
    'date': 'date',
    'dateSortType': 'DESC',
    'cmpySortType': 'ASC'
  };
  let options = {
    'method' : 'post',
    'payload' : formData,
    contentType: "application/x-www-form-urlencoded; charset=UTF-8",
    muteHttpExceptions: false,
  };

  console.log(dividendRightsInfoUrl);
  let response = UrlFetchApp.fetch(dividendRightsInfoUrl, options);
  let responseCode = response.getResponseCode();
  console.log('getDividendRightInfoPage(pageNo:= ' + pageNo + ', startIndex:= ' + startIndex + ') ResponseCode' + response.getResponseCode());
  let html = response.getContentText();
  //console.log(html);

  if (responseCode != 200) {
    return {};
  }

  let dict_data = {};
  let foundDataIndex = startIndex;
  let trStartTagIndex = html.indexOf(dividendRightsInfoRowStartText);
  do {
    let td1Start = html.indexOf(dividendRightsInfoCol1StartText, trStartTagIndex + dividendRightsInfoRowStartText.length);
    let td1MidEnd = html.indexOf("\">", td1Start + dividendRightsInfoCol1StartText.length);
    let td1End = html.indexOf(dividendRightsInfoCol1EndText, td1MidEnd + 2);
    let tdCompanyCode = html.substring(td1Start + dividendRightsInfoCol1StartText.length, td1MidEnd);
    
    let td2Start = html.indexOf(dividendRightsInfoCol2StartText, td1End + dividendRightsInfoCol1EndText.length);
    let td2End = html.indexOf(dividendRightsInfoCol2EndText, td2Start + dividendRightsInfoCol2StartText.length);
    let tdSecurityTypeCode = html.substring(td2Start + dividendRightsInfoCol2StartText.length, td2End);

    let td3Start = html.indexOf(dividendRightsInfoCol3StartText, td2End + dividendRightsInfoCol2EndText.length);
    let td3End = html.indexOf(dividendRightsInfoCol3EndText, td3Start + dividendRightsInfoCol3StartText.length);
    let tdTypeDividend = html.substring(td3Start + dividendRightsInfoCol3StartText.length, td3End);

    let td4Start = html.indexOf(dividendRightsInfoCol4StartText, td3End + dividendRightsInfoCol3EndText.length);
    let td4End = html.indexOf(dividendRightsInfoCol4EndText, td4Start + dividendRightsInfoCol4StartText.length);
    let tdDividendRate = html.substring(td4Start + dividendRightsInfoCol4StartText.length, td4End);

    let tdCleanDividendRate = tdDividendRate;
    // special case
    if (tdCleanDividendRate.indexOf('Three') >= 0 && tdCleanDividendRate.indexOf('Centavos') >= 0) {
      tdCleanDividendRate = '0.03';
    }
    let tdDividendRatePesoIndex = tdCleanDividendRate.indexOf('P');
    if (tdDividendRatePesoIndex > 0) {
      tdCleanDividendRate = tdCleanDividendRate.substring(tdDividendRatePesoIndex);
    }
    //let preClean = tdCleanDividendRate;
    tdCleanDividendRate = tdCleanDividendRate.replace('Php.','');
    tdCleanDividendRate = tdCleanDividendRate.replace('PhP.','');
    tdCleanDividendRate = tdCleanDividendRate.replace('PHP','');
    tdCleanDividendRate = tdCleanDividendRate.replace('Php','');
    tdCleanDividendRate = tdCleanDividendRate.replace('PhP','');
    tdCleanDividendRate = tdCleanDividendRate.replace('P','');
    //console.log('PreClean DivRate: [' + preClean + '] post clean [' + tdCleanDividendRate + ']');

    let tdDividendRateSlashIndex = tdCleanDividendRate.indexOf('/');
    if (tdDividendRateSlashIndex > 0) {
      tdCleanDividendRate = tdCleanDividendRate.substring(0, tdDividendRateSlashIndex);
    }
    let tdDividendRateSpaceIndex = tdCleanDividendRate.indexOf(' ');
    if (tdDividendRateSpaceIndex > 0) {
      tdCleanDividendRate = tdCleanDividendRate.substring(0, tdDividendRateSpaceIndex);
    }
    let tdDividendRateParenthesisIndex = tdCleanDividendRate.indexOf(')');
    if (tdDividendRateParenthesisIndex > 0) {
      tdCleanDividendRate = tdCleanDividendRate.substring(0, tdDividendRateParenthesisIndex);
    }
    tdCleanDividendRate = tdCleanDividendRate.replace(/[^\d.-]/g, '');

    let td5Start = html.indexOf(dividendRightsInfoCol5StartText, td4End + dividendRightsInfoCol4EndText.length);
    let td5End = html.indexOf(dividendRightsInfoCol5EndText, td5Start + dividendRightsInfoCol5StartText.length);
    let tdDividendExDate = html.substring(td5Start + dividendRightsInfoCol5StartText.length, td5End);

    let td6Start = html.indexOf(dividendRightsInfoCol6StartText, td5End + dividendRightsInfoCol5EndText.length);
    let td6End = html.indexOf(dividendRightsInfoCol6EndText, td6Start + dividendRightsInfoCol6StartText.length);
    let tdRecordDate = html.substring(td6Start + dividendRightsInfoCol6StartText.length, td6End);

    let td7Start = html.indexOf(dividendRightsInfoCol7StartText, td6End + dividendRightsInfoCol6EndText.length);
    let td7End = html.indexOf(dividendRightsInfoCol7EndText, td7Start + dividendRightsInfoCol7StartText.length);
    let tdPaymentDate = html.substring(td7Start + dividendRightsInfoCol7StartText.length, td7End);

    let foundStockCode = '';
    // PIZZA has dividend worded in english like "Three Centavo", for now we ignore it?
    //if (tdTypeDividend == 'Cash' && !tdDividendRate.includes('Centavo')) {
    if (tdTypeDividend == 'Cash') {
      // First check if the dividend rights info has equal company code and stock code
      for (let keyStockCode in ownedStockInfo) {
        if (ownedStockInfo[keyStockCode].CompanyId == tdCompanyCode && ownedStockInfo[keyStockCode].StockCode == tdSecurityTypeCode) {
          //console.log('TD1.1: ' + tdCompanyCode);
          //console.log('TD1.2: ' + tdSecurityTypeCode);
          //console.log('StockCode: ' + keyStockCode);
          foundStockCode = keyStockCode;
          break;
        }
      }
      // Then just try to match dividend rights info to the company code only
      if (foundStockCode == '') {
        for (let keyStockCode in ownedStockInfo) {
          if (ownedStockInfo[keyStockCode].CompanyId == tdCompanyCode && tdCompanyCode != '86' && tdSecurityTypeCode == 'COMMON') {
            //console.log('TD2.1: ' + tdCompanyCode);
            //console.log('StockCode: ' + keyStockCode);
            foundStockCode = keyStockCode;
            break;
          }
        }
      }
      // Possible 3rd check is for name but will do that last resort later

      if (foundStockCode != '') {
        let rowDictData = {};
        rowDictData['PageNo'] = pageNo;
        rowDictData['StockCode'] = foundStockCode;
        rowDictData['DividendPerShare'] = tdCleanDividendRate;
        rowDictData['ExDividendDate'] = Utilities.formatDate(new Date(tdDividendExDate), 'Asia/Tokyo', 'yyyy-MM-dd');
        if (tdDividendExDate == '' && tdPaymentDate != '') {
          rowDictData['ExDividendDate'] = Utilities.formatDate(new Date(tdPaymentDate), 'Asia/Tokyo', 'yyyy-MM-dd');
        }
        rowDictData['RecordDate'] = Utilities.formatDate(new Date(tdRecordDate), 'Asia/Tokyo', 'yyyy-MM-dd');
        rowDictData['PaymentDate'] = Utilities.formatDate(new Date(tdPaymentDate), 'Asia/Tokyo', 'yyyy-MM-dd');
        dict_data[foundDataIndex + 1] = rowDictData;
        foundDataIndex++;
      }
    }

    trStartTagIndex = html.indexOf(dividendRightsInfoRowStartText, td2End + dividendRightsInfoCol1EndText.length);
  } while (trStartTagIndex > -1);

  console.log(dict_data);
  return dict_data;
}

function getDividendRightInfoFromBackup_(ownedStockInfo) {
  let sheetBackupDividend = SpreadsheetApp.getActive().getSheetByName(SHEET_BACKUPDIVIDEND);
  let [rows, columns] = [sheetBackupDividend.getLastRow(), sheetBackupDividend.getLastColumn()];
  let dataBackupDividend = sheetBackupDividend.getRange(3, 1, rows, columns).getValues();

  let dict_data = {};
  let foundDataIndex = 0;
  for (let iRow = dataBackupDividend.length - 1; iRow > 0; iRow--) {
    let keyStockCode = dataBackupDividend[iRow][COLINDEX_BACKUPDIVIDEND_STOCKCODE];
    let dividendType = dataBackupDividend[iRow][COLINDEX_BACKUPDIVIDEND_DIVIDENDTYPE];
    if (keyStockCode == '') {
      continue;
    }
    if (dividendType != 'Cash') {
      continue;
    }
    // First check if the dividend rights info has equal stock code
    if (ownedStockInfo[keyStockCode] === undefined) {
      continue;
    }

    let cleanDividendPerShare = dataBackupDividend[iRow][COLINDEX_BACKUPDIVIDEND_DIVPERSHARE].toString();
    let exDividendDate = dataBackupDividend[iRow][COLINDEX_BACKUPDIVIDEND_EXDIVDATE].toString();
    let recordDate = dataBackupDividend[iRow][COLINDEX_BACKUPDIVIDEND_RECORDDATE].toString();
    let paymentDate = dataBackupDividend[iRow][COLINDEX_BACKUPDIVIDEND_PAYMENTDATE].toString();
    // special case
    if (cleanDividendPerShare.indexOf('Three') >= 0 && cleanDividendPerShare.indexOf('Centavos') >= 0) {
      cleanDividendPerShare = '0.03';
    }
    let dividendRatePesoIndex = cleanDividendPerShare.indexOf('P');
    if (dividendRatePesoIndex > 0) {
      cleanDividendPerShare = cleanDividendPerShare.substring(dividendRatePesoIndex);
    }
    cleanDividendPerShare = cleanDividendPerShare.replace('Php.','');
    cleanDividendPerShare = cleanDividendPerShare.replace('PHP','');
    cleanDividendPerShare = cleanDividendPerShare.replace('Php','');
    cleanDividendPerShare = cleanDividendPerShare.replace('PhP','');
    cleanDividendPerShare = cleanDividendPerShare.replace('P','');

    let tdDividendRateSlashIndex = cleanDividendPerShare.indexOf('/');
    if (tdDividendRateSlashIndex > 0) {
      cleanDividendPerShare = cleanDividendPerShare.substring(0, tdDividendRateSlashIndex);
    }
    let tdDividendRateSpaceIndex = cleanDividendPerShare.indexOf(' ');
    if (tdDividendRateSpaceIndex > 0) {
      cleanDividendPerShare = cleanDividendPerShare.substring(0, tdDividendRateSpaceIndex);
    }
    let tdDividendRateParenthesisIndex = cleanDividendPerShare.indexOf(')');
    if (tdDividendRateParenthesisIndex > 0) {
      cleanDividendPerShare = cleanDividendPerShare.substring(0, tdDividendRateParenthesisIndex);
    }

    cleanDividendPerShare = cleanDividendPerShare.replace(/[^\d.-]/g, '');

    let rowDictData = {};
    rowDictData['PageNo'] = '0';
    rowDictData['StockCode'] = keyStockCode;
    rowDictData['DividendPerShare'] = cleanDividendPerShare;
    rowDictData['ExDividendDate'] = Utilities.formatDate(new Date(exDividendDate), 'Asia/Tokyo', 'yyyy-MM-dd');
    rowDictData['RecordDate'] = Utilities.formatDate(new Date(recordDate), 'Asia/Tokyo', 'yyyy-MM-dd');
    rowDictData['PaymentDate'] = Utilities.formatDate(new Date(paymentDate), 'Asia/Tokyo', 'yyyy-MM-dd');
    dict_data[foundDataIndex + 1] = rowDictData;
    foundDataIndex++;
  }

  console.log(dict_data);

  return dict_data;
}

/**
 * Filters the dividend rights info for a specific page and only returns upcoming (unpaid) dividend rights info
 * - If the stock is new and is never paid before, get the first one and not the next
 *   - this has limitation but will go away through time once a share has first paid dividend
 * - If the stock is previously paid before, check the ex dividend date and compare if it is after the last paid ex dividend date
 */
function filterFutureDividendUnpaidOnly_(dividendPageInfo, allUnpaidDividend, ownedStockInfo, allUnpaidDividendSize) {
  let rowDictData = {};
  let unpaidDividendIndex = allUnpaidDividendSize;
  let nonTargetDividendCounter = 0;

  for (let keyDividendPageIndex in dividendPageInfo) {
    let aDividendPageInfo = dividendPageInfo[keyDividendPageIndex];
    let pageInfoExDividendDate = new Date(aDividendPageInfo.ExDividendDate);
    //console.log('Last Paid: ' + lastPaidExDividendDate);

    // An old dividend info, there was a paid value already
    if (ownedStockInfo[aDividendPageInfo.StockCode].LastPaidExDividedDate != null) {
      let lastPaidExDividendDate = new Date(ownedStockInfo[aDividendPageInfo.StockCode].LastPaidExDividedDate);
      if (pageInfoExDividendDate.getTime() < lastPaidExDividendDate.getTime()) {
        continue;
      }
    }
    if (ownedStockInfo[aDividendPageInfo.StockCode].ToBePaidExDividedDate != null) {
      let toBePaidExDividendDate = new Date(ownedStockInfo[aDividendPageInfo.StockCode].ToBePaidExDividedDate);
      if (pageInfoExDividendDate.getTime() < toBePaidExDividendDate.getTime()) {
        continue;
      }
    }

    //console.log('page info ' + JSON.stringify(aDividendPageInfo));
    let alreadyFoundUnpaidDividendInfo = false;
    for (let keyAllUnpaidIndex in allUnpaidDividend) {
      let theAllUnpaidDividendInfo = allUnpaidDividend[keyAllUnpaidIndex];
      let theAllUnpaidExDividendDate = new Date(theAllUnpaidDividendInfo.ExDividendDate);

      if (aDividendPageInfo.StockCode == theAllUnpaidDividendInfo.StockCode) {
      //console.log('>>> unpaid loop index=' + keyAllUnpaidIndex + ' ' + JSON.stringify(theAllUnpaidDividendInfo));
      //console.log('>>> exDate ' + pageInfoExDividendDate.valueOf() + ' vs ' + theAllUnpaidExDividendDate.valueOf());
        alreadyFoundUnpaidDividendInfo = true;
        if (pageInfoExDividendDate.valueOf() < theAllUnpaidExDividendDate.valueOf()) {
          // Several advance dividend were announced only use the nearest one
          // override to use new didivend info
          allUnpaidDividend[keyAllUnpaidIndex] = aDividendPageInfo;
          //console.log('))) Replace Unpaid ' + aDividendPageInfo.StockCode);
        } else if (pageInfoExDividendDate.valueOf() == theAllUnpaidExDividendDate.valueOf()) {
          // Same date but multiple line were declared (regular + special dividend declaration were done, merge)
          // Increment the dividend per share
          allUnpaidDividend[keyAllUnpaidIndex].DividendPerShare = parseFloat(allUnpaidDividend[keyAllUnpaidIndex].DividendPerShare) +  parseFloat(aDividendPageInfo.DividendPerShare);
          //console.log('))) Increment Unpaid ' + aDividendPageInfo.StockCode + ' to ' + allUnpaidDividend[keyAllUnpaidIndex].DividendPerShare);
        }
      }
    } // end for

    if (alreadyFoundUnpaidDividendInfo == false) {
      // totally new stock code, add to unpaid
      unpaidDividendIndex++;
      allUnpaidDividend[unpaidDividendIndex] = aDividendPageInfo;
      //console.log('))) New Unpaid ' + aDividendPageInfo.StockCode);

      if (ownedStockInfo[aDividendPageInfo.StockCode].ToBePaidExDividedDate == null) {
        ownedStockInfo[aDividendPageInfo.StockCode].ToBePaidExDividedDate = allUnpaidDividend[unpaidDividendIndex].ExDividendDate;
      }
    }

  } // end for
  return unpaidDividendIndex;
}

/**
 * Writes down all future dividend info to the "ProjectedDividend" excel sheet.
 * - All previous sheet values are cleared
 * - Also add summation formula to total certain columns
 */
function addFutureDividendExpectedToSheet_(colSettings, futureDividendInfo, ownedStockInfo) {

  getCashDividendEntitledQuantity_(futureDividendInfo, ownedStockInfo);

  // Clear the previous generated projeccted dividend info
  let sheetProjectedDividend = SpreadsheetApp.getActive().getSheetByName(SHEET_PROJDIV);
  let [rows, columns] = [sheetProjectedDividend.getLastRow(), sheetProjectedDividend.getLastColumn()];
  if (rows.length > 0) {
    sheetProjectedDividend.getRange(4, 1, rows, columns).clearContent();
    SpreadsheetApp.flush();
  }

  let rowIndex = 3;
  let cellRange = '';
  let formulaVal = '';
  for (let keyDivIndex in futureDividendInfo) {
    let aFutureDividendInfo = futureDividendInfo[keyDivIndex];

    cellRange = colSettings.PageNo + rowIndex;
    sheetProjectedDividend.getRange(cellRange).setValue(aFutureDividendInfo.PageNo);
    cellRange = colSettings.StockCode + rowIndex;
    sheetProjectedDividend.getRange(cellRange).setValue(aFutureDividendInfo.StockCode);
    cellRange = colSettings.ExDividendDate + rowIndex;
    sheetProjectedDividend.getRange(cellRange).setValue(aFutureDividendInfo.ExDividendDate);
    cellRange = colSettings.PaymentDate + rowIndex;
    sheetProjectedDividend.getRange(cellRange).setValue(aFutureDividendInfo.PaymentDate);
    cellRange = colSettings.DividendPerShare + rowIndex;
    sheetProjectedDividend.getRange(cellRange).setValue(aFutureDividendInfo.DividendPerShare);
    cellRange = colSettings.Quantity + rowIndex;
    //sheetProjectedDividend.getRange(cellRange).setValue(ownedStockInfo[aFutureDividendInfo.StockCode].ShareOwned);
    sheetProjectedDividend.getRange(cellRange).setValue(aFutureDividendInfo.EntitledQty);

    cellRange = colSettings.GrossAmount + rowIndex;
    formulaVal = "=" + colSettings.DividendPerShare + rowIndex + "*" + colSettings.Quantity + rowIndex;
    sheetProjectedDividend.getRange(cellRange).setFormula(formulaVal);

    cellRange = colSettings.NetAmount + rowIndex;
    formulaVal = "=" + colSettings.DividendPerShare + rowIndex + "*" + colSettings.Quantity + rowIndex + "*" + CASHDIV_WITHTAX_NETRATE;
    sheetProjectedDividend.getRange(cellRange).setFormula(formulaVal);

    cellRange = colSettings.DividendRate + rowIndex;
    formulaVal = "=" + colSettings.DividendPerShare + rowIndex + "/" + colSettings.PricePerShare + rowIndex;
    sheetProjectedDividend.getRange(cellRange).setFormula(formulaVal);

    cellRange = colSettings.PricePerShare + rowIndex;
    sheetProjectedDividend.getRange(cellRange).setValue(ownedStockInfo[aFutureDividendInfo.StockCode].AvgPricePerShare);

    formulaVal = "=" + colSettings.Quantity + rowIndex + "*" + colSettings.PricePerShare + rowIndex;
    cellRange = colSettings.MarketPrice + rowIndex;
    sheetProjectedDividend.getRange(cellRange).setFormula(formulaVal);
    
    rowIndex++;
  }

  if (rowIndex > 3) {
    cellRange = colSettings.GrossAmount + "1";
    formulaVal = "=SUM(" + colSettings.GrossAmount + "3:" + colSettings.GrossAmount + rowIndex + ")";
    sheetProjectedDividend.getRange(cellRange).setFormula(formulaVal);

    cellRange = colSettings.NetAmount + "1";
    formulaVal = "=SUM(" + colSettings.NetAmount + "3:" + colSettings.NetAmount + rowIndex + ")";
    sheetProjectedDividend.getRange(cellRange).setFormula(formulaVal);

    cellRange = colSettings.MarketPrice + "1";
    formulaVal = "=SUM(" + colSettings.MarketPrice + "3:" + colSettings.MarketPrice + rowIndex + ")";
    sheetProjectedDividend.getRange(cellRange).setFormula(formulaVal);
  }
}

function getCashDividendEntitledQuantity_(futureDividendInfo, ownedStockInfo) {

  let sheetStockTrans = SpreadsheetApp.getActive().getSheetByName(SHEET_STOCKTRANS);
  let [rows, columns] = [sheetStockTrans.getLastRow(), sheetStockTrans.getLastColumn()];
  let dataStockTrans = sheetStockTrans.getRange(3, 1, rows, columns).getValues();

  for (let keyDivIndex in futureDividendInfo) {
    let aFutureDividendInfo = futureDividendInfo[keyDivIndex];
    let entitledQty = ownedStockInfo[aFutureDividendInfo.StockCode].ShareOwned;
    let futureExDividendDate = new Date(aFutureDividendInfo.ExDividendDate);
    
    console.log('Stock ' + aFutureDividendInfo.StockCode + ' owned: ' + entitledQty + ' exDate ' + aFutureDividendInfo.ExDividendDate);
    console.log(JSON.stringify(aFutureDividendInfo));

    for (let iRow = 0; iRow < dataStockTrans.length; iRow++) {
      let stockCode = dataStockTrans[iRow][COLINDEX_STOCKTRANS_STOCKCODE];
      let transDateText = dataStockTrans[iRow][COLINDEX_STOCKTRANS_TRANSDATE];
      let transDate = new Date(transDateText);
      let quantity = dataStockTrans[iRow][COLINDEX_STOCKTRANS_QUANTITY];
      let stockTransType = dataStockTrans[iRow][COLINDEX_STOCKTRANS_TRANSTYPE];

      // Stop if transact type missing
      if (aFutureDividendInfo.StockCode != stockCode) {
        continue;
      }
      console.log('match ' + transDateText)
      if (transDate.valueOf() < futureExDividendDate.valueOf()) {
        break;
      }
      if (isBuyTransactionType(stockTransType) || stockTransType == TRANSTYPE_STOCKDIVIDEND) {
        entitledQty = entitledQty - parseInt(quantity);
      } else if (stockTransType == TRANSTYPE_SOLDSHARES) {
        entitledQty = entitledQty + (-1 * parseInt(quantity));
      }
    } // end for

    console.log('Stock ' + aFutureDividendInfo.StockCode + ' entitied: ' + entitledQty);
    aFutureDividendInfo.EntitledQty = entitledQty;

  } // end for
}
