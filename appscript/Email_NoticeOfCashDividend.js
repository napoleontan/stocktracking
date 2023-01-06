/**
 * Search Gmail to see the last 10 cash dividend notification email
 * For each message found, 
 * - parse the email for the cash dividend details
 * - check that the dividend info is already logged in the CashDividend sheet.
 * - if it is not yet logged, add a new row from top for each dividend
 */
function searchNewEmailColNoticeOfCashDividend() {
  
  let colSettings = loadColumnSettings(SHEET_CASHDIV);

  let sheetCashDividend = SpreadsheetApp.getActive().getSheetByName(SHEET_CASHDIV);
  let [rows, columns] = [sheetCashDividend.getLastRow(), sheetCashDividend.getLastColumn()];
  let dataCashDividend = sheetCashDividend.getRange(3, 1, rows, columns).getValues();
  //console.log(dataCashDividend);

  let sheetCurrentPrice = SpreadsheetApp.getActive().getSheetByName(SHEET_CURRENTPRICE);
  [rows, columns] = [sheetCurrentPrice.getLastRow(), sheetCurrentPrice.getLastColumn()];
  let dataCurrentPrice = sheetCurrentPrice.getRange(ROWINDEX_CURRENTPRICE_FIRSTROW, 1, rows, columns).getValues();
  let latestPriceInfo = {};

  let addToRow = 3;
  let emailInfo = readColNoticeOfCashDividendGmail_();
  for (let messageIndex in emailInfo) {
    let aNoticeInfo = emailInfo[messageIndex];
    let notificationLogged = checkColNoticeOfCashDividendExists_(dataCashDividend, aNoticeInfo);
    if (!notificationLogged) {
      console.log('Un logged Notification found');
      // Some mails have no transaction date, just use email date
      let paymentDate = (aNoticeInfo.TransactionInfo.PaymentDate == null) ? aNoticeInfo.MailDate : aNoticeInfo.TransactionInfo.PaymentDate;
      aNoticeInfo.TransactionInfo.PaymentDate = paymentDate;

      console.log(paymentDate + ' (' + aNoticeInfo.TransactionInfo.ExDividendDate + ') ' + aNoticeInfo.TransactionInfo.StockCode);
      getCashDividendNoticeSharePrice_(latestPriceInfo, dataCurrentPrice, aNoticeInfo);

      let computedAveragePrice = computeAveragePriceByStockCode(aNoticeInfo.TransactionInfo.StockCode, aNoticeInfo.TransactionInfo.ExDividendDate);
      aNoticeInfo.AverageBuyPricePerShare = computedAveragePrice.AverageBuyPricePerShare;

      logCashDividendNoticeToSheets_(colSettings, addToRow, aNoticeInfo, sheetCurrentPrice.getLastRow());
      console.log(aNoticeInfo);
      addToRow++;
    }
  } // end for 
}

/**
 * Search the latest 10 col Notice of Cash Dividend email thread
 */
function readColNoticeOfCashDividendGmail_() {

  let mailInfo = {};
  let messageIndex = 0;

  let searchKey = 'colfinancial subject:"Notice of Cash Dividend"';
  let threads = GmailApp.search(searchKey, 0, 10); 
  for (var i = 0; i < threads.length; i++) {
    let messages = threads[i].getMessages();
    for (var j = 0;  j < messages.length;  j++) {
      let message = messages[j];

      let messageInfo = {};
      messageInfo['MsgIndex'] = messageIndex;
      messageInfo['Subject'] = message.getSubject();
      messageInfo['MailDate'] = Utilities.formatDate(message.getDate(), 'Asia/Tokyo', 'yyyy-MM-dd');
      
      let mailBody = message.getBody().replace(/\r\n/g, ' ');

      messageInfo['TransactionInfo'] = readColNoticeOfCashDividendMessage_(messageIndex, mailBody);
      mailInfo[messageIndex] = messageInfo;

      // For DEBUG
      if (messageIndex == -1) {
        console.log(message.getBody().substring(9000));
        console.log(mailBody.substring(9000));
      } 
      //console.log(JSON.stringify(messageInfo));
      
      messageIndex++;
    }
  }

  console.log(JSON.stringify(mailInfo));
  return mailInfo;
}


/**
 * Read a COL Cash Dividend Notification email message
 */
function readColNoticeOfCashDividendMessage_(messageIndex, mailBody) {
  let marker = getColNoticeOfCashDividendMessageMarker_(messageIndex, mailBody);

  let messageInfo = {};

  messageInfo['MarkerType'] = marker['MarkerType'];

  // Get Record start for Transaction Date
  let field1Start = mailBody.indexOf(marker.Field1Start);
  let field1End = mailBody.indexOf(marker.CommonEnd, field1Start + marker.Field1Start.length);
  let paymentDateText = mailBody.substring(field1Start + marker.Field1Start.length, field1End);
  let paymentDate = Utilities.formatDate(new Date(paymentDateText), 'Asia/Tokyo', 'yyyy-MM-dd');
  if (paymentDate == '1970-01-01') {
    paymentDate = null;
  }
  messageInfo['PaymentDate'] = paymentDate;

  // Stock Code
  let stockCodeFieldParseInfo = readColNoticeOfCashDividendMessageField_(mailBody, 
    field1End + marker.CommonEnd.length, 
    marker.Field2Start, marker.CommonMid, marker.CommonEnd);
  messageInfo['StockCode'] = stockCodeFieldParseInfo.FieldValue;

  // Dividend Per Share
  let divPerShareFieldParseInfo = readColNoticeOfCashDividendMessageField_(mailBody, 
    stockCodeFieldParseInfo.FieldEnd + marker.CommonEnd.length, 
    marker.Field3Start, marker.CommonMid, marker.Field3End);

  if (divPerShareFieldParseInfo.FieldValue.replace(/\s/g, '') == '') {
    let divPerPriceRawValue = mailBody.substring(divPerShareFieldParseInfo.FieldMid, divPerShareFieldParseInfo.FieldEnd);
    //console.log('SharePriceRaw:[' + divPerPriceRawValue + ']');
    let spanIndexOf = divPerPriceRawValue.indexOf('</span>');
    if (spanIndexOf > 0) {
      divPerShareFieldParseInfo.FieldValue = divPerPriceRawValue.substring(0, spanIndexOf - 1);
      divPerShareFieldParseInfo.FieldValue = divPerShareFieldParseInfo.FieldValue.replace('\"', '');
      divPerShareFieldParseInfo.FieldValue = divPerShareFieldParseInfo.FieldValue.replace('>', '');
    }
  }
  messageInfo['DividendPerShare'] = divPerShareFieldParseInfo.FieldValue;

  // Ex Dividend Date
  let exDivDateFieldParseInfo = readColNoticeOfCashDividendMessageField_(mailBody, 
    divPerShareFieldParseInfo.FieldEnd + marker.CommonEnd.length, 
    marker.Field4Start, marker.CommonMid, marker.CommonEnd);
  let exDividendDateObj = new Date(exDivDateFieldParseInfo.FieldValue);
  // Sometimes the date parse do not contain the year and it defaults to 2001, so if year diff from current year > 20, set to current year
  var today = new Date();
  let exDividendDate = Utilities.formatDate(exDividendDateObj, 'Asia/Tokyo', 'yyyy-MM-dd');
  if (Math.abs(today.getFullYear() - exDividendDateObj.getFullYear()) > 20) {
    exDividendDateObj.setFullYear(today.getFullYear());
    exDividendDate = Utilities.formatDate(exDividendDateObj, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  messageInfo['ExDividendDate'] = exDividendDate;

  // No of Share
  let noOfShareFieldParseInfo = readColNoticeOfCashDividendMessageField_(mailBody, 
    exDivDateFieldParseInfo.FieldEnd + marker.CommonEnd.length, 
    marker.Field5Start, marker.CommonMid, marker.CommonEnd);
  messageInfo['NoOfShare'] = noOfShareFieldParseInfo.FieldValue.replace(',', '');

  // Gross Amount
  let grossAmtFieldParseInfo = readColNoticeOfCashDividendMessageField_(mailBody, 
    noOfShareFieldParseInfo.FieldEnd + marker.CommonEnd.length, 
    marker.Field6Start, marker.CommonMid, marker.CommonEnd);
  messageInfo['GrossAmount'] = grossAmtFieldParseInfo.FieldValue.replace(',', '');

  // Withholding Tax
  let witholdTaxFieldParseInfo = readColNoticeOfCashDividendMessageField_(mailBody, 
    grossAmtFieldParseInfo.FieldEnd + marker.CommonEnd.length, 
    marker.Field7Start, marker.CommonMid, marker.CommonEnd);
  messageInfo['WithholdingTax'] = witholdTaxFieldParseInfo.FieldValue.replace(',', '');

  // Net Amount
  let netAmtFieldParseInfo = readColNoticeOfCashDividendMessageField_(mailBody, 
    witholdTaxFieldParseInfo.FieldEnd + marker.CommonEnd.length, 
    marker.Field8Start, marker.CommonMid, marker.CommonEnd);
  messageInfo['NetAmount'] = netAmtFieldParseInfo.FieldValue.replace(',', '');

  return messageInfo;
}

function readColNoticeOfCashDividendMessageField_(mailBody, startIndex, startMarker, midMarker, endMarker) {
  let fieldParseInfo = {};
  fieldParseInfo['FieldStart'] = mailBody.indexOf(startMarker, startIndex);
  fieldParseInfo['FieldMid'] = mailBody.indexOf(midMarker, fieldParseInfo.FieldStart + startMarker.length);
  fieldParseInfo['FieldEnd'] = mailBody.indexOf(endMarker, fieldParseInfo.FieldMid + midMarker.length);
  let rawValue = mailBody.substring(fieldParseInfo.FieldMid + midMarker.length, fieldParseInfo.FieldEnd);
  
  // clean the value if contains tag
  let cleanedValue = rawValue;
  let endTagIndex = cleanedValue.lastIndexOf(">");
  if (endTagIndex > -1) {
    cleanedValue = cleanedValue.substring(endTagIndex + 1);
  }
  let endSemiColonIndex = cleanedValue.lastIndexOf(";");
  if (endSemiColonIndex > -1) {
    cleanedValue = cleanedValue.substring(endSemiColonIndex + 1);
  }

  fieldParseInfo['FieldValue'] = cleanedValue;

  return fieldParseInfo;
}

// Common
const readColNoticeOfCashDivMsgType1CheckerStartText = '<span style="background-color: #FFFFFF">';
const readColNoticeOfCashDivMsgType2CheckerStartText = '<span style="FONT-FAMILY: Arial; COLOR: black; FONT-SIZE: 10pt"><b>';
const readColNoticeOfCashDivMsgType3CheckerStartText = '<span style="FONT-FAMILY:Arial;COLOR:black;FONT-SIZE:10pt"><b>';

const readColNoticeOfCashDivMsgType1Common1MidText = '<span style="background-color: #FFFFFF">';
const readColNoticeOfCashDivMsgType2Common1MidText = '</span><b>';
const readColNoticeOfCashDivMsgType3Common1MidText = '</span><b>';
const readColNoticeOfCashDivMsgType4Common1MidText = '">';

const readColNoticeOfCashDivMsgType1Common1EndText = '</span>';
const readColNoticeOfCashDivMsgType2Common1EndText = '</b>';
const readColNoticeOfCashDivMsgType3Common1EndText = '</b>';
const readColNoticeOfCashDivMsgType4Common1EndText = '</span>';

// Payment Date
const readColNoticeOfCashDivMsgType1Fld1StartText = '<span style="background-color: #FFFFFF">';
const readColNoticeOfCashDivMsgType2Fld1StartText = '<span style="FONT-FAMILY: Arial; COLOR: black; FONT-SIZE: 10pt"><b>';
const readColNoticeOfCashDivMsgType3Fld1StartText = '<span style="FONT-FAMILY:Arial;COLOR:black;FONT-SIZE:10pt"><b>';
const readColNoticeOfCashDivMsgType4Fld1StartText = '<span style="font-size: 14px; font-family: Tahoma, Geneva, sans-serif;">';
// Stock Code
const readColNoticeOfCashDivMsgFld2StartText = 'Stock Code';
// Price per share
const readColNoticeOfCashDivMsgFld3StartText = 'Cash Dividend (Php)';
const readColNoticeOfCashDivMsgFld3EndText = '/ share';
// Ex Dividend Date
const readColNoticeOfCashDivMsgFld4StartText = 'Ex-Date';
// No of Share
const readColNoticeOfCashDivMsgFld5StartText = 'No. of Shares Entitled to Cash Dividend';
// Gross Amount
const readColNoticeOfCashDivMsgFld6StartText = 'Gross Amount';
// Withholding Tax
const readColNoticeOfCashDivMsgFld7StartText = 'Less Withholding Tax';
// Net Amount
const readColNoticeOfCashDivMsgFld8StartText = 'Net Amount';

/**
 * Prepare the cash dividend notification email markers 
 * - right now there were 4 major email patterns that were found with marker defined in constants
 * - there are a lot of outlier but cannot handle all of them
 */
function getColNoticeOfCashDividendMessageMarker_(messageIndex, mailBody) {
  let marker = {};

  if (mailBody.indexOf(readColNoticeOfCashDivMsgType1CheckerStartText) > 0) {
    marker['MarkerType'] = 1;
    marker['Field1Start'] = readColNoticeOfCashDivMsgType1Fld1StartText;
    marker['CommonEnd'] = readColNoticeOfCashDivMsgType1Common1EndText;
    marker['CommonMid'] = readColNoticeOfCashDivMsgType1Common1MidText;
  } else if (mailBody.indexOf(readColNoticeOfCashDivMsgType2CheckerStartText) > 0) {
    marker['MarkerType'] = 2;
    marker['Field1Start'] = readColNoticeOfCashDivMsgType2Fld1StartText;
    marker['CommonEnd'] = readColNoticeOfCashDivMsgType2Common1EndText;
    marker['CommonMid'] = readColNoticeOfCashDivMsgType2Common1MidText;
  } else if (mailBody.indexOf(readColNoticeOfCashDivMsgType3CheckerStartText) > 0) {
    marker['MarkerType'] = 3;
    marker['Field1Start'] = readColNoticeOfCashDivMsgType3Fld1StartText;
    marker['CommonEnd'] = readColNoticeOfCashDivMsgType3Common1EndText;
    marker['CommonMid'] = readColNoticeOfCashDivMsgType3Common1MidText;
  } else {
    marker['MarkerType'] = 4;
    marker['Field1Start'] = readColNoticeOfCashDivMsgType4Fld1StartText;
    marker['CommonEnd'] = readColNoticeOfCashDivMsgType4Common1EndText;
    marker['CommonMid'] = readColNoticeOfCashDivMsgType4Common1MidText;
  }
  // Stock Code
  marker['Field2Start'] = readColNoticeOfCashDivMsgFld2StartText;
  // Price per Share
  marker['Field3Start'] = readColNoticeOfCashDivMsgFld3StartText;
  marker['Field3End'] = readColNoticeOfCashDivMsgFld3EndText;
  // Ex Dividend Date
  marker['Field4Start'] = readColNoticeOfCashDivMsgFld4StartText;
  // No of Share
  marker['Field5Start'] = readColNoticeOfCashDivMsgFld5StartText;
  // Gross Amount
  marker['Field6Start'] = readColNoticeOfCashDivMsgFld6StartText;
  // Withholding Tax
  marker['Field7Start'] = readColNoticeOfCashDivMsgFld7StartText;
  // Net Amount
  marker['Field8Start'] = readColNoticeOfCashDivMsgFld8StartText;

  //console.log(marker);
  return marker;
}

/**
 * Check if the COL Notice of Cash Dividend is already logged in the CashDividend excel sheet
 * - Payment Date, Stock Code is used for checking
 */
function checkColNoticeOfCashDividendExists_(dataCashDividend, aNoticeInfo) {
  let paymentDate = (aNoticeInfo.TransactionInfo.PaymentDate == null) ? aNoticeInfo.MailDate : aNoticeInfo.TransactionInfo.PaymentDate;

  for (let iRow = dataCashDividend.length - 1; iRow >= 0; iRow--) {
    let stockCode = dataCashDividend[iRow][COLINDEX_CASHDIV_STOCKCODE];
    let paymentDateText = new Date(dataCashDividend[iRow][COLINDEX_CASHDIV_PAYMENTDATE]);
    let formattedPaymentDate = Utilities.formatDate(new Date(paymentDateText), 'Asia/Tokyo', 'yyyy-MM-dd');

    if (stockCode == '') {
      continue;
    }

    if (aNoticeInfo.TransactionInfo.StockCode != stockCode) {
      continue;
    }
    
    if (paymentDate != formattedPaymentDate) {
      continue;
    }

    return true;
  }
  return false;
}

/**
 * Retrieve the last price of the said stock before the ex-dividend date using pse edge api
 * - limitation is some stocks without security key cannot have their price computed like jfcpb
 */
function getCashDividendNoticeSharePrice_(latestPriceInfo, dataCurrentPrice, aNoticeInfo) {

  if (!(aNoticeInfo.TransactionInfo.StockCode in latestPriceInfo)) {
    latestPriceInfo[aNoticeInfo.TransactionInfo.StockCode] = {};

    // Get company id, security id and latest price info
    for (let iRow = 0; iRow < dataCurrentPrice.length; iRow++) {
      let stockCode = dataCurrentPrice[iRow][COLINDEX_CURRENTPRICE_STOCKCODE];
      if (stockCode == aNoticeInfo.TransactionInfo.StockCode) {
        latestPriceInfo[aNoticeInfo.TransactionInfo.StockCode]['CompanyId'] = dataCurrentPrice[iRow][COLINDEX_CURRENTPRICE_COMPANYID];
        latestPriceInfo[aNoticeInfo.TransactionInfo.StockCode]['SecurityId'] = dataCurrentPrice[iRow][COLINDEX_CURRENTPRICE_SECURITYID];
        latestPriceInfo[aNoticeInfo.TransactionInfo.StockCode]['LatestPrice'] = dataCurrentPrice[iRow][COLINDEX_CURRENTPRICE_LATESTPRICE];
        latestPriceInfo[aNoticeInfo.TransactionInfo.StockCode]['LatestPriceDate'] = 
          Utilities.formatDate(new Date(dataCurrentPrice[iRow][COLINDEX_CURRENTPRICE_LATESTPRICEDATE]), 'Asia/Manila', 'yyyy-MM-dd');
        break;
      }
    }
  }
  let currLastPriceInfo = latestPriceInfo[aNoticeInfo.TransactionInfo.StockCode];

  // TODO If latest price is the ex div date, maybe can not call price api already
  let exDividendDateDt = new Date(aNoticeInfo.TransactionInfo.ExDividendDate);

  // If past year, last price of the year
  let companyId = currLastPriceInfo.CompanyId;
  let securityId = currLastPriceInfo.SecurityId;

  let fromDateDt = new Date(exDividendDateDt.getTime()-(7*(24*3600*1000)));
  let toDateDt = new Date(exDividendDateDt.getTime()+(7*(24*3600*1000)));

  let fromDateText = Utilities.formatDate(fromDateDt, "GMT+9", "MM-dd-yyyy");
  let toDateText = Utilities.formatDate(toDateDt, "GMT+9", "MM-dd-yyyy");

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
  if (responseCode != 200) {
    console.log('Response failed: ' + aNoticeInfo.TransactionInfo.StockCode + ' + ' + aNoticeInfo.TransactionInfo.ExDividendDate);
    console.log(response.getResponseCode());
    console.log(response.getContentText());
    return {};
  }

  let json_response = JSON.parse(response);
  let chartData = json_response['chartData'];
  if (chartData.length <= 0) {
    console.log('Empty Chart Data: ' + aNoticeInfo.TransactionInfo.StockCode + ' + ' + aNoticeInfo.TransactionInfo.ExDividendDate);
    console.log(JSON.stringify(formData));
    console.log(response.getContentText());
    return {};
  }

  //console.log(chartData);
  let exDividendDateCompare = new Date(aNoticeInfo.TransactionInfo.ExDividendDate);
  let prevDateChartInfo = null;
  for (let chartId = 0; chartId < chartData.length; chartId++) {
    let aChartDataDate = new Date(chartData[chartId].CHART_DATE);
    //console.log('Compare '+ aChartDataDate + ' ' + exDividendDateCompare);
    if (aChartDataDate.getTime() > exDividendDateCompare.getTime()) {
      // Use the last value (possible no data for the ex-dividend date)
      aNoticeInfo.TransactionInfo['PricePerShare'] = prevDateChartInfo.PricePerShare;
      aNoticeInfo.TransactionInfo['PriceDate'] = prevDateChartInfo.PriceDate;
      return chartData[chartId];
    }
    // Remember the previous price info, in case no chart info for the exact ex-dividend date
    prevDateChartInfo = {};
    prevDateChartInfo['PricePerShare'] = chartData[chartId].CLOSE;
    prevDateChartInfo['PriceDate'] = Utilities.formatDate(aChartDataDate, "GMT+9", "MM-dd-yyyy");
  }
  return {};
}

/**
 * Write an individual CashDividend from the top
 * - a new row is inserted
 * - since email values are formatted with comma, before parsing value need to remove the comma
 */
function logCashDividendNoticeToSheets_(colSettings, addToRow, aNoticeInfo, currentPriceLastRow) {
  let sheetCashDiv = SpreadsheetApp.getActive().getSheetByName(SHEET_CASHDIV);
  sheetCashDiv.insertRowBefore(addToRow);
  console.log('Row added to ' + addToRow);

  let cellRange = '';
  let formulaVal = '';

  cellRange = colSettings.StockCode + addToRow;
  sheetCashDiv.getRange(cellRange).setValue(aNoticeInfo.TransactionInfo.StockCode);

  // Dates
  let paymentDate = (aNoticeInfo.TransactionInfo.PaymentDate == null) ? aNoticeInfo.MailDate : aNoticeInfo.TransactionInfo.PaymentDate;
  cellRange = colSettings.PaymentDate + addToRow;
  sheetCashDiv.getRange(cellRange).setValue(paymentDate);
  cellRange = colSettings.ExDividendDate + addToRow;
  sheetCashDiv.getRange(cellRange).setValue(aNoticeInfo.TransactionInfo.ExDividendDate);

  cellRange = colSettings.DayDiff + addToRow;
  formulaVal = "=DAYS(" + colSettings.PaymentDate + addToRow + "," + colSettings.ExDividendDate + addToRow + ")";
  sheetCashDiv.getRange(cellRange).setFormula(formulaVal);

  cellRange = colSettings.DividendPerShare + addToRow;
  sheetCashDiv.getRange(cellRange).setValue(aNoticeInfo.TransactionInfo.DividendPerShare);
  cellRange = colSettings.Quantity + addToRow;
  sheetCashDiv.getRange(cellRange).setValue(aNoticeInfo.TransactionInfo.NoOfShare);

  cellRange = colSettings.GrossAmount + addToRow;
  sheetCashDiv.getRange(cellRange).setValue(aNoticeInfo.TransactionInfo.GrossAmount);
  cellRange = colSettings.WithholdingTax + addToRow;
  sheetCashDiv.getRange(cellRange).setValue(aNoticeInfo.TransactionInfo.WithholdingTax);
  cellRange = colSettings.NetAmount + addToRow;
  sheetCashDiv.getRange(cellRange).setValue(aNoticeInfo.TransactionInfo.NetAmount);

  cellRange = colSettings.PriceDate + addToRow;
  sheetCashDiv.getRange(cellRange).setValue(aNoticeInfo.TransactionInfo.PriceDate);
  cellRange = colSettings.PricePerShare + addToRow;
  sheetCashDiv.getRange(cellRange).setValue(aNoticeInfo.TransactionInfo.PricePerShare);

  cellRange = colSettings.DividendRate + addToRow;
  formulaVal = "=" + colSettings.DividendPerShare + addToRow + "/" + colSettings.PricePerShare + addToRow;
  sheetCashDiv.getRange(cellRange).setFormula(formulaVal);
  sheetCashDiv.getRange(cellRange).setBackground("yellow");

  // Sector
  cellRange = colSettings.Sector + addToRow;
  formulaVal = "=LOOKUP(" + colSettings.StockCode + addToRow + ",CurrentPrice!$A$3:$A$" + 
    currentPriceLastRow + ",CurrentPrice!$C$3:$C$" + currentPriceLastRow + ")";
  sheetCashDiv.getRange(cellRange).setFormula(formulaVal);

  cellRange = colSettings.Broker + addToRow;
  sheetCashDiv.getRange(cellRange).setValue('COL');

  cellRange = colSettings.AverageBuyPricePerShare + addToRow;
  sheetCashDiv.getRange(cellRange).setValue(aNoticeInfo.AverageBuyPricePerShare);

  cellRange = colSettings.AverageBuyDividendRate + addToRow;
  formulaVal = "=" + colSettings.DividendPerShare + addToRow + "/" + colSettings.AverageBuyPricePerShare + addToRow;
  sheetCashDiv.getRange(cellRange).setFormula(formulaVal);

  SpreadsheetApp.flush();
}
