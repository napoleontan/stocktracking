/**
 * Search Gmail to see the last 5 trading confirmation email
 * For each message found, 
 * - parse the email for all transaction and each price per share of transaction
 * - check that the trading info is already logged in the StockInformation sheet.
 * - if it is not yet logged, add a new row from top for each transaction
 */
function searchNewEmailColTradingConfirmation() {

  let colSettings = loadColumnSettings(SHEET_STOCKTRANS);

  let sheetStockTrans = SpreadsheetApp.getActive().getSheetByName(SHEET_STOCKTRANS);
  let [rows, columns] = [sheetStockTrans.getLastRow(), sheetStockTrans.getLastColumn()];
  let dataStockTrans = sheetStockTrans.getRange(3, 1, rows, columns).getValues();
  //console.log(dataStockTrans);

  let addToRow = 3;
  let emailInfo = readColTradingConfirmationGmail_();
  for (let messageIndex in emailInfo) {
    let aMailInfo = emailInfo[messageIndex];
    for (let transactionIndex in aMailInfo.TransactionInfo) {
      let aTransactionInfo = aMailInfo.TransactionInfo[transactionIndex];
      for (let netAmountsIndex in aTransactionInfo.NetAmounts) {
        let aNetAmountInfo = aTransactionInfo.NetAmounts[netAmountsIndex];

        let transactionLogged = checkColTradingConfirmationTransactionExists_(dataStockTrans, aTransactionInfo, aNetAmountInfo);
        if (!transactionLogged) {
          console.log('Un logged Transaction found');
          console.log(aTransactionInfo.TransactionType + ' ' + aTransactionInfo.StockCode + ' ' + aTransactionInfo.TransactionDate);
          console.log('  ' + aNetAmountInfo.Quantity + ' ' + aNetAmountInfo.PricePerShare + ' ' + aNetAmountInfo.GrossAmount);
          logStockTransactionToSheets_(colSettings, addToRow, aTransactionInfo, aNetAmountInfo);
          addToRow++;
        }
      } // end for
    } // end for
  } // end for 
}

/**
 * Search the latest 5 col trading confirm email thread
 */
function readColTradingConfirmationGmail_() {

  let mailInfo = {};
  let messageIndex = 0;

  let searchKey = 'colfinancial subject:"COL Trading Confirmation"';
  let threads = GmailApp.search(searchKey, 0, 10);
  for (var i = 0; i < threads.length; i++) {
    let messages = threads[i].getMessages();
    for (var j = 0;  j < messages.length;  j++) {
      let message = messages[j];

      let messageInfo = {};
      messageInfo['Subject'] = message.getSubject();
      messageInfo['TransactionInfo'] = readColTradingConfirmationMessage_(message);
      mailInfo[messageIndex] = messageInfo;
      messageIndex++;
    }
  }

  //console.log(JSON.stringify(mailInfo));
  return mailInfo;
}

// Common
const readColTradingConfMsgCommon1MidText = '<font face="Arial" size="1">';
const readColTradingConfMsgCommon1EndText = '</font></p>';
const readColTradingConfMsgCommon2MidText = '<font size="1" face="Arial" color="#000000">';
const readColTradingConfMsgCommon2EndText = '</font></td>';

// Transaction Type
const readColTradingConfMsgFld1StartText = '<font size="1" face="Arial">Action:</font></td>';
// Transcation Date
const readColTradingConfMsgFld2StartText = '<font size="1" face="Arial">Trade Date:</font>';
// Symbol
const readColTradingConfMsgFld3StartText = '<font size="1" face="Arial" color="#000000">Symbol:</font>';
// Value Table
const readColTradingConfMsgNetAmtStartText = 'Net Amount</font>';
const readColTradingConfMsgNetAmtFldStartText = '<font face="Arial" size="1">';
const readColTradingConfMsgNetAmtFldEndText = '</font></td>';
// Value Table Fields
const readColTradingConfMsgTotalsStartText = 'Totals</font>';
// VAT total
const readColTradingConfMsgVATStartText = '<font face="Arial" size="1">';
const readColTradingConfMsgVATEndText = 'VAT</font></td>';

/**
 * Read a COL Trading Confirmation email message
 * - A single mail can contain multiple transaction
 * - A single transaction can contain multiple priced per share info
 *  - The VAT is only for the total info and not for each line so need to adjust manually if such happens
 */
function readColTradingConfirmationMessage_(message) {
  let mailBody = message.getBody();
  let transactStartIndex = 0;

  let messageInfo = {};
  let messageIndex = 0;

  // Get Record start for Transaction Type
  let field1Start = mailBody.indexOf(readColTradingConfMsgFld1StartText, transactStartIndex);
  let field1MidEnd = mailBody.indexOf(readColTradingConfMsgCommon1MidText, field1Start + readColTradingConfMsgFld1StartText.length);
  let field1End = mailBody.indexOf(readColTradingConfMsgCommon1EndText, field1MidEnd + readColTradingConfMsgCommon1MidText.length);

  while (field1Start > 0 && field1End > 0) {
    let messageTransactInfo = {};
    let transType = mailBody.substring(field1MidEnd + readColTradingConfMsgCommon1MidText.length, field1End).replace('-', '').replace(/\s/g, '');
    messageTransactInfo['TransactionType'] = transType;

    // Transcation Date
    let field2Start = mailBody.indexOf(readColTradingConfMsgFld2StartText, field1End + readColTradingConfMsgCommon1EndText.length);
    let field2MidEnd = mailBody.indexOf(readColTradingConfMsgCommon1MidText, field2Start + readColTradingConfMsgFld2StartText.length);
    let field2End = mailBody.indexOf(readColTradingConfMsgCommon1EndText, field2MidEnd + readColTradingConfMsgCommon1MidText.length);
    let transDateText = mailBody.substring(field2MidEnd + readColTradingConfMsgCommon1MidText.length, field2End);
    let transactDate = Utilities.formatDate(new Date(transDateText), 'Asia/Tokyo', 'yyyy-MM-dd').replace(/\s/g, '');
    messageTransactInfo['TransactionDate'] = transactDate;

    // Symbol
    let field3Start = mailBody.indexOf(readColTradingConfMsgFld3StartText, field2End + readColTradingConfMsgCommon1EndText.length);
    let field3MidEnd = mailBody.indexOf(readColTradingConfMsgCommon2MidText, field3Start + readColTradingConfMsgFld3StartText.length);
    let field3End = mailBody.indexOf(readColTradingConfMsgCommon2EndText, field3MidEnd + readColTradingConfMsgCommon2MidText.length);
    let stockCode = mailBody.substring(field3MidEnd + readColTradingConfMsgCommon2MidText.length, field3End).replace(/\s/g, '');
    messageTransactInfo['StockCode'] = stockCode;

    // Net Amounts
    let netAmtStart = mailBody.indexOf(readColTradingConfMsgNetAmtStartText, field3End + readColTradingConfMsgCommon2EndText.length);
    let totalsStart = mailBody.indexOf(readColTradingConfMsgTotalsStartText, netAmtStart + readColTradingConfMsgNetAmtStartText.length);

    let fieldIndex = 0;
    let rowIndex = 0;
    let rowValues = {};
    let rowFieldValues = {};
    let fieldStartIndex = mailBody.indexOf(readColTradingConfMsgNetAmtFldStartText, netAmtStart + readColTradingConfMsgNetAmtStartText.length);
    let fieldEndIndex = mailBody.indexOf(readColTradingConfMsgNetAmtFldEndText, fieldStartIndex + readColTradingConfMsgNetAmtFldStartText.length);
    while (fieldStartIndex > 0 && fieldStartIndex < totalsStart) {

      let fieldValue = mailBody.substring(fieldStartIndex + readColTradingConfMsgNetAmtFldStartText.length, fieldEndIndex);
      switch (fieldIndex) {
        case 0:
          rowFieldValues['Quantity'] = fieldValue;
          break;
        case 1:
          rowFieldValues['PricePerShare'] = fieldValue;
          break;
        case 2:
          rowFieldValues['GrossAmount'] = fieldValue;
          break;
        case 3:
          rowFieldValues['Commission'] = fieldValue;
          break;
        case 4:
          rowFieldValues['OtherCharges'] = fieldValue;
          break;
        case 5:
          rowFieldValues['SalesTax'] = fieldValue;
          break;
      }

      // If last field
      if (fieldIndex == 5) {
        rowValues[rowIndex] = rowFieldValues;
        rowFieldValues = {};
        rowIndex++;
        fieldIndex = -1;
      }

      // Iterate to next field
      fieldStartIndex = mailBody.indexOf(readColTradingConfMsgNetAmtFldStartText, fieldEndIndex + readColTradingConfMsgNetAmtFldEndText.length);
      fieldEndIndex = mailBody.indexOf(readColTradingConfMsgNetAmtFldEndText, fieldStartIndex + readColTradingConfMsgNetAmtFldStartText.length);
      fieldIndex++;
    } // while end
    messageTransactInfo['NetAmounts'] = rowValues;

    // VAT total
    let vatTotalEnd = mailBody.indexOf(readColTradingConfMsgVATEndText, totalsStart + readColTradingConfMsgTotalsStartText.length);
    let subStrUntilVatTotalEnd = mailBody.substring(0, vatTotalEnd + readColTradingConfMsgVATEndText.length);
    let vatTotalStart = subStrUntilVatTotalEnd.lastIndexOf(readColTradingConfMsgVATStartText);
    let vatTotal = mailBody.substring(vatTotalStart + readColTradingConfMsgVATStartText.length, vatTotalEnd);
    messageTransactInfo['VatTotal'] = vatTotal;

    // Iterate next Transaction Type
    transactStartIndex = totalsStart + readColTradingConfMsgTotalsStartText.length;
    field1Start = mailBody.indexOf(readColTradingConfMsgFld1StartText, transactStartIndex);
    field1MidEnd = mailBody.indexOf(readColTradingConfMsgCommon1MidText, field1Start + readColTradingConfMsgFld1StartText.length);
    field1End = mailBody.indexOf(readColTradingConfMsgCommon1EndText, field1MidEnd + readColTradingConfMsgCommon1MidText.length);

    messageInfo[messageIndex] = messageTransactInfo;
    messageIndex++;
  } // while end

  //console.log(JSON.stringify(messageInfo));
  return messageInfo;
}

/**
 * Check if the COL Trading Confirmation individual transaction is already logged in the StockTransaction excel sheet
 * - Transaction Type, Transaction Date, Stock Code and Quantity is used for checking
 */
function checkColTradingConfirmationTransactionExists_(dataStockTrans, aTransactionInfo, aNetAmountInfo) {
  for (let iRow = dataStockTrans.length - 1; iRow >= 0; iRow--) {
    let stockCode = dataStockTrans[iRow][COLINDEX_STOCKTRANS_STOCKCODE];
    let transDateText = new Date(dataStockTrans[iRow][COLINDEX_STOCKTRANS_TRANSDATE]);
    let formattedTransactDate = Utilities.formatDate(new Date(transDateText), 'Asia/Tokyo', 'yyyy-MM-dd');
    let quantity = dataStockTrans[iRow][COLINDEX_STOCKTRANS_QUANTITY];
    let transactType = dataStockTrans[iRow][COLINDEX_STOCKTRANS_TRANSTYPE];

    if (stockCode == '') {
      continue;
    }

    if (aTransactionInfo.StockCode != stockCode) {
      continue;
    }
    
    if (aTransactionInfo.TransactionDate != formattedTransactDate) {
      continue;
    }

    if (aTransactionInfo.TransactionType == 'BOUGHT' && transactType != TRANSTYPE_BOUGHTSHARES) {
      continue;
    }
    if (aTransactionInfo.TransactionType == 'SOLD' && transactType != TRANSTYPE_SOLDSHARES) {
      continue;
    }

    // The sold quantity has negative value but email is positive, so need to get absolute value
    // The email has comma, when parsing it consider comma as decima, so need to remove comma
    if (parseInt(aNetAmountInfo.Quantity.replace(/,/g, ''), 10) != Math.abs(parseInt(quantity, 10))) {
      continue;
    }

    return true;
  }

  return false;
}

/**
 * Write an individual StockTransaction from the top
 * - a new row is inserted
 * - fields are set whether it is Bought or Sold shares
 * - since email values are formatted with comma, before parsing value need to remove the comma
 */
function logStockTransactionToSheets_(colSettings, addToRow, aTransactionInfo, aNetAmountInfo) {
  let sheetStockTrans = SpreadsheetApp.getActive().getSheetByName(SHEET_STOCKTRANS);
  sheetStockTrans.insertRowBefore(addToRow);
  console.log('Row added to ' + addToRow);

  let cellRange = '';
  let formulaVal = '';

  cellRange = colSettings.StockCode + addToRow;
  sheetStockTrans.getRange(cellRange).setValue(aTransactionInfo.StockCode);
  cellRange = colSettings.TransactionDate + addToRow;
  sheetStockTrans.getRange(cellRange).setValue(aTransactionInfo.TransactionDate);

  let quantity = parseInt(aNetAmountInfo.Quantity.replace(',', ''), 10);
  let grossAmount = parseFloat(aNetAmountInfo.GrossAmount.replace(',', ''));
  let transactType = '';
  let pricePerShareCol = '';
  let grossAmountCol = '';

  if (aTransactionInfo.TransactionType == 'BOUGHT') {
    transactType = TRANSTYPE_BOUGHTSHARES;
    pricePerShareCol = colSettings.BuyPricePerShare;
    grossAmountCol = colSettings.GrossBuyAmount;
  } else if (aTransactionInfo.TransactionType == 'SOLD') {
    quantity *= -1;
    transactType = TRANSTYPE_SOLDSHARES;
    pricePerShareCol = colSettings.SellPricePerShare;
    grossAmountCol = colSettings.GrossSellAmount;
    grossAmount *= -1;
  }
  cellRange = colSettings.Quantity + addToRow;
  sheetStockTrans.getRange(cellRange).setValue(quantity);
  cellRange = colSettings.TransactionType + addToRow;
  sheetStockTrans.getRange(cellRange).setValue(transactType);
  cellRange = pricePerShareCol + addToRow;
  sheetStockTrans.getRange(cellRange).setValue(aNetAmountInfo.PricePerShare);
  cellRange = grossAmountCol + addToRow;
  sheetStockTrans.getRange(cellRange).setValue(grossAmount);

  // Fees
  cellRange = colSettings.CommissionFee + addToRow;
  sheetStockTrans.getRange(cellRange).setValue(aNetAmountInfo.Commission);

  let vatValue = parseFloat(aTransactionInfo.VatTotal.replace(',', ''));
  cellRange = colSettings.CommissionTax + addToRow;
  sheetStockTrans.getRange(cellRange).setValue(vatValue);
  // Highlight the vat as indication it is a new row, and to check if value is splitted between transacction
  sheetStockTrans.getRange(cellRange).setBackground("yellow");

  cellRange = colSettings.OtherCharges + addToRow;
  sheetStockTrans.getRange(cellRange).setValue(aNetAmountInfo.OtherCharges);

  if (aTransactionInfo.TransactionType == 'SOLD') {
    cellRange = colSettings.SalesTax + addToRow;
    sheetStockTrans.getRange(cellRange).setValue(aNetAmountInfo.SalesTax);
  }

  cellRange = colSettings.AllFess + addToRow;
  formulaVal = "=sum(" + colSettings.CommissionFee + addToRow + ":" + colSettings.SalesTax + addToRow + ")";
  sheetStockTrans.getRange(cellRange).setFormula(formulaVal);

  // Net Total
  let total = 0;
  total += parseFloat(aNetAmountInfo.Commission.replace(',', ''));
  total += vatValue;
  total += parseFloat(aNetAmountInfo.OtherCharges.replace(',', ''));
  total += parseFloat(aNetAmountInfo.SalesTax.replace(',', ''));

  let netAmountCol = '';
  if (aTransactionInfo.TransactionType == 'BOUGHT') {
    netAmountCol = colSettings.NetBuyAmount;
    total += grossAmount;
  } else if (aTransactionInfo.TransactionType == 'SOLD') {
    netAmountCol = colSettings.NetSellAmount;
    total += (-1 * grossAmount);
  }
  cellRange = netAmountCol + addToRow;
  sheetStockTrans.getRange(cellRange).setValue(total);

  // Sector
  let sheetCurrentPrice = SpreadsheetApp.getActive().getSheetByName(SHEET_CURRENTPRICE);
  let luLastRow = sheetCurrentPrice.getLastRow();

  cellRange = colSettings.Sector + addToRow;
  formulaVal = "=LOOKUP(" + colSettings.StockCode + addToRow + ",CurrentPrice!$A$3:$A$" + luLastRow + ",CurrentPrice!$C$3:$C$" + luLastRow + ")";
  sheetStockTrans.getRange(cellRange).setFormula(formulaVal);

  cellRange = colSettings.Broker + addToRow;
  sheetStockTrans.getRange(cellRange).setValue('COL');

  SpreadsheetApp.flush();
}