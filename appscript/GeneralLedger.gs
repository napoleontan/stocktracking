function generateLedger() {
  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let currentPriceSheet = activeSpreadsheet.getSheetByName(SHEET_CURRENTPRICE);

  let ledgerSheet = activeSpreadsheet.getSheetByName(SHEET_GENERATEDLEDGER);
  if (!ledgerSheet) {
    ledgerSheet = activeSpreadsheet.insertSheet();
    ledgerSheet.setName(SHEET_GENERATEDLEDGER);
    activeSpreadsheet.moveActiveSheet(currentPriceSheet.getIndex() + 1);
    activeSpreadsheet.setActiveSheet(ledgerSheet);
  }

  let sheetStockTrans = SpreadsheetApp.getActive().getSheetByName(SHEET_STOCKTRANS);10
  let [rows, columns] = [sheetStockTrans.getLastRow(), sheetStockTrans.getLastColumn()];
  let dataStockTrans = sheetStockTrans.getRange(1, 1, rows, columns).getValues();

  let sheetCashDividend = SpreadsheetApp.getActive().getSheetByName(SHEET_CASHDIV);
  [rows, columns] = [sheetCashDividend.getLastRow(), sheetCashDividend.getLastColumn()];
  let dataCashDividend = sheetCashDividend.getRange(1, 1, rows, columns).getValues();

  let sheetMoneyIn = SpreadsheetApp.getActive().getSheetByName(SHEET_MONEYIN);
  [rows, columns] = [sheetMoneyIn.getLastRow(), sheetMoneyIn.getLastColumn()];
  let dataMoneyIn = sheetMoneyIn.getRange(1, 1, rows, columns).getValues();

  let nextRecordExists = true;
  let nextDateToProcess = null;
  let currentMoneyAmount = 0;

  let currMoneyInIndex = 2;
  let currStockTransIndex = dataStockTrans.length - 1;
  let currCashDivIndex = dataCashDividend.length - 1;
  let ledgerRow = 0;

  let futureDate = new Date();
  futureDate.setDate(futureDate.getDate() + 1);
  let colSettings = loadColumnSettings(SHEET_GENERATEDLEDGER);

  while (nextRecordExists) {

    if (nextDateToProcess == null) {
      // this loop only checks the next date to process by checking minimum between the 3
      // If all are past the limit set flag to end loop

      let moneyInDate = futureDate;
      if (currMoneyInIndex < dataMoneyIn.length) {
        moneyInDate = new Date(dataMoneyIn[currMoneyInIndex][COLINDEX_MONEYIN_TRANSACTIONDATE]);
      }
      let stockTransDate = futureDate;
      if (currStockTransIndex >= 2) {
        stockTransDate = new Date(dataStockTrans[currStockTransIndex][COLINDEX_STOCKTRANS_TRANSDATE]);
      }
      let paymentDate = futureDate;
      if (currCashDivIndex >= 2) {
        paymentDate = new Date(dataCashDividend[currCashDivIndex][COLINDEX_CASHDIV_PAYMENTDATE]);
      }
      //console.log('Money In: ' + currMoneyInIndex + ' stockTrans: ' + currStockTransIndex + ' cashpay: ' + currCashDivIndex);
      //console.log('Money In: ' + moneyInDate + ' stockTrans: ' + stockTransDate + ' cashpay: ' + paymentDate);
      //console.log('Money In: ' + moneyInDate.getTime() + ' stockTrans: ' + stockTransDate.getTime() + ' cashpay: ' + paymentDate.getTime());

      if (stockTransDate.getTime() >= moneyInDate.getTime() && paymentDate.getTime() >= moneyInDate.getTime()) {
        nextDateToProcess = moneyInDate;
        //console.log('smallest is moneyin ');
      } else if (moneyInDate.getTime() >= stockTransDate.getTime() && paymentDate.getTime() >= stockTransDate.getTime()) {
        nextDateToProcess = stockTransDate;
        //console.log('smallest is strocktrans ');
      } else {
        nextDateToProcess = paymentDate;
        //console.log('smallest is cashdiv payment ');
      }
      if (nextDateToProcess.getTime() == futureDate.getTime()) {
        nextRecordExists = false;
      }
    } else {
      // Search items that are for the current date

      // Check the deposits for the day
      for (let moneyInIndex = currMoneyInIndex; moneyInIndex < dataMoneyIn.length; moneyInIndex++) {
        let moneyInDate = new Date(dataMoneyIn[moneyInIndex][COLINDEX_MONEYIN_TRANSACTIONDATE]);
        let rowAmount = dataMoneyIn[moneyInIndex][COLINDEX_MONEYIN_AMOUNT];

        if (moneyInDate.getTime() == nextDateToProcess.getTime()) {
          currentMoneyAmount += rowAmount;
          currentMoneyAmount = Math.round((currentMoneyAmount + Number.EPSILON) * 100) / 100;
          ledgerRow++;
          writeMoneyIn_(ledgerSheet, colSettings, ledgerRow, moneyInDate, rowAmount, currentMoneyAmount);
          //console.log(ledgerRow + '. MoneyIn Dep Date ' + moneyInDate + ' Add amount ' + rowAmount + ' Balance ' + currentMoneyAmount);
          currMoneyInIndex++;
        } else {
          break;
        }
      }

      // Check the transactions for the day
      // Bottom to top
      for (let stockTransIndex = currStockTransIndex; stockTransIndex >= 0; stockTransIndex--) {
        let stockCode = dataStockTrans[stockTransIndex][COLINDEX_STOCKTRANS_STOCKCODE];
        let transactType = dataStockTrans[stockTransIndex][COLINDEX_STOCKTRANS_TRANSTYPE];
        let transactionDate = new Date(dataStockTrans[stockTransIndex][COLINDEX_STOCKTRANS_TRANSDATE]);
        if (transactionDate.getTime() == nextDateToProcess.getTime()) {
          let rowAmount = 0;
          if (isBuyTransactionType(transactType) || transactType == TRANSTYPE_STOCKDIVIDEND) {
            rowAmount = dataStockTrans[stockTransIndex][COLINDEX_STOCKTRANS_NETBUYAMT];
          } else {
            rowAmount = dataStockTrans[stockTransIndex][COLINDEX_STOCKTRANS_NETSELLAMT];
          }
          currentMoneyAmount += (-1 * rowAmount);
          currentMoneyAmount = Math.round((currentMoneyAmount + Number.EPSILON) * 100) / 100;
          ledgerRow++;
          //console.log(ledgerRow + '. Stock Trans Date ' + transactionDate + ' ' + stockCode + ' ' + transactType + ' amount ' + rowAmount + ' Balance ' + currentMoneyAmount);
          writeStockTransaction_(ledgerSheet, colSettings, ledgerRow, transactionDate, transactType, stockCode, rowAmount, currentMoneyAmount);
          currStockTransIndex--;
        } else {
          break;
        }
      }

      // Check the cash dividend for the day
      // Bottom to top
      for (let cashDividendIndex = currCashDivIndex; cashDividendIndex >= 0; cashDividendIndex--) {
        let stockCode = dataCashDividend[cashDividendIndex][COLINDEX_CASHDIV_STOCKCODE];
        let paymentDate = new Date(dataCashDividend[cashDividendIndex][COLINDEX_CASHDIV_PAYMENTDATE]);
        let rowAmount = dataCashDividend[cashDividendIndex][COLINDEX_CASHDIV_NETAMT];
        if (paymentDate.getTime() == nextDateToProcess.getTime()) {
          currentMoneyAmount += rowAmount;
          currentMoneyAmount = Math.round((currentMoneyAmount + Number.EPSILON) * 100) / 100;
          ledgerRow++;
          //console.log(ledgerRow + '. CashDiv Pay Date ' + paymentDate + ' ' + stockCode + ' div amount ' + rowAmount + ' Balance ' + currentMoneyAmount);
          writeCashDividend_(ledgerSheet, colSettings, ledgerRow, paymentDate, stockCode, rowAmount, currentMoneyAmount);
          currCashDivIndex--;
        } else {
          break;
        }
      }

      // Reset this so next loop will find the minimum of the 3
      nextDateToProcess = null;
    }

  } // end while
  console.log('ledgerRow ' + ledgerRow);
}

function writeMoneyIn_(ledgerSheet, colSettings, ledgerRow, moneyInDate, rowAmount, currentMoneyAmount) {
  let rowNo = ledgerRow + 1;

  ledgerSheet.getRange(colSettings.No + rowNo).setValue(ledgerRow);
  ledgerSheet.getRange(colSettings.TransactDate + rowNo).setValue(moneyInDate);
  ledgerSheet.getRange(colSettings.TransactType + rowNo).setValue(TRANSTYPE_DEPOSIT);
  ledgerSheet.getRange(colSettings.Credit + rowNo).setValue(rowAmount);
  ledgerSheet.getRange(colSettings.Balance + rowNo).setValue(currentMoneyAmount);
}

function writeStockTransaction_(ledgerSheet, colSettings, ledgerRow, transactionDate, transactType, stockCode, rowAmount, currentMoneyAmount) {
  let rowNo = ledgerRow + 1;

  ledgerSheet.getRange(colSettings.No + rowNo).setValue(ledgerRow);
  ledgerSheet.getRange(colSettings.TransactDate + rowNo).setValue(transactionDate);
  ledgerSheet.getRange(colSettings.TransactType + rowNo).setValue(transactType);
  ledgerSheet.getRange(colSettings.StockCode + rowNo).setValue(stockCode);
  if (isBuyTransactionType(transactType) || transactType == TRANSTYPE_STOCKDIVIDEND) {
    // Bought, Stock Rights
    ledgerSheet.getRange(colSettings.Debit + rowNo).setValue(rowAmount);
  } else {
    // Sold
    ledgerSheet.getRange(colSettings.Credit + rowNo).setValue(-1 * rowAmount);
  }
  ledgerSheet.getRange(colSettings.Balance + rowNo).setValue(currentMoneyAmount);
}

function writeCashDividend_(ledgerSheet, colSettings, ledgerRow, paymentDate, stockCode, rowAmount, currentMoneyAmount) {
  let rowNo = ledgerRow + 1;

  ledgerSheet.getRange(colSettings.No + rowNo).setValue(ledgerRow);
  ledgerSheet.getRange(colSettings.TransactDate + rowNo).setValue(paymentDate);
  ledgerSheet.getRange(colSettings.TransactType + rowNo).setValue(TRANSTYPE_CASHDIVIDEND);
  ledgerSheet.getRange(colSettings.StockCode + rowNo).setValue(stockCode);
  ledgerSheet.getRange(colSettings.Credit + rowNo).setValue(rowAmount);
  ledgerSheet.getRange(colSettings.Balance + rowNo).setValue(currentMoneyAmount);
}
