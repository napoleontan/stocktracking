/**
 * Retrieve the latest stock price from the PSE.com.ph stock information page
 */
function getLatestPsePrice() {
  let colSettings = loadColumnSettings(SHEET_CURRENTPRICE);

  let currentPriceSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_CURRENTPRICE);
  for (var row = ROWINDEX_CURRENTPRICE_FIRSTROW; row <= currentPriceSheet.getLastRow(); row++) {
    const stockCode = currentPriceSheet.getRange(colSettings.StockCode + row).getValue();
    if (stockCode == '') {
      break;
    }
    let quantity = currentPriceSheet.getRange(colSettings.Quantity + row).getValue();
    if (quantity == 0) {
      continue;
    }

    currentPriceSheet.getRange(colSettings.LatestPrice + row).setBackground("yellow");
    currentPriceSheet.getRange(colSettings.LatestPriceDate + row).setBackground("yellow");
    SpreadsheetApp.flush();

    const companyId = currentPriceSheet.getRange(colSettings.CompanyId + row).getValue();
    console.log(companyInfoPageUrl + stockCode);
    let response = UrlFetchApp.fetch(companyInfoPageUrl + stockCode);
    let responseCode = response.getResponseCode();
    console.log(response.getResponseCode());
    if (responseCode != 200) {
      continue;
    }
    let html = response.getContentText();
    //console.log(html);

    let searchString = '<h3 class="last-price">';
    var index1 = html.search(searchString);
    console.log("index1=" + index1);
    if (index1 >= 0) {
      let searchString2 = '</h3>';
      var pos1 = index1 + searchString.length;
      var index2 = html.indexOf(searchString2, pos1);
      console.log("index2=" + index2);
      if (index2 >= 0) {
        var latestPrice = html.substring(pos1, index2).replace(/\s/g, "");
        console.log(latestPrice);
        currentPriceSheet.getRange(colSettings.LatestPrice + row).setValue(latestPrice);
      }
    } // index1 >=0

    let searchString3 = 'As of ';
    var index3 = html.search(searchString3);
    console.log("index3=" + index3);
    if (index3 >= 0) {
      let searchString4 = '</div>';
      var pos3 = index3 + searchString3.length;
      var index4 = html.indexOf(searchString4, pos3);
      console.log("index4=" + index4);
      if (index4 >= 0) {
        var latestPriceDate = html.substring(pos3, index4);
        console.log(latestPriceDate);
        currentPriceSheet.getRange(colSettings.LatestPriceDate + row).setValue(latestPriceDate);
      }
    } // index3 >=0
    
    currentPriceSheet.getRange(colSettings.LatestPrice + row).setBackground(null);
    currentPriceSheet.getRange(colSettings.LatestPriceDate + row).setBackground(null);
  } // for loop
}


function updatePriceFromBackup() {
  let currentPriceColSettings = loadColumnSettings(SHEET_CURRENTPRICE);

  let sheetBackupPrice = SpreadsheetApp.getActive().getSheetByName(SHEET_BACKUPPRICE);
  let [rows, columns] = [sheetBackupPrice.getLastRow(), sheetBackupPrice.getLastColumn()];
  let dataBackupPrice = sheetBackupPrice.getRange(2, 1, rows, columns).getValues();

  let nowTime = Utilities.formatDate(new Date(), 'Asia/Manila', 'yy-MM-dd HH:mm');
  let currentPriceSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_CURRENTPRICE);
  for (var row = ROWINDEX_CURRENTPRICE_FIRSTROW; row <= currentPriceSheet.getLastRow(); row++) {
    let stockCode = currentPriceSheet.getRange(currentPriceColSettings.StockCode + row).getValue();
    let currentPrice = currentPriceSheet.getRange(currentPriceColSettings.LatestPrice + row).getValue();
    if (stockCode == '') {
      break;
    }
    let quantity = currentPriceSheet.getRange(currentPriceColSettings.Quantity + row).getValue();
    if (quantity == 0) {
      continue;
    }

    currentPriceSheet.getRange(currentPriceColSettings.LatestPrice + row).setBackground("yellow");
    SpreadsheetApp.flush();

    let priceChangeFound = false;
    for (let iRow = 0; iRow < dataBackupPrice.length; iRow++) {
      let backupStockCode = dataBackupPrice[iRow][COLINDEX_BACKUPPRICE_STOCKCODE];
      let backupPrice = dataBackupPrice[iRow][COLINDEX_BACKUPPRICE_PRICE];
      if (backupStockCode != stockCode) {
        continue;
      } else {
        console.log('StockCode: ' + stockCode + ' backupPrice: ' + backupPrice + ' vs nowPrice: ' + currentPrice);
        if (backupPrice != currentPrice) {
          currentPriceSheet.getRange(currentPriceColSettings.LatestPrice + row).setValue(backupPrice);
          console.log(nowTime);
          currentPriceSheet.getRange(currentPriceColSettings.LatestPriceDate + row).setValue(nowTime);
          priceChangeFound = true;
        }
        break;
      }
    }
    if (!priceChangeFound) {
      //console.log('StockCode: ' + stockCode + ' backupPrice: none vs nowPrice: ' + currentPrice);
      currentPriceSheet.getRange(currentPriceColSettings.LatestPrice + row).setBackground(null);
    }

  } // for loop
}