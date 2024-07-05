// Use this code for Google Docs, Slides, Forms, or Sheets.
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('Extra Charts')
      .addItem('Open Diversification Graph', '_openDiversificationGraphDialog')
      .addToUi();
}

function _openDiversificationGraphDialog() {

  let template = HtmlService
    .createTemplateFromFile('DiversificationGraphDialog');

  let csvData = _createDiversificationGraphDialogCsvData();
  template.mainCsv = csvData.mainCsv;
  template.stockCsv = csvData.stockCsv;
  let html = template.evaluate()
    .setWidth(1000)
    .setHeight(800);

  SpreadsheetApp.getUi()
      .showModalDialog(html, 'Diversification Graph');
}

function _createDiversificationGraphDialogCsvData() {
  // Header
  let mainCsvText = 'Portfolio,Sector,SubSector,StockCode,value';
  let stockCsvText = 'StockCode,CompanyName';

  let colSettings = loadColumnSettings(SHEET_CURRENTPRICE);
  let currentPriceSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_CURRENTPRICE);

  let sectorSummary = {};
  let subSectorSummary = {};
  // Loop once to get data in an array for sorting and compute weight of each sector too
  let graphData = [];
  for (var row = ROWINDEX_CURRENTPRICE_FIRSTROW; row <= currentPriceSheet.getLastRow(); row++) {
    const stockCode = currentPriceSheet.getRange(colSettings.StockCode + row).getValue();
    if (stockCode == '') {
      break;
    }
    const quantity = currentPriceSheet.getRange(colSettings.Quantity + row).getValue();
    if (quantity == 0) {
      continue;
    }
    const companyName = currentPriceSheet.getRange(colSettings.CompanyName + row).getValue();
    const sector = currentPriceSheet.getRange(colSettings.Sector + row).getValue();
    const subSector = currentPriceSheet.getRange(colSettings.SubSector + row).getValue();
    const buyAmountPercentage = currentPriceSheet.getRange(colSettings.BuyAmountPercentage + row).getValue();

    graphData.push({ "stockCode" : stockCode, "companyName": companyName, "sector": sector, "subSector": subSector, "buyAmountPercentage" : Math.round(buyAmountPercentage * 10000) / 100 });
    if (!(sector in sectorSummary)) {
      sectorSummary[sector] = {};
      sectorSummary[sector]['BuyAmountPercentage'] = 0;
    }
    sectorSummary[sector]['BuyAmountPercentage'] = sectorSummary[sector]['BuyAmountPercentage'] + buyAmountPercentage;
    if (!(subSector in subSectorSummary)) {
      subSectorSummary[subSector] = {};
      subSectorSummary[subSector]['BuyAmountPercentage'] = 0;
    }
    subSectorSummary[subSector]['BuyAmountPercentage'] = subSectorSummary[subSector]['BuyAmountPercentage'] + buyAmountPercentage;
  }

  // Sort by descending sector total, then by subsector total amount, then by stock code total amount
  graphData.sort(function (a, b) {
    let compareResult = sectorSummary[b.sector]['BuyAmountPercentage'] - sectorSummary[a.sector]['BuyAmountPercentage'];
    if (compareResult == 0) {
      compareResult = subSectorSummary[b.subSector]['BuyAmountPercentage'] - subSectorSummary[a.subSector]['BuyAmountPercentage'];
    }
    if (compareResult == 0) {
      compareResult = b.buyAmountPercentage - a.buyAmountPercentage;
    } 
    return compareResult;
  });
  //Logger.log(JSON.stringify(graphData));

  for (const aKey in graphData) {
    let aStock = graphData[aKey];
    mainCsvText = mainCsvText + "\nPortfolio," + aStock.sector + "," + aStock.subSector + "," + aStock.stockCode + "," + aStock.buyAmountPercentage;
    stockCsvText = stockCsvText + "\n" + aStock.stockCode + "," + aStock.companyName;
  }
  Logger.log(mainCsvText);
  Logger.log(stockCsvText);

  return { "mainCsv" : mainCsvText, "stockCsv": stockCsvText };
}
