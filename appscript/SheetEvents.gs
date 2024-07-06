/**
 * Method that workaround the trigger of button via a checkbox for mobile phone use
 */
function onEdit(e) {
  const rg = e.range;
  if (rg.getSheet().getName() != SHEET_CURRENTPRICE) {
    return;
  }

  let colSettings = loadColumnSettings(SHEET_CURRENTPRICE);

  // When you check the analyze stock checkbox, then read the target stock and analyze and  open a new sheet
  let stockCodeAnalyzeCheckboxCell = colSettings.StockCode + ROWINDEX_CURRENTPRICE_CHECKBOX;
  let stockCodeAnalyzeTargetCell = colSettings.StockCode + ROWINDEX_CURRENTPRICE_SELECTEDSTOCKCODE;
  if (rg.getA1Notation() === stockCodeAnalyzeCheckboxCell && rg.isChecked() && rg.getSheet().getName() === SHEET_CURRENTPRICE) {
    let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let currentPriceSheet = activeSpreadsheet.getSheetByName(SHEET_CURRENTPRICE);
    let originalStockCode = currentPriceSheet.getRange(stockCodeAnalyzeTargetCell).getValue();
    if (originalStockCode != originalStockCode.toUpperCase()) {
      currentPriceSheet.getRange(stockCodeAnalyzeTargetCell).setValue(originalStockCode.toUpperCase());
    }
    analyzeStockCode(originalStockCode.toUpperCase());
    rg.uncheck();
  }
  // When you check the average buy price per share checkbox, compute the average buy price per share
  let computeBuyAveragePriceCheckboxCell = colSettings.AveragePricePerShare + ROWINDEX_CURRENTPRICE_CHECKBOX;
  if (rg.getA1Notation() === computeBuyAveragePriceCheckboxCell && rg.isChecked() && rg.getSheet().getName() === SHEET_CURRENTPRICE) {
    computeAverageBuyPrice();
    rg.uncheck();
  }

}

/**
 * Method that tracks the selected stock code and saves it to a special cell for use on analyze stock checkbox
 */
function onSelectionChange(e) {
  // Set background to red if a single empty cell is selected.
  const rg = e.range;
  if (rg.getSheet().getName() != SHEET_CURRENTPRICE) {
    return;
  }
  console.log(rg.getSheet().getName());

  let colSettings = loadColumnSettings(SHEET_CURRENTPRICE);
  let stockCodeAnalyzeTargetCell = colSettings.StockCode + ROWINDEX_CURRENTPRICE_SELECTEDSTOCKCODE;
  console.log(stockCodeAnalyzeTargetCell);

  if (rg.getA1Notation().indexOf(colSettings.StockCode) != 0) {
    return;
  }

  // Ignore selecting the title
  let rowNum = parseInt(rg.getA1Notation().replace(colSettings.StockCode, ''), 10);
  if (rowNum >= ROWINDEX_CURRENTPRICE_FIRSTROW && rg.getSheet().getName() === SHEET_CURRENTPRICE) {
    let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let currentPriceSheet = activeSpreadsheet.getSheetByName(SHEET_CURRENTPRICE);
    currentPriceSheet.getRange(stockCodeAnalyzeTargetCell).setValue(rg.getValue());
  }
}