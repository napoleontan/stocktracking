const disclousureUrl = 'https://edge.pse.com.ph/common/DisclosureCht.ax';
const companyPageUrl = 'https://edge.pse.com.ph/companyPage/stockData.do?cmpy_id=';
const companyInfoPageUrl = 'https://frames.pse.com.ph/security/';

// Dividend Withholding Tax Rate
const CASHDIV_WITHTAX_NETRATE = 0.9;
// Market Price should deduct all the commission and taxes = 0.99105 (or minus 0.00895)
// https://www.colfinancial.com/ape/final2/home/faq.asp
const COL_MARKETPRICE_NETRATE = 0.99105;

const TRANSTYPE_BOUGHTSHARES = 'Bought Shares';
const TRANSTYPE_SOLDSHARES = 'Sold Shares';
const TRANSTYPE_STOCKRIGHTS = 'Stock Rights';
const TRANSTYPE_STOCKDIVIDEND = 'Stock Dividend';
const TRANSTYPE_IPOBUYSHARES = 'IPO Buy Shares'
const TRANSTYPE_CASHDIVIDEND = 'Cash Dividend';
const TRANSTYPE_PROPERTYDIVIDEND = 'Property Dividend';

// Excel Columns
const SHEET_CURRENTPRICE = "CurrentPrice";
const COLINDEX_CURRENTPRICE_STOCKCODE = 0;
const COLINDEX_CURRENTPRICE_COMPANYNAME = 1;
const COLINDEX_CURRENTPRICE_SECTOR = 2;
const COLINDEX_CURRENTPRICE_COMPANYID = 3;
const COLINDEX_CURRENTPRICE_SECURITYID = 4;
const COLINDEX_CURRENTPRICE_LATESTPRICE = 5;
const COLINDEX_CURRENTPRICE_LATESTPRICEDATE = 6;
const COLINDEX_CURRENTPRICE_QUANTITY = 7;
const COLINDEX_CURRENTPRICE_AVGPRICEPERSHARE = 8;

const ROWINDEX_CURRENTPRICE_CHECKBOX = 1;
const ROWINDEX_CURRENTPRICE_SELECTEDSTOCKCODE = 2;
const ROWINDEX_CURRENTPRICE_FIRSTROW = 4;

const SHEET_STOCKTRANS = "StockTransaction";
const COLINDEX_STOCKTRANS_STOCKCODE = 0;
const COLINDEX_STOCKTRANS_TRANSDATE = 1;
const COLINDEX_STOCKTRANS_QUANTITY = 2;
const COLINDEX_STOCKTRANS_YEARENDQTY = 3;
const COLINDEX_STOCKTRANS_TRANSTYPE = 4;
const COLINDEX_STOCKTRANS_BUYPRICEPERSHARE = 5;
const COLINDEX_STOCKTRANS_SELLPRICEPERSHARE = 6;
const COLINDEX_STOCKTRANS_GROSSBUYAMT = 7;
const COLINDEX_STOCKTRANS_GROSSSELLAMT = 8;
const COLINDEX_STOCKTRANS_COMMISSIONFEE = 9;
const COLINDEX_STOCKTRANS_COMMISSIONTAX = 10;
const COLINDEX_STOCKTRANS_OTHERCHARGES = 11;
const COLINDEX_STOCKTRANS_SALESTAX = 12;
const COLINDEX_STOCKTRANS_ALLFEES = 13;
const COLINDEX_STOCKTRANS_NETBUYAMT = 14;
const COLINDEX_STOCKTRANS_NETSELLAMT = 15;
const COLINDEX_STOCKTRANS_SECTOR = 16;
const COLINDEX_STOCKTRANS_BROKER = 17;

const SHEET_TRANSTOTAL = "ST1";
const COLINDEX_TRANSTOTAL_STOCKCODE = 0;
const COLINDEX_TRANSTOTAL_QUANTITY = 4;
const COLINDEX_TRANSTOTAL_PRICEPERSHARE = 5;

const SHEET_CASHDIV = "CashDividend";
const COLINDEX_CASHDIV_STOCKCODE = 0;
const COLINDEX_CASHDIV_PAYMENTDATE = 1;
const COLINDEX_CASHDIV_EXDIVDATE = 2;
// No of days diff of Ex-Dividend Date vs Payment Date
const COLINDEX_CASHDIV_DAYDIFF = 3;
const COLINDEX_CASHDIV_DIVPERSHARE = 4;
const COLINDEX_CASHDIV_QUANTITY = 5;
const COLINDEX_CASHDIV_GROSSAMT = 6;
const COLINDEX_CASHDIV_WTAX = 7;
const COLINDEX_CASHDIV_NETAMT = 8;
const COLINDEX_CASHDIV_PRICEDATE = 9;
const COLINDEX_CASHDIV_PRICEPERSHARE = 10;
const COLINDEX_CASHDIV_DIVRATE = 11;
const COLINDEX_CASHDIV_SECTOR = 12;
const COLINDEX_CASHDIV_BROKER = 13;
const COLINDEX_CASHDIV_AVGBUYPRICEPERSHARE = 14;
const COLINDEX_CASHDIV_AVGBUYDIVRATE = 15;

const SHEET_YEARLYMARKETVALUE = "YearVal";
const COLINDEX_YEARLYMARKETVALUE_STOCKCODE = 0;

const SHEET_PROJDIV = 'ProjDiv';
const COLINDEX_PROJDIV_PAGENO = 0;
const COLINDEX_PROJDIV_STOCKCODE = 1;
const COLINDEX_PROJDIV_EXDIVDATE = 2;
const COLINDEX_PROJDIV_PAYMENTDATE = 3;
const COLINDEX_PROJDIV_DIVPERSHARE = 4;
const COLINDEX_PROJDIV_QUANTITY = 5;
const COLINDEX_PROJDIV_GROSSAMT = 6;
const COLINDEX_PROJDIV_NETAMT = 7;
const COLINDEX_PROJDIV_DIVRATE = 8;
const COLINDEX_PROJDIV_PRICEPERSHARE = 9;
const COLINDEX_PROJDIV_MARKETPRICE = 10;
const COLINDEX_PROJDIV_YEARRATE = 11;
const COLINDEX_PROJDIV_YEARAMOUNT = 12;
const COLINDEX_PROJDIV_SECTOR = 13;
const COLINDEX_PROJDIV_YEARSHARE = 14;

const ROWINDEX_PROJDIV_FIRSTROW = 3;

const SHEET_STOCKANALYZE = 'SA';
const COLINDEX_STOCKANALYZE_TRANSDATE = 0;
const COLINDEX_STOCKANALYZE_TRANSTYPE = 1;
const COLINDEX_STOCKANALYZE_BUYQTY = 2;
const COLINDEX_STOCKANALYZE_BUYPRICEPERSHARE = 3;
const COLINDEX_STOCKANALYZE_NETBUYAMOUNT = 4;
const COLINDEX_STOCKANALYZE_SELLQTY = 5;
const COLINDEX_STOCKANALYZE_SELLPRICEPERSHARE = 6;
const COLINDEX_STOCKANALYZE_SELLBUYAMOUNT = 7;
const COLINDEX_STOCKANALYZE_DIVPERSHARE = 8;
const COLINDEX_STOCKANALYZE_NETDIVAMOUNT = 9;
const COLINDEX_STOCKANALYZE_CLOSINGPRICEPERSHARE = 10;
const COLINDEX_STOCKANALYZE_DIVRATE = 11;
const COLINDEX_STOCKANALYZE_CURRQTY = 12;
const COLINDEX_STOCKANALYZE_AVGPRICEPERSHARE = 13;
const COLINDEX_STOCKANALYZE_BUYVSAVEPRICE = 14;
const COLINDEX_STOCKANALYZE_BUYVSAVEPRICEDIFF = 15;
const COLINDEX_STOCKANALYZE_SELLVSAVEPRICE = 16;
const COLINDEX_STOCKANALYZE_SELLVSAVEPRICEDIFF = 17;
const COLINDEX_STOCKANALYZE_ACTUALDIVRATE = 18;
const COLINDEX_STOCKANALYZE_DIVRATEDIFF = 19;
const COLINDEX_STOCKANALYZE_ANNUALACTUALDIVRATE = 20;

const COLTITLE_STOCKANALYZE_TRANSDATE = 'TransactDate';
const COLTITLE_STOCKANALYZE_TRANSTYPE = 'TransactType';
const COLTITLE_STOCKANALYZE_BUYQTY = 'BuyQty';
const COLTITLE_STOCKANALYZE_BUYPRICEPERSHARE = 'BuyPricePerShare';
const COLTITLE_STOCKANALYZE_NETBUYAMOUNT = 'NetBuyAmount';
const COLTITLE_STOCKANALYZE_SELLQTY = 'SellQty';
const COLTITLE_STOCKANALYZE_SELLPRICEPERSHARE = 'SellPricePerShare';
const COLTITLE_STOCKANALYZE_NETSELLAMOUNT = 'NetSellAmount';
const COLTITLE_STOCKANALYZE_DIVPERSHARE = 'DividendPerShare';
const COLTITLE_STOCKANALYZE_NETDIVAMOUNT = 'NetDividendAmt';
const COLTITLE_STOCKANALYZE_CLOSINGPRICEPERSHARE = 'ClosePrice';
const COLTITLE_STOCKANALYZE_DIVRATE = 'DivRate';
const COLTITLE_STOCKANALYZE_CURRQTY = 'Qty';
const COLTITLE_STOCKANALYZE_AVGPRICEPERSHARE = 'AvgPricePerShare';
const COLTITLE_STOCKANALYZE_BUYVSAVEPRICE = 'BuyVsAvePrice';
const COLTITLE_STOCKANALYZE_BUYVSAVEPRICEDIFF = 'Diff';
const COLTITLE_STOCKANALYZE_SELLVSAVEPRICE = 'SellVsAvePrice';
const COLTITLE_STOCKANALYZE_SELLVSAVEPRICEDIFF = 'Diff';
const COLTITLE_STOCKANALYZE_ACTUALDIVRATE = 'ActualDivRate';
const COLTITLE_STOCKANALYZE_DIVRATEDIFF = 'Diff';
const COLTITLE_STOCKANALYZE_ANNUALACTUALDIVRATE = 'AnnualDivRate';

const COLFORMAT_STOCKANALYZE_TRANSDATE = 'yyyy"-"mm"-"dd';
const COLFORMAT_STOCKANALYZE_TRANSTYPE = '#,##0';
const COLFORMAT_STOCKANALYZE_BUYQTY = '#,##0';
const COLFORMAT_STOCKANALYZE_BUYPRICEPERSHARE = '[$₱]#,##0.0000';
const COLFORMAT_STOCKANALYZE_NETBUYAMOUNT = '[$₱]#,##0.00';
const COLFORMAT_STOCKANALYZE_SELLQTY = '#,##0';
const COLFORMAT_STOCKANALYZE_SELLPRICEPERSHARE = '[$₱]#,##0.0000';
const COLFORMAT_STOCKANALYZE_NETSELLAMOUNT = '[$₱]#,##0.00';
const COLFORMAT_STOCKANALYZE_DIVPERSHARE = '[$₱]#,##0.0000';
const COLFORMAT_STOCKANALYZE_NETDIVAMOUNT = '[$₱]#,##0.00';
const COLFORMAT_STOCKANALYZE_CLOSINGPRICEPERSHARE = '[$₱]#,##0.0000';
const COLFORMAT_STOCKANALYZE_DIVRATE = '0.0000%';
const COLFORMAT_STOCKANALYZE_CURRQTY = '#,##0';
const COLFORMAT_STOCKANALYZE_AVGPRICEPERSHARE = '[$₱]#,##0.0000';
const COLFORMAT_STOCKANALYZE_BUYVSAVEPRICE = '[$₱]#,##0.0000';
const COLFORMAT_STOCKANALYZE_BUYVSAVEPRICEDIFF = '0.00%';
const COLFORMAT_STOCKANALYZE_SELLVSAVEPRICE = '[$₱]#,##0.0000';
const COLFORMAT_STOCKANALYZE_SELLVSAVEPRICEDIFF = '0.00%';
const COLFORMAT_STOCKANALYZE_ACTUALDIVRATE = '0.0000%';
const COLFORMAT_STOCKANALYZE_DIVRATEDIFF = '0.0000%';
const COLFORMAT_STOCKANALYZE_ANNUALACTUALDIVRATE = '0.0000%';

const SHEET_BACKUPPRICE = 'BackupPrice';
const COLINDEX_BACKUPPRICE_STOCKCODE = 1;
const COLINDEX_BACKUPPRICE_COMPANYNAME = 2;
const COLINDEX_BACKUPPRICE_PRICE = 4;

const SHEET_BACKUPDIVIDEND = 'BackupDividend';
const COLINDEX_BACKUPDIVIDEND_STOCKCODE = 1;
const COLINDEX_BACKUPDIVIDEND_DIVIDENDTYPE = 2;
const COLINDEX_BACKUPDIVIDEND_DIVPERSHARE = 3;
const COLINDEX_BACKUPDIVIDEND_EXDIVDATE = 4;
const COLINDEX_BACKUPDIVIDEND_RECORDDATE = 5;
const COLINDEX_BACKUPDIVIDEND_PAYMENTDATE = 6;

function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function isBuyTransactionType(transactType) {
  return transactType == TRANSTYPE_BOUGHTSHARES || transactType == TRANSTYPE_IPOBUYSHARES || transactType == TRANSTYPE_STOCKRIGHTS || transactType == TRANSTYPE_PROPERTYDIVIDEND;
}

function loadColumnSettings(sheetName) {
  if (sheetName == SHEET_CURRENTPRICE) {
    return {
      'StockCode': columnToLetter(COLINDEX_CURRENTPRICE_STOCKCODE + 1),
      'CompanyName': columnToLetter(COLINDEX_CURRENTPRICE_COMPANYNAME + 1),
      'Sector': columnToLetter(COLINDEX_CURRENTPRICE_SECTOR + 1),
      'CompanyId': columnToLetter(COLINDEX_CURRENTPRICE_COMPANYID + 1),
      'SecurityId': columnToLetter(COLINDEX_CURRENTPRICE_SECURITYID + 1),
      'LatestPrice': columnToLetter(COLINDEX_CURRENTPRICE_LATESTPRICE + 1),
      'LatestPriceDate': columnToLetter(COLINDEX_CURRENTPRICE_LATESTPRICEDATE + 1),
      'Quantity': columnToLetter(COLINDEX_CURRENTPRICE_QUANTITY + 1),
      'AveragePricePerShare': columnToLetter(COLINDEX_CURRENTPRICE_AVGPRICEPERSHARE + 1),
    };
  }
  if (sheetName == SHEET_STOCKTRANS) {
    return {
      'StockCode': columnToLetter(COLINDEX_STOCKTRANS_STOCKCODE + 1),
      'TransactionDate': columnToLetter(COLINDEX_STOCKTRANS_TRANSDATE + 1),
      'Quantity': columnToLetter(COLINDEX_STOCKTRANS_QUANTITY + 1),
      'YearEndQuantity': columnToLetter(COLINDEX_STOCKTRANS_YEARENDQTY + 1),

      'TransactionType': columnToLetter(COLINDEX_STOCKTRANS_TRANSTYPE + 1),
      'BuyPricePerShare': columnToLetter(COLINDEX_STOCKTRANS_BUYPRICEPERSHARE + 1),
      'SellPricePerShare': columnToLetter(COLINDEX_STOCKTRANS_SELLPRICEPERSHARE + 1),
      'GrossBuyAmount': columnToLetter(COLINDEX_STOCKTRANS_GROSSBUYAMT + 1),
      'GrossSellAmount': columnToLetter(COLINDEX_STOCKTRANS_GROSSSELLAMT + 1),

      'CommissionFee': columnToLetter(COLINDEX_STOCKTRANS_COMMISSIONFEE + 1),
      'CommissionTax': columnToLetter(COLINDEX_STOCKTRANS_COMMISSIONTAX + 1),
      'OtherCharges': columnToLetter(COLINDEX_STOCKTRANS_OTHERCHARGES + 1),
      'SalesTax': columnToLetter(COLINDEX_STOCKTRANS_SALESTAX + 1),
      'AllFess': columnToLetter(COLINDEX_STOCKTRANS_ALLFEES + 1),
      
      'NetBuyAmount': columnToLetter(COLINDEX_STOCKTRANS_NETBUYAMT + 1),
      'NetSellAmount': columnToLetter(COLINDEX_STOCKTRANS_NETSELLAMT + 1),
      'Sector': columnToLetter(COLINDEX_STOCKTRANS_SECTOR + 1),
      'Broker': columnToLetter(COLINDEX_STOCKTRANS_BROKER + 1),
    };
  }
  if (sheetName == SHEET_PROJDIV) {
    return {
      'PageNo': columnToLetter(COLINDEX_PROJDIV_PAGENO + 1),
      'StockCode': columnToLetter(COLINDEX_PROJDIV_STOCKCODE + 1),
      'ExDividendDate': columnToLetter(COLINDEX_PROJDIV_EXDIVDATE + 1),
      'PaymentDate': columnToLetter(COLINDEX_PROJDIV_PAYMENTDATE + 1),
      'DividendPerShare': columnToLetter(COLINDEX_PROJDIV_DIVPERSHARE + 1),
      'Quantity': columnToLetter(COLINDEX_PROJDIV_QUANTITY + 1),
      'GrossAmount': columnToLetter(COLINDEX_PROJDIV_GROSSAMT + 1),
      'NetAmount': columnToLetter(COLINDEX_PROJDIV_NETAMT + 1),
      'DividendRate': columnToLetter(COLINDEX_PROJDIV_DIVRATE + 1),
      'PricePerShare': columnToLetter(COLINDEX_PROJDIV_PRICEPERSHARE + 1),
      'MarketPrice': columnToLetter(COLINDEX_PROJDIV_MARKETPRICE + 1),
      'YearRate': columnToLetter(COLINDEX_PROJDIV_YEARRATE + 1),
      'YearAmount': columnToLetter(COLINDEX_PROJDIV_YEARAMOUNT + 1),
      'Sector': columnToLetter(COLINDEX_PROJDIV_SECTOR + 1),
      'YearShare': columnToLetter(COLINDEX_PROJDIV_YEARSHARE + 1),
    };
  }
  if (sheetName == SHEET_CASHDIV) {
    return {
      'StockCode': columnToLetter(COLINDEX_CASHDIV_STOCKCODE + 1),
      'PaymentDate': columnToLetter(COLINDEX_CASHDIV_PAYMENTDATE + 1),
      'ExDividendDate': columnToLetter(COLINDEX_CASHDIV_EXDIVDATE + 1),
      'DayDiff': columnToLetter(COLINDEX_CASHDIV_DAYDIFF + 1),
      'DividendPerShare': columnToLetter(COLINDEX_CASHDIV_DIVPERSHARE + 1),

      'Quantity': columnToLetter(COLINDEX_CASHDIV_QUANTITY + 1),
      'GrossAmount': columnToLetter(COLINDEX_CASHDIV_GROSSAMT + 1),
      'WithholdingTax': columnToLetter(COLINDEX_CASHDIV_WTAX + 1),
      'NetAmount': columnToLetter(COLINDEX_CASHDIV_NETAMT + 1),
      'PriceDate': columnToLetter(COLINDEX_CASHDIV_PRICEDATE + 1),
      'PricePerShare': columnToLetter(COLINDEX_CASHDIV_PRICEPERSHARE + 1),
      'DividendRate': columnToLetter(COLINDEX_CASHDIV_DIVRATE + 1),
      'Sector': columnToLetter(COLINDEX_CASHDIV_SECTOR + 1),
      'Broker': columnToLetter(COLINDEX_CASHDIV_BROKER + 1),
      'AverageBuyPricePerShare': columnToLetter(COLINDEX_CASHDIV_AVGBUYPRICEPERSHARE + 1),
      'AverageBuyDividendRate': columnToLetter(COLINDEX_CASHDIV_AVGBUYDIVRATE + 1),
    };
  }
  if (sheetName == SHEET_STOCKANALYZE) {
    return {
      'TransactionDate': columnToLetter(COLINDEX_STOCKANALYZE_TRANSDATE + 1),
      'TransactionType': columnToLetter(COLINDEX_STOCKANALYZE_TRANSTYPE + 1),

      'BuyQuantity': columnToLetter(COLINDEX_STOCKANALYZE_BUYQTY + 1),
      'BuyPricePerShare': columnToLetter(COLINDEX_STOCKANALYZE_BUYPRICEPERSHARE + 1),
      'NetBuyAmount': columnToLetter(COLINDEX_STOCKANALYZE_NETBUYAMOUNT + 1),

      'SellQuantity': columnToLetter(COLINDEX_STOCKANALYZE_SELLQTY + 1),
      'SellPricePerShare': columnToLetter(COLINDEX_STOCKANALYZE_SELLPRICEPERSHARE + 1),
      'NetSellAmount': columnToLetter(COLINDEX_STOCKANALYZE_SELLBUYAMOUNT + 1),

      'DividendPerShare': columnToLetter(COLINDEX_STOCKANALYZE_DIVPERSHARE + 1),
      'NetDividendAmount': columnToLetter(COLINDEX_STOCKANALYZE_NETDIVAMOUNT + 1),
      'ClosingPricePerShare': columnToLetter(COLINDEX_STOCKANALYZE_CLOSINGPRICEPERSHARE + 1),
      'DividendRate': columnToLetter(COLINDEX_STOCKANALYZE_DIVRATE + 1),

      'CurrentQuantity': columnToLetter(COLINDEX_STOCKANALYZE_CURRQTY + 1),
      'AveragePricePerShare': columnToLetter(COLINDEX_STOCKANALYZE_AVGPRICEPERSHARE + 1),

      'BuyVsAveragePrice': columnToLetter(COLINDEX_STOCKANALYZE_BUYVSAVEPRICE + 1),
      'BuyVsAveragePriceDiff': columnToLetter(COLINDEX_STOCKANALYZE_BUYVSAVEPRICEDIFF + 1),
      'SellVsAveragePrice': columnToLetter(COLINDEX_STOCKANALYZE_SELLVSAVEPRICE + 1),
      'SellVsAveragePriceDiff': columnToLetter(COLINDEX_STOCKANALYZE_SELLVSAVEPRICEDIFF + 1),
      'ActualDividendRate': columnToLetter(COLINDEX_STOCKANALYZE_ACTUALDIVRATE + 1),
      'DividendRateDiff': columnToLetter(COLINDEX_STOCKANALYZE_DIVRATEDIFF + 1),
      'AnnualActualDivRate': columnToLetter(COLINDEX_STOCKANALYZE_ANNUALACTUALDIVRATE + 1),
    };
  }
  if (sheetName == SHEET_BACKUPPRICE) {
    return {
      'StockCode': columnToLetter(COLINDEX_BACKUPPRICE_STOCKCODE + 1),
      'CompanyName': columnToLetter(COLINDEX_BACKUPPRICE_COMPANYNAME + 1),
      'Price': columnToLetter(COLINDEX_BACKUPPRICE_PRICE + 1),
    };
  };
  if (sheetName == SHEET_BACKUPDIVIDEND) {
    return {
      'StockCode': columnToLetter(COLINDEX_BACKUPDIVIDEND_STOCKCODE + 1),
      'DividendType': columnToLetter(COLINDEX_BACKUPDIVIDEND_DIVIDENDTYPE + 1),
      'DividendPerShare': columnToLetter(COLINDEX_BACKUPDIVIDEND_DIVPERSHARE + 1),
      'ExDividendDate': columnToLetter(COLINDEX_BACKUPDIVIDEND_EXDIVDATE + 1),
      'RecordDate': columnToLetter(COLINDEX_BACKUPDIVIDEND_RECORDDATE + 1),
      'PaymentDate': columnToLetter(COLINDEX_BACKUPDIVIDEND_PAYMENTDATE + 1),
    };
  };
  return null;
}

getStockShareEquallyPerYear = function(stockCode) {
  let colSettings = loadColumnSettings(SHEET_PROJDIV);
  //console.log('isStockShareEquallyPerYear = ' + stockCode);

  let projDivSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_PROJDIV);
  for (let row = ROWINDEX_PROJDIV_FIRSTROW; row <= projDivSheet.getLastRow(); row++) {
    let rowStockCode = projDivSheet.getRange(colSettings.StockCode + row).getValue();
    //console.log('   rowStockCode = ' + rowStockCode);
    if (rowStockCode != stockCode) {
      continue;
    }
    let yearShare = projDivSheet.getRange(colSettings.YearShare + row).getValue();
    let yearShareNumber = Math.round(yearShare * 100);
    //console.log('Year Share: ' + yearShareNumber);
    if (yearShareNumber == 25 || yearShareNumber == 50 || yearShareNumber == 100) {
      return yearShareNumber / 100;
    }

  } // for loop
  return null;
}
