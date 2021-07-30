function setup() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Portfolio Menu')
      .addItem('Refresh','refreshPortfolio')
      .addToUi();
  
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet = activeSpreadsheet.getSheetByName("portfolio");

  if(sheet == null) {
    Logger.log("Portfolio sheet is missing. Creating it...");
    sheet = activeSpreadsheet.insertSheet();
    sheet.setName("portfolio");

    sheet.getRange("A1").setValue("Ticker");
    sheet.getRange("B1").setValue("Quantity");
  }

  settingsSheet = activeSpreadsheet.getSheetByName("settings");

  if (settingsSheet == null) {
    Logger.log("Settings sheet is missing. Creating it...");
    settingsSheet = activeSpreadsheet.insertSheet();
    settingsSheet.setName("settings");

    settingsSheet.getRange("A1").setValue("CoinMarketCap API Key =>");
    settingsSheet.getRange("C1").setValue("get an API key from https://pro.coinmarketcap.com/signup");
    settingsSheet.getRange("A2").setValue("Currency conversion #1 =>");
    settingsSheet.getRange("B2").setValue("BTC");
    settingsSheet.getRange("A3").setValue("Currency conversion #2 =>");
    settingsSheet.getRange("B3").setValue("USD");
    settingsSheet.getRange("A4").setValue("Currency conversion #3 =>");
    settingsSheet.getRange("B4").setValue("EUR");
    settingsSheet.getRange("A5").setValue("Currency conversion #4 =>");
    settingsSheet.getRange("A6").setValue("Currency conversion #5 =>");

    settingsSheet.autoResizeColumns(1, 3);
  }

  historySheet = activeSpreadsheet.getSheetByName("history");
  
  if (historySheet == null) {
    Logger.log("History sheet is missing. Creating it...");
    historySheet = activeSpreadsheet.insertSheet();
    historySheet.setName("history");
  }

  sheet.activate();
}

const tickersColumn = "A";
const tickersRowStart = 2;
const quantityColumn = "B";
const quantityRowStart = "B";
var tickers = [];
var tickersCount;
var cmcApiKey;
var currencies = [];
var currenciesCount;
var sheet;
var settingsSheet;

function refreshPortfolio() {

  getAllVariables();
  cleanSheet();

  if (currenciesCount == 0 || tickersCount == 0) {
    Logger.log("No tickers or no currencies specified.");
    throw new Error( "No tickers or no currencies specified." );
  } else if (cmcApiKey == "") {
    throw new Error( "Please add the CoinMarketCap API key in the sheet 'settings', cell B2." );
  }

  // Call CMC and populate the data.  
  for (var counter = 0; counter < currenciesCount; counter = counter +1) {
    // Fetch prices and populate the table.
    // + 4 accounts for the columns ticker, quantity and an empty column.
    getPrices(currencies[counter], counter + 4);

    // Add grand totals title.
    sheet.getRange(tickersCount + 3, currenciesCount + 4).setValue("Grand Totals:"); 

    // Compute the totals per currency.
    computeTotalsPerCurrency(currencies[counter], counter + 4);
  }

  computePortfolioAllocation();
  applyFormatting();
  createChart();
  saveToHistory();
}

function getAllVariables() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet = activeSpreadsheet.getSheetByName("portfolio");
  settingsSheet = activeSpreadsheet.getSheetByName("settings");
  historySheet = activeSpreadsheet.getSheetByName("history");

  // Get the tickers.
  var range = sheet.getRange(tickersColumn + tickersRowStart + ":" + tickersColumn);
  var values = range.getValues().filter(String);

  for (var row in values) {
    for (var col in values[row]) {
      tickers.push(values[row][col]);
    }
  }

  tickersCount = tickers.length;
  Logger.log("Found " + tickersCount + " tickers: " + tickers);

  // Get the CMC API key.
  cmcApiKey = settingsSheet.getRange("B1").getValue();

  // Get the conversion currencies.
  var currenciesRange = settingsSheet.getRange("B2:B");
  var currenciesValues = currenciesRange.getValues().filter(String);
  for (var row in currenciesValues) {
    for (var col in currenciesValues[row]) {
      currencies.push(currenciesValues[row][col]);
    }
  }
  currenciesCount = currencies.length;
  Logger.log("Found " + currenciesCount + " currencies: " + currencies);

}

function cleanSheet() {
  var range = sheet.getRange(1, 3, sheet.getLastRow(), sheet.getLastColumn() + 3);
  Logger.log("range to clean: " + range.getA1Notation());
  range.clearContent();

  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearFormat();  
}

function applyFormatting() {
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .setFontSize(10)
    .setFontFamily("Verdana");
  
  var firstRowRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  firstRowRange.setFontWeight("bold");
  firstRowRange.setBackgroundColor("#9BC4E2");
  firstRowRange.setHorizontalAlignment("center");

  var grandTotalRange = sheet.getRange(tickersCount + 3, currenciesCount + 4, 1, sheet.getLastColumn());
  grandTotalRange.setFontWeight("bold");

  var allocationColumnRange = sheet.getRange(2, 2 * currenciesCount + 5, tickersCount, 1);
  allocationColumnRange.setFontWeight("bold");
}

function computeTotalsPerCurrency(currency, currencyColumn) {

  var columnToPopulate = currencyColumn + currenciesCount + 1;
  
  // Add title for total per currency.
  sheet.getRange(1, columnToPopulate).setValue("Totals in " + currency); 

  for (var counter = 0; counter < tickersCount; counter = counter +1) {
    var cellToPopulate = sheet.getRange(counter + 2, columnToPopulate).getA1Notation();    
    var quantityCell = "B" + (counter + 2);
    var priceCell = sheet.getRange(counter + 2, currencyColumn).getA1Notation();
    
    var formula = "=" + quantityCell + "*" + priceCell;
    Logger.log("cell: " + cellToPopulate + " - formula: " + formula);

    sheet.getRange(cellToPopulate).setFormula(formula).setNumberFormat(getNumberFormatForCurrency(currency));
  }

  // Grand total per currency.
  var grandTotalCellToPopulate = sheet.getRange(tickersCount + 3, columnToPopulate).getA1Notation();
  var grandTotalRange = sheet.getRange(2, columnToPopulate, tickersCount).getA1Notation();
  var grandTotalFormula = "=SUM(" + grandTotalRange + ")"
  Logger.log("grandTotalCellToPopulate: " + grandTotalCellToPopulate);
  Logger.log("grandTotalFormula: " + grandTotalFormula);  
  sheet.getRange(grandTotalCellToPopulate).setFormula(grandTotalFormula);  
}

function getPrices(currency, currencyColumn) {
  var options = {
    'headers': { 'X-CMC_PRO_API_KEY': cmcApiKey}
  };

  // Call the CMC API
  var response = UrlFetchApp.fetch("https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest?symbol=" + tickers.join() + "&convert=" + currency, options);

  // Parse the JSON reply
  var json = response.getContentText();
  var data = JSON.parse(json);

  // Add a title to the price column.
  sheet.getRange(1, currencyColumn).setValue("Price in " + currency);  

  // Populate the prices.
  for (var counter = 0; counter < tickers.length; counter = counter +1) {
    var ticker = tickers[counter];    
    // + 2 here accounts for the title row.
    sheet.getRange(counter + 2, currencyColumn).setValue(data["data"][ticker]["quote"][currency]["price"]).setNumberFormat(getNumberFormatForCurrency(currency));  
  }
}

function createChart() {

    var charts = sheet.getCharts();
    for(var i = 0; i < charts.length; i++){
      sheet.removeChart(charts[i]);
    }

    var startRow = tickersRowStart;
    var endRow = tickersCount + 1;
    var dataColumn = 4 + currenciesCount + 1;

    var chart = sheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(sheet.getRange(tickersColumn + startRow + ":" + tickersColumn + endRow))                
        .addRange(sheet.getRange(sheet.getRange(startRow, dataColumn).getA1Notation() + ":" + sheet.getRange(endRow, dataColumn).getA1Notation()))
        .setPosition(endRow + 3, 2, 0, 0)
        .setOption('is3D', true)
        .build();
    sheet.insertChart(chart);
}

function computePortfolioAllocation() {  

  var columnForAllocation = 3 + currenciesCount + 1 + currenciesCount + 1;

  var dataColumn = 4 + currenciesCount + 1;
  var grandTotalRow = 1 + tickersCount + 2;
  var grandTotalCell = sheet.getRange(grandTotalRow, dataColumn).getA1Notation();  
  Logger.log("Grand total cell: " + grandTotalCell);

  sheet.getRange(1, columnForAllocation).setValue("Allocation");  
  for (var i = 1; i <= tickers.length; i = i +1) {
    var cellWithData = sheet.getRange(i + 1, dataColumn).getA1Notation();
    var cellWithAveragedData = sheet.getRange(i + 1, columnForAllocation).getA1Notation();
    Logger.log("cellWithData/cellWithAveragedData: " + cellWithData + "->" + cellWithAveragedData);
    
    //var valueOfTicker = sheet.getRange(cellWithData).getValue();
    sheet.getRange(cellWithAveragedData).setFormula("=" + cellWithData + "/" + grandTotalCell).setNumberFormat("0.00%");
  }

}

function saveToHistory() {
  // Create a date object for the current date and time.
  var now = new Date();

  var sourceRange = sheet.getDataRange();
  Logger.log("Portfolio range: " + sourceRange.getA1Notation());
  
  var lastRowInTargetSheet = historySheet.getLastRow();
  var rowWithDate,rowWithCopiedData = 0; 
  if(lastRowInTargetSheet == 0) {
    rowWithDate = 2; 
    rowWithCopiedData = 4; // 2 rows for space.
  } else {
    rowWithDate = lastRowInTargetSheet + 3; // 2 rows for space.
    rowWithCopiedData = lastRowInTargetSheet + 5; // 2 rows for space.
  }

  var targetRange = historySheet.getRange("A" + rowWithCopiedData);
  Logger.log(targetRange.getA1Notation());
  Logger.log("getlastrow: " + lastRowInTargetSheet);
  Logger.log("getdatarange: " + historySheet.getDataRange().getA1Notation());

  historySheet.getRange("A" + rowWithDate).setValue("Date: " + now);
  sourceRange.copyTo(targetRange);
}

function getNumberFormatForCurrency(currency) {
  switch(currency) {
    case "BTC":
      return "₿#,##0.00000000";
    case "USD":
    case "AUD":
    case "CAD":
    case "NZD":
      return "$#,##0.00";
    case "GBP":
      return "£#,##0.00";
    case "EUR":
      return "€#,##0.00";
    case "CHF":
      return "CHF#,##0.00";
    case "CNY":
    case "JPY":
      return "¥#,##0.00";
    case "SEK":
      return "kr#,##0.00";
    case "ILS":
      return "₪#,##0.00";  
    default:
      return "#,##0.00";
  }
}
