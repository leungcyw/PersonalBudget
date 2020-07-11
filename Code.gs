/**
 * Google App Script code to manage my budget input and visual display
 * Functions interact with Google Forms for data input
 * Functions interact with Google Sheets for data storage and display
 */

// Constants for Google Form data spreadsheet
const RAW_DATASHEET_NAME = 'Form Responses';
const RAW_DATASHEET_TIMESHEET_COL = 0;
const RAW_DATASHEET_CATEGORY_COL = 1;
const RAW_DATASHEET_DESCRIPTION_COL = 2;
const RAW_DATASHEET_AMOUNT = 3;

// Constants for stored data spreadsheet
const DATASTORE_DATASHEET_NAME = 'Data';
const DATASTORE_DEFAULT_MEASUREMENTS_CELL = 'A1';
const DATASTORE_TEMP_CELL = 'D1';
const DATASTORE_TEMP_RANGE = 'D1:D';
const DATASTORE_TEMP_ROW = 1;
const DATASTORE_TEMP_COL1 = 4;
const DATASTORE_TEMP_COL2 = 5;
const DATASTORE_TEMP_TWO_COLS = 2;
const DATASTORE_MONTH_KEY_RANGE = 'B1:B';
const DATASTORE_MONTH_VAL_RANGE = 'C1:C';
const DATASTORE_MONTH_KEY_COL = 2;
const DATASTORE_MONTH_VAL_COL = 3;

// Constants for dashboard spreadsheet
const DASHBOARD_DATASHEET_NAME = 'Dashboard';
const DASHBOARD_DATE_EXPENSE_CELL = 'B2';
const DASHBOARD_MONTH_EXPENSE_CELL = 'B3';
const DASHBOARD_YEAR_EXPENSE_CELL = 'B4';
const DASHBOARD_MONTH_INCOME_CELL = 'B5';
const DASHBOARD_YEAR_INCOME_CELL = 'B6';
const DASHBOARD_MONTH_BALANCE_CELL = 'B7';
const DASHBOARD_YEAR_BALANCE_CELL = 'B8';
const DASHBOARD_TOTAL_BALANCE_CELL = 'B9';

// Query and formatting constants
const CURRENCY_FORMAT = "$#,##0.00";
const POSITIVE_AMOUNT_COLOR = 'green';
const NEGATIVE_AMOUNT_COLOR = 'red';
const FORM_QUERY_TARGET = '\'Form Responses\'!A:D';
const FORM_DATE_EXPENSE_SQL = 'select D where A >= date \'"&TEXT(TODAY(), "yyyy-mm-dd")&"\' and B != \'Income\'';
const FORM_MONTH_EXPENSE_SQL = 'select D where MONTH(A) = MONTH(date \'"&TEXT(TODAY(), "yyyy-mm-dd")&"\') and YEAR(A) = YEAR(date \'"&TEXT(TODAY(), "yyyy-mm-dd")&"\') and B != \'Income\'';
const FORM_YEAR_EXPENSE_SQL = 'select D where YEAR(A) = YEAR(date \'"&TEXT(TODAY(), "yyyy-mm-dd")&"\') and B != \'Income\'';
const FORM_MONTH_INCOME_SQL = 'select D where MONTH(A) = MONTH(date \'"&TEXT(TODAY(), "yyyy-mm-dd")&"\') and YEAR(A) = YEAR(date \'"&TEXT(TODAY(), "yyyy-mm-dd")&"\') and B = \'Income\'';
const FORM_YEAR_INCOME_SQL = 'select D where YEAR(A) = YEAR(date \'"&TEXT(TODAY(), "yyyy-mm-dd")&"\') and B = \'Income\'';
const FORM_MONTH_CATEGORY_EXPENSE_SQL = 'select B,D where MONTH(A) = MONTH(date \'"&TEXT(TODAY(), "yyyy-mm-dd")&"\') and YEAR(A) = YEAR(date \'"&TEXT(TODAY(), "yyyy-mm-dd")&"\') and B != \'Income\'';

// Budget Overview table constants
const TALBE_NUM_ROWS = 9;

// Graph constants
const DATASTORE_MONTH_GRAPH_ROW_POS = 1;
const DATASTORE_MONTH_GRAPH_COL_POS = 4;
const MONTH_GRAPH_TITLE = 'Current Month Expenses';
const MONTH_GRAPH_COL_RANGE = [4, 8];


/**
 * Creates a Trigger to update the dashboard whenever the Google Form is submitted
 * NOTE: MUST BE MANUALLY CALLED TO CREATE THE TRIGGER 
 */
function createUpdateTrigger() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('updateDashboard')
  .forSpreadsheet(spreadsheet)
  .onFormSubmit()
  .create();
}

/**
 * Deletes all currently active Triggers
 * NOTE: MUST BE MANUALLY CALLED TO DELE
 */
function deleteTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for(var i = 0; i < triggers.length; i++){
      ScriptApp.deleteTrigger(triggers[i]);
  }
}

/**
 * Gets the total balance from all data from the Google Form spreadsheet
 */
function totalBalance() {
  // Gets the raw data from the Google Form spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var rawDatasheet = spreadsheet.getSheetByName(RAW_DATASHEET_NAME);
  var data = rawDatasheet.getDataRange().getValues();
  
  // Calculates the total income and expenses
  var totalIncome = 0;
  var totalExpenses = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][RAW_DATASHEET_CATEGORY_COL] == 'Income') {
      totalIncome += data[i][RAW_DATASHEET_AMOUNT];
    } else {
      totalExpenses += data[i][RAW_DATASHEET_AMOUNT];
    }
  }
  
  // Sets the result on the dashboard
  var totalBalance = totalIncome - totalExpenses;
  var dashboard = spreadsheet.getSheetByName(DASHBOARD_DATASHEET_NAME);
  var cell = dashboard.getRange(DASHBOARD_TOTAL_BALANCE_CELL);
  cell.setNumberFormat(CURRENCY_FORMAT);
  cell.setValue(totalBalance);
  
  // Sets the color of the text
  if (totalBalance >= 0) {
    cell.setFontColor(POSITIVE_AMOUNT_COLOR);
  } else {
    cell.setFontColor(NEGATIVE_AMOUNT_COLOR);
  }
}

/**
 * Sets the sum of a query sum to the specified sheet and cell
 * @param {String} target The target of the query
 * @param {String} sql The query parameters to search for
 * @param {String} resultSheet The sheet to display the result in
 * @param {String} resultCell The cell on the sheet to display the result in
 * @param {String} resultFormat The format to display the result
 */
function totalQuerySum(target, sql, resultSheet, resultCell, resultFormat) {
  // Gets the sheet to display the result
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(resultSheet);
  
  // Queries the data and sums the query output
  var queryData = query(target, sql);
  var sum = 0;
  for (var i = 1; i < queryData.length; i++) {
    sum += queryData[i][0];
  }
  
  // Sets the query sum to resultCell using resultFormat
  var cell = sheet.getRange(resultCell);
  cell.setNumberFormat(resultFormat);
  cell.setValue(sum);
}

/**
 * Displays pie chart with data from datasheet with given key and value ranges
 * @param {String} keyRange The range of the keys
 * @param {String} valRange The range of the values
 * @param {Integer} rowPos The row of the top-left corner of the chart (Google Sheets uses 1-based indexing)
 * @param {Integer} colPos The column of the top-left corner of the chart (Google Sheets uses 1-based indexing)
 * @param {String} title The title of the plot
 */
function displayPieChart(keyRange, valRange, rowPos, colPos, title) {
  // Gets the necessary sheets
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = spreadsheet.getSheetByName(DATASTORE_DATASHEET_NAME); 
  var dashboard = spreadsheet.getSheetByName(DASHBOARD_DATASHEET_NAME);
  
  // Sets default column widths
  setDefaultWidths(DASHBOARD_DATASHEET_NAME, MONTH_GRAPH_COL_RANGE);

  // Gets the ranges for the labels and data values
  var totalChartLabels = dataSheet.getRange(keyRange);
  var totalChartValues = dataSheet.getRange(valRange);
  
  // Gets width and height of chart based on the Budget Overview table
  var width = dashboard.getColumnWidth(MONTH_GRAPH_COL_RANGE[0]) * (MONTH_GRAPH_COL_RANGE[1] - MONTH_GRAPH_COL_RANGE[0] + 1);
  var height = dashboard.getRowHeight(1) * TALBE_NUM_ROWS;

  // Sets parameters for the pie chart
  var totalsChart = dataSheet.newChart()
  .setChartType(Charts.ChartType.PIE)
  .addRange(totalChartLabels)
  .addRange(totalChartValues)
  .setPosition(rowPos, colPos, 0, 0)
  .setOption('legend.position', 'right')
  .setOption('pieSliceText', 'value-and-percentage')
  .setOption('title', title)
  .setOption('width', width)
  .setOption('height', height)
  .setNumHeaders(1)
  .build();

  // Inserts the chart into the dashboard
  dashboard.insertChart(totalsChart);
}

/**
 * Sets the pie chart data of aggregates of each expense type to the data sheet
 * @param {String} target The target of the query
 * @param {String} sql The query parameters to search for
 * @param {Integer} keyCol The column to store keys (Google Sheets uses 1-based indexing)
 * @param {Integer} valCol The column to store values (Google Sheets used 1-based indexing)
 */
function setPieChartData(target, sql, keyCol, valCol) {
  // Uses the data sheet for query process and data storage
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = spreadsheet.getSheetByName(DATASTORE_DATASHEET_NAME);
  
  // Deletes currently stored data
  dataSheet.deleteColumn(valCol);
  dataSheet.deleteColumn(keyCol);

  
  // Queries data and stores in it in columns dedicated to temporary storage
  var query = '=QUERY(' + target + ', \"' + sql + '\")';
  var pushQuery = dataSheet.getRange(DATASTORE_TEMP_CELL).setFormula(query);
  var range = dataSheet.getRange(DATASTORE_TEMP_RANGE).getValues();
  var numRows = range.filter(String).length;
  var pullResult = dataSheet.getRange(DATASTORE_TEMP_ROW, DATASTORE_TEMP_COL1, numRows, DATASTORE_TEMP_TWO_COLS).getValues();

  // Adds queried data to a map to aggregate data of each category
  var expenses = new Map();
  for (var i = 1; i < pullResult.length; i++) {
    var key = pullResult[i][0];
    var val = pullResult[i][1];
    if (expenses.has(key)) {
      expenses.set(key, expenses.get(key) + val);
    } else {
      expenses.set(key, val);
    }
  }
  
  // Sets aggregate data to the the specified keyCol and valCol, starting on the second row
  var row = 2;
  expenses.forEach((val, key, map) => {
    dataSheet.getRange(row, keyCol).setValue(key);
    dataSheet.getRange(row, valCol).setValue(val);
    row += 1;
  });
  
  // Deletes temporary data 
  dataSheet.deleteColumn(DATASTORE_TEMP_COL1);
  dataSheet.deleteColumn(DATASTORE_TEMP_COL2);
}

/**
 * Queries data from the spreadsheet with given parameters
 * @param {String} target The target of the query
 * @param {String} sql The query parameters to search for
 * @return {Array} the result
 */
function query(target, sql) { 
  // Uses the data sheet for query process
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = spreadsheet.getSheetByName(DATASTORE_DATASHEET_NAME);
  
  // Creates the query string
  var query = '=QUERY(' + target + ', \"' + sql + '\")';
  
  // Sets query results in the data sheet for temporary storage
  var pushQuery = dataSheet.getRange(DATASTORE_TEMP_CELL).setFormula(query);
  
  // Finds the number of rows in the data sheet for the temp column and gets the results
  var range = dataSheet.getRange(DATASTORE_TEMP_RANGE).getValues();
  var numRows = range.filter(String).length;
  var pullResult = dataSheet.getRange(DATASTORE_TEMP_ROW, DATASTORE_TEMP_COL1, numRows).getValues();
  
  // Deletes the column returns the queried data
  dataSheet.deleteColumn(DATASTORE_TEMP_COL1);
  return pullResult;
}

/**
 * Clears all figures from a sheet
 * @param {String} sheetName The name of the sheet to clear figures of
 */
function clearFigures(sheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  var charts = sheet.getCharts();
  for (var i = 0; i < charts.length; i++) {
    sheet.removeChart(charts[i]);
  }
}

/**
 * Sets a range of columns to default widths, determined by column 1 of the Data sheet
 * @param {String} sheetName The name of the sheet to set columns to default width
 * @param {Array} cols The array storing the start and end columns to set the widths, inclusive
 */
function setDefaultWidths(sheetName, cols) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = spreadsheet.getSheetByName(DATASTORE_DATASHEET_NAME);
  var sheet = spreadsheet.getSheetByName(sheetName);
  var width = dataSheet.getColumnWidth(1);
  
  var col = cols[0];
  while (col <= cols[1]) {
    sheet.setColumnWidth(col, width);
    col += 1;
  }
}

/**
 * Main function to update the dashboard
 */
function updateDashboard() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = spreadsheet.getSheetByName(DASHBOARD_DATASHEET_NAME);
  
  // Clears current figures to generate figures that reflect potential updates
  clearFigures(DASHBOARD_DATASHEET_NAME);
  
  // Queries data from Google Form responses and sets the data
  totalQuerySum(FORM_QUERY_TARGET, FORM_DATE_EXPENSE_SQL, DASHBOARD_DATASHEET_NAME, DASHBOARD_DATE_EXPENSE_CELL, CURRENCY_FORMAT);
  totalQuerySum(FORM_QUERY_TARGET, FORM_MONTH_EXPENSE_SQL, DASHBOARD_DATASHEET_NAME, DASHBOARD_MONTH_EXPENSE_CELL, CURRENCY_FORMAT);
  totalQuerySum(FORM_QUERY_TARGET, FORM_YEAR_EXPENSE_SQL, DASHBOARD_DATASHEET_NAME, DASHBOARD_YEAR_EXPENSE_CELL, CURRENCY_FORMAT);
  totalQuerySum(FORM_QUERY_TARGET, FORM_MONTH_INCOME_SQL, DASHBOARD_DATASHEET_NAME, DASHBOARD_MONTH_INCOME_CELL, CURRENCY_FORMAT);
  totalQuerySum(FORM_QUERY_TARGET, FORM_YEAR_INCOME_SQL, DASHBOARD_DATASHEET_NAME, DASHBOARD_YEAR_INCOME_CELL, CURRENCY_FORMAT);
  totalBalance();

  // Computes and sets the total monthly balance
  var monthBalanceCell = dashboard.getRange(DASHBOARD_MONTH_BALANCE_CELL);
  var monthIncomeValue = dashboard.getRange(DASHBOARD_MONTH_INCOME_CELL).getValue();
  var monthExpenseValue = dashboard.getRange(DASHBOARD_MONTH_EXPENSE_CELL).getValue();
  var monthBalanceValue = monthIncomeValue - monthExpenseValue;
  monthBalanceCell.setNumberFormat(CURRENCY_FORMAT);
  monthBalanceCell.setValue(monthBalanceValue);
  if (monthBalanceValue >= 0) {
    monthBalanceCell.setFontColor(POSITIVE_AMOUNT_COLOR);
  } else {
    monthBalanceCell.setFontColor(NEGATIVE_AMOUNT_COLOR);
  }
     
  // Computes and sets the total yearly balance
  var yearBalanceCell = dashboard.getRange(DASHBOARD_YEAR_BALANCE_CELL);
  var yearIncomeValue = dashboard.getRange(DASHBOARD_YEAR_INCOME_CELL).getValue();
  var yearExpenseValue = dashboard.getRange(DASHBOARD_YEAR_EXPENSE_CELL).getValue();
  var yearBalanceValue = yearIncomeValue - yearExpenseValue;
  yearBalanceCell.setNumberFormat(CURRENCY_FORMAT);
  yearBalanceCell.setValue(yearBalanceValue);
  if (yearBalanceValue >= 0) {
    yearBalanceCell.setFontColor(POSITIVE_AMOUNT_COLOR);
  } else {
    yearBalanceCell.setFontColor(NEGATIVE_AMOUNT_COLOR);
  }
  
  // Plots the monthly expenses as a pie chart
  setPieChartData(FORM_QUERY_TARGET, FORM_MONTH_CATEGORY_EXPENSE_SQL, DATASTORE_MONTH_KEY_COL, DATASTORE_MONTH_VAL_COL);
  displayPieChart(DATASTORE_MONTH_KEY_RANGE, DATASTORE_MONTH_VAL_RANGE, DATASTORE_MONTH_GRAPH_ROW_POS, DATASTORE_MONTH_GRAPH_COL_POS, MONTH_GRAPH_TITLE);
}