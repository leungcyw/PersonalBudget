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
const DATASTORE_TEMP_CELL = 'B1';
const DATASTORE_TEMP_RANGE = 'B1:B';
const DATASTORE_TEMP_ROW = 1;
const DATASTORE_TEMP_COL = 2;

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


/**
 * Displays data when the Google Sheets spreadsheet is opened
 * @param {Event} e The onOpen event
 */
function onOpen(e) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = spreadsheet.getSheetByName(DASHBOARD_DATASHEET_NAME);
  
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
    sum += queryData[i];
  }
  
  // Sets the query sum to resultCell using resultFormat
  var cell = sheet.getRange(resultCell);
  cell.setNumberFormat(resultFormat);
  cell.setValue(sum);
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
  var pullResult = dataSheet.getRange(DATASTORE_TEMP_ROW, DATASTORE_TEMP_COL, numRows).getValues();
  
  // Deletes the column returns the queried data
  dataSheet.deleteColumn(DATASTORE_TEMP_COL);
  return pullResult;
}