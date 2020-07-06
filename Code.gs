/**
 * Code to manage my budget input and visual display
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

// Constants for dashboard spreadsheet
const DASHBOARD_DATASHEET_NAME = 'Dashboard';
const DASHBOARD_DATE_EXPENSE_CELL = 'B2';



/**
 * Displays data when the Google Sheets spreadsheet is opened
 * @param {Event} e The onOpen event
 */
function onOpen(e) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = spreadsheet.getSheetByName(DASHBOARD_DATASHEET_NAME);
  
  CURRENT_DATE_EXPENSES();
}


/**
 * Gets the total balance from all data from the Google Form spreadsheet
 * @return The total balance that is the total income minus the total expenses
 */
function TOTAL_BALANCE() {
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
  return totalIncome - totalExpenses;
}


/**
 * Gets the total expenses of the current date and displays it on the dashboard
 */
function CURRENT_DATE_EXPENSES() {
  // Gets the results of the query of all expenses from the current date
  const QUERY_TARGET = '\'Form Responses\'!A:D';
  const QUERY_SQL = 'select D where A >= date \'"&TEXT(TODAY(), "yyyy-mm-dd")&"\' and B != \'Income\'';
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = spreadsheet.getSheetByName(DASHBOARD_DATASHEET_NAME);
  var queryData = QUERY_DATA(QUERY_TARGET, QUERY_SQL);
 
  // Computes the total expenses of the current date
  var currentDateExpenses = 0;
  for (var i = 1; i < queryData.length; i++) {
    currentDateExpenses += queryData[i];
  }
  
  // Sets the result on the dashboard
  var cell = dashboard.getRange(DASHBOARD_DATE_EXPENSE_CELL);
  cell.setNumberFormat("$#,##0.00;$(#,##0.00)");
  cell.setValue(currentDateExpenses);
}


/**
 * Queries data from the spreadsheet with given parameters
 * @param {String} target The target of the query
 * @param {String} sql The query parameters to search for
 * @return {Array} the result
 */
function QUERY_DATA(target, sql) { 
  // Creates a temp sheet for query process
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var tempSheet = spreadsheet.insertSheet();
  
  // Creates the query string
  var query = '=QUERY(' + target + ', \"' + sql + '\")';
  
  // Sets query results in the temp sheet for temporary storage and gets the values as an array
  var pushQuery = tempSheet.getRange(1, 1).setFormula(query);
  var pullResult = tempSheet.getDataRange().getValues();
  
  // Deletes the temp sheet and returns the queried data
  spreadsheet.deleteSheet(tempSheet);
  return pullResult;
}