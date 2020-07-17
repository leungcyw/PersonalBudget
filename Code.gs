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
const DATASTORE_TEMP_CELL = 'G1';
const DATASTORE_TEMP_RANGE = 'G1:G';
const DATASTORE_TEMP_ROW = 1;
const DATASTORE_TEMP_COL1 = 7;
const DATASTORE_TEMP_COL2 = 8;
const DATASTORE_TEMP_TWO_COLS = 2;
const DATASTORE_MONTH_KEY_RANGE = 'E1:E';
const DATASTORE_MONTH_VAL_RANGE = 'F1:F';
const DATASTORE_MONTH_KEY_COL = 5;
const DATASTORE_MONTH_VAL_COL = 6;
const DATASTORE_BAR_INIT_CELLS = ['B2', 'C2', 'D2'];
const DATASTORE_BAR_INCOME_COL = 3;
const DATASTORE_BAR_EXPENSE_COL = 4;
const DATASTORE_BAR_MONTHS_RANGE = 'B1:B';
const DATASTORE_BAR_MONTHS_DISPLAY_RANGE = 'B2:B';
const DATASTORE_BAR_EXCESS_MONTHS = 'B13:B';
const DATASTORE_BAR_TOTAL_RANGE = 'B:D';
const DATASTORE_BAR_MONTHS_ARR = ['Jan.', 'Feb.', 'Mar.', 'Apr.', 'May', 'June', 'July', 'Aug.', 'Sept.', 'Oct.', 'Nov.', 'Dec.'];

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
const TABLE_ROWS = [1, 9];
const TABLE_COLS = [1, 2];
const TABLE_CELL_HEIGHT = 35;
const TABLE_CELL_KEY_WIDTH = 250;
const TABLE_CELL_VAL_WIDTH = 180;
const TABLE_STRINGS = new Map([
  ['A1', 'Budget Overview'],
  ['B1', ''],
  ['A2', 'Current Date Expenses:'],
  ['A3', 'Current Month Expenses:'],
  ['A4', 'Current Year Expenses:'],
  ['A5', 'Current Month Income:'],
  ['A6', 'Current Year Income:'],
  ['A7', 'Current Month Balance:'],
  ['A8', 'Current Year Balance:'],
  ['A9', 'Total Balance:']
]);
const TABLE_TITLE_RANGE = 'A1:B1';
const TABLE_TITLE_FONT_SIZE = 24;
const TABLE_TITLE_COLOR = '#a4c2f4';
const TABLE_KEY_VAL_RANGE = 'A2:B9';
const TABLE_BODY_FONT_SIZE = 14;
const TABLE_BODY_ALTER_COLOR_RANGES = ['A2:B4', 'A7:B8'];
const TABLE_BODY_ALTER_COLOR = '#d9d9d9';
const TABLE_BODY_WHITE_COLOR_RANGES = ['A5:B6', 'A9:B9'];
const TABLE_BODY_WHITE_COLOR = '#ffffff';
const TABLE_RANGE = 'A1:B9';

// Graph constants
const DATASTORE_MONTH_GRAPH_ROW_POS = 1;
const DATASTORE_MONTH_GRAPH_COL_POS = 4;
const MONTH_GRAPH_TITLE = 'Current Month Expenses';
const MONTH_GRAPH_COL_RANGE = [4, 8];
const BAR_LABEL_INCOME = 'Income';
const BAR_LABEL_EXPENSE = 'Expense';
const BAR_GRAPH_ROW_RANGE = [11, 29];
const BAR_GRAPH_EXTRA_COL_RANGE = [3, 8];
const DATASTORE_BAR_GRAPH_ROW_POS = 11;
const DATASTORE_BAR_GRAPH_COL_POS = 1;
const NUM_MONTHS = 12;




function init() {
  // Call functions to initialize various components
  formatOverviewTable();
  deleteTriggers();
  createUpdateTriggers();
  
  // Sets initial cell for monthly data info
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = spreadsheet.getSheetByName(DATASTORE_DATASHEET_NAME);
  dataSheet.getRange(DATASTORE_BAR_INIT_CELLS[0]).setValue(DATASTORE_BAR_MONTHS_ARR[new Date().getMonth()]);
  dataSheet.getRange(DATASTORE_BAR_INIT_CELLS[1]).setValue(0);
  dataSheet.getRange(DATASTORE_BAR_INIT_CELLS[2]).setValue(0);
  
  // Updates the dashboard
  updateDashboard();
}

/**
 * Creates the 'Budget Overview' table on the Dashboard 
 */
function formatOverviewTable() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = spreadsheet.getSheetByName("Dashboard");
  
  //Sets the dimensions of the Budget Overview table
  dashboard.setRowHeights(TABLE_ROWS[0], TABLE_ROWS[1] - TABLE_ROWS[0] + 1, TABLE_CELL_HEIGHT);
  dashboard.setColumnWidth(TABLE_COLS[0], TABLE_CELL_KEY_WIDTH);
  dashboard.setColumnWidth(TABLE_COLS[1], TABLE_CELL_VAL_WIDTH);
  
  // Sets table text
  for (const [k, v] of TABLE_STRINGS) {
    var cell = dashboard.getRange(k);
    cell.setValue(v);
  }
  
  // Merges cells of table heading
  dashboard.getRange(TABLE_TITLE_RANGE).mergeAcross();
  
  // Formats cells for table heading
  var titleRange = dashboard.getRange(TABLE_TITLE_RANGE);
  titleRange.setHorizontalAlignment('center')
  titleRange.setVerticalAlignment('middle');
  titleRange.setFontWeight('bold');
  titleRange.setFontSize(TABLE_TITLE_FONT_SIZE);
  titleRange.setBackground(TABLE_TITLE_COLOR);
  
  // Formats cells for table body
  var bodyRange = dashboard.getRange(TABLE_KEY_VAL_RANGE);
  bodyRange.setHorizontalAlignment('right');
  bodyRange.setVerticalAlignment('middle');
  bodyRange.setFontSize(TABLE_BODY_FONT_SIZE);
  
  // Formats alternating colors for table body
  var alterColorRanges = dashboard.getRangeList(TABLE_BODY_ALTER_COLOR_RANGES);
  alterColorRanges.setBackground(TABLE_BODY_ALTER_COLOR);
  var whiteColorRanges = dashboard.getRangeList(TABLE_BODY_WHITE_COLOR_RANGES);
  whiteColorRanges.setBackground(TABLE_BODY_WHITE_COLOR);
  
  // Sets the table border
  var tableRange = dashboard.getRange(TABLE_RANGE);
  tableRange.setBorder(true, true, true, true, true, true);
}

/**
 * Creates Triggers to update the dashboard whenever the Google Form is submitted and each new day
 * NOTE: MUST BE MANUALLY CALLED TO CREATE THE TRIGGER 
 */
function createUpdateTriggers() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Checks if the Triggers currently exist
  var createFormTrigger = true;
  var createDateTrigger = true;
  var createMonthTrigger = true;
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    if (trigger.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT && trigger.getHandlerFunction() === 'updateDashboard') {
      createFormTrigger = false;
    } else if (trigger.getEventType() === ScriptApp.EventType.CLOCK && trigger.getHandlerFunction() === 'updateDashboard') {
      createDateTrigger = false;
    } else if (trigger.getEventType() === ScriptApp.EventType.CLOCK && trigger.getHandlerFunction() === 'updateMonthData') {
      createMonthTrigger = false;
    }
  });

  // Creates a Trigger for Form submits
  if (createFormTrigger) {
    ScriptApp.newTrigger('updateDashboard')
    .forSpreadsheet(spreadsheet)
    .onFormSubmit()
    .create();
  }

  // Creates a Trigger for date changes
  if (createDateTrigger) {
    ScriptApp.newTrigger('updateDashboard')
    .timeBased()
    .atHour(0)
    .nearMinute(30)
    .everyDays(1)
    .create();
  }

  if (createMonthTrigger) {
    ScriptApp.newTrigger('updateMonthData')
    .timeBased()
    .atHour(0)
    .nearMinute(30)
    .everyDays(1)
    .create();
  }
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
 * Updates the month data at the beginning of each month for the column chart
 */
function updateMonthData() {
  // Only runs on the first of the month
  var date = new Date();
  if (date.getDate() != 1) {
    return;
  }
  
  // Gets the necessary data sheet to edit the sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = spreadsheet.getSheetByName(DATASTORE_DATASHEET_NAME);
  var currentMonth = DATASTORE_BAR_MONTHS_ARR[new Date().getMonth()];
  var rangeVals = dataSheet.getRange(DATASTORE_BAR_MONTHS_DISPLAY_RANGE).getValues();
  var numRows = rangeVals.filter(String).length;
  
  // Reduces the table if there are already 12 months displayed, and then sets the data
  if (numRows >= NUM_MONTHS) {
    for (var i = 3; i <= numRows + 1; i++) {
      dataSheet.getRange(i-1, 2, 1, 3).setValues(dataSheet.getRange(i, 2, 1, 3).getValues());
      Logger.log(dataSheet.getRange(i, 2, 1, 3).getValues());
    }
    dataSheet.getRange(DATASTORE_BAR_EXCESS_MONTHS).clear();
    dataSheet.getRange(13, 2).setValue(currentMonth);
    dataSheet.getRange(13, 3).setValue(0);
    dataSheet.getRange(13, 4).setValue(0);
  } else {
    dataSheet.getRange(numRows + 2, 2).setValue(currentMonth);
    dataSheet.getRange(numRows + 2, 3).setValue(0);
    dataSheet.getRange(numRows + 2, 4).setValue(0);
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
 * Displays column chart with monthly expense and income data
 * @param {String} dataRange The cell range of data
 * @param {Integer} rowPos The row of the top-left corner of the chart (Google Sheets uses 1-based indexing)
 * @param {Integer} colPos The column of the top-left corner of the chart (Google Sheets uses 1-based indexing)
 */
function displayBarChart(dataRange, rowPos, colPos) {
  // Gets the data sheet to pull data
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(DATASTORE_DATASHEET_NAME);
  var dashboard = spreadsheet.getSheetByName(DASHBOARD_DATASHEET_NAME);
  var data = sheet.getRange(dataRange);
  
  // Calculates the dimensions of the graph
  var width = TABLE_CELL_KEY_WIDTH + TABLE_CELL_VAL_WIDTH + dashboard.getColumnWidth(BAR_GRAPH_EXTRA_COL_RANGE[0]) * (BAR_GRAPH_EXTRA_COL_RANGE[1] - BAR_GRAPH_EXTRA_COL_RANGE[0] + 1);
  var height = dashboard.getRowHeight(BAR_GRAPH_ROW_RANGE[0]) * (BAR_GRAPH_ROW_RANGE[1] - BAR_GRAPH_ROW_RANGE[0] + 1);

  // Sets the parameters of the chart
  var chartBuilder = sheet.newChart()
  .setChartType(Charts.ChartType.COLUMN)
  .addRange(data)
  .setOption('colors', [POSITIVE_AMOUNT_COLOR, NEGATIVE_AMOUNT_COLOR])
  .setPosition(rowPos, colPos, 1, 1)
  .setOption('legend.position', 'right')
  .setOption('series', {0:{labelInLegend:BAR_LABEL_INCOME}, 1:{labelInLegend:BAR_LABEL_EXPENSE}})
  .setOption('width', width)
  .setOption('height', height);

  var chart = chartBuilder.build();
  dashboard.insertChart(chart);  
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
  
  // Updates data for bar graph
  var dataSheet = spreadsheet.getSheetByName(DATASTORE_DATASHEET_NAME);
  var currentMonth = DATASTORE_BAR_MONTHS_ARR[new Date().getMonth()];
  var months = dataSheet.getRange(DATASTORE_BAR_MONTHS_RANGE).getValues();
  var row = -1;
  for (var i = 0; i < months.length; i++) {
    if (months[i][0] == currentMonth) {
      row = i;
      break;
    }
  }
  if (row != -1) {
    dataSheet.getRange(row + 1, DATASTORE_BAR_INCOME_COL).setValue(dashboard.getRange(DASHBOARD_MONTH_INCOME_CELL).getValue());
    dataSheet.getRange(row + 1, DATASTORE_BAR_EXPENSE_COL).setValue(dashboard.getRange(DASHBOARD_MONTH_EXPENSE_CELL).getValue());
  }
  
  // Plots the income and expenses column chart
  displayBarChart(DATASTORE_BAR_TOTAL_RANGE, DATASTORE_BAR_GRAPH_ROW_POS, DATASTORE_BAR_GRAPH_COL_POS);
}