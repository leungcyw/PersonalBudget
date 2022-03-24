/**
 * Google App Script code to manage my budget input and visual display
 * Functions interact with Google Forms for data input
 * Functions interact with Google Sheets for data storage and display
 */

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
 * Main function to update the dashboard
 */
function updateDashboard() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = spreadsheet.getSheetByName(DASHBOARD_DATASHEET_NAME);
  
  // Clears current figures to generate figures that reflect potential updates
  clearFigures(DASHBOARD_DATASHEET_NAME);
  
  // Queries data from Google Form responses and sets the data
  // First tries to build from cached data, but recomputes all if cached data is not found
  cachedTotalQuerySum(FORM_QUERY_TARGET, FORM_DATE_EXPENSE_SQL, DurationEnum.day, DASHBOARD_DATASHEET_NAME, DASHBOARD_DATE_EXPENSE_CELL, CURRENCY_FORMAT, null);
  cachedTotalQuerySum(FORM_QUERY_TARGET, FORM_MONTH_EXPENSE_SQL, DurationEnum.month, DASHBOARD_DATASHEET_NAME, DASHBOARD_MONTH_EXPENSE_CELL, CURRENCY_FORMAT, DATASTORE_LAST_COMPUTED_MONTH_EXPENSES);
  cachedTotalQuerySum(FORM_QUERY_TARGET, FORM_YEAR_EXPENSE_SQL, DurationEnum.year, DASHBOARD_DATASHEET_NAME, DASHBOARD_YEAR_EXPENSE_CELL, CURRENCY_FORMAT, DATASTORE_LAST_COMPUTED_YEAR_EXPENSES);
  cachedTotalQuerySum(FORM_QUERY_TARGET, FORM_MONTH_INCOME_SQL, DurationEnum.month, DASHBOARD_DATASHEET_NAME, DASHBOARD_MONTH_INCOME_CELL, CURRENCY_FORMAT, DATASTORE_LAST_COMPUTED_MONTH_INCOME);
  cachedTotalQuerySum(FORM_QUERY_TARGET, FORM_YEAR_INCOME_SQL, DurationEnum.year, DASHBOARD_DATASHEET_NAME, DASHBOARD_YEAR_INCOME_CELL, CURRENCY_FORMAT, DATASTORE_LAST_COMPUTED_YEAR_INCOME);
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
  displayBarChart(DATASTORE_BAR_TOTAL_RANGE, DATASTORE_BAR_GRAPH_ROW_POS, DATASTORE_BAR_GRAPH_COL_POS, BAR_GRAPH_TITLE);
  
  
  // Plots the monthly expenses as a pie chart
  setPieChartData(FORM_QUERY_TARGET, FORM_MONTH_CATEGORY_EXPENSE_SQL, DATASTORE_MONTH_KEY_COL, DATASTORE_MONTH_VAL_COL);
  displayPieChart(DATASTORE_MONTH_KEY_RANGE, DATASTORE_MONTH_VAL_RANGE, DATASTORE_MONTH_GRAPH_ROW_POS, DATASTORE_MONTH_GRAPH_COL_POS, MONTH_GRAPH_TITLE);

  // Plots the yearly expenses as a pie chart
  setPieChartData(FORM_QUERY_TARGET, FORM_YEAR_CATEGORY_EXPENSE_SQL, DATASTORE_YEAR_KEY_COL, DATASTORE_YEAR_VAL_COL);
  displayPieChart(DATASTORE_YEAR_KEY_RANGE, DATASTORE_YEAR_VAL_RANGE, DATASTORE_YEAR_GRAPH_ROW_POS, DATASTORE_YEAR_GRAPH_COL_POS, YEAR_GRAPH_TITLE);
}