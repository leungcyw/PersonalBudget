/**
 * Contains table visualization
 */


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
  titleRange.setFontSize(TABLE_FONT_SIZE);
  titleRange.setBackground(TABLE_TITLE_COLOR);
  
  // Formats cells for table body
  var bodyRange = dashboard.getRange(TABLE_KEY_VAL_RANGE);
  bodyRange.setHorizontalAlignment('right');
  bodyRange.setVerticalAlignment('middle');
  bodyRange.setFontSize(TABLE_FONT_SIZE);
  
  // Formats alternating colors for table body
  var expenseRange = dashboard.getRange(TABLE_BODY_EXPENSE_COLOR_RANGE);
  expenseRange.setBackground(TABLE_BODY_EXPENSE_COLOR);
  var incomeRange = dashboard.getRange(TABLE_BODY_INCOME_COLOR_RANGE);
  incomeRange.setBackground(TABLE_BODY_INCOME_COLOR);
  var balanceRange = dashboard.getRange(TABLE_BODY_BALANCE_COLOR_RANGE);
  balanceRange.setBackground(TABLE_BODY_BALANCE_COLOR);
  
  // Sets the table border
  var tableRange = dashboard.getRange(TABLE_RANGE);
  tableRange.setBorder(true, true, true, true, true, true);
}