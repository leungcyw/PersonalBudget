/**
 * Contains bar chart visualization
 */


/**
 * Displays column chart with monthly expense and income data
 * @param {String} dataRange The cell range of data
 * @param {Integer} rowPos The row of the top-left corner of the chart (Google Sheets uses 1-based indexing)
 * @param {Integer} colPos The column of the top-left corner of the chart (Google Sheets uses 1-based indexing)
 */
function displayBarChart(dataRange, rowPos, colPos, title) {
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
  .setOption('title', title)
  .setOption('width', width)
  .setOption('height', height);

  var chart = chartBuilder.build();
  dashboard.insertChart(chart);
}