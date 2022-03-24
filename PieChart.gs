/**
 * Contains pie chart visualization
 */


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
  if (title === YEAR_GRAPH_TITLE) {
    width = TABLE_CELL_KEY_WIDTH + TABLE_CELL_VAL_WIDTH + dashboard.getColumnWidth(BAR_GRAPH_EXTRA_COL_RANGE[0]) * (BAR_GRAPH_EXTRA_COL_RANGE[1] - BAR_GRAPH_EXTRA_COL_RANGE[0] + 1);
    height = dashboard.getRowHeight(BAR_GRAPH_ROW_RANGE[0]) * (BAR_GRAPH_ROW_RANGE[1] - BAR_GRAPH_ROW_RANGE[0] + 1);
  }

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
  
  // Clears currently stored data
  dataSheet.getRange(1, valCol, dataSheet.getLastRow()).clear();
  dataSheet.getRange(1, keyCol, dataSheet.getLastRow()).clear();
  
  // Queries data and stores in it in columns dedicated to temporary storage
  var query = '=QUERY(' + target + ', \"' + sql + '\")';
  var pushQuery = dataSheet.getRange(DATASTORE_TEMP_CELL).setFormula(query);
  var range = dataSheet.getRange(DATASTORE_TEMP_RANGE).getValues();
  var numRows = range.filter(String).length;
  var pullResult = dataSheet.getRange(DATASTORE_TEMP_ROW, DATASTORE_TEMP_COL1, numRows, DATASTORE_TEMP_TWO_COLS).getValues();

  // Adds queried data to a map to aggregate data of each category
  var expenses = {};
  for (var i = 1; i < pullResult.length; i++) {
    var key = pullResult[i][0];
    var val = pullResult[i][1];
    if (expenses.hasOwnProperty(key)) {
      expenses[key] = expenses[key] + val;
    } else {
      expenses[key] = val;
    }
  }
  
  // Sets aggregate data to the the specified keyCol and valCol, starting on the second row
  var row = 2;  
  for (const [key, val] of Object.entries(expenses)) {
    dataSheet.getRange(row, keyCol).setValue(key);
    dataSheet.getRange(row, valCol).setValue(val);
    row += 1;    
  }
  
  // Clears temporary data 
  dataSheet.getRange(1, DATASTORE_TEMP_COL1, dataSheet.getLastRow()).clear();
  dataSheet.getRange(1, DATASTORE_TEMP_COL2, dataSheet.getLastRow()).clear();
}