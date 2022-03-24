/**
 * Misc helper functions for updating dashboard
 */


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
 * NOTE: MUST BE MANUALLY CALLED TO DELETE
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
  // Tries to get cached values exist
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = spreadsheet.getSheetByName(DATASTORE_DATASHEET_NAME);
  var rawDatasheet = spreadsheet.getSheetByName(RAW_DATASHEET_NAME);
  var storedTimestamp = datasheet.getRange(DATASTORE_LAST_COMPUTED_TIMESTAMP_CELL).getValue();
  var storedValue = datasheet.getRange(DATASTORE_LAST_COMPUTED_TOTAL_BALANCE).getValue();
  
  // Defines computation variables
  var totalIncome = 0;
  var totalExpenses = 0;
  
  // Accumulates the cached values if it exists
  if (typeof storedTimestamp == 'object' && typeof storedValue == 'number') {
    var sql = 'select D where A >= datetime \'"&TEXT($A2, "yyyy-mm-dd HH:mm:ss")&"\' and ';
    var queryIncomeData = query(FORM_QUERY_TARGET, sql + INCOME_QUERY_CONTENT);
    for (var i = 1; i < queryIncomeData.length; i++) {
      totalIncome += queryIncomeData[i];
    }
    var queryExpensesData = query(FORM_QUERY_TARGET, sql + EXPENSES_QUERY_CONTENT);
    for (var i = 1; i < queryExpensesData.length; i++) {
      totalExpenses += queryExpensesData[i]; 
    }
    
  // Cached values are invalid, so we iterate through all data entries
  } else {
    storedValue = 0;
    
    // Gets the raw data from the Google Form spreadsheet
    var data = rawDatasheet.getDataRange().getValues();
  
    // Calculates the total income and expenses
    for (var i = 1; i < data.length; i++) {
      if (data[i][RAW_DATASHEET_CATEGORY_COL] == 'Income') {
        totalIncome += data[i][RAW_DATASHEET_AMOUNT];
      } else {
        totalExpenses += data[i][RAW_DATASHEET_AMOUNT];
      }
    }        
  }
  
  // Sets the result on the dashboard
  var totalBalance = totalIncome - totalExpenses + storedValue;
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
 
  // Tries to cache the timestamp of the last entry and current total
  // Errors when totalBalance is called when there are no data entries (e.g. when called on init)
  try {
    var lastComputedTimestamp = rawDatasheet.getRange(rawDatasheet.getLastRow(), RAW_DATASHEET_TIMESTAMP + 1).getValue();
    datasheet.getRange(DATASTORE_LAST_COMPUTED_TIMESTAMP_CELL).setValue(new Date(lastComputedTimestamp.getTime() + 1000));
    datasheet.getRange(DATASTORE_LAST_COMPUTED_TOTAL_BALANCE).setValue(totalBalance);
  } catch (err) {}
}

/**
 * Sets the sum of a query sum to the specified sheet and cell
 * @param {String} target The target of the query
 * @param {String} sql The query parameters to search for
 * @param {String} resultSheet The sheet to display the result in
 * @param {String} resultCell The cell on the sheet to display the result in
 * @param {String} resultFormat The format to display the result
 * @param {String} cacheCell The cell on the data sheet to cache the current result
 * @param {Number} cachedValue The value of previously cached data
 */
function totalQuerySum(target, sql, resultSheet, resultCell, resultFormat, cacheCell, cachedValue) {
  // Gets the sheet to display the result
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(resultSheet);
  
  // Queries the data and sums the query output
  var queryData = query(target, sql);
  var sum = cachedValue;
  for (var i = 1; i < queryData.length; i++) {
    sum += queryData[i][0];
  }
  
  // Sets the query sum to resultCell using resultFormat
  var cell = sheet.getRange(resultCell);
  cell.setNumberFormat(resultFormat);
  cell.setValue(sum);
  
  // Sets the query sum to the cacheCell
  if (cacheCell != null) {
    var datasheet = spreadsheet.getSheetByName(DATASTORE_DATASHEET_NAME);
    datasheet.getRange(cacheCell).setValue(sum);
  }
}

/**
 * Checks if a cached value could be used to reduce query amount before computing the sum
 * @param {String} target The target of the query
 * @param {String} sql The query parameters to search for
 * @param {String} duration The time duration the query is for
 * @param {String} resultSheet The sheet to display the result in
 * @param {String} resultCell The cell on the sheet to display the result in
 * @param {String} resultFormat The format to display the result
 * @param {String} cacheCell The cell on the data sheet to cache the current result
 */
function cachedTotalQuerySum(target, sql, duration, resultSheet, resultCell, resultFormat, cacheCell) {
  // Gets the cached values in the data sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = spreadsheet.getSheetByName(DATASTORE_DATASHEET_NAME);
  var storedTimestamp = datasheet.getRange(DATASTORE_LAST_COMPUTED_TIMESTAMP_CELL).getValue();
  var storedValue = (cacheCell == null) ? null : datasheet.getRange(cacheCell).getValue();
    
  // Checks if the cached value is valid for the given duration
  var validCache = false;
  if (typeof storedTimestamp == 'object' && typeof storedValue == 'number') {
    var today = new Date();
    switch (duration) {
      case 1:
        validCache = false;
        break;
      case 2:
        validCache = (storedTimestamp.getMonth() == today.getMonth() && storedTimestamp.getFullYear() == today.getFullYear());
        break;
      case 3:
        validCache = (storedTimestamp.getFullYear() == today.getFullYear());
        break;
      default:
        validCache = true;
    }
  }  
  
  // Runs an updated query sum on valid caches
  if (validCache) {
    var isIncomeQuery = sql.includes(INCOME_QUERY_CONTENT);
    var updated_sql = 'select D where A >= datetime \'"&TEXT($A2, "yyyy-mm-dd HH:mm:ss")&"\' and ';
    updated_sql += (isIncomeQuery) ? INCOME_QUERY_CONTENT : EXPENSES_QUERY_CONTENT;
    totalQuerySum(target, updated_sql, resultSheet, resultCell, resultFormat, cacheCell, storedValue);
    return;
  }
  
  // Uses original query sum on invalid caches
  totalQuerySum(target, sql, resultSheet, resultCell, resultFormat, cacheCell, 0);
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
  dataSheet.getRange(1, DATASTORE_TEMP_COL1, dataSheet.getLastRow()).clear();
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