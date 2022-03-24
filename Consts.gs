/**
 * Constants to manage dashboard appearance
 */


// Constants for Google Form data spreadsheet
const RAW_DATASHEET_NAME = 'Form Responses';
const RAW_DATASHEET_TIMESHEET_COL = 0;
const RAW_DATASHEET_CATEGORY_COL = 1;
const RAW_DATASHEET_DESCRIPTION_COL = 2;
const RAW_DATASHEET_AMOUNT = 3;
const RAW_DATASHEET_TIMESTAMP = 0;

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
const DATASTORE_LAST_COMPUTED_TIMESTAMP_CELL = 'A2';
const DATASTORE_LAST_COMPUTED_TOTAL_BALANCE = 'A3';
const DATASTORE_LAST_COMPUTED_MONTH_INCOME = 'A4';
const DATASTORE_LAST_COMPUTED_MONTH_EXPENSES = 'A5';
const DATASTORE_LAST_COMPUTED_YEAR_INCOME = 'A6';
const DATASTORE_LAST_COMPUTED_YEAR_EXPENSES = 'A7';

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
const EXPENSES_QUERY_CONTENT = 'B != \'Income\'';
const INCOME_QUERY_CONTENT = 'B = \'Income\'';
const DurationEnum = Object.freeze({"day": 1, "month": 2, "year": 3, "beginning_of_time": 0});

// Budget Overview table constants
const TALBE_NUM_ROWS = 9;
const TABLE_ROWS = [1, 9];
const TABLE_COLS = [1, 2];
const TABLE_CELL_HEIGHT = 28;
const TABLE_CELL_KEY_WIDTH = 200;
const TABLE_CELL_VAL_WIDTH = 110;
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
const TABLE_FONT_SIZE = 11;
const TABLE_TITLE_COLOR = '#ffffff';
const TABLE_KEY_VAL_RANGE = 'A2:B9';
const TABLE_BODY_EXPENSE_COLOR_RANGE = 'A2:B4';
const TABLE_BODY_EXPENSE_COLOR = '#ffd1d1';
const TABLE_BODY_INCOME_COLOR_RANGE = 'A5:B6';
const TABLE_BODY_INCOME_COLOR = '#e2f5d7';
const TABLE_BODY_BALANCE_COLOR_RANGE = 'A7:B9';
const TABLE_BODY_BALANCE_COLOR = '#d9d9d9';
const TABLE_RANGE = 'A1:B9';

// Graph constants
const DATASTORE_MONTH_GRAPH_ROW_POS = 1;
const DATASTORE_MONTH_GRAPH_COL_POS = 4;
const MONTH_GRAPH_TITLE = 'Current Month Expenses';
const MONTH_GRAPH_COL_RANGE = [4, 8];
const BAR_GRAPH_TITLE = 'Current Year Expenses';
const BAR_LABEL_INCOME = 'Income';
const BAR_LABEL_EXPENSE = 'Expense';
const BAR_GRAPH_ROW_RANGE = [11, 29];
const BAR_GRAPH_EXTRA_COL_RANGE = [3, 8];
const DATASTORE_BAR_GRAPH_ROW_POS = 11;
const DATASTORE_BAR_GRAPH_COL_POS = 1;
const NUM_MONTHS = 12;