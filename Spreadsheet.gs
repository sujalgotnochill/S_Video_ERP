/*************************************************
 * S Video ERP
 * Module : Core
 * File   : Spreadsheet.gs
 * Version: 1.0.0
 *
 * Responsibility:
 * Single gateway for Google Sheets operations.
 *
 * Rules:
 * - No business logic.
 * - No module-specific logic.
 * - No object mapping.
 * - No validation.
 *************************************************/

/*************************************************
 * Returns active spreadsheet.
 *************************************************/
function getSpreadsheet() {

  return SpreadsheetApp.getActiveSpreadsheet();

}

/*************************************************
 * Returns sheet by configured name.
 *************************************************/
function getSheet(sheetName) {

  const sheet =
    getSpreadsheet().getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(
      "Sheet not found: " + sheetName
    );
  }

  return sheet;

}

/*************************************************
 * Returns entire data range values.
 *************************************************/
function getSheetData(sheetName) {

  return getSheet(sheetName)
    .getDataRange()
    .getValues();

}

/*************************************************
 * Returns last used row.
 *************************************************/
function getSheetLastRow(sheetName) {

  return getSheet(sheetName)
    .getLastRow();

}

/*************************************************
 * Returns last used column.
 *************************************************/
function getSheetLastColumn(sheetName) {

  return getSheet(sheetName)
    .getLastColumn();

}

/*************************************************
 * Returns range object.
 *************************************************/
function getSheetRange(
  sheetName,
  row,
  column,
  numRows,
  numColumns
) {

  return getSheet(sheetName)
    .getRange(
      row,
      column,
      numRows,
      numColumns
    );

}

/*************************************************
 * Reads values from a range.
 *************************************************/
function getRangeValues(
  sheetName,
  row,
  column,
  numRows,
  numColumns
) {

  return getSheetRange(
    sheetName,
    row,
    column,
    numRows,
    numColumns
  ).getValues();

}

/*************************************************
 * Writes values to a range.
 *************************************************/
function setRangeValues(
  sheetName,
  row,
  column,
  values
) {

  getSheetRange(
    sheetName,
    row,
    column,
    values.length,
    values[0].length
  ).setValues(values);

}

/*************************************************
 * Appends a single row.
 *************************************************/
function appendSheetRow(
  sheetName,
  values
) {

  getSheet(sheetName)
    .appendRow(values);

}

/*************************************************
 * Clears a range.
 *************************************************/
function clearSheetRange(
  sheetName,
  row,
  column,
  numRows,
  numColumns
) {

  getSheetRange(
    sheetName,
    row,
    column,
    numRows,
    numColumns
  ).clearContent();

}