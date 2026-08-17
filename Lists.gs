/**
 * Returns all non-empty values under a header
 * from the Lists sheet.
 *
 * Example:
 * getList("Effects")
 */
function getList(headerName) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lists");

  const data = sheet.getDataRange().getValues();

  const headers = data[0];

  const column = headers.indexOf(headerName);

  if (column === -1) {
    throw new Error("Header not found: " + headerName);
  }

  const values = [];

  for (let i = 1; i < data.length; i++) {

    const value = data[i][column];

    if (value !== "") {
      values.push(value);
    }

  }

  return values;

}

/**
 * Applies a dropdown from Lists sheet
 * to an entire column in another sheet.
 */
function applyDropdown(targetSheetName, targetHeader, listHeader) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const targetSheet = ss.getSheetByName(targetSheetName);
  const listSheet = ss.getSheetByName("Lists");

  // Target Header
  const targetHeaders = targetSheet
    .getRange(1, 1, 1, targetSheet.getLastColumn())
    .getValues()[0];

  const targetColumn = targetHeaders.indexOf(targetHeader) + 1;

  if (targetColumn === 0) {
    throw new Error("Target header not found: " + targetHeader);
  }

  // Lists Header
  const listHeaders = listSheet
    .getRange(1, 1, 1, listSheet.getLastColumn())
    .getValues()[0];

  const listColumn = listHeaders.indexOf(listHeader) + 1;

  if (listColumn === 0) {
    throw new Error("List header not found: " + listHeader);
  }

  // Source values
  const lastRow = listSheet.getLastRow();

  const numRows = Math.max(lastRow - 1, 1);

  const sourceRange = listSheet.getRange(
    2,
    listColumn,
    numRows,
    1
  );

  // Dropdown Rule
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(false)
    .build();

  // Entire usable column (except header)
  const maxRows = targetSheet.getMaxRows();

  targetSheet
    .getRange(2, targetColumn, maxRows - 1, 1)
    .setDataValidation(rule);

}

const BOOKING_DROPDOWNS = [
  "Song Preference",
  "Effects",
  "Status"
];

function syncBookingMasterDropdowns() {

  BOOKING_DROPDOWNS.forEach(header => {
    applyDropdown(
      "Booking Master",
      header,
      header
    );
  });

}


function handleLists(e) {

  const sheet = e.range.getSheet();

  if (sheet.getName() !== "Lists") return;
  if (e.range.getRow() === 1) return;

  syncBookingMasterDropdowns();

}