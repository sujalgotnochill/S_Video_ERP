/*************************************************
 * S Video ERP
 * Module : Core
 * File   : Utilities.gs
 * Version: 1.0.0
 *
 * Responsibility:
 * Generic helper functions used throughout the ERP.
 *************************************************/

/*************************************************
 * Returns TRUE if value is blank.
 *************************************************/
function isBlank(value) {

  return value === null ||
         value === undefined ||
         String(value).trim() === "";

}

/*************************************************
 * Safely converts value to number.
 *************************************************/
function toNumber(value) {

  const number = Number(value);

  return isNaN(number)
    ? 0
    : number;

}

/*************************************************
 * Formats a date.
 *************************************************/
function formatERPDate(value) {

  if (!value) {
    return "";
  }

  return Utilities.formatDate(
    new Date(value),
    Session.getScriptTimeZone(),
    BOOKING_CONFIG.DATE_FORMAT
  );

}

/*************************************************
 * Generates next ID.
 *************************************************/
function generateNextId(sheet, prefix) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return prefix + "001";
  }

  const ids = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat();

  let max = 0;

  ids.forEach(function(id){

    if (
      typeof id === "string" &&
      id.startsWith(prefix)
    ) {

      const number =
        parseInt(id.replace(prefix, ""), 10);

      if (!isNaN(number)) {
        max = Math.max(max, number);
      }

    }

  });

  return prefix + String(max + 1).padStart(3, "0");

}

/*************************************************
 * Generates next Booking Number.
 *************************************************/
function generateNextBookingNo(sheet) {

  const values =
    sheet.getRange("A:A")
      .getValues();

  for (let i = values.length - 1; i >= 1; i--) {

    const bookingNo =
      parseInt(values[i][0], 10);

    if (!isNaN(bookingNo)) {
      return bookingNo + 1;
    }

  }

  return BOOKING_CONFIG.START_NUMBER;

}