/*************************************************
 * S Video ERP
 * Module : Booking
 * File   : BookingRepository.gs
 * Version: 1.0.0
 *
 * Responsibility:
 * Data access for Booking Master.
 *
 * Rules:
 * - No business logic.
 * - No validation.
 * - No calculations.
 * - No formatting.
 *************************************************/

/*************************************************
 * Returns Booking sheet.
 *************************************************/
function bookingRepositoryGetSheet() {

  return getSheet(SHEETS.BOOKING);

}

/*************************************************
 * Returns all booking rows
 * excluding the header.
 *************************************************/
function bookingRepositoryGetRows() {

  const sheet = bookingRepositoryGetSheet();

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  return sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      BOOKING_COLUMNS.REMARKS
    )
    .getValues();

}

/*************************************************
 * Returns booking row index
 * (sheet row number).
 *************************************************/
function bookingRepositoryFindRowByBookingNo(
  bookingNo
) {

  const rows =
    bookingRepositoryGetRows();

  for (let i = 0; i < rows.length; i++) {

    if (
      String(rows[i][BOOKING_COLUMNS.BOOKING_NO - 1]).trim() ===
      String(bookingNo).trim()
    ) {

      return i + 2;

    }

  }

  return null;

}

/*************************************************
 * Reads a single booking row.
 *************************************************/
function bookingRepositoryGetRow(
  bookingNo
) {

  const row =
    bookingRepositoryFindRowByBookingNo(
      bookingNo
    );

  if (!row) {
    return null;
  }

  return bookingRepositoryGetSheet()
    .getRange(
      row,
      1,
      1,
      BOOKING_COLUMNS.REMARKS
    )
    .getValues()[0];

}

/*************************************************
 * Writes an entire booking row.
 *************************************************/
function bookingRepositoryUpdateRow(
  row,
  values
) {

  bookingRepositoryGetSheet()
    .getRange(
      row,
      1,
      1,
      values.length
    )
    .setValues([values]);

}

/*************************************************
 * Inserts an entire booking row.
 *************************************************/
function bookingRepositoryInsertRow(
  row,
  values
) {

  bookingRepositoryGetSheet()
    .getRange(
      row,
      1,
      1,
      values.length
    )
    .setValues([values]);

}