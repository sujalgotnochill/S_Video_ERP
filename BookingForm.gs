/*************************************************
 * S Video ERP
 * Module : Booking
 * File   : BookingForm.gs
 * Version: 2.0.0
 *
 * Responsibility:
 * Supplies data required by the Booking Form.
 *************************************************/

/*************************************************
 * Opens New Booking Dialog
 *************************************************/
function showNewBookingDialog() {

  const template =
    HtmlService.createTemplateFromFile("NewBooking");

  template.editMode = false;
  template.booking = null;

  SpreadsheetApp
    .getUi()
    .showModalDialog(
      template
        .evaluate()
        .setWidth(
          DIALOG.NEW_BOOKING.WIDTH
        )
        .setHeight(
          DIALOG.NEW_BOOKING.HEIGHT
        ),
      "New Booking"
    );

}

/*************************************************
 * Returns all Studios
 *************************************************/
function getStudios() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(
      SHEETS.STUDIO
    );

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  return sheet
    .getRange(2, 1, lastRow - 1, 8)
    .getValues()
    .filter(row => row[0] && row[1])
    .map(function(row){

      return {

        id: row[0],
        name: row[1],
        displayName: row[2],
        owner: row[3],
        mobile: row[4],
        whatsapp: row[5],
        city: row[6],
        notes: row[7]

      };

    });

}

/*************************************************
 * Returns all Services
 *************************************************/
function getServices() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(
      SHEETS.SERVICE
    );

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  return sheet
    .getRange(2, 1, lastRow - 1, 6)
    .getValues()
    .filter(row => row[0] && row[1])
    .map(function(row){

      return {

        id: row[0],
        name: row[1],
        price: row[2],
        category: row[3],
        active: row[4],
        notes: row[5]

      };

    });

}

/*************************************************
 * Returns Booking Dropdown Lists
 *************************************************/
function getBookingLists() {

  return {

    songPreference:
      getList("Song Preference"),

    effects:
      getList("Effects"),

    status:
      getList("Status")

  };

}

/*************************************************
 * Returns all data required by Booking Form
 *************************************************/
function getBookingFormData() {

  return {

    studios:
      getStudios(),

    services:
      getServices(),

    lists:
      getBookingLists(),

    today:
      formatERPDate(new Date())

  };

}

/*************************************************
 * Creates New Booking
 *************************************************/
function submitBooking(data) {

  return saveBooking(data);

}