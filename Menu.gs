/*************************************************
 * S Video ERP
 * Module : Core
 * File   : Menu.gs
 * Version: 2.0.0
 *************************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(APP.NAME)
    .addItem("➕ New Studio", "showNewStudioDialog")
    .addItem("📋 New Booking", "showNewBookingDialog")
    .addItem("🔍 Booking Search", "showBookingSearchDialog")
    .addItem("📊 Dashboard", "showDashboardDialog")
    .addToUi();
}

function showNewStudioDialog() {
  const html = HtmlService
    .createTemplateFromFile("NewStudio")
    .evaluate()
    .setWidth(DIALOG.NEW_STUDIO.WIDTH)
    .setHeight(DIALOG.NEW_STUDIO.HEIGHT);

  SpreadsheetApp.getUi()
    .showModalDialog(html, "New Studio");
}

function showDashboardDialog() {
  const html = HtmlService.createTemplateFromFile("Dashboard_Page").evaluate()
    .setWidth(700)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, "📊 ERP Dashboard");
}

function showBookingSearchDialog() {
  const html = HtmlService
    .createTemplateFromFile("BookingSearch")
    .evaluate()
    .setWidth(DIALOG.BOOKING_SEARCH.WIDTH)
    .setHeight(DIALOG.BOOKING_SEARCH.HEIGHT);

  SpreadsheetApp.getUi()
    .showModalDialog(html, "Booking Search");
}

function showEditBookingDialog(bookingNo) {
  const booking = getBooking(bookingNo);

  if (!booking) {
    throw new Error("Booking not found.");
  }

  const template = HtmlService.createTemplateFromFile("NewBooking");
  template.editMode = true;
  template.booking = booking;

  SpreadsheetApp.getUi()
    .showModalDialog(
      template.evaluate()
        .setWidth(DIALOG.EDIT_BOOKING.WIDTH)
        .setHeight(DIALOG.EDIT_BOOKING.HEIGHT),
      "Edit Booking"
    );
}