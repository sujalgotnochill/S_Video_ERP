/*************************************************
 * S Video ERP
 * Module : Booking
 * File   : BookingForm.gs
 * Version: 2.2.0 (Display Names & QC)
 *************************************************/

function showNewBookingDialog() {
  const template = HtmlService.createTemplateFromFile("NewBooking");
  template.editMode = false;
  template.booking = null;

  const width = (typeof DIALOG !== "undefined" && DIALOG.NEW_BOOKING) ? DIALOG.NEW_BOOKING.WIDTH : 650;
  const height = (typeof DIALOG !== "undefined" && DIALOG.NEW_BOOKING) ? DIALOG.NEW_BOOKING.HEIGHT : 600;

  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(width).setHeight(height),
    "New Booking / Edit"
  );
}

function getStudios() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Studio Master");
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, 8).getValues()
    .filter(row => row[0] && row[1])
    .map(row => ({
      id: String(row[0] || ""),
      name: String(row[1] || ""),
      displayName: String(row[2] || (row[1] + " - " + row[3] + " - " + row[6])),
      owner: String(row[3] || ""),
      mobile: String(row[4] || ""),
      whatsapp: String(row[5] || ""),
      city: String(row[6] || ""),
      notes: String(row[7] || "")
    }));
}

function getServices() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Services Master");
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, 6).getValues()
    .filter(row => row[0] && row[1])
    .map(row => ({
      id: String(row[0] || ""),
      name: String(row[1] || ""),
      price: Number(row[2]) || 0,
      category: String(row[3] || ""),
      active: Boolean(row[4]),
      notes: String(row[5] || "")
    }));
}

function getStaff() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Staff Master");
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Fetch 3 columns so we can access Column C (Display Name)
  return sheet.getRange(2, 1, lastRow - 1, 3).getValues()
    .filter(row => row[0] && row[1])
    .map(row => ({
      id: String(row[0] || ""),
      name: String(row[2] || row[1] || "") // Grabs Column C (Display Name), falls back to B if empty
    }));
}

function getBookingLists() {
  let songPref = [], fx = [], stat = [];
  try { songPref = getList("Song Preference"); } catch (e) { songPref = ["STANDARD", "CUSTOM"]; }
  try { fx = getList("Effects"); } catch (e) { fx = ["NORMAL", "CINEMATIC"]; }
  try { stat = getList("Status"); } catch (e) { stat = ["Pending", "In Progress", "Completed", "Delivered"]; }

  return { songPreference: songPref, effects: fx, status: stat };
}

function getBookingFormData() {
  return {
    studios: getStudios(),
    services: getServices(),
    staff: getStaff(),
    lists: getBookingLists(),
    today: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")
  };
}

function submitBooking(data) {
  return saveBooking(data);
}