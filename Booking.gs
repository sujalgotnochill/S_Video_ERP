/*************************************************
 * Booking.gs
 * Business Logic for Booking Module
 *************************************************/

const BOOKING_SHEET = "Booking Master";

const BOOKING_COLUMNS = {
  BOOKING_NO: 1, BOOKING_DATE: 2, DELIVERY_DATE: 3, STUDIO_ID: 4, STUDIO_NAME: 5,
  JOB_TITLE: 6, MOBILE_NUMBER: 7, SERVICE: 8, QUANTITY: 9, SONG_PREFERENCE: 10,
  EFFECTS: 11, DRONE: 12, PENDRIVE: 13, HARD_DISK: 14, MEMORY_CARD: 15,
  ASSIGNED_MIXER: 16, QC: 17, STATUS: 18, ADVANCE: 19, AMOUNT: 20, BALANCE: 21, REMARKS: 22
};

function getBookingSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BOOKING_SHEET);
}

function generateBookingNumber() {
  return generateNextBookingNo(getBookingSheet());
}

function getService(serviceId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Services Master");
  if (!sheet) return null;
  const data = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 6).getValues();
  const row = data.find(r => r[0] === serviceId);
  return row ? { id: row[0], name: row[1], price: Number(row[2]) || 0 } : null;
}

function saveBooking(data) {
  const sheet = getBookingSheet();
  const bookingNo = generateBookingNumber();
  
  const advance = Number(data.advancePaid) || 0;
  // If quantity is blank, leave it blank. Don't force a 1.
  const quantity = data.quantity === "" ? "" : Number(data.quantity);
  
  const serviceObj = getService(data.serviceId);
  const serviceName = serviceObj ? serviceObj.name : "";
  const unitPrice = serviceObj ? serviceObj.price : 0;
  
  // Calculate Amount ONLY if Quantity is provided
  const amount = quantity !== "" ? quantity * unitPrice : "";
  // Calculate Balance: If Amount is blank, just show negative advance (or blank).
  const balance = amount !== "" ? amount - advance : (advance > 0 ? -advance : "");

  const bookingNumbers = sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 1).getValues().flat();
  const emptyIndex = bookingNumbers.findIndex(value => value === "" || value === null);
  const row = emptyIndex === -1 ? sheet.getLastRow() + 1 : emptyIndex + 2;

  sheet.getRange(row, 1, 1, 22).setValues([[
    bookingNo, data.bookingDate, data.deliveryDate, data.studioId || "", data.studioName,
    data.jobTitle, data.mobileNumber, serviceName, quantity, data.songPreference, data.effects,
    data.droneFootage ? true : false, data.pendrive ? true : false, data.hardDisk ? true : false,
    data.memoryCard ? true : false, data.assignedMixer || "", data.qc || "", data.status || "Pending", 
    advance, amount, balance, data.remarks || ""
  ]]);

  return { success: true, bookingNo: bookingNo };
}

function getBooking(bookingNo) {
  const sheet = getBookingSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, 22).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(bookingNo)) {
      const r = data[i];
      const formatVal = (v) => v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd") : v;
      return {
        row: i + 2, bookingNo: r[0], bookingDate: formatVal(r[1]), deliveryDate: formatVal(r[2]),
        studioId: r[3], studioName: r[4], jobTitle: r[5], mobileNumber: r[6], serviceName: r[7],
        quantity: r[8], songPreference: r[9], effects: r[10], drone: Boolean(r[11]), pendrive: Boolean(r[12]),
        hardDisk: Boolean(r[13]), memoryCard: Boolean(r[14]), assignedMixer: r[15], qc: r[16],
        status: r[17], advance: r[18], amount: r[19], balance: r[20], remarks: r[21]
      };
    }
  }
  return null;
}

function updateBooking(data) {
  const sheet = getBookingSheet();
  const row = Number(data.row);
  if (!row || row < 2) throw new Error("Invalid row index.");

  const existing = sheet.getRange(row, 1, 1, 22).getValues()[0];
  
  const advance = Number(data.advancePaid !== undefined ? data.advancePaid : existing[18]) || 0;
  const quantity = data.quantity !== undefined ? (data.quantity === "" ? "" : Number(data.quantity)) : existing[8];
  
  const serviceObj = getService(data.serviceId);
  const serviceName = serviceObj ? serviceObj.name : (data.serviceName || existing[7]);
  const unitPrice = serviceObj ? serviceObj.price : 0;
  
  // Calculate Amount ONLY if Quantity is provided
  const amount = quantity !== "" ? quantity * unitPrice : "";
  const balance = amount !== "" ? amount - advance : (advance > 0 ? -advance : "");

  sheet.getRange(row, 1, 1, 22).setValues([[
    data.bookingNo || existing[0], data.bookingDate, data.deliveryDate, data.studioId || "", data.studioName,
    data.jobTitle, data.mobileNumber, serviceName, quantity, data.songPreference, data.effects,
    data.droneFootage ? true : false, data.pendrive ? true : false, data.hardDisk ? true : false,
    data.memoryCard ? true : false, data.assignedMixer || existing[15], data.qc || existing[16], 
    data.status || existing[17], advance, amount, balance, data.remarks || ""
  ]]);

  return { success: true, bookingNo: data.bookingNo };
}