/*************************************************
 * Booking.gs
 * Business Logic for Booking Module
 *************************************************/

const BOOKING_SHEET = "Booking Master";

const BOOKING_COLUMNS = {
  BOOKING_NO: 1,
  BOOKING_DATE: 2,
  DELIVERY_DATE: 3,
  STUDIO_ID: 4,
  STUDIO_NAME: 5,
  JOB_TITLE: 6,
  MOBILE_NUMBER: 7,
  SERVICE: 8,
  QUANTITY: 9,
  SONG_PREFERENCE: 10,
  EFFECTS: 11,
  DRONE: 12,
  PENDRIVE: 13,
  HARD_DISK: 14,
  MEMORY_CARD: 15,
  ASSIGNED_MIXER: 16,
  QC: 17,
  STATUS: 18,
  ADVANCE: 19,
  AMOUNT: 20,
  BALANCE: 21,
  REMARKS: 22
};

/*************************************************
 * Returns Booking Master sheet
 *************************************************/
function getBookingSheet() {

  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(BOOKING_SHEET);

}

/*************************************************
 * Generates next Booking Number
 *************************************************/
function generateBookingNumber() {

  return generateNextBookingNo(
    getBookingSheet()
  );

}

/*************************************************
 * Returns Service by ID
 *************************************************/
function getService(serviceId) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Services Master");

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return null;

  const data = sheet
    .getRange(2, 1, lastRow - 1, 6)
    .getValues();

  const row = data.find(r => r[0] === serviceId);

  if (!row) return null;

  return {
    id: row[0],
    name: row[1],
    price: Number(row[2]) || 0,
    category: row[3],
    active: row[4],
    notes: row[5]
  };

}

/*************************************************
 * Returns Service Price
 *************************************************/
function getServicePrice(serviceId) {

  const service = getService(serviceId);

  return service
    ? service.price
    : 0;

}

/*************************************************
 * Calculates Balance
 *************************************************/
function calculateBalance(amount, advance) {

  amount = Number(amount) || 0;
  advance = Number(advance) || 0;

  return amount - advance;

}

/*************************************************
 * Saves Booking
 *************************************************/
function saveBooking(data) {

  const sheet = getBookingSheet();

  const bookingNo =
    generateBookingNumber();

  const advance =
    Number(data.advancePaid) || 0;

  const defaultStatus =
    getList("Status")[0] || "Pending";

  const bookingNumbers = sheet
    .getRange(
      2,
      BOOKING_COLUMNS.BOOKING_NO,
      sheet.getMaxRows() - 1,
      1
    )
    .getValues()
    .flat();

  const emptyIndex =
    bookingNumbers.findIndex(
      value => value === ""
    );

  const row =
    emptyIndex === -1
      ? sheet.getLastRow() + 1
      : emptyIndex + 2;

  sheet
    .getRange(row, 1, 1, 22)
    .setValues([[
      bookingNo,
      data.bookingDate,
      data.deliveryDate,
      data.studioId,
      data.studioName,
      data.jobTitle,
      data.mobileNumber,
      getService(data.serviceId).name,
      "",                 // Quantity
      data.songPreference,
      data.effects,
      data.droneFootage,
      data.pendrive,
      data.hardDisk,
      data.memoryCard,
      "",                 // Assigned Mixer
      "",                 // QC
      defaultStatus,
      advance,
      "",                 // Amount
      "",                 // Balance
      data.remarks
    ]]);


  return {
    success: true,
    bookingNo: bookingNo
  };

}

/*************************************************
 * Returns Booking by Booking Number
 *************************************************/
/*************************************************
 * Returns Booking by Booking Number
 *************************************************/
function getBooking(bookingNo) {

  const row =
    bookingRepositoryGetRow(bookingNo);

  if (!row) {
    return null;
  }

  const sheetRow =
    bookingRepositoryFindRowByBookingNo(
      bookingNo
    );

  return bookingRowToModel(
    row,
    sheetRow
  );

}

/*************************************************
 * Returns Studio Map
 *************************************************/
function getStudioMap() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Studio Master");

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return {};

  const data = sheet
    .getRange(2, 1, lastRow - 1, 8)
    .getValues();

  const map = {};

  data.forEach(row => {

    map[row[0]] = {
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

  return map;

}


/*************************************************
 * Returns Service Map
 *************************************************/
function getServiceMap() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Services Master");

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return {};

  const data = sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .getValues();

  const map = {};

  data.forEach(row => {

    map[row[0]] = {
      id: row[0],
      name: row[1],
      price: Number(row[2]) || 0
    };

  });

  return map;

}


/*************************************************
 * Update Booking
 *************************************************/
function updateBooking(data) {

  const sheet = getBookingSheet();

  const row = Number(data.row);

  const amount =
    Number(data.quantity || 0) *
    getServicePrice(data.serviceId);

  const balance =
    amount - Number(data.advance || 0);

  sheet
    .getRange(row, 1, 1, 22)
    .setValues([[
      data.bookingNo,
      data.bookingDate,
      data.deliveryDate,
      data.studioId,
      data.studioName,
      data.jobTitle,
      data.mobileNumber,
      data.serviceName,
      data.quantity,
      data.songPreference,
      data.effects,
      data.drone,
      data.pendrive,
      data.hardDisk,
      data.memoryCard,
      data.assignedMixer,
      data.qc,
      data.status,
      data.advance,
      amount,
      balance,
      data.remarks
    ]]);

  return {
    success: true,
    amount: amount,
    balance: balance
  };

}

