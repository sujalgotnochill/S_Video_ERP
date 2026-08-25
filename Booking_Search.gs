/*************************************************
 * S Video ERP
 * Module : Booking Search
 * File   : Booking_Search.gs
 *************************************************/

function showBookingSearchDialog() {
  const html = HtmlService.createTemplateFromFile("BookingSearch").evaluate()
    .setWidth(750)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, "🔍 Search Bookings");
}

function searchBookings(query) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bookingSheet = ss.getSheetByName("Booking Master");
  const studioSheet = ss.getSheetByName("Studio Master");
  
  if (!bookingSheet) return [];
  
  // 1. Build a quick dictionary to match Studio IDs to their City
  const cityMap = {};
  if (studioSheet) {
    const studioData = studioSheet.getDataRange().getValues();
    for (let i = 1; i < studioData.length; i++) {
      const sId = String(studioData[i][0]).trim();
      const sCity = String(studioData[i][6]).trim(); // Column G is City
      if (sId) cityMap[sId] = sCity;
    }
  }
  
  const data = bookingSheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const results = [];
  const q = String(query).toLowerCase().trim();

  // Loop backward to show the newest bookings first
  for (let i = data.length - 1; i > 0; i--) {
    const row = data[i];
    
    // Ignore completely blank ghost rows
    if (!row[0] || String(row[0]).trim() === "") {
      continue; 
    }

    const bookingNo = String(row[0] || "").toLowerCase();
    const studioId = String(row[3] || "").trim();
    const studioName = String(row[4] || "").toLowerCase();
    const mobile = String(row[6] || "").toLowerCase();
    const status = String(row[17] || "");
    
    // Grab the city from our dictionary
    const city = cityMap[studioId] || ""; 
    const cityLower = city.toLowerCase();

    // Now it searches by Booking No, Name, Mobile, OR City!
    if (q === "" || bookingNo.includes(q) || studioName.includes(q) || mobile.includes(q) || cityLower.includes(q)) {
      results.push({
        bookingNo: row[0],
        date: row[1] instanceof Date ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), "dd-MM-yyyy") : row[1],
        studioName: row[4] || "One-Time Customer",
        city: city, // Pass the city to the frontend
        service: row[7] || "N/A",
        status: status
      });
    }
    
    // Stop after finding 20 matches so the UI remains lightning fast
    if (results.length >= 20) break;
  }
  
  return results;
}

function openEditBookingDialog(bookingNo) {
  const bookingData = getBooking(bookingNo); 
  if (!bookingData) throw new Error("Could not find booking data for #" + bookingNo);

  const template = HtmlService.createTemplateFromFile("NewBooking");
  template.editMode = true;
  template.booking = bookingData;

  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(650).setHeight(600),
    "📋 Edit Booking #" + bookingNo
  );
}