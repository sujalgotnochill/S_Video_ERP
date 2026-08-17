function searchBookingRecords(keyword) {

  return searchBookings(keyword);

}

function searchBookings(keyword) {

  const sheet = getBookingSheet();

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const data = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      BOOKING_COLUMNS.REMARKS
    )
    .getValues();

  keyword = String(keyword || "")
    .trim()
    .toLowerCase();

  const results = [];

  data.forEach(function(row){

    if (!row[BOOKING_COLUMNS.BOOKING_NO - 1]) {
      return;
    }

    const booking = {

      bookingNo:
        row[BOOKING_COLUMNS.BOOKING_NO - 1],

      bookingDate:
        row[BOOKING_COLUMNS.BOOKING_DATE - 1]
          ? Utilities.formatDate(
              new Date(row[BOOKING_COLUMNS.BOOKING_DATE - 1]),
              Session.getScriptTimeZone(),
              "yyyy-MM-dd"
            )
          : "",

      deliveryDate:
        row[BOOKING_COLUMNS.DELIVERY_DATE - 1]
          ? Utilities.formatDate(
              new Date(row[BOOKING_COLUMNS.DELIVERY_DATE - 1]),
              Session.getScriptTimeZone(),
              "yyyy-MM-dd"
            )
          : "",

      studioId:
        row[BOOKING_COLUMNS.STUDIO_ID - 1],

      studioName:
        row[BOOKING_COLUMNS.STUDIO_NAME - 1],

      mobileNumber:
        row[BOOKING_COLUMNS.MOBILE_NUMBER - 1],

      jobTitle:
        row[BOOKING_COLUMNS.JOB_TITLE - 1],

      serviceName:
        row[BOOKING_COLUMNS.SERVICE - 1],

      status:
        row[BOOKING_COLUMNS.STATUS - 1]

    };

    const searchable = (

      String(booking.bookingNo) + " " +
      String(booking.studioName) + " " +
      String(booking.mobileNumber) + " " +
      String(booking.jobTitle) + " " +
      String(booking.serviceName) + " " +
      String(booking.status)

    ).toLowerCase();

    if (
      keyword === "" ||
      searchable.includes(keyword)
    ) {

      results.push(booking);

    }

  });

  Logger.log(JSON.stringify(results, null, 2));

  return results;

}

function testSearch() {
  Logger.log(JSON.stringify(searchBookings(""), null, 2));
}


