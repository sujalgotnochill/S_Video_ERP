/*************************************************
 * S Video ERP
 * Module : Booking
 * File   : BookingMapper.gs
 * Version: 1.0.0
 *
 * Responsibility:
 * Converts between sheet rows and Booking objects.
 *
 * This is the ONLY file that should know the
 * Booking Master column layout.
 *************************************************/

/*************************************************
 * Sheet Row -> Booking Model
 *************************************************/
function bookingRowToModel(row, rowNumber) {

  return {

    row: rowNumber,

    bookingNo: row[BOOKING_COLUMNS.BOOKING_NO - 1],

    bookingDate: formatERPDate(
      row[BOOKING_COLUMNS.BOOKING_DATE - 1]
    ),

    deliveryDate: formatERPDate(
      row[BOOKING_COLUMNS.DELIVERY_DATE - 1]
    ),

    studioId: row[BOOKING_COLUMNS.STUDIO_ID - 1],

    studioName: row[BOOKING_COLUMNS.STUDIO_NAME - 1],

    jobTitle: row[BOOKING_COLUMNS.JOB_TITLE - 1],

    mobileNumber: row[BOOKING_COLUMNS.MOBILE_NUMBER - 1],

    serviceName: row[BOOKING_COLUMNS.SERVICE - 1],

    quantity: row[BOOKING_COLUMNS.QUANTITY - 1],

    songPreference: row[BOOKING_COLUMNS.SONG_PREFERENCE - 1],

    effects: row[BOOKING_COLUMNS.EFFECTS - 1],

    drone: row[BOOKING_COLUMNS.DRONE - 1],

    pendrive: row[BOOKING_COLUMNS.PENDRIVE - 1],

    hardDisk: row[BOOKING_COLUMNS.HARD_DISK - 1],

    memoryCard: row[BOOKING_COLUMNS.MEMORY_CARD - 1],

    assignedMixer: row[BOOKING_COLUMNS.ASSIGNED_MIXER - 1],

    qc: row[BOOKING_COLUMNS.QC - 1],

    status: row[BOOKING_COLUMNS.STATUS - 1],

    advance: row[BOOKING_COLUMNS.ADVANCE - 1],

    amount: row[BOOKING_COLUMNS.AMOUNT - 1],

    balance: row[BOOKING_COLUMNS.BALANCE - 1],

    remarks: row[BOOKING_COLUMNS.REMARKS - 1]

  };

}

/*************************************************
 * Booking Model -> Sheet Row
 *************************************************/
function bookingModelToRow(booking) {

  return [

    booking.bookingNo,
    booking.bookingDate,
    booking.deliveryDate,
    booking.studioId,
    booking.studioName,
    booking.jobTitle,
    booking.mobileNumber,
    booking.serviceName,
    booking.quantity,
    booking.songPreference,
    booking.effects,
    booking.drone,
    booking.pendrive,
    booking.hardDisk,
    booking.memoryCard,
    booking.assignedMixer,
    booking.qc,
    booking.status,
    booking.advance,
    booking.amount,
    booking.balance,
    booking.remarks

  ];

}