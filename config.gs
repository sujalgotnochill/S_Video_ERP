/*************************************************
 * S Video ERP
 * Module : Core
 * File   : Config.gs
 * Version: 1.0.0
 *
 * Responsibility:
 * Central configuration for the entire ERP.
 *************************************************/

/*************************************************
 * Application
 *************************************************/

const APP = {
  NAME: "S Video ERP",
  VERSION: "1.0.0"
};

/*************************************************
 * Sheet Names
 *************************************************/

const SHEETS = {
  BOOKING: "Booking Master",
  STUDIO: "Studio Master",
  SERVICE: "Services Master",
  STAFF: "Staff Master",
  LISTS: "Lists"
};

/*************************************************
 * Booking Configuration
 *************************************************/

const BOOKING_CONFIG = {
  START_NUMBER: 34396,
  DEFAULT_STATUS: "Pending",
  DATE_FORMAT: "yyyy-MM-dd"
};

/*************************************************
 * Dialog Sizes
 *************************************************/

const DIALOG = {

  NEW_BOOKING: {
    WIDTH: 900,
    HEIGHT: 700
  },

  EDIT_BOOKING: {
    WIDTH: 900,
    HEIGHT: 760
  },

  BOOKING_SEARCH: {
    WIDTH: 1200,
    HEIGHT: 700
  },

  NEW_STUDIO: {
    WIDTH: 500,
    HEIGHT: 650
  }

};