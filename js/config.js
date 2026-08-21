/**
 * config.js — App settings
 * -----------------------------------------------------------------
 * API_URL     = Google Apps Script Web App /exec URL
 * MONTHS      = Sheet date words ("11 August 2026")
 *
 * Moderator PIN is NOT stored here anymore.
 * It lives in the Google Sheet tab named "Access", cell B1.
 */
const API_URL = "https://script.google.com/macros/s/AKfycbxp2RyVJLAgsi13190nIHNk4C8WJThjEVbBmAL1uBOtGav1z_20cnbyaVfHtANwdfzMCw/exec";
const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const SESSION_KEY = "sasol_session_v1";
const ROSTER_TTL_MS = 45 * 1000;  // reuse roster so Find Me / dup checks stay fast
const SYNC_POLL_MS = 60 * 1000;   // moderator background sync (Refresh still forces now)
