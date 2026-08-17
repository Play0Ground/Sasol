/**
 * config.js — App settings (change these if your deploy URL / PIN changes)
 * -----------------------------------------------------------------
 * API_URL  = Google Apps Script Web App /exec URL (Version 7+)
 * MOD_PIN  = Moderator soft PIN (not real security — lives in page source)
 * MONTHS   = Used to format sheet dates like "11 August 2026"
 */
const API_URL = "https://script.google.com/macros/s/AKfycbxp2RyVJLAgsi13190nIHNk4C8WJThjEVbBmAL1uBOtGav1z_20cnbyaVfHtANwdfzMCw/exec";
const MOD_PIN = "2026";
const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const SESSION_KEY = "sasol_session_v1";
const ROSTER_TTL_MS = 60 * 1000;
