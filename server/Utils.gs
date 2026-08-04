// Utils.gs — Own-It Microsite Shared Utilities

var CHAT_URL = PropertiesService.getScriptProperties().getProperty('GCHAT_WEBHOOK') || '';

// Standalone script — open the master workbook by ID.
// ID is read from Script Properties first (MASTER_SHEET_ID), then falls back to the hardcoded value.
var MASTER_SHEET_ID = PropertiesService.getScriptProperties().getProperty('MASTER_SHEET_ID')
                      || '105hKEYXITPS2zD7570sQqOjm2RcQ_aGY8uh9BtSdBvY';

function getMasterSheet() {
  return SpreadsheetApp.openById(MASTER_SHEET_ID);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function logStatus(message, section) {
  Logger.log('[' + (section || 'OWN-IT') + '] ' + message);
}

function safe(v) {
  if (v === null || v === undefined || v === '' || v === false) return 0;
  var n = Number(String(v).replace(/[$,%]/g, '').replace(/\(([^)]+)\)/, '-$1').trim());
  return isNaN(n) ? 0 : n;
}
