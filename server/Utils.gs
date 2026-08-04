// Utils.gs — Own-It Microsite Shared Utilities

var CHAT_URL = PropertiesService.getScriptProperties().getProperty('GCHAT_WEBHOOK') || '';

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
