// Auth.gs — Own-It Microsite Role-Based Access Control

var ACCESS_COLUMN_MAP = {
  'WORKSTREAM 1 - PRICING':        'WS1 PRICING',
  'WORKSTREAM 2 - BID GOVERNANCE': 'WS2 BID GOV',
  'WORKSTREAM 3 - CREDIT A-R':     'WS3 CREDIT',
  'WORKSTREAM 4 - CONTRACT':       'WS4 CONTRACT',
  'WORKSTREAM 5 - PROPOSAL':       'WS5 PROPOSAL',
  'ALL WORKSTREAMS':               'ALL ACCESS'
};

var ACCESS_TAB_NAME    = 'Admin - Access';
var ACCESS_REQ_TAB     = 'Access Requests';
var ADMINS_COLUMN      = 'ADMINS';
var ALL_ACCESS_COLUMN  = 'ALL ACCESS';

function checkAccess(permissionColumn) {
  try {
    var email = Session.getEffectiveUser().getEmail().trim().toLowerCase();
    var ownerEmail = (PropertiesService.getScriptProperties().getProperty('OWNER_EMAIL') || '').trim().toLowerCase();
    if (ownerEmail && email === ownerEmail) return true;

    var ss  = getMasterSheet();
    var tab = ss.getSheetByName(ACCESS_TAB_NAME);
    if (!tab) {
      Logger.log('checkAccess: ' + ACCESS_TAB_NAME + ' tab not found');
      return false;
    }

    var data    = tab.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });

    var adminIdx = headers.indexOf(ADMINS_COLUMN);
    var allIdx   = headers.indexOf(ALL_ACCESS_COLUMN);
    var permIdx  = headers.indexOf(String(permissionColumn).trim().toUpperCase());

    // Each row may have an email in any one column — scan the relevant columns across all rows
    for (var r = 1; r < data.length; r++) {
      if (adminIdx >= 0 && String(data[r][adminIdx] || '').trim().toLowerCase() === email) return true;
      if (allIdx   >= 0 && String(data[r][allIdx]   || '').trim().toLowerCase() === email) return true;
      if (permIdx  >= 0 && String(data[r][permIdx]  || '').trim().toLowerCase() === email) return true;
    }

    Logger.log('checkAccess: DENIED for ' + email + ' column=' + permissionColumn);
    return false;
  } catch (err) {
    Logger.log('checkAccess error: ' + err.message);
    return false;
  }
}

function getSessionEmail() {
  try {
    return { email: Session.getEffectiveUser().getEmail() };
  } catch (err) {
    Logger.log('getSessionEmail error: ' + err.message);
    return { email: '' };
  }
}

function getAccessColumnDescriptions() {
  return [
    { key: 'WORKSTREAM 1 - PRICING',        label: 'Pricing',            desc: 'Deal support volumes, win/loss analysis, cost models' },
    { key: 'WORKSTREAM 2 - BID GOVERNANCE', label: 'Bid Governance',     desc: 'Approval matrix, deal volumes, bid waiver tools' },
    { key: 'WORKSTREAM 3 - CREDIT A-R',     label: 'Credit / A-R',       desc: 'Credit policy, decision criteria, documentation' },
    { key: 'WORKSTREAM 4 - CONTRACT',       label: 'Contract',            desc: 'Contract clauses, pain points, AI tools' },
    { key: 'WORKSTREAM 5 - PROPOSAL',       label: 'Proposal & Content',  desc: 'Loopio, Highspot, proposal automation' }
  ];
}

function getAccessDiag() {
  try {
    var email = Session.getEffectiveUser().getEmail().trim().toLowerCase();
    var ss    = getMasterSheet();
    var tab   = ss.getSheetByName(ACCESS_TAB_NAME);
    if (!tab) return { error: ACCESS_TAB_NAME + ' tab not found', email: email };

    var data    = tab.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var found   = [];

    for (var r = 1; r < data.length; r++) {
      for (var c = 0; c < data[r].length; c++) {
        if (String(data[r][c] || '').trim().toLowerCase() === email) {
          found.push(headers[c] || ('Column ' + c));
        }
      }
    }

    return { email: email, columnsFound: found };
  } catch (err) {
    return { error: err.message };
  }
}

function syncAccessColumnSchema() {
  var ss  = getMasterSheet();
  var tab = ss.getSheetByName(ACCESS_TAB_NAME);
  if (!tab) {
    Logger.log('syncAccessColumnSchema: ' + ACCESS_TAB_NAME + ' tab not found. Create it first.');
    return;
  }

  var headers = tab.getRange(1, 1, 1, tab.getLastColumn()).getValues()[0]
                   .map(function(h) { return String(h).trim().toUpperCase(); });
  var required = [ADMINS_COLUMN, 'WS1 PRICING', 'WS2 BID GOV', 'WS3 CREDIT', 'WS4 CONTRACT', 'WS5 PROPOSAL', ALL_ACCESS_COLUMN];

  required.forEach(function(col) {
    if (headers.indexOf(col) === -1) {
      var nextCol = tab.getLastColumn() + 1;
      tab.getRange(1, nextCol).setValue(col);
      Logger.log('syncAccessColumnSchema: added column ' + col);
    }
  });

  Logger.log('syncAccessColumnSchema: done');
}

function installAccessRequestTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'onEditAccessRequests') ScriptApp.deleteTrigger(t);
  });

  var ss = getMasterSheet();
  ScriptApp.newTrigger('onEditAccessRequests')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  Logger.log('installAccessRequestTrigger: installed onEditAccessRequests trigger');
}

function onEditAccessRequests(e) {
  try {
    var sheet = e.range.getSheet();
    if (sheet.getName() !== ACCESS_REQ_TAB) return;

    var editedCol = e.range.getColumn();
    var editedRow = e.range.getRow();
    if (editedRow < 2) return;

    // Column H (8) is the approve checkbox
    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var approveColIdx = headerRow.indexOf('✓ Approve') + 1;
    if (approveColIdx === 0) approveColIdx = 8;

    if (editedCol !== approveColIdx) return;
    if (!e.range.getValue()) return;

    var row         = sheet.getRange(editedRow, 1, 1, 8).getValues()[0];
    var reqEmail    = String(row[0] || '').trim().toLowerCase();
    var accessCol   = String(row[2] || '').trim().toUpperCase();
    if (!reqEmail || !accessCol) return;

    var ss       = getMasterSheet();
    var tab      = ss.getSheetByName(ACCESS_TAB_NAME);
    if (!tab) return;

    var headers  = tab.getRange(1, 1, 1, tab.getLastColumn()).getValues()[0]
                      .map(function(h) { return String(h).trim().toUpperCase(); });
    var colIndex = headers.indexOf(accessCol);
    if (colIndex === -1) return;

    // Find empty row in that column
    var colData = tab.getRange(2, colIndex + 1, tab.getMaxRows() - 1, 1).getValues();
    var targetRow = 2;
    for (var i = 0; i < colData.length; i++) {
      if (!String(colData[i][0] || '').trim()) { targetRow = i + 2; break; }
    }
    tab.getRange(targetRow, colIndex + 1).setValue(reqEmail);

    // Mark as approved
    var approverEmail = Session.getEffectiveUser().getEmail();
    sheet.getRange(editedRow, 6).setValue(approverEmail);
    sheet.getRange(editedRow, 7).setValue(new Date());
    sheet.getRange(editedRow, 5).setValue('Approved');

    Logger.log('onEditAccessRequests: granted ' + accessCol + ' to ' + reqEmail);
  } catch (err) {
    Logger.log('onEditAccessRequests error: ' + err.message);
  }
}
