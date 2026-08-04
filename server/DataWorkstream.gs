// DataWorkstream.gs — Own-It Microsite Workstream Data Functions

function getWorkstream1Data() {
  if (!checkAccess('WS1 PRICING')) return { error: 'ACCESS_DENIED' };
  try {
    var ss  = getMasterSheet();
    var tab = ss.getSheetByName('Workstream 1 - Pricing');
    if (!tab) return { error: 'Tab not found: Workstream 1 - Pricing' };
    var rows = tab.getDataRange().getValues();
    return { data: rows, error: false };
  } catch (e) {
    Logger.log('getWorkstream1Data error: ' + e.message);
    return { error: e.message };
  }
}

function getWorkstream2Data() {
  if (!checkAccess('WS2 BID GOV')) return { error: 'ACCESS_DENIED' };
  try {
    var ss  = getMasterSheet();
    var tab = ss.getSheetByName('Workstream 2 - Bid Governance');
    if (!tab) return { error: 'Tab not found: Workstream 2 - Bid Governance' };
    var rows = tab.getDataRange().getValues();
    return { data: rows, error: false };
  } catch (e) {
    Logger.log('getWorkstream2Data error: ' + e.message);
    return { error: e.message };
  }
}

function getWorkstream3Data() {
  if (!checkAccess('WS3 CREDIT')) return { error: 'ACCESS_DENIED' };
  try {
    var ss  = getMasterSheet();
    var tab = ss.getSheetByName('Workstream 3 - Credit AR');
    if (!tab) return { error: 'Tab not found: Workstream 3 - Credit AR' };
    var rows = tab.getDataRange().getValues();
    return { data: rows, error: false };
  } catch (e) {
    Logger.log('getWorkstream3Data error: ' + e.message);
    return { error: e.message };
  }
}

function getWorkstream4Data() {
  if (!checkAccess('WS4 CONTRACT')) return { error: 'ACCESS_DENIED' };
  try {
    var ss  = getMasterSheet();
    var tab = ss.getSheetByName('Workstream 4 - Contract');
    if (!tab) return { error: 'Tab not found: Workstream 4 - Contract' };
    var rows = tab.getDataRange().getValues();
    return { data: rows, error: false };
  } catch (e) {
    Logger.log('getWorkstream4Data error: ' + e.message);
    return { error: e.message };
  }
}

function getWorkstream5Data() {
  if (!checkAccess('WS5 PROPOSAL')) return { error: 'ACCESS_DENIED' };
  try {
    var ss  = getMasterSheet();
    var tab = ss.getSheetByName('Workstream 5 - Proposal Content');
    if (!tab) return { error: 'Tab not found: Workstream 5 - Proposal Content' };
    var rows = tab.getDataRange().getValues();
    return { data: rows, error: false };
  } catch (e) {
    Logger.log('getWorkstream5Data error: ' + e.message);
    return { error: e.message };
  }
}

// Notes tab: A=WorkstreamNum, B=Issue, C=Criticality, D=Actioned, E=Submitted By, F=Submitted At
function getWorkstreamNotes(wsNum) {
  try {
    var ss  = getMasterSheet();
    var tab = ss.getSheetByName('Notes');
    if (!tab) return { rows: [] };
    var data = tab.getDataRange().getValues();
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(wsNum)) {
        var ts = data[i][5];
        var tsStr = '';
        try { tsStr = ts ? Utilities.formatDate(new Date(ts), Session.getScriptTimeZone(), 'MMM d HH:mm') : ''; } catch(e2) {}
        rows.push({
          rowNum:      i + 1,
          issue:       String(data[i][1] || ''),
          criticality: String(data[i][2] || ''),
          actioned:    data[i][3] === true || String(data[i][3]).toUpperCase() === 'TRUE',
          submittedBy: String(data[i][4] || ''),
          submittedAt: tsStr
        });
      }
    }
    return { rows: rows };
  } catch (e) {
    Logger.log('getWorkstreamNotes error: ' + e.message);
    return { rows: [] };
  }
}

function addWorkstreamNote(wsNum, issue, criticality) {
  try {
    var ss  = getMasterSheet();
    var tab = ss.getSheetByName('Notes');
    if (!tab) {
      tab = ss.insertSheet('Notes');
      tab.getRange(1, 1, 1, 6).setValues([['WorkstreamNum', 'Issue', 'Criticality', 'Actioned', 'Submitted By', 'Submitted At']]);
    }
    var email   = Session.getEffectiveUser().getEmail();
    var now     = new Date();
    var nextRow = tab.getLastRow() + 1;
    tab.getRange(nextRow, 1, 1, 6).setValues([[Number(wsNum), issue, criticality || '', false, email, now]]);
    var tsStr = '';
    try { tsStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM d HH:mm'); } catch(e2) {}
    return {
      success: true,
      row: { rowNum: nextRow, issue: issue, criticality: criticality || '', actioned: false, submittedBy: email, submittedAt: tsStr }
    };
  } catch (e) {
    Logger.log('addWorkstreamNote error: ' + e.message);
    return { success: false, error: e.message };
  }
}

function setNoteActioned(rowNum, actioned) {
  try {
    var ss  = getMasterSheet();
    var tab = ss.getSheetByName('Notes');
    if (!tab) return { success: false, error: 'Notes tab not found' };
    tab.getRange(rowNum, 4).setValue(actioned ? true : false);
    return { success: true };
  } catch (e) {
    Logger.log('setNoteActioned error: ' + e.message);
    return { success: false, error: e.message };
  }
}
