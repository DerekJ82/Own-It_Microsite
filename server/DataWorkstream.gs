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

// Notes are stored in a 'Notes' tab: col A = wsNum (1-5), col B = notes text
function getWorkstreamNotes(wsNum) {
  try {
    var ss  = getMasterSheet();
    var tab = ss.getSheetByName('Notes');
    if (!tab) return { notes: '' };

    var data = tab.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(wsNum)) {
        return { notes: String(data[i][1] || '') };
      }
    }
    return { notes: '' };
  } catch (e) {
    Logger.log('getWorkstreamNotes error: ' + e.message);
    return { notes: '' };
  }
}

function saveWorkstreamNotes(wsNum, text) {
  try {
    var ss    = getMasterSheet();
    var tab   = ss.getSheetByName('Notes');
    if (!tab) {
      tab = ss.insertSheet('Notes');
      tab.getRange(1, 1, 1, 4).setValues([['WorkstreamNum', 'Notes', 'Last Updated By', 'Last Updated At']]);
    }

    var email = Session.getEffectiveUser().getEmail();
    var now   = new Date();
    var data  = tab.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(wsNum)) {
        tab.getRange(i + 1, 2, 1, 3).setValues([[text, email, now]]);
        return { success: true };
      }
    }

    // No existing row — append
    var nextRow = tab.getLastRow() + 1;
    tab.getRange(nextRow, 1, 1, 4).setValues([[wsNum, text, email, now]]);
    return { success: true };
  } catch (e) {
    Logger.log('saveWorkstreamNotes error: ' + e.message);
    return { success: false, error: e.message };
  }
}
