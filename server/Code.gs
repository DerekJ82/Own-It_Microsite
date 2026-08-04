// Code.gs — Own-It Microsite Web App Entry Point

function doGet(e) {
  try {
    var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    if (authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED) {
      var authUrl = authInfo.getAuthorizationUrl();
      return HtmlService.createHtmlOutput(
        '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
        '<style>' +
        'body{margin:0;font-family:"Segoe UI",Arial,sans-serif;background:#faf8fc;' +
        'display:flex;align-items:center;justify-content:center;min-height:100vh;}' +
        '.card{background:#fff;border-radius:16px;border:1px solid rgba(75,40,109,0.12);' +
        'box-shadow:0 4px 24px rgba(75,40,109,0.08);padding:48px 40px;max-width:440px;' +
        'width:calc(100% - 48px);text-align:center;}' +
        'h2{color:#4b286d;font-size:1.05rem;font-weight:800;margin:0 0 12px;}' +
        'p{color:#6b7280;font-size:0.78rem;line-height:1.7;margin:0 0 28px;}' +
        'a{display:inline-block;background:#4b286d;color:#fff;text-decoration:none;' +
        'font-size:0.8rem;font-weight:700;padding:13px 30px;border-radius:8px;}' +
        '</style></head><body>' +
        '<div class="card">' +
        '<h2>Own-It Microsite &mdash; Authorization Required</h2>' +
        '<p>New permissions are required. Click below to authorize.</p>' +
        '<a href="' + authUrl + '" target="_top">Authorize &amp; Continue</a>' +
        '</div></body></html>'
      ).setTitle('Own-It — Authorization Required')
       .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (authErr) {
    Logger.log('doGet authInfo error: ' + authErr.message);
  }

  if (e && e.parameter && e.parameter.view) {
    var v = e.parameter.view;
    var wsMap = {
      'ws1': 'client/components/Workstream1',
      'ws2': 'client/components/Workstream2',
      'ws3': 'client/components/Workstream3',
      'ws4': 'client/components/Workstream4',
      'ws5': 'client/components/Workstream5'
    };
    if (wsMap[v]) {
      return HtmlService.createHtmlOutputFromFile(wsMap[v])
        .setTitle('Own-It — ' + v.toUpperCase())
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }

  return HtmlService.createTemplateFromFile('client/Index')
    .evaluate()
    .setTitle('Own-It Microsite')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}
