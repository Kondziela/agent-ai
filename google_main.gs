function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = new Date().toISOString();

  try {
    // Parse incoming data (JSON body or URL-encoded form)
    let data = {};
    if (e.parameter && Object.keys(e.parameter).length) {
      data = e.parameter;
    } else if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (_) {
        data = e.postData.contents.split('&').reduce(function(acc, p) {
          var parts = p.split('=');
          var k = decodeURIComponent(parts[0] || '');
          var v = decodeURIComponent((parts[1] || '').replace(/\+/g, ' '));
          if (k) acc[k] = v;
          return acc;
        }, {});
      }
    }

    const source   = data.source   || '';
    const campaign = data.campaign || '';
    const ts       = data.timestamp || timestamp;

    // Route: email signup
    if (data.email) {
      let sheet = ss.getSheetByName('signups');
      if (!sheet) {
        sheet = ss.insertSheet('signups');
        sheet.appendRow(['email', 'timestamp', 'source', 'campaign']);
      }
      sheet.appendRow([data.email, ts, source, campaign]);

    // Route: page view
    } else if (data.sessionId) {
      let sheet = ss.getSheetByName('page_views');
      if (!sheet) {
        sheet = ss.insertSheet('page_views');
        sheet.appendRow(['session_id', 'page', 'source', 'campaign', 'timestamp']);
      }
      sheet.appendRow([data.sessionId, data.page || '', source, campaign, ts]);

    } else {
      throw new Error('Unknown event type');
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'OK' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
