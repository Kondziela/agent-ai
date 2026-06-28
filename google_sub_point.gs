function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSheet();
  let email = '';
  let timestamp = new Date().toISOString();

  try {
    // If JSON was sent
    if (e.postData && e.postData.type && e.postData.type.indexOf('application/json') !== -1) {
      const data = JSON.parse(e.postData.contents || '{}');
      email = data.email || '';
      timestamp = data.timestamp || timestamp;

    // If URL-encoded form data is available via e.parameter
    } else if (e.parameter && Object.keys(e.parameter).length) {
      email = e.parameter.email || '';
      timestamp = e.parameter.timestamp || timestamp;

    // Fallback: parse postData.contents manually (e.g., "email=...&timestamp=...")
    } else if (e.postData && e.postData.contents) {
      const params = e.postData.contents.split('&').reduce(function(acc, p) {
        var parts = p.split('=');
        var k = decodeURIComponent(parts[0] || '');
        var v = decodeURIComponent((parts[1] || '').replace(/\+/g, ' '));
        if (k) acc[k] = v;
        return acc;
      }, {});
      email = params.email || '';
      timestamp = params.timestamp || timestamp;
    }

    if (!email) throw new Error('Missing email');

    sheet.appendRow([email, timestamp]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'OK' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
