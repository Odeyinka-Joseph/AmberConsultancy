// FINAL AMBER GitHub Pages -> Google Apps Script bridge
// Replace ONLY the existing doPost(e) function in the live Apps Script project
// with this function, save, then deploy a NEW VERSION of the existing Web App.

function doPost(e) {
  let out;
  const isBridge = !!(e && e.parameter && e.parameter.payload);
  const nonce = e && e.parameter ? String(e.parameter.bridge_nonce || '') : '';

  try {
    const raw = isBridge
      ? e.parameter.payload
      : (e && e.postData ? e.postData.contents : '{}');

    out = { data: route(JSON.parse(raw)) };
  } catch (err) {
    out = {
      error: err && err.message ? err.message : String(err),
      code: err && err.code ? err.code : ''
    };
  }

  if (isBridge) {
    const message = JSON.stringify({
      __amberBridge: true,
      nonce: nonce,
      response: out
    }).replace(/</g, '\\u003c');

    // IMPORTANT: emit a real closing </script> tag. The previous full backend
    // accidentally emitted <\\/script>, which leaves the relay script malformed
    // in the browser and causes every login request to time out.
    const html = '<!doctype html><html><body><script>' +
      'window.top.postMessage(' + message + ', "*");' +
      '</script></body></html>';

    return HtmlService
      .createHtmlOutput(html)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
