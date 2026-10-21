// AMBER GitHub Pages -> Google Apps Script bridge
// Replace the existing doPost(e) function in your Apps Script project with this one,
// then deploy a NEW VERSION of the existing Web App deployment.
// Keep the same /exec URL and set access to Anyone.

function doPost(e) {
  let out;
  const isBridge = !!(e && e.parameter && e.parameter.payload);
  const nonce = e && e.parameter ? String(e.parameter.bridge_nonce || '') : '';

  try {
    const raw = isBridge ? e.parameter.payload : (e && e.postData ? e.postData.contents : '{}');
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

    // ALLOWALL is required because the response is loaded inside a hidden iframe
    // on the GitHub Pages domain. Without this, the browser can block the frame
    // and the frontend eventually reports "Failed to fetch" / no response.
    return HtmlService.createHtmlOutput(
      '<!doctype html><html><body><script>' +
      'window.parent.postMessage(' + message + ', "*");' +
      '<\/script></body></html>'
    ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
