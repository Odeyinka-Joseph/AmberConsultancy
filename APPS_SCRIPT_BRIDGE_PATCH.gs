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
    out = { error: err.message, code: err.code || '' };
  }

  // Browser requests from GitHub Pages arrive through a hidden form/iframe.
  // Return HTML that posts the JSON response safely to the parent page.
  if (isBridge) {
    const message = JSON.stringify({
      __amberBridge: true,
      nonce: nonce,
      response: out
    }).replace(/</g, '\\u003c');

    return HtmlService.createHtmlOutput(
      '<!doctype html><html><body><script>' +
      'window.parent.postMessage(' + message + ', "*");' +
      '</script></body></html>'
    );
  }

  // Retain normal JSON behavior for non-browser/API callers.
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
