// AMBER one-time self-service admin setup patch.
// In the live Apps Script project:
// 1) Replace the existing doPost(e) with the doPost(e) below.
// 2) Add the helper functions below anywhere in Code.gs.
// 3) Save and deploy a NEW VERSION of the existing Web App.

function getSettingValue_(key) {
  const rows = table(T.SET, ['key','value']);
  const hit = rows.find(r => String(r.key) === String(key));
  return hit ? String(hit.value || '') : '';
}

function setSettingValue_(key, value) {
  const s = sh(T.SET);
  const rows = table(T.SET, ['key','value']);
  const hit = rows.find(r => String(r.key) === String(key));
  if (hit) {
    s.getRange(hit._row, 2).setValue(String(value));
  } else {
    put(T.SET, [[String(key), String(value)]]);
  }
}

function bootstrapAdmin_(b) {
  if (getSettingValue_('admin_setup_enabled').toUpperCase() !== 'TRUE') {
    throw fail('Admin setup is disabled.');
  }

  const expected = getSettingValue_('admin_setup_code');
  if (!expected || String(b.setup_code || '').trim() !== expected) {
    throw fail('Invalid setup code.');
  }

  const existingAdmin = table(T.USERS, UH).find(
    u => String(u.role || '').toLowerCase() === 'admin' && isTrue(u.active)
  );
  if (existingAdmin) {
    throw fail('An administrator already exists.');
  }

  const fullName = String(b.full_name || '').trim();
  const username = String(b.username || '').trim().toLowerCase();
  const password = String(b.password || '');

  if (!fullName) throw fail('Full name is required.');
  if (!/^[a-z0-9._-]{3,40}$/.test(username)) throw fail('Invalid username.');
  if (password.length < 8) throw fail('Password must be at least 8 characters.');
  if (getUser(username)) throw fail('That username already exists.');

  put(T.USERS, [[
    username,
    fullName,
    'admin',
    'TRUE',
    makeHash(password),
    'FALSE',
    new Date().toISOString()
  ]]);

  setSettingValue_('admin_setup_enabled', 'FALSE');
  setSettingValue_('admin_setup_code', '');
  audit(username, 'admin_self_setup', 'Initial administrator created');

  return { ok: true, username: username };
}

function doPost(e) {
  let out;
  const isBridge = !!(e && e.parameter && e.parameter.payload);
  const nonce = e && e.parameter ? String(e.parameter.bridge_nonce || '') : '';

  try {
    const raw = isBridge
      ? e.parameter.payload
      : (e && e.postData ? e.postData.contents : '{}');

    const body = JSON.parse(raw);
    out = {
      data: body && body.action === 'bootstrap_admin'
        ? bootstrapAdmin_(body)
        : route(body)
    };
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
