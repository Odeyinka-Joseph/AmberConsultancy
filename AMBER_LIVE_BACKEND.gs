// AMBER Consulting — Kwara 2027 Phone Poll backend
// Google Apps Script web app backend.
// Paste this entire file into Code.gs, save, then deploy as Web app:
// Execute as: Me | Who has access: Anyone

const SPREADSHEET_ID = '1PIEcKV59dS_Fpz-C9iew1RoAIxsXvny78DpfNfMKrC4';

const T = {
  USERS: 'Users',
  CONTACTS: 'Contacts',
  RESP: 'Responses',
  ATT: 'Attempts',
  SET: 'Settings',
  AUDIT: 'Audit',
  STAGE: 'Staging',
  ERR: 'Import errors',
  UP: 'Uploads'
};

const UH = ['username','full_name','role','active','pw_hash','must_change','created_at'];
const CH = ['id','name','phone_normalised','phone_original','lga','ward','polling_unit','status','in_directory','created_at'];
const AH = ['id','contact_id','enumerator','started_at','ended_at','disposition'];

const RCOLS = [
  ['Response ID','id'],['Submitted at','submitted_at'],['Interview date','interview_date'],
  ['Enumerator','_enumerator'],['Enumerator username','_enum_user'],['Status','status'],
  ['Disposition','disposition'],['Call start','call_start'],['Call end','call_end'],
  ['Duration (seconds)','duration_seconds'],['Contact ID','contact_id'],
  ['Language of interview','language'],['Respondent name','respondent_name'],
  ['Phone number','respondent_phone'],['Local government','respondent_lga'],
  ['Ward','respondent_ward'],['Polling unit','respondent_polling_unit'],
  ['Respondent reached','respondent_reached'],['Consent','consent'],
  ['S1 Registration check','s1_registration_check'],['S1 New local government','s1_new_lga'],
  ['S1 New ward','s1_new_ward'],['Voting local government','vote_lga'],
  ['Senate district','senate_district'],['Q1 Rating of Governor AbdulRazaq','q1_governor_rating'],
  ['Q2 Biggest problem in area','q2_biggest_problem'],['Q2 Other','q2_other'],
  ['Q3 Governorship vote','q3_governor_vote'],['Q3 Someone else','q3_other'],
  ['Q3 Rotation','q3_rotation'],['Q4 Mind made up','q4_mind_made_up'],
  ['Q5 Leaning','q5_leaning'],['Q6 Main reason','q6_main_reason'],
  ['Q6 Which leader','q6_leader'],['Q6 Other','q6_other'],
  ['Q7 Heard of Danladi (APC)','q7_known_danladi'],['Q7 Heard of Kawu (PDP)','q7_known_kawu'],
  ['Q7 Heard of Zakari (ADC)','q7_known_zakari'],['Q7 Heard of Apaokagi (SDP)','q7_known_apaokagi'],
  ['Q8 Feeling Danladi (APC)','q8_feeling_danladi'],['Q8 Feeling Kawu (PDP)','q8_feeling_kawu'],
  ['Q8 Feeling Zakari (ADC)','q8_feeling_zakari'],['Q8 Feeling Apaokagi (SDP)','q8_feeling_apaokagi'],
  ['Q9 Senate vote','q9_senate_vote'],['Q9 Someone else','q9_other'],['Q9 Rotation','q9_rotation'],
  ['Q10 Same party in both','q10_same_party'],['Q11 Zoning','q11_zoning'],
  ['Q12 2023 Governor vote','q12_2023_vote'],['Q13 Turnout likelihood','q13_turnout_likelihood'],
  ['Q14 Voter card (PVC)','q14_pvc'],['D1 Religion','d1_religion'],
  ['D2 People or tribe','d2_tribe'],['D2 Other','d2_other']
];

const LGAS = {
  Asa:'Kwara Central', 'Ilorin East':'Kwara Central', 'Ilorin South':'Kwara Central',
  'Ilorin West':'Kwara Central', Offa:'Kwara South', Oyun:'Kwara South',
  Ifelodun:'Kwara South', Irepodun:'Kwara South', Isin:'Kwara South',
  Ekiti:'Kwara South', 'Oke-Ero':'Kwara South', Moro:'Kwara North',
  Baruten:'Kwara North', Edu:'Kwara North', Kaiama:'Kwara North', Patigi:'Kwara North'
};

function db_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function sh_(name) { return db_().getSheetByName(name); }
function props_() { return PropertiesService.getScriptProperties(); }
function isTrue_(v) { return String(v).toUpperCase() === 'TRUE'; }
function fail_(message, code) { const e = new Error(message); e.code = code || ''; return e; }

function ensureSecret_() {
  let value = props_().getProperty('SECRET');
  if (!value) {
    value = Utilities.getUuid() + Utilities.getUuid();
    props_().setProperty('SECRET', value);
  }
  return value;
}

function table_(sheetName, headers) {
  const s = sh_(sheetName);
  if (!s) throw fail_('Backend sheet missing: ' + sheetName, 'backend');
  const lastRow = s.getLastRow();
  if (lastRow < 2) return [];
  return s.getRange(2, 1, lastRow - 1, headers.length).getValues().map((row, i) => {
    const out = {_row: i + 2};
    headers.forEach((h, j) => out[h] = row[j]);
    return out;
  });
}

function appendRows_(sheetName, rows) {
  if (!rows || !rows.length) return;
  const s = sh_(sheetName);
  if (!s) throw fail_('Backend sheet missing: ' + sheetName, 'backend');
  const start = s.getLastRow() + 1;
  s.getRange(start, 1, rows.length, rows[0].length).setNumberFormat('@').setValues(rows);
}

function audit_(who, action, details) {
  try {
    appendRows_(T.AUDIT, [[new Date().toISOString(), String(who || ''), String(action || ''), String(details || '')]]);
  } catch (e) {}
}

function getSetting_(key) {
  const row = table_(T.SET, ['key','value']).find(r => String(r.key) === String(key));
  return row ? String(row.value || '') : '';
}

function setSetting_(key, value) {
  const s = sh_(T.SET);
  const rows = table_(T.SET, ['key','value']);
  const row = rows.find(r => String(r.key) === String(key));
  if (row) s.getRange(row._row, 2).setValue(String(value));
  else appendRows_(T.SET, [[String(key), String(value)]]);
}

function nsn_(x) {
  let d = String(x || '').replace(/\D/g, '');
  if (d.indexOf('234') === 0) d = d.slice(3);
  else if (d[0] === '0') d = d.slice(1);
  return d;
}

function e164_(x) {
  const n = nsn_(x);
  return /^[789]\d{9}$/.test(n) ? '+234' + n : null;
}

function getUser_(username) {
  const wanted = String(username || '').trim().toLowerCase();
  return table_(T.USERS, UH).find(u => String(u.username || '').trim().toLowerCase() === wanted);
}

function publicUser_(u) {
  return {
    username: String(u.username || ''),
    full_name: String(u.full_name || ''),
    role: String(u.role || ''),
    must_change: isTrue_(u.must_change)
  };
}

function hashPw_(password, salt) {
  let h = salt + password;
  for (let i = 0; i < 1000; i++) {
    h = Utilities.base64Encode(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt)
    );
  }
  return h;
}

function makeHash_(password) {
  const salt = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  return salt + '$' + hashPw_(password, salt);
}

function checkPw_(password, stored) {
  const p = String(stored || '').split('$');
  return p.length === 2 && hashPw_(password, p[0]) === p[1];
}

function sign_(text) {
  return Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(text, ensureSecret_())
  );
}

function makeToken_(username) {
  const payload = Utilities.base64EncodeWebSafe(JSON.stringify({
    u: username,
    exp: Date.now() + 3 * 864e5
  }));
  return payload + '.' + sign_(payload);
}

function readToken_(token) {
  const p = String(token || '').split('.');
  if (p.length !== 2 || sign_(p[0]) !== p[1]) throw fail_('Please sign in again.', 'auth');
  const obj = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(p[0])).getDataAsString());
  if (Number(obj.exp) < Date.now()) throw fail_('Session expired. Please sign in again.', 'auth');
  return String(obj.u || '');
}

function auth_(token) {
  const u = getUser_(readToken_(token));
  if (!u || !isTrue_(u.active)) throw fail_('Account inactive. Contact your administrator.', 'auth');
  return u;
}

function login_(body) {
  const u = getUser_(body.username);
  if (!u || !isTrue_(u.active) || !checkPw_(String(body.password || ''), String(u.pw_hash || ''))) {
    throw fail_('Wrong username or password.', 'auth');
  }
  return {token: makeToken_(String(u.username)), user: publicUser_(u)};
}

function bootstrapAdmin_(body) {
  if (getSetting_('admin_setup_enabled').toUpperCase() !== 'TRUE') {
    throw fail_('Admin setup is disabled.', 'setup');
  }
  const expected = getSetting_('admin_setup_code');
  if (!expected || String(body.setup_code || '').trim() !== expected) {
    throw fail_('Invalid setup code.', 'setup');
  }
  const existingAdmin = table_(T.USERS, UH).find(
    u => String(u.role || '').toLowerCase() === 'admin' && isTrue_(u.active)
  );
  if (existingAdmin) throw fail_('An administrator already exists.', 'setup');

  const fullName = String(body.full_name || '').trim();
  const username = String(body.username || '').trim().toLowerCase();
  const password = String(body.password || '');
  if (!fullName) throw fail_('Full name is required.', 'validation');
  if (!/^[a-z0-9._-]{3,40}$/.test(username)) throw fail_('Invalid username.', 'validation');
  if (password.length < 8) throw fail_('Password must be at least 8 characters.', 'validation');
  if (getUser_(username)) throw fail_('That username already exists.', 'validation');

  appendRows_(T.USERS, [[
    username, fullName, 'admin', 'TRUE', makeHash_(password), 'FALSE', new Date().toISOString()
  ]]);
  setSetting_('admin_setup_enabled', 'FALSE');
  setSetting_('admin_setup_code', '');
  audit_(username, 'admin_self_setup', 'Initial administrator created');
  return {ok: true, username: username};
}

function requireAdmin_(u) {
  if (!u || String(u.role || '').toLowerCase() !== 'admin') {
    throw fail_('Admin access required.', 'forbidden');
  }
}

function temporaryPassword_() {
  const chars = 'abcdefghjkmnpqrstuvwxyz23456789';
  let out = '';
  for (let i = 0; i < 9; i++) out += chars[Math.floor(Math.random() * chars.length)];
  return out;
}

function adminListUsers_(u) {
  requireAdmin_(u);
  return table_(T.USERS, UH).map(x => ({
    username: String(x.username || ''),
    full_name: String(x.full_name || ''),
    role: String(x.role || ''),
    active: isTrue_(x.active),
    must_change: isTrue_(x.must_change),
    created_at: String(x.created_at || '')
  }));
}

function adminCreateUser_(u, body) {
  requireAdmin_(u);
  const fullName = String(body.full_name || '').trim();
  const username = String(body.username || '').trim().toLowerCase();
  const role = String(body.role || 'enumerator').trim().toLowerCase();
  if (!fullName) throw fail_('Full name is required.', 'validation');
  if (!/^[a-z0-9._-]{3,40}$/.test(username)) throw fail_('Invalid username.', 'validation');
  if (['admin','enumerator'].indexOf(role) < 0) throw fail_('Invalid role.', 'validation');
  if (getUser_(username)) throw fail_('That username already exists.', 'validation');
  const pw = temporaryPassword_();
  appendRows_(T.USERS, [[username, fullName, role, 'TRUE', makeHash_(pw), 'TRUE', new Date().toISOString()]]);
  audit_(u.username, 'user_created_web', username);
  return {ok: true, username: username, temporary_password: pw};
}

function adminResetPassword_(u, body) {
  requireAdmin_(u);
  const target = getUser_(body.username);
  if (!target) throw fail_('User not found.', 'not_found');
  const pw = temporaryPassword_();
  const s = sh_(T.USERS);
  s.getRange(target._row, UH.indexOf('pw_hash') + 1).setValue(makeHash_(pw));
  s.getRange(target._row, UH.indexOf('must_change') + 1).setValue('TRUE');
  audit_(u.username, 'password_reset_web', target.username);
  return {ok: true, username: String(target.username), temporary_password: pw};
}

function adminToggleUser_(u, body) {
  requireAdmin_(u);
  const target = getUser_(body.username);
  if (!target) throw fail_('User not found.', 'not_found');
  if (String(target.username).toLowerCase() === String(u.username).toLowerCase() && isTrue_(target.active)) {
    throw fail_('You cannot deactivate your own account.', 'validation');
  }
  const active = !isTrue_(target.active);
  sh_(T.USERS).getRange(target._row, UH.indexOf('active') + 1).setValue(active ? 'TRUE' : 'FALSE');
  audit_(u.username, active ? 'user_reactivated_web' : 'user_deactivated_web', target.username);
  return {ok: true, username: String(target.username), active: active};
}

function allowManual_() {
  return getSetting_('allow_not_in_directory').toUpperCase() === 'TRUE';
}

function searchContacts_(query) {
  const digits = nsn_(query);
  if (digits.length < 7) return [];
  const s = sh_(T.CONTACTS);
  if (!s || s.getLastRow() < 2) return [];
  const rows = s.getRange(2, 1, s.getLastRow() - 1, CH.length).getValues();
  const attemptSheet = sh_(T.ATT);
  const attempts = attemptSheet && attemptSheet.getLastRow() > 1
    ? attemptSheet.getRange(2, 2, attemptSheet.getLastRow() - 1, 1).getValues().flat()
    : [];

  return rows.map((r, i) => {
    const o = {_row: i + 2};
    CH.forEach((k, j) => o[k] = r[j]);
    return o;
  }).filter(o => nsn_(o.phone_normalised).indexOf(digits) >= 0)
    .slice(0, 10)
    .map(o => ({
      id: String(o.id || ''), name: String(o.name || ''),
      phone_normalised: String(o.phone_normalised || ''), lga: String(o.lga || ''),
      ward: String(o.ward || ''), polling_unit: String(o.polling_unit || ''),
      status: String(o.status || ''), in_directory: isTrue_(o.in_directory),
      attempts: attempts.filter(x => String(x) === String(o.id)).length
    }));
}

function findContact_(id) {
  return table_(T.CONTACTS, CH).find(c => String(c.id) === String(id));
}

function startInterview_(u, id) {
  const c = findContact_(id);
  if (!c) throw fail_('Contact not found.', 'not_found');
  if (String(c.status) === 'Completed' && String(u.role).toLowerCase() !== 'admin') {
    throw fail_('Interview already completed.', 'validation');
  }
  return {
    contact: {
      id: String(c.id || ''), name: String(c.name || ''), phone: String(c.phone_normalised || ''),
      lga: String(c.lga || ''), ward: String(c.ward || ''), polling_unit: String(c.polling_unit || ''),
      status: String(c.status || '')
    },
    server_now: new Date().toISOString()
  };
}

function manualContact_(u, body) {
  if (!allowManual_()) throw fail_('Interviews for numbers not in the directory are disabled.', 'forbidden');
  const phone = e164_(body.phone);
  if (!phone) throw fail_('Invalid Nigerian phone number.', 'validation');
  const fullName = String(body.name || '').trim();
  const lga = String(body.lga || '').trim();
  const ward = String(body.ward || '').trim();
  const pu = String(body.polling_unit || '').trim();
  if (!LGAS[lga] || !fullName || !ward || !pu) throw fail_('Name, LGA, ward and polling unit are required.', 'validation');
  const id = Utilities.getUuid();
  appendRows_(T.CONTACTS, [[id, fullName, phone, String(body.phone || ''), lga, ward, pu, 'Not called', 'FALSE', new Date().toISOString()]]);
  audit_(u.username, 'manual_contact', id);
  return {id: id};
}

function submitInterview_(u, row) {
  if (!row || !row.id || !row.contact_id) throw fail_('Invalid submission.', 'validation');
  const c = findContact_(row.contact_id);
  if (!c) throw fail_('Contact not found.', 'not_found');

  const now = new Date().toISOString();
  row.submitted_at = now;
  row.interview_date = now;
  row._enumerator = String(u.full_name || '');
  row._enum_user = String(u.username || '');
  row.respondent_name = String(c.name || '');
  row.respondent_phone = String(c.phone_normalised || '');
  row.respondent_lga = String(c.lga || '');
  row.respondent_ward = String(c.ward || '');
  row.respondent_polling_unit = String(c.polling_unit || '');
  row.vote_lga = String(row.s1_new_lga || c.lga || '');
  row.senate_district = LGAS[row.vote_lga] || '';

  const start = row.call_start ? new Date(row.call_start).getTime() : Date.now();
  const end = row.call_end ? new Date(row.call_end).getTime() : Date.now();
  row.duration_seconds = Math.max(0, Math.round((end - start) / 1000));

  appendRows_(T.RESP, [RCOLS.map(x => row[x[1]] == null ? '' : String(row[x[1]]))]);
  appendRows_(T.ATT, [[
    Utilities.getUuid(), String(c.id), String(u.username),
    String(row.call_start || ''), String(row.call_end || ''), String(row.disposition || 'Completed')
  ]]);

  if (String(c.status) !== 'Completed') {
    sh_(T.CONTACTS).getRange(c._row, CH.indexOf('status') + 1)
      .setValue(String(row.status) === 'complete' ? 'Completed' : String(row.disposition || 'Call back'));
  }
  audit_(u.username, 'interview_submit', row.id);
  return {ok: true};
}

function route_(body) {
  if (!body || typeof body !== 'object') throw fail_('Invalid request.', 'validation');
  const action = String(body.action || '');
  if (action === 'health') return {ok: true, service: 'AMBER API', version: '2026-10-21'};
  if (action === 'bootstrap_admin') return bootstrapAdmin_(body);
  if (action === 'login') return login_(body);

  const u = auth_(body.token);
  if (isTrue_(u.must_change) && ['me','change_password'].indexOf(action) < 0) {
    throw fail_('Change your password first.', 'auth');
  }

  switch (action) {
    case 'me': return publicUser_(u);
    case 'change_password': {
      const pw = String(body.new_password || '');
      if (pw.length < 8) throw fail_('Use at least 8 characters.', 'validation');
      const s = sh_(T.USERS);
      s.getRange(u._row, UH.indexOf('pw_hash') + 1).setValue(makeHash_(pw));
      s.getRange(u._row, UH.indexOf('must_change') + 1).setValue('FALSE');
      audit_(u.username, 'password_changed', '');
      return {ok: true};
    }
    case 'settings': return {allow_not_in_directory: allowManual_()};
    case 'admin_list_users': return adminListUsers_(u);
    case 'admin_create_user': return adminCreateUser_(u, body);
    case 'admin_reset_password': return adminResetPassword_(u, body);
    case 'admin_toggle_user': return adminToggleUser_(u, body);
    case 'search': return searchContacts_(body.q);
    case 'start': return startInterview_(u, body.contact_id);
    case 'submit': return submitInterview_(u, body.row);
    case 'manual_contact': return manualContact_(u, body);
    default: throw fail_('Unknown action.', 'validation');
  }
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ok: true, service: 'AMBER API', version: '2026-10-21'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let out;
  const isBridge = !!(e && e.parameter && e.parameter.payload);
  const nonce = e && e.parameter ? String(e.parameter.bridge_nonce || '') : '';

  try {
    const raw = isBridge
      ? String(e.parameter.payload || '{}')
      : (e && e.postData ? String(e.postData.contents || '{}') : '{}');
    out = {data: route_(JSON.parse(raw))};
  } catch (err) {
    out = {
      error: err && err.message ? String(err.message) : String(err),
      code: err && err.code ? String(err.code) : ''
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

    return HtmlService.createHtmlOutput(html)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function setup() {
  const defs = [
    [T.USERS, UH], [T.CONTACTS, CH], [T.RESP, RCOLS.map(x => x[0])], [T.ATT, AH],
    [T.SET, ['key','value']], [T.AUDIT, ['time','who','action','details']],
    [T.STAGE, ['Name','Phone number','LGA','Ward','Polling unit']],
    [T.ERR, ['Row in Staging','Name','Phone number','LGA','Reason']],
    [T.UP, ['time','who','mode','rows imported','rows rejected']]
  ];

  defs.forEach(def => {
    let s = sh_(def[0]);
    if (!s) s = db_().insertSheet(def[0]);
    if (s.getMaxColumns() < def[1].length) {
      s.insertColumnsAfter(s.getMaxColumns(), def[1].length - s.getMaxColumns());
    }
    if (s.getLastRow() === 0) {
      s.getRange(1, 1, 1, def[1].length).setValues([def[1]]).setFontWeight('bold');
    }
  });

  ensureSecret_();
  if (!getSetting_('allow_not_in_directory')) setSetting_('allow_not_in_directory', 'FALSE');
  if (!getSetting_('admin_setup_enabled')) setSetting_('admin_setup_enabled', 'FALSE');
  return 'AMBER backend setup complete';
}

function enableAdminSetup() {
  const existingAdmin = table_(T.USERS, UH).find(
    u => String(u.role || '').toLowerCase() === 'admin' && isTrue_(u.active)
  );
  if (existingAdmin) throw new Error('An active administrator already exists.');
  const code = 'AMBER-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
  setSetting_('admin_setup_code', code);
  setSetting_('admin_setup_enabled', 'TRUE');
  SpreadsheetApp.getUi().alert('Admin setup enabled.\nOne-time setup code: ' + code + '\nKeep this code private.');
  return code;
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('AMBER')
      .addItem('Initialize / repair backend', 'setup')
      .addItem('Enable first-admin setup', 'enableAdminSetup')
      .addToUi();
  } catch (e) {}
}
