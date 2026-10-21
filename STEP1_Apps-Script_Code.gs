// AMBER Kwara 2027 Phone Poll: backend on Google Sheets + Apps Script.
// Paste into Extensions -> Apps Script of the Google Sheet, then run AMBER menu -> "1. Set up sheets".
// Deploy -> New deployment -> Web app -> Execute as: Me, Who has access: Anyone -> copy the /exec URL.
const T = { USERS: 'Users', CONTACTS: 'Contacts', RESP: 'Responses', ATT: 'Attempts', SET: 'Settings', AUDIT: 'Audit', STAGE: 'Staging', ERR: 'Import errors', UP: 'Uploads' };
const LGAS = { Asa:'Kwara Central','Ilorin East':'Kwara Central','Ilorin South':'Kwara Central','Ilorin West':'Kwara Central',
  Offa:'Kwara South',Oyun:'Kwara South',Ifelodun:'Kwara South',Irepodun:'Kwara South',Isin:'Kwara South',Ekiti:'Kwara South','Oke-Ero':'Kwara South',
  Moro:'Kwara North',Baruten:'Kwara North',Edu:'Kwara North',Kaiama:'Kwara North',Patigi:'Kwara North' };
const UH = ['username','full_name','role','active','pw_hash','must_change','created_at'];
const CH = ['id','name','phone_normalised','phone_original','lga','ward','polling_unit','status','in_directory','created_at'];
const AH = ['id','contact_id','enumerator','started_at','ended_at','disposition'];
// [Responses column header, key]. Wordings only; human-readable headers; no codes.
const RCOLS = [['Response ID','id'],['Submitted at','submitted_at'],['Interview date','interview_date'],['Enumerator','_enumerator'],['Enumerator username','_enum_user'],
  ['Status','status'],['Disposition','disposition'],['Call start','call_start'],['Call end','call_end'],['Duration (seconds)','duration_seconds'],['Contact ID','contact_id'],
  ['Language of interview','language'],['Respondent name','respondent_name'],['Phone number','respondent_phone'],['Local government','respondent_lga'],['Ward','respondent_ward'],['Polling unit','respondent_polling_unit'],
  ['Respondent reached','respondent_reached'],['Consent','consent'],['S1 Registration check','s1_registration_check'],['S1 New local government','s1_new_lga'],['S1 New ward','s1_new_ward'],
  ['Voting local government','vote_lga'],['Senate district','senate_district'],['Q1 Rating of Governor AbdulRazaq','q1_governor_rating'],['Q2 Biggest problem in area','q2_biggest_problem'],['Q2 Other','q2_other'],
  ['Q3 Governorship vote','q3_governor_vote'],['Q3 Someone else','q3_other'],['Q3 Rotation','q3_rotation'],['Q4 Mind made up','q4_mind_made_up'],['Q5 Leaning','q5_leaning'],
  ['Q6 Main reason','q6_main_reason'],['Q6 Which leader','q6_leader'],['Q6 Other','q6_other'],
  ['Q7 Heard of Danladi (APC)','q7_known_danladi'],['Q7 Heard of Kawu (PDP)','q7_known_kawu'],['Q7 Heard of Zakari (ADC)','q7_known_zakari'],['Q7 Heard of Apaokagi (SDP)','q7_known_apaokagi'],
  ['Q8 Feeling Danladi (APC)','q8_feeling_danladi'],['Q8 Feeling Kawu (PDP)','q8_feeling_kawu'],['Q8 Feeling Zakari (ADC)','q8_feeling_zakari'],['Q8 Feeling Apaokagi (SDP)','q8_feeling_apaokagi'],
  ['Q9 Senate vote','q9_senate_vote'],['Q9 Someone else','q9_other'],['Q9 Rotation','q9_rotation'],['Q10 Same party in both','q10_same_party'],['Q11 Zoning','q11_zoning'],
  ['Q12 2023 Governor vote','q12_2023_vote'],['Q13 Turnout likelihood','q13_turnout_likelihood'],['Q14 Voter card (PVC)','q14_pvc'],['D1 Religion','d1_religion'],['D2 People or tribe','d2_tribe'],['D2 Other','d2_other']];

// ---------------------------------------------------------------- helpers
const sh = n => SpreadsheetApp.getActive().getSheetByName(n);
const props = () => PropertiesService.getScriptProperties();
const isTrue = v => String(v).toUpperCase() === 'TRUE';
function fail(msg, code) { const e = new Error(msg); e.code = code || ''; return e; }
function table(name, headers) {
  const s = sh(name), n = s.getLastRow(); if (n < 2) return [];
  return s.getRange(2, 1, n - 1, headers.length).getValues().map((r, i) => { const o = { _row: i + 2 }; headers.forEach((k, j) => o[k] = r[j]); return o; });
}
function put(name, values) { const s = sh(name), r = s.getLastRow() + 1; s.getRange(r, 1, values.length ? 1 : 0, values[0].length).setNumberFormat('@').setValues(values); return r; }
function audit(who, action, details) { put(T.AUDIT, [[new Date().toISOString(), who, action, details || '']]); }
const nsn = s => { let d = String(s || '').replace(/\D/g, ''); if (d.indexOf('234') === 0) d = d.slice(3); else if (d[0] === '0') d = d.slice(1); return d; };
const e164 = s => { const n = nsn(s); return /^[789]\d{9}$/.test(n) ? '+234' + n : null; };
const lagos = (v, f) => Utilities.formatDate(new Date(v), 'Africa/Lagos', f || 'dd/MM/yyyy HH:mm:ss');

// ---------------------------------------------------------------- auth
const sign = s => Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, props().getProperty('SECRET')));
function makeToken(u) { const p = Utilities.base64EncodeWebSafe(JSON.stringify({ u: u, exp: Date.now() + 3 * 864e5 })); return p + '.' + sign(p); }
function readToken(t) {
  const parts = String(t || '').split('.'); if (parts.length !== 2 || sign(parts[0]) !== parts[1]) throw fail('Please sign in again.', 'auth');
  const o = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString());
  if (o.exp < Date.now()) throw fail('Session expired. Please sign in again.', 'auth'); return o.u;
}
function hashPw(pw, salt) { let h = salt + pw; for (let i = 0; i < 1000; i++) h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt)); return h; }
function makeHash(pw) { const salt = Utilities.getUuid().replace(/-/g, '').slice(0, 16); return salt + '$' + hashPw(pw, salt); }
function checkPw(pw, stored) { const p = String(stored).split('$'); return p.length === 2 && hashPw(pw, p[0]) === p[1]; }
const getUser = name => table(T.USERS, UH).filter(u => String(u.username).toLowerCase() === String(name || '').trim().toLowerCase())[0];
const pub = u => ({ username: u.username, full_name: u.full_name, role: u.role, must_change: isTrue(u.must_change) });
function auth(token) { const u = getUser(readToken(token)); if (!u || !isTrue(u.active)) throw fail('Account inactive. Contact your admin.', 'auth'); return u; }
function login(b) {
  const cache = CacheService.getScriptCache(), k = 'fail:' + String(b.username).toLowerCase(), n = Number(cache.get(k) || 0);
  if (n >= 5) throw fail('Too many attempts. Try again in 10 minutes.');
  const u = getUser(b.username);
  if (!u || !isTrue(u.active) || !checkPw(String(b.password || ''), u.pw_hash)) { cache.put(k, String(n + 1), 600); throw fail('Wrong username or password.'); }
  cache.remove(k); return { token: makeToken(u.username), user: pub(u) };
}

// ---------------------------------------------------------------- API
function doGet() { return ContentService.createTextOutput('AMBER API running'); }
function doPost(e) {
  let out;
  try { out = { data: route(JSON.parse(e.postData.contents)) }; } catch (err) { out = { error: err.message, code: err.code || '' }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function route(b) {
  if (b.action === 'login') return login(b);
  const u = auth(b.token);
  if (isTrue(u.must_change) && ['me', 'change_password'].indexOf(b.action) < 0) throw fail('Change your password first.');
  switch (b.action) {
    case 'me': return pub(u);
    case 'change_password':
      if (String(b.new_password || '').length < 8) throw fail('Use at least 8 characters.');
      sh(T.USERS).getRange(u._row, UH.indexOf('pw_hash') + 1).setValue(makeHash(b.new_password));
      sh(T.USERS).getRange(u._row, UH.indexOf('must_change') + 1).setValue('FALSE'); return { ok: true };
    case 'settings': return { allow_not_in_directory: allowManual() };
    case 'admin_list_users': return adminListUsers(u);
    case 'admin_create_user': return adminCreateUser(u, b);
    case 'admin_reset_password': return adminResetPassword(u, b);
    case 'admin_toggle_user': return adminToggleUser(u, b);
    case 'search': return search(b.q);
    case 'start': return start(u, b.contact_id);
    case 'submit': return submit(u, b.row);
    case 'manual_contact': return manualContact(u, b);
    default: throw fail('Unknown action');
  }
}
function requireAdmin(u) { if (!u || String(u.role).toLowerCase() !== 'admin') throw fail('Admin access required.', 'forbidden'); }
function adminListUsers(u) {
  requireAdmin(u);
  return table(T.USERS, UH).map(x => ({ username: String(x.username), full_name: String(x.full_name), role: String(x.role), active: isTrue(x.active), must_change: isTrue(x.must_change), created_at: String(x.created_at || '') }));
}
function adminCreateUser(u, b) {
  requireAdmin(u);
  const name = String(b.full_name || '').trim(), un = String(b.username || '').trim().toLowerCase(), role = String(b.role || 'enumerator').toLowerCase();
  if (!name) throw fail('Full name is required.');
  if (!/^[a-z0-9._-]{3,40}$/.test(un)) throw fail('Username must be 3-40 letters/numbers and may include . _ -');
  if (['admin','enumerator'].indexOf(role) < 0) throw fail('Role must be admin or enumerator.');
  if (getUser(un)) throw fail('That username already exists.');
  const pw = tempPw();
  put(T.USERS, [[un, name, role, 'TRUE', makeHash(pw), 'TRUE', new Date().toISOString()]]);
  audit(u.username, 'user_created_web', un + ' (' + role + ')');
  return { ok:true, username:un, temporary_password:pw };
}
function adminResetPassword(u, b) {
  requireAdmin(u);
  const target = getUser(b.username); if (!target) throw fail('User not found.');
  const pw = tempPw();
  sh(T.USERS).getRange(target._row, 5).setValue(makeHash(pw));
  sh(T.USERS).getRange(target._row, 6).setValue('TRUE');
  audit(u.username, 'password_reset_web', target.username);
  return { ok:true, username:target.username, temporary_password:pw };
}
function adminToggleUser(u, b) {
  requireAdmin(u);
  const target = getUser(b.username); if (!target) throw fail('User not found.');
  if (String(target.username).toLowerCase() === String(u.username).toLowerCase() && isTrue(target.active)) throw fail('You cannot deactivate your own signed-in admin account.');
  const on = !isTrue(target.active);
  sh(T.USERS).getRange(target._row, UH.indexOf('active') + 1).setValue(on ? 'TRUE' : 'FALSE');
  audit(u.username, on ? 'user_reactivated_web' : 'user_deactivated_web', target.username);
  return { ok:true, username:target.username, active:on };
}
const allowManual = () => table(T.SET, ['key', 'value']).some(r => r.key === 'allow_not_in_directory' && isTrue(r.value));
const asContact = o => ({ id: o.id, name: o.name, phone_normalised: o.phone_normalised, lga: o.lga, ward: o.ward, polling_unit: o.polling_unit, status: o.status, in_directory: isTrue(o.in_directory) });

function search(q) {
  const d = nsn(q), s = sh(T.CONTACTS), n = s.getLastRow(); if (d.length < 7 || n < 2) return [];
  const phones = s.getRange(2, 3, n - 1, 1).getValues(), hits = [];
  for (let i = 0; i < phones.length && hits.length < 10; i++) if (String(phones[i][0]).indexOf(d) >= 0) hits.push(i + 2);
  const att = sh(T.ATT).getLastRow() > 1 ? sh(T.ATT).getRange(2, 2, sh(T.ATT).getLastRow() - 1, 1).getValues().flat() : [];
  return hits.map(r => { const v = s.getRange(r, 1, 1, CH.length).getValues()[0], o = {}; CH.forEach((k, j) => o[k] = v[j]);
    return Object.assign(asContact(o), { attempts: att.filter(x => x === o.id).length }); }).sort((a, b) => a.name < b.name ? -1 : 1);
}
const findContact = id => table(T.CONTACTS, CH).filter(c => c.id === id)[0];
function start(u, id) {
  const c = findContact(id); if (!c) throw fail('Contact not found');
  if (c.status === 'Completed' && u.role !== 'admin') throw fail('Interview already completed');
  return { contact: { id: c.id, name: c.name, phone: c.phone_normalised, lga: c.lga, ward: c.ward, polling_unit: c.polling_unit, status: c.status }, server_now: new Date().toISOString() };
}
function manualContact(u, b) {
  if (!allowManual()) throw fail('Interviews for numbers not in the directory are disabled.');
  const p = e164(b.phone); if (!p) throw fail('Invalid Nigerian phone number');
  if (!LGAS[b.lga] || !String(b.name).trim() || !String(b.ward).trim() || !String(b.polling_unit).trim()) throw fail('Name, LGA, ward and polling unit are required');
  const id = Utilities.getUuid();
  put(T.CONTACTS, [[id, String(b.name).trim(), p, b.phone, b.lga, String(b.ward).trim(), String(b.polling_unit).trim(), 'Not called', 'FALSE', new Date().toISOString()]]);
  audit(u.username, 'manual_contact', id); return { id: id };
}
const STATUS = { 'Completed': 'Completed', 'Refused': 'Refused', 'Not available': 'Not reachable', 'Wrong number': 'Wrong number', 'Ineligible': 'Ineligible' };
function submit(u, r) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    if (!r || !r.id || ['complete', 'partial', 'terminated'].indexOf(r.status) < 0) throw fail('Invalid submission');
    const rs = sh(T.RESP), n = rs.getLastRow();
    if (n > 1 && rs.getRange(2, 1, n - 1, 1).getValues().some(x => x[0] === r.id)) return { ok: true, duplicate: true };
    const c = findContact(r.contact_id); if (!c) throw fail('Contact not found');
    if (r.status === 'complete' && c.status === 'Completed' && u.role !== 'admin') throw fail('Interview already completed for this contact');
    const rec = Object.assign({}, r);
    rec.respondent_name = c.name; rec.respondent_phone = c.phone_normalised; rec.respondent_lga = c.lga; rec.respondent_ward = c.ward; rec.respondent_polling_unit = c.polling_unit;
    rec.vote_lga = r.s1_registration_check === 'I now vote somewhere else in Kwara' ? r.s1_new_lga : c.lga;
    if (!LGAS[rec.vote_lga]) throw fail('Invalid voting local government');
    rec.senate_district = LGAS[rec.vote_lga];
    rec.duration_seconds = Math.max(0, Math.round((new Date(r.call_end) - new Date(r.call_start)) / 1000));
    rec.interview_date = lagos(r.call_start, 'yyyy-MM-dd'); rec.submitted_at = new Date().toISOString();
    rec._enumerator = u.full_name; rec._enum_user = u.username;
    const skipped = {}; (r.skipped_fields || []).forEach(k => skipped[k] = 1);
    const row = RCOLS.map(function (c2) { const k = c2[1], v = rec[k];
      if (skipped[k]) return 'Skipped by logic';
      if (v === undefined || v === null) return '';
      return ['submitted_at', 'call_start', 'call_end'].indexOf(k) >= 0 ? lagos(v) : String(v); });
    put(T.RESP, [row]);
    const disp = r.status === 'complete' ? 'Completed' : (r.disposition || 'Call back');
    put(T.ATT, [[Utilities.getUuid(), c.id, u.username, lagos(r.call_start), lagos(r.call_end), disp]]);
    let st = STATUS[disp] || 'Call back'; if (r.s1_registration_check === 'I no longer vote in Kwara') st = 'Ineligible';
    if (c.status !== 'Completed') sh(T.CONTACTS).getRange(c._row, CH.indexOf('status') + 1).setValue(st);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

// ---------------------------------------------------------------- admin menu (run from the Sheet)
function onOpen() {
  SpreadsheetApp.getUi().createMenu('AMBER').addItem('1. Set up sheets', 'setup').addItem('Create admin login (username: admin)', 'ensureAdmin').addSeparator().addItem('Add user', 'addUser')
    .addItem('Reset a user\'s password', 'resetPassword').addItem('Deactivate / reactivate user', 'toggleUser').addSeparator()
    .addItem('Import directory from Staging tab', 'importDirectory').addItem('Toggle: allow numbers not in directory', 'toggleManual').addToUi();
}
const ui = () => SpreadsheetApp.getUi();
const ask = m => { const r = ui().prompt(m); return r.getSelectedButton() === ui().Button.OK ? r.getResponseText().trim() : null; };
const tempPw = () => { const a = 'abcdefghjkmnpqrstuvwxyz23456789'; let s = ''; for (let i = 0; i < 9; i++) s += a[Math.floor(Math.random() * a.length)]; return s; };
const me = () => Session.getActiveUser().getEmail() || 'sheet-owner';
function setup() {
  const defs = [[T.USERS, UH], [T.CONTACTS, CH], [T.RESP, RCOLS.map(c => c[0])], [T.ATT, AH], [T.SET, ['key', 'value']], [T.AUDIT, ['time', 'who', 'action', 'details']],
    [T.STAGE, ['Name', 'Phone number', 'LGA', 'Ward', 'Polling unit']], [T.ERR, ['Row in Staging', 'Name', 'Phone number', 'LGA', 'Reason']], [T.UP, ['time', 'who', 'mode', 'rows imported', 'rows rejected']]];
  defs.forEach(function (d) { const s = sh(d[0]) || SpreadsheetApp.getActive().insertSheet(d[0]);
    if (s.getLastRow() === 0) { s.getRange(1, 1, 1, d[1].length).setNumberFormat('@').setValues([d[1]]).setFontWeight('bold'); s.setFrozenRows(1); }
    s.getRange(2, 1, 2000, d[1].length).setNumberFormat('@'); });
  if (!props().getProperty('SECRET')) props().setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  if (table(T.SET, ['key', 'value']).length === 0) put(T.SET, [['allow_not_in_directory', 'FALSE']]);
  ui().alert('Sheets are ready. Now run AMBER -> Create admin login (username: admin).\nThen Deploy -> Manage deployments -> Edit -> New version -> Deploy.');
}
function ensureAdmin() {
  const existing = getUser('admin');
  if (existing) {
    sh(T.USERS).getRange(existing._row, UH.indexOf('role') + 1).setValue('admin');
    sh(T.USERS).getRange(existing._row, UH.indexOf('active') + 1).setValue('TRUE');
    ui().alert('Admin login already exists.\nUsername: admin\nUse AMBER -> Reset a user\'s password if you need a new temporary password.');
    return;
  }
  const pw = tempPw();
  put(T.USERS, [['admin', 'Administrator', 'admin', 'TRUE', makeHash(pw), 'TRUE', new Date().toISOString()]]);
  audit(me(), 'admin_bootstrap', 'admin');
  ui().alert('Admin login created.\nUsername: admin\nTemporary password: ' + pw + '\nYou must change it at first login. Copy it now.');
}
function addUser() {
  const name = ask('Full name:'); if (!name) return; const un = ask('Username (letters/numbers, no spaces):'); if (!un) return;
  const role = (ask('Role: type admin or enumerator') || '').toLowerCase(); if (['admin', 'enumerator'].indexOf(role) < 0) return ui().alert('Role must be admin or enumerator.');
  if (getUser(un)) return ui().alert('That username already exists.');
  const pw = tempPw(); put(T.USERS, [[un.toLowerCase(), name, role, 'TRUE', makeHash(pw), 'TRUE', new Date().toISOString()]]);
  audit(me(), 'user_created', un + ' (' + role + ')'); ui().alert('User created.\nUsername: ' + un.toLowerCase() + '\nTemporary password: ' + pw + '\nThey must change it at first login. Copy it now.');
}
function resetPassword() {
  const u = getUser(ask('Username:')); if (!u) return ui().alert('User not found.'); const pw = tempPw();
  sh(T.USERS).getRange(u._row, 5).setValue(makeHash(pw)); sh(T.USERS).getRange(u._row, 6).setValue('TRUE');
  audit(me(), 'password_reset', u.username); ui().alert('New temporary password for ' + u.username + ': ' + pw);
}
function toggleUser() {
  const u = getUser(ask('Username:')); if (!u) return ui().alert('User not found.'); const on = !isTrue(u.active);
  sh(T.USERS).getRange(u._row, 4).setValue(on ? 'TRUE' : 'FALSE'); audit(me(), on ? 'user_reactivated' : 'user_deactivated', u.username);
  ui().alert(u.username + (on ? ' is now ACTIVE.' : ' is now DEACTIVATED (cannot sign in).'));
}
function toggleManual() {
  const s = sh(T.SET), rows = table(T.SET, ['key', 'value']), r = rows.filter(x => x.key === 'allow_not_in_directory')[0], on = !isTrue(r.value);
  s.getRange(r._row, 2).setValue(on ? 'TRUE' : 'FALSE'); audit(me(), 'allow_not_in_directory', String(on)); ui().alert('Allow numbers not in directory: ' + (on ? 'ON' : 'OFF'));
}
function lev(a, b) { const m = []; for (let i = 0; i <= a.length; i++) { m[i] = [i]; } for (let j = 1; j <= b.length; j++) m[0][j] = j;
  for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++) m[i][j] = Math.min(m[i - 1][j] + 1, m[i][j - 1] + 1, m[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)); return m[a.length][b.length]; }
function matchLga(x) {
  const t = String(x || '').trim().toLowerCase(); if (!t) return null; let best = null, bd = 3;
  for (const n in LGAS) { if (n.toLowerCase() === t) return { lga: n, fixed: false }; const d = lev(t, n.toLowerCase()); if (d < bd) { bd = d; best = n; } }
  return best ? { lga: best, fixed: true } : null;
}
function importDirectory() {
  const mode = ui().alert('Import mode', 'YES = Append to the existing directory.\nNO = Replace (removes contacts not yet called, then imports).\nCancel = stop.', ui().ButtonSet.YES_NO_CANCEL);
  if (mode === ui().Button.CANCEL) return;
  const st = sh(T.STAGE), n = st.getLastRow(); if (n < 2) return ui().alert('The Staging tab is empty. Paste your directory there (columns: Name, Phone number, LGA, Ward, Polling unit).');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const replace = mode === ui().Button.NO, all = table(T.CONTACTS, CH), keep = replace ? all.filter(c => c.status !== 'Not called') : all;
    const seen = {}; keep.forEach(c => seen[c.phone_normalised + '|' + String(c.name).toLowerCase()] = 1);
    const ok = [], bad = []; let fixed = 0;
    st.getRange(2, 1, n - 1, 5).getValues().forEach(function (r, i) {
      const name = String(r[0]).trim(), ph = e164(r[1]), l = matchLga(r[2]), ward = String(r[3]).trim(), pu = String(r[4]).trim(), why = [];
      if (!name || !ward || !pu) why.push('missing field'); if (!ph) why.push('bad phone number'); if (!l) why.push('unknown LGA');
      if (!why.length) { const key = ph + '|' + name.toLowerCase(); if (seen[key]) why.push('duplicate (same phone and name)'); else seen[key] = 1; }
      if (why.length) bad.push([i + 2, r[0], r[1], r[2], why.join(', ')]);
      else { if (l.fixed) fixed++; ok.push([Utilities.getUuid(), name, ph, String(r[1]), l.lga, ward, pu, 'Not called', 'TRUE', new Date().toISOString()]); }
    });
    const es = sh(T.ERR); es.getRange(2, 1, Math.max(es.getLastRow(), 2), 5).clearContent();
    if (bad.length) es.getRange(2, 1, bad.length, 5).setNumberFormat('@').setValues(bad);
    const go = ui().alert('Preview', ok.length + ' valid rows (' + fixed + ' with a corrected LGA spelling).\n' + bad.length + ' rejected (see the "Import errors" tab).\n\nImport the valid rows now?', ui().ButtonSet.YES_NO);
    if (go !== ui().Button.YES) return;
    const cs = sh(T.CONTACTS);
    if (replace) { cs.getRange(2, 1, Math.max(cs.getLastRow(), 2), CH.length).clearContent(); if (keep.length) cs.getRange(2, 1, keep.length, CH.length).setNumberFormat('@').setValues(keep.map(c => CH.map(k => c[k]))); }
    if (ok.length) cs.getRange(cs.getLastRow() + 1, 1, ok.length, CH.length).setNumberFormat('@').setValues(ok);
    put(T.UP, [[new Date().toISOString(), me(), replace ? 'replace' : 'append', ok.length, bad.length]]);
    audit(me(), 'directory_import', ok.length + ' imported, ' + bad.length + ' rejected');
    st.getRange(2, 1, n - 1, 5).clearContent();
    ui().alert('Imported ' + ok.length + ' contacts.');
  } finally { lock.releaseLock(); }
}
