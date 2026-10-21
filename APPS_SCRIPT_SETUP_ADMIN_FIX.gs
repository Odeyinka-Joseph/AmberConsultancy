// AMBER setup/admin repair patch.
// In your Google Apps Script project, REPLACE the existing setup(), toggleManual(),
// and add this setAdminPassword() function. Then save and redeploy a NEW VERSION.

function ensureGrid(s, rows, cols) {
  if (s.getMaxRows() < rows) s.insertRowsAfter(s.getMaxRows(), rows - s.getMaxRows());
  if (s.getMaxColumns() < cols) s.insertColumnsAfter(s.getMaxColumns(), cols - s.getMaxColumns());
}

function setup() {
  const defs = [
    [T.USERS, UH],
    [T.CONTACTS, CH],
    [T.RESP, RCOLS.map(c => c[0])],
    [T.ATT, AH],
    [T.SET, ['key', 'value']],
    [T.AUDIT, ['time', 'who', 'action', 'details']],
    [T.STAGE, ['Name', 'Phone number', 'LGA', 'Ward', 'Polling unit']],
    [T.ERR, ['Row in Staging', 'Name', 'Phone number', 'LGA', 'Reason']],
    [T.UP, ['time', 'who', 'mode', 'rows imported', 'rows rejected']]
  ];

  defs.forEach(function(d) {
    const s = sh(d[0]) || SpreadsheetApp.getActive().insertSheet(d[0]);
    ensureGrid(s, 2001, d[1].length);
    if (s.getLastRow() === 0) {
      s.getRange(1, 1, 1, d[1].length)
        .setNumberFormat('@')
        .setValues([d[1]])
        .setFontWeight('bold');
      s.setFrozenRows(1);
    }
    s.getRange(2, 1, 2000, d[1].length).setNumberFormat('@');
  });

  if (!props().getProperty('SECRET')) {
    props().setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  }

  if (table(T.SET, ['key', 'value']).length === 0) {
    put(T.SET, [['allow_not_in_directory', 'FALSE']]);
  }

  ui().alert('Sheets are ready. Next set/reset the admin password, then redeploy the Web App as a new version.');
}

function setAdminPassword() {
  const pw = ask('New admin password (minimum 6 characters):');
  if (pw === null) return;
  if (String(pw).length < 6) return ui().alert('Use at least 6 characters.');

  let u = getUser('admin');
  if (!u) {
    put(T.USERS, [['admin', 'Administrator', 'admin', 'TRUE', makeHash(pw), 'FALSE', new Date().toISOString()]]);
  } else {
    const s = sh(T.USERS);
    s.getRange(u._row, UH.indexOf('role') + 1).setValue('admin');
    s.getRange(u._row, UH.indexOf('active') + 1).setValue('TRUE');
    s.getRange(u._row, UH.indexOf('pw_hash') + 1).setValue(makeHash(pw));
    s.getRange(u._row, UH.indexOf('must_change') + 1).setValue('FALSE');
  }

  audit(me(), 'admin_password_set', 'admin');
  ui().alert('Admin password updated successfully. Username: admin');
}

function toggleManual() {
  const s = sh(T.SET), rows = table(T.SET, ['key', 'value']);
  let r = rows.filter(x => x.key === 'allow_not_in_directory')[0];
  if (!r) {
    put(T.SET, [['allow_not_in_directory', 'FALSE']]);
    r = table(T.SET, ['key', 'value']).filter(x => x.key === 'allow_not_in_directory')[0];
  }
  const on = !isTrue(r.value);
  s.getRange(r._row, 2).setValue(on ? 'TRUE' : 'FALSE');
  audit(me(), 'allow_not_in_directory', String(on));
  ui().alert('Allow numbers not in directory: ' + (on ? 'ON' : 'OFF'));
}
