/**
 * Respirator Fitment Attendance — Google Apps Script backend
 *
 * After editing: Deploy → Manage deployments → Edit → New version → Deploy
 * Who has access: Anyone | Execute as: Me
 *
 * Sheets: "Attendance" (A–E) + "Config" (A1 date, A2 location, A3 facilitator)
 */

const ATTENDANCE_SHEET = 'Attendance';
const CONFIG_SHEET = 'Config';
const TZ = 'Africa/Johannesburg';
const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function todayParts_() {
  const now = new Date();
  const d = Number(Utilities.formatDate(now, TZ, 'd'));
  const m = Number(Utilities.formatDate(now, TZ, 'M'));
  const y = Number(Utilities.formatDate(now, TZ, 'yyyy'));
  const dd = Utilities.formatDate(now, TZ, 'dd');
  const mm = Utilities.formatDate(now, TZ, 'MM');
  const numeric = dd + '/' + mm + '/' + y;
  const words = d + ' ' + MONTHS[m - 1] + ' ' + y;
  return { numeric: numeric, words: words, display: numeric + ' (' + words + ')' };
}

function getAttendanceSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ATTENDANCE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ATTENDANCE_SHEET);
    sh.appendRow([
      'Attendee Name and Surname',
      'Sasol Control Number',
      'Company',
      'Date of Training',
      'Respirator Size Required (S/M/L)'
    ]);
  }
  return sh;
}

function getConfigSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG_SHEET);
    sh.getRange('A1').setValue('');
    sh.getRange('A2').setValue('Sasol Club');
    sh.getRange('A3').setValue('');
  }
  return sh;
}

function readConfig_() {
  const sh = getConfigSheet_();
  const today = todayParts_();
  sh.getRange('A1').setValue(today.display);
  return {
    date: today.display,
    dateNumeric: today.numeric,
    dateWords: today.words,
    location: String(sh.getRange('A2').getValue() || 'Sasol Club'),
    facilitator: String(sh.getRange('A3').getValue() || '')
  };
}

function writeConfig_(cfg) {
  const sh = getConfigSheet_();
  if (cfg.date !== undefined) sh.getRange('A1').setValue(cfg.date);
  if (cfg.location !== undefined) sh.getRange('A2').setValue(cfg.location);
  if (cfg.facilitator !== undefined) sh.getRange('A3').setValue(cfg.facilitator);
  return {
    date: String(sh.getRange('A1').getValue() || ''),
    location: String(sh.getRange('A2').getValue() || 'Sasol Club'),
    facilitator: String(sh.getRange('A3').getValue() || '')
  };
}

function formatDateCell_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    const dd = Utilities.formatDate(value, TZ, 'dd');
    const mm = Utilities.formatDate(value, TZ, 'MM');
    const y = Utilities.formatDate(value, TZ, 'yyyy');
    const d = Number(Utilities.formatDate(value, TZ, 'd'));
    const m = Number(Utilities.formatDate(value, TZ, 'M'));
    return dd + '/' + mm + '/' + y + ' (' + d + ' ' + MONTHS[m - 1] + ' ' + y + ')';
  }
  return String(value || '');
}

function rowToRecord_(row) {
  return {
    name: String(row[0] || ''),
    employeeNumber: String(row[1] || ''),
    company: String(row[2] || ''),
    date: formatDateCell_(row[3]),
    size: String(row[4] || '').trim().toUpperCase()
  };
}

function listRecords_() {
  const sh = getAttendanceSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last, 5).getValues();
  return values
    .filter(r => String(r[1] || '').trim() !== '')
    .map(rowToRecord_);
}

function findRowByEmployee_(emp) {
  const sh = getAttendanceSheet_();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const nums = sh.getRange(2, 2, last, 1).getValues();
  const target = String(emp).trim();
  for (let i = 0; i < nums.length; i++) {
    if (String(nums[i][0]).trim() === target) {
      return { sheet: sh, row: i + 2 };
    }
  }
  return null;
}

function handleRegister_(name, employeeNumber, company, size) {
  name = String(name || '').trim();
  employeeNumber = String(employeeNumber || '').trim();
  company = String(company || '').trim();
  size = String(size || '').trim().toUpperCase();
  const today = todayParts_();
  const date = today.display;

  if (!name || !employeeNumber || !company) {
    return { success: false, error: 'Name, control number, and company are required.' };
  }

  const existing = findRowByEmployee_(employeeNumber);
  if (existing) {
    const rowVals = existing.sheet.getRange(existing.row, 1, 1, 5).getValues()[0];
    return { success: true, duplicate: true, data: rowToRecord_(rowVals) };
  }

  getAttendanceSheet_().appendRow([name, employeeNumber, company, date, size]);
  writeConfig_({ date: date });
  return {
    success: true,
    duplicate: false,
    data: { name: name, employeeNumber: employeeNumber, company: company, date: date, size: size, dateNumeric: today.numeric, dateWords: today.words }
  };
}

function handleUpdate_(employeeNumber, name, company, date, size) {
  employeeNumber = String(employeeNumber || '').trim();
  const found = findRowByEmployee_(employeeNumber);
  if (!found) return { success: false, error: 'Employee not found' };

  const sh = found.sheet;
  const row = found.row;
  if (name !== undefined && name !== null && String(name).length) sh.getRange(row, 1).setValue(String(name).trim());
  if (company !== undefined && company !== null && String(company).length) sh.getRange(row, 3).setValue(String(company).trim());
  if (date !== undefined && date !== null && String(date).length) sh.getRange(row, 4).setValue(String(date).trim());
  if (size !== undefined && size !== null && String(size).length) sh.getRange(row, 5).setValue(String(size).trim().toUpperCase());

  const rowVals = sh.getRange(row, 1, 1, 5).getValues()[0];
  return { success: true, data: rowToRecord_(rowVals) };
}

function handleDelete_(employeeNumber) {
  employeeNumber = String(employeeNumber || '').trim();
  const found = findRowByEmployee_(employeeNumber);
  if (!found) return { success: false, error: 'Employee not found' };
  found.sheet.deleteRow(found.row);
  return { success: true };
}

function doGet(e) {
  try {
    e = e || { parameter: {} };
    const p = e.parameter || {};
    const action = String(p.action || '').toLowerCase();

    if (action === 'ping' || action === 'health') {
      return jsonOut_({ success: true, ok: true, time: todayParts_().display });
    }
    if (action === 'config') {
      return jsonOut_({ success: true, data: readConfig_() });
    }
    if (action === 'list') {
      return jsonOut_({ success: true, data: listRecords_() });
    }
    if (action === 'search') {
      const q = String(p.query || '').trim();
      const matches = listRecords_().filter(r => String(r.employeeNumber).slice(-4) === q);
      return jsonOut_({ success: true, data: matches });
    }
    // GET fallbacks (more reliable from GitHub Pages than POST)
    if (action === 'register') {
      return jsonOut_(handleRegister_(p.name, p.employeeNumber, p.company, p.size));
    }
    if (action === 'update') {
      return jsonOut_(handleUpdate_(p.employeeNumber, p.name, p.company, p.date, p.size));
    }
    if (action === 'delete') {
      return jsonOut_(handleDelete_(p.employeeNumber));
    }
    if (action === 'setconfig') {
      const today = todayParts_();
      const data = writeConfig_({
        date: today.display,
        location: p.location,
        facilitator: p.facilitator
      });
      return jsonOut_({ success: true, data: data });
    }
    return jsonOut_({ success: false, error: 'Unknown action' });
  } catch (err) {
    return jsonOut_({ success: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(body.action || '').toLowerCase();

    if (action === 'setconfig') {
      const today = todayParts_();
      const data = writeConfig_({
        date: today.display,
        location: body.location,
        facilitator: body.facilitator
      });
      return jsonOut_({ success: true, data: data });
    }
    if (action === 'register') {
      return jsonOut_(handleRegister_(body.name, body.employeeNumber, body.company, body.size));
    }
    if (action === 'update') {
      return jsonOut_(handleUpdate_(body.employeeNumber, body.name, body.company, body.date, body.size));
    }
    if (action === 'delete') {
      return jsonOut_(handleDelete_(body.employeeNumber));
    }
    return jsonOut_({ success: false, error: 'Unknown action' });
  } catch (err) {
    return jsonOut_({ success: false, error: String(err) });
  }
}
