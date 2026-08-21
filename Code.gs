/**
 * Respirator Fitment Attendance — Google Apps Script backend
 *
 * After editing: Deploy → Manage deployments → Edit → New version → Deploy
 * Who has access: Anyone | Execute as: Me
 *
 * Sheets used:
 *  - Register tab (auto-found): Name | Control Number | Company | Date | Size | Time Signed In
 *  - Config: session location / facilitator
 *  - Access: moderator PIN lives in B1 (label in A1). Created automatically.
 */

const CONFIG_SHEET = 'Config';
const ACCESS_SHEET = 'Access'; // moderator PIN — put the PIN in cell B1
const TZ = 'Africa/Johannesburg';
const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const DEFAULT_PIN = '2026'; // only used once when Access sheet is first created

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
  const time = Utilities.formatDate(now, TZ, 'HH:mm'); // 24h Joburg time
  const stamp = words + ' · ' + time;
  return { numeric: numeric, words: words, display: words, time: time, stamp: stamp, now: now };
}

function norm_(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function empStr_(v) {
  // Control number / ID is TEXT — keep leading zeros (e.g. 0123456)
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return '';
  if (typeof v === 'number' && isFinite(v)) {
    return String(Math.round(v));
  }
  return String(v).replace(/^\u200B+/, '').trim();
}

function setEmpCell_(sheet, row, col1Based, emp) {
  const cell = sheet.getRange(row, col1Based);
  cell.setNumberFormat('@');
  cell.setValue(empStr_(emp));
}

function formatDateCell_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    const d = Number(Utilities.formatDate(value, TZ, 'd'));
    const m = Number(Utilities.formatDate(value, TZ, 'M'));
    const y = Utilities.formatDate(value, TZ, 'yyyy');
    return d + ' ' + MONTHS[m - 1] + ' ' + y;
  }
  return String(value || '').trim();
}

function formatTimeCell_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TZ, 'HH:mm');
  }
  const s = String(value || '').trim();
  const m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?/);
  if (m) {
    const hh = ('0' + m[1]).slice(-2);
    return hh + ':' + m[2];
  }
  return s;
}

/** Access sheet: A1 = "Moderator PIN", B1 = the secret PIN (not exposed to the website). */
function ensureAccessSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ACCESS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ACCESS_SHEET);
    sh.getRange('A1').setValue('Moderator PIN');
    sh.getRange('B1').setValue(DEFAULT_PIN);
    sh.getRange('A2').setValue('Change the PIN in cell B1. The website never shows this value.');
    sh.getRange('A3').setValue('After changing B1, moderators use the new PIN immediately (no redeploy).');
    sh.setColumnWidth(1, 160);
    sh.setColumnWidth(2, 120);
  } else {
    if (!String(sh.getRange('A1').getValue() || '').trim()) sh.getRange('A1').setValue('Moderator PIN');
    if (!String(sh.getRange('B1').getValue() || '').trim()) sh.getRange('B1').setValue(DEFAULT_PIN);
  }
  return sh;
}

function readModeratorPin_() {
  const sh = ensureAccessSheet_();
  return String(sh.getRange('B1').getDisplayValue() || '').trim();
}

function verifyPin_(pin) {
  const expected = readModeratorPin_();
  const got = String(pin || '').trim();
  if (!expected) return { success: true, ok: false, error: 'No PIN set in Access!B1' };
  if (!got) return { success: true, ok: false, error: 'Enter the moderator PIN' };
  return { success: true, ok: got === expected };
}

/**
 * Find the register sheet + header row by looking for Name + Control Number headers.
 * Ensures a "Time Signed In" column exists.
 */
function locateRegister_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let best = null;

  for (let s = 0; s < sheets.length; s++) {
    const sh = sheets[s];
    const name = sh.getName();
    if (name === CONFIG_SHEET || name === ACCESS_SHEET) continue;

    const lastRow = Math.max(sh.getLastRow(), 1);
    const lastCol = Math.max(sh.getLastColumn(), 6);
    const scanRows = Math.min(lastRow, 15);
    const values = sh.getRange(1, 1, scanRows, Math.min(lastCol, 12)).getValues();

    for (let r = 0; r < values.length; r++) {
      const row = values[r].map(norm_);
      let nameCol = -1;
      let empCol = -1;
      let companyCol = -1;
      let dateCol = -1;
      let sizeCol = -1;
      let timeCol = -1;

      for (let c = 0; c < row.length; c++) {
        const h = row[c];
        if (!h) continue;
        if (nameCol < 0 && (h.indexOf('attendee') >= 0 || h.indexOf('name') >= 0 || h.indexOf('surname') >= 0)) nameCol = c;
        if (empCol < 0 && (h.indexOf('control') >= 0 || h.indexOf('employee') >= 0 || h === 'empno' || h === 'id' || h.indexOf('identity') >= 0 || h.indexOf('sasolcontrol') >= 0)) empCol = c;
        if (companyCol < 0 && h.indexOf('company') >= 0) companyCol = c;
        if (dateCol < 0 && h.indexOf('date') >= 0) dateCol = c;
        if (sizeCol < 0 && (h.indexOf('size') >= 0 || h.indexOf('respirator') >= 0)) sizeCol = c;
        if (timeCol < 0 && (h.indexOf('time') >= 0 || h.indexOf('signedin') >= 0 || h.indexOf('timestamp') >= 0)) timeCol = c;
      }

      if (nameCol < 0 || empCol < 0) continue;

      const dataStart = r + 2;
      const endRow = sh.getLastRow();
      let count = 0;
      if (endRow >= dataStart) {
        const empVals = sh.getRange(dataStart, empCol + 1, endRow, 1).getDisplayValues();
        for (let i = 0; i < empVals.length; i++) {
          if (empStr_(empVals[i][0]) !== '') count++;
        }
      }

      const candidate = {
        sheet: sh,
        headerRow: r + 1,
        nameCol: nameCol,
        empCol: empCol,
        companyCol: companyCol >= 0 ? companyCol : 2,
        dateCol: dateCol >= 0 ? dateCol : 3,
        sizeCol: sizeCol >= 0 ? sizeCol : 4,
        timeCol: timeCol,
        count: count
      };

      if (!best || candidate.count > best.count) best = candidate;
      break;
    }
  }

  if (!best) {
    let sh = ss.getSheetByName('Attendance');
    if (!sh) {
      sh = ss.insertSheet('Attendance');
      sh.appendRow([
        'Attendee Name and Surname',
        'Sasol Control Number',
        'Company',
        'Date of Training',
        'Respirator Size Required (S/M/L)',
        'Time Signed In'
      ]);
    }
    best = {
      sheet: sh,
      headerRow: 1,
      nameCol: 0,
      empCol: 1,
      companyCol: 2,
      dateCol: 3,
      sizeCol: 4,
      timeCol: 5,
      count: 0
    };
  }

  ensureTimeColumn_(best);
  return best;
}

/** Add "Time Signed In" header if the register sheet does not have a time column yet. */
function ensureTimeColumn_(map) {
  if (map.timeCol >= 0) return map;
  const sh = map.sheet;
  const headerRow = map.headerRow;
  const nextCol = Math.max(sh.getLastColumn(), map.sizeCol + 1, map.dateCol + 1, map.empCol + 1, map.companyCol + 1) + 1;
  sh.getRange(headerRow, nextCol).setValue('Time Signed In');
  map.timeCol = nextCol - 1; // 0-based
  return map;
}

function maxCol_(map) {
  return Math.max(map.nameCol, map.empCol, map.companyCol, map.dateCol, map.sizeCol, map.timeCol >= 0 ? map.timeCol : 0) + 1;
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
  ensureAccessSheet_();
  const sh = getConfigSheet_();
  const today = todayParts_();
  sh.getRange('A1').setValue(today.display);
  const loc = locateRegister_();
  return {
    date: today.display,
    dateNumeric: today.numeric,
    dateWords: today.words,
    time: today.time,
    location: String(sh.getRange('A2').getValue() || 'Sasol Club'),
    facilitator: String(sh.getRange('A3').getValue() || ''),
    sheetName: loc.sheet.getName(),
    accessSheet: ACCESS_SHEET,
    recordCount: loc.count
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

function rowToRecordFrom_(row, map, sheetRow) {
  const time = map.timeCol >= 0 ? formatTimeCell_(row[map.timeCol]) : '';
  return {
    name: String(row[map.nameCol] || '').trim(),
    employeeNumber: empStr_(row[map.empCol]),
    company: String(row[map.companyCol] || '').trim(),
    date: formatDateCell_(row[map.dateCol]),
    time: time,
    size: String(row[map.sizeCol] || '').trim().toUpperCase(),
    row: sheetRow || 0
  };
}

function listRecords_() {
  const map = locateRegister_();
  const sh = map.sheet;
  const start = map.headerRow + 1;
  const last = sh.getLastRow();
  if (last < start) return [];

  const maxCol = maxCol_(map);
  const values = sh.getRange(start, 1, last, maxCol).getValues();
  const displays = sh.getRange(start, 1, last, maxCol).getDisplayValues();
  const out = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i].slice();
    row[map.empCol] = displays[i][map.empCol];
    if (map.timeCol >= 0) row[map.timeCol] = displays[i][map.timeCol];
    const rec = rowToRecordFrom_(row, map, start + i);
    if (rec.employeeNumber !== '') out.push(rec);
  }
  return out;
}

function findRowByEmployee_(emp) {
  const map = locateRegister_();
  const sh = map.sheet;
  const start = map.headerRow + 1;
  const last = sh.getLastRow();
  if (last < start) return null;

  const nums = sh.getRange(start, map.empCol + 1, last, 1).getDisplayValues();
  const target = empStr_(emp);
  for (let i = 0; i < nums.length; i++) {
    if (empStr_(nums[i][0]) === target) {
      return { sheet: sh, row: start + i, map: map };
    }
  }
  return null;
}

function findRowBySheetRow_(sheetRow) {
  const map = locateRegister_();
  const row = Number(sheetRow);
  if (!row || row <= map.headerRow) return null;
  const sh = map.sheet;
  if (row > sh.getLastRow()) return null;
  return { sheet: sh, row: row, map: map };
}

function readRowRecord_(found) {
  const map = found.map;
  const maxCol = maxCol_(map);
  const rowVals = found.sheet.getRange(found.row, 1, 1, maxCol).getValues()[0];
  const displays = found.sheet.getRange(found.row, 1, 1, maxCol).getDisplayValues()[0];
  rowVals[map.empCol] = displays[map.empCol];
  if (map.timeCol >= 0) rowVals[map.timeCol] = displays[map.timeCol];
  return rowToRecordFrom_(rowVals, map, found.row);
}

function handleRegister_(name, employeeNumber, company, size) {
  name = String(name || '').trim();
  employeeNumber = empStr_(employeeNumber);
  company = String(company || '').trim();
  size = String(size || '').trim().toUpperCase();
  const today = todayParts_();
  const date = today.display;
  const time = today.time;

  if (!name || !employeeNumber || !company) {
    return { success: false, error: 'Name, control number / ID, and company are required.' };
  }

  const existing = findRowByEmployee_(employeeNumber);
  if (existing) {
    return { success: true, duplicate: true, data: readRowRecord_(existing) };
  }

  const map = locateRegister_();
  const sh = map.sheet;
  const maxCol = maxCol_(map);
  const row = new Array(maxCol);
  for (let i = 0; i < maxCol; i++) row[i] = '';
  row[map.nameCol] = name;
  row[map.empCol] = employeeNumber;
  row[map.companyCol] = company;
  row[map.dateCol] = date;
  row[map.sizeCol] = size;
  if (map.timeCol >= 0) row[map.timeCol] = time;
  sh.appendRow(row);
  const newRow = sh.getLastRow();
  setEmpCell_(sh, newRow, map.empCol + 1, employeeNumber);
  if (map.timeCol >= 0) {
    sh.getRange(newRow, map.timeCol + 1).setNumberFormat('@');
    sh.getRange(newRow, map.timeCol + 1).setValue(time);
  }

  writeConfig_({ date: date });
  return {
    success: true,
    duplicate: false,
    data: {
      name: name,
      employeeNumber: employeeNumber,
      company: company,
      date: date,
      time: time,
      size: size,
      row: newRow,
      dateNumeric: today.numeric,
      dateWords: today.words
    },
    sheetName: sh.getName()
  };
}

function handleUpdate_(employeeNumber, name, company, date, size, newEmployeeNumber, sheetRow, time) {
  employeeNumber = empStr_(employeeNumber);
  let found = null;
  if (sheetRow) found = findRowBySheetRow_(sheetRow);
  if (!found) found = findRowByEmployee_(employeeNumber);
  if (!found) return { success: false, error: 'Employee not found' };

  const sh = found.sheet;
  const row = found.row;
  const map = found.map;

  if (newEmployeeNumber !== undefined && newEmployeeNumber !== null && String(newEmployeeNumber).trim() !== '') {
    const nextEmp = empStr_(newEmployeeNumber);
    const currentEmp = empStr_(sh.getRange(row, map.empCol + 1).getDisplayValue());
    if (nextEmp !== currentEmp) {
      const clash = findRowByEmployee_(nextEmp);
      if (clash && clash.row !== row) {
        return { success: false, error: 'That control number / ID already exists (duplicate not allowed).' };
      }
      setEmpCell_(sh, row, map.empCol + 1, nextEmp);
    }
  }
  if (name !== undefined && name !== null && String(name).length) sh.getRange(row, map.nameCol + 1).setValue(String(name).trim());
  if (company !== undefined && company !== null && String(company).length) sh.getRange(row, map.companyCol + 1).setValue(String(company).trim());
  if (date !== undefined && date !== null && String(date).length) sh.getRange(row, map.dateCol + 1).setValue(String(date).trim());
  if (size !== undefined && size !== null && String(size).length) sh.getRange(row, map.sizeCol + 1).setValue(String(size).trim().toUpperCase());
  if (time !== undefined && time !== null && String(time).length && map.timeCol >= 0) {
    sh.getRange(row, map.timeCol + 1).setNumberFormat('@');
    sh.getRange(row, map.timeCol + 1).setValue(formatTimeCell_(time));
  }

  return { success: true, data: readRowRecord_(found) };
}

function handleDelete_(employeeNumber, sheetRow) {
  let found = null;
  if (sheetRow) found = findRowBySheetRow_(sheetRow);
  if (!found) found = findRowByEmployee_(empStr_(employeeNumber));
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
      const loc = locateRegister_();
      ensureAccessSheet_();
      return jsonOut_({
        success: true,
        ok: true,
        time: todayParts_().stamp,
        sheetName: loc.sheet.getName(),
        accessSheet: ACCESS_SHEET,
        recordCount: loc.count
      });
    }
    if (action === 'config') {
      return jsonOut_({ success: true, data: readConfig_() });
    }
    if (action === 'verifypin') {
      return jsonOut_(verifyPin_(p.pin));
    }
    if (action === 'list') {
      const loc = locateRegister_();
      return jsonOut_({
        success: true,
        data: listRecords_(),
        sheetName: loc.sheet.getName()
      });
    }
    if (action === 'search') {
      const q = String(p.query || '').trim();
      const matches = listRecords_().filter(function(r) {
        return String(r.employeeNumber).slice(-4) === q;
      });
      return jsonOut_({ success: true, data: matches });
    }
    if (action === 'register') {
      return jsonOut_(handleRegister_(p.name, p.employeeNumber, p.company, p.size));
    }
    if (action === 'update') {
      return jsonOut_(handleUpdate_(p.employeeNumber, p.name, p.company, p.date, p.size, p.newEmployeeNumber, p.row, p.time));
    }
    if (action === 'delete') {
      return jsonOut_(handleDelete_(p.employeeNumber, p.row));
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

    if (action === 'verifypin') {
      return jsonOut_(verifyPin_(body.pin));
    }
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
      return jsonOut_(handleUpdate_(body.employeeNumber, body.name, body.company, body.date, body.size, body.newEmployeeNumber, body.row, body.time));
    }
    if (action === 'delete') {
      return jsonOut_(handleDelete_(body.employeeNumber, body.row));
    }
    return jsonOut_({ success: false, error: 'Unknown action' });
  } catch (err) {
    return jsonOut_({ success: false, error: String(err) });
  }
}
