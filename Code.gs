/**
 * Respirator Fitment Attendance — Google Apps Script backend
 *
 * After editing: Deploy → Manage deployments → Edit → New version → Deploy
 * Who has access: Anyone | Execute as: Me
 *
 * Auto-finds the sheet tab that has columns:
 * Name | Control Number | Company | Date | Size
 */

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
  return { numeric: numeric, words: words, display: words };
}

function norm_(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function empStr_(v) {
  // Control number / ID is TEXT — keep leading zeros (e.g. 0123456)
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return '';
  if (typeof v === 'number' && isFinite(v)) {
    // Already stored as a number in Sheets (zeros already lost) — keep digits only
    return String(Math.round(v));
  }
  return String(v).replace(/^\u200B+/, '').trim();
}

function setEmpCell_(sheet, row, col1Based, emp) {
  const cell = sheet.getRange(row, col1Based);
  cell.setNumberFormat('@'); // plain text — preserves leading zeros
  cell.setValue(empStr_(emp));
}

/**
 * Find the register sheet + header row by looking for Name + Control Number headers.
 * Picks the sheet with the most real data rows.
 */
function locateRegister_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let best = null;

  for (let s = 0; s < sheets.length; s++) {
    const sh = sheets[s];
    if (sh.getName() === CONFIG_SHEET) continue;

    const lastRow = Math.max(sh.getLastRow(), 1);
    const lastCol = Math.max(sh.getLastColumn(), 5);
    const scanRows = Math.min(lastRow, 15);
    const values = sh.getRange(1, 1, scanRows, Math.min(lastCol, 10)).getValues();

    for (let r = 0; r < values.length; r++) {
      const row = values[r].map(norm_);
      let nameCol = -1;
      let empCol = -1;
      let companyCol = -1;
      let dateCol = -1;
      let sizeCol = -1;

      for (let c = 0; c < row.length; c++) {
        const h = row[c];
        if (!h) continue;
        if (nameCol < 0 && (h.indexOf('attendee') >= 0 || h.indexOf('name') >= 0 || h.indexOf('surname') >= 0)) nameCol = c;
        if (empCol < 0 && (h.indexOf('control') >= 0 || h.indexOf('employee') >= 0 || h === 'empno' || h === 'id' || h.indexOf('identity') >= 0 || h.indexOf('sasolcontrol') >= 0)) empCol = c;
        if (companyCol < 0 && h.indexOf('company') >= 0) companyCol = c;
        if (dateCol < 0 && h.indexOf('date') >= 0) dateCol = c;
        if (sizeCol < 0 && (h.indexOf('size') >= 0 || h.indexOf('respirator') >= 0)) sizeCol = c;
      }

      if (nameCol < 0 || empCol < 0) continue;

      // Count data rows under this header
      const dataStart = r + 2; // 1-based sheet row after header
      const endRow = sh.getLastRow();
      let count = 0;
      if (endRow >= dataStart) {
        const empVals = sh.getRange(dataStart, empCol + 1, endRow, 1).getValues();
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
        count: count
      };

      if (!best || candidate.count > best.count) best = candidate;
      break; // use first matching header row on this sheet
    }
  }

  if (best) return best;

  // Fallback: create Attendance sheet
  let sh = ss.getSheetByName('Attendance');
  if (!sh) {
    sh = ss.insertSheet('Attendance');
    sh.appendRow([
      'Attendee Name and Surname',
      'Sasol Control Number',
      'Company',
      'Date of Training',
      'Respirator Size Required (S/M/L)'
    ]);
  }
  return {
    sheet: sh,
    headerRow: 1,
    nameCol: 0,
    empCol: 1,
    companyCol: 2,
    dateCol: 3,
    sizeCol: 4,
    count: 0
  };
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
  const loc = locateRegister_();
  return {
    date: today.display,
    dateNumeric: today.numeric,
    dateWords: today.words,
    location: String(sh.getRange('A2').getValue() || 'Sasol Club'),
    facilitator: String(sh.getRange('A3').getValue() || ''),
    sheetName: loc.sheet.getName(),
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

function formatDateCell_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    const dd = Utilities.formatDate(value, TZ, 'dd');
    const mm = Utilities.formatDate(value, TZ, 'MM');
    const y = Utilities.formatDate(value, TZ, 'yyyy');
    const d = Number(Utilities.formatDate(value, TZ, 'd'));
    const m = Number(Utilities.formatDate(value, TZ, 'M'));
    return d + ' ' + MONTHS[m - 1] + ' ' + y;
  }
  return String(value || '').trim();
}

function rowToRecordFrom_(row, map, sheetRow) {
  return {
    name: String(row[map.nameCol] || '').trim(),
    employeeNumber: empStr_(row[map.empCol]),
    company: String(row[map.companyCol] || '').trim(),
    date: formatDateCell_(row[map.dateCol]),
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

  const maxCol = Math.max(map.nameCol, map.empCol, map.companyCol, map.dateCol, map.sizeCol) + 1;
  const values = sh.getRange(start, 1, last, maxCol).getValues();
  const displays = sh.getRange(start, 1, last, maxCol).getDisplayValues();
  const out = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i].slice();
    // Prefer display text for control number / ID so leading zeros survive
    row[map.empCol] = displays[i][map.empCol];
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

function handleRegister_(name, employeeNumber, company, size) {
  name = String(name || '').trim();
  employeeNumber = empStr_(employeeNumber);
  company = String(company || '').trim();
  size = String(size || '').trim().toUpperCase();
  const today = todayParts_();
  const date = today.display;

  if (!name || !employeeNumber || !company) {
    return { success: false, error: 'Name, control number / ID, and company are required.' };
  }

  const existing = findRowByEmployee_(employeeNumber);
  if (existing) {
    const map = existing.map;
    const maxCol = Math.max(map.nameCol, map.empCol, map.companyCol, map.dateCol, map.sizeCol) + 1;
    const rowVals = existing.sheet.getRange(existing.row, 1, 1, maxCol).getValues()[0];
    const displays = existing.sheet.getRange(existing.row, 1, 1, maxCol).getDisplayValues()[0];
    rowVals[map.empCol] = displays[map.empCol];
    return { success: true, duplicate: true, data: rowToRecordFrom_(rowVals, map, existing.row) };
  }

  const map = locateRegister_();
  const sh = map.sheet;
  const maxCol = Math.max(map.nameCol, map.empCol, map.companyCol, map.dateCol, map.sizeCol) + 1;
  const row = new Array(maxCol);
  for (let i = 0; i < maxCol; i++) row[i] = '';
  row[map.nameCol] = name;
  row[map.empCol] = employeeNumber;
  row[map.companyCol] = company;
  row[map.dateCol] = date;
  row[map.sizeCol] = size;
  sh.appendRow(row);
  const newRow = sh.getLastRow();
  setEmpCell_(sh, newRow, map.empCol + 1, employeeNumber);

  writeConfig_({ date: date });
  return {
    success: true,
    duplicate: false,
    data: { name: name, employeeNumber: employeeNumber, company: company, date: date, size: size, row: newRow, dateNumeric: today.numeric, dateWords: today.words },
    sheetName: sh.getName()
  };
}

function handleUpdate_(employeeNumber, name, company, date, size, newEmployeeNumber, sheetRow) {
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

  const maxCol = Math.max(map.nameCol, map.empCol, map.companyCol, map.dateCol, map.sizeCol) + 1;
  const rowVals = sh.getRange(row, 1, 1, maxCol).getValues()[0];
  const displays = sh.getRange(row, 1, 1, maxCol).getDisplayValues()[0];
  rowVals[map.empCol] = displays[map.empCol];
  return { success: true, data: rowToRecordFrom_(rowVals, map, row) };
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
      return jsonOut_({
        success: true,
        ok: true,
        time: todayParts_().display,
        sheetName: loc.sheet.getName(),
        recordCount: loc.count
      });
    }
    if (action === 'config') {
      return jsonOut_({ success: true, data: readConfig_() });
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
      return jsonOut_(handleUpdate_(p.employeeNumber, p.name, p.company, p.date, p.size, p.newEmployeeNumber, p.row));
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
      return jsonOut_(handleUpdate_(body.employeeNumber, body.name, body.company, body.date, body.size, body.newEmployeeNumber, body.row));
    }
    if (action === 'delete') {
      return jsonOut_(handleDelete_(body.employeeNumber, body.row));
    }
    return jsonOut_({ success: false, error: 'Unknown action' });
  } catch (err) {
    return jsonOut_({ success: false, error: String(err) });
  }
}
