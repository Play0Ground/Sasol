/**
 * Respirator Fitment Attendance — Google Apps Script backend
 *
 * SETUP
 * 1. Open your Google Sheet (attendance register).
 * 2. Extensions → Apps Script → paste this entire file (replace old code).
 * 3. Deploy → Manage deployments → Edit (pencil) → New version → Deploy
 *    OR Deploy → New deployment → Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copy the /exec URL into index.html as API_URL if it changed.
 *
 * SHEETS
 * - "Attendance" columns: A Name | B Control Number | C Company | D Date | E Size
 * - "Config" cells: A1=date, A2=location, A3=facilitator  (created automatically)
 */

const ATTENDANCE_SHEET = 'Attendance';
const CONFIG_SHEET = 'Config';

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
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
  return {
    date: String(sh.getRange('A1').getValue() || ''),
    location: String(sh.getRange('A2').getValue() || 'Sasol Club'),
    facilitator: String(sh.getRange('A3').getValue() || '')
  };
}

function writeConfig_(cfg) {
  const sh = getConfigSheet_();
  if (cfg.date !== undefined) sh.getRange('A1').setValue(cfg.date);
  if (cfg.location !== undefined) sh.getRange('A2').setValue(cfg.location);
  if (cfg.facilitator !== undefined) sh.getRange('A3').setValue(cfg.facilitator);
  return readConfig_();
}

function rowToRecord_(row) {
  return {
    name: String(row[0] || ''),
    employeeNumber: String(row[1] || ''),
    company: String(row[2] || ''),
    date: String(row[3] || ''),
    size: String(row[4] || '').trim().toUpperCase()
  };
}

function listRecords_() {
  const sh = getAttendanceSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, 5).getValues();
  return values
    .filter(r => String(r[1] || '').trim() !== '')
    .map(rowToRecord_);
}

function findRowByEmployee_(emp) {
  const sh = getAttendanceSheet_();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const nums = sh.getRange(2, 2, last - 1, 1).getValues();
  const target = String(emp).trim();
  for (let i = 0; i < nums.length; i++) {
    if (String(nums[i][0]).trim() === target) {
      return { sheet: sh, row: i + 2 };
    }
  }
  return null;
}

function doGet(e) {
  try {
    const action = (e.parameter.action || '').toLowerCase();
    if (action === 'config') {
      return jsonOut_({ success: true, data: readConfig_() });
    }
    if (action === 'list') {
      return jsonOut_({ success: true, data: listRecords_() });
    }
    if (action === 'search') {
      const q = String(e.parameter.query || '').trim();
      const all = listRecords_();
      const matches = all.filter(r => String(r.employeeNumber).slice(-4) === q);
      return jsonOut_({ success: true, data: matches });
    }
    return jsonOut_({ success: false, error: 'Unknown action' });
  } catch (err) {
    return jsonOut_({ success: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = String(body.action || '').toLowerCase();

    if (action === 'setconfig') {
      const data = writeConfig_({
        date: body.date,
        location: body.location,
        facilitator: body.facilitator
      });
      return jsonOut_({ success: true, data: data });
    }

    if (action === 'register') {
      const name = String(body.name || '').trim();
      const employeeNumber = String(body.employeeNumber || '').trim();
      const company = String(body.company || '').trim();
      const date = String(body.date || '').trim();
      // Size optional — moderator assigns later
      const size = String(body.size || '').trim().toUpperCase();

      if (!name || !employeeNumber || !company) {
        return jsonOut_({ success: false, error: 'Name, control number, and company are required.' });
      }

      const existing = findRowByEmployee_(employeeNumber);
      if (existing) {
        const rowVals = existing.sheet.getRange(existing.row, 1, 1, 5).getValues()[0];
        return jsonOut_({ success: true, duplicate: true, data: rowToRecord_(rowVals) });
      }

      getAttendanceSheet_().appendRow([name, employeeNumber, company, date, size]);
      return jsonOut_({
        success: true,
        duplicate: false,
        data: { name, employeeNumber, company, date, size }
      });
    }

    if (action === 'update') {
      const employeeNumber = String(body.employeeNumber || '').trim();
      const found = findRowByEmployee_(employeeNumber);
      if (!found) return jsonOut_({ success: false, error: 'Employee not found' });

      const sh = found.sheet;
      const row = found.row;
      if (body.name !== undefined) sh.getRange(row, 1).setValue(String(body.name).trim());
      if (body.company !== undefined) sh.getRange(row, 3).setValue(String(body.company).trim());
      if (body.date !== undefined) sh.getRange(row, 4).setValue(String(body.date).trim());
      if (body.size !== undefined) sh.getRange(row, 5).setValue(String(body.size).trim().toUpperCase());

      const rowVals = sh.getRange(row, 1, 1, 5).getValues()[0];
      return jsonOut_({ success: true, data: rowToRecord_(rowVals) });
    }

    if (action === 'delete') {
      const employeeNumber = String(body.employeeNumber || '').trim();
      const found = findRowByEmployee_(employeeNumber);
      if (!found) return jsonOut_({ success: false, error: 'Employee not found' });
      found.sheet.deleteRow(found.row);
      return jsonOut_({ success: true });
    }

    return jsonOut_({ success: false, error: 'Unknown action' });
  } catch (err) {
    return jsonOut_({ success: false, error: String(err) });
  }
}
