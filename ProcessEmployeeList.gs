/**
 * ============================================================
 * EMPLOYEE LIST PROCESSOR
 * ============================================================
 * Trigger this function once a day (separate from Reconciliation).
 *
 * Looks for "Employee List.csv" in the Darwinbox folder and fully
 * replaces the Emp_Details sheet with its content.
 * Any employee not present in the CSV is removed.
 */

function processEmployeeListFromFolder() {
  const FOLDER_ID = '1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1';

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    Logger.log('Another Employee List run is already in progress. Exiting.');
    return;
  }

  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const empFile = findFileInFolder(folder, 'employee list.csv');

    if (!empFile) {
      Logger.log('Employee List.csv not found in the folder.');
      return;
    }

    Logger.log('Found Employee List file: ' + empFile.getName());

    const csvText = empFile.getBlob().getDataAsString('UTF-8');
    if (!csvText.trim()) {
      Logger.log('Employee List CSV is empty.');
      return;
    }

    const csvData = Utilities.parseCsv(csvText);
    if (!csvData || csvData.length < 2) {
      Logger.log('Employee List CSV has no data rows.');
      return;
    }

    Logger.log('CSV rows including header: ' + csvData.length);

    const result = processEmployeeListUpdate(csvData);
    Logger.log('Employee List result: ' + JSON.stringify(result));

  } catch (err) {
    Logger.log('Employee List ERROR: ' + err.toString());
  } finally {
    lock.releaseLock();
  }
}

/**
 * Full-replace the Emp_Details sheet using the provided CSV data.
 * CSV is the complete source of truth.
 *
 * Sheet columns:
 *   A: S/N | B: COMPANY | C: Emp ID | D: GENDER
 *   E: TRAIN TYPE I | F: Category | G: Emp Name
 *
 * CSV columns:
 *   [0] Emp ID | [1] Emp Name | [2] Category | [3] Company | [4] Gender
 */
function processEmployeeListUpdate(csvData) {
  const SPREADSHEET_ID = '1Wmo3BU1ht3VCiNx5hn0-PmrT9_VHivsnFoe6tLFAa74';
  const SHEET_ID = 342193858; // Emp_Details

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetById(SHEET_ID);

  if (!sheet) {
    throw new Error('Target sheet Emp_Details could not be found. GID: ' + SHEET_ID);
  }

  const lastRow = sheet.getLastRow();
  const previousCount = lastRow > 1 ? lastRow - 1 : 0;
  Logger.log('Existing Emp_Details records (will be replaced): ' + previousCount);

  const records = new Map();

  for (let i = 1; i < csvData.length; i++) {
    const row = csvData[i];
    if (!row || row.length < 5) continue;

    const empId    = cleanValue(row[0]);
    const empName  = cleanValue(row[1]);
    const category = cleanValue(row[2]);
    const company  = cleanValue(row[3]);
    const gender   = cleanValue(row[4]);

    if (!empId) continue;

    const key = empId.toUpperCase();
    records.set(key, [
      0,          // S/N (re-numbered later)
      company,
      empId,
      gender,
      '',         // TRAIN TYPE I (not in CSV)
      category,
      empName
    ]);
  }

  const finalData = Array.from(records.values());

  finalData.sort(function(a, b) {
    const compCmp = String(a[1]).localeCompare(String(b[1]));
    if (compCmp !== 0) return compCmp;
    return String(a[2]).localeCompare(String(b[2]));
  });

  for (let i = 0; i < finalData.length; i++) {
    finalData[i][0] = i + 1;
  }

  Logger.log('Final Emp_Details records to write: ' + finalData.length);

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 7).clearContent();
  }

  if (finalData.length > 0) {
    sheet.getRange(2, 1, finalData.length, 7).setValues(finalData);
  }

  SpreadsheetApp.flush();

  return {
    success: true,
    previousRecords: previousCount,
    finalRecords: finalData.length
  };
}
