/**
 * ============================================================
 * RECONCILIATION PROCESSOR
 * ============================================================
 * Trigger this function once a day (separate from Employee List).
 *
 * Looks for "Reconciliation Update.csv" in the Darwinbox folder
 * and fully replaces the Reconciliation sheet with its content.
 * Any Emp + Date combination not present in the CSV is removed.
 */

function processReconciliationFromFolder() {
  const FOLDER_ID = '1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1';

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    Logger.log('Another Reconciliation run is already in progress. Exiting.');
    return;
  }

  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const reconFile = findFileInFolder(folder, 'reconciliation update.csv');

    if (!reconFile) {
      Logger.log('Reconciliation Update.csv not found in the folder.');
      return;
    }

    Logger.log('Found Reconciliation file: ' + reconFile.getName());

    const csvText = reconFile.getBlob().getDataAsString('UTF-8');
    if (!csvText.trim()) {
      Logger.log('Reconciliation CSV is empty.');
      return;
    }

    const csvData = Utilities.parseCsv(csvText);
    if (!csvData || csvData.length < 2) {
      Logger.log('Reconciliation CSV has no data rows.');
      return;
    }

    Logger.log('CSV rows including header: ' + csvData.length);

    const result = processReconciliationUpdate(csvData);
    Logger.log('Reconciliation result: ' + JSON.stringify(result));

  } catch (err) {
    Logger.log('Reconciliation ERROR: ' + err.toString());
  } finally {
    lock.releaseLock();
  }
}

/**
 * Full-replace the Reconciliation sheet using the provided CSV data.
 * CSV is the complete source of truth.
 */
function processReconciliationUpdate(csvData) {
  const SPREADSHEET_ID = '1Wmo3BU1ht3VCiNx5hn0-PmrT9_VHivsnFoe6tLFAa74';
  const SHEET_ID = 251798960;

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetById(SHEET_ID);

  if (!sheet) {
    throw new Error('Target sheet could not be found. GID: ' + SHEET_ID);
  }

  const lastRow = sheet.getLastRow();
  const existingCount = lastRow > 1 ? lastRow - 1 : 0;
  Logger.log('Existing reconciliation records before replace: ' + existingCount);

  const records = new Map();
  let incomingCount = 0;

  for (let i = 1; i < csvData.length; i++) {
    const row = csvData[i];
    if (!row || row.length < 2) continue;

    const employeeId = cleanValue(row[0]);
    const date = normalizeDate(row[1]);
    const status = cleanValue(row[2]);

    if (!employeeId || !date) continue;

    const key = createReconciliationKey(employeeId, date);
    records.set(key, [employeeId, date, status]);
    incomingCount++;
  }

  Logger.log('CSV records used: ' + incomingCount);

  const finalData = Array.from(records.values());
  finalData.sort(function(a, b) {
    const empCmp = String(a[0]).localeCompare(String(b[0]));
    if (empCmp !== 0) return empCmp;
    return String(a[1]).localeCompare(String(b[1]));
  });

  Logger.log('Final records to write: ' + finalData.length);

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  }

  if (finalData.length > 0) {
    sheet.getRange(2, 1, finalData.length, 3).setValues(finalData);
    sheet.getRange(2, 2, finalData.length, 1).setNumberFormat('M/d/yyyy');
  }

  SpreadsheetApp.flush();

  return {
    success: true,
    previousRecords: existingCount,
    csvRecords: incomingCount,
    finalRecords: finalData.length
  };
}
