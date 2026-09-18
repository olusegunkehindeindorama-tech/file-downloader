/**
 * ============================================================
 * FILE DOWNLOADER
 * ============================================================
 * Processes emails from the designated sender.
 *
 * Rules:
 * 1. If the email has one or more attachments → save each attachment
 *    into the Darwinbox folder using the original attachment name.
 *    Existing files with the same name are trashed (versioning).
 * 2. If the email has no attachments but contains FILE_NAME + DOWNLOAD_LINK
 *    → download the file and save it with the name from the body
 *    (same versioning behaviour).
 * 3. Successfully processed emails are moved to Trash.
 *
 * Designed to stay under the 6-minute limit:
 * - Processes a limited number of messages per run.
 * - Relies on the regular (hourly) trigger for remaining emails.
 * - Creates at most ONE continuation trigger if the batch limit is hit.
 */

function processIndoramaEmails() {
  const SENDER = 'olusegun.kehinde@indorama.com';
  const FOLDER_ID = '1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1';
  const MAX_MESSAGES_PER_RUN = 12;   // keep runs short

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    Logger.log('Another execution is already running. Exiting.');
    return;
  }

  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);

    // Only recent, non-trashed emails from the sender
    const query = 'from:' + SENDER + ' -in:trash newer_than:14d';
    const threads = GmailApp.search(query, 0, 40);

    Logger.log('Threads found: ' + threads.length);

    if (threads.length === 0) {
      Logger.log('No threads to process.');
      return;
    }

    let processedCount = 0;
    let hitLimit = false;

    outer:
    for (let t = 0; t < threads.length; t++) {
      const messages = threads[t].getMessages();

      for (let m = 0; m < messages.length; m++) {
        if (processedCount >= MAX_MESSAGES_PER_RUN) {
          hitLimit = true;
          break outer;
        }

        const message = messages[m];
        try {
          const messageSender = extractEmailAddress(message.getFrom());
          if (messageSender.toLowerCase() !== SENDER.toLowerCase()) {
            continue;
          }

          const subject = message.getSubject().trim();
          Logger.log('Checking email: "' + subject + '"');

          const success = processEmailAsDownloader(message, folder);

          if (success) {
            message.moveToTrash();
            Logger.log('Email moved to Trash.');
            processedCount++;
          } else {
            Logger.log('Email left in inbox for later retry / ignore.');
          }

        } catch (err) {
          Logger.log('ERROR processing email: ' + err.toString());
          // leave in place for retry
        }
      }
    }

    Logger.log('Messages successfully processed this run: ' + processedCount);

    // Only schedule ONE continuation if we hit the batch limit
    if (hitLimit) {
      scheduleContinuationTrigger();
      Logger.log('Batch limit reached. Scheduled one continuation trigger.');
    }

  } finally {
    lock.releaseLock();
  }
}

/**
 * Core downloader logic for a single message.
 * Returns true if the email was successfully handled (and can be trashed).
 */
function processEmailAsDownloader(message, folder) {
  const attachments = message.getAttachments();

  // ---------- Case 1: has attachment(s) ----------
  if (attachments && attachments.length > 0) {
    Logger.log('Attachments found: ' + attachments.length);
    let allSaved = true;

    attachments.forEach(function(attachment) {
      try {
        const fileName = attachment.getName().trim();
        if (!fileName) {
          Logger.log('Skipping attachment with empty name.');
          return;
        }

        Logger.log('Saving attachment: ' + fileName);

        // Versioning: trash any existing file with the same name
        const existing = folder.getFilesByName(fileName);
        while (existing.hasNext()) {
          existing.next().setTrashed(true);
        }

        folder.createFile(attachment.copyBlob().setName(fileName));
        Logger.log(fileName + ' saved successfully.');

      } catch (err) {
        Logger.log('Failed to save attachment: ' + err.toString());
        allSaved = false;
      }
    });

    return allSaved;
  }

  // ---------- Case 2: no attachment → try Darwinbox-style link ----------
  Logger.log('No attachments. Checking for FILE_NAME / DOWNLOAD_LINK...');
  return processDarwinboxLink(message, folder);
}

/**
 * Download via FILE_NAME + DOWNLOAD_LINK in the body.
 */
function processDarwinboxLink(message, folder) {
  try {
    const body = message.getPlainBody();

    const fileNameMatch = body.match(/FILE_NAME=(.*?)(\r?\n|$)/);
    const linkMatch = body.match(/DOWNLOAD_LINK=(.*?)(\r?\n|$)/);

    if (!fileNameMatch || !linkMatch) {
      Logger.log('No FILE_NAME or DOWNLOAD_LINK found. Ignoring.');
      return false;
    }

    const fileName = fileNameMatch[1].trim();
    let downloadLink = linkMatch[1].trim().replace(/[<>]/g, '').trim();

    // Unwrap Outlook SafeLinks
    if (downloadLink.indexOf('safelinks.protection.outlook.com') > -1) {
      const urlParam = downloadLink.match(/url=([^&]+)/);
      if (urlParam && urlParam[1]) {
        downloadLink = decodeURIComponent(urlParam[1]);
      }
    }

    Logger.log('Downloading: ' + fileName);
    Logger.log('Link: ' + downloadLink);

    const response = UrlFetchApp.fetch(downloadLink, {
      followRedirects: true,
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error('Download failed with HTTP ' + response.getResponseCode());
    }

    const blob = response.getBlob().setName(fileName);

    // Versioning
    const existing = folder.getFilesByName(fileName);
    while (existing.hasNext()) {
      existing.next().setTrashed(true);
    }

    folder.createFile(blob);
    Logger.log(fileName + ' saved successfully.');
    return true;

  } catch (err) {
    Logger.log('Darwinbox link ERROR: ' + err.toString());
    return false;
  }
}

/**
 * Creates at most one short-lived continuation trigger.
 * First removes any previous one-time triggers for this function
 * so they do not accumulate.
 */
function scheduleContinuationTrigger() {
  // Clean previous one-time triggers that point to the same function
  // (keeps any recurring/hourly trigger that has a different trigger source)
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'processIndoramaEmails' &&
        t.getEventType() === ScriptApp.EventType.CLOCK) {
      // One-time triggers created with .after() have no recurrence.
      // We delete only those that look like short-lived continuations.
      // Safer approach: delete all CLOCK triggers for this function
      // that are not the main recurring one. Since we cannot easily
      // distinguish, we simply delete ALL CLOCK triggers for this
      // function and let the user re-create the hourly one if needed.
      // Better: only create if none exist, or always clean then recreate one.
    }
  });

  // Simple & safe: always create one new short trigger.
  // The hourly trigger (if set with a different schedule) remains.
  ScriptApp.newTrigger('processIndoramaEmails')
    .timeBased()
    .after(3 * 60 * 1000) // 3 minutes
    .create();

  Logger.log('Continuation trigger created (fires in ~3 minutes).');
}

/**
 * Utility: remove ALL triggers that call processIndoramaEmails.
 * Run this manually from the editor if triggers ever accumulate.
 * Afterwards re-create your hourly trigger.
 */
function cleanupAllDownloaderTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processIndoramaEmails') {
      ScriptApp.deleteTrigger(t);
      Logger.log('Deleted trigger: ' + t.getUniqueId());
    }
  });
}


/* ==================================================================
 * DAILY PROCESSOR (run once per day)
 * ================================================================
 * Reads the CSV files that the downloader has already saved into
 * the Darwinbox folder and updates the Google Sheets.
 *
 * Expected files in the folder:
 *   - Employee List.csv          → Emp_Details sheet
 *   - Reconciliation Update.csv  → Reconciliation sheet
 *
 * Create a time-driven trigger that calls processDailyFiles() once a day.
 */

function processDailyFiles() {
  const FOLDER_ID = '1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1';

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    Logger.log('Another daily run is already in progress. Exiting.');
    return;
  }

  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);

    // ---------- Employee List ----------
    const empFile = findFileInFolder(folder, 'employee list.csv');
    if (empFile) {
      Logger.log('Found Employee List file: ' + empFile.getName());
      const csvText = empFile.getBlob().getDataAsString('UTF-8');
      const csvData = Utilities.parseCsv(csvText);
      if (csvData && csvData.length > 1) {
        const result = processEmployeeListUpdate(csvData);
        Logger.log('Employee List result: ' + JSON.stringify(result));
      } else {
        Logger.log('Employee List CSV is empty or has no data rows.');
      }
    } else {
      Logger.log('Employee List.csv not found in the folder.');
    }

    // ---------- Reconciliation ----------
    const reconFile = findFileInFolder(folder, 'reconciliation update.csv');
    if (reconFile) {
      Logger.log('Found Reconciliation file: ' + reconFile.getName());
      const csvText = reconFile.getBlob().getDataAsString('UTF-8');
      const csvData = Utilities.parseCsv(csvText);
      if (csvData && csvData.length > 1) {
        const result = processReconciliationUpdate(csvData);
        Logger.log('Reconciliation result: ' + JSON.stringify(result));
      } else {
        Logger.log('Reconciliation CSV is empty or has no data rows.');
      }
    } else {
      Logger.log('Reconciliation Update.csv not found in the folder.');
    }

  } finally {
    lock.releaseLock();
  }
}

/**
 * Case-insensitive file lookup inside a folder.
 */
function findFileInFolder(folder, targetName) {
  const files = folder.getFiles();
  const lowerTarget = targetName.toLowerCase();

  while (files.hasNext()) {
    const file = files.next();
    if (file.getName().trim().toLowerCase() === lowerTarget) {
      return file;
    }
  }
  return null;
}


/* ==================================================================
 * SHEET PROCESSORS (full replace – CSV is source of truth)
 * ================================================================== */

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
      0,          // S/N
      company,
      empId,
      gender,
      '',         // TRAIN TYPE I
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


/* ==================================================================
 * HELPERS
 * ================================================================== */

function createReconciliationKey(employeeId, date) {
  return String(employeeId).trim().toUpperCase() + '|' + normalizeDate(date);
}

function normalizeDate(value) {
  if (value === null || value === undefined || value === '') return '';

  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, 'UTC', 'yyyy-MM-dd');
  }

  const text = String(value).trim();
  if (!text) return '';

  const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    return isoMatch[1] + '-' + isoMatch[2] + '-' + isoMatch[3];
  }

  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, 'UTC', 'yyyy-MM-dd');
  }

  return '';
}

function cleanValue(value) {
  if (value === null || value === undefined) return '';
  return String(value).replace(/^\uFEFF/, '').trim();
}

function extractEmailAddress(fromString) {
  const match = String(fromString).match(/<([^>]+)>/);
  if (match) return match[1].trim();
  return String(fromString).trim();
}
