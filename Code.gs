/**
 * Main entry point.
 * Processes at most ONE thread per execution.
 * If more threads remain, schedules a one-time continuation trigger.
 */
function processIndoramaEmails() {
  const SENDER = 'olusegun.kehinde@indorama.com';
  const FOLDER_ID = '1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1';
  const MAX_THREADS_BEFORE_SPLIT = 2; // as requested: >2 threads → 1 per trigger

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    Logger.log('Another execution is already running. Exiting.');
    return;
  }

  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);

    // Search only recent non-trashed emails from the sender
    // (adjust newer_than: if you need older mail)
    const query = 'from:' + SENDER + ' -in:trash newer_than:14d';
    const threads = GmailApp.search(query, 0, 50);

    Logger.log('Threads found: ' + threads.length);

    if (threads.length === 0) {
      Logger.log('No threads to process.');
      return;
    }

    // ---- Process ONLY the first thread ----
    const thread = threads[0];
    processSingleThread(thread, SENDER, folder);

    // ---- If more than MAX_THREADS_BEFORE_SPLIT remain, schedule continuation(s) ----
    // Because we just processed (and possibly trashed) the first one,
    // the next search will see the remaining threads.
    if (threads.length > MAX_THREADS_BEFORE_SPLIT) {
      // Schedule one continuation. The next run will again take only the first
      // remaining thread and schedule another if needed. This keeps us at
      // 1 thread per trigger as requested.
      scheduleContinuationTrigger();
      Logger.log(
        'More than ' + MAX_THREADS_BEFORE_SPLIT +
        ' threads were present. Scheduled a continuation trigger for the next thread.'
      );
    } else if (threads.length > 1) {
      // 2 threads total → we already did the first; schedule one more for the last
      scheduleContinuationTrigger();
      Logger.log('One more thread remains. Scheduled continuation trigger.');
    }

  } finally {
    lock.releaseLock();
  }
}

/**
 * Process every message inside a single Gmail thread.
 */
function processSingleThread(thread, SENDER, folder) {
  const messages = thread.getMessages();
  Logger.log('Processing thread with ' + messages.length + ' message(s).');

  messages.forEach(function(message) {
    try {
      const messageSender = extractEmailAddress(message.getFrom());
      if (messageSender.toLowerCase() !== SENDER.toLowerCase()) {
        return;
      }

      const subject = message.getSubject().trim();
      Logger.log('Checking email: "' + subject + '"');

      // ---------- Darwinbox (contains match) ----------
      if (subject.toLowerCase().includes('darwinbox download')) {
        const success = processDarwinboxMessage(message, folder);
        if (success) {
          message.moveToTrash();
          Logger.log('Darwinbox email moved to Trash.');
        }
        return;
      }

      // ---------- Reconciliation (exact match) ----------
      if (subject.toLowerCase() === 'reconciliation') {
        const success = processReconciliationMessage(message);
        if (success) {
          message.moveToTrash();
          Logger.log('Reconciliation email moved to Trash.');
        }
        return;
      }

      // Other subjects – ignore
      Logger.log('Ignoring email with subject: ' + subject);

    } catch (err) {
      Logger.log('ERROR processing email: ' + err.toString());
      // leave message in place for retry
    }
  });
}

/**
 * Creates a one-time time-based trigger that will call processIndoramaEmails
 * again after a short delay. This realises "one thread per trigger".
 */
function scheduleContinuationTrigger() {
  // Clean any previous continuation triggers first to avoid accumulation
  const existing = ScriptApp.getProjectTriggers();
  existing.forEach(function(t) {
    if (t.getHandlerFunction() === 'processIndoramaEmails' &&
        t.getEventType() === ScriptApp.EventType.CLOCK) {
      // Keep the regular recurring trigger; only delete pure one-time ones
      // (one-time triggers have no recurrence). We simply allow a new one.
    }
  });

  ScriptApp.newTrigger('processIndoramaEmails')
    .timeBased()
    .after(2 * 60 * 1000) // 2 minutes
    .create();

  Logger.log('Continuation trigger created (fires in ~2 minutes).');
}

/**
 * Optional utility – run once from the editor if you ever need to
 * remove leftover one-time triggers.
 */
function cleanupOneTimeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processIndoramaEmails') {
      ScriptApp.deleteTrigger(t);
      Logger.log('Deleted trigger: ' + t.getUniqueId());
    }
  });
}


/* ==================================================================
 * DARWINBOX PROCESSOR (unchanged logic, only subject matching moved up)
 * ================================================================== */

function processDarwinboxMessage(message, folder) {
  try {
    const body = message.getPlainBody();

    const fileNameMatch = body.match(/FILE_NAME=(.*?)(\r?\n|$)/);
    const linkMatch = body.match(/DOWNLOAD_LINK=(.*?)(\r?\n|$)/);

    if (!fileNameMatch || !linkMatch) {
      Logger.log('Darwinbox email does not contain FILE_NAME or DOWNLOAD_LINK.');
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

    Logger.log('Processing Darwinbox file: ' + fileName);
    Logger.log('Actual Download Link: ' + downloadLink);

    const response = UrlFetchApp.fetch(downloadLink, {
      followRedirects: true,
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error('Download failed with code: ' + response.getResponseCode());
    }

    const blob = response.getBlob().setName(fileName);

    // Trash any existing file with the same name
    const existingFiles = folder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    folder.createFile(blob);
    Logger.log(fileName + ' saved successfully.');
    return true;

  } catch (err) {
    Logger.log('Darwinbox ERROR: ' + err.toString());
    return false;
  }
}


/* ==================================================================
 * RECONCILIATION EMAIL PROCESSOR
 * ================================================================== */

function processReconciliationMessage(message) {
  try {
    const attachments = message.getAttachments();
    Logger.log('Attachments found: ' + attachments.length);

    let csvAttachment = null;

    // Prefer exact name (case-insensitive)
    attachments.forEach(function(attachment) {
      const name = attachment.getName().trim().toLowerCase();
      if (name === 'reconciliation update.csv') {
        csvAttachment = attachment;
      }
    });

    // Fallback: any CSV
    if (!csvAttachment) {
      attachments.forEach(function(attachment) {
        const name = attachment.getName().trim().toLowerCase();
        if (!csvAttachment && name.endsWith('.csv')) {
          csvAttachment = attachment;
        }
      });
    }

    if (!csvAttachment) {
      Logger.log('No CSV attachment found in Reconciliation email.');
      return false;
    }

    Logger.log('Reading attachment: ' + csvAttachment.getName());

    const csvText = csvAttachment.getDataAsString('UTF-8');
    if (!csvText.trim()) {
      throw new Error('CSV attachment is empty.');
    }

    Logger.log('CSV size: ' + csvText.length + ' characters');

    const csvData = Utilities.parseCsv(csvText);
    if (!csvData || csvData.length < 2) {
      throw new Error('CSV does not contain any data rows.');
    }

    Logger.log('CSV rows including header: ' + csvData.length);

    const result = processReconciliationUpdate(csvData);
    Logger.log('Reconciliation processing result: ' + JSON.stringify(result));
    return true;

  } catch (err) {
    Logger.log('Reconciliation ERROR: ' + err.toString());
    return false;
  }
}


/* ==================================================================
 * RECONCILIATION GOOGLE SHEET PROCESSOR
 * Optimised for large existing data (~25k rows)
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
  const lastColumn = sheet.getLastColumn();

  let existingData = [];
  if (lastRow > 1 && lastColumn >= 3) {
    existingData = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  }

  Logger.log('Existing records: ' + existingData.length);

  // ---------- Build map of existing records ----------
  const records = new Map();

  for (let i = 0; i < existingData.length; i++) {
    const row = existingData[i];
    const employeeId = cleanValue(row[0]);
    const date = normalizeDate(row[1]);
    const status = cleanValue(row[2]);

    if (!employeeId || !date) continue;

    const key = createReconciliationKey(employeeId, date);
    records.set(key, [employeeId, date, status]);
  }

  // ---------- Apply incoming CSV (new data always wins) ----------
  let incomingCount = 0;

  for (let i = 1; i < csvData.length; i++) {
    const row = csvData[i];
    if (!row || row.length < 2) continue;

    const employeeId = cleanValue(row[0]);
    const date = normalizeDate(row[1]);
    const status = cleanValue(row[2]);

    if (!employeeId || !date) {
      // skip silently for speed on large CSVs
      continue;
    }

    const key = createReconciliationKey(employeeId, date);
    records.set(key, [employeeId, date, status]);
    incomingCount++;
  }

  Logger.log('Incoming records applied: ' + incomingCount);

  // ---------- Convert + sort ----------
  const finalData = Array.from(records.values());

  // Pre-normalise dates once for faster sort
  finalData.forEach(function(r) {
    r[1] = normalizeDate(r[1]); // already normalised, but ensure
  });

  finalData.sort(function(a, b) {
    const empCmp = String(a[0]).localeCompare(String(b[0]));
    if (empCmp !== 0) return empCmp;
    return String(a[1]).localeCompare(String(b[1]));
  });

  Logger.log('Final records to write: ' + finalData.length);

  // ---------- Write back (single bulk operation) ----------
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  }

  if (finalData.length > 0) {
    // Write in one go
    sheet.getRange(2, 1, finalData.length, 3).setValues(finalData);

    // Format date column
    sheet.getRange(2, 2, finalData.length, 1).setNumberFormat('M/d/yyyy');
  }

  SpreadsheetApp.flush();

  return {
    success: true,
    existingRecords: existingData.length,
    incomingRecords: incomingCount,
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

  // Already a Date object from Sheets
  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, 'UTC', 'yyyy-MM-dd');
  }

  const text = String(value).trim();
  if (!text) return '';

  // Fast path for ISO-like strings (most common from CSV)
  const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    return isoMatch[1] + '-' + isoMatch[2] + '-' + isoMatch[3];
  }

  // Fallback
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
