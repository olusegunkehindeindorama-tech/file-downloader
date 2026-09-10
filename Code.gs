

function processIndoramaEmails() {

  const SENDER = 'olusegun.kehinde@indorama.com';

  const folder = DriveApp.getFolderById(
    '1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1'
  );

  /*
   * Get ALL non-trashed emails from the sender.
   *
   * We intentionally do NOT put a subject filter here because
   * we want one function to decide what to do based on subject.
   */
  const threads = GmailApp.search(
    'from:' + SENDER + ' -in:trash'
  );

  Logger.log('Threads found: ' + threads.length);


  threads.forEach(function(thread) {

    const messages = thread.getMessages();

    messages.forEach(function(message) {

      try {

        /*
         * Gmail thread search can return a thread containing
         * messages from other people, so make sure the individual
         * message is actually from our expected sender.
         */
        const messageSender = extractEmailAddress(
          message.getFrom()
        );

        if (
          messageSender.toLowerCase() !==
          SENDER.toLowerCase()
        ) {
          return;
        }


        const subject = message
          .getSubject()
          .trim();


        Logger.log(
          'Checking email: "' + subject + '"'
        );


        /*
         * ======================================================
         * DARWINBOX
         * ======================================================
         */

        if (
          subject.toLowerCase() ===
          'darwinbox download'.toLowerCase()
        ) {

          const success =
            processDarwinboxMessage(
              message,
              folder
            );

          if (success) {
            message.moveToTrash();
            Logger.log(
              'Darwinbox email moved to Trash.'
            );
          }

          return;
        }


        /*
         * ======================================================
         * RECONCILIATION
         * ======================================================
         */

        if (
          subject.toLowerCase() ===
          'reconciliation'.toLowerCase()
        ) {

          const success =
            processReconciliationMessage(
              message
            );

          if (success) {
            message.moveToTrash();
            Logger.log(
              'Reconciliation email moved to Trash.'
            );
          }

          return;
        }


        /*
         * ======================================================
         * OTHER SUBJECT
         * ======================================================
         *
         * Do nothing.
         */

        Logger.log(
          'Ignoring email with subject: ' +
          subject
        );

      } catch (err) {

        /*
         * IMPORTANT:
         * If processing fails, the email is NOT moved to Trash.
         *
         * This allows you to investigate/retry it.
         */

        Logger.log(
          'ERROR processing email: ' +
          err.toString()
        );

      }

    });

  });

}


/**
 * ============================================================
 * DARWINBOX PROCESSOR
 * ============================================================
 *
 * This is essentially your existing Darwinbox logic.
 */

function processDarwinboxMessage(message, folder) {

  try {

    const body = message.getPlainBody();


    /*
     * Extract file name
     */
    const fileNameMatch = body.match(
      /FILE_NAME=(.*?)(\r?\n|$)/
    );


    /*
     * Extract download link
     */
    const linkMatch = body.match(
      /DOWNLOAD_LINK=(.*?)(\r?\n|$)/
    );


    if (!fileNameMatch || !linkMatch) {

      Logger.log(
        'Darwinbox email does not contain FILE_NAME or DOWNLOAD_LINK.'
      );

      return false;

    }


    const fileName =
      fileNameMatch[1].trim();


    const rawLink =
      linkMatch[1].trim();


    /*
     * Remove < >
     */
    let downloadLink =
      rawLink
        .replace(/[<>]/g, '')
        .trim();


    /*
     * Unwrap Outlook SafeLinks
     */
    if (
      downloadLink.indexOf(
        'safelinks.protection.outlook.com'
      ) > -1
    ) {

      const urlParam =
        downloadLink.match(
          /url=([^&]+)/
        );

      if (
        urlParam &&
        urlParam[1]
      ) {

        downloadLink =
          decodeURIComponent(
            urlParam[1]
          );

      }

    }


    Logger.log(
      'Processing Darwinbox file: ' +
      fileName
    );

    Logger.log(
      'Actual Download Link: ' +
      downloadLink
    );


    /*
     * Download file
     */
    const response =
      UrlFetchApp.fetch(
        downloadLink,
        {
          followRedirects: true,
          muteHttpExceptions: true
        }
      );


    if (
      response.getResponseCode() !== 200
    ) {

      throw new Error(
        'Download failed with code: ' +
        response.getResponseCode()
      );

    }


    const blob =
      response
        .getBlob()
        .setName(fileName);


    /*
     * Delete/Trash existing version
     */
    const existingFiles =
      folder.getFilesByName(fileName);


    while (
      existingFiles.hasNext()
    ) {

      existingFiles
        .next()
        .setTrashed(true);

    }


    /*
     * Save new file
     */
    folder.createFile(blob);


    Logger.log(
      fileName +
      ' saved successfully.'
    );


    return true;


  } catch (err) {

    Logger.log(
      'Darwinbox ERROR: ' +
      err.toString()
    );

    return false;

  }

}


/**
 * ============================================================
 * RECONCILIATION EMAIL PROCESSOR
 * ============================================================
 *
 * Finds the attachment:
 *
 * Reconciliation Update.csv
 *
 * Reads the CSV directly from the email.
 *
 * It does NOT save the CSV to Google Drive.
 */

function processReconciliationMessage(message) {

  try {

    const attachments =
      message.getAttachments();


    Logger.log(
      'Attachments found: ' +
      attachments.length
    );


    let csvAttachment = null;


    /*
     * Find the exact attachment.
     *
     * Case-insensitive comparison is used.
     */
    attachments.forEach(function(attachment) {

      const name =
        attachment
          .getName()
          .trim()
          .toLowerCase();


      if (
        name ===
        'reconciliation update.csv'
      ) {

        csvAttachment =
          attachment;

      }

    });


    /*
     * If exact name wasn't found, look for any CSV
     * as a fallback.
     */
    if (!csvAttachment) {

      attachments.forEach(function(attachment) {

        const name =
          attachment
            .getName()
            .trim()
            .toLowerCase();


        if (
          !csvAttachment &&
          name.endsWith('.csv')
        ) {

          csvAttachment =
            attachment;

        }

      });

    }


    if (!csvAttachment) {

      Logger.log(
        'No CSV attachment found in Reconciliation email.'
      );

      return false;

    }


    Logger.log(
      'Reading attachment: ' +
      csvAttachment.getName()
    );


    /*
     * ========================================================
     * GET CSV CONTENT
     * ========================================================
     */

    const csvText =
      csvAttachment.getDataAsString(
        'UTF-8'
      );


    if (!csvText.trim()) {

      throw new Error(
        'CSV attachment is empty.'
      );

    }


    Logger.log(
      'CSV size: ' +
      csvText.length +
      ' characters'
    );


    /*
     * ========================================================
     * PARSE CSV
     * ========================================================
     */

    const csvData =
      Utilities.parseCsv(
        csvText
      );


    if (
      !csvData ||
      csvData.length < 2
    ) {

      throw new Error(
        'CSV does not contain any data rows.'
      );

    }


    Logger.log(
      'CSV rows including header: ' +
      csvData.length
    );


    /*
     * ========================================================
     * SEND DATA TO RECONCILIATION PROCESSOR
     * ========================================================
     */

    const result =
      processReconciliationUpdate(
        csvData
      );


    Logger.log(
      'Reconciliation processing result: ' +
      JSON.stringify(result)
    );


    return true;


  } catch (err) {

    Logger.log(
      'Reconciliation ERROR: ' +
      err.toString()
    );

    return false;

  }

}


/**
 * ============================================================
 * RECONCILIATION GOOGLE SHEET PROCESSOR
 * ============================================================
 *
 * Target spreadsheet:
 *
 * 1Wmo3BU1ht3VCiNx5hn0-PmrT9_VHivsnFoe6tLFAa74
 *
 * Duplicate key:
 *
 * Emp_No + Absent_Date
 *
 *
 * Incoming CSV columns:
 *
 * Emp No
 * Tr Date
 * Reconciliation
 *
 *
 * Google Sheet columns:
 *
 * Emp_No
 * Absent_Date
 * Reconciliation_Status
 */

function processReconciliationUpdate(csvData) {

  /*
   * ==========================================================
   * CONFIGURATION
   * ==========================================================
   */

  const SPREADSHEET_ID =
    '1Wmo3BU1ht3VCiNx5hn0-PmrT9_VHivsnFoe6tLFAa74';


  const SHEET_ID =
    251798960;


  /*
   * ==========================================================
   * OPEN SHEET
   * ==========================================================
   */

  const spreadsheet =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );


  const sheet =
    spreadsheet.getSheetById(
      SHEET_ID
    );


  if (!sheet) {

    throw new Error(
      'Target sheet could not be found. GID: ' +
      SHEET_ID
    );

  }


  /*
   * ==========================================================
   * READ EXISTING DATA INTO MEMORY
   * ==========================================================
   */

  const lastRow =
    sheet.getLastRow();


  const lastColumn =
    sheet.getLastColumn();


  let existingData = [];


  if (
    lastRow > 1 &&
    lastColumn >= 3
  ) {

    existingData =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          3
        )
        .getValues();

  }


  Logger.log(
    'Existing records: ' +
    existingData.length
  );


  /*
   * ==========================================================
   * BUILD EXISTING RECORD MAP
   * ==========================================================
   *
   * Key:
   *
   * Employee ID + Date
   *
   * Example:
   *
   * FDT20092|2026-07-05
   */

  const records =
    new Map();


  existingData.forEach(function(row) {

    const employeeId =
      cleanValue(row[0]);


    const date =
      normalizeDate(row[1]);


    const status =
      cleanValue(row[2]);


    /*
     * Ignore incomplete records.
     */
    if (
      !employeeId ||
      !date
    ) {

      return;

    }


    const key =
      createReconciliationKey(
        employeeId,
        date
      );


    /*
     * Existing duplicate records are automatically
     * collapsed because Map only stores one value
     * per key.
     */
    records.set(
      key,
      [
        employeeId,
        date,
        status
      ]
    );

  });


  /*
   * ==========================================================
   * PROCESS INCOMING CSV
   * ==========================================================
   *
   * CSV structure:
   *
   * Row 0 = headers
   *
   * Emp No
   * Tr Date
   * Reconciliation
   *
   */

  let incomingCount = 0;


  /*
   * Skip header row.
   */
  for (
    let i = 1;
    i < csvData.length;
    i++
  ) {

    const row =
      csvData[i];


    if (
      !row ||
      row.length < 2
    ) {

      continue;

    }


    /*
     * CSV columns:
     *
     * [0] Emp No
     * [1] Tr Date
     * [2] Reconciliation
     */

    const employeeId =
      cleanValue(row[0]);


    const date =
      normalizeDate(row[1]);


    const status =
      cleanValue(row[2]);


    /*
     * Skip invalid rows.
     */
    if (
      !employeeId ||
      !date
    ) {

      Logger.log(
        'Skipping invalid CSV row ' +
        (i + 1)
      );

      continue;

    }


    const key =
      createReconciliationKey(
        employeeId,
        date
      );


    /*
     * IMPORTANT:
     *
     * If this key already exists,
     * this incoming record replaces it.
     *
     * If it doesn't exist,
     * it is added.
     *
     * Therefore:
     *
     * NEW DATA ALWAYS WINS.
     */
    records.set(
      key,
      [
        employeeId,
        date,
        status
      ]
    );


    incomingCount++;

  }


  /*
   * ==========================================================
   * CONVERT MAP BACK TO ARRAY
   * ==========================================================
   */

  const finalData =
    Array.from(
      records.values()
    );


  /*
   * ==========================================================
   * SORT
   * ==========================================================
   *
   * Employee first,
   * then date.
   */

  finalData.sort(
    function(a, b) {

      const employeeCompare =
        String(a[0]).localeCompare(
          String(b[0])
        );


      if (
        employeeCompare !== 0
      ) {

        return employeeCompare;

      }


      return normalizeDate(a[1])
        .localeCompare(
          normalizeDate(b[1])
        );

    }
  );


  /*
   * ==========================================================
   * WRITE BACK TO GOOGLE SHEET
   * ==========================================================
   *
   * Header remains untouched.
   *
   * Existing data is cleared.
   *
   * Complete final dataset is written in ONE operation.
   */

  if (lastRow > 1) {

    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        3
      )
      .clearContent();

  }


  if (
    finalData.length > 0
  ) {

    sheet
      .getRange(
        2,
        1,
        finalData.length,
        3
      )
      .setValues(
        finalData
      );

  }


  /*
   * ==========================================================
   * FORMAT DATE COLUMN
   * ==========================================================
   */

  if (
    finalData.length > 0
  ) {

    sheet
      .getRange(
        2,
        2,
        finalData.length,
        1
      )
      .setNumberFormat(
        'M/d/yyyy'
      );

  }


  SpreadsheetApp.flush();


  /*
   * ==========================================================
   * RETURN RESULT
   * ==========================================================
   */

  return {

    success: true,

    existingRecords:
      existingData.length,

    incomingRecords:
      incomingCount,

    finalRecords:
      finalData.length

  };

}


/**
 * ============================================================
 * CREATE RECONCILIATION KEY
 * ============================================================
 */

function createReconciliationKey(
  employeeId,
  date
) {

  return (
    String(employeeId)
      .trim()
      .toUpperCase()
    +
    '|'
    +
    normalizeDate(date)
  );

}


/**
 * ============================================================
 * DATE NORMALIZATION
 * ============================================================
 *
 * Converts:
 *
 * 2026-07-05T00:00:00
 *
 * to:
 *
 * 2026-07-05
 *
 *
 * It also handles Google Sheets Date objects.
 */

function normalizeDate(value) {

  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {

    return '';

  }


  /*
   * Google Sheets Date object
   */
  if (
    Object.prototype.toString.call(value)
    === '[object Date]'
  ) {

    if (
      isNaN(value.getTime())
    ) {

      return '';

    }


    return Utilities.formatDate(
      value,
      'UTC',
      'yyyy-MM-dd'
    );

  }


  const text =
    String(value).trim();


  if (!text) {

    return '';

  }


  /*
   * ISO / UTC timestamp.
   *
   * This is the important part for your CSV.
   *
   * 2026-07-05T00:00:00
   *
   * becomes:
   *
   * 2026-07-05
   */

  const isoMatch =
    text.match(
      /^(\d{4})-(\d{2})-(\d{2})/
    );


  if (isoMatch) {

    return (
      isoMatch[1] +
      '-' +
      isoMatch[2] +
      '-' +
      isoMatch[3]
    );

  }


  /*
   * Normal date strings such as:
   *
   * 7/5/2026
   *
   * 07/05/2026
   */

  const parsed =
    new Date(text);


  if (
    !isNaN(parsed.getTime())
  ) {

    return Utilities.formatDate(
      parsed,
      'UTC',
      'yyyy-MM-dd'
    );

  }


  return '';

}


/**
 * ============================================================
 * CLEAN VALUE
 * ============================================================
 */

function cleanValue(value) {

  if (
    value === null ||
    value === undefined
  ) {

    return '';

  }


  /*
   * Remove BOM that sometimes appears in CSV headers/data.
   */
  return String(value)
    .replace(/^\uFEFF/, '')
    .trim();

}


/**
 * ============================================================
 * EXTRACT EMAIL ADDRESS
 * ============================================================
 *
 * Converts:
 *
 * Olusegun Kehinde <olusegun.kehinde@indorama.com>
 *
 * into:
 *
 * olusegun.kehinde@indorama.com
 */

function extractEmailAddress(fromString) {

  const match =
    String(fromString).match(
      /<([^>]+)>/
    );


  if (match) {

    return match[1].trim();

  }


  return String(fromString).trim();

}
