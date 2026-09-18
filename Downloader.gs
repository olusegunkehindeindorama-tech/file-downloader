/**
 * ============================================================
 * FILE DOWNLOADER
 * ============================================================
 * Processes emails from the designated sender and saves files
 * into the Darwinbox folder.
 *
 * Rules:
 * 1. If the email has attachment(s) → save each attachment with
 *    its original name (existing same-name files are trashed).
 * 2. If no attachment but FILE_NAME + DOWNLOAD_LINK exist in the
 *    body → download and save with the name from the body.
 * 3. Successfully processed emails are moved to Trash.
 *
 * Trigger this function hourly (or as needed).
 * It processes a limited batch per run to stay under the time limit.
 */

function processIndoramaEmails() {
  const SENDER = 'olusegun.kehinde@indorama.com';
  const FOLDER_ID = '1DZ2MYPvTR1HMSVUIE3fcCIBVLyrBqxD1';
  const MAX_MESSAGES_PER_RUN = 12;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    Logger.log('Another downloader run is already in progress. Exiting.');
    return;
  }

  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
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
        }
      }
    }

    Logger.log('Messages successfully processed this run: ' + processedCount);

    if (hitLimit) {
      scheduleDownloaderContinuation();
      Logger.log('Batch limit reached. Scheduled one continuation trigger.');
    }

  } finally {
    lock.releaseLock();
  }
}

/**
 * Core downloader logic for a single message.
 * Returns true if the email was successfully handled.
 */
function processEmailAsDownloader(message, folder) {
  const attachments = message.getAttachments();

  // Case 1: has attachment(s)
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

  // Case 2: no attachment → try Darwinbox-style link
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
 * Creates one short-lived continuation trigger for the downloader.
 */
function scheduleDownloaderContinuation() {
  ScriptApp.newTrigger('processIndoramaEmails')
    .timeBased()
    .after(3 * 60 * 1000)
    .create();

  Logger.log('Downloader continuation trigger created (fires in ~3 minutes).');
}

/**
 * Utility: remove ALL triggers that call processIndoramaEmails.
 * Run once from the editor if triggers accumulate, then re-create
 * your regular hourly trigger.
 */
function cleanupDownloaderTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processIndoramaEmails') {
      ScriptApp.deleteTrigger(t);
      Logger.log('Deleted trigger: ' + t.getUniqueId());
    }
  });
}
