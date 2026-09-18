/**
 * ============================================================
 * PROJECT STRUCTURE
 * ============================================================
 *
 * This Apps Script project is split into separate files for clarity:
 *
 *   Downloader.gs
 *     - processIndoramaEmails()          ← trigger HOURLY
 *     - Saves email attachments / Darwinbox links into the folder
 *     - Moves processed emails to Trash
 *
 *   ProcessReconciliation.gs
 *     - processReconciliationFromFolder() ← trigger ONCE A DAY
 *     - Reads "Reconciliation Update.csv" from the folder
 *     - Fully replaces the Reconciliation sheet
 *
 *   ProcessEmployeeList.gs
 *     - processEmployeeListFromFolder()   ← trigger ONCE A DAY
 *     - Reads "Employee List.csv" from the folder
 *     - Fully replaces the Emp_Details sheet
 *
 *   Helpers.gs
 *     - Shared utility functions
 *
 * RECOMMENDED TRIGGERS
 * -------------------
 * 1. Hourly  → processIndoramaEmails
 * 2. Daily   → processReconciliationFromFolder
 * 3. Daily   → processEmployeeListFromFolder
 *    (you can schedule the two daily ones a few minutes apart)
 *
 * If old one-time triggers have accumulated, run:
 *   cleanupDownloaderTriggers()
 * once from the editor, then recreate the hourly trigger.
 */
