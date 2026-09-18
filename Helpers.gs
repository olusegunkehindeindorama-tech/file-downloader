/**
 * ============================================================
 * SHARED HELPERS
 * ============================================================
 * Used by Downloader, Reconciliation and Employee List scripts.
 */

/**
 * Case-insensitive file lookup inside a Drive folder.
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
