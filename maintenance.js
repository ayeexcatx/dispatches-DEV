/**
 * This file contains cleanup jobs for old dispatch files.
 *
 * In plain language, it goes through archive folders and removes old files to
 * keep recent history while avoiding clutter.
 */

/**
 * Move archived dispatch files older than 100 days to Google Drive trash.
 *
 * It checks each company folder, then each truck folder, then each file, and skips HTML pages.
 * It reads the date in each file name to decide whether the file is too old.
 */
function cleanupOldDispatches() {
  // Main archive folder that holds all company folders and truck subfolders.
  const parentFolder = DriveApp.getFolderById(DISPATCH_ARCHIVES_FOLDER_ID);
  const now = new Date();
  const cutoff = new Date(now.getTime() - 100 * 24 * 60 * 60 * 1000); // Date cutoff: anything older than this is considered old.

  // Step 1: loop through each company folder.
  const companyFolders = parentFolder.getFolders();
  while (companyFolders.hasNext()) {
    const companyFolder = companyFolders.next();
    // Step 2: inside that company, loop through each truck folder.
    const truckFolders = companyFolder.getFolders();
    while (truckFolders.hasNext()) {
      const truckFolder = truckFolders.next();
      // Step 3: inside that truck folder, loop through each file.
      const files = truckFolder.getFiles();
      while (files.hasNext()) {
        const file = files.next();
        const name = file.getName();

        // Skip .html files because those are company web pages, not dispatch docs.
        if (name.endsWith('.html')) continue;

        // File names start with date/time, so read that date to check age.
        const match = name.match(/^(\d{4}-\d{2}-\d{2})_(\d{4})_/);
        if (match) {
          const fileDate = new Date(match[1]);
          if (fileDate < cutoff) {
            Logger.log(`🗑️ Deleting ${name} from ${truckFolder.getName()}`);
            file.setTrashed(true); // Use trash so files can still be restored if needed.
          }
        }
      }
    }
  }
}

/**
 * Move old files (over 100 days) to trash in each truck's "Replaced Dispatches" folder.
 *
 * This keeps recent replacement history available without storing very old replacements forever.
 */
function deleteOldReplacedDispatches() {
  // Keep folder ID here too so this function works even if run by itself.
  const DISPATCH_ARCHIVES_FOLDER_ID = "1Fic0PvyH2B-Dq7P0hYQLsn0jB09qOWLE"; // ID for the main Dispatch Archives folder.
  const DAYS_TO_KEEP = 100;
  const now = new Date();
  const cutoffDate = new Date(now.getTime() - DAYS_TO_KEEP * 24 * 60 * 60 * 1000);

  // Open the main folder that contains all companies.
  const mainFolder = DriveApp.getFolderById(DISPATCH_ARCHIVES_FOLDER_ID);
  const companyFolders = mainFolder.getFolders();

  while (companyFolders.hasNext()) {
    const companyFolder = companyFolders.next();
    // Step 2: inside that company, loop through each truck folder.
    const truckFolders = companyFolder.getFolders();

    while (truckFolders.hasNext()) {
      const truckFolder = truckFolders.next();
      // Some truck folders have a subfolder named "Replaced Dispatches".
      const replacedFolderIterator = truckFolder.getFoldersByName("Replaced Dispatches");

      // If that subfolder is missing, skip this truck.
      if (!replacedFolderIterator.hasNext()) continue;

      const replacedFolder = replacedFolderIterator.next();
      const files = replacedFolder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        const created = file.getDateCreated();

        if (created < cutoffDate) {
          Logger.log(`Deleting old file: ${file.getName()} (created: ${created})`);
          file.setTrashed(true);
        }
      }
    }
  }
}
// Shared root folder ID used by cleanup jobs and web page lookup code.