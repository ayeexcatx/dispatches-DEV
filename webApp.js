/**
 * This file serves company dispatch web pages stored in Google Drive.
 */

// Main Drive folder where company dispatch HTML files are stored.
const DISPATCH_ARCHIVES_FOLDER_ID = '1Fic0PvyH2B-Dq7P0hYQLsn0jB09qOWLE';

/**
 * Handle a web request and return the matching company dispatch page.
 *
 * Expected URL format: ?company=Company_Name_With_Underscores
 */
function doGet(e) {
  // Read the company name from the URL and make sure it was provided.
  const companyRaw = e.parameter.company;
  if (!companyRaw) {
    return ContentService.createTextOutput('Error: Missing company parameter.');
  }

  // Use that company value to build the expected HTML file name.
  const company = companyRaw;
  const fileName = `${company}_dispatch_list.html`;
  // Start searching from the top archive folder, including nested folders.
  const archivesFolder = DriveApp.getFolderById(DISPATCH_ARCHIVES_FOLDER_ID);
  const file = findFileRecursively(archivesFolder, fileName);

  if (!file) {
    Logger.log(`❌ File ${fileName} not found`);
    return ContentService.createTextOutput(`No dispatch list found for company: ${company}.`);
  }

  // Read the HTML file text and send it back as the web page response.
  const blob = file.getBlob();
  const content = blob.getDataAsString();

  Logger.log(`📦 File size: ${blob.getBytes().length} bytes`);
  Logger.log(`🧪 First 1000 characters: ${content.slice(0, 1000)}`);

  return HtmlService.createHtmlOutput(content)
    .setTitle(`${company} Dispatch List`)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Search folder-by-folder until we find the exact file name.
 *
 * @param {GoogleAppsScript.Drive.Folder} folder - Folder currently being checked.
 * @param {string} targetFileName - Exact file name we are looking for.
 * @returns {GoogleAppsScript.Drive.File|null} The file if found, otherwise null.
 */
function findFileRecursively(folder, targetFileName) {
  // First check files in this folder.
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getName() === targetFileName) {
      return file;
    }
  }
  // If not found, go into each subfolder and keep searching.
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    const result = findFileRecursively(subfolder, targetFileName);
    if (result) return result;
  }
  return null;
}
