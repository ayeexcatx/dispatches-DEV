/**
 * This file serves company dispatch web pages stored in Google Drive.
 */

// Main Drive folder where company dispatch HTML files are stored.
const DISPATCH_ARCHIVES_FOLDER_ID = '1Fic0PvyH2B-Dq7P0hYQLsn0jB09qOWLE';

/**
 * Handle a web request and return the matching scoped dispatch page.
 *
 * Expected URL format: ?t=TOKEN
 */
function doGet(e) {
  const token = String((e && e.parameter && e.parameter.t) || '').trim();
  if (!token) {
    return renderErrorPage_('Missing token. Please use your dispatch portal link.');
  }

  const user = getActivePortalUserByToken_(token);
  if (!user) {
    return renderErrorPage_('Invalid or inactive token. Please contact dispatch admin.');
  }

  let content = '';
  let pageTitle = 'Dispatch List';

  if (user.truckNumber) {
    content = buildTruckScopedPortalHtml_(user.truckNumber);
    pageTitle = `${user.truckNumber} Dispatch List`;
  } else if (user.company) {
    const companyFileName = `${user.company}_dispatch_list.html`;
    const archivesFolder = DriveApp.getFolderById(DISPATCH_ARCHIVES_FOLDER_ID);
    const file = findFileRecursively(archivesFolder, companyFileName);

    if (!file) {
      return renderErrorPage_(`No dispatch list found for company: ${user.company}.`);
    }

    content = file.getBlob().getDataAsString();
    pageTitle = `${user.company} Dispatch List`;
  } else {
    return renderErrorPage_('Token is missing truck/company scope. Please contact dispatch admin.');
  }

  return HtmlService.createHtmlOutput(content)
    .setTitle(pageTitle)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Look up a Users row by token and only allow active rows.
 *
 * @param {string} token - Portal token from URL.
 * @returns {{userId:string,displayName:string,role:string,company:string,truckNumber:string}|null}
 */
function getActivePortalUserByToken_(token) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');
  if (!usersSheet || usersSheet.getLastRow() < 2) return null;

  const lastColumn = usersSheet.getLastColumn();
  const headers = usersSheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(h => String(h || '').trim());
  const headerMap = {};
  headers.forEach((h, idx) => { if (h) headerMap[h] = idx; });

  const requiredHeaders = ['user_id', 'display_name', 'role', 'company', 'truck_number', 'token', 'is_active'];
  const missingHeader = requiredHeaders.find(name => headerMap[name] === undefined);
  if (missingHeader) {
    throw new Error(`Users tab is missing required header: ${missingHeader}`);
  }

  const rows = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, lastColumn).getValues();
  const row = rows.find((values) => {
    const rowToken = String(values[headerMap.token] || '').trim();
    if (rowToken !== token) return false;

    const activeRaw = values[headerMap.is_active];
    const isActive = activeRaw === true || String(activeRaw || '').toUpperCase() === 'TRUE';
    return isActive;
  });

  if (!row) return null;

  return {
    userId: String(row[headerMap.user_id] || '').trim(),
    displayName: String(row[headerMap.display_name] || '').trim(),
    role: String(row[headerMap.role] || '').trim(),
    company: String(row[headerMap.company] || '').trim(),
    truckNumber: String(row[headerMap.truck_number] || '').trim()
  };
}

/**
 * Build a truck-filtered portal page from the company dispatch HTML source.
 *
 * @param {string} truckNumber - Truck identity to keep in rendered blocks.
 * @returns {string} Filtered portal HTML.
 */
function buildTruckScopedPortalHtml_(truckNumber) {
  const archivesFolder = DriveApp.getFolderById(DISPATCH_ARCHIVES_FOLDER_ID);
  const companyFolderIterator = archivesFolder.getFolders();
  const truckToken = `_${truckNumber}_`;

  while (companyFolderIterator.hasNext()) {
    const companyFolder = companyFolderIterator.next();
    const htmlFiles = companyFolder.getFilesByName(`${companyFolder.getName()}_dispatch_list.html`);
    if (!htmlFiles.hasNext()) continue;

    const file = htmlFiles.next();
    const rawHtml = file.getBlob().getDataAsString();
    if (rawHtml.indexOf(truckToken) === -1 && rawHtml.indexOf(`data-truck-number="${truckNumber}"`) === -1) {
      continue;
    }

    return filterDispatchBlocksByTruck_(rawHtml, truckNumber, companyFolder.getName());
  }

  return renderErrorHtml_(`No dispatch list found for truck: ${truckNumber}.`);
}

/**
 * Keep only dispatch rows for a given truck and preserve confirmReceipt wiring.
 *
 * @param {string} html - Source portal HTML.
 * @param {string} truckNumber - Truck to preserve.
 * @param {string} companyName - Company label.
 * @returns {string} Filtered HTML.
 */
function filterDispatchBlocksByTruck_(html, truckNumber, companyName) {
  const sections = ['upcoming', 'today', 'past'];
  let output = html;

  sections.forEach((sectionClass) => {
    const sectionPattern = new RegExp(`(<div class="section ${sectionClass}">[\\s\\S]*?<h2>[\\s\\S]*?<\\/h2>)([\\s\\S]*?)(<\\/div>)`);
    const sectionMatch = output.match(sectionPattern);
    if (!sectionMatch) return;

    const prefix = sectionMatch[1];
    const body = sectionMatch[2];
    const suffix = sectionMatch[3];

    const blockPattern = /<div class="dispatch-block">[\s\S]*?<\/div>/g;
    const blocks = body.match(blockPattern) || [];
    const filtered = blocks.filter(block => block.indexOf(`data-truck-number="${truckNumber}"`) !== -1 || block.indexOf(`_${truckNumber}_`) !== -1);

    const replacementBody = filtered.length > 0 ? `\n${filtered.join('\n')}\n` : '\n<p>No dispatches in this section.</p>\n';
    output = output.replace(sectionPattern, `${prefix}${replacementBody}${suffix}`);
  });

  output = output.replace(/<h1 class="dispatch-title">[\s\S]*?<\/h1>/, `<h1 class="dispatch-title">${companyName} Dispatches — ${truckNumber}</h1>`);
  return output;
}

/**
 * Build a simple full-page HTML error response.
 *
 * @param {string} message - Human-friendly error text.
 * @returns {GoogleAppsScript.HTML.HtmlOutput} Error output.
 */
function renderErrorPage_(message) {
  return HtmlService.createHtmlOutput(renderErrorHtml_(message))
    .setTitle('Dispatch Portal Error')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Build error HTML.
 *
 * @param {string} message - Human-friendly error text.
 * @returns {string} HTML content.
 */
function renderErrorHtml_(message) {
  const safeMessage = String(message || 'Unknown portal error.');
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Dispatch Portal Error</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 40px; background: #f9f9f9; color: #333; }
    .card { max-width: 720px; margin: 60px auto; background: #fff; border: 1px solid #e0e0e0; border-radius: 12px; padding: 24px; }
    h1 { margin-top: 0; color: #c62828; }
  </style>
</head>
<body>
  <div class="card">
    <h1>Unable to open dispatch portal</h1>
    <p>${safeMessage}</p>
  </div>
</body>
</html>`;
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
