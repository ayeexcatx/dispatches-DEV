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
  const companyParam = String((e && e.parameter && e.parameter.company) || '').trim();
  const isDev = isDevEnvironment_();
  const user = token ? getActivePortalUserByToken_(token) : null;
  const shouldUseCompanyFallback = isDev && companyParam && (!token || !user);

  if (!shouldUseCompanyFallback) {
    if (!token) {
      return renderErrorPage_('Missing token. Please use your dispatch portal link.');
    }

    if (!user) {
      return renderErrorPage_('Invalid or inactive token. Please contact dispatch admin.');
    }
  }

  let content = '';
  let pageTitle = 'Dispatch List';

  if (!shouldUseCompanyFallback && isAdminOrDispatcherUser_(user)) {
    content = buildAdminDashboardHtml_();
    pageTitle = 'Admin Dispatch Dashboard';
    return HtmlService.createHtmlOutput(content)
      .setTitle(pageTitle)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (!shouldUseCompanyFallback && user.truckNumber) {
    content = buildTruckScopedPortalHtml_(user.truckNumber);
    pageTitle = `${user.truckNumber} Dispatch List`;
  } else {
    const resolvedCompany = shouldUseCompanyFallback ? companyParam : user.company;
    if (!resolvedCompany) {
      return renderErrorPage_('Token is missing truck/company scope. Please contact dispatch admin.');
    }

    const companyFileName = `${resolvedCompany}_dispatch_list.html`;
    const archivesFolder = DriveApp.getFolderById(DISPATCH_ARCHIVES_FOLDER_ID);
    const file = findFileRecursively(archivesFolder, companyFileName);

    if (!file) {
      return renderErrorPage_(`No dispatch list found for company: ${resolvedCompany}.`);
    }

    content = file.getBlob().getDataAsString();
    if (shouldUseCompanyFallback) {
      content = addDevFallbackBanner_(content);
    }
    pageTitle = `${resolvedCompany} Dispatch List`;
  }

  return HtmlService.createHtmlOutput(content)
    .setTitle(pageTitle)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Determine whether a user should see the admin/dispatcher dashboard.
 *
 * @param {{role:string}|null} user - Active portal user.
 * @returns {boolean} True when role is admin or dispatcher.
 */
function isAdminOrDispatcherUser_(user) {
  if (!user) return false;
  const role = String(user.role || '').trim().toLowerCase();
  return role === 'admin' || role === 'dispatcher';
}

/**
 * Build the Admin/Dispatcher dashboard HTML for active dispatches.
 *
 * @returns {string} Dashboard markup.
 */
function buildAdminDashboardHtml_() {
  const sheet = SpreadsheetApp.openById(DEV_DB_SPREADSHEET_ID).getSheetByName('Dispatches');
  if (!sheet || sheet.getLastRow() < 2) {
    return `<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>Admin Dispatch Dashboard</title></head>
<body><h1>Admin Dispatch Dashboard</h1><p>No active dispatches found.</p></body></html>`;
  }

  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(h => String(h || '').trim());
  const headerMap = {};
  headers.forEach((h, idx) => { if (h) headerMap[h] = idx; });

  const requiredHeaders = ['dispatch_id', 'date', 'truck_numbers', 'client', 'job_number', 'start_time', 'status', 'doc_url'];
  const missingHeader = requiredHeaders.find(name => headerMap[name] === undefined);
  if (missingHeader) {
    throw new Error(`Dispatches tab is missing required header: ${missingHeader}`);
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastColumn).getValues();
  const activeRows = rows.filter(row => String(row[headerMap.status] || '').trim() !== 'Completed');

  const tableRows = activeRows.map((row) => {
    const dispatchId = String(row[headerMap.dispatch_id] || '').trim();
    const dateValue = row[headerMap.date];
    const dateText = dateValue instanceof Date
      ? Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(dateValue || '').trim();
    const truckNumbers = String(row[headerMap.truck_numbers] || '').trim();
    const client = String(row[headerMap.client] || '').trim();
    const jobNumber = String(row[headerMap.job_number] || '').trim();
    const startTime = String(row[headerMap.start_time] || '').trim();
    const status = String(row[headerMap.status] || '').trim();
    const docUrl = String(row[headerMap.doc_url] || '').trim();

    const viewAction = docUrl
      ? `<button onclick="window.open('${docUrl}', '_blank')">View Dispatch</button>`
      : '<button disabled>View Dispatch</button>';

    return `<tr>
      <td>${dateText}</td>
      <td>${truckNumbers}</td>
      <td>${client}</td>
      <td>${jobNumber}</td>
      <td>${startTime}</td>
      <td>${status}</td>
      <td>
        ${viewAction}
        <button onclick="markCompleted('${dispatchId}')" ${dispatchId ? '' : 'disabled'}>Mark Completed</button>
        <button onclick="amendDispatch('${dispatchId}')" ${dispatchId ? '' : 'disabled'}>Amend</button>
        <button onclick="cancelDispatch('${dispatchId}')" ${dispatchId ? '' : 'disabled'}>Cancel</button>
      </td>
    </tr>`;
  }).join('\n');

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Admin Dispatch Dashboard</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 24px; background: #f6f8fb; color: #222; }
    h1 { margin-top: 0; }
    table { width: 100%; border-collapse: collapse; background: #fff; }
    th, td { border: 1px solid #d9d9d9; padding: 10px; text-align: left; vertical-align: top; }
    th { background: #eef3ff; }
    .actions button { margin-right: 8px; margin-bottom: 4px; }
  </style>
</head>
<body>
  <h1>Admin Dispatch Dashboard</h1>
  <h2>Active Dispatches</h2>
  <table>
    <thead>
      <tr>
        <th>Date</th>
        <th>Truck Numbers</th>
        <th>Client</th>
        <th>Job Number</th>
        <th>Start Time</th>
        <th>Status</th>
        <th>Actions</th>
      </tr>
    </thead>
    <tbody>
      ${tableRows || '<tr><td colspan="7">No active dispatches found.</td></tr>'}
    </tbody>
  </table>
  <script>
    function markCompleted(dispatchId) {
      if (!dispatchId) return;
      google.script.run
        .withSuccessHandler(function () { location.reload(); })
        .withFailureHandler(function (error) { alert('Mark complete failed: ' + (error && error.message ? error.message : error)); })
        .markDispatchCompleted(dispatchId);
    }

    function amendDispatch(dispatchId) {
      if (!dispatchId) return;
      google.script.run
        .withSuccessHandler(function () { location.reload(); })
        .withFailureHandler(function (error) { alert('Amend failed: ' + (error && error.message ? error.message : error)); })
        .amendDispatch(dispatchId);
    }

    function cancelDispatch(dispatchId) {
      if (!dispatchId) return;
      google.script.run
        .withSuccessHandler(function () { location.reload(); })
        .withFailureHandler(function (error) { alert('Cancel failed: ' + (error && error.message ? error.message : error)); })
        .cancelDispatch(dispatchId);
    }
  </script>
</body>
</html>`;
}

/**
 * Determine whether this script is running in DEV.
 *
 * @returns {boolean} True when ENV is set to DEV.
 */
function isDevEnvironment_() {
  return typeof ENV !== 'undefined' && String(ENV).toUpperCase() === 'DEV';
}

/**
 * Prepend a visible DEV banner when company fallback mode is used.
 *
 * @param {string} html - Existing page HTML.
 * @returns {string} HTML with a fallback banner inserted.
 */
function addDevFallbackBanner_(html) {
  const banner = '<div style="position:sticky;top:0;z-index:9999;background:#fff3cd;color:#856404;border:1px solid #ffeeba;padding:12px 16px;font-weight:bold;text-align:center;">DEV FALLBACK MODE (company param)</div>';
  const rawHtml = String(html || '');

  if (/<body[^>]*>/i.test(rawHtml)) {
    return rawHtml.replace(/<body([^>]*)>/i, `<body$1>${banner}`);
  }

  return `${banner}${rawHtml}`;
}

/**
 * Look up a Users row by token and only allow active rows.
 *
 * @param {string} token - Portal token from URL.
 * @returns {{userId:string,displayName:string,role:string,company:string,truckNumber:string}|null}
 */
function getActivePortalUserByToken_(token) {
  const ss = SpreadsheetApp.openById(DEV_DB_SPREADSHEET_ID);
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
  const sheet = SpreadsheetApp.openById(DEV_DB_SPREADSHEET_ID).getSheetByName('Dispatches');
  if (!sheet || sheet.getLastRow() < 2) {
    return renderNoDispatchesHtml_(truckNumber);
  }

  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(h => String(h || '').trim());
  const headerMap = {};
  headers.forEach((h, idx) => { if (h) headerMap[h] = idx; });

  const requiredHeaders = ['dispatch_id', 'date', 'start_time', 'job_number', 'truck_numbers', 'status', 'doc_url', 'is_confirmed'];
  const missingHeader = requiredHeaders.find(name => headerMap[name] === undefined);
  if (missingHeader) {
    throw new Error(`Dispatches tab is missing required header: ${missingHeader}`);
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastColumn).getValues();

  function stripTime(date) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  }

  const today = stripTime(new Date());
  const normalizedTruck = String(truckNumber || '').trim();
  const matchingDispatches = rows.filter((row) => {
    const truckNumbers = String(row[headerMap.truck_numbers] || '').trim();
    const truckTokens = truckNumbers.split(/[\s,]+/).map(t => t.trim()).filter(Boolean);
    return truckTokens.indexOf(normalizedTruck) !== -1;
  });

  if (!matchingDispatches.length) {
    return renderNoDispatchesHtml_(truckNumber);
  }

  const upcoming = [];
  const todayList = [];
  const past = [];

  matchingDispatches.forEach((row) => {
    const dispatchId = String(row[headerMap.dispatch_id] || '').trim();
    const dateValue = row[headerMap.date];
    const dateString = dateValue instanceof Date
      ? Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(dateValue || '').trim();

    const startTimeRaw = String(row[headerMap.start_time] || '').trim();
    const timeDigits = startTimeRaw.replace(/\D/g, '');
    const timePart = (timeDigits + '0000').slice(0, 4);
    const dispatchDate = dateString
      ? new Date(`${dateString}T${timePart.slice(0, 2)}:${timePart.slice(2)}:00`)
      : new Date(0);

    const friendlyDate = dispatchDate.toLocaleDateString('en-US', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit'
    });
    const friendlyTime = dispatchDate.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' });

    const jobNumber = String(row[headerMap.job_number] || '').trim();
    const status = String(row[headerMap.status] || '').trim();
    const docUrl = String(row[headerMap.doc_url] || '').trim();
    const rawIsConfirmed = row[headerMap.is_confirmed];
    const isConfirmed = rawIsConfirmed === true
      || String(rawIsConfirmed || '').toLowerCase() === 'true';

    let statusLabel = '';
    if (status === 'Canceled') {
      statusLabel = " <span style='color: #d93025; font-weight: bold;'>CANCELLED</span>";
    } else if (status === 'Amended') {
      statusLabel = " <span style='color: #FFBF00; font-weight: bold;'>AMENDMENT</span>";
    }

    const label = `${friendlyDate} @ ${friendlyTime} – ${normalizedTruck} – ${jobNumber}${statusLabel}`;
    const showConfirmButton = !isConfirmed;
    const confirmButton = !dispatchId
      ? `<button class="confirm-btn" disabled title="Dispatch index missing">Unavailable</button>`
      : showConfirmButton
        ? `<button class="confirm-btn" data-dispatch-id="${dispatchId}" data-truck-number="${normalizedTruck}" onclick="confirmReceipt(this)">Confirm Receipt</button>`
        : '<button class="confirm-btn" disabled>Confirmed ✓</button>';

    const link = `<div class="dispatch-block"><a href="${docUrl}">${label}</a>${confirmButton}</div>`;
    const entry = { date: dispatchDate, html: link };

    const dispatchDay = stripTime(dispatchDate);
    if (dispatchDay.getTime() < today.getTime()) {
      past.push(entry);
    } else if (dispatchDay.getTime() === today.getTime()) {
      todayList.push(entry);
    } else {
      upcoming.push(entry);
    }
  });

  upcoming.sort((a, b) => b.date - a.date);
  todayList.sort((a, b) => b.date - a.date);
  past.sort((a, b) => b.date - a.date);

  const upcomingHTML = upcoming.map(e => e.html).join('\n');
  const todayHTML = todayList.map(e => e.html).join('\n');
  const pastHTML = past.map(e => e.html).join('\n');

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>${normalizedTruck} Dispatches</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 40px; background: #f9f9f9; }
    .section {
      margin-top: 40px;
      padding: 20px;
      border-radius: 10px;
      background-color: #ffffff;
      box-shadow: 0 0 8px rgba(0,0,0,0.05);
    }
    .upcoming { border: 3px solid #4CAF50; }
    .today { border: 3px solid #2196F3; }
    .past { border: 3px solid #9E9E9E; }
    h2 {
      font-size: 20px;
      margin-top: 0;
    }
    .dispatch-block {
      padding: 10px 0;
      border-bottom: 1px solid #e0e0e0;
    }
    .dispatch-block a {
      color: #1a73e8;
      text-decoration: underline;
      font-weight: bold;
    }
    .confirm-btn {
      margin-left: 12px;
      vertical-align: middle;
      border: none;
      border-radius: 4px;
      padding: 4px 6px;
      font-size: 14px;
      font-weight: bold;
      cursor: pointer;
      background-color: #1a73e8;
      color: #fff;
    }
    .confirm-btn.confirmed,
    .confirm-btn:disabled {
      background-color: #d9ead3;
      color: #2e7d32;
      cursor: default;
    }
    .title-container {
      text-align: center;
      margin-bottom: 24px;
    }
    .dispatch-title {
      display: inline-block;
      background-color: #FFD700;
      padding: 12px 24px;
      border-radius: 999px;
      font-size: 2em;
      font-weight: bold;
      color: #000;
      box-shadow: 0 2px 6px rgba(0,0,0,0.15);
    }
  </style>
</head>
<body>
    <div class="title-container">
      <h1 class="dispatch-title">Dispatches — ${normalizedTruck}</h1>
    </div>

  <div class="section upcoming">
    <h2>Upcoming</h2>
    ${upcomingHTML || '<p>No upcoming dispatches.</p>'}
  </div>

  <div class="section today">
    <h2>Today</h2>
    ${todayHTML || '<p>No dispatches for today.</p>'}
  </div>

  <div class="section past">
    <h2>Past</h2>
    ${pastHTML || '<p>No past dispatches.</p>'}
  </div>

  <script>
    function confirmReceipt(buttonEl) {
      const dispatchId = buttonEl.getAttribute('data-dispatch-id');
      const truckNumber = buttonEl.getAttribute('data-truck-number');
      if (!dispatchId || !truckNumber) return;

      const originalLabel = buttonEl.textContent;
      buttonEl.disabled = true;
      buttonEl.textContent = 'Confirming...';

      google.script.run
        .withSuccessHandler(function () {
          buttonEl.textContent = 'Confirmed ✓';
        })
        .withFailureHandler(function (error) {
          buttonEl.disabled = false;
          buttonEl.textContent = originalLabel || 'Confirm Receipt';
          alert('Confirmation failed: ' + (error && error.message ? error.message : error));
        })
        .confirmDispatchReceipt(dispatchId, truckNumber);
    }
  </script>
</body>
</html>`;
}

function renderNoDispatchesHtml_(truckNumber) {
  const normalizedTruck = String(truckNumber || '').trim();
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>${normalizedTruck} Dispatches</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 40px; background: #f9f9f9; color: #333; }
    .card { max-width: 720px; margin: 60px auto; background: #fff; border: 1px solid #e0e0e0; border-radius: 12px; padding: 24px; }
  </style>
</head>
<body>
  <div class="card">
    <p>No dispatches found for ${normalizedTruck}</p>
  </div>
</body>
</html>`;
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
