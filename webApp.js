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
  const pageParam = String((e && e.parameter && e.parameter.p) || '').trim().toLowerCase();
  try {
    if (pageParam === 'ping') {
      return renderPingPage_(token, pageParam);
    }

    const companyParam = String((e && e.parameter && e.parameter.company) || '').trim();
    const isDev = isDevEnvironment_();
    const user = token ? getActivePortalUserByToken_(token) : null;
    const shouldUseCompanyFallback = isDev && companyParam && (!token || !user);

    if (!shouldUseCompanyFallback) {
      if (!token) {
        return renderErrorPage_('Invalid or missing token. Please use your dispatch portal link.');
      }

      if (!user) {
        return renderErrorPage_('Invalid or missing token. Please contact dispatch admin.');
      }
    }

    let content = '';
    let pageTitle = 'Dispatch List';

    if (!shouldUseCompanyFallback && isAdminOrDispatcherUser_(user)) {
      if (pageParam === 'repair') {
        repairDispatchDocLinks_DEV_();
        return HtmlService.createHtmlOutput('<script>window.location.replace("?t=' + encodeURIComponent(token) + '&p=dashboard&msg=repair_done");</script>')
          .setTitle('Repairing Dispatch Links')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      let resolvedPage = pageParam || 'dashboard';
      if (resolvedPage === 'create-dispatch') resolvedPage = 'create';
      if (resolvedPage === 'companies-trucks') resolvedPage = 'companies';
      const allowedPages = { dashboard: true, create: true, companies: true, users: true };
      if (!allowedPages[resolvedPage]) {
        return renderAdminFallbackPage_(token, user, pageParam, new Error('Unknown admin page parameter'));
      }

      try {
        content = buildAdminHtml_({
          user: user,
          token: token,
          params: e && e.parameter,
          forcedPage: resolvedPage
        });
        Logger.log('[DEV] admin render p=' + resolvedPage + ' htmlLen=' + (content ? content.length : 0));
        if (!content || String(content).trim().length < 50) {
          throw new Error('Admin HTML empty/too short');
        }
      } catch (adminError) {
        return renderAdminFallbackPage_(token, user, pageParam, adminError);
      }
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
  } catch (error) {
    let user = null;
    try {
      user = token ? getActivePortalUserByToken_(token) : null;
    } catch (userLookupError) {
      user = null;
    }
    return renderAdminFallbackPage_(token, user, pageParam, error);
  }
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
 * Build the Admin/Dispatcher dashboard HTML as a single hardcoded string.
 *
 * @param {{user:Object,token:string,params:Object,forcedPage:string}} opts - Render options.
 * @returns {string} Dashboard markup.
 */
function buildAdminHtml_(opts) {
  const user = opts && opts.user ? opts.user : null;
  const token = opts && opts.token ? opts.token : '';
  const params = opts && opts.params ? opts.params : null;
  const forcedPage = opts && opts.forcedPage ? opts.forcedPage : '';
  const messageParam = String((params && params.msg) || '').trim().toLowerCase();
  const rawErrorParam = String((params && params.err) || '').trim();
  const errorParam = rawErrorParam ? decodeURIComponent(rawErrorParam) : '';
  const pageParamRaw = String((params && params.p) || '').trim().toLowerCase();
  const pageAliases = {
    'create-dispatch': 'create',
    'companies-trucks': 'companies'
  };
  const page = forcedPage || pageAliases[pageParamRaw] || pageParamRaw || 'dashboard';

  const notices = {
    completed_ok: { kind: 'success', text: 'Dispatch marked completed.' },
    amend_ok: { kind: 'success', text: 'Dispatch amended.' },
    cancel_ok: { kind: 'success', text: 'Dispatch canceled.' },
    create_ok: { kind: 'success', text: 'Dispatch created successfully.' },
    repair_done: { kind: 'success', text: 'Dispatch doc links repair completed.' }
  };
  const notice = notices[messageParam] || null;
  const renderId = Utilities.getUuid().slice(0, 8);
  const renderTimestamp = new Date().toISOString();
  const debugRole = String((user && user.role) || '').trim() || 'unknown';
  const debugTokenPresent = token ? 'yes' : 'no';

  const navItems = [
    { id: 'dashboard', label: 'Dashboard' },
    { id: 'create', label: 'Create Dispatch' },
    { id: 'companies', label: 'Companies/Trucks' },
    { id: 'users', label: 'Users' }
  ];

  const navHtml = navItems.map((item) => {
    const activeClass = item.id === page ? 'active' : '';
    const paramsForLink = ['t=' + encodeURIComponent(token || '')];
    paramsForLink.push('p=' + encodeURIComponent(item.id));
    const href = '?' + paramsForLink.join('&');
    return `<a class="tab ${activeClass}" href="${href}">${item.label}</a>`;
  }).join('\n');

  let pageContent = '';

  if (page === 'create') {
    pageContent = `
      <h2>create</h2>
      <section class="card page-shell">
        <p>Create Dispatch page.</p>
        <form id="createDispatchForm" onsubmit="submitCreateDispatch(event)">
        <label>Date
          <input type="date" name="date" required>
        </label>
        <label>Shift
          <input type="text" name="shift" placeholder="e.g. 6:00 AM" required>
        </label>
        <label>Client
          <input type="text" name="client" required>
        </label>
        <label>Job Number
          <input type="text" name="jobNumber" required>
        </label>
        <label>Start Time
          <input type="text" name="startTime" placeholder="e.g. 7:30 AM" required>
        </label>
        <label>Start Location
          <input type="text" name="startLocation" required>
        </label>
        <label>Instructions
          <textarea name="instructions" rows="4" required></textarea>
        </label>
        <label>Truck Numbers (comma-separated)
          <input type="text" name="truckNumbers" placeholder="RT03, RT12" required>
        </label>
        <div class="form-actions">
          <button type="submit">Create Dispatch</button>
        </div>
        </form>
      </section>
    `;
  } else if (page === 'companies') {
    pageContent = '<h2>companies</h2><section class="card page-shell"><p>Companies/Trucks placeholder content.</p></section>';
  } else if (page === 'users') {
    pageContent = '<h2>users</h2><section class="card page-shell"><p>Users placeholder content.</p></section>';
  } else {
    pageContent = `
      <h2 id="dashboardTitle">dashboard</h2>
      <div class="form-actions" style="margin: 0 0 12px;">
        <button onclick="repairDocLinks()">Repair Doc Links</button>
      </div>
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
        <tbody id="dispatchTableBody">
          <tr><td colspan="7">Loading dashboard data…</td></tr>
        </tbody>
      </table>
    `;
  }

  let html = `<!DOCTYPE html>
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
    .banner { margin: 12px 0 16px; padding: 10px 12px; border-radius: 4px; font-weight: 600; display: none; }
    .banner.show { display: block; }
    .banner.success { background: #e6f4ea; color: #137333; border: 1px solid #b7dfbf; }
    .banner.error { background: #fce8e6; color: #a50e0e; border: 1px solid #f6c7c3; }
    .tabs { display: flex; gap: 8px; margin: 8px 0 16px; }
    .tab { text-decoration: none; color: #174ea6; background: #e8f0fe; border: 1px solid #d2e3fc; border-radius: 4px; padding: 8px 12px; font-weight: 600; }
    .tab.active { background: #174ea6; color: #fff; border-color: #174ea6; }
    .card { background: #fff; border: 1px solid #d9d9d9; border-radius: 6px; padding: 16px; max-width: 760px; }
    form label { display: block; margin-bottom: 12px; font-weight: 600; }
    form input, form textarea { width: 100%; box-sizing: border-box; margin-top: 6px; padding: 8px; font: inherit; }
    .form-actions { margin-top: 8px; }
    .debug-strip { margin: 10px 0 14px; padding: 8px 10px; background: #fff8e1; border: 1px solid #f2cf6f; border-radius: 4px; font-size: 12px; }
    .debug-strip code { background: #fff; border: 1px solid #ead69b; border-radius: 3px; padding: 1px 4px; }
    #clientErrors { margin-top: 8px; background: #fff; border: 1px solid #ead69b; border-radius: 4px; padding: 8px; white-space: pre-wrap; min-height: 24px; }
    #clientLog { margin-top: 8px; background: #f3f7ff; border: 1px solid #bcd2ff; border-radius: 4px; padding: 8px; white-space: pre-wrap; min-height: 24px; }
  </style>
</head>
<body>
  <h1>CCG Dispatch DEV — Admin</h1>
  <section class="debug-strip" id="debugStrip">
    <strong>Debug</strong> — p: <code>${page || 'dashboard'}</code> | token: <code>${debugTokenPresent}</code> | role: <code>${debugRole}</code> | ts: <code>${renderTimestamp}</code> | renderId: <code>${renderId}</code> | htmlSize: <code id="htmlSizeValue">pending</code>
    <pre id="clientErrors">No client errors.</pre>
    <pre id="clientLog">No client log entries.</pre>
  </section>
  <div id="statusBanner" class="banner"></div>
  <nav class="tabs">${navHtml}</nav>
  ${pageContent}
  <script>
    const TOKEN = ${JSON.stringify(token || '')};
    const INITIAL_NOTICE = ${JSON.stringify(notice)};
    const INITIAL_ERROR = ${JSON.stringify(errorParam)};
    const CURRENT_PAGE = ${JSON.stringify(page)};
    const SERVER_RENDER_ID = ${JSON.stringify(renderId)};

    function appendClientError(message, source) {
      const el = document.getElementById('clientErrors');
      if (!el) return;
      const existing = String(el.textContent || '').trim();
      const prefix = existing && existing !== 'No client errors.' ? (existing + '\n') : '';
      el.textContent = prefix + '[' + new Date().toISOString() + '] ' + (source ? source + ': ' : '') + String(message || 'Unknown client error');
    }

    function appendClientLog(message) {
      const el = document.getElementById('clientLog');
      if (!el) return;
      const existing = String(el.textContent || '').trim();
      const prefix = existing && existing !== 'No client log entries.' ? (existing + '\n') : '';
      el.textContent = prefix + '[' + new Date().toISOString() + '] ' + String(message || 'log');
    }


    window.onerror = function (message, source, lineno, colno, error) {
      const stack = error && error.stack ? ('\n' + error.stack) : '';
      appendClientError(String(message || 'window.onerror') + ' @ ' + String(source || 'inline') + ':' + String(lineno || 0) + ':' + String(colno || 0) + stack, 'onerror');
    };

    window.addEventListener('unhandledrejection', function (event) {
      const reason = event && event.reason;
      const text = reason && reason.stack ? reason.stack : (reason && reason.message ? reason.message : String(reason || 'Unhandled promise rejection'));
      appendClientError(text, 'unhandledrejection');
    });

    function buildUrlWithParams(values) {
      const base = window.location.pathname;
      const params = new URLSearchParams();
      Object.keys(values || {}).forEach(function (key) {
        const value = values[key];
        if (value === undefined || value === null || String(value) === '') return;
        params.set(key, String(value));
      });
      return base + '?' + params.toString();
    }

    function showBanner(kind, text) {
      const el = document.getElementById('statusBanner');
      if (!el || !text) return;
      el.className = 'banner show ' + (kind === 'success' ? 'success' : 'error');
      el.textContent = text;
    }

    function escapeHtml(value) {
      return String(value || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
    }

    if (INITIAL_NOTICE && INITIAL_NOTICE.text) {
      showBanner(INITIAL_NOTICE.kind, INITIAL_NOTICE.text);
    }
    if (INITIAL_ERROR) {
      showBanner('error', INITIAL_ERROR);
    }

    function renderDispatchTable(data) {
      const tableBody = document.getElementById('dispatchTableBody');
      if (!tableBody) return;
      const dispatches = (data && data.activeDispatches) || [];
      const titleEl = document.getElementById('dashboardTitle');
      if (titleEl && data && data.currentUser) {
        const displayName = String(data.currentUser.display_name || '').trim();
        const role = String(data.currentUser.role || '').trim();
        titleEl.textContent = displayName
          ? 'dashboard — ' + displayName + (role ? ' (' + role + ')' : '')
          : 'dashboard';
      }

      if (!dispatches.length) {
        tableBody.innerHTML = '<tr><td colspan="7">No active dispatches found.</td></tr>';
        return;
      }

      tableBody.innerHTML = dispatches.map(function (dispatch) {
        const dispatchId = String(dispatch.dispatch_id || '').trim();
        const docUrl = String(dispatch.doc_url || '').trim();
        const docId = String(dispatch.doc_id || '').trim();
        const hasLinkedDoc = isValidDocUrl(docUrl) || isValidDocId(docId);
        const viewButtonLabel = hasLinkedDoc ? 'View Dispatch' : 'No doc linked';
        return '<tr>' +
          '<td>' + escapeHtml(dispatch.date) + '</td>' +
          '<td>' + escapeHtml(dispatch.truck_numbers) + '</td>' +
          '<td>' + escapeHtml(dispatch.client) + '</td>' +
          '<td>' + escapeHtml(dispatch.job_number) + '</td>' +
          '<td>' + escapeHtml(dispatch.start_time) + '</td>' +
          '<td>' + escapeHtml(dispatch.status) + '</td>' +
          '<td class="actions">' +
            '<button onclick="viewDispatch(' + JSON.stringify(dispatchId) + ', ' + JSON.stringify(encodeURIComponent(docUrl)) + ', ' + JSON.stringify(encodeURIComponent(docId)) + ')" ' + ((dispatchId && hasLinkedDoc) ? '' : 'disabled') + '>' + viewButtonLabel + '</button>' +
            '<button onclick="markCompleted(' + JSON.stringify(dispatchId) + ')" ' + (dispatchId ? '' : 'disabled') + '>Mark Completed</button>' +
            '<button onclick="amendDispatchAction(' + JSON.stringify(dispatchId) + ')" ' + (dispatchId ? '' : 'disabled') + '>Amend</button>' +
            '<button onclick="cancelDispatchAction(' + JSON.stringify(dispatchId) + ')" ' + (dispatchId ? '' : 'disabled') + '>Cancel</button>' +
          '</td>' +
        '</tr>';
      }).join('');
    }

    function loadDashboardData(token) {
      const sizeEl = document.getElementById('htmlSizeValue');
      if (sizeEl) {
        sizeEl.textContent = String((document.documentElement && document.documentElement.outerHTML && document.documentElement.outerHTML.length) || 0);
      }

      appendClientLog('calling getAdminDashboardData');

      let settled = false;
      const watchdog = window.setTimeout(function () {
        if (settled) return;
        settled = true;
        showBanner('error', 'Dashboard data load timed out');
        appendClientLog('timeout');
      }, 10000);

      google.script.run
        .withSuccessHandler(function (data) {
          if (settled) return;
          settled = true;
          window.clearTimeout(watchdog);
          appendClientLog('success');
          renderDispatchTable(data || {});
        })
        .withFailureHandler(function (error) {
          if (settled) return;
          settled = true;
          window.clearTimeout(watchdog);
          const message = (error && error.message) || String(error || 'Unknown error');
          appendClientLog('failure: ' + message);
          showBanner('error', message);
          const tableBody = document.getElementById('dispatchTableBody');
          if (tableBody) {
            tableBody.innerHTML = '<tr><td colspan="7">Failed to load dashboard data.</td></tr>';
          }
        })
        .getAdminDashboardData(token);
    }

    function navigateToPage(pageId) {
      window.location.href = buildUrlWithParams({ t: TOKEN, p: pageId });
    }

    function submitCreateDispatch(event) {
      event.preventDefault();
      const form = event.target;
      const formData = new FormData(form);
      const payload = {
        date: String(formData.get('date') || '').trim(),
        shift: String(formData.get('shift') || '').trim(),
        client: String(formData.get('client') || '').trim(),
        jobNumber: String(formData.get('jobNumber') || '').trim(),
        startTime: String(formData.get('startTime') || '').trim(),
        startLocation: String(formData.get('startLocation') || '').trim(),
        instructions: String(formData.get('instructions') || '').trim(),
        truckNumbers: String(formData.get('truckNumbers') || '').trim()
      };

      if (!payload.date || !payload.shift || !payload.client || !payload.jobNumber || !payload.startTime || !payload.startLocation || !payload.instructions || !payload.truckNumbers) {
        showBanner('error', 'All fields are required.');
        return;
      }

      google.script.run
        .withSuccessHandler(function () {
          window.location.href = buildUrlWithParams({ t: TOKEN, p: 'dashboard', msg: 'create_ok' });
        })
        .withFailureHandler(function (error) {
          showBanner('error', (error && error.message) || String(error || 'Unknown error'));
        })
        .createDispatchFromDashboard(TOKEN, payload);
    }

    function isValidDocUrl(value) {
      return /^https:\/\//i.test(String(value || '').trim());
    }

    function isValidDocId(value) {
      return /^[A-Za-z0-9_-]{21,}$/.test(String(value || '').trim());
    }

    function resolveDispatchDocUrl(docUrl, docId) {
      const trimmedUrl = String(docUrl || '').trim();
      const trimmedId = String(docId || '').trim();
      if (isValidDocUrl(trimmedUrl)) {
        return trimmedUrl;
      }
      if (isValidDocId(trimmedId)) {
        return 'https://docs.google.com/document/d/' + encodeURIComponent(trimmedId) + '/edit';
      }
      return '';
    }

    function viewDispatch(dispatchId, encodedDocUrl, encodedDocId) {
      const docUrl = decodeURIComponent(encodedDocUrl || '');
      const docId = decodeURIComponent(encodedDocId || '');
      const finalUrl = resolveDispatchDocUrl(docUrl, docId);
      if (!finalUrl) {
        showBanner('error', 'No document is linked for this dispatch.');
        return;
      }
      console.log('Opening dispatch doc URL', { dispatchId: dispatchId, docUrl: docUrl, docId: docId, finalUrl: finalUrl });
      window.open(finalUrl, '_blank', 'noopener');
    }


    function repairDocLinks() {
      if (!window.confirm('Run one-time repair for legacy doc links?')) return;
      google.script.run
        .withSuccessHandler(function () {
          showBanner('success', 'Dispatch doc links repair completed.');
          loadDashboardData(TOKEN);
        })
        .withFailureHandler(function (error) {
          showBanner('error', (error && error.message) || String(error || 'Unknown error'));
        })
        .repairDispatchDocLinksFromDashboard(TOKEN);
    }

    function markCompleted(dispatchId) {
      if (!dispatchId) return;
      google.script.run
        .withSuccessHandler(function () {
          showBanner('success', 'Dispatch marked completed.');
          loadDashboardData(TOKEN);
        })
        .withFailureHandler(function (error) {
          showBanner('error', (error && error.message) || String(error || 'Unknown error'));
        })
        .markDispatchCompletedFromDashboard(TOKEN, dispatchId);
    }

    function amendDispatchAction(dispatchId) {
      if (!dispatchId) return;
      const summary = window.prompt('Enter amendment summary:');
      if (summary === null) return;
      if (!String(summary || '').trim()) {
        showBanner('error', 'Amendment summary is required.');
        return;
      }
      google.script.run
        .withSuccessHandler(function () {
          showBanner('success', 'Dispatch amended.');
          loadDashboardData(TOKEN);
        })
        .withFailureHandler(function (error) {
          showBanner('error', (error && error.message) || String(error || 'Unknown error'));
        })
        .amendDispatchFromDashboard(TOKEN, dispatchId, summary);
    }

    function cancelDispatchAction(dispatchId) {
      if (!dispatchId) return;
      const reason = window.prompt('Enter cancellation reason:');
      if (reason === null) return;
      if (!String(reason || '').trim()) {
        showBanner('error', 'Cancellation reason is required.');
        return;
      }
      google.script.run
        .withSuccessHandler(function () {
          showBanner('success', 'Dispatch canceled.');
          loadDashboardData(TOKEN);
        })
        .withFailureHandler(function (error) {
          showBanner('error', (error && error.message) || String(error || 'Unknown error'));
        })
        .cancelDispatchFromDashboard(TOKEN, dispatchId, reason);
    }

    (function updateHtmlSize() {
      const sizeEl = document.getElementById('htmlSizeValue');
      if (sizeEl) sizeEl.textContent = String((document.documentElement && document.documentElement.outerHTML && document.documentElement.outerHTML.length) || 0);
    })();

    if (CURRENT_PAGE === 'dashboard') {
      loadDashboardData(TOKEN);
    }
  </script>
</body>
</html>`;

  if (!html || !String(html).trim()) {
    html = '<!DOCTYPE html><html><body><h1>admin</h1><h2>' + (page || 'dashboard') + '</h2><div>Fallback admin content.</div></body></html>';
  }

  return html;
}

/**
 * Render a minimal no-dependency ping page for doGet diagnostics.
 *
 * @param {string} token - Portal token.
 * @param {string} pageParam - Requested p value.
 * @returns {GoogleAppsScript.HTML.HtmlOutput} Static diagnostic HtmlOutput.
 */
function renderPingPage_(token, pageParam) {
  const html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>PING</title></head><body style="font-family:Arial,sans-serif;padding:16px;">'
    + '<h1 style="margin-top:0;">PING OK</h1>'
    + '<p><strong>token present:</strong> ' + (token ? 'yes' : 'no') + '</p>'
    + '<p><strong>p:</strong> ' + String(pageParam || '').replace(/</g, '&lt;') + '</p>'
    + '<p><strong>timestamp:</strong> ' + new Date().toISOString() + '</p>'
    + '</body></html>';
  return HtmlService.createHtmlOutput(html)
    .setTitle('PING')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Retrieve dashboard user/profile and active dispatch rows for client-side rendering.
 *
 * @param {string} token - Portal token.
 * @returns {{currentUser:{display_name:string,role:string},activeDispatches:Array<Object>}}
 */
function getAdminDashboardData(token) {
  const tokenPresent = String(token || '').trim() ? 'yes' : 'no';
  Logger.log('[DEV] getAdminDashboardData start tokenPresent=' + tokenPresent);

  let actor = null;
  try {
    actor = getAuthorizedDashboardActor_(token);
    const actorRole = String((actor && actor.role) || '').trim().toLowerCase() || 'unknown';
    Logger.log('[DEV] getAdminDashboardData actor role=' + actorRole);

    const sheet = SpreadsheetApp.openById(DEV_DB_SPREADSHEET_ID).getSheetByName('Dispatches');
    const response = {
      currentUser: {
        display_name: String((actor && actor.displayName) || '').trim(),
        role: String((actor && actor.role) || '').trim().toLowerCase()
      },
      activeDispatches: []
    };

    if (!sheet || sheet.getLastRow() < 2) {
      Logger.log('[DEV] getAdminDashboardData end role=' + response.currentUser.role + ' rows=0');
      return response;
    }

    const lastColumn = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(h => String(h || '').trim());
    const headerMap = {};
    headers.forEach((h, idx) => { if (h) headerMap[h] = idx; });

    const requiredHeaders = ['dispatch_id', 'date', 'truck_numbers', 'client', 'job_number', 'start_time', 'status'];
    const missingHeader = requiredHeaders.find(name => headerMap[name] === undefined);
    if (missingHeader) {
      throw new Error(`Dispatches tab is missing required header: ${missingHeader}`);
    }

    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastColumn).getValues();
    response.activeDispatches = rows
      .filter(row => String(row[headerMap.status] || '').trim() !== 'Completed')
      .map((row) => {
        const dateValue = row[headerMap.date];
        const dateText = dateValue instanceof Date
          ? dateValue.toISOString()
          : String(dateValue || '').trim();
        return {
          dispatch_id: String(row[headerMap.dispatch_id] || '').trim(),
          date: dateText,
          truck_numbers: String(row[headerMap.truck_numbers] || '').trim(),
          client: String(row[headerMap.client] || '').trim(),
          job_number: String(row[headerMap.job_number] || '').trim(),
          start_time: String(row[headerMap.start_time] || '').trim(),
          status: String(row[headerMap.status] || '').trim(),
          doc_url: headerMap.doc_url !== undefined ? String(row[headerMap.doc_url] || '').trim() : '',
          doc_id: headerMap.doc_id !== undefined ? String(row[headerMap.doc_id] || '').trim() : ''
        };
      });

    Logger.log('[DEV] getAdminDashboardData end role=' + response.currentUser.role + ' rows=' + response.activeDispatches.length);
    return response;
  } catch (error) {
    const roleForError = String((actor && actor.role) || 'unknown').trim().toLowerCase() || 'unknown';
    Logger.log('[DEV] getAdminDashboardData error role=' + roleForError + ' message=' + ((error && error.message) ? error.message : String(error || 'unknown')));
    throw error;
  }
}

/**
 * Render an explicit non-empty fallback page for admin routes.
 *
 * @param {string} token - Portal token.
 * @param {{role:string}|null} user - Active user object when available.
 * @param {string} pageParam - Requested page parameter.
 * @param {Error|string} error - Original error.
 * @returns {GoogleAppsScript.HTML.HtmlOutput} Visible fallback HtmlOutput.
 */
function renderAdminFallbackPage_(token, user, pageParam, error) {
  const message = (error && error.message) ? error.message : String(error || 'Unknown error');
  const stack = error && error.stack ? String(error.stack) : '';
  const role = user ? String(user.role || '').trim() : 'unknown';
  const safeMessage = String(message || '').replace(/</g, '&lt;');
  const safeStack = String(stack || '').replace(/</g, '&lt;');
  const safePageParam = String(pageParam || '(empty)').replace(/</g, '&lt;');
  const safeRole = String(role || 'unknown').replace(/</g, '&lt;');
  const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Admin Fallback</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 24px; background: #fff; color: #111; }
    h1 { margin: 0 0 16px; background: #c62828; color: #fff; padding: 16px; font-size: 36px; letter-spacing: 1px; }
    .card { border: 2px solid #c62828; border-radius: 6px; padding: 16px; max-width: 980px; }
    code, pre { background: #f5f5f5; border: 1px solid #ddd; border-radius: 4px; padding: 8px; display: block; white-space: pre-wrap; }
    .meta p { margin: 8px 0; }
  </style>
</head>
<body>
  <h1>ADMIN FALLBACK</h1>
  <section class="card">
    <div class="meta">
      <p><strong>timestamp:</strong> ${new Date().toISOString()}</p>
      <p><strong>p:</strong> ${safePageParam}</p>
      <p><strong>token present:</strong> ${token ? 'yes' : 'no'}</p>
      <p><strong>role:</strong> ${safeRole}</p>
    </div>
    <h2>Error message</h2>
    <code>${safeMessage || 'none'}</code>
    <h2>Error stack</h2>
    <pre>${safeStack || 'No stack available.'}</pre>
  </section>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html)
    .setTitle('Admin Fallback')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
 * Validate dispatcher/admin token for dashboard actions.
 *
 * @param {string} token - Portal token.
 * @returns {{userId:string,displayName:string,role:string}} Authenticated actor.
 */
function getAuthorizedDashboardActor_(token) {
  const safeToken = String(token || '').trim();
  if (!safeToken) {
    throw new Error('Missing token for dashboard action.');
  }

  const user = getActivePortalUserByToken_(safeToken);
  if (!isAdminOrDispatcherUser_(user)) {
    throw new Error('Unauthorized dashboard action.');
  }

  return user;
}

/**
 * Dashboard action: mark dispatch completed.
 *
 * @param {string} token - Portal token.
 * @param {string} dispatchId - Dispatch UUID.
 * @returns {{status:string,rowNumber:number}} Server result.
 */
function markDispatchCompletedFromDashboard(token, dispatchId) {
  const actor = getAuthorizedDashboardActor_(token);
  const updatedBy = actor.userId || actor.displayName || 'unknown';
  return markDispatchCompleted(dispatchId, updatedBy);
}

/**
 * Dashboard action: amend dispatch.
 *
 * @param {string} token - Portal token.
 * @param {string} dispatchId - Dispatch UUID.
 * @param {string} changeSummary - Required amendment note.
 * @returns {{status:string,rowNumber:number}} Server result.
 */
function amendDispatchFromDashboard(token, dispatchId, changeSummary) {
  const actor = getAuthorizedDashboardActor_(token);
  const updatedBy = actor.userId || actor.displayName || 'unknown';
  return amendDispatch(dispatchId, changeSummary, updatedBy);
}

/**
 * Dashboard action: cancel dispatch.
 *
 * @param {string} token - Portal token.
 * @param {string} dispatchId - Dispatch UUID.
 * @param {string} cancelReason - Required cancellation reason.
 * @returns {{status:string,rowNumber:number}} Server result.
 */
function cancelDispatchFromDashboard(token, dispatchId, cancelReason) {
  const actor = getAuthorizedDashboardActor_(token);
  const updatedBy = actor.userId || actor.displayName || 'unknown';
  return cancelDispatch(dispatchId, cancelReason, updatedBy);
}

/**
 * Dashboard action: create a dispatch and generated document.
 *
 * @param {string} token - Portal token.
 * @param {{date:string,shift:string,client:string,jobNumber:string,startTime:string,startLocation:string,instructions:string,truckNumbers:string}} payload
 * @returns {{status:string}} Server result.
 */
function createDispatchFromDashboard(token, payload) {
  const actor = getAuthorizedDashboardActor_(token);
  createDispatchFromPortalForm(payload, actor.userId || actor.displayName || 'unknown');
  return { status: 'ok' };
}


/**
 * Dashboard action: run legacy doc-link repair (admin/dispatcher only).
 *
 * @param {string} token - Portal token.
 * @returns {{scanned:number,updated:number,movedFromLastUpdatedBy:number,synthesizedFromDocId:number,overwroteInvalidDocUrl:number}}
 */
function repairDispatchDocLinksFromDashboard(token) {
  getAuthorizedDashboardActor_(token);
  return repairDispatchDocLinks_DEV_();
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
  return html;
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
