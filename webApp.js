/**
 * This file serves company dispatch web pages stored in Google Drive.
 */

// Main Drive folder where company dispatch HTML files are stored.
const DISPATCH_ARCHIVES_FOLDER_ID = '1Fic0PvyH2B-Dq7P0hYQLsn0jB09qOWLE';
const DEV_DEBUG = isDevEnvironment_();
const SHOW_REPAIR_BUTTON = false;

const NY_TIMEZONE = 'America/New_York';

function formatMinutesAsTime12_(totalMinutes) {
  const safeMinutes = ((Number(totalMinutes) % 1440) + 1440) % 1440;
  const hour24 = Math.floor(safeMinutes / 60);
  const minute = safeMinutes % 60;
  const suffix = hour24 >= 12 ? 'PM' : 'AM';
  const hour12 = hour24 % 12 || 12;
  return `${String(hour12).padStart(2, '0')}:${String(minute).padStart(2, '0')} ${suffix}`;
}

function normalizeTimeString_(value) {
  const text = String(value || '').trim();
  const match = text.match(/^(\d{1,2}):(\d{2})(?:\s*([AaPp][Mm]))?$/);
  if (!match) return text;

  const hour = Number(match[1]);
  const minute = Number(match[2]);
  const suffix = match[3] ? String(match[3]).toUpperCase() : '';
  if (!isFinite(hour) || !isFinite(minute)) return text;

  if (suffix) {
    const safeHour = Math.min(Math.max(hour, 1), 12);
    const safeMinute = Math.min(Math.max(minute, 0), 59);
    return `${String(safeHour).padStart(2, '0')}:${String(safeMinute).padStart(2, '0')} ${suffix}`;
  }

  return `${String(Math.min(Math.max(hour, 0), 23)).padStart(2, '0')}:${String(Math.min(Math.max(minute, 0), 59)).padStart(2, '0')}`;
}

function parseSheetDateOrTimeValue_(value) {
  if (value instanceof Date) return new Date(value.getTime());
  if (typeof value === 'number' && isFinite(value)) {
    if (Math.abs(value) >= 1) {
      const millis = Math.round((value - 25569) * 86400000);
      return new Date(millis);
    }
    return { type: 'sheetSerialTime', minutes: Math.round(value * 24 * 60) };
  }

  const text = String(value || '').trim();
  if (!text) return null;

  if (/^\d{1,2}:\d{2}/.test(text)) {
    return normalizeTimeString_(text);
  }

  const isoDateOnly = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoDateOnly) {
    return new Date(Number(isoDateOnly[1]), Number(isoDateOnly[2]) - 1, Number(isoDateOnly[3]));
  }

  const slashDate = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashDate) {
    return new Date(Number(slashDate[3]), Number(slashDate[1]) - 1, Number(slashDate[2]));
  }

  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function formatDateNY_(value) {
  const parsed = parseSheetDateOrTimeValue_(value);
  if (!parsed) return String(value || '').trim();
  return Utilities.formatDate(parsed, NY_TIMEZONE, 'MM/dd/yyyy');
}

function formatTimeNY_(value) {
  const parsed = parseSheetDateOrTimeValue_(value);
  if (!parsed) return String(value || '').trim();
  if (typeof parsed === 'string') return parsed;
  if (parsed && parsed.type === 'sheetSerialTime') {
    return formatMinutesAsTime12_(parsed.minutes);
  }
  return Utilities.formatDate(parsed, NY_TIMEZONE, 'hh:mm a');
}

function isValidDocUrl_(value) {
  const trimmed = String(value || '').trim();
  return /^https:\/\/docs\.google\.com\/document\/d\//i.test(trimmed)
    || /^https:\/\//i.test(trimmed);
}

function isValidDocId_(value) {
  return /^[A-Za-z0-9_-]{21,}$/.test(String(value || '').trim());
}

function resolveDocUrl_(docUrl, docId) {
  if (isValidDocUrl_(docUrl)) return docUrl;
  if (isValidDocId_(docId)) return 'https://docs.google.com/document/d/' + docId + '/edit';
  return '';
}

/**
 * Handle a web request and return the matching scoped dispatch page.
 *
 * Expected URL format: ?t=TOKEN
 */
function doGet(e) {
  const token = String((e && e.parameter && e.parameter.t) || '').trim();
  const pageParam = String((e && e.parameter && e.parameter.p) || '').trim().toLowerCase();
  let user = null;
  let isAdmin = false;
  let p = pageParam || 'dashboard';

  function logDoGet_(htmlLen) {
    Logger.log('[DEV] doGet p=' + p + ' tokenPresent=' + (!!token) + ' isAdmin=' + isAdmin + ' userRole=' + (user && user.role));
    Logger.log('[DEV] doGet htmlLen=' + Number(htmlLen || 0));
  }

  try {
    if (pageParam === 'ping') {
      p = 'ping';
      logDoGet_(58);
      return renderPingPage_(token, pageParam);
    }

    const companyParam = String((e && e.parameter && e.parameter.company) || '').trim();
    const isDev = isDevEnvironment_();
    user = token ? getActivePortalUserByToken_(token) : null;
    isAdmin = isAdminOrDispatcherUser_(user);
    const shouldUseCompanyFallback = isDev && companyParam && (!token || !user);

    if (pageParam === 'lite') {
      p = 'lite';
      if (!isAdmin) {
        const unauthorizedHtml = 'not authorized';
        logDoGet_(unauthorizedHtml.length);
        return HtmlService.createHtmlOutput(unauthorizedHtml)
          .setTitle('Not Authorized');
      }
      const liteHtml = '<!doctype html><html><body><div id="boot">PENDING</div><pre id="log">empty</pre><script>document.getElementById(\'boot\').textContent=\'OK\';document.getElementById(\'log\').textContent=\'script ran at \' + new Date().toISOString();</script></body></html>';
      logDoGet_(liteHtml.length);
      return HtmlService.createHtmlOutput(liteHtml)
        .setTitle('Lite Test');
    }

    if (pageParam === 'safe') {
      p = 'safe';
      if (!isAdmin) {
        const unauthorizedHtml = 'not authorized';
        logDoGet_(unauthorizedHtml.length);
        return HtmlService.createHtmlOutput(unauthorizedHtml)
          .setTitle('Not Authorized');
      }
      const safeHtml = renderSafeAdminPage_(token);
      logDoGet_(safeHtml.length);
      return HtmlService.createHtmlOutput(safeHtml)
        .setTitle('Safe Dashboard');
    }

    if (pageParam === 'source') {
      p = 'source';
      if (!isAdmin) {
        const unauthorizedHtml = 'not authorized';
        logDoGet_(unauthorizedHtml.length);
        return HtmlService.createHtmlOutput(unauthorizedHtml)
          .setTitle('Not Authorized');
      }

      const dashboardHtml = buildAdminHtml_({
        user: user,
        token: token,
        params: e && e.parameter,
        forcedPage: 'dashboard'
      });
      Logger.log('[DEV] dashboard html head=' + String(dashboardHtml || '').slice(0, 500));
      Logger.log('[DEV] dashboard html tail=' + String(dashboardHtml || '').slice(-500));

      const sourceHtml = '<!doctype html><html><head><meta charset="UTF-8"><title>Dashboard Source</title></head><body><h1>DASHBOARD SOURCE</h1><pre>'
        + escapeHtmlForPre_(dashboardHtml || '')
        + '</pre></body></html>';
      logDoGet_(sourceHtml.length);
      return HtmlService.createHtmlOutput(sourceHtml)
        .setTitle('Dashboard Source');
    }

    if (isAdmin) {
      const pageAliases = { 'create-dispatch': 'create', 'companies-trucks': 'companies' };
      const allowedPages = { dashboard: true, create: true, edit: true, companies: true, users: true };
      p = pageAliases[p] || p;
      if (!allowedPages[p]) p = 'dashboard';
    }

    if (!shouldUseCompanyFallback) {
      if (!token) {
        logDoGet_(0);
        return renderErrorPage_('Invalid or missing token. Please use your dispatch portal link.');
      }

      if (!user) {
        logDoGet_(0);
        return renderErrorPage_('Invalid or missing token. Please contact dispatch admin.');
      }
    }

    let content = '';
    let pageTitle = 'Dispatch List';

    if (!shouldUseCompanyFallback && isAdmin) {
      if (pageParam === 'repair') {
        p = 'dashboard';
        repairDispatchDocLinks_DEV_();
        content = buildAdminHtml_({
          user: user,
          token: token,
          params: Object.assign({}, (e && e.parameter) || {}, { msg: 'repair_done' }),
          forcedPage: 'dashboard'
        });
        pageTitle = 'Admin Dispatch Dashboard';
        logDoGet_(content.length);
        return HtmlService.createHtmlOutput(content)
          .setTitle(pageTitle);
      }

      const resolvedPage = p;
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
        logDoGet_(0);
        return renderAdminFallbackPage_(token, user, pageParam, adminError);
      }
      pageTitle = 'Admin Dispatch Dashboard';
      logDoGet_(content.length);
      return HtmlService.createHtmlOutput(content)
        .setTitle(pageTitle);
    }

    if (!shouldUseCompanyFallback && user.truckNumber) {
      content = buildTruckScopedPortalHtml_(user.truckNumber, token);
      pageTitle = `${user.truckNumber} Dispatch List`;
    } else {
      const resolvedCompany = shouldUseCompanyFallback ? companyParam : user.company;
      if (!resolvedCompany) {
        logDoGet_(0);
        return renderErrorPage_('Token is missing truck/company scope. Please contact dispatch admin.');
      }

      const companyFileName = `${resolvedCompany}_dispatch_list.html`;
      const archivesFolder = DriveApp.getFolderById(DISPATCH_ARCHIVES_FOLDER_ID);
      const file = findFileRecursively(archivesFolder, companyFileName);

      if (!file) {
        logDoGet_(0);
        return renderErrorPage_(`No dispatch list found for company: ${resolvedCompany}.`);
      }

      content = file.getBlob().getDataAsString();
      if (shouldUseCompanyFallback) {
        content = addDevFallbackBanner_(content);
      }
      pageTitle = `${resolvedCompany} Dispatch List`;
    }

    logDoGet_(content ? content.length : 0);
    return HtmlService.createHtmlOutput(content)
      .setTitle(pageTitle);
  } catch (error) {
    try {
      user = token ? getActivePortalUserByToken_(token) : null;
      isAdmin = isAdminOrDispatcherUser_(user);
    } catch (userLookupError) {
      user = null;
      isAdmin = false;
    }
    logDoGet_(0);
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
  const allowedPages = { dashboard: true, create: true, edit: true, companies: true, users: true };
  let page = forcedPage || pageAliases[pageParamRaw] || pageParamRaw || 'dashboard';
  if (!allowedPages[page]) page = 'dashboard';

  const notices = {
    completed_ok: { kind: 'success', text: 'Dispatch marked completed.' },
    amend_ok: { kind: 'success', text: 'Dispatch amended.' },
    edit_ok: { kind: 'success', text: 'Dispatch amended.' },
    cancel_ok: { kind: 'success', text: 'Dispatch canceled.' },
    create_ok: { kind: 'success', text: 'Dispatch created successfully.' },
    repair_done: { kind: 'success', text: 'Dispatch doc links repair completed.' }
  };
  const notice = notices[messageParam] || null;
  const renderId = Utilities.getUuid().slice(0, 8);
  const renderTimestamp = new Date().toISOString();
  const debugRole = String((user && user.role) || '').trim() || 'unknown';
  const debugTokenPresent = token ? 'yes' : 'no';
  const showDebugUi = DEV_DEBUG;
  const showRepairButton = DEV_DEBUG || SHOW_REPAIR_BUTTON;

  const navItems = [
    { id: 'dashboard', label: 'Dashboard' },
    { id: 'create', label: 'Create Dispatch' },
    { id: 'companies', label: 'Companies/Trucks' },
    { id: 'users', label: 'Users' }
  ];

  const baseUrl = ScriptApp.getService().getUrl();
  const navHtml = navItems.map((item) => {
    const activeClass = item.id === page ? 'active' : '';
    const paramsForLink = ['t=' + encodeURIComponent(token || '')];
    paramsForLink.push('p=' + encodeURIComponent(item.id));
    const href = baseUrl + '?' + paramsForLink.join('&');
    return `<a class="tab ${activeClass}" target="_top" href="${href}">${item.label}</a>`;
  }).join('\n');

  let pageContent = '';

  if (page === 'create' || page === 'edit') {
    const editDispatchId = String((params && params.dispatch_id) || '').trim();
    const editRecord = page === 'edit' && editDispatchId ? getDispatchForDashboardEdit(token, editDispatchId) : null;
    const initialMode = editRecord ? 'edit' : 'create';
    const initialPayload = editRecord || null;
    pageContent = `
      <h2>${page === 'edit' ? 'Edit Dispatch' : 'Create Dispatch'}</h2>
      <section class="card page-shell">
        <form id="createDispatchForm" data-form-mode="${initialMode}" data-dispatch-id="${editDispatchId}" onsubmit="submitCreateDispatch(event)">
        <label>Status
          <select name="status" id="statusField" required>
            <option value="Confirmed">Confirmed</option>
            <option value="Dispatched">Dispatched</option>
          </select>
        </label>
        <label>Date
          <input type="date" name="date" required>
        </label>
        <label>Shift
          <select name="shift" required>
            <option value="">Select shift</option>
            <option value="day">day</option>
            <option value="night">night</option>
          </select>
        </label>
        <label>Client
          <input type="text" name="client">
        </label>
        <div id="assignmentBlocks"></div>
        <button type="button" onclick="addAssignmentBlock()">Add Assignment</button>
        <label>Truck Numbers (comma-separated)
          <input type="text" name="truckNumbers" placeholder="RT03, RT12" required>
        </label>
        <label>Tolls Policy
          <select name="tollsPolicy">
            <option value="">Select</option>
            <option value="included">included</option>
            <option value="authorized">authorized</option>
            <option value="unauthorized">unauthorized</option>
          </select>
        </label>
        <label>Notes
          <textarea name="notes" rows="3"></textarea>
        </label>
        <label id="changeSummaryWrap" style="display:none;">Change Summary
          <textarea name="changeSummary" rows="3"></textarea>
        </label>
        <div class="form-actions">
          ${page === 'edit'
            ? '<button type="submit">Save Dispatch</button>'
            : '<button type="submit" onclick="window.__submitMode = &quot;dashboard&quot;">Submit & Return to Dashboard</button>\n          <button type="submit" onclick="window.__submitMode = &quot;new&quot;">Submit & Create New Dispatch</button>'}
        </div>
        </form>
      </section>
      <script>window.__INITIAL_DISPATCH_PAYLOAD__ = ${JSON.stringify(initialPayload)};</script>
    `;
  } else if (page === 'companies') {
    pageContent = '<h2>companies</h2><section class="card page-shell"><p>Companies/Trucks placeholder content.</p></section>';
  } else if (page === 'users') {
    pageContent = '<h2>users</h2><section class="card page-shell"><p>Users placeholder content.</p></section>';
  } else {
    pageContent = `
      <h2 id="dashboardTitle">dashboard</h2>
      ${showRepairButton ? `<div class="form-actions" style="margin: 0 0 12px;"><button type="button" onclick="repairDocLinks()">Repair Doc Links</button></div>` : ''}
      <table>
        <thead>
          <tr>
            <th>Date</th>
            <th>Start Time</th>
            <th>Truck Number</th>
            <th>Job Number</th>
            <th>Status</th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody id="dispatchTableBody">
          <tr><td colspan="6">Loading dashboard data…</td></tr>
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
    .debug-strip { margin: 10px 0 14px; padding: 8px 10px; background: #fff8e1; border: 1px solid #f2cf6f; border-radius: 4px; font-size: 12px; position: static; pointer-events: none; z-index: 0; max-height: 160px; overflow: auto; }
    .debug-strip code { background: #fff; border: 1px solid #ead69b; border-radius: 3px; padding: 1px 4px; }
    #clientErrors { margin-top: 8px; background: #fff; border: 1px solid #ead69b; border-radius: 4px; padding: 8px; white-space: pre-wrap; min-height: 24px; position: static; pointer-events: none; z-index: 0; max-height: 160px; overflow: auto; }
    #clientLog { margin-top: 8px; background: #f3f7ff; border: 1px solid #bcd2ff; border-radius: 4px; padding: 8px; white-space: pre-wrap; min-height: 24px; position: static; pointer-events: none; z-index: 0; max-height: 160px; overflow: auto; }
    details.errors-box { margin-top: 8px; }
    details.errors-box > summary { cursor: pointer; font-weight: 600; }
    nav.tabs, .tabs a, button { pointer-events: auto; position: relative; z-index: 10; }
  </style>
</head>
<body>
  ${showDebugUi ? '<div id="jsBoot" style="padding:6px;border:2px solid #000;display:inline-block;">JS BOOT PENDING</div>' : ''}
  ${showDebugUi ? '<pre id="clientLog">No client log entries.</pre>' : ''}
  ${showDebugUi
    ? '<pre id="clientErrors">No client errors.</pre>'
    : '<details class="errors-box"><summary>Errors</summary><pre id="clientErrors">No client errors.</pre></details>'}
  <script>
    function appendLine(existing, line, emptySentinel) {
      var cur = String(existing || '').trim();
      var prefix = cur && cur !== emptySentinel ? (cur + '\\n') : '';
      return prefix + line;
    }

    try {
      (function () {
        var bootEl = document.getElementById('jsBoot');
        if (bootEl) bootEl.textContent = 'JS BOOT OK';
        var logEl = document.getElementById('clientLog');
        if (logEl) {
          logEl.textContent = appendLine(logEl.textContent, '[boot] script executed at ' + new Date().toISOString(), 'No client log entries.');
        }
      })();
    } catch (bootError) {
      var errorEl = document.getElementById('clientErrors');
      if (errorEl) {
        errorEl.textContent = appendLine(errorEl.textContent, '[boot] ' + (bootError && bootError.message ? bootError.message : String(bootError || 'Unknown boot error')), 'No client errors.');
      }
    }

    loadNotifications();
  </script>
  <h1>CCG Dispatch DEV — Admin</h1>
  ${showDebugUi ? `<section class="debug-strip" id="debugStrip"><strong>Debug</strong> — p: <code>${page || 'dashboard'}</code> | token: <code>${debugTokenPresent}</code> | role: <code>${debugRole}</code> | ts: <code>${renderTimestamp}</code> | renderId: <code>${renderId}</code> | htmlSize: <code id="htmlSizeValue">pending</code></section>` : ''}
  <div id="statusBanner" class="banner"></div>
  <nav class="tabs">${navHtml}</nav>
  ${pageContent}
  <script>
    const TOKEN = ${JSON.stringify(token || '')};
    const BASE_URL = ${JSON.stringify(baseUrl || '')};
    const INITIAL_NOTICE = ${JSON.stringify(notice)};
    const INITIAL_ERROR = ${JSON.stringify(errorParam)};
    const CURRENT_PAGE = ${JSON.stringify(page)};
    const SERVER_RENDER_ID = ${JSON.stringify(renderId)};
    const DEV_DEBUG = ${JSON.stringify(showDebugUi)};
    window.__submitMode = null;

    function appendLine(existing, line, emptySentinel) {
      const cur = String(existing || '').trim();
      const prefix = cur && cur !== emptySentinel ? (cur + '\\n') : '';
      return prefix + line;
    }

    function appendClientError(message, source) {
      const el = document.getElementById('clientErrors');
      if (!el) return;
      el.textContent = appendLine(el.textContent, '[' + new Date().toISOString() + '] ' + (source ? source + ': ' : '') + String(message || 'Unknown client error'), 'No client errors.');
    }

    function appendClientLog(message) {
      if (!DEV_DEBUG) return;
      const el = document.getElementById('clientLog');
      if (!el) return;
      el.textContent = appendLine(el.textContent, '[' + new Date().toISOString() + '] ' + String(message || 'log'), 'No client log entries.');
    }

    (function patchEventPrototypeForTrace() {
      if (window.__dispatchEventTracePatched) return;
      window.__dispatchEventTracePatched = true;

      const originalPreventDefault = Event.prototype.preventDefault;
      Event.prototype.preventDefault = function () {
        this.__pd = true;
        return originalPreventDefault.apply(this, arguments);
      };

      const originalStopPropagation = Event.prototype.stopPropagation;
      Event.prototype.stopPropagation = function () {
        this.__sp = true;
        return originalStopPropagation.apply(this, arguments);
      };
    })();


    window.onerror = function (message, source, lineno, colno, error) {
      const stack = error && error.stack ? ('\\n' + error.stack) : '';
      appendClientError(String(message || 'window.onerror') + ' @ ' + String(source || 'inline') + ':' + String(lineno || 0) + ':' + String(colno || 0) + stack, 'onerror');
    };

    window.addEventListener('unhandledrejection', function (event) {
      const reason = event && event.reason;
      const text = reason && reason.stack ? reason.stack : (reason && reason.message ? reason.message : String(reason || 'Unhandled promise rejection'));
      appendClientError(text, 'unhandledrejection');
    });

    document.addEventListener('click', function (e) {
      const target = e && e.target ? e.target : null;
      if (!target) {
        appendClientLog('[click-trace] target=unknown id= class= href= defaultPrevented(before)=' + String(Boolean(e && e.defaultPrevented)) + ' pdFlag=' + String(Boolean(e && e.__pd)) + ' spFlag=' + String(Boolean(e && e.__sp)));
        return;
      }

      const beforeDefaultPrevented = Boolean(e.defaultPrevented);
      const anchor = target.closest ? target.closest('a.tab') : null;
      const hrefAttr = anchor ? String(anchor.getAttribute('href') || '') : (target.getAttribute ? String(target.getAttribute('href') || '') : '');
      const absoluteHref = anchor ? String(anchor.href || '') : '';
      appendClientLog('[click-trace] target=' + String(target.tagName || 'unknown') + ' id=' + String(target.id || '') + ' class=' + String(target.className || '') + ' href=' + hrefAttr + ' defaultPrevented(before)=' + String(beforeDefaultPrevented) + ' pdFlag=' + String(Boolean(e.__pd)) + ' spFlag=' + String(Boolean(e.__sp)));

      if (anchor && absoluteHref) {
        appendClientLog('[nav-force] assign=' + absoluteHref);
        setTimeout(function () {
          const beforeHref = String(window.location.href || '');
          try {
            window.location.assign(absoluteHref);
          } catch (assignError) {
            appendClientLog('[nav-force] assign error=' + String(assignError && assignError.message ? assignError.message : assignError));
          }
          window.setTimeout(function () {
            const afterHref = String(window.location.href || '');
            if (afterHref === beforeHref) {
              appendClientLog('[nav-fallback] opening _top');
              window.open(absoluteHref, '_top');
            }
          }, 300);
        }, 0);
      }

      const actionButton = target.closest ? target.closest('button.actionBtn') : null;
      if (actionButton) {
        const action = String(actionButton.dataset.action || '').trim();
        const dispatchId = String(actionButton.dataset.dispatchId || '').trim();
        const docUrl = String(actionButton.dataset.docUrl || '').trim();
        const docId = String(actionButton.dataset.docId || '').trim();

        appendClientLog('[action-click] ' + action + ' id=' + dispatchId);
        if (!dispatchId) {
          appendClientError('[action-click] missing dispatchId for action=' + action, 'action-click');
        }

        switch (action) {
          case 'view': {
            const finalUrl = resolveDispatchDocUrl(docUrl, docId);
            if (!finalUrl) {
              showBanner('error', 'No document is linked for this dispatch.');
              return;
            }
            window.open(finalUrl, '_blank', 'noopener');
            return;
          }
          case 'complete': {
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
            return;
          }
          case 'amend': {
            if (!dispatchId) return;
            goAdmin('edit', { dispatch_id: dispatchId });
            return;
          }
          case 'cancel': {
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
            return;
          }
          default:
            return;
        }
      }

      setTimeout(function () {
        appendClientLog('[click-trace-post] target=' + String(target.tagName || 'unknown') + ' defaultPrevented(after)=' + String(Boolean(e.defaultPrevented)) + ' pdFlag=' + String(Boolean(e.__pd)) + ' spFlag=' + String(Boolean(e.__sp)));
      }, 0);
    }, true);

    function forceSafeDebugBox(el) {
      if (!el || !el.style) return;
      el.style.position = 'static';
      el.style.pointerEvents = 'none';
      el.style.zIndex = '0';
      el.style.maxHeight = '160px';
      el.style.overflow = 'auto';
      el.style.display = 'block';
      el.style.width = 'auto';
      el.style.height = 'auto';
    }

    function hardenAdminDiagnostics() {
      ['statusBanner', 'debugStrip', 'clientLog', 'clientErrors'].forEach(function (id) {
        const el = document.getElementById(id);
        if (!el) return;
        const computed = window.getComputedStyle(el);
        const rect = el.getBoundingClientRect();
        if (computed.position === 'fixed' || computed.position === 'absolute' || rect.width >= window.innerWidth * 0.95 || rect.height >= window.innerHeight * 0.95) {
          appendClientLog('[overlay-guard] hardening #' + id + ' position=' + computed.position + ' rect=' + [Math.round(rect.left), Math.round(rect.top), Math.round(rect.width), Math.round(rect.height)].join(','));
        }
        forceSafeDebugBox(el);
      });
    }

    function findViewportOverlays() {
      const viewportW = Math.max(1, window.innerWidth || 0);
      const viewportH = Math.max(1, window.innerHeight || 0);
      const overlayIds = { clientLog: true, clientErrors: true, debugStrip: true, jsBoot: true };
      const elements = document.body ? document.body.querySelectorAll('*') : [];

      Array.prototype.forEach.call(elements, function (el) {
        if (!el || !window.getComputedStyle) return;
        const style = window.getComputedStyle(el);
        if (style.pointerEvents === 'none') return;
        if (style.position !== 'fixed' && style.position !== 'absolute') return;
        const rect = el.getBoundingClientRect();
        if (!rect || rect.width <= 0 || rect.height <= 0) return;

        const coversMostViewport = rect.width >= viewportW * 0.9 && rect.height >= viewportH * 0.9;
        if (!coversMostViewport) return;

        const descriptor = (el.tagName || 'unknown') + '#' + String(el.id || '(no-id)') + '.' + String(el.className || '(no-class)');
        const box = 'x=' + Math.round(rect.left) + ',y=' + Math.round(rect.top) + ',w=' + Math.round(rect.width) + ',h=' + Math.round(rect.height);
        const detail = '[overlay] ' + descriptor + ' z=' + String(style.zIndex || 'auto') + ' pointer=' + String(style.pointerEvents || 'auto') + ' pos=' + String(style.position || '') + ' ' + box;
        appendClientError(detail, 'overlay-scan');
        appendClientLog(detail);

        if (overlayIds[String(el.id || '')]) {
          forceSafeDebugBox(el);
          appendClientLog('[overlay-fix] forced safe styles on #' + String(el.id || 'unknown'));
        }
      });
    }

    function buildUrlWithParams(values) {
      const base = BASE_URL || window.location.href.split('?')[0];
      const params = new URLSearchParams();
      Object.keys(values || {}).forEach(function (key) {
        const value = values[key];
        if (value === undefined || value === null || String(value) === '') return;
        params.set(key, String(value));
      });
      return base + '?' + params.toString();
    }

    function goAdmin(page, extraParams) {
      const base = BASE_URL || window.location.href.split('?')[0];
      const parts = [
        't=' + encodeURIComponent(TOKEN),
        'p=' + encodeURIComponent(String(page || 'dashboard'))
      ];
      Object.keys(extraParams || {}).forEach(function (key) {
        const value = extraParams[key];
        if (value === undefined || value === null || String(value) === '') return;
        parts.push(encodeURIComponent(key) + '=' + encodeURIComponent(String(value)));
      });
      const url = base + '?' + parts.join('&');
      window.open(url, '_top');
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
        tableBody.innerHTML = '<tr><td colspan="6">No active dispatches found.</td></tr>';
        appendClientLog('[render] buttons view=0 complete=0 amend=0 cancel=0 disabledView=0 disabledTotal=0');
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
          '<td>' + escapeHtml(dispatch.start_time) + '</td>' +
          '<td>' + escapeHtml(dispatch.truck_numbers) + '</td>' +
          '<td>' + escapeHtml(dispatch.job_number) + '</td>' +
          '<td>' + escapeHtml(dispatch.status) + '</td>' +
          '<td class="actions">' +
            '<button type="button" class="actionBtn" data-action="view" data-dispatch-id="' + escapeHtml(dispatchId) + '" data-doc-url="' + escapeHtml(docUrl) + '" data-doc-id="' + escapeHtml(docId) + '" ' + (hasLinkedDoc ? '' : 'disabled') + '>' + viewButtonLabel + '</button>' +
            '<button type="button" class="actionBtn" data-action="complete" data-dispatch-id="' + escapeHtml(dispatchId) + '" ' + (dispatchId ? '' : 'disabled') + '>Mark Completed</button>' +
            '<button type="button" class="actionBtn" data-action="amend" data-dispatch-id="' + escapeHtml(dispatchId) + '" ' + (dispatchId ? '' : 'disabled') + '>Amend</button>' +
            '<button type="button" class="actionBtn" data-action="cancel" data-dispatch-id="' + escapeHtml(dispatchId) + '" ' + (dispatchId ? '' : 'disabled') + '>Cancel</button>' +
          '</td>' +
        '</tr>';
      }).join('');

      const actionButtons = tableBody.querySelectorAll('button.actionBtn');
      const actionCounts = { view: 0, complete: 0, amend: 0, cancel: 0 };
      let disabledViewCount = 0;
      let disabledTotalCount = 0;

      Array.prototype.forEach.call(actionButtons, function (button) {
        const action = String((button && button.dataset && button.dataset.action) || '').trim();
        if (Object.prototype.hasOwnProperty.call(actionCounts, action)) {
          actionCounts[action] += 1;
        }
        if (button && button.disabled) {
          disabledTotalCount += 1;
          if (action === 'view') {
            disabledViewCount += 1;
          }
        }
      });

      appendClientLog('[render] buttons view=' + actionCounts.view + ' complete=' + actionCounts.complete + ' amend=' + actionCounts.amend + ' cancel=' + actionCounts.cancel + ' disabledView=' + disabledViewCount + ' disabledTotal=' + disabledTotalCount);
    }

    function loadDashboardData(token) {
      const sizeEl = document.getElementById('htmlSizeValue');
      if (sizeEl) {
        sizeEl.textContent = String((document.documentElement && document.documentElement.outerHTML && document.documentElement.outerHTML.length) || 0);
      }

      appendClientLog('[load] calling getAdminDashboardData');

      let settled = false;
      const watchdog = window.setTimeout(function () {
        if (settled) return;
        settled = true;
        showBanner('error', 'Dashboard data load timed out');
        appendClientLog('[load] timeout');
      }, 10000);

      google.script.run
        .withSuccessHandler(function (data) {
          if (settled) return;
          settled = true;
          window.clearTimeout(watchdog);
          const dispatchCount = (data && data.activeDispatches && data.activeDispatches.length) || 0;
          appendClientLog('[load] success, dispatches=' + dispatchCount);
          renderDispatchTable(data || {});
        })
        .withFailureHandler(function (error) {
          if (settled) return;
          settled = true;
          window.clearTimeout(watchdog);
          const fullError = (error && error.stack)
            ? String(error.stack)
            : ((error && error.message) ? String(error.message) : String(error || 'Unknown error'));
          appendClientLog('[load] failure: ' + fullError);
          showBanner('error', fullError);
          const tableBody = document.getElementById('dispatchTableBody');
          if (tableBody) {
            tableBody.innerHTML = '<tr><td colspan="7">Failed to load dashboard data.</td></tr>';
          }
        })
        .getAdminDashboardData(token);
    }


    function assignmentTemplate(index, value) {
      var v = value || {};
      return '<fieldset class="assignment-block" data-assignment-index="' + index + '">'
        + '<legend>Assignment ' + (index + 1) + '</legend>'
        + '<label>Job Number<input type="text" name="assignment_jobNumber_' + index + '" value="' + escapeHtml(String(v.job_number || v.jobNumber || '')) + '"></label>'
        + '<label>Start Time<input type="text" name="assignment_startTime_' + index + '" value="' + escapeHtml(String(v.start_time || v.startTime || '')) + '"></label>'
        + '<label>Start Location<textarea name="assignment_startLocation_' + index + '" rows="2">' + escapeHtml(String(v.start_location || v.startLocation || '')) + '</textarea></label>'
        + '<label>Instructions<textarea name="assignment_instructions_' + index + '" rows="3">' + escapeHtml(String(v.instructions || '')) + '</textarea></label>'
        + '<label>Notes<textarea name="assignment_notes_' + index + '" rows="2">' + escapeHtml(String(v.notes || '')) + '</textarea></label>'
        + '</fieldset>';
    }

    function addAssignmentBlock(value) {
      var container = document.getElementById('assignmentBlocks');
      if (!container) return;
      var index = container.querySelectorAll('.assignment-block').length;
      container.insertAdjacentHTML('beforeend', assignmentTemplate(index, value));
    }

    function hydrateCreateOrEditForm() {
      var form = document.getElementById('createDispatchForm');
      if (!form) return;
      var payload = window.__INITIAL_DISPATCH_PAYLOAD__ || null;
      addAssignmentBlock();
      if (!payload) return;
      form.elements.status.value = payload.status || 'Dispatched';
      form.elements.date.value = payload.date || '';
      form.elements.shift.value = payload.shift || '';
      form.elements.client.value = payload.client || '';
      form.elements.truckNumbers.value = payload.truck_numbers || '';
      form.elements.tollsPolicy.value = payload.tolls_policy || '';
      form.elements.notes.value = payload.notes || '';
      var blocks = document.getElementById('assignmentBlocks');
      if (blocks) blocks.innerHTML = '';
      var assignments = Array.isArray(payload.assignments) && payload.assignments.length ? payload.assignments : [{
        job_number: payload.job_number || '',
        start_time: payload.start_time || '',
        start_location: payload.start_location || '',
        instructions: payload.instructions || '',
        notes: payload.notes || ''
      }];
      assignments.forEach(function (a) { addAssignmentBlock(a); });
      var summaryWrap = document.getElementById('changeSummaryWrap');
      if (summaryWrap) summaryWrap.style.display = 'block';
    }

    function collectAssignments(formData) {
      var assignments = [];
      var index = 0;
      while (formData.has('assignment_jobNumber_' + index) || formData.has('assignment_startTime_' + index) || formData.has('assignment_startLocation_' + index)) {
        assignments.push({
          jobNumber: String(formData.get('assignment_jobNumber_' + index) || '').trim(),
          startTime: String(formData.get('assignment_startTime_' + index) || '').trim(),
          startLocation: String(formData.get('assignment_startLocation_' + index) || '').trim(),
          instructions: String(formData.get('assignment_instructions_' + index) || '').trim(),
          notes: String(formData.get('assignment_notes_' + index) || '').trim()
        });
        index += 1;
      }
      return assignments;
    }

    function resetCreateDispatchForm(form) {
      if (!form) return;
      form.reset();
      if (form.elements.status) {
        form.elements.status.value = 'Confirmed';
      }
      var blocks = document.getElementById('assignmentBlocks');
      if (blocks) {
        blocks.innerHTML = '';
      }
      addAssignmentBlock();
    }

    function submitCreateDispatch(event) {
      event.preventDefault();
      const form = event.target;
      const submitMode = String(window.__submitMode || 'dashboard').trim();
      const formData = new FormData(form);
      const assignments = collectAssignments(formData);
      const payload = {
        status: String(formData.get('status') || '').trim(),
        date: String(formData.get('date') || '').trim(),
        shift: String(formData.get('shift') || '').trim(),
        client: String(formData.get('client') || '').trim(),
        jobNumber: String((assignments[0] && assignments[0].jobNumber) || '').trim(),
        startTime: String((assignments[0] && assignments[0].startTime) || '').trim(),
        startLocation: String((assignments[0] && assignments[0].startLocation) || '').trim(),
        instructions: String((assignments[0] && assignments[0].instructions) || '').trim(),
        notes: String(formData.get('notes') || '').trim(),
        tollsPolicy: String(formData.get('tollsPolicy') || '').trim(),
        truckNumbers: String(formData.get('truckNumbers') || '').trim(),
        assignments: assignments,
        changeSummary: String(formData.get('changeSummary') || '').trim()
      };

      if (!payload.date || !payload.shift || !payload.truckNumbers) {
        showBanner('error', 'Date, shift, and truck numbers are required.');
        window.__submitMode = null;
        return;
      }
      if (payload.status === 'Dispatched' && (!payload.client || !payload.jobNumber || !payload.startTime || !payload.startLocation || !payload.instructions)) {
        showBanner('error', 'Dispatched requires client and assignment details.');
        window.__submitMode = null;
        return;
      }

      const mode = String(form.dataset.formMode || 'create').trim();
      const dispatchId = String(form.dataset.dispatchId || '').trim();
      if (mode === 'edit') {
        if (!payload.changeSummary) {
          showBanner('error', 'Amendment change summary is required.');
          window.__submitMode = null;
          return;
        }
        google.script.run
          .withSuccessHandler(function () {
            showBanner('success', 'Created.');
            window.open(BASE_URL + '?t=' + encodeURIComponent(TOKEN) + '&p=dashboard&msg=edit_ok', '_top');
            window.__submitMode = null;
          })
          .withFailureHandler(function (error) {
            const fullError = (error && error.stack)
              ? String(error.stack)
              : ((error && error.message) ? String(error.message) : String(error || 'Unknown error'));
            showBanner('error', fullError);
            window.__submitMode = null;
          })
          .saveDispatchEditFromDashboard(TOKEN, dispatchId, payload);
        return;
      }

      google.script.run
        .withSuccessHandler(function () {
          if (submitMode === 'dashboard') {
            window.open(BASE_URL + '?t=' + encodeURIComponent(TOKEN) + '&p=dashboard&msg=create_ok', '_top');
            window.__submitMode = null;
            return;
          }
          resetCreateDispatchForm(form);
          showBanner('success', 'Created');
          window.__submitMode = null;
        })
        .withFailureHandler(function (error) {
          const fullError = (error && error.stack)
            ? String(error.stack)
            : ((error && error.message) ? String(error.message) : String(error || 'Unknown error'));
          showBanner('error', fullError);
          window.__submitMode = null;
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

    function repairDocLinks() {
      if (!window.confirm('Run one-time repair for legacy doc links?')) return;
      google.script.run
        .withSuccessHandler(function () {
          goAdmin('dashboard', { msg: 'repair_done' });
        })
        .withFailureHandler(function (error) {
          showBanner('error', (error && error.message) || String(error || 'Unknown error'));
        })
        .repairDispatchDocLinksFromDashboard(TOKEN);
    }

    (function updateHtmlSize() {
      const sizeEl = document.getElementById('htmlSizeValue');
      if (sizeEl) sizeEl.textContent = String((document.documentElement && document.documentElement.outerHTML && document.documentElement.outerHTML.length) || 0);
    })();

    hardenAdminDiagnostics();
    findViewportOverlays();

    if (CURRENT_PAGE === 'dashboard') {
      loadDashboardData(TOKEN);
    }
    if (CURRENT_PAGE === 'create' || CURRENT_PAGE === 'edit') {
      hydrateCreateOrEditForm();
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
 * Escape text for safe display inside an HTML <pre> element.
 *
 * @param {string} value - Raw text.
 * @returns {string} Escaped HTML entities.
 */
function escapeHtmlForPre_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Render a minimal admin-safe page for dashboard script debugging.
 *
 * @param {string} token - Active portal token.
 * @returns {string} Safe HTML markup.
 */
function renderSafeAdminPage_(token) {
  const baseUrl = ScriptApp.getService().getUrl();
  return `<!doctype html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Safe Dashboard</title>
</head>
<body>
  <nav>
    <a href="${baseUrl}?t=${encodeURIComponent(token || '')}&p=dashboard">Dashboard</a>
    <a href="${baseUrl}?t=${encodeURIComponent(token || '')}&p=create">Create Dispatch</a>
    <a href="${baseUrl}?t=${encodeURIComponent(token || '')}&p=companies">Companies/Trucks</a>
    <a href="${baseUrl}?t=${encodeURIComponent(token || '')}&p=users">Users</a>
  </nav>
  <h1>SAFE DASHBOARD</h1>
  <pre id="out">loading...</pre>
  <script>
    (function () {
      var out = document.getElementById('out');
      function write(text) {
        if (out) out.textContent = String(text || '');
      }

      write('boot ok');

      google.script.run
        .withSuccessHandler(function (resp) {
          var dispatches = (resp && resp.activeDispatches && resp.activeDispatches.slice)
            ? resp.activeDispatches.slice(0, 5)
            : [];
          write(JSON.stringify({
            ok: true,
            activeDispatches: dispatches,
            raw: resp
          }, null, 2));
        })
        .withFailureHandler(function (error) {
          var errorText = (error && error.stack)
            ? String(error.stack)
            : ((error && error.message) ? String(error.message) : String(error || 'Unknown error'));
          write('failure:\\n' + errorText);
        })
        .getAdminDashboardData(${JSON.stringify(token || '')});
    })();
  </script>
</body>
</html>`;
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
    .setTitle('PING');
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
        const normalizeText = (value) => String(value || '').trim();
        return {
          dispatch_id: normalizeText(row[headerMap.dispatch_id]),
          date: formatDateNY_(row[headerMap.date]),
          truck_numbers: normalizeText(row[headerMap.truck_numbers]),
          client: normalizeText(row[headerMap.client]),
          job_number: normalizeText(row[headerMap.job_number]),
          start_time: formatTimeNY_(row[headerMap.start_time]),
          status: normalizeText(row[headerMap.status]),
          doc_url: headerMap.doc_url !== undefined ? normalizeText(row[headerMap.doc_url]) : '',
          doc_id: headerMap.doc_id !== undefined ? normalizeText(row[headerMap.doc_id]) : ''
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
  <div id="jsBoot" style="padding:6px;border:2px solid #000;display:inline-block;">JS BOOT PENDING</div>
  <pre id="clientLog">No client log entries.</pre>
  <pre id="clientErrors">No client errors.</pre>
  <script>
    function appendLine(existing, line, emptySentinel) {
      var cur = String(existing || '').trim();
      var prefix = cur && cur !== emptySentinel ? (cur + '\\n') : '';
      return prefix + line;
    }

    try {
      (function () {
        var bootEl = document.getElementById('jsBoot');
        if (bootEl) bootEl.textContent = 'JS BOOT OK';
        var logEl = document.getElementById('clientLog');
        if (logEl) {
          logEl.textContent = appendLine(logEl.textContent, '[boot] script executed at ' + new Date().toISOString(), 'No client log entries.');
        }
      })();
    } catch (bootError) {
      var errorEl = document.getElementById('clientErrors');
      if (errorEl) {
        errorEl.textContent = appendLine(errorEl.textContent, '[boot] ' + (bootError && bootError.message ? bootError.message : String(bootError || 'Unknown boot error')), 'No client errors.');
      }
    }
  </script>
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
    .setTitle('Admin Fallback');
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
 * @returns {{status:string,mode:string,createdCount:number,dispatchIds:string[]}} Server result.
 */
function createDispatchFromDashboard(token, payload) {
  const actor = getAuthorizedDashboardActor_(token);
  return createDispatchFromPortalForm(payload, actor.userId || actor.displayName || 'unknown');
}


/**
 * Dashboard action: load one dispatch for edit form.
 *
 * @param {string} token
 * @param {string} dispatchId
 * @returns {Object}
 */
function getDispatchForDashboardEdit(token, dispatchId) {
  getAuthorizedDashboardActor_(token);
  return getDispatchById_(dispatchId);
}

/**
 * Dashboard action: save full dispatch edits.
 *
 * @param {string} token
 * @param {string} dispatchId
 * @param {Object} payload
 * @returns {{status:string,rowNumber:number}}
 */
function saveDispatchEditFromDashboard(token, dispatchId, payload) {
  const actor = getAuthorizedDashboardActor_(token);
  const updatedBy = actor.userId || actor.displayName || 'unknown';
  return saveDispatchEdit(dispatchId, payload, updatedBy);
}

/**
 * Read current user notifications.
 *
 * @param {string} token
 * @returns {Array<Object>}
 */
function getMyNotifications(token) {
  const actor = getActivePortalUserByToken_(String(token || '').trim());
  if (!actor) throw new Error('Unauthorized.');
  const notifications = getNotificationsForUser(actor.userId) || [];
  if (ENV === 'DEV') {
    log_('notifications_fetch token_user_id=' + String(actor.userId || '').trim() + ' returned=' + notifications.length);
  }
  return notifications;
}

/**
 * Mark one notification as read.
 *
 * @param {string} token
 * @param {string} notificationId
 * @returns {{status:string}}
 */
function markMyNotificationRead(token, notificationId) {
  const actor = getActivePortalUserByToken_(String(token || '').trim());
  if (!actor) throw new Error('Unauthorized.');
  return markNotificationRead(notificationId, actor.userId);
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
function buildTruckScopedPortalHtml_(truckNumber, token) {
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

  function escapeHtmlText_(value) {
    return String(value || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
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
    const startValue = row[headerMap.start_time];
    const dispatchDate = parseSheetDateOrTimeValue_(dateValue) || new Date(0);

    const friendlyDate = formatDateNY_(dateValue);
    const friendlyTime = formatTimeNY_(startValue);

    const jobNumber = String(row[headerMap.job_number] || '').trim();
    const status = String(row[headerMap.status] || '').trim();
    const docUrl = String(row[headerMap.doc_url] || '').trim();
    const docId = headerMap.doc_id === undefined ? '' : String(row[headerMap.doc_id] || '').trim();
    const resolvedDocUrl = String(resolveDocUrl_(docUrl, docId) || '').trim();
    const rawIsConfirmed = row[headerMap.is_confirmed];
    const isConfirmed = rawIsConfirmed === true
      || String(rawIsConfirmed || '').toLowerCase() === 'true';

    const statusClass = status === 'Canceled'
      ? ' status-canceled'
      : (status === 'Amended' ? ' status-amended' : '');
    const statusText = status || 'Unknown';
    const showConfirmButton = !isConfirmed;
    const confirmButton = !dispatchId
      ? '<button type="button" class="confirmBtn" disabled title="Dispatch index missing">Unavailable</button>'
      : showConfirmButton
        ? `<button type="button" class="confirmBtn" data-dispatch-id="${dispatchId}">Confirm receipt</button>`
        : '<button type="button" class="confirmBtn" disabled>Confirmed ✓</button>';
    const viewButton = isValidDocUrl_(resolvedDocUrl)
      ? `<a class="viewBtn" href="${resolvedDocUrl}" target="_blank" rel="noopener">View</a>`
      : '<button type="button" class="viewBtn" disabled>No doc</button>';

    const block = '<div class="dispatch-block">'
      + '<div class="dispatch-meta"><strong>Date:</strong> ' + escapeHtmlText_(friendlyDate) + '</div>'
      + '<div class="dispatch-meta"><strong>Start Time:</strong> ' + escapeHtmlText_(friendlyTime) + '</div>'
      + '<div class="dispatch-meta"><strong>Job Number:</strong> ' + escapeHtmlText_(jobNumber) + '</div>'
      + '<div class="dispatch-meta"><strong>Status:</strong> <span class="status-pill' + statusClass + '">' + escapeHtmlText_(statusText) + '</span></div>'
      + '<div class="dispatch-actions">' + viewButton + confirmButton + '</div>'
      + '</div>';
    const entry = { date: dispatchDate, html: block };

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
      padding: 10px;
      border: 1px solid #e0e0e0;
      border-radius: 8px;
      margin-bottom: 10px;
      background: #fff;
    }
    .dispatch-meta {
      margin: 2px 0;
    }
    .dispatch-actions {
      margin-top: 8px;
      display: flex;
      gap: 8px;
      align-items: center;
    }
    .viewBtn {
      display: inline-block;
      text-decoration: none;
      border: 1px solid #c8d7f2;
      background: #eef3ff;
      color: #1a4fb3;
      border-radius: 4px;
      padding: 4px 8px;
      font-weight: 600;
    }
    .viewBtn[disabled] {
      background: #f3f4f6;
      color: #888;
      border-color: #ddd;
    }
    .status-pill {
      display: inline-block;
      border-radius: 999px;
      padding: 2px 8px;
      background: #eef2ff;
      color: #1e40af;
      font-weight: 700;
    }
    .status-pill.status-canceled {
      background: #fde8e8;
      color: #b42318;
    }
    .status-pill.status-amended {
      background: #fff4cc;
      color: #8a6d00;
    }
    .notification-item {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
      padding: 10px;
      border: 1px solid #e5e7eb;
      border-radius: 8px;
      margin-top: 8px;
    }
    .notification-item.unread {
      background: #eef4ff;
      border-color: #c7dafc;
      font-weight: 700;
    }
    .notification-item .notification-message {
      flex: 1;
      min-width: 0;
    }
    .notification-error {
      color: #b42318;
      font-weight: 700;
    }
    .confirmBtn {
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
    .confirmBtn.confirmed,
    .confirmBtn:disabled {
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

  <div class="section" id="notificationsSection">
    <h2 id="notificationsHeader">Notifications</h2>
    <div id="notifBox">Loading notifications…</div>
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
  <details class="section" style="margin-top:12px;"><summary>Errors</summary><pre id="clientErrors">No client errors.</pre></details>
  <details class="section" style="margin-top:12px;"><summary>Client Log</summary><pre id="clientLog">No client log entries.</pre></details>

  <script>
    const TOKEN = ${JSON.stringify(token || '')};

    function appendLine(existing, line, emptySentinel) {
      const cur = String(existing || '').trim();
      const prefix = cur && cur !== emptySentinel ? (cur + '\n') : '';
      return prefix + line;
    }

    function appendPortalLog(message) {
      const line = '[' + new Date().toISOString() + '] ' + String(message || 'log');
      const el = document.getElementById('clientLog');
      if (el) {
        el.textContent = appendLine(el.textContent, line, 'No client log entries.');
      } else {
        console.log('[portal-notifications] ' + line);
      }
    }

    function appendPortalError(message) {
      const line = '[' + new Date().toISOString() + '] ' + String(message || 'Unknown client error');
      const el = document.getElementById('clientErrors');
      if (el) {
        el.textContent = appendLine(el.textContent, line, 'No client errors.');
      } else {
        console.error('[portal-notifications] ' + line);
      }
    }

    function escapeHtml(value) {
      return String(value || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
    }

    function renderNotifications(items) {
      const host = document.getElementById('notifBox');
      const header = document.getElementById('notificationsHeader');
      if (!host) return;

      const list = Array.isArray(items) ? items.slice() : [];
      list.sort(function (a, b) {
        const aTime = Date.parse(String((a && a.created_at) || '')) || 0;
        const bTime = Date.parse(String((b && b.created_at) || '')) || 0;
        return bTime - aTime;
      });

      const unreadCount = list.filter(function (n) {
        return !(n && n.is_read);
      }).length;

      if (header) {
        header.textContent = unreadCount > 0
          ? ('Notifications (' + unreadCount + ' unread)')
          : 'Notifications (0 unread)';
      }

      if (!list.length) {
        host.textContent = 'No notifications.';
        return;
      }

      host.innerHTML = list.map(function (n) {
        const isRead = Boolean(n && n.is_read);
        const unreadClass = isRead ? '' : ' unread';
        const link = n.dispatch_id ? ('?t=' + encodeURIComponent(TOKEN) + '&dispatch_id=' + encodeURIComponent(n.dispatch_id)) : '#';
        const notificationId = String((n && n.notification_id) || '');
        const message = escapeHtml(String((n && n.message) || 'Notification'));
        const markReadButton = isRead
          ? ''
          : (' <button type="button" class="notifReadBtn" data-notification-id="' + escapeHtml(notificationId) + '">Mark read</button>');
        return '<div class="notification-item' + unreadClass + '"><a class="notification-message" href="' + link + '">' + message + '</a>' + markReadButton + '</div>';
      }).join('');
    }

    function loadNotifications() {
      const host = document.getElementById('notifBox');
      if (host) host.textContent = 'Loading notifications… (requesting)';
      appendClientLog('[notif] start');
      google.script.run
        .withSuccessHandler(function (items) {
          const list = Array.isArray(items) ? items : [];
          appendClientLog('[notif] success count=' + list.length);
          renderNotifications(list);
        })
        .withFailureHandler(function (error) {
          const message = (error && error.message) ? String(error.message) : String(error || 'Unknown error');
          appendClientLog('[notif] failure ' + message);
          appendPortalError('Failed to load notifications: ' + message);
          if (host) host.textContent = 'Failed to load notifications: ' + message;
        })
        .getMyNotifications(TOKEN);
    }


    document.addEventListener('click', function (e) {
      const btn = e && e.target && e.target.closest ? e.target.closest('.notifReadBtn') : null;
      if (!btn) return;
      markRead(String(btn.dataset.notificationId || '').trim());
    }, true);

    function markRead(notificationId) {
      if (!notificationId) return;
      google.script.run
        .withSuccessHandler(function () { loadNotifications(); })
        .withFailureHandler(function () {})
        .markMyNotificationRead(TOKEN, notificationId);
    }

    function appendClientLog(message) {
      appendPortalLog(message);
    }

    function appendClientError(message) {
      appendPortalError(message);
    }

    function reloadDispatchList() {
      window.open(window.location.href, '_top');
    }

    document.addEventListener('click', function (e) {
      const btn = e && e.target && e.target.closest ? e.target.closest('button.confirmBtn') : null;
      if (!btn || btn.disabled) return;
      const dispatchId = String(btn.dataset.dispatchId || '').trim();
      if (!dispatchId) return;
      const originalLabel = btn.textContent;
      btn.disabled = true;
      btn.textContent = 'Confirming...';
      appendClientLog('[confirm-click] ' + dispatchId);
      google.script.run
        .withSuccessHandler(function () {
          appendClientLog('[confirm] ok');
          reloadDispatchList();
        })
        .withFailureHandler(function (err) {
          btn.disabled = false;
          btn.textContent = originalLabel || 'Confirm receipt';
          appendClientError((err && err.message) || String(err));
        })
        .confirmDispatchReceipt(TOKEN, dispatchId);
    }, true);

    window.addEventListener('load', function () {
      setTimeout(loadNotifications, 0);
    });
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
    .setTitle('Dispatch Portal Error');
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
