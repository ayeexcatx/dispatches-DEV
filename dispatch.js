/**
 * This file runs the day-to-day dispatch process.
 *
 * In plain terms, this file does all of these jobs:
 * - reads new form answers for one truck or many trucks
 * - makes saved archive copies of each dispatch
 * - rebuilds each company's dispatch web page
 * - adds a spreadsheet menu button people can click
 */

const ENV = 'DEV';
const DEV_DB_SPREADSHEET_ID = '1BRkmpO0PoYyDVfK5zskSN9UZjAV9cJpcUd76xXLQHss';

const DEV_SCHEMA_HEADERS = {
  Dispatches: [
    'dispatch_id',
    'created_at',
    'date',
    'shift',
    'client',
    'client_name',
    'job_number',
    'start_time',
    'start_location',
    'instructions',
    'notes',
    'tolls_policy',
    'assignments_json',
    'truck_numbers',
    'status',
    'last_updated_at',
    'last_updated_by',
    'change_summary',
    'cancel_reason',
    'doc_id',
    'doc_url',
    'last_confirmed_at',
    'is_confirmed'
  ],
  Notifications: [
    'notification_id',
    'created_at',
    'recipient_user_id',
    'dispatch_id',
    'type',
    'message',
    'is_read',
    'read_at'
  ],
  Trucks: [
    'truckNumber',
    'company',
    'active',
    'updatedAt',
    'notes'
  ],
  Users: [
    'user_id',
    'display_name',
    'role',
    'company',
    'truck_number',
    'token',
    'is_active'
  ],
  Confirmations: [
    'dispatch_id',
    'truck_number',
    'confirmation_type',
    'timestamp',
    'confirmed_by'
  ],
  TimeEntries: [
    'timeEntryId',
    'dispatchId',
    'truckNumber',
    'userId',
    'startTime',
    'endTime',
    'hours',
    'createdAt',
    'notes'
  ],
  AmendmentHistory: [
    'dispatch_id',
    'amended_at',
    'amended_by',
    'change_summary',
    'previous_doc_id',
    'new_doc_id'
  ]
};

/**
 * Emit a log line with an environment prefix so DEV/PROD logs are easier to separate.
 *
 * @param {...*} args - Values to log.
 */
function log_(...args) {
  Logger.log(`[${ENV}] ${args.map(value => String(value)).join(' ')}`);
}

/**
 * Generate a unique dispatch identifier.
 *
 * @returns {string} UUID value for a new dispatch record.
 */
function newDispatchId_() {
  return Utilities.getUuid();
}

/**
 * Safely get a sheet by name from the active spreadsheet.
 *
 * @param {string} name - Exact tab name.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Matching sheet tab.
 */
function getSheet_(name) {
  const ss = SpreadsheetApp.openById(DEV_DB_SPREADSHEET_ID);
  if (!ss) {
    throw new Error(`[${ENV}] No active spreadsheet is available.`);
  }

  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error(`[${ENV}] Required sheet tab "${name}" was not found.`);
  }

  return sheet;
}

/**
 * Ensure all DEV schema tabs exist with header rows.
 *
 * Missing tabs are created and initialized with the corresponding header row.
 * Existing tabs are validated against expected headers and can be force-overwritten.
 *
 * @param {{forceHeaders?: boolean}=} opts - Optional schema repair controls.
 */
function ensureDevSchema_(opts) {
  const ss = SpreadsheetApp.openById(DEV_DB_SPREADSHEET_ID);
  if (!ss) {
    throw new Error(`[${ENV}] Cannot ensure schema because no active spreadsheet is available.`);
  }

  const options = opts || {};
  const forceHeaders = options.forceHeaders === true;
  const updatedSheets = [];

  Object.keys(DEV_SCHEMA_HEADERS).forEach((sheetName) => {
    const expectedHeaders = DEV_SCHEMA_HEADERS[sheetName];
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
      updatedSheets.push(sheetName);
      log_(`Created sheet with headers: ${sheetName}`);
      return;
    }

    const rowWidth = Math.max(sheet.getLastColumn(), expectedHeaders.length);
    const allRowHeaders = sheet.getRange(1, 1, 1, rowWidth).getValues()[0]
      .map(value => String(value || '').trim());
    const existingHeaders = allRowHeaders.slice(0, expectedHeaders.length);
    const normalizedExpected = expectedHeaders.map(value => String(value || '').trim());

    const hasMissingRequiredHeaders = normalizedExpected.some(header => existingHeaders.indexOf(header) === -1);
    const headersMatchExactly = normalizedExpected.every((header, idx) => existingHeaders[idx] === header);
    const hasUnexpectedExtraHeaders = allRowHeaders.slice(expectedHeaders.length).some(header => header !== '');
    const shouldOverwriteHeaders = forceHeaders || hasMissingRequiredHeaders || !headersMatchExactly || hasUnexpectedExtraHeaders;

    if (shouldOverwriteHeaders) {
      sheet.getRange(1, 1, 1, rowWidth).clearContent();
      sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
      updatedSheets.push(sheetName);
      log_(`Updated headers for sheet: ${sheetName}`);
    }
  });

  if (updatedSheets.length) {
    log_(`Schema header updates applied to: ${updatedSheets.join(', ')}`);
  } else {
    log_('Schema headers already match expected values.');
  }
}

/**
 * Append one dispatch row to the Dispatches table.
 *
 * @param {Object} row - Dispatch row values.
 */


/**
 * Build a unique notification ID.
 *
 * @returns {string}
 */
function newNotificationId_() {
  return Utilities.getUuid();
}

/**
 * Split truck numbers from comma/space separated text.
 *
 * @param {string} raw
 * @returns {string[]}
 */
function parseTruckNumbers_(raw) {
  return String(raw || '')
    .split(/[\s,]+/)
    .map((truck) => String(truck || '').trim())
    .filter(Boolean);
}

/**
 * Return active users that should receive dispatch notifications.
 *
 * @param {string[]} truckNumbers
 * @returns {Array<{userId:string,displayName:string,role:string,company:string,truckNumber:string}>}
 */
function getNotificationRecipientsForDispatch_(truckNumbers) {
  const sheet = getSheet_('Users');
  const headerMap = getHeaderMap_(sheet);
  const requiredHeaders = ['user_id', 'display_name', 'role', 'company', 'truck_number', 'is_active'];
  requiredHeaders.forEach((name) => {
    if (headerMap[name] === undefined) {
      throw new Error(`[${ENV}] Users tab is missing required header: ${name}`);
    }
  });

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const normalizedTrucks = (truckNumbers || []).map((truck) => String(truck || '').trim()).filter(Boolean);
  const truckSet = {};
  normalizedTrucks.forEach((truck) => { truckSet[truck] = true; });

  const recipientsByUserId = {};
  for (let rowIndex = 2; rowIndex <= lastRow; rowIndex++) {
    const rowObj = getRowObject_(sheet, rowIndex, headerMap);
    const activeRaw = rowObj.is_active;
    const isActive = activeRaw === true || String(activeRaw || '').toUpperCase() === 'TRUE';
    if (!isActive) continue;

    const role = String(rowObj.role || '').trim().toLowerCase();
    const userId = String(rowObj.user_id || '').trim();
    if (!userId) continue;

    const truckNumber = String(rowObj.truck_number || '').trim();
    const company = String(rowObj.company || '').trim();

    const isTruckRecipient = truckNumber && truckSet[truckNumber];
    const isCompanyOwner = role === 'company_owner' && company;

    let include = isTruckRecipient;
    if (!include && isCompanyOwner) {
      include = normalizedTrucks.some((truck) => {
        return normalizedTrucks.indexOf(truck) !== -1 && company && getCompanyForTruck_(truck) === company;
      });
    }

    if (!include) continue;

    recipientsByUserId[userId] = {
      userId: userId,
      displayName: String(rowObj.display_name || '').trim(),
      role: String(rowObj.role || '').trim(),
      company: company,
      truckNumber: truckNumber
    };
  }

  return Object.keys(recipientsByUserId).map((userId) => recipientsByUserId[userId]);
}

/**
 * Add notification rows for one dispatch event.
 *
 * @param {string} dispatchId
 * @param {string[]} truckNumbers
 * @param {string} type
 * @param {string} message
 */
function createNotificationsForDispatch_(dispatchId, truckNumbers, type, message) {
  const recipients = getNotificationRecipientsForDispatch_(truckNumbers);
  if (!recipients.length) return;

  const sheet = getSheet_('Notifications');
  const headerMap = getHeaderMap_(sheet);
  recipients.forEach((recipient) => {
    appendDispatchRowByHeader_(sheet, headerMap, {
      notification_id: newNotificationId_(),
      created_at: new Date().toISOString(),
      recipient_user_id: recipient.userId,
      dispatch_id: dispatchId,
      type: type,
      message: message,
      is_read: false,
      read_at: ''
    });
  });
}

/**
 * Add one confirmation history row.
 *
 * @param {string} dispatchId
 * @param {string} truckNumber
 * @param {string} confirmationType
 * @param {string} confirmedBy
 */
function appendConfirmationHistory_(dispatchId, truckNumber, confirmationType, confirmedBy) {
  const confirmationSheet = getSheet_('Confirmations');
  confirmationSheet.appendRow([
    dispatchId,
    truckNumber,
    confirmationType,
    new Date().toISOString(),
    confirmedBy || 'unknown'
  ]);
}

/**
 * Find one dispatch row by dispatch_id.
 *
 * @param {string} dispatchId
 * @returns {{sheet:GoogleAppsScript.Spreadsheet.Sheet,rowNumber:number,headerMap:Object<string,number>,rowObject:Object<string,*>}}
 */
function getDispatchRecordById_(dispatchId) {
  const sheet = getSheet_('Dispatches');
  const headerMap = getHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
  for (let rowIndex = 2; rowIndex <= lastRow; rowIndex++) {
    const rowObj = getRowObject_(sheet, rowIndex, headerMap);
    if (String(rowObj.dispatch_id || '').trim() === String(dispatchId || '').trim()) {
      return { sheet: sheet, rowNumber: rowIndex, headerMap: headerMap, rowObject: rowObj };
    }
  }
  throw new Error(`[${ENV}] Dispatch record not found for dispatch_id=${dispatchId}`);
}

/**
 * Map truck to company.
 *
 * @param {string} truckNumber
 * @returns {string}
 */
function getCompanyForTruck_(truckNumber) {
  const sheet = getSheet_('Trucks');
  const headerMap = getHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
  for (let rowIndex = 2; rowIndex <= lastRow; rowIndex++) {
    const rowObj = getRowObject_(sheet, rowIndex, headerMap);
    if (String(rowObj.truckNumber || '').trim() === String(truckNumber || '').trim()) {
      return String(rowObj.company || '').trim();
    }
  }
  return '';
}
function appendDispatchIndexRow_(row) {
  const sheet = getSheet_('Dispatches');
  const headerMap = getHeaderMap_(sheet);
  appendDispatchRowByHeader_(sheet, headerMap, {
    dispatch_id: row.dispatchId,
    created_at: row.createdAt,
    date: row.date,
    shift: row.shift,
    client: row.client,
    client_name: row.clientName || row.client,
    job_number: row.jobNumber,
    start_time: row.startTime,
    start_location: row.startLocation,
    instructions: row.instructions,
    notes: row.notes || '',
    tolls_policy: row.tollsPolicy || '',
    assignments_json: row.assignmentsJson || '[]',
    truck_numbers: row.truckNumbers,
    status: row.status,
    last_updated_at: row.lastUpdatedAt,
    last_updated_by: row.lastUpdatedBy || '',
    change_summary: row.changeSummary || '',
    cancel_reason: row.cancelReason || '',
    doc_id: row.docId,
    doc_url: row.docUrl,
    last_confirmed_at: row.lastConfirmedAt,
    is_confirmed: row.isConfirmed
  });
  log_(`Dispatches row write success dispatch_id=${row.dispatchId} doc_id=${row.docId} row=${sheet.getLastRow()}`);
}


/**
 * Build a header-name to column-index map from row 1.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet to inspect.
 * @returns {Object<string, number>} Header -> 0-based index map.
 */
function getHeaderMap_(sheet) {
  const lastColumn = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const map = {};

  headers.forEach((header, index) => {
    const key = String(header || '').trim();
    if (key) map[key] = index + 1;
  });

  return map;
}

/**
 * Read one row as an object keyed by header names.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowIndex
 * @param {Object<string, number>} headerMap
 * @returns {Object<string, *>}
 */
function getRowObject_(sheet, rowIndex, headerMap) {
  const width = Math.max(sheet.getLastColumn(), Object.keys(headerMap || {}).length, 1);
  const row = sheet.getRange(rowIndex, 1, 1, width).getValues()[0];
  const obj = {};
  Object.keys(headerMap || {}).forEach((header) => {
    obj[header] = row[(headerMap[header] || 1) - 1];
  });
  return obj;
}

/**
 * Update one row by header names.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowIndex
 * @param {Object<string, number>} headerMap
 * @param {Object<string, *>} updatesObject
 */
function setRowValuesByHeader_(sheet, rowIndex, headerMap, updatesObject) {
  Object.keys(updatesObject || {}).forEach((header) => {
    if (!headerMap[header]) return;
    sheet.getRange(rowIndex, headerMap[header]).setValue(updatesObject[header]);
  });
}

/**
 * Append one Dispatches row by header names.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object<string, number>} headerMap
 * @param {Object<string, *>} dispatchObject
 */
function appendDispatchRowByHeader_(sheet, headerMap, dispatchObject) {
  const width = Math.max(sheet.getLastColumn(), Object.keys(headerMap || {}).length, 1);
  const row = new Array(width).fill('');
  Object.keys(dispatchObject || {}).forEach((header) => {
    const column = headerMap[header];
    if (!column) return;
    row[column - 1] = dispatchObject[header];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
}

function getHeaderIndexMap_(sheet) {
  const oneBased = getHeaderMap_(sheet);
  const zeroBased = {};
  Object.keys(oneBased).forEach((key) => {
    zeroBased[key] = oneBased[key] - 1;
  });
  return zeroBased;
}

/**
 * Build lookup by doc_id for records in the Dispatches sheet.
 *
 * @returns {Object<string, {dispatchId:string,isConfirmed:boolean}>} Map keyed by doc_id.
 */
function getDispatchLookupByDocId_() {
  const sheet = getSheet_('Dispatches');
  const headerMap = getHeaderIndexMap_(sheet);
  const requiredHeaders = ['dispatch_id', 'doc_id', 'is_confirmed'];

  requiredHeaders.forEach((name) => {
    if (headerMap[name] === undefined) {
      throw new Error(`[${ENV}] Dispatches tab is missing required header: ${name}`);
    }
  });

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const lookup = {};

  data.forEach((row) => {
    const docId = String(row[headerMap.doc_id] || '').trim();
    if (!docId) return;
    lookup[docId] = {
      dispatchId: String(row[headerMap.dispatch_id] || '').trim(),
      isConfirmed: String(row[headerMap.is_confirmed] || '').toUpperCase() === 'TRUE' || row[headerMap.is_confirmed] === true
    };
  });

  return lookup;
}

/**
 * Update Dispatches row to mark a dispatch as confirmed.
 *
 * @param {string} dispatchId - Dispatch UUID.
 * @returns {{rowNumber:number,status:string}} Updated row metadata.
 */
function markDispatchConfirmed_(dispatchId) {
  const sheet = getSheet_('Dispatches');
  const headerMap = getHeaderMap_(sheet);
  const requiredHeaders = ['dispatch_id', 'status', 'last_confirmed_at', 'is_confirmed'];

  requiredHeaders.forEach((name) => {
    if (headerMap[name] === undefined) {
      throw new Error(`[${ENV}] Dispatches tab is missing required header: ${name}`);
    }
  });

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error(`[${ENV}] Dispatch record not found for dispatch_id=${dispatchId}`);
  }

  let targetRow = -1;
  for (let rowIndex = 2; rowIndex <= lastRow; rowIndex++) {
    const rowObj = getRowObject_(sheet, rowIndex, headerMap);
    if (String(rowObj.dispatch_id || '').trim() === dispatchId) {
      targetRow = rowIndex;
      break;
    }
  }

  if (targetRow === -1) {
    throw new Error(`[${ENV}] Dispatch record not found for dispatch_id=${dispatchId}`);
  }

  const currentStatus = String(getRowObject_(sheet, targetRow, headerMap).status || '').trim();
  setRowValuesByHeader_(sheet, targetRow, headerMap, {
    last_confirmed_at: new Date().toISOString(),
    is_confirmed: true
  });
  return {
    rowNumber: targetRow,
    status: currentStatus
  };
}

/**
 * Update dispatch status and last_updated_at timestamp.
 *
 * @param {string} dispatchId - Dispatch UUID.
 * @param {string} status - New status value.
 * @returns {{rowNumber:number,status:string}} Updated row metadata.
 */
function updateDispatchStatus_(dispatchId, status, opts) {
  const sheet = getSheet_('Dispatches');
  const headerMap = getHeaderMap_(sheet);
  const requiredHeaders = ['dispatch_id', 'status', 'last_updated_at'];

  requiredHeaders.forEach((name) => {
    if (headerMap[name] === undefined) {
      throw new Error(`[${ENV}] Dispatches tab is missing required header: ${name}`);
    }
  });

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error(`[${ENV}] Dispatch record not found for dispatch_id=${dispatchId}`);
  }

  let targetRow = -1;
  for (let rowIndex = 2; rowIndex <= lastRow; rowIndex++) {
    const rowObj = getRowObject_(sheet, rowIndex, headerMap);
    if (String(rowObj.dispatch_id || '').trim() === dispatchId) {
      targetRow = rowIndex;
      break;
    }
  }

  if (targetRow === -1) {
    throw new Error(`[${ENV}] Dispatch record not found for dispatch_id=${dispatchId}`);
  }

  const options = opts || {};
  const actor = String(options.updatedBy || '').trim();
  const changeSummary = String(options.changeSummary || '').trim();
  const cancelReason = String(options.cancelReason || '').trim();

  const updates = {
    status: status,
    last_updated_at: new Date().toISOString(),
    last_updated_by: actor,
    change_summary: changeSummary,
    cancel_reason: cancelReason
  };

  if (status === 'Amended' || status === 'Canceled') {
    updates.is_confirmed = false;
  }

  setRowValuesByHeader_(sheet, targetRow, headerMap, updates);

  return {
    rowNumber: targetRow,
    status: status
  };
}

/**
 * Mark one dispatch as completed from the admin dashboard.
 *
 * @param {string} dispatchId - Dispatch UUID.
 * @returns {{status:string,rowNumber:number}} Result payload.
 */
function markDispatchCompleted(dispatchId, updatedBy) {
  ensureDevSchema_();
  const update = updateDispatchStatus_(dispatchId, 'Completed', {
    updatedBy: updatedBy,
    changeSummary: 'Marked completed'
  });
  log_(`Dispatch marked completed dispatch_id=${dispatchId} row=${update.rowNumber}`);
  return { status: 'ok', rowNumber: update.rowNumber };
}

/**
 * Mark one dispatch as amended from the admin dashboard.
 *
 * @param {string} dispatchId - Dispatch UUID.
 * @returns {{status:string,rowNumber:number}} Result payload.
 */
function amendDispatch(dispatchId, changeSummary, updatedBy) {
  const summary = String(changeSummary || '').trim();
  if (!summary) {
    throw new Error('Amendment change summary is required.');
  }

  ensureDevSchema_();
  const update = updateDispatchStatus_(dispatchId, 'Amended', {
    updatedBy: updatedBy,
    changeSummary: summary
  });
  const record = getDispatchRecordById_(dispatchId);
  const trucks = parseTruckNumbers_(record.rowObject.truck_numbers);
  createNotificationsForDispatch_(dispatchId, trucks, 'amended', `Dispatch amended: ${summary}`);
  log_(`Dispatch marked amended dispatch_id=${dispatchId} row=${update.rowNumber}`);
  return { status: 'ok', rowNumber: update.rowNumber };
}

/**
 * Mark one dispatch as canceled from the admin dashboard.
 *
 * @param {string} dispatchId - Dispatch UUID.
 * @returns {{status:string,rowNumber:number}} Result payload.
 */
function cancelDispatch(dispatchId, cancelReason, updatedBy) {
  const reason = String(cancelReason || '').trim();
  if (!reason) {
    throw new Error('Cancellation reason is required.');
  }

  ensureDevSchema_();
  const update = updateDispatchStatus_(dispatchId, 'Canceled', {
    updatedBy: updatedBy,
    changeSummary: `Canceled: ${reason}`,
    cancelReason: reason
  });
  const record = getDispatchRecordById_(dispatchId);
  const trucks = parseTruckNumbers_(record.rowObject.truck_numbers);
  createNotificationsForDispatch_(dispatchId, trucks, 'canceled', `Dispatch canceled: ${reason}`);
  log_(`Dispatch marked canceled dispatch_id=${dispatchId} row=${update.rowNumber}`);
  return { status: 'ok', rowNumber: update.rowNumber };
}


/**
 * Convert a dispatch status into the confirmation type written to Confirmations.
 *
 * @param {string} status - Current dispatch status from Dispatches tab.
 * @returns {string} Confirmation type for Confirmations tab.
 */
function getConfirmationTypeForStatus_(status) {
  if (status === 'Amended') return 'amended';
  if (status === 'Canceled') return 'canceled';
  if (status === 'Confirmed') return 'confirmed';
  if (status === 'Dispatched') return 'dispatched';
  if (status === 'Completed') return 'completed';
  return 'dispatch';
}

/**
 * Append a confirmation row and mark the dispatch as confirmed.
 *
 * @param {string} dispatchId - Dispatch UUID.
 * @param {string} truckNumber - Truck identity shown in portal.
 * @returns {{status:string}} Result payload for client.
 */
function confirmDispatchReceipt(dispatchId, truckNumber) {
  ensureDevSchema_();

  const activeUserEmail = Session.getActiveUser().getEmail();
  const confirmedBy = activeUserEmail || 'unknown';

  const dispatchUpdate = markDispatchConfirmed_(dispatchId);
  const confirmationType = getConfirmationTypeForStatus_(dispatchUpdate.status);

  appendConfirmationHistory_(dispatchId, truckNumber, confirmationType, confirmedBy);

  log_(`Confirmation success dispatch_id=${dispatchId} truck_number=${truckNumber} confirmation_type=${confirmationType} dispatches_row=${dispatchUpdate.rowNumber}`);

  return { status: 'ok' };
}

/**
 * Looks through a Google Doc and makes a label bold when it finds that exact text.
 *
 * @param {GoogleAppsScript.Document.Body} body - The document text area to search through.
 * @param {string} labelText - The exact words to find and bold, like "Start Time".
 */
function boldLabel(body, labelText) {
  const paragraphs = body.getParagraphs(); // Get every paragraph so we can scan the whole document from top to bottom.
  for (let i = 0; i < paragraphs.length; i++) {
    const textElement = paragraphs[i].editAsText(); // Turn this paragraph into editable text so formatting can be changed.
    const text = textElement.getText(); // Read the words in this paragraph.
    const index = text.indexOf(labelText); // Check whether this paragraph contains the label we are looking for.
    if (index !== -1) {
      // If we found that label, make those words bold so they stand out.
      textElement.setBold(index, index + labelText.length - 1, true);
    }
  }
}

/**
 * Apply bold + underline to an inserted title paragraph without changing any surrounding formatting.
 *
 * @param {GoogleAppsScript.Document.Body} body - The document text area to search through.
 * @param {number} assignmentNumber - Assignment number to format, like 1 or 2.
 */
function formatInsertedTitleParagraph(body, assignmentNumber) {
  const titlePattern = new RegExp(`ASSIGNMENT\\s+${assignmentNumber}`, 'i');
  const paragraphs = body.getParagraphs();
  for (let i = 0; i < paragraphs.length; i++) {
    const textElement = paragraphs[i].editAsText();
    const text = textElement.getText();
    const match = text.match(titlePattern);
    if (match) {
      const index = match.index;
      const endIndex = index + match[0].length - 1;
      textElement.setBold(index, endIndex, true);
      textElement.setUnderline(index, endIndex, true);
    }
  }
}



/**
 * Load dispatch for edit page.
 *
 * @param {string} dispatchId
 * @returns {Object}
 */
function getDispatchById_(dispatchId) {
  const record = getDispatchRecordById_(dispatchId);
  const row = record.rowObject;
  let assignments = [];
  try {
    assignments = JSON.parse(String(row.assignments_json || '[]'));
  } catch (error) {
    assignments = [];
  }
  return {
    dispatch_id: String(row.dispatch_id || '').trim(),
    date: String(row.date || '').trim(),
    shift: String(row.shift || '').trim(),
    client: String(row.client || '').trim(),
    client_name: String(row.client_name || row.client || '').trim(),
    job_number: String(row.job_number || '').trim(),
    start_time: String(row.start_time || '').trim(),
    start_location: String(row.start_location || '').trim(),
    instructions: String(row.instructions || '').trim(),
    notes: String(row.notes || '').trim(),
    tolls_policy: String(row.tolls_policy || '').trim(),
    truck_numbers: String(row.truck_numbers || '').trim(),
    status: String(row.status || '').trim(),
    assignments: assignments
  };
}

/**
 * Save edits on an existing dispatch row.
 *
 * @param {string} dispatchId
 * @param {Object} payload
 * @param {string} updatedBy
 * @returns {{status:string,rowNumber:number}}
 */
function saveDispatchEdit(dispatchId, payload, updatedBy) {
  ensureDevSchema_();
  const summary = String((payload && payload.changeSummary) || '').trim();
  if (!summary) throw new Error('Amendment change summary is required.');

  const record = getDispatchRecordById_(dispatchId);
  const truckNumbers = parseTruckNumbers_(payload && payload.truckNumbers ? payload.truckNumbers : record.rowObject.truck_numbers);
  const assignments = Array.isArray(payload && payload.assignments) ? payload.assignments : [];

  const updates = {
    date: String((payload && payload.date) || record.rowObject.date || '').trim(),
    shift: String((payload && payload.shift) || record.rowObject.shift || '').trim(),
    client: String((payload && payload.client) || record.rowObject.client || '').trim(),
    client_name: String((payload && (payload.clientName || payload.client)) || record.rowObject.client_name || record.rowObject.client || '').trim(),
    job_number: String((payload && payload.jobNumber) || record.rowObject.job_number || '').trim(),
    start_time: String((payload && payload.startTime) || record.rowObject.start_time || '').trim(),
    start_location: String((payload && payload.startLocation) || record.rowObject.start_location || '').trim(),
    instructions: String((payload && payload.instructions) || record.rowObject.instructions || '').trim(),
    notes: String((payload && payload.notes) || record.rowObject.notes || '').trim(),
    tolls_policy: String((payload && payload.tollsPolicy) || record.rowObject.tolls_policy || '').trim(),
    assignments_json: JSON.stringify(assignments),
    truck_numbers: truckNumbers.join(', '),
    status: 'Amended',
    is_confirmed: false,
    change_summary: summary,
    last_updated_at: new Date().toISOString(),
    last_updated_by: String(updatedBy || '').trim() || 'unknown'
  };

  setRowValuesByHeader_(record.sheet, record.rowNumber, record.headerMap, updates);
  createNotificationsForDispatch_(dispatchId, truckNumbers, 'amended', `Dispatch amended: ${summary}`);
  return { status: 'ok', rowNumber: record.rowNumber };
}
/**
 * This runs automatically each time someone submits the dispatch form.
 *
 * If more than one truck was selected, this repeats the same process for each truck:
 * save an archive copy and refresh the company dispatch page.
 *
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e - The submitted form answers.
 */
function onFormSubmit(e) {
  const values = e.values; // Get all answers that were submitted in this form response.
  const actorId = String((e && e.actorId) || '').trim();

  // These positions match the exact order of fields in the Google Form.
  const date = values[1];
  const rawShiftTime = values[2];
  const company = values[3];
  const jobNumber = values[4];
  const rawStartTime = values[5];
  const startLocation = values[6];
  const instructions = values[7];
  const notes = values[8];
  const truckNumbersRaw = values[9]; // Example: "RT03, RT12" when multiple trucks are chosen.
  const tolls = values[10];
  const add01 = values[11];
  const add02 = values[12];
  const startTime02Raw = values[13];
  const startLocation02 = values[14];
  const instructions02 = values[15];
  
  // Split the truck list by commas so we can handle one truck at a time.
  const truckNumbers = truckNumbersRaw.split(',').map(truck => truck.trim()); // This becomes a clean list of truck numbers.

  Logger.log(values);  // Log full form values so staff can troubleshoot bad input later.

  // Log raw times so we can see exactly what the form sent before formatting.
  console.log("Raw Shift Time:", rawShiftTime);
  console.log("Raw Start Time:", rawStartTime);
  console.log("Raw Start Time 02:", startTime02Raw);

  ensureDevSchema_();
  const assignedTruckNumbers = truckNumbers.join(', ');

truckNumbers.forEach(truckNumber => { // Run the full archive flow for each selected truck, one by one.

  // Build a date like 2025-06-01 so file names sort correctly by day.
  const formattedDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");

  // Convert raw time text into readable time values for drivers and office staff.
  const shiftTime = formatTime(rawShiftTime);
  const startTime = formatTime(rawStartTime);
  const startTime02 = formatTime(startTime02Raw);

  // Make a 24-hour HHMM value (like 0615) so archived files sort by start time.
  const sortableStartTime = convertToSortableTime(rawStartTime);

  /**
   * Turns a time into HHMM (for example 6:15 AM -> 0615) for file name sorting.
   * If time cannot be read, use 0000 so the file can still be named safely.
   */
  function convertToSortableTime(rawTime) {
    if (!rawTime) return "0000";
    const match = rawTime.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/i);
    if (!match) return "0000";

    let hour = parseInt(match[1], 10);
    const minutes = match[2];
    const period = match[3] ? match[3].toUpperCase() : "";

    if (period === "PM" && hour < 12) hour += 12;
    if (period === "AM" && hour === 12) hour = 0;

    return `${hour.toString().padStart(2, '0')}${minutes}`; // End of helper that builds sortable HHMM time.
  }

  // Google Doc template used to create each archived dispatch copy.
  const templateId = (jobNumber.toUpperCase() === 'MCRC' || jobNumber.toUpperCase() === 'CMPM')
    ? "1dS1Q0s55yBZHBGEUp4mHkIEroEWmtFPowRBDMLZQ84c"
    : "1S4mTIjbAGT-xTLAabY7wWzy92dkxKuri4FNHOxIV6aQ";

  // Fallback archive folder if this truck does not have its own folder mapped.
  const archiveFolderId = "1Fic0PvyH2B-Dq7P0hYQLsn0jB09qOWLE";

  // Map each truck number to the Drive folder where its archive copies are saved.
  const truckArchiveFolders = {
    "DT02": "1k3LqwTyN4VprEZp4WjAwCutnTPDn76FR",
    "RT03": "1OHpftrKXETwNesEijjrlxO8WGzVEWjvL",
    "RT12": "1pm1UVd42sIAKGvnY6Hv7nqOBb3LPsefN",
    "EXCL1": "-1ixUnhC3UHhu6i5mFzPiFdjXUGpt_Y4tu",
    "WAJA01": "1VofxiA7vNfDUYMBXVAVUvxKI6DQuJPQE",
    "WAJA03": "1SakeHV9M9Ly_Yx1dY_Q6bTazXxW9Oekc",
    "WMBA11": "1kZZym6R-b33j7_2SQtGLw6eiIHzURQYn"
    };  

  // We do not create a brand-new blank doc here.
  // Dispatches are archived by creating a copy from the template doc.

  // ===== Create and fill an archive copy =====
  try {
    const template = DriveApp.getFileById(templateId); // Load the template file from Drive.
    const folderId = truckArchiveFolders[truckNumber] || archiveFolderId; // Save to this truck's folder when mapped, otherwise use the fallback folder.
    const archiveFolder = DriveApp.getFolderById(folderId); // Open the destination Drive folder.


    // Use the same HHMM start time in the archive file name.
    const newName = `${formattedDate}_${sortableStartTime}_Dispatch_${truckNumber}_${jobNumber}`; // File name pattern: YYYY-MM-DD_HHMM_Dispatch_TRUCK_JOB.

    // Create a new archive doc from the template, then open it to fill in details.
    const archiveCopy = template.makeCopy(newName, archiveFolder);
    archiveCopy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // Allow anyone with the link to view this archive copy.
    const archiveDoc = DocumentApp.openById(archiveCopy.getId());
    const archiveBody = archiveDoc.getBody();

    // Fill all required placeholders in the archive copy.
    archiveBody.replaceText("{{DATE}}", date);
    archiveBody.replaceText("{{SHIFT TIME}}", shiftTime + " ");
    archiveBody.replaceText("{{COMPANY}}", company);
    archiveBody.replaceText("{{JOB NUMBER}}", jobNumber);
    archiveBody.replaceText("{{START TIME}}", startTime);
    archiveBody.replaceText("{{START LOCATION}}", startLocation);
    archiveBody.replaceText("{{INSTRUCTIONS}}", instructions);
    archiveBody.replaceText("{{NOTES}}", notes);
    archiveBody.replaceText("{{TRUCK NUMBER}}", truckNumber);
    archiveBody.replaceText("{{TOLLS}}", tolls);

    // Archive ADD 01: add text if present, otherwise remove the placeholder line.
    if (add01.trim()) {
      archiveBody.replaceText("{{ADD 01}}", add01.trim());
    } else {
      const paragraph = findParagraphContaining(archiveBody, "{{ADD 01}}");
      if (paragraph) paragraph.removeFromParent(); // Remove this line completely when field is empty.
    }

    // Archive ADD 02: add text plus spacing if present, otherwise remove the line.
    if (add02.trim()) {
      const para = findParagraphContaining(archiveBody, "{{ADD 02}}");
      if (para) {
        const index = archiveBody.getChildIndex(para);
        archiveBody.insertParagraph(index, ""); // Add an empty line above this section for readability.
      }
      archiveBody.replaceText("{{ADD 02}}", add02.trim());
    } else {
      const paragraph = findParagraphContaining(archiveBody, "{{ADD 02}}");
      if (paragraph) paragraph.removeFromParent();
    }

    // Archive second assignment start time: include only when provided.
    if (startTime02Raw.trim()) {
      const label = "Start Time 02:";
      const formatted = `${label} ${startTime02}`;
      archiveBody.replaceText("{{START TIME 02:}}", formatted);
      boldLabel(archiveBody, label);
    } else {
      const paragraph = findParagraphContaining(archiveBody, "{{START TIME 02:}}");
      if (paragraph) paragraph.removeFromParent();
    }

    // Archive second assignment location: include only when provided.
    if (startLocation02.trim()) {
      const label = "Start Location 02:";
      const formatted = `${label}\n${startLocation02}`;
      archiveBody.replaceText("{{START LOCATION 02:}}", formatted);
      boldLabel(archiveBody, label);
    } else {
      const paragraph = findParagraphContaining(archiveBody, "{{START LOCATION 02:}}");
      if (paragraph) paragraph.removeFromParent();
    }

    // Archive second assignment instructions: include only when provided.
    if (instructions02.trim()) {
      const label = "Instructions 02:";
      const formatted = `${label}\n${instructions02}`;
      archiveBody.replaceText("{{INSTRUCTIONS 02:}}", formatted);
      boldLabel(archiveBody, label);
    } else {
      const paragraph = findParagraphContaining(archiveBody, "{{INSTRUCTIONS 02:}}");
      if (paragraph) paragraph.removeFromParent();
    }

    // Bold key labels in the archive copy so the layout matches the live version.
    boldLabel(archiveBody, "CONFIRM DISPATCH");
    boldLabel(archiveBody, "Start Time");
    boldLabel(archiveBody, "Start Location");
    boldLabel(archiveBody, "Instructions");
    boldLabel(archiveBody, "Start Time 02:");
    boldLabel(archiveBody, "Start Location 02:");
    boldLabel(archiveBody, "Instructions 02:");
    formatInsertedTitleParagraph(archiveBody, 1);
    formatInsertedTitleParagraph(archiveBody, 2);

    archiveDoc.saveAndClose(); // Save all archive changes to Drive.
    console.log(`Archive copy created successfully: ${newName}`);

    const dispatchId = String((e && e.dispatchId) || '').trim() || newDispatchId_();
    const docId = archiveCopy.getId();
    appendDispatchIndexRow_({
      dispatchId: dispatchId,
      createdAt: new Date().toISOString(),
      date: date,
      shift: shiftTime,
      client: company,
      clientName: String((e && e.clientName) || company || '').trim(),
      jobNumber: jobNumber,
      startTime: startTime,
      startLocation: startLocation,
      instructions: instructions,
      notes: String(notes || '').trim(),
      tollsPolicy: String(tolls || '').trim(),
      assignmentsJson: String((e && e.assignmentsJson) || '[]'),
      truckNumbers: assignedTruckNumbers,
      status: 'Dispatched',
      lastUpdatedAt: new Date().toISOString(),
      lastUpdatedBy: actorId || 'system',
      docId: docId,
      docUrl: `https://docs.google.com/document/d/${docId}/edit`,
      lastConfirmedAt: '',
      isConfirmed: false
    });
  } catch (error) {
    console.error("Error creating archive copy:", error); // If archive creation fails, write the error to logs for troubleshooting.
  }

  /**
   * Read a raw time value and return a clean time like 6:15 AM in this script's time zone.
   *
   * @param {string} rawTime - Original time text from the form.
   * @returns {string} A cleaned-up time string, or the original text if it cannot be parsed.
   */
  function formatTime(rawTime) {
    Logger.log("Raw value received in formatTime: " + rawTime);
    if (!rawTime) return "";

    // Try to read times that look like 5:00:00 AM or 17:00.
    const timePattern = /^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/i;
    const match = rawTime.match(timePattern);
    if (!match) return rawTime;  // If the format is unknown, keep the original value.

    let [, hourStr, minuteStr, ampm] = match;
    let hour = parseInt(hourStr, 10);
    const minute = parseInt(minuteStr, 10);

    // Convert AM/PM time into 24-hour values before final formatting.
    if (ampm) {
      const isPM = ampm.toUpperCase() === "PM";
      if (isPM && hour !== 12) hour += 12;
      if (!isPM && hour === 12) hour = 0;
    }

    // Build a temporary Date using today's date plus this hour and minute.
    const date = new Date();
    date.setHours(hour);
    date.setMinutes(minute);
    date.setSeconds(0);
    date.setMilliseconds(0);

    return Utilities.formatDate(date, Session.getScriptTimeZone(), "h:mm a");
  }

  /**
   * Scan each paragraph and return the first one that contains the target words.
   *
   * This is used when we need to remove placeholder lines that were left blank.
   */
  function findParagraphContaining(body, searchText) {
    const paragraphs = body.getParagraphs();
    for (let i = 0; i < paragraphs.length; i++) {
      if (paragraphs[i].getText().includes(searchText)) {
        return paragraphs[i];
      }
    }
    return null;
  }

  /**
   * Read one truck's archive files and build a sorted list of dispatch details.
   *
   * Expected file name shape:
   * YYYY-MM-DD_HHMM_Dispatch_TRUCKNUMBER_JOBNUMBER
   *
   * @param {string} truckNumber - Truck number used to find the right archive folder.
   * @returns {Array<{date:string,time24hr:string,jobNumber:string,url:string}>}
   *   Dispatch info sorted from oldest to newest.
   */
  function getTruckDispatches(truckNumber) {
    const truckFolders = {
    "DT02": "1k3LqwTyN4VprEZp4WjAwCutnTPDn76FR",
    "RT03": "1OHpftrKXETwNesEijjrlxO8WGzVEWjvL",
    "RT12": "1pm1UVd42sIAKGvnY6Hv7nqOBb3LPsefN",
    "EXCL1": "-1ixUnhC3UHhu6i5mFzPiFdjXUGpt_Y4tu",
    "WAJA01": "1VofxiA7vNfDUYMBXVAVUvxKI6DQuJPQE",
    "WAJA03": "1SakeHV9M9Ly_Yx1dY_Q6bTazXxW9Oekc",
    "WMBA11": "1kZZym6R-b33j7_2SQtGLw6eiIHzURQYn"
      };

    const folderId = truckFolders[truckNumber];
    if (!folderId) return [];

    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const dispatches = [];

    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName(); // Example file name: 2025-05-13_0615_Dispatch_RT03_00123
      const url = file.getUrl();

      const match = name.match(/^(\d{4}-\d{2}-\d{2})_(\d{4})_Dispatch_[^_]+_(\d+)/);
      if (match) {
        dispatches.push({
          date: match[1],
          time24hr: match[2],
          jobNumber: match[3],
          url: url
        });
      }
    }

    // Sort by date and time so results are in order.
    dispatches.sort((a, b) => {
      const aKey = `${a.date}_${a.time24hr}`;
      const bKey = `${b.date}_${b.time24hr}`;
      return aKey.localeCompare(bKey);
    });

    return dispatches;
  }

  // Map truck number to company name so we refresh the correct company page only.
  const truckToCompanyMap = {
  'DT02': 'CCG',
  'RT03': 'RTM',
  'RT12': 'RTM',
  'EXCL1': 'EXCL',
  'WAJA01': 'WAJA',
  'WAJA03': 'WAJA',
  'WMBA11': 'WMBA'
  // Add new truck-to-company pairs here when new trucks are added.
};

const companyName = truckToCompanyMap[truckNumber];
if (companyName) {
  updateCompanyDispatchPage(companyName);
} else {
  Logger.log(`Company not found for truck number: ${truckNumber}`);
}

}); // Finished processing all selected trucks.
} // End of the form submit workflow.


/**
 * Create dispatch rows/documents for Confirmed or Dispatched workflows.
 *
 * @param {Object} payload
 * @param {string=} actorId
 * @returns {{status:string,dispatchId:string,mode:string}}
 */
function createDispatchFromPortalForm(payload, actorId) {
  ensureDevSchema_();
  const safePayload = payload || {};
  const mode = String(safePayload.status || 'Dispatched').trim();
  const createdAt = new Date().toISOString();
  const updatedBy = String(actorId || '').trim() || 'system';

  const truckNumbers = parseTruckNumbers_(safePayload.truckNumbers);
  if (!truckNumbers.length) {
    throw new Error('Truck numbers are required.');
  }

  const baseRequired = ['date', 'shift'];
  const missingBase = baseRequired.find((field) => !String(safePayload[field] || '').trim());
  if (missingBase) throw new Error(`Missing required dispatch field: ${missingBase}`);

  const isConfirmedMode = mode === 'Confirmed';
  const isDispatchedMode = mode === 'Dispatched';
  if (!isConfirmedMode && !isDispatchedMode) {
    throw new Error('Status must be Confirmed or Dispatched.');
  }

  const assignments = Array.isArray(safePayload.assignments) ? safePayload.assignments : [];
  const normalizedAssignments = assignments
    .map((item) => ({
      job_number: String((item && item.jobNumber) || '').trim(),
      start_time: String((item && item.startTime) || '').trim(),
      start_location: String((item && item.startLocation) || '').trim(),
      instructions: String((item && item.instructions) || '').trim(),
      notes: String((item && item.notes) || '').trim()
    }))
    .filter((item) => item.job_number || item.start_time || item.start_location || item.instructions || item.notes);

  const defaultAssignment = {
    job_number: String(safePayload.jobNumber || '').trim(),
    start_time: String(safePayload.startTime || '').trim(),
    start_location: String(safePayload.startLocation || '').trim(),
    instructions: String(safePayload.instructions || '').trim(),
    notes: String(safePayload.notes || '').trim()
  };

  if (!normalizedAssignments.length && (defaultAssignment.job_number || defaultAssignment.start_time || defaultAssignment.start_location || defaultAssignment.instructions || defaultAssignment.notes)) {
    normalizedAssignments.push(defaultAssignment);
  }

  if (isDispatchedMode) {
    const required = ['client', 'jobNumber', 'startTime', 'startLocation', 'instructions'];
    const missing = required.find((field) => !String(safePayload[field] || '').trim());
    if (missing) throw new Error(`Missing required dispatch field: ${missing}`);
  }

  const dispatchId = newDispatchId_();
  const baseRow = {
    dispatchId: dispatchId,
    createdAt: createdAt,
    date: String(safePayload.date || '').trim(),
    shift: String(safePayload.shift || '').trim(),
    client: String(safePayload.client || '').trim(),
    clientName: String(safePayload.clientName || safePayload.client || '').trim(),
    jobNumber: String(safePayload.jobNumber || '').trim(),
    startTime: String(safePayload.startTime || '').trim(),
    startLocation: String(safePayload.startLocation || '').trim(),
    instructions: String(safePayload.instructions || '').trim(),
    notes: String(safePayload.notes || '').trim(),
    tollsPolicy: String(safePayload.tollsPolicy || '').trim(),
    assignmentsJson: JSON.stringify(normalizedAssignments),
    truckNumbers: truckNumbers.join(', '),
    status: mode,
    lastUpdatedAt: createdAt,
    lastUpdatedBy: updatedBy,
    changeSummary: '',
    cancelReason: '',
    docId: '',
    docUrl: '',
    lastConfirmedAt: '',
    isConfirmed: false
  };

  if (isConfirmedMode) {
    appendDispatchIndexRow_(baseRow);
    createNotificationsForDispatch_(dispatchId, truckNumbers, 'confirmed_created', `Confirmed dispatch created for ${truckNumbers.join(', ')}`);
    return { status: 'ok', dispatchId: dispatchId, mode: mode };
  }

  const formValues = [];
  formValues[1] = baseRow.date;
  formValues[2] = baseRow.shift;
  formValues[3] = baseRow.client;
  formValues[4] = baseRow.jobNumber;
  formValues[5] = baseRow.startTime;
  formValues[6] = baseRow.startLocation;
  formValues[7] = baseRow.instructions;
  formValues[8] = baseRow.notes;
  formValues[9] = baseRow.truckNumbers;
  formValues[10] = baseRow.tollsPolicy;

  const second = normalizedAssignments[1] || {};
  formValues[11] = second.job_number || '';
  formValues[12] = '';
  formValues[13] = second.start_time || '';
  formValues[14] = second.start_location || '';
  formValues[15] = second.instructions || '';

  onFormSubmit({ values: formValues, actorId: updatedBy, dispatchId: dispatchId, clientName: baseRow.clientName, assignmentsJson: baseRow.assignmentsJson });

  const record = getDispatchRecordById_(dispatchId);
  setRowValuesByHeader_(record.sheet, record.rowNumber, record.headerMap, {
    notes: baseRow.notes,
    tolls_policy: baseRow.tollsPolicy,
    assignments_json: baseRow.assignmentsJson,
    client_name: baseRow.clientName,
    status: 'Dispatched'
  });

  createNotificationsForDispatch_(dispatchId, truckNumbers, 'dispatched_created', `Dispatched job ${baseRow.jobNumber} created for ${truckNumbers.join(', ')}`);
  return { status: 'ok', dispatchId: dispatchId, mode: mode };
}

/**
 * Create or update one company web page
/**
 * Create or update one company web page that lists dispatches by Upcoming, Today, and Past.
 *
 * @param {string} companyName - The company name used to find folders and title the page.
 */
function updateCompanyDispatchPage(companyName) {
  ensureDevSchema_();
  const dispatchLookupByDocId = getDispatchLookupByDocId_();
  Logger.log(`Dispatch lookup key count: ${Object.keys(dispatchLookupByDocId).length}`);

  const folderMap = {
    "CCG": ["DT02"],
    "RTM": ["RT03", "RT12"],
    "EXCL": ["EXCL1"],
    "WAJA": ["WAJA01", "WAJA03"],
    "WMBA": ["WMBA11"]
  };

  const parentFolder = DriveApp.getFolderById("1Fic0PvyH2B-Dq7P0hYQLsn0jB09qOWLE");
  const companyFolder = parentFolder.getFoldersByName(companyName).next();
  const truckFolders = folderMap[companyName];

  let allFiles = [];

  // Collect dispatch files from every truck folder that belongs to this company.
  for (const truckNum of truckFolders) {
    const truckFolder = companyFolder.getFoldersByName(truckNum).next();
    const files = truckFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().endsWith(".html")) continue; // Skip HTML files because they are web pages, not dispatch docs.
      allFiles.push(file);
    }
  }

  // Only show dispatches from the last 18 days on the web page. Older files are kept, not deleted.
  const EIGHTEEN_DAYS_MS = 18 * 24 * 60 * 60 * 1000;
  const now = new Date().getTime();

  const recentFiles = [];
  for (const file of allFiles) {
    const createdTime = file.getDateCreated().getTime();
    if ((now - createdTime) <= EIGHTEEN_DAYS_MS) {
      recentFiles.push(file);
    }
  }
  allFiles = recentFiles; // Use only recent files for the page output.


  // Old cleanup idea below is commented out and does not run.
 // Example old setting: 18 days in milliseconds.
  // Example old value: current time in milliseconds.

  // Example old loop through files.
    // Example old file creation timestamp.
    // Example old condition: file older than 18 days.
      // Example old try block.
        // Example old action: move old file to trash.
        // Example old log message after trashing.
      // Example old catch block.
        // Example old error log.
      //}
    //}
  //}


  // Read date/time from each dispatch file name so we can place each item in the right section.
  const dispatches = allFiles.map(file => {
    const name = file.getName();
    const dateMatch = name.match(/(\d{4}-\d{2}-\d{2})_(\d{4})/);
    const date = dateMatch ? new Date(`${dateMatch[1]}T${dateMatch[2].slice(0,2)}:${dateMatch[2].slice(2)}:00`) : new Date(0);
    return { file, name, date };
  }).sort((a, b) => a.date - b.date);

    // Helper: remove time-of-day so comparisons are based on calendar day only.
    function stripTime(date) {
      return new Date(date.getFullYear(), date.getMonth(), date.getDate());
    }

    const today = stripTime(new Date());

  const upcoming = [], todayList = [], past = [];

  for (const d of dispatches) {
    const url = d.file.getUrl();
    const parts = d.name.split('_');
    const [dispatchDateStr, timeStr, , truckNumber, jobNumberWithExt] = parts;
    const jobNumber = jobNumberWithExt.replace(/\..+$/, '');

    const dispatchDate = new Date(`${dispatchDateStr}T${timeStr.slice(0, 2)}:${timeStr.slice(2)}:00`);
    const friendlyDate = dispatchDate.toLocaleDateString('en-US', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit'
    });
    const friendlyTime = dispatchDate.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' });

  let labelContent = `${friendlyDate} @ ${friendlyTime} – ${truckNumber} – ${jobNumber}`;
  let statusLabel = "";
  let labelStyle = "";

  if (d.name.includes("CANCELLED")) {
    statusLabel = " <span style='color: #d93025; font-weight: bold;'>CANCELLED</span>";
    labelStyle = "color: grey; text-decoration: line-through;";
  } else if (d.name.includes("AMEND")) {
    statusLabel = " <span style='color: #FFBF00; font-weight: bold;'>AMENDMENT</span>";
  }

  const label = `<span style="${labelStyle}">${labelContent}</span>${statusLabel}`;
  const docId = d.file.getId();
  const dispatchRecord = dispatchLookupByDocId[docId] || {};
  const dispatchId = String(dispatchRecord.dispatchId || dispatchRecord.dispatch_id || '').trim();
  const rawStatus = String(dispatchRecord.status || '').trim();
  const normalizedStatus = rawStatus.toLowerCase();
  const rawIsConfirmed = dispatchRecord.isConfirmed !== undefined
    ? dispatchRecord.isConfirmed
    : dispatchRecord.is_confirmed;
  const isConfirmed = rawIsConfirmed === true
    || (typeof rawIsConfirmed === 'string' && rawIsConfirmed.toLowerCase() === 'true');
  // Keep confirm controls visible for unconfirmed dispatches, including amended/canceled items.
  // Only treat truly finalized statuses (for example Completed) as non-confirmable.
  const isFinalizedStatus = normalizedStatus === 'completed';
  const showConfirmButton = !isConfirmed && !isFinalizedStatus;
  const confirmButton = !dispatchId
    ? `<button class="confirm-btn" disabled title="Dispatch index missing">Unavailable</button>`
    : showConfirmButton
      ? `<button class="confirm-btn" data-dispatch-id="${dispatchId}" data-truck-number="${truckNumber}" onclick="confirmReceipt(this)">Confirm Receipt</button>`
      : `<button class="confirm-btn" disabled>${isConfirmed ? 'Confirmed ✓' : 'Finalized'}</button>`;
  const link = `<div class="dispatch-block"><a href="${url}">${label}</a>${confirmButton}</div>`;

    const entry = { date: d.date, html: link };

    const dispatchDay = stripTime(dispatchDate);
    
    if (dispatchDay.getTime() < today.getTime()) {
      past.push(entry);
    } else if (dispatchDay.getTime() === today.getTime()) {
      todayList.push(entry);
    } else {
      upcoming.push(entry);
    }
  }
  upcoming.sort((a, b) => b.date - a.date);
  todayList.sort((a, b) => b.date - a.date);
  past.sort((a, b) => b.date - a.date);

  const upcomingHTML = upcoming.map(e => e.html).join('\n');
  const todayHTML = todayList.map(e => e.html).join('\n');
  const pastHTML = past.map(e => e.html).join('\n');


  // Build the final HTML page that users open to see dispatch links.
  const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>${companyName} Dispatches</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 40px; background: #f9f9f9; }
    h1 {
      font-size: 28px;
      text-align: center;
      padding: 10px 20px;
      display: inline-block;
      background-color: #FFD700; /* Bright yellow background so the title looks like a construction header. */
      border-radius: 25px;
      box-shadow: 0 2px 5px rgba(0,0,0,0.2);
    }
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
      text-decoration: none;
      color: #333;
    }
    .dispatch-block a:hover {
      text-decoration: underline;
    }
    .dispatch-block a {
      color: #1a73e8; /* Blue link color so dispatch links are easy to spot. */
      text-decoration: underline;
      font-weight: bold; /* Make dispatch links bold so they stand out. */
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
      background-color: #FFD700; /* Keep the title bubble in construction-style yellow. */
      padding: 12px 24px;
      border-radius: 999px; /* Rounded shape to make the title look like a pill bubble. */
      font-size: 2em;
      font-weight: bold;
      color: #000;
      box-shadow: 0 2px 6px rgba(0,0,0,0.15);
    }
  </style>
</head>
<body>
    <div class="title-container">
      <h1 class="dispatch-title">${companyName} Dispatches</h1>
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

  const fileName = `${companyName.replace(/\s+/g, '_')}_dispatch_list.html`;
  const existing = companyFolder.getFilesByName(fileName);
  let htmlFile;

  if (existing.hasNext()) {
    htmlFile = existing.next();
    htmlFile.setContent(html);
  } else {
    htmlFile = companyFolder.createFile(fileName, html, MimeType.HTML);
  }

  htmlFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
}


/**
 * When the spreadsheet opens, add a custom menu called Dispatch Tools.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚚 Dispatch Tools')
    .addItem('🔄 Refresh All Company Pages', 'updateAllCompanyPages')
    .addToUi();
}

/**
 * Rebuild dispatch pages for every company (can be run by hourly trigger).
 */
function updateAllCompanyPages() {
  const companies = [
    "CCG",
    "RTM",
    "EXCL",
    "WAJA",
    "WMBA"
  ];

  for (const company of companies) {
    updateCompanyDispatchPage(company);
  }
}

/**
 * Create or rotate a portal token for a truck/company user row.
 *
 * This helper is admin-only and stores a UUID in Users.token.
 *
 * @param {string} truckNumber - Truck number for the user record.
 * @param {string} company - Company code for the user record.
 * @returns {{user_id:string, token:string}} Token payload for sharing.
 */
function createUserToken_(truckNumber, company) {
  const activeEmail = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  const ownerEmail = String(SpreadsheetApp.openById(DEV_DB_SPREADSHEET_ID).getOwner().getEmail() || '').trim().toLowerCase();

  if (!activeEmail || activeEmail !== ownerEmail) {
    throw new Error(`[${ENV}] Only spreadsheet admins may create user tokens.`);
  }

  ensureDevSchema_();

  const sheet = getSheet_('Users');
  const headerMap = getHeaderIndexMap_(sheet);
  const requiredHeaders = ['user_id', 'display_name', 'role', 'company', 'truck_number', 'token', 'is_active'];

  requiredHeaders.forEach((name) => {
    if (headerMap[name] === undefined) {
      throw new Error(`[${ENV}] Users tab is missing required header: ${name}`);
    }
  });

  const normalizedTruck = String(truckNumber || '').trim();
  const normalizedCompany = String(company || '').trim();
  if (!normalizedTruck && !normalizedCompany) {
    throw new Error(`[${ENV}] Provide truckNumber and/or company when creating a user token.`);
  }

  const token = Utilities.getUuid();
  const lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const rowIndex = data.findIndex((row) => {
      const rowTruck = String(row[headerMap.truck_number] || '').trim();
      const rowCompany = String(row[headerMap.company] || '').trim();
      return rowTruck === normalizedTruck && rowCompany === normalizedCompany;
    });

    if (rowIndex !== -1) {
      const targetRow = rowIndex + 2;
      sheet.getRange(targetRow, headerMap.token + 1).setValue(token);
      sheet.getRange(targetRow, headerMap.is_active + 1).setValue(true);
      const userId = String(sheet.getRange(targetRow, headerMap.user_id + 1).getValue() || '').trim() || Utilities.getUuid();
      if (!String(sheet.getRange(targetRow, headerMap.user_id + 1).getValue() || '').trim()) {
        sheet.getRange(targetRow, headerMap.user_id + 1).setValue(userId);
      }
      return { user_id: userId, token };
    }
  }

  const row = new Array(sheet.getLastColumn()).fill('');
  const userId = Utilities.getUuid();
  row[headerMap.user_id] = userId;
  row[headerMap.role] = 'driver';
  row[headerMap.company] = normalizedCompany;
  row[headerMap.truck_number] = normalizedTruck;
  row[headerMap.token] = token;
  row[headerMap.is_active] = true;
  sheet.appendRow(row);

  return { user_id: userId, token };
}

function runCreateDevToken() {
  // CHANGE THESE TWO to something real you want
  const res = createUserToken_('DT02', 'CCG'); 
  Logger.log('TOKEN: ' + JSON.stringify(res));
}

function runRepairDevSchema() {
  ensureDevSchema_({ forceHeaders: true });
}

function repairUsersHeader_DEV_() {
  const EXPECTED = ['user_id', 'display_name', 'role', 'company', 'truck_number', 'token', 'is_active'];

  // This targets the spreadsheet the script is bound to (the one you opened via Extensions → Apps Script)
  const ss = SpreadsheetApp.openById(DEV_DB_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Users') || ss.insertSheet('Users');

  const lastCol = Math.max(sheet.getLastColumn(), EXPECTED.length);
  const current = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());

  // If already correct, stop
  const hasAll = EXPECTED.every(h => current.includes(h));
  if (hasAll && current.slice(0, EXPECTED.length).join('|') === EXPECTED.join('|')) {
    Logger.log('Users headers already correct.');
    return;
  }

  // Snapshot old header row for safety
  Logger.log('Old Users headers: ' + JSON.stringify(current));

  // Overwrite header row (row 1) with expected headers
  sheet.getRange(1, 1, 1, EXPECTED.length).setValues([EXPECTED]);

  // Optional: clear any extra header cells to the right so headerMap can't pick up junk
  if (lastCol > EXPECTED.length) {
    sheet.getRange(1, EXPECTED.length + 1, 1, lastCol - EXPECTED.length).clearContent();
  }

  Logger.log('Users headers repaired to: ' + JSON.stringify(EXPECTED));
}

function runRepairUsersHeader() {
  repairUsersHeader_DEV_();
}

/**
 * Validate a Google Docs document ID.
 *
 * @param {string} value
 * @returns {boolean}
 */
function isValidDocId_(value) {
  return /^[A-Za-z0-9_-]{21,}$/.test(String(value || '').trim());
}

/**
 * Parse a Google Docs document ID from a URL.
 *
 * @param {string} url
 * @returns {string}
 */
function parseDocIdFromUrl_(url) {
  const match = String(url || '').match(/docs\.google\.com\/document\/d\/([A-Za-z0-9_-]{21,})/i);
  return match ? match[1] : '';
}

/**
 * One-time DEV migration to repair Dispatches doc_id/doc_url fields.
 *
 * @returns {{scanned:number,updated:number,movedFromLastUpdatedBy:number,synthesizedFromDocId:number,overwroteInvalidDocUrl:number}}
 */
function repairDispatchDocLinks_DEV_() {
  ensureDevSchema_();
  const sheet = getSheet_('Dispatches');
  const headerMap = getHeaderMap_(sheet);
  const requiredHeaders = ['doc_id', 'doc_url', 'last_updated_by'];
  requiredHeaders.forEach((name) => {
    if (headerMap[name] === undefined) {
      throw new Error(`[${ENV}] Dispatches tab is missing required header: ${name}`);
    }
  });

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { scanned: 0, updated: 0, movedFromLastUpdatedBy: 0, synthesizedFromDocId: 0, overwroteInvalidDocUrl: 0 };
  }

  const stats = { scanned: lastRow - 1, updated: 0, movedFromLastUpdatedBy: 0, synthesizedFromDocId: 0, overwroteInvalidDocUrl: 0 };

  for (let rowIndex = 2; rowIndex <= lastRow; rowIndex++) {
    const row = getRowObject_(sheet, rowIndex, headerMap);
    const currentDocId = String(row.doc_id || '').trim();
    const currentDocUrl = String(row.doc_url || '').trim();
    const currentLastUpdatedBy = String(row.last_updated_by || '').trim();

    const updates = {};
    let effectiveDocId = currentDocId;
    let effectiveDocUrl = currentDocUrl;

    const docUrlLooksValid = /docs\.google\.com\/document\/d\//i.test(currentDocUrl);
    const legacyUrlInActor = /docs\.google\.com\/document\/d\//i.test(currentLastUpdatedBy);

    if ((!currentDocUrl || !docUrlLooksValid) && legacyUrlInActor) {
      const parsedFromActorUrl = parseDocIdFromUrl_(currentLastUpdatedBy);
      updates.doc_url = currentLastUpdatedBy;
      effectiveDocUrl = currentLastUpdatedBy;
      if (parsedFromActorUrl) {
        updates.doc_id = parsedFromActorUrl;
        effectiveDocId = parsedFromActorUrl;
      }
      updates.last_updated_by = 'legacy-system';
      stats.movedFromLastUpdatedBy += 1;
    }

    const docIdToUse = isValidDocId_(effectiveDocId) ? effectiveDocId : '';
    const canonicalDocUrl = docIdToUse ? `https://docs.google.com/document/d/${docIdToUse}/edit` : '';

    if (docIdToUse && (!effectiveDocUrl || !/https?:\/\//i.test(effectiveDocUrl))) {
      updates.doc_url = canonicalDocUrl;
      stats.synthesizedFromDocId += 1;
    }

    if (docIdToUse && effectiveDocUrl && !/https?:\/\//i.test(effectiveDocUrl)) {
      updates.doc_url = canonicalDocUrl;
      stats.overwroteInvalidDocUrl += 1;
    }

    if (Object.keys(updates).length) {
      if (headerMap.last_updated_at !== undefined) {
        updates.last_updated_at = new Date().toISOString();
      }
      setRowValuesByHeader_(sheet, rowIndex, headerMap, updates);
      stats.updated += 1;
    }
  }

  log_(`repairDispatchDocLinks_DEV_ scanned=${stats.scanned} updated=${stats.updated}`);
  return stats;
}


/**
 * Read notifications for a user newest-first.
 *
 * @param {string} userId
 * @returns {Array<Object>}
 */
function getNotificationsForUser(userId) {
  ensureDevSchema_();
  const sheet = getSheet_('Notifications');
  const headerMap = getHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const oneBased = headerMap;
  const idx = {};
  Object.keys(oneBased).forEach((k)=>idx[k]=oneBased[k]-1);
  return rows
    .map((row)=>({
      notification_id: String(row[idx.notification_id] || '').trim(),
      created_at: String(row[idx.created_at] || '').trim(),
      recipient_user_id: String(row[idx.recipient_user_id] || '').trim(),
      dispatch_id: String(row[idx.dispatch_id] || '').trim(),
      type: String(row[idx.type] || '').trim(),
      message: String(row[idx.message] || '').trim(),
      is_read: row[idx.is_read] === true || String(row[idx.is_read] || '').toLowerCase() === 'true',
      read_at: String(row[idx.read_at] || '').trim()
    }))
    .filter((n)=>n.recipient_user_id === String(userId || '').trim())
    .sort((a,b)=>String(b.created_at).localeCompare(String(a.created_at)));
}

/**
 * Mark one notification as read.
 *
 * @param {string} notificationId
 * @param {string} userId
 * @returns {{status:string}}
 */
function markNotificationRead(notificationId, userId) {
  ensureDevSchema_();
  const sheet = getSheet_('Notifications');
  const headerMap = getHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
  for (let rowIndex = 2; rowIndex <= lastRow; rowIndex++) {
    const rowObj = getRowObject_(sheet, rowIndex, headerMap);
    if (String(rowObj.notification_id || '').trim() !== String(notificationId || '').trim()) continue;
    if (String(rowObj.recipient_user_id || '').trim() !== String(userId || '').trim()) continue;
    setRowValuesByHeader_(sheet, rowIndex, headerMap, { is_read: true, read_at: new Date().toISOString() });
    return { status: 'ok' };
  }
  throw new Error('Notification not found.');
}
