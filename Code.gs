/**
 * Butler Property Management (BPM) — Production Apps Script Backend
 *
 * Handles:
 *  - Public Tenant Maintenance Request Portal (?view=request)
 *  - PIN / Email OTP Session Authentication & Server-Side RBAC
 *  - Closed-Loop Work Orders -> Work Logs / 0-Hr Expenses -> Billing -> Payments
 *  - 5-Tier Billable Rate Resolution & Custom Flat Override
 *  - Invoicing, Multi-Property Owner Billing, Monthly Owner Settlements, Gmail PDF & Lump-Sum Payments
 *  - Academic (May-Apr) & Calendar (Jan-Dec) Rent Roll with Prepayment Rollover & Deposit Settlements
 *  - Hourly + 1099 Payroll, Out-of-Pocket Reimbursements & Paystub PDFs
 *  - Automated Weekly Monday "Nothing Missed" Digest & Safe Schema-Mapped Archival
 */

const BPM_DEFAULT_SPREADSHEET_ID = '1bELSUI7xnIyqNxUPMOFTU8uRc9sGsIhkEre6LkVenmI';
const SESSION_TTL_DAYS = 30;
const OTP_TTL_SECONDS = 900; // 15 minutes

const SCHEMA = {
  sys_Settings: ['Setting_Key', 'Setting_Value', 'Description'],
  dim_Owners: [
    'Owner_ID', 'Owner_Name', 'Company_Name', 'Email', 'CC_Emails', 'Phone',
    'Billing_Mode', 'Mgmt_Fee_Pct', 'Default_Rate_Override', 'Material_Markup_Pct',
    'Payment_Terms_Days', 'Notes', 'Status'
  ],
  Properties: [
    'Property_ID', 'Property_Name', 'Owner_ID', 'Owner_Name', 'Address', 'Unit_Type',
    'Billing_Mode_Override', 'Mgmt_Fee_Override_Pct', 'Default_Lease_Cycle',
    'Aliases', 'Gate_Lock_Notes', 'Status'
  ],
  Tasks: [
    'Task_ID', 'Task_Name', 'Default_Billable_Rate', 'Default_Billing_Unit',
    'Category', 'Description', 'Active'
  ],
  dim_Employees: [
    'Employee_ID', 'Name', 'Email', 'Phone', 'Pay_Type', 'Default_Pay_Rate',
    'Role', 'Pin_Hash', 'Active', 'Default_Billable_Rate_Override', 'Notes'
  ],
  dim_Billing_Rates: [
    'Rate_ID', 'Task_ID', 'Task_Name', 'Scope_Type', 'Scope_ID', 'Scope_Name',
    'Rate_Type', 'Billable_Rate', 'Effective_Date', 'Notes'
  ],
  sys_Sessions: [
    'Token_Hash', 'Employee_ID', 'Email', 'Role', 'Expires_At', 'Created_At', 'Last_Seen'
  ],
  fact_Work_Orders: [
    'WO_ID', 'Created_At', 'Source', 'Property_ID', 'Property_Name', 'Unit_Room',
    'Requester_Name', 'Requester_Phone', 'Requester_Email', 'Permission_To_Enter',
    'Urgency', 'Task_Category', 'Title', 'Description', 'Photo_URLs',
    'Assigned_Employee_ID', 'Assigned_To_Name', 'Due_Date', 'Status',
    'Completed_At', 'Resolution_Notes', 'NTE_Limit'
  ],
  fact_Work_Logs: [
    'Log_ID', 'WO_ID', 'Work_Date', 'Timestamp', 'Employee_ID', 'Employee_Name',
    'Property_ID', 'Property_Name', 'Task_Name', 'Hours', 'Start_Time', 'End_Time',
    'Notes', 'Work_Type', 'Payroll_Status', 'Payroll_ID', 'Employee_Rate', 'Gross_Pay',
    'Billing_Status', 'Invoice_ID', 'Charge_Target', 'Billable_Rate',
    'Calculated_Billable', 'Charge_Amount_Custom', 'Rate_Source', 'Edited_By', 'Edited_At'
  ],
  fact_Material_Expenses: [
    'Expense_ID', 'WO_ID', 'Log_ID', 'Expense_Date', 'Timestamp', 'Employee_ID',
    'Employee_Name', 'Property_ID', 'Property_Name', 'Vendor', 'Item_Description',
    'Quantity', 'Unit_Cost', 'Total_Cost', 'Markup_Pct', 'Billable_Amount',
    'Paid_By', 'Charge_Target', 'Receipt_URL', 'Payroll_Status', 'Payroll_ID',
    'Billing_Status', 'Invoice_ID', 'Notes'
  ],
  fact_Invoices: [
    'Invoice_ID', 'Invoice_Number', 'Doc_Type', 'Owner_ID', 'Owner_Name',
    'Recipient_Email', 'Property_ID', 'Property_Name', 'Period_Start', 'Period_End',
    'Issue_Date', 'Due_Date', 'Status', 'Labor_Subtotal', 'Materials_Subtotal',
    'Rent_Collected', 'Mgmt_Fee_Amount', 'Adjustments_Amount', 'Total_Amount',
    'Amount_Paid', 'Balance_Due', 'Sent_At', 'Last_Reminder_At', 'PDF_URL',
    'Line_Items_JSON', 'Notes_To_Client', 'Created_By'
  ],
  fact_Invoice_Payments: [
    'Payment_ID', 'Invoice_ID', 'Invoice_Number', 'Owner_ID', 'Payment_Date',
    'Amount', 'Payment_Method', 'Reference_Number', 'Batch_Payment_ID',
    'Recorded_By', 'Notes', 'Created_At'
  ],
  fact_Leases: [
    'Lease_ID', 'Property_ID', 'Property_Name', 'Unit_Label', 'Lease_Cycle',
    'Cycle_Label', 'Start_Date', 'End_Date', 'Monthly_Rent_Total', 'Deposit_Total',
    'Grace_Day', 'Late_Fee_Amount', 'Proration_Notes', 'Status', 'Updated_At'
  ],
  dim_Tenants: [
    'Tenant_ID', 'Lease_ID', 'Property_ID', 'Property_Name', 'Name', 'Room_Unit',
    'Email', 'Phone', 'Rent_Share', 'Deposit_Required', 'Deposit_Paid',
    'Deposit_Paid_Date', 'Lease_Override_Months_JSON', 'Move_In_Date',
    'Move_Out_Date', 'Guarantor_Info', 'Status'
  ],
  fact_Rent_Ledger: [
    'Ledger_ID', 'Lease_ID', 'Tenant_ID', 'Tenant_Name', 'Property_ID', 'Property_Name',
    'Entry_Date', 'Period_Month', 'Months_Covered', 'Covered_Months_JSON',
    'Entry_Type', 'Amount', 'Payment_Method', 'Reference_Note', 'Recorded_By', 'Created_At'
  ],
  fact_Payroll_History: [
    'Payroll_ID', 'Run_Date', 'Pay_Period_Start', 'Pay_Period_End', 'Employee_ID',
    'Employee_Name', 'Pay_Type', 'Total_Hours', 'Regular_Pay', 'Reimbursements_Pay',
    'Adjustment_Amount', 'Adjustment_Note', 'Net_Payout', 'Payment_Method',
    'Reference_Number', 'Paystub_PDF_URL', 'Work_Log_IDs_JSON', 'Expense_IDs_JSON',
    'Paid_By', 'Status'
  ],
  archive_fact_Work_Logs: [],
  archive_fact_Material_Expenses: [],
  archive_fact_Work_Orders: [],
  Log: ['Timestamp', 'Actor_Email', 'Action', 'Entity_Type', 'Entity_ID', 'Details']
};

SCHEMA.archive_fact_Work_Logs = SCHEMA.fact_Work_Logs.slice();
SCHEMA.archive_fact_Material_Expenses = SCHEMA.fact_Material_Expenses.slice();
SCHEMA.archive_fact_Work_Orders = SCHEMA.fact_Work_Orders.slice();

const DEFAULT_SETTINGS = {
  COMPANY_NAME: 'Butler Property Management, LLC',
  COMPANY_ADDRESS: '702 Clarendon Place, Indianapolis, IN 46208',
  COMPANY_PHONE: '(317) 432-0054',
  DISPATCH_EMAIL: 'butlerpropmgmt@gmail.com',
  DEFAULT_FALLBACK_RATE: '65',
  DEFAULT_MGMT_FEE_PCT: '8',
  DEFAULT_MATERIAL_MARKUP_PCT: '0',
  DEFAULT_GRACE_DAY: '5',
  DEFAULT_LATE_FEE: '50',
  INVOICE_DUE_DAYS: '15',
  PAYMENT_INSTRUCTIONS:
    'Pay via ACH, Check payable to Butler Property Management, LLC (702 Clarendon Place, Indianapolis, IN 46208), or Zelle/Venmo to butlerpropmgmt@gmail.com.',
  PAYMENT_URL: '',
  RECEIPTS_FOLDER_ID: '',
  INVOICES_FOLDER_ID: '',
  NEXT_INVOICE_SEQ: '1001',
  NEXT_WO_SEQ: '1001'
};

const TENANT_CATEGORIES = [
  'Plumbing',
  'Electrical',
  'HVAC / Heating & Cooling',
  'Appliance',
  'Doors / Locks / Keys',
  'Pest Control',
  'Exterior / Lawn / Snow',
  'General Maintenance'
];

const MONTH_NAMES_SHORT = [
  'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
];

// ============================================================================
// WEB ENTRYPOINTS & SCRIPT TRIGGERS
// ============================================================================

/**
 * Web app HTTP GET entrypoint supporting ?view=request for tenant maintenance intake.
 */
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const viewParam = String(params.view || '').toLowerCase();
  const initialView = (viewParam === 'request' || viewParam === 'intake') ? 'request' : 'app';
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.initialView = initialView;
  tpl.presetWoId = String(params.wo || '');
  tpl.presetProperty = String(params.property || '');
  let serviceUrl = '';
  try {
    serviceUrl = ScriptApp.getService().getUrl() || '';
  } catch (err) {
    serviceUrl = '';
  }
  tpl.webAppUrl = serviceUrl;
  return tpl
    .evaluate()
    .setTitle('Butler Property Management')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

/**
 * Includes an HTML partial (Stylesheet or JavaScript) inside Index.html.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Adds a Spreadsheet custom menu for quick database setup and weekly digest runs.
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('BPM Operations')
      .addItem('Initialize / Heal Database Schema', 'setupDatabase')
      .addItem('Run Weekly "Nothing Missed" Digest Now', 'runWeeklyDigestTrigger')
      .addItem('Install Weekly Automation Triggers', 'installAutomationTriggersFromMenu')
      .addToUi();
  } catch (err) {
    // Ignore when executed headlessly
  }
}

function installAutomationTriggersFromMenu() {
  installTriggersInternal();
}

// ============================================================================
// CORE HELPERS: LOCKING, WIRE SANITIZATION, DATES, HASHING & TABLE STORAGE
// ============================================================================

function getSS() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  let id = BPM_DEFAULT_SPREADSHEET_ID;
  try {
    if (typeof PropertiesService !== 'undefined') {
      const propId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
      if (propId) id = propId;
    }
  } catch (err) {
    // Fallback to default constant
  }
  return SpreadsheetApp.openById(id);
}

function withScriptLock(fn) {
  if (typeof LockService === 'undefined' || !LockService.getScriptLock) {
    return fn();
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    return fn();
  } finally {
    try {
      lock.releaseLock();
    } catch (err) {
      // Ignore release failure
    }
  }
}

function round2(val) {
  const n = Number(val || 0);
  if (!isFinite(n)) return 0;
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

function toNum(val, fallback) {
  if (val === '' || val === null || val === undefined) return fallback !== undefined ? fallback : 0;
  const cleaned = typeof val === 'string' ? val.replace(/[$,\s]/g, '') : val;
  const n = Number(cleaned);
  return isFinite(n) ? n : (fallback !== undefined ? fallback : 0);
}

function toBool(val, defaultVal) {
  if (val === true || val === 'TRUE' || val === 'true' || val === 'Yes' || val === '1' || val === 1 || val === 'Active') {
    return true;
  }
  if (val === false || val === 'FALSE' || val === 'false' || val === 'No' || val === '0' || val === 0 || val === 'Inactive') {
    return false;
  }
  return defaultVal !== undefined ? defaultVal : false;
}

function nowIso() {
  return new Date().toISOString();
}

function getTz() {
  try {
    return getSS().getSpreadsheetTimeZone() || 'America/New_York';
  } catch (err) {
    return 'America/New_York';
  }
}

function todayYmd() {
  return toYmd(new Date());
}

function toYmd(val) {
  if (!val) return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return Utilities.formatDate(val, getTz(), 'yyyy-MM-dd');
  }
  const s = String(val).trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) return s.slice(0, 10);
  const slashMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (slashMatch) {
    const mm = slashMatch[1].padStart(2, '0');
    const dd = slashMatch[2].padStart(2, '0');
    const yrRaw = slashMatch[3];
    const yyyy = yrRaw.length === 2 ? '20' + yrRaw : yrRaw;
    return yyyy + '-' + mm + '-' + dd;
  }
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, getTz(), 'yyyy-MM-dd');
  }
  return s;
}

function toMonthKey(val) {
  if (!val) return '';
  const s = String(val).trim();
  if (/^\d{4}-\d{2}$/.test(s)) return s;
  const ymd = toYmd(val);
  return /^\d{4}-\d{2}/.test(ymd) ? ymd.slice(0, 7) : '';
}

function addDaysYmd(ymdStr, days) {
  const ymd = toYmd(ymdStr) || todayYmd();
  const parts = ymd.split('-').map(Number);
  const dt = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  dt.setUTCDate(dt.getUTCDate() + Number(days || 0));
  return dt.toISOString().slice(0, 10);
}

function formatMonthLabel(monthKey) {
  const m = String(monthKey || '').match(/^(\d{4})-(\d{2})$/);
  if (!m) return monthKey || '';
  const idx = Number(m[2]) - 1;
  return (MONTH_NAMES_SHORT[idx] || m[2]) + ' \'' + m[1].slice(2);
}

function escapeHtml(str) {
  return String(str === null || str === undefined ? '' : str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function sanitizeForWire(val) {
  if (val === null || val === undefined) return val;
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? '' : val.toISOString();
  }
  if (Array.isArray(val)) {
    return val.map(sanitizeForWire);
  }
  if (typeof val === 'object') {
    const out = {};
    const keys = Object.keys(val);
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      if (k === '_rowIndex') continue;
      out[k] = sanitizeForWire(val[k]);
    }
    return out;
  }
  return val;
}

function sha256Hex(str) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(str || ''), Utilities.Charset.UTF_8);
  let hex = '';
  for (let i = 0; i < digest.length; i++) {
    const b = (digest[i] + 256) % 256;
    hex += (b < 16 ? '0' : '') + b.toString(16);
  }
  return hex;
}

function newId(prefix) {
  return prefix + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
}

function isDateHeader(headerName) {
  return (
    /_Date$/.test(headerName) ||
    headerName === 'Pay_Period_Start' ||
    headerName === 'Pay_Period_End' ||
    headerName === 'Period_Start' ||
    headerName === 'Period_End'
  );
}

function ensureSheetSchema(ss, sheetName, expectedHeaders) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h || '').trim();
  });
  const missing = expectedHeaders.filter(function (h) {
    return currentHeaders.indexOf(h) === -1;
  });
  if (missing.length > 0) {
    sheet.getRange(1, currentHeaders.length + 1, 1, missing.length).setValues([missing]);
  }
  sheet.setFrozenRows(1);
  return sheet;
}

/**
 * Creates missing tabs, heals missing columns on existing tabs, and seeds baseline settings.
 */
function setupDatabase() {
  return withScriptLock(function () {
    const ss = getSS();
    const names = Object.keys(SCHEMA);
    for (let i = 0; i < names.length; i++) {
      ensureSheetSchema(ss, names[i], SCHEMA[names[i]]);
    }
    const settingsTbl = getTableData('sys_Settings');
    const existingKeys = {};
    settingsTbl.rows.forEach(function (r) {
      existingKeys[String(r.Setting_Key)] = true;
    });
    const toSeed = [];
    Object.keys(DEFAULT_SETTINGS).forEach(function (k) {
      if (!existingKeys[k]) {
        toSeed.push({
          Setting_Key: k,
          Setting_Value: DEFAULT_SETTINGS[k],
          Description: 'BPM default configuration'
        });
      }
    });
    if (toSeed.length > 0) {
      appendRecords('sys_Settings', toSeed);
    }

    const empTbl = getTableData('dim_Employees');
    if (empTbl.rows.length === 0) {
      let activeEmail = '';
      try {
        activeEmail = Session.getActiveUser().getEmail() || '';
      } catch (err) {
        activeEmail = '';
      }
      if (activeEmail) {
        appendRecords('dim_Employees', [{
          Employee_ID: 'EMP-001',
          Name: 'Owner',
          Email: activeEmail,
          Phone: '',
          Pay_Type: 'Hourly',
          Default_Pay_Rate: 35,
          Role: 'Owner',
          Pin_Hash: '',
          Active: 'TRUE',
          Default_Billable_Rate_Override: '',
          Notes: 'Bootstrapped owner account'
        }]);
      }
    }
    return { ok: true, sheetsVerified: names.length };
  });
}

function getTableData(sheetName) {
  const ss = getSS();
  const expected = SCHEMA[sheetName];
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() === 0) {
    sheet = ensureSheetSchema(ss, sheetName, expected || []);
  }
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), expected ? expected.length : 1);
  if (lastRow === 0) {
    return { sheet: sheet, headers: expected || [], colMap: {}, rows: [] };
  }
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(function (h) {
    return String(h || '').trim();
  });
  const colMap = {};
  for (let c = 0; c < headers.length; c++) {
    if (headers[c]) colMap[headers[c]] = c;
  }
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const rawRow = values[r];
    let hasAny = false;
    for (let c = 0; c < rawRow.length; c++) {
      if (rawRow[c] !== '' && rawRow[c] !== null && rawRow[c] !== undefined) {
        hasAny = true;
        break;
      }
    }
    if (!hasAny) continue;
    const obj = { _rowIndex: r + 1 };
    for (let c = 0; c < headers.length; c++) {
      const h = headers[c];
      if (!h) continue;
      let cell = rawRow[c];
      if (cell instanceof Date) {
        cell = isDateHeader(h) ? toYmd(cell) : cell.toISOString();
      } else if (isDateHeader(h) && typeof cell === 'string' && cell.trim()) {
        cell = toYmd(cell);
      }
      obj[h] = cell;
    }
    rows.push(obj);
  }
  return { sheet: sheet, headers: headers, colMap: colMap, rows: rows };
}

function mapObjectToRow(headers, obj) {
  return headers.map(function (h) {
    const v = obj[h];
    return v === undefined || v === null ? '' : v;
  });
}

function appendRecords(sheetName, records) {
  if (!records || records.length === 0) return [];
  const ss = getSS();
  const expected = SCHEMA[sheetName];
  const sheet = ensureSheetSchema(ss, sheetName, expected || Object.keys(records[0]));
  const lastCol = Math.max(sheet.getLastColumn(), expected ? expected.length : 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h || '').trim();
  });
  const rows2d = records.map(function (rec) {
    return mapObjectToRow(headers, rec);
  });
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows2d.length, headers.length).setValues(rows2d);
  return records;
}

function updateRowById(sheetName, idHeader, idVal, patchObj) {
  const tbl = getTableData(sheetName);
  const target = String(idVal);
  for (let i = 0; i < tbl.rows.length; i++) {
    const row = tbl.rows[i];
    if (String(row[idHeader]) === target) {
      Object.keys(patchObj).forEach(function (k) {
        row[k] = patchObj[k];
      });
      const rowArr = mapObjectToRow(tbl.headers, row);
      tbl.sheet.getRange(row._rowIndex, 1, 1, tbl.headers.length).setValues([rowArr]);
      return row;
    }
  }
  return null;
}

function batchPatchRowsById(sheetName, idHeader, patchesById) {
  const ids = Object.keys(patchesById || {});
  if (ids.length === 0) return 0;
  const tbl = getTableData(sheetName);
  let count = 0;
  for (let i = 0; i < tbl.rows.length; i++) {
    const row = tbl.rows[i];
    const id = String(row[idHeader]);
    if (Object.prototype.hasOwnProperty.call(patchesById, id)) {
      const patch = patchesById[id];
      Object.keys(patch).forEach(function (k) {
        row[k] = patch[k];
      });
      const rowArr = mapObjectToRow(tbl.headers, row);
      tbl.sheet.getRange(row._rowIndex, 1, 1, tbl.headers.length).setValues([rowArr]);
      count++;
    }
  }
  return count;
}

function deleteRowsWhere(sheetName, predicate) {
  const tbl = getTableData(sheetName);
  const toDelete = [];
  for (let i = 0; i < tbl.rows.length; i++) {
    if (predicate(tbl.rows[i])) toDelete.push(tbl.rows[i]._rowIndex);
  }
  for (let j = toDelete.length - 1; j >= 0; j--) {
    tbl.sheet.deleteRow(toDelete[j]);
  }
  return toDelete.length;
}

function addAuditLog(actorEmail, action, entityType, entityId, details) {
  try {
    appendRecords('Log', [{
      Timestamp: nowIso(),
      Actor_Email: actorEmail || 'system',
      Action: action,
      Entity_Type: entityType || '',
      Entity_ID: entityId || '',
      Details: typeof details === 'string' ? details : JSON.stringify(details || {})
    }]);
  } catch (err) {
    // Non-fatal
  }
}

function getSettingsMap() {
  const tbl = getTableData('sys_Settings');
  const map = Object.assign({}, DEFAULT_SETTINGS);
  tbl.rows.forEach(function (r) {
    const k = String(r.Setting_Key || '').trim();
    if (k) map[k] = r.Setting_Value === undefined || r.Setting_Value === null ? '' : String(r.Setting_Value);
  });
  return map;
}

function saveSettingValue(key, value, description) {
  const tbl = getTableData('sys_Settings');
  const existing = tbl.rows.find(function (r) {
    return String(r.Setting_Key) === String(key);
  });
  if (existing) {
    updateRowById('sys_Settings', 'Setting_Key', key, {
      Setting_Value: String(value),
      Description: description || existing.Description || ''
    });
  } else {
    appendRecords('sys_Settings', [{
      Setting_Key: String(key),
      Setting_Value: String(value),
      Description: description || ''
    }]);
  }
}

function nextSeqId(prefix, settingKey, defaultStart) {
  const settings = getSettingsMap();
  const current = parseInt(settings[settingKey] || defaultStart || 1001, 10) || 1001;
  saveSettingValue(settingKey, String(current + 1), 'Auto-increment sequence');
  if (prefix === 'BPM') {
    const yr = String(new Date().getFullYear());
    return 'BPM-' + yr + '-' + String(current).padStart(4, '0');
  }
  return prefix + '-' + String(current);
}

// ============================================================================
// DRIVE & PDF HELPERS
// ============================================================================

function getOrCreateFolderByKey(settingKey, folderName) {
  const settings = getSettingsMap();
  const cachedId = settings[settingKey];
  if (cachedId) {
    try {
      return DriveApp.getFolderById(cachedId);
    } catch (err) {
      // Fall through to find/create
    }
  }
  const rootName = 'BPM Documents';
  let rootFolder = null;
  const rootIter = DriveApp.getFoldersByName(rootName);
  if (rootIter.hasNext()) {
    rootFolder = rootIter.next();
  } else {
    rootFolder = DriveApp.createFolder(rootName);
  }
  let targetFolder = null;
  const subIter = rootFolder.getFoldersByName(folderName);
  if (subIter.hasNext()) {
    targetFolder = subIter.next();
  } else {
    targetFolder = rootFolder.createFolder(folderName);
  }
  saveSettingValue(settingKey, targetFolder.getId(), 'Drive folder ID for ' + folderName);
  return targetFolder;
}

function saveBase64Upload(folderSettingKey, folderName, fileName, dataUrlOrBase64, defaultMime) {
  if (!dataUrlOrBase64) return { fileId: '', url: '' };
  let base64 = String(dataUrlOrBase64);
  let mimeType = defaultMime || 'image/jpeg';
  const match = base64.match(/^data:([^;]+);base64,(.*)$/);
  if (match) {
    mimeType = match[1];
    base64 = match[2];
  }
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mimeType, fileName || ('upload_' + Date.now() + '.jpg'));
  const folder = getOrCreateFolderByKey(folderSettingKey, folderName);
  const file = folder.createFile(blob);
  return { fileId: file.getId(), url: file.getUrl() };
}

function renderAndSavePdf(folderSettingKey, folderName, pdfFilename, html) {
  const htmlBlob = Utilities.newBlob(html, 'text/html', pdfFilename.replace(/\.pdf$/i, '.html'));
  const pdfBlob = htmlBlob.getAs('application/pdf').setName(pdfFilename);
  let pdfUrl = '';
  try {
    const folder = getOrCreateFolderByKey(folderSettingKey, folderName);
    const file = folder.createFile(pdfBlob);
    pdfUrl = file.getUrl();
  } catch (err) {
    pdfUrl = '';
  }
  const pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
  return { pdfUrl: pdfUrl, pdfBase64: pdfBase64, blob: pdfBlob, html: html };
}

// ============================================================================
// AUTHENTICATION & SERVER-SIDE RBAC
// ============================================================================

function findEmployeeByEmail(email) {
  const clean = String(email || '').trim().toLowerCase();
  if (!clean) return null;
  const tbl = getTableData('dim_Employees');
  return tbl.rows.find(function (r) {
    return String(r.Email || '').trim().toLowerCase() === clean && toBool(r.Active, true);
  }) || null;
}

function normalizePhoneDigits(val) {
  const d = String(val || '').replace(/\D+/g, '');
  if (d.length === 11 && d.charAt(0) === '1') return d.slice(1);
  return d;
}

function findEmployeeByIdentifier(identifier) {
  const byEmail = findEmployeeByEmail(identifier);
  if (byEmail) return byEmail;
  const digits = normalizePhoneDigits(identifier);
  if (digits.length < 7) return null;
  const tbl = getTableData('dim_Employees');
  return tbl.rows.find(function (r) {
    if (!toBool(r.Active, true)) return false;
    const empDigits = normalizePhoneDigits(r.Phone);
    if (!empDigits || empDigits.length < 7) return false;
    if (empDigits === digits) return true;
    if (digits.length === 7 && empDigits.slice(-7) === digits) return true;
    if (empDigits.length === 7 && digits.slice(-7) === empDigits) return true;
    return false;
  }) || null;
}

function findEmployeeById(empId) {
  const tbl = getTableData('dim_Employees');
  return tbl.rows.find(function (r) {
    return String(r.Employee_ID) === String(empId) && toBool(r.Active, true);
  }) || null;
}

function buildPublicUser(emp) {
  return {
    employeeId: String(emp.Employee_ID),
    name: String(emp.Name || ''),
    email: String(emp.Email || ''),
    phone: String(emp.Phone || ''),
    role: String(emp.Role || 'Field Crew'),
    payType: String(emp.Pay_Type || 'Hourly'),
    hasPin: Boolean(String(emp.Pin_Hash || '').trim())
  };
}

function isAdminRole(role) {
  return role === 'Owner' || role === 'Admin';
}

function issueSessionForEmployee(emp) {
  const token = 'bpm_' + Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  const tokenHash = sha256Hex(token);
  const now = new Date();
  const expires = new Date(now.getTime() + SESSION_TTL_DAYS * 86400000);
  appendRecords('sys_Sessions', [{
    Token_Hash: tokenHash,
    Employee_ID: emp.Employee_ID,
    Email: emp.Email || emp.Phone || '',
    Role: emp.Role,
    Expires_At: expires.toISOString(),
    Created_At: now.toISOString(),
    Last_Seen: now.toISOString()
  }]);
  return token;
}

function resolveSessionUser(token) {
  if (token) {
    const tokenHash = sha256Hex(String(token));
    const sessTbl = getTableData('sys_Sessions');
    const nowMs = Date.now();
    const match = sessTbl.rows.find(function (s) {
      return String(s.Token_Hash) === tokenHash && new Date(s.Expires_At).getTime() > nowMs;
    });
    if (match) {
      const emp = findEmployeeById(match.Employee_ID);
      if (emp) {
        return buildPublicUser(emp);
      }
    }
  }
  try {
    const activeEmail = Session.getActiveUser().getEmail();
    if (activeEmail) {
      const emp = findEmployeeByEmail(activeEmail);
      if (emp) return buildPublicUser(emp);
    }
  } catch (err) {
    // Ignore when Session unavailable
  }
  return null;
}

function requireAuth(token, allowedRoles) {
  const user = resolveSessionUser(token);
  if (!user) {
    throw new Error('AUTH_REQUIRED: Session expired or not signed in.');
  }
  if (allowedRoles && allowedRoles.length > 0 && allowedRoles.indexOf(user.role) === -1) {
    throw new Error('FORBIDDEN: Your role (' + user.role + ') cannot perform this action.');
  }
  return user;
}

/**
 * Initial app bootstrap. Validates existing session token or active Google user.
 */
function bootstrapApp(clientToken) {
  let user = null;
  let activeToken = clientToken || '';
  if (clientToken) {
    const tokenHash = sha256Hex(clientToken);
    const sessTbl = getTableData('sys_Sessions');
    const match = sessTbl.rows.find(function (s) {
      return String(s.Token_Hash) === tokenHash && new Date(s.Expires_At).getTime() > Date.now();
    });
    if (match) {
      const emp = findEmployeeById(match.Employee_ID);
      if (emp) user = buildPublicUser(emp);
    }
  }
  if (!user) {
    try {
      const activeEmail = Session.getActiveUser().getEmail();
      if (activeEmail) {
        const emp = findEmployeeByEmail(activeEmail);
        if (emp) {
          activeToken = withScriptLock(function () {
            return issueSessionForEmployee(emp);
          });
          user = buildPublicUser(emp);
        }
      }
    } catch (err) {
      // External anonymous visit
    }
  }
  const settings = getSettingsMap();
  if (!user) {
    return sanitizeForWire({
      authenticated: false,
      companyName: settings.COMPANY_NAME,
      companyPhone: settings.COMPANY_PHONE
    });
  }
  return sanitizeForWire({
    authenticated: true,
    token: activeToken,
    user: user,
    bootstrapData: buildAuthenticatedBootstrap(user)
  });
}

/**
 * Generates and emails a 6-digit login code for registered active employees.
 */
function requestLoginCode(email) {
  const clean = String(email || '').trim().toLowerCase();
  if (!clean) throw new Error('Please enter your email address.');
  const emp = findEmployeeByEmail(clean);
  if (emp) {
    const code = String(Math.floor(100000 + Math.random() * 900000));
    const cache = CacheService.getScriptCache();
    cache.put('OTP_' + clean, sha256Hex(clean + ':' + code), OTP_TTL_SECONDS);
    const settings = getSettingsMap();
    const html =
      '<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;padding:24px;border:1px solid #e2e8f0;border-radius:10px">' +
      '<h2 style="margin:0 0 8px;color:#0f172a">' + escapeHtml(settings.COMPANY_NAME) + '</h2>' +
      '<p style="color:#475569;font-size:14px">Hello ' + escapeHtml(emp.Name) + ',</p>' +
      '<p style="color:#475569;font-size:14px">Use the 6-digit verification code below to sign in (valid for 15 minutes):</p>' +
      '<div style="font-size:32px;font-weight:bold;letter-spacing:6px;color:#1e40af;background:#eff6ff;padding:16px;text-align:center;border-radius:8px;margin:16px 0">' +
      code + '</div>' +
      '<p style="color:#94a3b8;font-size:12px">If you did not request this code, you can ignore this email.</p></div>';
    MailApp.sendEmail({
      to: emp.Email,
      subject: settings.COMPANY_NAME + ' Login Code: ' + code,
      body: 'Your login code is: ' + code + ' (expires in 15 minutes).',
      htmlBody: html
    });
    addAuditLog(emp.Email, 'REQUEST_OTP', 'Employee', emp.Employee_ID, 'Login code sent');
  }
  return { ok: true, message: 'If your email is registered, a 6-digit login code has been sent.' };
}

/**
 * Verifies 6-digit OTP code and creates a 30-day session token.
 */
function verifyLoginCode(email, code) {
  const clean = String(email || '').trim().toLowerCase();
  const cleanCode = String(code || '').trim();
  if (!clean || !cleanCode) throw new Error('Email and 6-digit code are required.');
  const cache = CacheService.getScriptCache();
  const expectedHash = cache.get('OTP_' + clean);
  if (!expectedHash || expectedHash !== sha256Hex(clean + ':' + cleanCode)) {
    throw new Error('Invalid or expired login code.');
  }
  const emp = findEmployeeByEmail(clean);
  if (!emp) throw new Error('Account not found or inactive.');
  cache.remove('OTP_' + clean);
  return withScriptLock(function () {
    const token = issueSessionForEmployee(emp);
    const user = buildPublicUser(emp);
    addAuditLog(emp.Email, 'LOGIN_OTP', 'Session', emp.Employee_ID, 'Signed in with OTP');
    return sanitizeForWire({
      ok: true,
      token: token,
      user: user,
      needsPinSetup: !user.hasPin,
      bootstrapData: buildAuthenticatedBootstrap(user)
    });
  });
}

/**
 * Verifies Email or Phone Number + numeric PIN and creates a 30-day session token.
 */
function loginWithPin(identifier, pin) {
  const clean = String(identifier || '').trim();
  const cleanPin = String(pin || '').trim();
  if (!clean || !cleanPin) throw new Error('Email or phone number and PIN are required.');
  const emp = findEmployeeByIdentifier(clean);
  if (!emp) throw new Error('Invalid email/phone or PIN.');
  const storedHash = String(emp.Pin_Hash || '').trim();
  if (!storedHash) {
    throw new Error('No PIN set yet for this account. Ask an Admin to set a PIN or use Email 6-Digit Code.');
  }
  const attemptHash = sha256Hex(emp.Employee_ID + ':' + cleanPin);
  if (attemptHash !== storedHash) {
    throw new Error('Invalid email/phone or PIN.');
  }
  return withScriptLock(function () {
    const token = issueSessionForEmployee(emp);
    const user = buildPublicUser(emp);
    addAuditLog(emp.Email || emp.Phone || emp.Name, 'LOGIN_PIN', 'Session', emp.Employee_ID, 'Signed in with PIN');
    return sanitizeForWire({
      ok: true,
      token: token,
      user: user,
      bootstrapData: buildAuthenticatedBootstrap(user)
    });
  });
}

function setMyPin(token, newPin) {
  const user = requireAuth(token);
  const cleanPin = String(newPin || '').trim();
  if (!/^\d{4,8}$/.test(cleanPin)) {
    throw new Error('PIN must be 4 to 8 digits.');
  }
  return withScriptLock(function () {
    updateRowById('dim_Employees', 'Employee_ID', user.employeeId, {
      Pin_Hash: sha256Hex(user.employeeId + ':' + cleanPin)
    });
    addAuditLog(user.email, 'SET_PIN', 'Employee', user.employeeId, 'Updated own PIN');
    return { ok: true };
  });
}

function adminSetEmployeePin(token, employeeId, newPin) {
  const actor = requireAuth(token, ['Owner', 'Admin']);
  const cleanPin = String(newPin || '').trim();
  if (!/^\d{4,8}$/.test(cleanPin)) {
    throw new Error('PIN must be 4 to 8 digits.');
  }
  return withScriptLock(function () {
    const emp = updateRowById('dim_Employees', 'Employee_ID', employeeId, {
      Pin_Hash: sha256Hex(String(employeeId) + ':' + cleanPin)
    });
    if (!emp) throw new Error('Employee not found.');
    addAuditLog(actor.email, 'ADMIN_SET_PIN', 'Employee', employeeId, 'PIN reset by admin');
    return { ok: true };
  });
}

function logoutSession(token) {
  if (!token) return { ok: true };
  const tokenHash = sha256Hex(token);
  return withScriptLock(function () {
    deleteRowsWhere('sys_Sessions', function (r) {
      return String(r.Token_Hash) === tokenHash;
    });
    return { ok: true };
  });
}

function buildAuthenticatedBootstrap(user) {
  const settings = getSettingsMap();
  const propsTbl = getTableData('Properties');
  const tasksTbl = getTableData('Tasks');
  const woTbl = getTableData('fact_Work_Orders');

  const activeProperties = propsTbl.rows
    .filter(function (p) { return toBool(p.Status, true); })
    .map(function (p) {
      return {
        Property_ID: p.Property_ID,
        Property_Name: p.Property_Name,
        Owner_ID: p.Owner_ID,
        Owner_Name: p.Owner_Name,
        Address: p.Address,
        Unit_Type: p.Unit_Type,
        Default_Lease_Cycle: p.Default_Lease_Cycle,
        Aliases: p.Aliases,
        Gate_Lock_Notes: p.Gate_Lock_Notes
      };
    });

  const activeTasks = tasksTbl.rows
    .filter(function (t) { return toBool(t.Active, true); })
    .map(function (t) {
      return {
        Task_ID: t.Task_ID,
        Task_Name: t.Task_Name,
        Default_Billable_Rate: toNum(t.Default_Billable_Rate, 0),
        Default_Billing_Unit: t.Default_Billing_Unit || 'Hourly',
        Category: t.Category || ''
      };
    });

  if (isAdminRole(user.role)) {
    const ownersTbl = getTableData('dim_Owners');
    const empsTbl = getTableData('dim_Employees');
    const ratesTbl = getTableData('dim_Billing_Rates');
    return {
      settings: settings,
      properties: activeProperties,
      allProperties: propsTbl.rows,
      tasks: activeTasks,
      allTasks: tasksTbl.rows,
      owners: ownersTbl.rows,
      employees: empsTbl.rows.map(function (e) {
        const copy = Object.assign({}, e);
        copy.Has_Pin = Boolean(String(e.Pin_Hash || '').trim());
        delete copy.Pin_Hash;
        return copy;
      }),
      billingRates: ratesTbl.rows,
      openWorkOrdersCount: woTbl.rows.filter(function (w) {
        return w.Status !== 'Completed' && w.Status !== 'Billed' && w.Status !== 'Cancelled';
      }).length
    };
  }

  // Scoped payload for Field Crew / Contractor (zero owner financials or employee wages of others)
  return {
    settings: {
      COMPANY_NAME: settings.COMPANY_NAME,
      COMPANY_PHONE: settings.COMPANY_PHONE,
      DISPATCH_EMAIL: settings.DISPATCH_EMAIL
    },
    properties: activeProperties,
    tasks: activeTasks.map(function (t) {
      return { Task_ID: t.Task_ID, Task_Name: t.Task_Name, Category: t.Category };
    }),
    openWorkOrdersCount: woTbl.rows.filter(function (w) {
      return (
        String(w.Assigned_Employee_ID) === user.employeeId &&
        w.Status !== 'Completed' &&
        w.Status !== 'Billed' &&
        w.Status !== 'Cancelled'
      );
    }).length
  };
}

// ============================================================================
// PUBLIC TENANT MAINTENANCE INTAKE (?view=request)
// ============================================================================

/**
 * Returns ONLY public active property labels & intake categories (no owner/tenant/financial data).
 */
function getTenantPortalBootstrap() {
  const settings = getSettingsMap();
  const propsTbl = getTableData('Properties');
  const properties = propsTbl.rows
    .filter(function (p) { return toBool(p.Status, true); })
    .map(function (p) {
      return {
        propertyId: String(p.Property_ID),
        propertyName: String(p.Property_Name),
        address: String(p.Address || '')
      };
    })
    .sort(function (a, b) {
      return a.propertyName.localeCompare(b.propertyName);
    });
  return sanitizeForWire({
    companyName: settings.COMPANY_NAME,
    companyPhone: settings.COMPANY_PHONE,
    properties: properties,
    categories: TENANT_CATEGORIES
  });
}

function findPropertyRecord(propertyId, propertyName) {
  const propsTbl = getTableData('Properties');
  const cleanId = String(propertyId || '').trim();
  const cleanName = String(propertyName || '').trim().toLowerCase();
  return propsTbl.rows.find(function (p) {
    if (cleanId && String(p.Property_ID) === cleanId) return true;
    if (!cleanName) return false;
    if (String(p.Property_Name || '').trim().toLowerCase() === cleanName) return true;
    const aliases = String(p.Aliases || '').split(',').map(function (a) { return a.trim().toLowerCase(); });
    return aliases.indexOf(cleanName) !== -1;
  }) || null;
}

/**
 * Public tenant submission endpoint. Creates Work Order and sends email alert (prefixed [EMERGENCY] when urgent).
 */
function submitTenantRequest(payload) {
  if (!payload) throw new Error('Request payload is required.');
  const prop = findPropertyRecord(payload.propertyId, payload.propertyName);
  if (!prop) throw new Error('Please select a valid property.');
  const requesterName = String(payload.requesterName || '').trim();
  const requesterPhone = String(payload.requesterPhone || '').trim();
  const requesterEmail = String(payload.requesterEmail || '').trim();
  const description = String(payload.description || '').trim();
  if (!requesterName || (!requesterPhone && !requesterEmail) || !description) {
    throw new Error('Name, phone or email, and issue description are required.');
  }
  const urgency = ['Normal', 'Urgent', 'Emergency'].indexOf(payload.urgency) !== -1
    ? payload.urgency
    : 'Normal';

  return withScriptLock(function () {
    const woId = nextSeqId('WO', 'NEXT_WO_SEQ', 1001);
    const photoUrls = [];
    const photos = Array.isArray(payload.photos) ? payload.photos.slice(0, 3) : [];
    for (let i = 0; i < photos.length; i++) {
      const ph = photos[i];
      if (!ph || !ph.dataUrl) continue;
      const saved = saveBase64Upload(
        'RECEIPTS_FOLDER_ID',
        'Receipts & Intake Photos',
        woId + '_photo_' + (i + 1) + '.jpg',
        ph.dataUrl,
        ph.mimeType || 'image/jpeg'
      );
      if (saved.url) photoUrls.push(saved.url);
    }

    const title = (payload.taskCategory || 'Maintenance') + ' — ' + description.slice(0, 60);
    const woRecord = {
      WO_ID: woId,
      Created_At: nowIso(),
      Source: 'Tenant Portal',
      Property_ID: prop.Property_ID,
      Property_Name: prop.Property_Name,
      Unit_Room: String(payload.unitRoom || '').trim(),
      Requester_Name: requesterName,
      Requester_Phone: requesterPhone,
      Requester_Email: requesterEmail,
      Permission_To_Enter: String(payload.permissionToEnter || 'Yes - Anytime'),
      Urgency: urgency,
      Task_Category: String(payload.taskCategory || 'General Maintenance'),
      Title: title,
      Description: description,
      Photo_URLs: photoUrls.join('\n'),
      Assigned_Employee_ID: '',
      Assigned_To_Name: '',
      Due_Date: urgency === 'Emergency' ? todayYmd() : addDaysYmd(todayYmd(), urgency === 'Urgent' ? 1 : 3),
      Status: 'New',
      Completed_At: '',
      Resolution_Notes: '',
      NTE_Limit: ''
    };
    appendRecords('fact_Work_Orders', [woRecord]);
    addAuditLog(requesterEmail || requesterName, 'TENANT_REQUEST_CREATED', 'WorkOrder', woId, prop.Property_Name + ' [' + urgency + ']');

    // Immediate Email Dispatch Alert
    const settings = getSettingsMap();
    const dispatchTo = settings.DISPATCH_EMAIL || 'butlerpropmgmt@gmail.com';
    const urgentPrefix = (urgency === 'Emergency' || urgency === 'Urgent') ? '[' + urgency.toUpperCase() + '] ' : '';
    const photoLinksHtml = photoUrls.length > 0
      ? '<p><strong>Attached Photos:</strong><br>' + photoUrls.map(function (u, idx) {
          return '<a href="' + escapeHtml(u) + '">Photo ' + (idx + 1) + '</a>';
        }).join(' &bull; ') + '</p>'
      : '';
    try {
      MailApp.sendEmail({
        to: dispatchTo,
        replyTo: requesterEmail || dispatchTo,
        subject: urgentPrefix + 'Maintenance Request ' + woId + ': ' + prop.Property_Name,
        body:
          'Ticket: ' + woId + '\nProperty: ' + prop.Property_Name + '\nUrgency: ' + urgency +
          '\nTenant: ' + requesterName + ' (' + requesterPhone + ' / ' + requesterEmail + ')' +
          '\nPermission to Enter: ' + woRecord.Permission_To_Enter +
          '\nCategory: ' + woRecord.Task_Category + '\nDescription:\n' + description,
        htmlBody:
          '<div style="font-family:Arial,sans-serif;max-width:600px;border:1px solid #e2e8f0;border-radius:10px;padding:20px">' +
          '<h2 style="margin-top:0;color:' + (urgency === 'Emergency' ? '#b91c1c' : '#0f172a') + '">' +
          escapeHtml(urgentPrefix + 'New Tenant Request — ' + woId) + '</h2>' +
          '<p><strong>Property:</strong> ' + escapeHtml(prop.Property_Name) + (woRecord.Unit_Room ? ' (' + escapeHtml(woRecord.Unit_Room) + ')' : '') + '</p>' +
          '<p><strong>Urgency:</strong> ' + escapeHtml(urgency) + ' &bull; <strong>Category:</strong> ' + escapeHtml(woRecord.Task_Category) + '</p>' +
          '<p><strong>Tenant:</strong> ' + escapeHtml(requesterName) + ' &bull; ' + escapeHtml(requesterPhone) + ' &bull; ' + escapeHtml(requesterEmail) + '</p>' +
          '<p><strong>Permission to Enter:</strong> ' + escapeHtml(woRecord.Permission_To_Enter) + '</p>' +
          '<div style="background:#f8fafc;padding:12px;border-radius:8px;margin:12px 0"><strong>Issue Details:</strong><br>' +
          escapeHtml(description).replace(/\n/g, '<br>') + '</div>' +
          photoLinksHtml +
          '<p style="color:#64748b;font-size:12px">Open the BPM Operations Portal to assign this work order so nothing falls through the cracks.</p></div>'
      });
    } catch (mailErr) {
      addAuditLog('system', 'EMAIL_ALERT_FAILED', 'WorkOrder', woId, String(mailErr));
    }
    return sanitizeForWire({
      ok: true,
      woId: woId,
      propertyName: prop.Property_Name,
      urgency: urgency,
      createdAt: woRecord.Created_At
    });
  });
}

// ============================================================================
// 5-TIER BILLABLE RATE RESOLUTION ENGINE
// ============================================================================

/**
 * Resolves hourly/flat client billable rate strictly independent from employee pay cost.
 * Order: Custom Override > Property+Task > Owner+Task > Employee+Task > Global Task > Global Fallback.
 */
function resolveBillableRate(params, ctx) {
  const settings = ctx.settings || getSettingsMap();
  const rates = ctx.rates || getTableData('dim_Billing_Rates').rows;
  const tasks = ctx.tasks || getTableData('Tasks').rows;
  const owners = ctx.owners || getTableData('dim_Owners').rows;
  const employees = ctx.employees || getTableData('dim_Employees').rows;

  const hours = toNum(params.hours, 0);
  const workType = String(params.workType || 'Time');
  const chargeTarget = String(params.chargeTarget || 'Owner');
  const taskName = String(params.taskName || '').trim();
  const taskLower = taskName.toLowerCase();

  if (chargeTarget === 'BPM Internal') {
    return { billableRate: 0, calculatedBillable: 0, effectiveBillable: 0, rateSource: 'BPM Internal ($0)', rateType: 'Hourly' };
  }
  if (workType === 'Expense Only' || hours === 0) {
    return { billableRate: 0, calculatedBillable: 0, effectiveBillable: 0, rateSource: 'Expense Only', rateType: 'Flat' };
  }

  const customRaw = params.customCharge;
  const hasCustom = customRaw !== '' && customRaw !== null && customRaw !== undefined && !isNaN(Number(customRaw));

  const taskRow = tasks.find(function (t) {
    return String(t.Task_Name || '').trim().toLowerCase() === taskLower;
  });
  const ownerRow = owners.find(function (o) {
    return String(o.Owner_ID) === String(params.ownerId || '');
  });
  const empRow = employees.find(function (e) {
    return String(e.Employee_ID) === String(params.employeeId || '');
  });

  function matchScopeRate(scopeType, matchId, matchName) {
    let exactTask = null;
    let allTask = null;
    for (let i = 0; i < rates.length; i++) {
      const r = rates[i];
      if (String(r.Scope_Type) !== scopeType) continue;
      const scopeMatch =
        scopeType === 'Global' ||
        (matchId && String(r.Scope_ID) === String(matchId)) ||
        (matchName && String(r.Scope_Name || '').trim().toLowerCase() === String(matchName).trim().toLowerCase());
      if (!scopeMatch) continue;
      const rTask = String(r.Task_Name || '').trim().toLowerCase();
      if (rTask === taskLower) { exactTask = r; break; }
      if (rTask === 'all' || rTask === 'default') allTask = r;
    }
    return exactTask || allTask;
  }

  let resolvedRate = 0;
  let rateType = (taskRow && taskRow.Default_Billing_Unit === 'Flat') ? 'Flat' : 'Hourly';
  let rateSource = '';

  const propRule = matchScopeRate('Property', params.propertyId, params.propertyName);
  if (propRule && toNum(propRule.Billable_Rate, 0) > 0) {
    resolvedRate = toNum(propRule.Billable_Rate, 0);
    rateType = propRule.Rate_Type || rateType;
    rateSource = 'Property Rate';
  } else {
    const ownerRule = params.ownerId ? matchScopeRate('Owner', params.ownerId, ownerRow ? ownerRow.Owner_Name : '') : null;
    if (ownerRule && toNum(ownerRule.Billable_Rate, 0) > 0) {
      resolvedRate = toNum(ownerRule.Billable_Rate, 0);
      rateType = ownerRule.Rate_Type || rateType;
      rateSource = 'Owner Task Rate';
    } else if (ownerRow && toNum(ownerRow.Default_Rate_Override, 0) > 0) {
      resolvedRate = toNum(ownerRow.Default_Rate_Override, 0);
      rateSource = 'Owner Default Override';
    } else {
      const empRule = params.employeeId ? matchScopeRate('Employee', params.employeeId, empRow ? empRow.Name : '') : null;
      if (empRule && toNum(empRule.Billable_Rate, 0) > 0) {
        resolvedRate = toNum(empRule.Billable_Rate, 0);
        rateType = empRule.Rate_Type || rateType;
        rateSource = 'Employee Rate Rule';
      } else if (empRow && toNum(empRow.Default_Billable_Rate_Override, 0) > 0) {
        resolvedRate = toNum(empRow.Default_Billable_Rate_Override, 0);
        rateSource = 'Employee Billable Override';
      } else {
        const globalRule = matchScopeRate('Global', '', '');
        if (globalRule && toNum(globalRule.Billable_Rate, 0) > 0) {
          resolvedRate = toNum(globalRule.Billable_Rate, 0);
          rateType = globalRule.Rate_Type || rateType;
          rateSource = 'Global Rate Matrix';
        } else if (taskRow && toNum(taskRow.Default_Billable_Rate, 0) > 0) {
          resolvedRate = toNum(taskRow.Default_Billable_Rate, 0);
          rateType = taskRow.Default_Billing_Unit === 'Flat' ? 'Flat' : 'Hourly';
          rateSource = 'Task Default Rate';
        } else {
          resolvedRate = toNum(settings.DEFAULT_FALLBACK_RATE, 65);
          rateSource = 'Fallback Rate';
        }
      }
    }
  }

  if (workType === 'Flat') rateType = 'Flat';
  const calculatedBillable = rateType === 'Flat' ? round2(resolvedRate) : round2(hours * resolvedRate);
  const effectiveBillable = hasCustom ? round2(Number(customRaw)) : calculatedBillable;
  if (hasCustom) {
    rateSource = 'Custom Override';
  }
  return {
    billableRate: round2(resolvedRate),
    calculatedBillable: calculatedBillable,
    effectiveBillable: effectiveBillable,
    rateSource: rateSource,
    rateType: rateType
  };
}

function getEffectiveLogBillable(log) {
  if (String(log.Charge_Target) === 'BPM Internal') return 0;
  const custom = log.Charge_Amount_Custom;
  if (custom !== '' && custom !== null && custom !== undefined && !isNaN(Number(custom))) {
    return round2(Number(custom));
  }
  return round2(toNum(log.Calculated_Billable, 0));
}

// ============================================================================
// WORK ORDERS, WORK LOGS & 0-HR MATERIAL EXPENSES
// ============================================================================

function getWorkOrders(token, filters) {
  const user = requireAuth(token);
  const tbl = getTableData('fact_Work_Orders');
  const f = filters || {};
  let rows = tbl.rows.slice();
  if (!isAdminRole(user.role)) {
    rows = rows.filter(function (w) {
      return String(w.Assigned_Employee_ID) === user.employeeId || String(w.Status) === 'New';
    });
  }
  if (f.status && f.status !== 'All') {
    if (f.status === 'Open') {
      rows = rows.filter(function (w) {
        return ['Completed', 'Billed', 'Cancelled'].indexOf(w.Status) === -1;
      });
    } else {
      rows = rows.filter(function (w) { return String(w.Status) === f.status; });
    }
  }
  if (f.urgency && f.urgency !== 'All') {
    rows = rows.filter(function (w) { return String(w.Urgency) === f.urgency; });
  }
  if (f.propertyId) {
    rows = rows.filter(function (w) { return String(w.Property_ID) === String(f.propertyId); });
  }
  rows.sort(function (a, b) {
    const urgRank = { Emergency: 0, Urgent: 1, Normal: 2, Low: 3 };
    const statusDone = function (s) { return s === 'Completed' || s === 'Billed' || s === 'Cancelled' ? 1 : 0; };
    if (statusDone(a.Status) !== statusDone(b.Status)) return statusDone(a.Status) - statusDone(b.Status);
    const ua = urgRank[a.Urgency] !== undefined ? urgRank[a.Urgency] : 9;
    const ub = urgRank[b.Urgency] !== undefined ? urgRank[b.Urgency] : 9;
    if (ua !== ub) return ua - ub;
    return String(b.Created_At || '').localeCompare(String(a.Created_At || ''));
  });
  return sanitizeForWire(rows);
}

function saveWorkOrder(token, payload) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  if (!payload) throw new Error('Work order payload required.');
  return withScriptLock(function () {
    const prop = findPropertyRecord(payload.propertyId, payload.propertyName);
    if (!prop) throw new Error('Valid Property is required.');

    let assignedEmp = null;
    if (payload.assignedEmployeeId) {
      assignedEmp = findEmployeeById(payload.assignedEmployeeId);
    }

    const patch = {
      Property_ID: prop.Property_ID,
      Property_Name: prop.Property_Name,
      Unit_Room: String(payload.unitRoom || ''),
      Requester_Name: String(payload.requesterName || ''),
      Requester_Phone: String(payload.requesterPhone || ''),
      Requester_Email: String(payload.requesterEmail || ''),
      Permission_To_Enter: String(payload.permissionToEnter || 'Yes - Coordinate'),
      Urgency: String(payload.urgency || 'Normal'),
      Task_Category: String(payload.taskCategory || 'General Maintenance'),
      Title: String(payload.title || payload.taskCategory || 'Work Order'),
      Description: String(payload.description || ''),
      Assigned_Employee_ID: assignedEmp ? assignedEmp.Employee_ID : '',
      Assigned_To_Name: assignedEmp ? assignedEmp.Name : '',
      Due_Date: toYmd(payload.dueDate || ''),
      Status: String(payload.status || (assignedEmp ? 'Assigned' : 'New')),
      Resolution_Notes: String(payload.resolutionNotes || ''),
      NTE_Limit: payload.nteLimit !== undefined && payload.nteLimit !== '' ? toNum(payload.nteLimit, '') : ''
    };
    if (patch.Status === 'Completed' && !payload.completedAt) {
      patch.Completed_At = nowIso();
    }

    let saved = null;
    let isNew = false;
    if (payload.woId) {
      saved = updateRowById('fact_Work_Orders', 'WO_ID', payload.woId, patch);
      if (!saved) throw new Error('Work Order not found: ' + payload.woId);
    } else {
      isNew = true;
      const woId = nextSeqId('WO', 'NEXT_WO_SEQ', 1001);
      saved = Object.assign({
        WO_ID: woId,
        Created_At: nowIso(),
        Source: String(payload.source || 'Owner/Admin'),
        Photo_URLs: '',
        Completed_At: ''
      }, patch);
      appendRecords('fact_Work_Orders', [saved]);
    }

    if (payload.notifyAssignee && assignedEmp && assignedEmp.Email) {
      try {
        const settings = getSettingsMap();
        MailApp.sendEmail({
          to: assignedEmp.Email,
          subject: '[' + saved.Urgency + '] Assigned Work Order ' + saved.WO_ID + ': ' + saved.Property_Name,
          body:
            'Hi ' + assignedEmp.Name + ',\n\nYou have been assigned Work Order ' + saved.WO_ID +
            ' at ' + saved.Property_Name + '.\nCategory: ' + saved.Task_Category +
            '\nDue Date: ' + (saved.Due_Date || 'ASAP') +
            '\nDescription:\n' + saved.Description +
            '\n\n' + (prop.Gate_Lock_Notes ? 'Access Notes: ' + prop.Gate_Lock_Notes : ''),
          htmlBody:
            '<div style="font-family:Arial,sans-serif;max-width:560px;padding:18px;border:1px solid #e2e8f0;border-radius:10px">' +
            '<h3 style="margin-top:0;color:#0f172a">Work Order Assigned: ' + escapeHtml(saved.WO_ID) + '</h3>' +
            '<p><strong>Property:</strong> ' + escapeHtml(saved.Property_Name) + '</p>' +
            '<p><strong>Urgency:</strong> ' + escapeHtml(saved.Urgency) + ' &bull; <strong>Due:</strong> ' + escapeHtml(saved.Due_Date || 'ASAP') + '</p>' +
            '<p><strong>Task:</strong> ' + escapeHtml(saved.Task_Category) + '</p>' +
            '<div style="background:#f8fafc;padding:12px;border-radius:8px">' + escapeHtml(saved.Description) + '</div>' +
            (prop.Gate_Lock_Notes ? '<p><strong>Access / Lock Notes:</strong> ' + escapeHtml(prop.Gate_Lock_Notes) + '</p>' : '') +
            '<p style="color:#64748b;font-size:12px">' + escapeHtml(settings.COMPANY_NAME) + '</p></div>'
        });
      } catch (err) {
        // Non-fatal
      }
    }

    addAuditLog(user.email, isNew ? 'CREATE_WORK_ORDER' : 'UPDATE_WORK_ORDER', 'WorkOrder', saved.WO_ID, saved.Property_Name);
    return sanitizeForWire(saved);
  });
}

function cancelWorkOrder(token, woId) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  return withScriptLock(function () {
    const updated = updateRowById('fact_Work_Orders', 'WO_ID', woId, { Status: 'Cancelled' });
    if (!updated) throw new Error('Work Order not found.');
    addAuditLog(user.email, 'CANCEL_WORK_ORDER', 'WorkOrder', woId, 'Cancelled');
    return sanitizeForWire(updated);
  });
}

/**
 * Logs Time, Flat-Rate work, and/or 0-Hour Standalone Expenses & Material Receipts.
 */
function submitWorkEntry(token, payload) {
  const user = requireAuth(token);
  if (!payload) throw new Error('Work entry payload required.');

  return withScriptLock(function () {
    const prop = findPropertyRecord(payload.propertyId, payload.propertyName);
    if (!prop) throw new Error('Please select a valid property.');

    const targetEmpId = isAdminRole(user.role) && payload.employeeId ? String(payload.employeeId) : user.employeeId;
    const emp = findEmployeeById(targetEmpId);
    if (!emp) throw new Error('Active employee record not found.');

    const settings = getSettingsMap();
    const ownersTbl = getTableData('dim_Owners');
    const owner = ownersTbl.rows.find(function (o) { return String(o.Owner_ID) === String(prop.Owner_ID); }) || {};

    const workType = ['Time', 'Flat', 'Expense Only'].indexOf(payload.workType) !== -1
      ? payload.workType
      : (toNum(payload.hours, 0) === 0 ? 'Expense Only' : 'Time');

    let hours = workType === 'Expense Only' ? 0 : round2(toNum(payload.hours, 0));
    if (hours === 0 && payload.startTime && payload.endTime && workType === 'Time') {
      hours = computeHoursFromTimeStrings(payload.startTime, payload.endTime);
    }

    const materials = Array.isArray(payload.materials) ? payload.materials : [];
    if (workType !== 'Expense Only' && workType !== 'Flat' && hours <= 0 && materials.length === 0) {
      throw new Error('Enter hours worked OR add at least one material/receipt expense.');
    }
    if (workType === 'Expense Only' && materials.length === 0) {
      throw new Error('Standalone expense mode requires at least one receipt or material line item.');
    }

    const workDate = toYmd(payload.workDate) || todayYmd();
    const chargeTarget = ['Owner', 'Tenant', 'BPM Internal'].indexOf(payload.chargeTarget) !== -1
      ? payload.chargeTarget
      : 'Owner';
    const woId = String(payload.woId || '').trim();
    let createdLog = null;

    // Only write a fact_Work_Logs row if there is actual labor time or a flat labor charge
    if (workType !== 'Expense Only' && (hours > 0 || workType === 'Flat')) {
      const logId = newId('LOG');
      const empRate = toNum(emp.Default_Pay_Rate, 0);
      const grossPay = workType === 'Flat'
        ? round2(payload.flatPayAmount !== undefined && payload.flatPayAmount !== ''
            ? toNum(payload.flatPayAmount, 0)
            : (emp.Pay_Type === 'Flat/1099' ? empRate : round2(hours * empRate)))
        : round2(hours * empRate);

      const rateInfo = resolveBillableRate({
        propertyId: prop.Property_ID,
        propertyName: prop.Property_Name,
        ownerId: prop.Owner_ID,
        employeeId: emp.Employee_ID,
        taskName: payload.taskName || 'General Maintenance',
        hours: hours || 1,
        workType: workType,
        customCharge: isAdminRole(user.role) ? payload.customCharge : ''
      }, {
        settings: settings,
        owners: ownersTbl.rows,
        employees: [emp]
      });

      const autoApprove = isAdminRole(user.role) && Boolean(payload.autoApprove);
      createdLog = {
        Log_ID: logId,
        WO_ID: woId,
        Work_Date: workDate,
        Timestamp: nowIso(),
        Employee_ID: emp.Employee_ID,
        Employee_Name: emp.Name,
        Property_ID: prop.Property_ID,
        Property_Name: prop.Property_Name,
        Task_Name: String(payload.taskName || 'General Maintenance'),
        Hours: hours,
        Start_Time: String(payload.startTime || ''),
        End_Time: String(payload.endTime || ''),
        Notes: String(payload.notes || ''),
        Work_Type: workType,
        Payroll_Status: autoApprove ? 'Approved' : 'Pending',
        Payroll_ID: '',
        Employee_Rate: empRate,
        Gross_Pay: grossPay,
        Billing_Status: chargeTarget === 'BPM Internal' ? 'Internal/Non-Billable' : 'Unbilled',
        Invoice_ID: '',
        Charge_Target: chargeTarget,
        Billable_Rate: rateInfo.billableRate,
        Calculated_Billable: rateInfo.calculatedBillable,
        Charge_Amount_Custom: isAdminRole(user.role) && payload.customCharge !== undefined && payload.customCharge !== ''
          ? round2(toNum(payload.customCharge, 0))
          : '',
        Rate_Source: rateInfo.rateSource,
        Edited_By: user.email,
        Edited_At: ''
      };
      appendRecords('fact_Work_Logs', [createdLog]);
    }

    const createdMaterials = [];
    for (let i = 0; i < materials.length; i++) {
      const m = materials[i];
      if (!m) continue;
      const qty = toNum(m.quantity, 1) || 1;
      const unitCost = round2(toNum(m.unitCost, m.totalCost || 0));
      const totalCost = m.unitCost !== undefined && m.unitCost !== ''
        ? round2(qty * unitCost)
        : round2(toNum(m.totalCost, 0));
      if (totalCost <= 0 && !m.itemDescription) continue;

      const defaultMarkup = owner.Material_Markup_Pct !== '' && owner.Material_Markup_Pct !== undefined
        ? toNum(owner.Material_Markup_Pct, 0)
        : toNum(settings.DEFAULT_MATERIAL_MARKUP_PCT, 0);
      const markupPct = (isAdminRole(user.role) && m.markupPct !== undefined && m.markupPct !== '')
        ? toNum(m.markupPct, defaultMarkup)
        : defaultMarkup;
      const mTarget = ['Owner', 'Tenant', 'BPM Internal'].indexOf(m.chargeTarget) !== -1
        ? m.chargeTarget
        : chargeTarget;
      const billableAmount = mTarget === 'BPM Internal' ? 0 : round2(totalCost * (1 + markupPct / 100));
      const paidBy = (m.paidBy === 'Employee Out-of-Pocket' || m.paidBy === 'Stock') ? m.paidBy : 'Company Account/Card';

      const expenseId = newId('EXP');
      let receiptUrl = String(m.receiptUrl || '');
      if (m.receiptDataUrl) {
        const uploaded = saveBase64Upload(
          'RECEIPTS_FOLDER_ID',
          'Receipts & Intake Photos',
          expenseId + '_' + (m.receiptFileName || 'receipt.jpg'),
          m.receiptDataUrl,
          m.receiptMimeType || 'image/jpeg'
        );
        receiptUrl = uploaded.url;
      }

      const expPayrollStatus = paidBy === 'Employee Out-of-Pocket'
        ? ((isAdminRole(user.role) && payload.autoApprove) ? 'Approved' : 'Pending')
        : 'N/A';

      createdMaterials.push({
        Expense_ID: expenseId,
        WO_ID: woId,
        Log_ID: createdLog ? createdLog.Log_ID : '',
        Expense_Date: toYmd(m.expenseDate) || workDate,
        Timestamp: nowIso(),
        Employee_ID: emp.Employee_ID,
        Employee_Name: emp.Name,
        Property_ID: prop.Property_ID,
        Property_Name: prop.Property_Name,
        Vendor: String(m.vendor || ''),
        Item_Description: String(m.itemDescription || 'Materials'),
        Quantity: qty,
        Unit_Cost: unitCost,
        Total_Cost: totalCost,
        Markup_Pct: markupPct,
        Billable_Amount: billableAmount,
        Paid_By: paidBy,
        Charge_Target: mTarget,
        Receipt_URL: receiptUrl,
        Payroll_Status: expPayrollStatus,
        Payroll_ID: '',
        Billing_Status: mTarget === 'BPM Internal' ? 'Internal/Non-Billable' : 'Unbilled',
        Invoice_ID: '',
        Notes: String(m.notes || payload.notes || '')
      });
    }
    if (createdMaterials.length > 0) {
      appendRecords('fact_Material_Expenses', createdMaterials);
    }

    if (woId) {
      const woTbl = getTableData('fact_Work_Orders');
      const woRow = woTbl.rows.find(function (w) { return String(w.WO_ID) === woId; });
      if (woRow && woRow.Status !== 'Billed' && woRow.Status !== 'Cancelled') {
        const nextStatus = payload.markWoCompleted ? 'Completed' : (woRow.Status === 'New' || woRow.Status === 'Assigned' ? 'In Progress' : woRow.Status);
        const woPatch = { Status: nextStatus };
        if (payload.markWoCompleted) {
          woPatch.Completed_At = nowIso();
          if (payload.notes) woPatch.Resolution_Notes = String(payload.notes);
        }
        updateRowById('fact_Work_Orders', 'WO_ID', woId, woPatch);
      }
    }

    addAuditLog(
      user.email,
      workType === 'Expense Only' ? 'SUBMIT_EXPENSE' : 'SUBMIT_WORK_LOG',
      createdLog ? 'WorkLog' : 'Expense',
      createdLog ? createdLog.Log_ID : (createdMaterials[0] ? createdMaterials[0].Expense_ID : ''),
      prop.Property_Name + ' (' + hours + ' hrs, ' + createdMaterials.length + ' materials)'
    );

    return sanitizeForWire({
      ok: true,
      workLog: createdLog,
      materials: createdMaterials
    });
  });
}

function computeHoursFromTimeStrings(startStr, endStr) {
  const s = String(startStr || '').split(':').map(Number);
  const e = String(endStr || '').split(':').map(Number);
  if (s.length < 2 || e.length < 2 || isNaN(s[0]) || isNaN(e[0])) return 0;
  let diffMin = (e[0] * 60 + e[1]) - (s[0] * 60 + s[1]);
  if (diffMin < 0) diffMin += 1440;
  return round2(diffMin / 60);
}

/**
 * Updates a Work Log entry. Field Crew may edit own Pending/Unbilled logs; Admin/Owner may edit rate, custom charge & status.
 */
function updateWorkLogEntry(token, payload) {
  const user = requireAuth(token);
  if (!payload || !payload.logId) throw new Error('logId is required.');

  return withScriptLock(function () {
    const tbl = getTableData('fact_Work_Logs');
    const existing = tbl.rows.find(function (r) { return String(r.Log_ID) === String(payload.logId); });
    if (!existing) throw new Error('Work log not found.');

    const admin = isAdminRole(user.role);
    if (!admin) {
      if (String(existing.Employee_ID) !== user.employeeId) {
        throw new Error('FORBIDDEN: You can only edit your own entries.');
      }
      if (existing.Payroll_Status !== 'Pending' || existing.Billing_Status !== 'Unbilled') {
        throw new Error('Approved, paid, or invoiced entries can only be modified by an Owner/Admin.');
      }
    }

    const patch = {
      Edited_By: user.email,
      Edited_At: nowIso()
    };
    if (payload.workDate) patch.Work_Date = toYmd(payload.workDate);
    if (payload.notes !== undefined) patch.Notes = String(payload.notes);
    if (payload.taskName) patch.Task_Name = String(payload.taskName);
    if (payload.woId !== undefined) patch.WO_ID = String(payload.woId);

    if (payload.propertyId || payload.propertyName) {
      const prop = findPropertyRecord(payload.propertyId, payload.propertyName);
      if (prop) {
        patch.Property_ID = prop.Property_ID;
        patch.Property_Name = prop.Property_Name;
      }
    }

    const hours = payload.hours !== undefined ? round2(toNum(payload.hours, existing.Hours)) : toNum(existing.Hours, 0);
    patch.Hours = hours;
    const empRate = (admin && payload.employeeRate !== undefined) ? toNum(payload.employeeRate, existing.Employee_Rate) : toNum(existing.Employee_Rate, 0);
    patch.Employee_Rate = empRate;
    if (existing.Work_Type === 'Flat' && admin && payload.grossPay !== undefined) {
      patch.Gross_Pay = round2(toNum(payload.grossPay, existing.Gross_Pay));
    } else if (existing.Work_Type !== 'Flat') {
      patch.Gross_Pay = round2(hours * empRate);
    }

    if (admin) {
      if (payload.chargeTarget) patch.Charge_Target = payload.chargeTarget;
      if (payload.payrollStatus && existing.Payroll_Status !== 'Paid') patch.Payroll_Status = payload.payrollStatus;

      const target = patch.Charge_Target || existing.Charge_Target;
      if (target === 'BPM Internal') {
        patch.Billing_Status = 'Internal/Non-Billable';
        patch.Calculated_Billable = 0;
        patch.Charge_Amount_Custom = '';
      } else {
        if (existing.Billing_Status === 'Internal/Non-Billable') {
          patch.Billing_Status = 'Unbilled';
        }
        if (payload.billableRate !== undefined && payload.billableRate !== '') {
          const newRate = round2(toNum(payload.billableRate, existing.Billable_Rate));
          patch.Billable_Rate = newRate;
          patch.Calculated_Billable = existing.Work_Type === 'Flat' ? newRate : round2(hours * newRate);
          patch.Rate_Source = 'Manual Rate Edit';
        } else if (payload.hours !== undefined || payload.taskName !== undefined || payload.propertyId !== undefined) {
          const propObj = findPropertyRecord(patch.Property_ID || existing.Property_ID, patch.Property_Name || existing.Property_Name) || {};
          const resolved = resolveBillableRate({
            propertyId: propObj.Property_ID || existing.Property_ID,
            propertyName: propObj.Property_Name || existing.Property_Name,
            ownerId: propObj.Owner_ID || '',
            employeeId: existing.Employee_ID,
            taskName: patch.Task_Name || existing.Task_Name,
            hours: hours,
            workType: existing.Work_Type
          }, {});
          patch.Billable_Rate = resolved.billableRate;
          patch.Calculated_Billable = resolved.calculatedBillable;
          patch.Rate_Source = resolved.rateSource;
        }
        if (payload.chargeAmountCustom !== undefined) {
          patch.Charge_Amount_Custom = payload.chargeAmountCustom === '' || payload.chargeAmountCustom === null
            ? ''
            : round2(toNum(payload.chargeAmountCustom, 0));
        }
      }
    }

    const updated = updateRowById('fact_Work_Logs', 'Log_ID', payload.logId, patch);
    addAuditLog(user.email, 'UPDATE_WORK_LOG', 'WorkLog', payload.logId, 'Updated work log');
    return sanitizeForWire(updated);
  });
}

function updateMaterialExpense(token, payload) {
  const user = requireAuth(token);
  if (!payload || !payload.expenseId) throw new Error('expenseId is required.');

  return withScriptLock(function () {
    const tbl = getTableData('fact_Material_Expenses');
    const existing = tbl.rows.find(function (r) { return String(r.Expense_ID) === String(payload.expenseId); });
    if (!existing) throw new Error('Material expense not found.');

    const admin = isAdminRole(user.role);
    if (!admin) {
      if (String(existing.Employee_ID) !== user.employeeId) throw new Error('FORBIDDEN');
      if (existing.Billing_Status !== 'Unbilled' || existing.Payroll_Status === 'Paid' || existing.Payroll_Status === 'Approved') {
        throw new Error('Locked expense can only be edited by Admin/Owner.');
      }
    }

    const qty = payload.quantity !== undefined ? toNum(payload.quantity, existing.Quantity) : toNum(existing.Quantity, 1);
    const unitCost = payload.unitCost !== undefined ? toNum(payload.unitCost, existing.Unit_Cost) : toNum(existing.Unit_Cost, 0);
    const totalCost = payload.totalCost !== undefined ? round2(toNum(payload.totalCost, existing.Total_Cost)) : round2(qty * unitCost);
    const markupPct = admin && payload.markupPct !== undefined ? toNum(payload.markupPct, existing.Markup_Pct) : toNum(existing.Markup_Pct, 0);
    const target = admin && payload.chargeTarget ? payload.chargeTarget : existing.Charge_Target;
    const paidBy = payload.paidBy || existing.Paid_By;

    const patch = {
      Expense_Date: payload.expenseDate ? toYmd(payload.expenseDate) : existing.Expense_Date,
      Vendor: payload.vendor !== undefined ? String(payload.vendor) : existing.Vendor,
      Item_Description: payload.itemDescription !== undefined ? String(payload.itemDescription) : existing.Item_Description,
      Quantity: qty,
      Unit_Cost: round2(unitCost),
      Total_Cost: totalCost,
      Markup_Pct: markupPct,
      Charge_Target: target,
      Paid_By: paidBy,
      Billable_Amount: target === 'BPM Internal' ? 0 : round2(totalCost * (1 + markupPct / 100)),
      Notes: payload.notes !== undefined ? String(payload.notes) : existing.Notes
    };

    if (admin) {
      patch.Billing_Status = target === 'BPM Internal'
        ? 'Internal/Non-Billable'
        : (existing.Billing_Status === 'Internal/Non-Billable' ? 'Unbilled' : existing.Billing_Status);
      if (paidBy !== 'Employee Out-of-Pocket') {
        patch.Payroll_Status = 'N/A';
      } else if (payload.payrollStatus && existing.Payroll_Status !== 'Paid') {
        patch.Payroll_Status = payload.payrollStatus;
      } else if (existing.Payroll_Status === 'N/A') {
        patch.Payroll_Status = 'Pending';
      }
    }

    const updated = updateRowById('fact_Material_Expenses', 'Expense_ID', payload.expenseId, patch);
    addAuditLog(user.email, 'UPDATE_EXPENSE', 'Expense', payload.expenseId, 'Updated expense');
    return sanitizeForWire(updated);
  });
}

function batchApproveItems(token, payload) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  const logIds = Array.isArray(payload.logIds) ? payload.logIds : [];
  const expenseIds = Array.isArray(payload.expenseIds) ? payload.expenseIds : [];
  const nextStatus = payload.status === 'Pending' ? 'Pending' : 'Approved';

  return withScriptLock(function () {
    const logPatches = {};
    const logTbl = getTableData('fact_Work_Logs');
    logTbl.rows.forEach(function (r) {
      if (logIds.indexOf(r.Log_ID) !== -1 && r.Payroll_Status !== 'Paid' && r.Payroll_Status !== 'N/A') {
        logPatches[r.Log_ID] = { Payroll_Status: nextStatus, Edited_By: user.email, Edited_At: nowIso() };
      }
    });
    const logsUpdated = batchPatchRowsById('fact_Work_Logs', 'Log_ID', logPatches);

    const expPatches = {};
    const expTbl = getTableData('fact_Material_Expenses');
    expTbl.rows.forEach(function (r) {
      if (expenseIds.indexOf(r.Expense_ID) !== -1 && r.Paid_By === 'Employee Out-of-Pocket' && r.Payroll_Status !== 'Paid') {
        expPatches[r.Expense_ID] = { Payroll_Status: nextStatus };
      }
    });
    const expsUpdated = batchPatchRowsById('fact_Material_Expenses', 'Expense_ID', expPatches);
    addAuditLog(user.email, 'BATCH_APPROVAL', 'Payroll', '', 'Set ' + nextStatus + ' on ' + logsUpdated + ' logs, ' + expsUpdated + ' expenses');
    return { ok: true, logsUpdated: logsUpdated, expensesUpdated: expsUpdated };
  });
}

function getWorkLogsAndExpenses(token, filters) {
  const user = requireAuth(token);
  const f = filters || {};
  const admin = isAdminRole(user.role);
  const logsTbl = getTableData('fact_Work_Logs');
  const expTbl = getTableData('fact_Material_Expenses');

  const filterRow = function (r, dateCol) {
    if (!admin && String(r.Employee_ID) !== user.employeeId) return false;
    if (f.employeeId && String(r.Employee_ID) !== String(f.employeeId)) return false;
    if (f.propertyId && String(r.Property_ID) !== String(f.propertyId)) return false;
    if (f.startDate && r[dateCol] && String(r[dateCol]) < f.startDate) return false;
    if (f.endDate && r[dateCol] && String(r[dateCol]) > f.endDate) return false;
    if (f.payrollStatus && f.payrollStatus !== 'All' && String(r.Payroll_Status) !== f.payrollStatus) return false;
    if (f.billingStatus && f.billingStatus !== 'All' && String(r.Billing_Status) !== f.billingStatus) return false;
    return true;
  };

  const logs = logsTbl.rows.filter(function (r) { return filterRow(r, 'Work_Date'); }).map(function (r) {
    const out = Object.assign({}, r);
    out.Effective_Billable = getEffectiveLogBillable(r);
    if (!admin) {
      delete out.Billable_Rate;
      delete out.Calculated_Billable;
      delete out.Charge_Amount_Custom;
      delete out.Effective_Billable;
      delete out.Rate_Source;
    }
    return out;
  });
  logs.sort(function (a, b) { return String(b.Work_Date || '').localeCompare(String(a.Work_Date || '')); });

  const expenses = expTbl.rows.filter(function (r) { return filterRow(r, 'Expense_Date'); }).map(function (r) {
    const out = Object.assign({}, r);
    if (!admin) {
      delete out.Markup_Pct;
      delete out.Billable_Amount;
    }
    return out;
  });
  expenses.sort(function (a, b) { return String(b.Expense_Date || '').localeCompare(String(a.Expense_Date || '')); });

  return sanitizeForWire({ workLogs: logs, expenses: expenses });
}

// ============================================================================
// ADMIN OVERVIEW & "NOTHING MISSED" COMMAND CENTER
// ============================================================================

function computeInvoiceDisplayStatus(inv, todayStr) {
  const s = String(inv.Status || 'Draft');
  if ((s === 'Sent' || s === 'Partial') && inv.Due_Date && String(inv.Due_Date) < (todayStr || todayYmd()) && toNum(inv.Balance_Due, 0) > 0.009) {
    return 'Overdue';
  }
  return s;
}

function getAdminOverview(token) {
  requireAuth(token, ['Owner', 'Admin']);
  const today = todayYmd();
  const currentMonth = today.slice(0, 7);

  const settings = getSettingsMap();
  const owners = getTableData('dim_Owners').rows;
  const props = getTableData('Properties').rows;
  const wos = getTableData('fact_Work_Orders').rows;
  const logs = getTableData('fact_Work_Logs').rows;
  const exps = getTableData('fact_Material_Expenses').rows;
  const invoices = getTableData('fact_Invoices').rows;

  const ownerById = {};
  owners.forEach(function (o) { ownerById[o.Owner_ID] = o; });
  const propById = {};
  props.forEach(function (p) { propById[p.Property_ID] = p; });

  // Unbilled Work Logs & Materials Pipeline
  const unbilledLogs = logs.filter(function (l) {
    return l.Billing_Status === 'Unbilled' && l.Charge_Target !== 'BPM Internal';
  });
  const unbilledExps = exps.filter(function (e) {
    return e.Billing_Status === 'Unbilled' && e.Charge_Target !== 'BPM Internal';
  });

  let unbilledTotal = 0;
  const ownerMap = {};

  function ensureOwnerBucket(propObj) {
    const ownerId = (propObj && propObj.Owner_ID) ? propObj.Owner_ID : 'UNASSIGNED';
    const owner = ownerById[ownerId] || { Owner_ID: ownerId, Owner_Name: propObj && propObj.Owner_Name ? propObj.Owner_Name : 'Unassigned Owner', Billing_Mode: 'Direct Invoice', Mgmt_Fee_Pct: settings.DEFAULT_MGMT_FEE_PCT };
    if (!ownerMap[ownerId]) {
      ownerMap[ownerId] = {
        ownerId: ownerId,
        ownerName: owner.Owner_Name,
        ownerEmail: owner.Email || '',
        billingMode: owner.Billing_Mode || 'Direct Invoice',
        mgmtFeePct: toNum(owner.Mgmt_Fee_Pct, toNum(settings.DEFAULT_MGMT_FEE_PCT, 8)),
        totalUnbilled: 0,
        properties: {}
      };
    }
    const pId = propObj ? propObj.Property_ID : 'UNKNOWN';
    if (!ownerMap[ownerId].properties[pId]) {
      ownerMap[ownerId].properties[pId] = {
        propertyId: pId,
        propertyName: propObj ? propObj.Property_Name : 'Unknown Property',
        laborItems: [],
        materialItems: [],
        totalLabor: 0,
        totalMaterials: 0,
        totalUnbilled: 0,
        hasPendingPayroll: false
      };
    }
    return { ownerBucket: ownerMap[ownerId], propBucket: ownerMap[ownerId].properties[pId] };
  }

  unbilledLogs.forEach(function (l) {
    const bill = getEffectiveLogBillable(l);
    unbilledTotal = round2(unbilledTotal + bill);
    const pObj = propById[l.Property_ID] || { Property_ID: l.Property_ID, Property_Name: l.Property_Name, Owner_ID: 'UNASSIGNED' };
    const buckets = ensureOwnerBucket(pObj);
    const item = Object.assign({}, l, { Effective_Billable: bill });
    buckets.propBucket.laborItems.push(item);
    buckets.propBucket.totalLabor = round2(buckets.propBucket.totalLabor + bill);
    buckets.propBucket.totalUnbilled = round2(buckets.propBucket.totalUnbilled + bill);
    buckets.ownerBucket.totalUnbilled = round2(buckets.ownerBucket.totalUnbilled + bill);
    if (l.Payroll_Status === 'Pending') buckets.propBucket.hasPendingPayroll = true;
  });

  unbilledExps.forEach(function (e) {
    const bill = round2(toNum(e.Billable_Amount, 0));
    unbilledTotal = round2(unbilledTotal + bill);
    const pObj = propById[e.Property_ID] || { Property_ID: e.Property_ID, Property_Name: e.Property_Name, Owner_ID: 'UNASSIGNED' };
    const buckets = ensureOwnerBucket(pObj);
    buckets.propBucket.materialItems.push(e);
    buckets.propBucket.totalMaterials = round2(buckets.propBucket.totalMaterials + bill);
    buckets.propBucket.totalUnbilled = round2(buckets.propBucket.totalUnbilled + bill);
    buckets.ownerBucket.totalUnbilled = round2(buckets.ownerBucket.totalUnbilled + bill);
    if (e.Payroll_Status === 'Pending') buckets.propBucket.hasPendingPayroll = true;
  });

  const unbilledByOwner = Object.keys(ownerMap).map(function (k) {
    const ob = ownerMap[k];
    ob.propertyList = Object.keys(ob.properties).map(function (pk) { return ob.properties[pk]; });
    delete ob.properties;
    return ob;
  }).sort(function (a, b) { return b.totalUnbilled - a.totalUnbilled; });

  // Open Work Orders & Alerts
  const openWos = wos.filter(function (w) {
    return ['Completed', 'Billed', 'Cancelled'].indexOf(w.Status) === -1;
  });
  const unassignedWos = openWos.filter(function (w) { return w.Status === 'New' || !w.Assigned_Employee_ID; });
  const emergencyWos = openWos.filter(function (w) { return w.Urgency === 'Emergency'; });
  const pendingLogs = logs.filter(function (l) { return l.Payroll_Status === 'Pending'; });
  const pendingExps = exps.filter(function (e) { return e.Paid_By === 'Employee Out-of-Pocket' && e.Payroll_Status === 'Pending'; });
  const fallbackLogs = unbilledLogs.filter(function (l) { return String(l.Rate_Source) === 'Fallback Rate'; });

  // Invoices & Receivables
  let overdueTotal = 0;
  let overdueCount = 0;
  let outstandingReceivables = 0;
  invoices.forEach(function (inv) {
    if (inv.Status === 'Void' || inv.Status === 'Draft') return;
    if (inv.Doc_Type === 'Owner Settlement' && toNum(inv.Balance_Due, 0) === 0) return;
    const bal = toNum(inv.Balance_Due, 0);
    if (bal > 0.009) {
      outstandingReceivables = round2(outstandingReceivables + bal);
      if (computeInvoiceDisplayStatus(inv, today) === 'Overdue') {
        overdueTotal = round2(overdueTotal + bal);
        overdueCount++;
      }
    }
  });

  // Current Rent Snapshot
  const rentSnap = getRentRollDataInternal({ cycleType: 'Academic', cycleLabel: inferActiveAcademicCycle(today) }, settings);
  let pastGraceCount = 0;
  let currentMonthExpected = 0;
  let currentMonthCollected = 0;
  rentSnap.leases.forEach(function (lease) {
    lease.tenants.forEach(function (t) {
      const mCell = t.months[currentMonth];
      if (mCell) {
        currentMonthExpected = round2(currentMonthExpected + toNum(mCell.scheduled, 0));
        currentMonthCollected = round2(currentMonthCollected + toNum(mCell.applied, 0));
      }
      if (t.hasPastDue) pastGraceCount++;
    });
  });

  // Action Alerts ("Nothing Missed")
  const alerts = [];
  if (emergencyWos.length > 0) {
    alerts.push({ severity: 'red', tab: 'work-orders', title: emergencyWos.length + ' Emergency Work Order(s)', detail: 'Immediate dispatch attention required.' });
  }
  if (unassignedWos.length > 0) {
    alerts.push({ severity: 'amber', tab: 'work-orders', title: unassignedWos.length + ' Unassigned Work Request(s)', detail: 'Assign crew so tenant requests do not sit.' });
  }
  if (pendingLogs.length + pendingExps.length > 0) {
    alerts.push({ severity: 'amber', tab: 'time-expenses', title: (pendingLogs.length + pendingExps.length) + ' Entries Awaiting Approval', detail: 'Review crew hours and out-of-pocket receipts.' });
  }
  if (fallbackLogs.length > 0) {
    alerts.push({ severity: 'blue', tab: 'time-expenses', title: fallbackLogs.length + ' Log(s) Using Fallback Rate ($' + settings.DEFAULT_FALLBACK_RATE + '/hr)', detail: 'Consider adding explicit task/property rates.' });
  }
  if (unbilledTotal > 0) {
    alerts.push({ severity: 'emerald', tab: 'billing', title: '$' + unbilledTotal.toFixed(2) + ' Unbilled Labor & Materials Ready', detail: 'Generate and send owner/tenant invoices.' });
  }
  if (overdueCount > 0) {
    alerts.push({ severity: 'red', tab: 'billing', title: overdueCount + ' Overdue Invoice(s) ($' + overdueTotal.toFixed(2) + ')', detail: 'Send payment reminders.' });
  }
  if (pastGraceCount > 0) {
    alerts.push({ severity: 'amber', tab: 'rentals', title: pastGraceCount + ' Tenant(s) Past Grace Day', detail: 'Review rent roll and apply late fees/reminders.' });
  }

  // Monthly 6-month trend for chart
  const trendMonths = buildRecentMonthKeys(6, currentMonth);
  const rentLedger = getTableData('fact_Rent_Ledger').rows;
  const payrollHist = getTableData('fact_Payroll_History').rows;
  const monthlyTrend = trendMonths.map(function (mk) {
    let billed = 0;
    let payroll = 0;
    let rent = 0;
    invoices.forEach(function (inv) {
      if (inv.Status !== 'Void' && String(inv.Issue_Date || '').slice(0, 7) === mk && inv.Doc_Type !== 'Owner Settlement') {
        billed = round2(billed + toNum(inv.Total_Amount, 0));
      }
    });
    payrollHist.forEach(function (p) {
      if (String(p.Run_Date || '').slice(0, 7) === mk) {
        payroll = round2(payroll + toNum(p.Net_Payout, 0));
      }
    });
    rentLedger.forEach(function (r) {
      if (String(r.Entry_Date || '').slice(0, 7) === mk && ['Rent', 'Upfront Prepay', 'Late Fee'].indexOf(r.Entry_Type) !== -1) {
        rent = round2(rent + toNum(r.Amount, 0));
      }
    });
    return { month: mk, label: formatMonthLabel(mk), billedRevenue: billed, payrollCost: payroll, rentCollected: rent };
  });

  return sanitizeForWire({
    kpis: {
      unbilledTotal: unbilledTotal,
      unbilledLogCount: unbilledLogs.length,
      unbilledExpenseCount: unbilledExps.length,
      pendingApprovalCount: pendingLogs.length + pendingExps.length,
      openWoCount: openWos.length,
      emergencyWoCount: emergencyWos.length,
      overdueInvoiceTotal: overdueTotal,
      overdueInvoiceCount: overdueCount,
      outstandingReceivables: outstandingReceivables,
      currentMonthRentExpected: currentMonthExpected,
      currentMonthRentCollected: currentMonthCollected,
      pastGraceTenantCount: pastGraceCount
    },
    alerts: alerts,
    unbilledByOwner: unbilledByOwner,
    monthlyTrend: monthlyTrend,
    recentLogs: logs.slice(-10).reverse()
  });
}

function inferActiveAcademicCycle(todayStr) {
  const y = Number((todayStr || todayYmd()).slice(0, 4));
  const m = Number((todayStr || todayYmd()).slice(5, 7));
  const startY = m >= 5 ? y : (y - 1);
  return String(startY).slice(-2) + '-' + String(startY + 1).slice(-2);
}

function buildRecentMonthKeys(count, endMonthKey) {
  const parts = endMonthKey.split('-').map(Number);
  const out = [];
  for (let i = count - 1; i >= 0; i--) {
    const dt = new Date(Date.UTC(parts[0], parts[1] - 1 - i, 1));
    out.push(dt.toISOString().slice(0, 7));
  }
  return out;
}

// ============================================================================
// INVOICING, MONTHLY OWNER SETTLEMENTS, EMAIL DELIVERY & PAYMENT ALLOCATION
// ============================================================================

function getBillingDashboard(token) {
  requireAuth(token, ['Owner', 'Admin']);
  const overview = getAdminOverview(token);
  const today = todayYmd();
  const invRows = getTableData('fact_Invoices').rows.map(function (inv) {
    const copy = Object.assign({}, inv);
    copy.Display_Status = computeInvoiceDisplayStatus(inv, today);
    return copy;
  }).sort(function (a, b) {
    return String(b.Issue_Date || '').localeCompare(String(a.Issue_Date || '')) ||
      String(b.Invoice_Number || '').localeCompare(String(a.Invoice_Number || ''));
  });
  const paymentRows = getTableData('fact_Invoice_Payments').rows.slice().reverse();
  return sanitizeForWire({
    unbilledByOwner: overview.unbilledByOwner,
    invoices: invRows,
    payments: paymentRows
  });
}

/**
 * Creates a Single-Property Invoice, Multi-Property Owner Invoice, Monthly Owner Settlement, or Tenant Charge.
 */
function createInvoice(token, payload) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  if (!payload) throw new Error('Invoice payload required.');

  return withScriptLock(function () {
    const settings = getSettingsMap();
    const docType = ['Invoice', 'Owner Settlement', 'Tenant Charge'].indexOf(payload.docType) !== -1
      ? payload.docType
      : 'Invoice';

    const ownersTbl = getTableData('dim_Owners');
    const propsTbl = getTableData('Properties');
    const logsTbl = getTableData('fact_Work_Logs');
    const expTbl = getTableData('fact_Material_Expenses');

    const logIds = Array.isArray(payload.logIds) ? payload.logIds : [];
    const expIds = Array.isArray(payload.expenseIds) ? payload.expenseIds : [];
    const propIds = Array.isArray(payload.propertyIds) ? payload.propertyIds : [];

    const selectedLogs = logsTbl.rows.filter(function (l) {
      return logIds.indexOf(l.Log_ID) !== -1 && l.Billing_Status === 'Unbilled';
    });
    const selectedExps = expTbl.rows.filter(function (e) {
      return expIds.indexOf(e.Expense_ID) !== -1 && e.Billing_Status === 'Unbilled';
    });

    if (docType !== 'Owner Settlement' && selectedLogs.length === 0 && selectedExps.length === 0 && toNum(payload.adjustmentsAmount, 0) === 0) {
      throw new Error('Select at least one unbilled labor entry, material expense, or adjustment.');
    }

    const firstProp = propsTbl.rows.find(function (p) {
      return propIds.indexOf(p.Property_ID) !== -1 ||
        (selectedLogs[0] && p.Property_ID === selectedLogs[0].Property_ID) ||
        (selectedExps[0] && p.Property_ID === selectedExps[0].Property_ID);
    });
    const ownerId = payload.ownerId || (firstProp ? firstProp.Owner_ID : 'OWN-001');
    const owner = ownersTbl.rows.find(function (o) { return String(o.Owner_ID) === String(ownerId); }) || {
      Owner_ID: ownerId,
      Owner_Name: payload.ownerName || (firstProp ? firstProp.Owner_Name : 'Property Owner'),
      Email: '',
      Mgmt_Fee_Pct: settings.DEFAULT_MGMT_FEE_PCT,
      Payment_Terms_Days: settings.INVOICE_DUE_DAYS
    };

    // Determine distinct properties covered
    const propSet = {};
    propIds.forEach(function (id) {
      const p = propsTbl.rows.find(function (r) { return r.Property_ID === id; });
      if (p) propSet[id] = p.Property_Name;
    });
    selectedLogs.forEach(function (l) { propSet[l.Property_ID] = l.Property_Name; });
    selectedExps.forEach(function (e) { propSet[e.Property_ID] = e.Property_Name; });
    const propKeys = Object.keys(propSet);
    const propertyIdField = propKeys.length > 1 ? 'MULTI' : (propKeys[0] || '');
    const propertyNameField = propKeys.length > 1
      ? ('Multiple (' + propKeys.map(function (k) { return propSet[k]; }).join(', ') + ')')
      : (propSet[propKeys[0]] || payload.propertyName || '');

    let laborSubtotal = 0;
    const lineItems = [];
    selectedLogs.forEach(function (l) {
      const amt = getEffectiveLogBillable(l);
      laborSubtotal = round2(laborSubtotal + amt);
      lineItems.push({
        kind: 'Labor',
        refId: l.Log_ID,
        date: l.Work_Date,
        propertyName: l.Property_Name,
        description: l.Task_Name + (l.Notes ? ' — ' + l.Notes : ''),
        qty: l.Work_Type === 'Flat' ? 'Flat' : (l.Hours + ' hrs'),
        unitRate: toNum(l.Billable_Rate, 0),
        amount: amt
      });
    });

    let materialsSubtotal = 0;
    selectedExps.forEach(function (e) {
      const amt = round2(toNum(e.Billable_Amount, 0));
      materialsSubtotal = round2(materialsSubtotal + amt);
      lineItems.push({
        kind: 'Material',
        refId: e.Expense_ID,
        date: e.Expense_Date,
        propertyName: e.Property_Name,
        description: (e.Vendor ? e.Vendor + ': ' : '') + e.Item_Description,
        qty: e.Quantity || 1,
        unitRate: toNum(e.Unit_Cost, amt),
        amount: amt
      });
    });

    const periodStart = toYmd(payload.periodStart) || (selectedLogs[0] ? selectedLogs[0].Work_Date : todayYmd().slice(0, 8) + '01');
    const periodEnd = toYmd(payload.periodEnd) || todayYmd();
    const issueDate = toYmd(payload.issueDate) || todayYmd();
    const termsDays = toNum(owner.Payment_Terms_Days, toNum(settings.INVOICE_DUE_DAYS, 15));
    const dueDate = toYmd(payload.dueDate) || addDaysYmd(issueDate, termsDays);
    const adjustments = round2(toNum(payload.adjustmentsAmount, 0));

    let rentCollected = 0;
    let mgmtFeePct = toNum(payload.mgmtFeePct, toNum(owner.Mgmt_Fee_Pct, toNum(settings.DEFAULT_MGMT_FEE_PCT, 8)));
    let mgmtFeeAmount = 0;
    let totalAmount = 0;
    let netOwnerPayout = 0;

    if (docType === 'Owner Settlement') {
      const rentTbl = getTableData('fact_Rent_Ledger').rows;
      rentTbl.forEach(function (r) {
        if (propKeys.length > 0 && propKeys.indexOf(r.Property_ID) === -1) return;
        const d = toYmd(r.Entry_Date);
        if (d < periodStart || d > periodEnd) return;
        if (['Rent', 'Upfront Prepay', 'Late Fee', 'Utility/Misc'].indexOf(r.Entry_Type) !== -1) {
          rentCollected = round2(rentCollected + toNum(r.Amount, 0));
        } else if (r.Entry_Type === 'Credit/Waiver') {
          rentCollected = round2(rentCollected - toNum(r.Amount, 0));
        }
      });
      mgmtFeeAmount = round2(rentCollected * (mgmtFeePct / 100));
      netOwnerPayout = round2(rentCollected - mgmtFeeAmount - laborSubtotal - materialsSubtotal + adjustments);
      // Receivable by BPM is >0 only when deductions exceed collected rent
      totalAmount = netOwnerPayout < 0 ? round2(Math.abs(netOwnerPayout)) : 0;
    } else {
      totalAmount = round2(laborSubtotal + materialsSubtotal + adjustments);
    }

    const invoiceId = newId('INV');
    const invoiceNumber = nextSeqId('BPM', 'NEXT_INVOICE_SEQ', 1001);
    const saveAsDraft = Boolean(payload.saveAsDraft);
    const status = saveAsDraft ? 'Draft' : (docType === 'Owner Settlement' && totalAmount === 0 ? 'Paid' : 'Sent');

    const lineSummaryObj = {
      summary: {
        isSettlement: docType === 'Owner Settlement',
        mgmtFeePct: mgmtFeePct,
        netOwnerPayout: netOwnerPayout,
        adjustmentNote: String(payload.adjustmentNote || '')
      },
      items: lineItems
    };

    const invoiceRec = {
      Invoice_ID: invoiceId,
      Invoice_Number: invoiceNumber,
      Doc_Type: docType,
      Owner_ID: owner.Owner_ID,
      Owner_Name: owner.Owner_Name,
      Recipient_Email: String(payload.recipientEmail || owner.Email || ''),
      Property_ID: propertyIdField,
      Property_Name: propertyNameField,
      Period_Start: periodStart,
      Period_End: periodEnd,
      Issue_Date: issueDate,
      Due_Date: dueDate,
      Status: status,
      Labor_Subtotal: laborSubtotal,
      Materials_Subtotal: materialsSubtotal,
      Rent_Collected: rentCollected,
      Mgmt_Fee_Amount: mgmtFeeAmount,
      Adjustments_Amount: adjustments,
      Total_Amount: totalAmount,
      Amount_Paid: 0,
      Balance_Due: totalAmount,
      Sent_At: '',
      Last_Reminder_At: '',
      PDF_URL: '',
      Line_Items_JSON: JSON.stringify(lineSummaryObj),
      Notes_To_Client: String(payload.notesToClient || ''),
      Created_By: user.email
    };

    const html = buildInvoiceHtml(invoiceRec, lineSummaryObj, owner, settings);
    const pdfRes = renderAndSavePdf('INVOICES_FOLDER_ID', 'Invoices & Settlements', invoiceNumber + '.pdf', html);
    invoiceRec.PDF_URL = pdfRes.pdfUrl;

    appendRecords('fact_Invoices', [invoiceRec]);

    const nextItemBillStatus = saveAsDraft
      ? 'Draft Invoice'
      : (docType === 'Owner Settlement' && totalAmount === 0 ? 'Paid' : 'Invoiced');

    const logPatches = {};
    const woIdsToMarkBilled = {};
    selectedLogs.forEach(function (l) {
      logPatches[l.Log_ID] = { Billing_Status: nextItemBillStatus, Invoice_ID: invoiceId };
      if (!saveAsDraft && l.WO_ID) woIdsToMarkBilled[l.WO_ID] = true;
    });
    batchPatchRowsById('fact_Work_Logs', 'Log_ID', logPatches);

    const expPatches = {};
    selectedExps.forEach(function (e) {
      expPatches[e.Expense_ID] = { Billing_Status: nextItemBillStatus, Invoice_ID: invoiceId };
      if (!saveAsDraft && e.WO_ID) woIdsToMarkBilled[e.WO_ID] = true;
    });
    batchPatchRowsById('fact_Material_Expenses', 'Expense_ID', expPatches);

    const woTbl = getTableData('fact_Work_Orders');
    const woPatches = {};
    woTbl.rows.forEach(function (w) {
      if (woIdsToMarkBilled[w.WO_ID] && w.Status === 'Completed') {
        woPatches[w.WO_ID] = { Status: 'Billed' };
      }
    });
    batchPatchRowsById('fact_Work_Orders', 'WO_ID', woPatches);

    addAuditLog(user.email, 'CREATE_INVOICE', 'Invoice', invoiceId, invoiceNumber + ' (' + docType + ') $' + totalAmount);
    return sanitizeForWire({
      ok: true,
      invoice: invoiceRec,
      html: html,
      pdfBase64: pdfRes.pdfBase64,
      pdfFilename: invoiceNumber + '.pdf'
    });
  });
}

function buildInvoiceHtml(inv, lineObj, owner, settings) {
  const s = lineObj.summary || {};
  const items = lineObj.items || [];
  const rowsHtml = items.map(function (it) {
    return (
      '<tr>' +
      '<td style="padding:8px;border-bottom:1px solid #e2e8f0">' + escapeHtml(it.date || '') + '</td>' +
      '<td style="padding:8px;border-bottom:1px solid #e2e8f0">' + escapeHtml(it.propertyName || '') + '</td>' +
      '<td style="padding:8px;border-bottom:1px solid #e2e8f0"><strong>' + escapeHtml(it.kind) + '</strong>: ' + escapeHtml(it.description) + '</td>' +
      '<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right">' + escapeHtml(String(it.qty || '')) + '</td>' +
      '<td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right">$' + round2(it.amount).toFixed(2) + '</td>' +
      '</tr>'
    );
  }).join('');

  let summaryBox = '';
  if (s.isSettlement) {
    summaryBox =
      '<div style="background:#f8fafc;border:1px solid #cbd5e1;border-radius:8px;padding:16px;margin-top:18px">' +
      '<h3 style="margin:0 0 10px;color:#0f172a">Owner Settlement Summary (' + escapeHtml(inv.Period_Start) + ' to ' + escapeHtml(inv.Period_End) + ')</h3>' +
      '<table style="width:100%;font-size:14px">' +
      '<tr><td>Rent & Fee Collections</td><td style="text-align:right;color:#047857">+' + '$' + round2(inv.Rent_Collected).toFixed(2) + '</td></tr>' +
      '<tr><td>Management Fee (' + round2(s.mgmtFeePct) + '%)</td><td style="text-align:right;color:#b91c1c">-$' + round2(inv.Mgmt_Fee_Amount).toFixed(2) + '</td></tr>' +
      '<tr><td>Maintenance Labor</td><td style="text-align:right;color:#b91c1c">-$' + round2(inv.Labor_Subtotal).toFixed(2) + '</td></tr>' +
      '<tr><td>Materials & Parts</td><td style="text-align:right;color:#b91c1c">-$' + round2(inv.Materials_Subtotal).toFixed(2) + '</td></tr>' +
      (toNum(inv.Adjustments_Amount, 0) !== 0 ? '<tr><td>Adjustments</td><td style="text-align:right">$' + round2(inv.Adjustments_Amount).toFixed(2) + '</td></tr>' : '') +
      '<tr style="font-size:18px;font-weight:bold;border-top:2px solid #94a3b8"><td style="padding-top:8px">' +
      (s.netOwnerPayout >= 0 ? 'Net Owner Distribution' : 'Balance Due From Owner') +
      '</td><td style="padding-top:8px;text-align:right">$' + round2(Math.abs(s.netOwnerPayout)).toFixed(2) + '</td></tr>' +
      '</table></div>';
  } else {
    summaryBox =
      '<div style="width:320px;margin-left:auto;margin-top:18px;background:#f8fafc;border:1px solid #cbd5e1;border-radius:8px;padding:14px">' +
      '<table style="width:100%;font-size:14px">' +
      '<tr><td>Labor Subtotal</td><td style="text-align:right">$' + round2(inv.Labor_Subtotal).toFixed(2) + '</td></tr>' +
      '<tr><td>Materials Subtotal</td><td style="text-align:right">$' + round2(inv.Materials_Subtotal).toFixed(2) + '</td></tr>' +
      (toNum(inv.Adjustments_Amount, 0) !== 0 ? '<tr><td>Adjustments' + (s.adjustmentNote ? ' (' + escapeHtml(s.adjustmentNote) + ')' : '') + '</td><td style="text-align:right">$' + round2(inv.Adjustments_Amount).toFixed(2) + '</td></tr>' : '') +
      '<tr style="font-size:17px;font-weight:bold;border-top:2px solid #94a3b8"><td style="padding-top:8px">Total</td><td style="padding-top:8px;text-align:right">$' + round2(inv.Total_Amount).toFixed(2) + '</td></tr>' +
      (toNum(inv.Amount_Paid, 0) > 0 ? '<tr><td>Paid</td><td style="text-align:right">-$' + round2(inv.Amount_Paid).toFixed(2) + '</td></tr>' +
        '<tr style="font-weight:bold;color:#1e40af"><td>Balance Due</td><td style="text-align:right">$' + round2(inv.Balance_Due).toFixed(2) + '</td></tr>' : '') +
      '</table></div>';
  }

  const payLinkBtn = settings.PAYMENT_URL
    ? '<p style="margin-top:10px"><a href="' + escapeHtml(settings.PAYMENT_URL) + '" style="display:inline-block;background:#1d4ed8;color:#fff;text-decoration:none;padding:10px 18px;border-radius:6px;font-weight:bold">Pay Online Now</a></p>'
    : '';

  return (
    '<!DOCTYPE html><html><head><meta charset="utf-8"><style>body{font-family:Helvetica,Arial,sans-serif;color:#1e293b;margin:28px;font-size:13px}</style></head><body>' +
    '<div style="display:flex;justify-content:space-between;align-items:flex-start;border-bottom:3px solid #1e3a8a;padding-bottom:14px">' +
    '<div><h1 style="margin:0;color:#1e3a8a;font-size:22px">' + escapeHtml(settings.COMPANY_NAME) + '</h1>' +
    '<p style="margin:4px 0 0;color:#475569">' + escapeHtml(settings.COMPANY_ADDRESS) + '<br>' +
    escapeHtml(settings.COMPANY_PHONE) + ' &bull; ' + escapeHtml(settings.DISPATCH_EMAIL) + '</p></div>' +
    '<div style="text-align:right"><h2 style="margin:0;color:#0f172a">' + escapeHtml(inv.Doc_Type.toUpperCase()) + '</h2>' +
    '<p style="margin:4px 0 0"><strong>#' + escapeHtml(inv.Invoice_Number) + '</strong><br>Issued: ' + escapeHtml(inv.Issue_Date) + '<br>Due: ' + escapeHtml(inv.Due_Date) + '</p></div></div>' +
    '<div style="margin:18px 0"><strong>Bill To / Owner:</strong> ' + escapeHtml(inv.Owner_Name || owner.Owner_Name || '') +
    (inv.Property_Name ? '<br><strong>Property:</strong> ' + escapeHtml(inv.Property_Name) : '') +
    (inv.Recipient_Email ? '<br><strong>Email:</strong> ' + escapeHtml(inv.Recipient_Email) : '') + '</div>' +
    '<table style="width:100%;border-collapse:collapse;font-size:12px">' +
    '<thead><tr style="background:#f1f5f9;text-align:left">' +
    '<th style="padding:8px">Date</th><th style="padding:8px">Property</th><th style="padding:8px">Description</th>' +
    '<th style="padding:8px;text-align:right">Qty / Hrs</th><th style="padding:8px;text-align:right">Amount</th>' +
    '</tr></thead><tbody>' + (rowsHtml || '<tr><td colspan="5" style="padding:10px;color:#64748b">No itemized maintenance charges in period.</td></tr>') + '</tbody></table>' +
    summaryBox +
    (inv.Notes_To_Client ? '<p style="margin-top:16px"><strong>Notes:</strong> ' + escapeHtml(inv.Notes_To_Client) + '</p>' : '') +
    '<div style="margin-top:24px;padding:14px;background:#eff6ff;border:1px solid #bfdbfe;border-radius:8px">' +
    '<strong style="color:#1e3a8a">Payment Instructions</strong><p style="margin:6px 0 0">' + escapeHtml(settings.PAYMENT_INSTRUCTIONS || '') + '</p>' +
    payLinkBtn + '</div></body></html>'
  );
}

function finalizeOrUpdateDraftInvoice(token, payload) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  if (!payload || !payload.invoiceId) throw new Error('invoiceId is required.');
  return withScriptLock(function () {
    const inv = getTableData('fact_Invoices').rows.find(function (r) { return String(r.Invoice_ID) === String(payload.invoiceId); });
    if (!inv) throw new Error('Invoice not found.');
    const settings = getSettingsMap();
    const owner = getTableData('dim_Owners').rows.find(function (o) { return String(o.Owner_ID) === String(inv.Owner_ID); }) || {};
    let lineObj = { summary: {}, items: [] };
    try { lineObj = JSON.parse(inv.Line_Items_JSON || '{}'); } catch (e) {}

    const adjustments = payload.adjustmentsAmount !== undefined ? round2(toNum(payload.adjustmentsAmount, inv.Adjustments_Amount)) : toNum(inv.Adjustments_Amount, 0);
    if (payload.adjustmentNote !== undefined) {
      lineObj.summary = lineObj.summary || {};
      lineObj.summary.adjustmentNote = String(payload.adjustmentNote);
    }

    let newTotal = round2(toNum(inv.Labor_Subtotal, 0) + toNum(inv.Materials_Subtotal, 0) + adjustments);
    if (inv.Doc_Type === 'Owner Settlement') {
      const net = round2(toNum(inv.Rent_Collected, 0) - toNum(inv.Mgmt_Fee_Amount, 0) - toNum(inv.Labor_Subtotal, 0) - toNum(inv.Materials_Subtotal, 0) + adjustments);
      lineObj.summary.netOwnerPayout = net;
      newTotal = net < 0 ? round2(Math.abs(net)) : 0;
    }

    const markSent = Boolean(payload.finalizeSent) && inv.Status === 'Draft';
    const nextStatus = markSent ? (inv.Doc_Type === 'Owner Settlement' && newTotal === 0 ? 'Paid' : 'Sent') : inv.Status;
    const paid = round2(toNum(inv.Amount_Paid, 0));
    const bal = Math.max(0, round2(newTotal - paid));

    const updatedDraft = Object.assign({}, inv, {
      Recipient_Email: payload.recipientEmail !== undefined ? String(payload.recipientEmail) : inv.Recipient_Email,
      Due_Date: payload.dueDate ? toYmd(payload.dueDate) : inv.Due_Date,
      Adjustments_Amount: adjustments,
      Total_Amount: newTotal,
      Balance_Due: bal,
      Notes_To_Client: payload.notesToClient !== undefined ? String(payload.notesToClient) : inv.Notes_To_Client,
      Line_Items_JSON: JSON.stringify(lineObj),
      Status: nextStatus
    });

    const html = buildInvoiceHtml(updatedDraft, lineObj, owner, settings);
    const pdf = renderAndSavePdf('INVOICES_FOLDER_ID', 'Invoices & Settlements', inv.Invoice_Number + '.pdf', html);
    updatedDraft.PDF_URL = pdf.pdfUrl || inv.PDF_URL;

    updateRowById('fact_Invoices', 'Invoice_ID', inv.Invoice_ID, updatedDraft);
    if (markSent) {
      promoteDraftItemsToInvoiced(inv.Invoice_ID);
    }
    addAuditLog(user.email, 'UPDATE_INVOICE', 'Invoice', inv.Invoice_ID, 'Status ' + nextStatus + ' Total $' + newTotal);
    return sanitizeForWire({ ok: true, invoice: updatedDraft, pdfBase64: pdf.pdfBase64, html: html });
  });
}

function getInvoicePdfData(token, invoiceId) {
  requireAuth(token, ['Owner', 'Admin']);
  const inv = getTableData('fact_Invoices').rows.find(function (r) { return String(r.Invoice_ID) === String(invoiceId); });
  if (!inv) throw new Error('Invoice not found.');
  const settings = getSettingsMap();
  const owner = getTableData('dim_Owners').rows.find(function (o) { return String(o.Owner_ID) === String(inv.Owner_ID); }) || {};
  let lineObj = { summary: {}, items: [] };
  try {
    lineObj = JSON.parse(inv.Line_Items_JSON || '{}');
  } catch (err) {
    lineObj = { summary: {}, items: [] };
  }
  const html = buildInvoiceHtml(inv, lineObj, owner, settings);
  const pdf = renderAndSavePdf('INVOICES_FOLDER_ID', 'Invoices & Settlements', inv.Invoice_Number + '.pdf', html);
  return sanitizeForWire({
    invoice: inv,
    html: html,
    pdfUrl: inv.PDF_URL || pdf.pdfUrl,
    pdfBase64: pdf.pdfBase64,
    pdfFilename: inv.Invoice_Number + '.pdf'
  });
}

function sendInvoiceEmail(token, payload) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  if (!payload || !payload.invoiceId) throw new Error('invoiceId required.');

  return withScriptLock(function () {
    const inv = getTableData('fact_Invoices').rows.find(function (r) { return String(r.Invoice_ID) === String(payload.invoiceId); });
    if (!inv) throw new Error('Invoice not found.');
    const toEmail = String(payload.toEmail || inv.Recipient_Email || '').trim();
    if (!toEmail) throw new Error('Recipient email address is required.');

    const settings = getSettingsMap();
    const owner = getTableData('dim_Owners').rows.find(function (o) { return String(o.Owner_ID) === String(inv.Owner_ID); }) || {};
    let lineObj = { summary: {}, items: [] };
    try { lineObj = JSON.parse(inv.Line_Items_JSON || '{}'); } catch (e) {}
    const html = buildInvoiceHtml(inv, lineObj, owner, settings);
    const pdfRes = renderAndSavePdf('INVOICES_FOLDER_ID', 'Invoices & Settlements', inv.Invoice_Number + '.pdf', html);

    const subject = String(payload.subject || (settings.COMPANY_NAME + ' — ' + inv.Doc_Type + ' ' + inv.Invoice_Number));
    const msg = String(payload.messageBody || ('Attached is your ' + inv.Doc_Type + ' #' + inv.Invoice_Number + '.'));
    const payUrlSection = settings.PAYMENT_URL
      ? '<p><a href="' + escapeHtml(settings.PAYMENT_URL) + '" style="display:inline-block;background:#1d4ed8;color:#fff;padding:10px 18px;border-radius:6px;text-decoration:none;font-weight:bold">View Payment Portal</a></p>'
      : '';

    const emailOpts = {
      to: toEmail,
      replyTo: settings.DISPATCH_EMAIL || 'butlerpropmgmt@gmail.com',
      subject: subject,
      body: msg + '\n\n' + (settings.PAYMENT_INSTRUCTIONS || '') + (settings.PAYMENT_URL ? '\nPay online: ' + settings.PAYMENT_URL : ''),
      htmlBody:
        '<div style="font-family:Arial,sans-serif;max-width:600px">' +
        '<p>' + escapeHtml(msg).replace(/\n/g, '<br>') + '</p>' +
        '<div style="background:#f8fafc;padding:12px;border-radius:8px;border:1px solid #e2e8f0;margin:14px 0">' +
        '<strong>Payment Instructions:</strong><br>' + escapeHtml(settings.PAYMENT_INSTRUCTIONS || '') + '</div>' +
        payUrlSection + '</div>',
      attachments: [pdfRes.blob]
    };
    const cc = String(payload.ccEmails || owner.CC_Emails || '').trim();
    if (cc) emailOpts.cc = cc;

    MailApp.sendEmail(emailOpts);

    const nextStatus = inv.Status === 'Paid' ? 'Paid' : (toNum(inv.Amount_Paid, 0) > 0 ? 'Partial' : 'Sent');
    const updatedInv = updateRowById('fact_Invoices', 'Invoice_ID', inv.Invoice_ID, {
      Recipient_Email: toEmail,
      Status: nextStatus,
      Sent_At: nowIso(),
      PDF_URL: inv.PDF_URL || pdfRes.pdfUrl
    });

    // Promote any Draft Invoice linked items to Invoiced
    promoteDraftItemsToInvoiced(inv.Invoice_ID);
    addAuditLog(user.email, 'SEND_INVOICE_EMAIL', 'Invoice', inv.Invoice_ID, 'Sent to ' + toEmail);
    return sanitizeForWire({ ok: true, invoice: updatedInv });
  });
}

function sendInvoiceReminder(token, payload) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  const invId = typeof payload === 'string' ? payload : payload.invoiceId;
  const inv = getTableData('fact_Invoices').rows.find(function (r) { return String(r.Invoice_ID) === String(invId); });
  if (!inv) throw new Error('Invoice not found.');
  const customNote = (payload && payload.customNote)
    ? payload.customNote
    : ('Friendly reminder: ' + inv.Doc_Type + ' #' + inv.Invoice_Number + ' has an open balance of $' + round2(inv.Balance_Due).toFixed(2) + ' (due ' + inv.Due_Date + ').');
  const res = sendInvoiceEmail(token, {
    invoiceId: inv.Invoice_ID,
    toEmail: (payload && payload.toEmail) || inv.Recipient_Email,
    subject: 'Payment Reminder: ' + inv.Invoice_Number + ' ($' + round2(inv.Balance_Due).toFixed(2) + ' Due)',
    messageBody: customNote
  });
  updateRowById('fact_Invoices', 'Invoice_ID', inv.Invoice_ID, { Last_Reminder_At: nowIso() });
  addAuditLog(user.email, 'SEND_REMINDER', 'Invoice', inv.Invoice_ID, 'Reminder sent');
  return res;
}

function promoteDraftItemsToInvoiced(invoiceId) {
  const woIds = {};
  const logsTbl = getTableData('fact_Work_Logs');
  const logPatches = {};
  logsTbl.rows.forEach(function (l) {
    if (String(l.Invoice_ID) === String(invoiceId)) {
      if (l.WO_ID) woIds[l.WO_ID] = true;
      if (l.Billing_Status === 'Draft Invoice') {
        logPatches[l.Log_ID] = { Billing_Status: 'Invoiced' };
      }
    }
  });
  batchPatchRowsById('fact_Work_Logs', 'Log_ID', logPatches);

  const expTbl = getTableData('fact_Material_Expenses');
  const expPatches = {};
  expTbl.rows.forEach(function (e) {
    if (String(e.Invoice_ID) === String(invoiceId)) {
      if (e.WO_ID) woIds[e.WO_ID] = true;
      if (e.Billing_Status === 'Draft Invoice') {
        expPatches[e.Expense_ID] = { Billing_Status: 'Invoiced' };
      }
    }
  });
  batchPatchRowsById('fact_Material_Expenses', 'Expense_ID', expPatches);

  const woTbl = getTableData('fact_Work_Orders');
  const woPatches = {};
  woTbl.rows.forEach(function (w) {
    if (woIds[w.WO_ID] && w.Status === 'Completed') {
      woPatches[w.WO_ID] = { Status: 'Billed' };
    }
  });
  batchPatchRowsById('fact_Work_Orders', 'WO_ID', woPatches);
}

/**
 * Records a single-invoice payment OR a lump-sum multi-invoice batch allocation.
 */
function recordInvoicePayment(token, payload) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  if (!payload) throw new Error('Payment payload required.');

  return withScriptLock(function () {
    const invTbl = getTableData('fact_Invoices');
    const allocations = Array.isArray(payload.allocations) && payload.allocations.length > 0
      ? payload.allocations
      : [{ invoiceId: payload.invoiceId, amount: payload.amount }];

    const batchId = allocations.length > 1 ? newId('BATCH') : '';
    const paymentDate = toYmd(payload.paymentDate) || todayYmd();
    const method = String(payload.paymentMethod || 'ACH');
    const refNum = String(payload.referenceNumber || '');
    const notes = String(payload.notes || '');

    const paymentRecords = [];
    const invPatches = {};
    const paidInvoiceIds = [];

    for (let i = 0; i < allocations.length; i++) {
      const alloc = allocations[i];
      const amt = round2(toNum(alloc.amount, 0));
      if (amt <= 0) continue;
      const inv = invTbl.rows.find(function (r) { return String(r.Invoice_ID) === String(alloc.invoiceId); });
      if (!inv) throw new Error('Invoice not found: ' + alloc.invoiceId);
      if (inv.Status === 'Void') throw new Error('Cannot apply payment to void invoice ' + inv.Invoice_Number);

      const newPaid = round2(toNum(inv.Amount_Paid, 0) + amt);
      const newBal = Math.max(0, round2(toNum(inv.Total_Amount, 0) - newPaid));
      const newStatus = newBal <= 0.009 ? 'Paid' : 'Partial';

      invPatches[inv.Invoice_ID] = {
        Amount_Paid: newPaid,
        Balance_Due: newBal,
        Status: newStatus
      };
      if (newStatus === 'Paid') paidInvoiceIds.push(inv.Invoice_ID);

      paymentRecords.push({
        Payment_ID: newId('PMT'),
        Invoice_ID: inv.Invoice_ID,
        Invoice_Number: inv.Invoice_Number,
        Owner_ID: inv.Owner_ID,
        Payment_Date: paymentDate,
        Amount: amt,
        Payment_Method: method,
        Reference_Number: refNum,
        Batch_Payment_ID: batchId,
        Recorded_By: user.email,
        Notes: notes,
        Created_At: nowIso()
      });
    }

    if (paymentRecords.length === 0) throw new Error('Payment amount must be greater than $0.');
    appendRecords('fact_Invoice_Payments', paymentRecords);
    batchPatchRowsById('fact_Invoices', 'Invoice_ID', invPatches);

    // Stamp linked work logs and materials as Paid for any invoice settled in full
    if (paidInvoiceIds.length > 0) {
      const logsTbl = getTableData('fact_Work_Logs');
      const logPatches = {};
      logsTbl.rows.forEach(function (l) {
        if (paidInvoiceIds.indexOf(String(l.Invoice_ID)) !== -1) {
          logPatches[l.Log_ID] = { Billing_Status: 'Paid' };
        }
      });
      batchPatchRowsById('fact_Work_Logs', 'Log_ID', logPatches);

      const expTbl = getTableData('fact_Material_Expenses');
      const expPatches = {};
      expTbl.rows.forEach(function (e) {
        if (paidInvoiceIds.indexOf(String(e.Invoice_ID)) !== -1) {
          expPatches[e.Expense_ID] = { Billing_Status: 'Paid' };
        }
      });
      batchPatchRowsById('fact_Material_Expenses', 'Expense_ID', expPatches);
    }

    addAuditLog(user.email, 'RECORD_INVOICE_PAYMENT', 'Payment', batchId || paymentRecords[0].Payment_ID, 'Recorded ' + paymentRecords.length + ' allocation(s)');
    return sanitizeForWire({ ok: true, payments: paymentRecords, paidInvoiceIds: paidInvoiceIds });
  });
}

/**
 * Voids an invoice and releases all linked work logs and material expenses back to Unbilled.
 */
function voidInvoice(token, invoiceId, reason, forceVoid) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  return withScriptLock(function () {
    const inv = getTableData('fact_Invoices').rows.find(function (r) { return String(r.Invoice_ID) === String(invoiceId); });
    if (!inv) throw new Error('Invoice not found.');
    if (toNum(inv.Amount_Paid, 0) > 0 && !forceVoid) {
      throw new Error('Invoice has recorded payments. Refund or confirm void explicitly.');
    }

    const updated = updateRowById('fact_Invoices', 'Invoice_ID', invoiceId, {
      Status: 'Void',
      Balance_Due: 0,
      Notes_To_Client: (inv.Notes_To_Client ? inv.Notes_To_Client + ' | ' : '') + 'VOIDED: ' + (reason || 'Cancelled')
    });

    const logsTbl = getTableData('fact_Work_Logs');
    const logPatches = {};
    logsTbl.rows.forEach(function (l) {
      if (String(l.Invoice_ID) === String(invoiceId)) {
        logPatches[l.Log_ID] = {
          Billing_Status: l.Charge_Target === 'BPM Internal' ? 'Internal/Non-Billable' : 'Unbilled',
          Invoice_ID: ''
        };
      }
    });
    const releasedLogs = batchPatchRowsById('fact_Work_Logs', 'Log_ID', logPatches);

    const expTbl = getTableData('fact_Material_Expenses');
    const expPatches = {};
    expTbl.rows.forEach(function (e) {
      if (String(e.Invoice_ID) === String(invoiceId)) {
        expPatches[e.Expense_ID] = {
          Billing_Status: e.Charge_Target === 'BPM Internal' ? 'Internal/Non-Billable' : 'Unbilled',
          Invoice_ID: ''
        };
      }
    });
    const releasedExps = batchPatchRowsById('fact_Material_Expenses', 'Expense_ID', expPatches);

    const woIdsToReopen = {};
    logsTbl.rows.forEach(function (l) {
      if (String(l.Invoice_ID) === String(invoiceId) && l.WO_ID) woIdsToReopen[l.WO_ID] = true;
    });
    expTbl.rows.forEach(function (e) {
      if (String(e.Invoice_ID) === String(invoiceId) && e.WO_ID) woIdsToReopen[e.WO_ID] = true;
    });
    const woTbl = getTableData('fact_Work_Orders');
    const woPatches = {};
    woTbl.rows.forEach(function (w) {
      if (woIdsToReopen[w.WO_ID] && w.Status === 'Billed') {
        woPatches[w.WO_ID] = { Status: 'Completed' };
      }
    });
    batchPatchRowsById('fact_Work_Orders', 'WO_ID', woPatches);

    addAuditLog(user.email, 'VOID_INVOICE', 'Invoice', invoiceId, 'Released ' + releasedLogs + ' logs, ' + releasedExps + ' expenses');
    return sanitizeForWire({ ok: true, invoice: updated, releasedLogs: releasedLogs, releasedExpenses: releasedExps });
  });
}

// ============================================================================
// RENTALS, LEASES, ACADEMIC/CALENDAR RENT ROLL & PREPAYMENT ROLLOVER
// ============================================================================

function getCycleMonths(cycleType, cycleLabel, leasesRows) {
  if (cycleType === 'Calendar') {
    const yr = parseInt(String(cycleLabel || new Date().getFullYear()).slice(0, 4), 10) || new Date().getFullYear();
    const months = [];
    for (let m = 1; m <= 12; m++) {
      months.push(yr + '-' + String(m).padStart(2, '0'));
    }
    return months;
  }
  // Academic cycle May -> Apr (+ optional 13th lease-end May when active leases span May-to-May)
  const parts = String(cycleLabel || inferActiveAcademicCycle()).split('-');
  const startYear = parts[0].length === 2 ? 2000 + parseInt(parts[0], 10) : parseInt(parts[0], 10);
  const months = [];
  for (let i = 0; i < 12; i++) {
    const dt = new Date(Date.UTC(startYear, 4 + i, 1));
    months.push(dt.toISOString().slice(0, 7));
  }
  const thirteenth = (startYear + 1) + '-05';
  if (leasesRows && leasesRows.some(function (l) {
    return String(l.Cycle_Label) === String(cycleLabel) && String(l.End_Date || '').slice(0, 7) >= thirteenth;
  })) {
    months.push(thirteenth);
  }
  return months;
}

function getRentRollData(token, params) {
  requireAuth(token, ['Owner', 'Admin']);
  const settings = getSettingsMap();
  return sanitizeForWire(getRentRollDataInternal(params || {}, settings));
}

function leaseOverlapsRange(lease, windowStart, windowEnd) {
  const s = toMonthKey(lease.Start_Date);
  const e = toMonthKey(lease.End_Date);
  if (s && windowEnd && s > windowEnd) return false;
  if (e && windowStart && e < windowStart) return false;
  return Boolean(s || e);
}

function getRentRollDataInternal(params, settings) {
  const cycleType = params.cycleType === 'Calendar' ? 'Calendar' : 'Academic';
  const defaultCycle = cycleType === 'Calendar' ? String(new Date().getFullYear()) : inferActiveAcademicCycle();
  const cycleLabel = String(params.cycleLabel || defaultCycle);

  const leasesAll = getTableData('fact_Leases').rows;
  const tenantsAll = getTableData('dim_Tenants').rows;
  const ledgerAll = getTableData('fact_Rent_Ledger').rows;

  const propScopedLeases = leasesAll.filter(function (l) {
    return !params.propertyId || String(l.Property_ID) === String(params.propertyId);
  });
  const months = getCycleMonths(cycleType, cycleLabel, propScopedLeases);
  const windowStart = months[0];
  const windowEnd = months[months.length - 1];

  const filteredLeases = propScopedLeases.filter(function (l) {
    if (String(l.Lease_Cycle) === cycleType && String(l.Cycle_Label) === cycleLabel) return true;
    return leaseOverlapsRange(l, windowStart, windowEnd);
  });
  const today = todayYmd();
  const currentMonth = today.slice(0, 7);
  const currentDay = parseInt(today.slice(8, 10), 10) || 1;
  const defaultGrace = toNum(settings.DEFAULT_GRACE_DAY, 5);

  const summary = {
    monthlyScheduled: 0,
    cycleScheduledTotal: 0,
    cycleCollectedTotal: 0,
    depositsRequired: 0,
    depositsHeld: 0,
    pastDueBalance: 0,
    pastDueTenants: 0
  };

  const enrichedLeases = filteredLeases.map(function (lease) {
    const graceDay = toNum(lease.Grace_Day, defaultGrace);
    const leaseTenants = tenantsAll.filter(function (t) { return String(t.Lease_ID) === String(lease.Lease_ID); });
    const leaseLedger = ledgerAll.filter(function (r) { return String(r.Lease_ID) === String(lease.Lease_ID); });

    const leaseStartMonth = toMonthKey(lease.Start_Date);
    const leaseEndMonth = toMonthKey(lease.End_Date);
    let leasePaidTotal = 0;
    let leasePastDue = 0;
    let leaseDepositsHeld = 0;

    const computedTenants = leaseTenants.map(function (t) {
      let overrides = {};
      try {
        if (t.Lease_Override_Months_JSON) overrides = JSON.parse(t.Lease_Override_Months_JSON);
      } catch (err) {
        overrides = {};
      }
      const tMoveIn = toMonthKey(t.Move_In_Date);
      const tMoveOut = toMonthKey(t.Move_Out_Date);
      const baseShare = round2(toNum(t.Rent_Share, 0));
      const depHeld = round2(toNum(t.Deposit_Paid, 0));
      const depReq = round2(toNum(t.Deposit_Required, 0));
      leaseDepositsHeld = round2(leaseDepositsHeld + depHeld);
      summary.depositsHeld = round2(summary.depositsHeld + depHeld);
      summary.depositsRequired = round2(summary.depositsRequired + depReq);

      const tLedger = leaseLedger.filter(function (r) { return String(r.Tenant_ID) === String(t.Tenant_ID); });
      const monthState = {};
      let activeCountBeforeLast = 0;
      const thirteenthKey = months.length === 13 ? months[12] : '';

      // 1. Initialize scheduled rent + explicit late fees / waivers
      months.forEach(function (mk) {
        let scheduled = baseShare;
        let isWaived = false;
        let isOutside = false;

        if (tMoveIn && mk < tMoveIn) isOutside = true;
        if (tMoveOut && mk > tMoveOut) isOutside = true;
        if (leaseStartMonth && mk < leaseStartMonth) isOutside = true;
        if (leaseEndMonth && mk > leaseEndMonth) isOutside = true;

        if (Object.prototype.hasOwnProperty.call(overrides, mk)) {
          const ov = overrides[mk];
          if (ov === 'WAIVED' || ov === 'FREE' || Number(ov) === 0) {
            scheduled = 0;
            isWaived = true;
          } else {
            scheduled = round2(toNum(ov, baseShare));
          }
        } else if (isOutside) {
          scheduled = 0;
        } else if (mk === thirteenthKey && activeCountBeforeLast >= 12) {
          scheduled = 0;
          isOutside = true;
        }

        if (mk !== thirteenthKey && !isOutside) activeCountBeforeLast++;

        monthState[mk] = {
          month: mk,
          scheduled: scheduled,
          lateFees: 0,
          credits: 0,
          applied: 0,
          prepaidCashAnchor: 0,
          isPrepaidCovered: false,
          balance: scheduled,
          isWaived: isWaived,
          isOutside: isOutside,
          paidDate: '',
          status: isOutside ? 'Outside' : (isWaived ? 'Waived' : 'Open')
        };
      });

      // 2. Apply Late Fees, Credits, Explicit Covered-Period Prepays & Tagged Rent Payments
      const unallocatedCredits = [];
      tLedger.forEach(function (entry) {
        const amt = round2(toNum(entry.Amount, 0));
        const pm = toMonthKey(entry.Period_Month) || toMonthKey(entry.Entry_Date);
        if (entry.Entry_Type === 'Late Fee' || entry.Entry_Type === 'Utility/Misc') {
          if (monthState[pm]) {
            monthState[pm].lateFees = round2(monthState[pm].lateFees + amt);
            monthState[pm].scheduled = round2(monthState[pm].scheduled + amt);
            monthState[pm].balance = round2(monthState[pm].balance + amt);
          }
        } else if (entry.Entry_Type === 'Credit/Waiver') {
          if (monthState[pm]) {
            monthState[pm].credits = round2(monthState[pm].credits + amt);
            monthState[pm].scheduled = Math.max(0, round2(monthState[pm].scheduled - amt));
          }
        } else if (entry.Entry_Type === 'Upfront Prepay') {
          let coveredList = [];
          try {
            if (entry.Covered_Months_JSON) coveredList = JSON.parse(entry.Covered_Months_JSON);
          } catch (err) {
            coveredList = [];
          }
          if (Array.isArray(coveredList) && coveredList.length > 0) {
            coveredList.forEach(function (cmk) {
              if (monthState[cmk]) {
                monthState[cmk].isPrepaidCovered = true;
                monthState[cmk].applied = monthState[cmk].scheduled;
                monthState[cmk].balance = 0;
                monthState[cmk].paidDate = entry.Entry_Date || '';
                if (cmk === pm) monthState[cmk].prepaidCashAnchor = round2(monthState[cmk].prepaidCashAnchor + amt);
              }
            });
            leasePaidTotal = round2(leasePaidTotal + amt);
            summary.cycleCollectedTotal = round2(summary.cycleCollectedTotal + amt);
          } else {
            unallocatedCredits.push({ amount: amt, entryDate: entry.Entry_Date, startMonth: pm || months[0] });
            leasePaidTotal = round2(leasePaidTotal + amt);
            summary.cycleCollectedTotal = round2(summary.cycleCollectedTotal + amt);
          }
        } else if (entry.Entry_Type === 'Rent') {
          leasePaidTotal = round2(leasePaidTotal + amt);
          summary.cycleCollectedTotal = round2(summary.cycleCollectedTotal + amt);
          if (pm && monthState[pm]) {
            const cell = monthState[pm];
            const needed = Math.max(0, round2(cell.scheduled - cell.applied));
            const useNow = Math.min(needed, amt);
            cell.applied = round2(cell.applied + useNow);
            if (!cell.paidDate) cell.paidDate = entry.Entry_Date || '';
            const leftover = round2(amt - useNow);
            if (leftover > 0.009) {
              unallocatedCredits.push({ amount: leftover, entryDate: entry.Entry_Date, startMonth: pm });
            }
          } else {
            unallocatedCredits.push({ amount: amt, entryDate: entry.Entry_Date, startMonth: months[0] });
          }
        }
      });

      // 3. FIFO Prepayment Rollover across remaining unpaid scheduled months
      unallocatedCredits.forEach(function (cred) {
        let remaining = cred.amount;
        for (let i = 0; i < months.length && remaining > 0.009; i++) {
          const mk = months[i];
          if (cred.startMonth && mk < cred.startMonth) continue;
          const cell = monthState[mk];
          if (cell.isPrepaidCovered || cell.scheduled <= 0) continue;
          const need = Math.max(0, round2(cell.scheduled - cell.applied));
          if (need <= 0.009) continue;
          const take = Math.min(need, remaining);
          cell.applied = round2(cell.applied + take);
          if (!cell.paidDate) cell.paidDate = cred.entryDate || '';
          remaining = round2(remaining - take);
        }
      });

      // 4. Finalize balances & status flags per month
      let tenantScheduledTotal = 0;
      let tenantPastDue = 0;
      months.forEach(function (mk) {
        const cell = monthState[mk];
        if (!cell.isPrepaidCovered) {
          cell.balance = Math.max(0, round2(cell.scheduled - cell.applied));
        }
        tenantScheduledTotal = round2(tenantScheduledTotal + cell.scheduled);
        summary.cycleScheduledTotal = round2(summary.cycleScheduledTotal + cell.scheduled);
        if (mk === currentMonth) {
          summary.monthlyScheduled = round2(summary.monthlyScheduled + cell.scheduled);
        }
        if (cell.isOutside) {
          cell.status = 'Outside';
        } else if (cell.isPrepaidCovered) {
          cell.status = 'Prepaid';
        } else if (cell.scheduled <= 0 && cell.isWaived) {
          cell.status = 'Waived';
        } else if (cell.balance <= 0.009 && cell.scheduled > 0) {
          cell.status = 'Paid';
        } else if (mk < currentMonth || (mk === currentMonth && currentDay > graceDay)) {
          cell.status = cell.applied > 0 ? 'Partial Past Due' : 'Past Due';
          tenantPastDue = round2(tenantPastDue + cell.balance);
        } else if (cell.applied > 0 && cell.balance > 0.009) {
          cell.status = 'Partial';
        } else {
          cell.status = mk === currentMonth ? 'Due' : 'Upcoming';
        }
      });

      if (tenantPastDue > 0.009) {
        summary.pastDueTenants++;
        summary.pastDueBalance = round2(summary.pastDueBalance + tenantPastDue);
        leasePastDue = round2(leasePastDue + tenantPastDue);
      }

      return Object.assign({}, t, {
        months: monthState,
        tenantScheduledTotal: tenantScheduledTotal,
        tenantPastDue: tenantPastDue,
        hasPastDue: tenantPastDue > 0.009
      });
    });

    return Object.assign({}, lease, {
      tenants: computedTenants,
      leasePaidTotal: leasePaidTotal,
      leaseDepositsHeld: leaseDepositsHeld,
      leasePastDue: leasePastDue
    });
  });

  return {
    cycleType: cycleType,
    cycleLabel: cycleLabel,
    months: months.map(function (mk) { return { key: mk, label: formatMonthLabel(mk) }; }),
    currentMonth: currentMonth,
    summary: summary,
    leases: enrichedLeases,
    recentLedger: ledgerAll.slice(-25).reverse()
  };
}

function saveLease(token, payload) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  if (!payload) throw new Error('Lease payload required.');
  return withScriptLock(function () {
    const prop = findPropertyRecord(payload.propertyId, payload.propertyName);
    if (!prop) throw new Error('Valid Property is required.');
    const fields = {
      Property_ID: prop.Property_ID,
      Property_Name: prop.Property_Name,
      Unit_Label: String(payload.unitLabel || ''),
      Lease_Cycle: payload.leaseCycle === 'Calendar' ? 'Calendar' : 'Academic',
      Cycle_Label: String(payload.cycleLabel || inferActiveAcademicCycle()),
      Start_Date: toYmd(payload.startDate),
      End_Date: toYmd(payload.endDate),
      Monthly_Rent_Total: round2(toNum(payload.monthlyRentTotal, 0)),
      Deposit_Total: round2(toNum(payload.depositTotal, 0)),
      Grace_Day: toNum(payload.graceDay, 5),
      Late_Fee_Amount: round2(toNum(payload.lateFeeAmount, 50)),
      Proration_Notes: String(payload.prorationNotes || ''),
      Status: String(payload.status || 'Active'),
      Updated_At: nowIso()
    };

    let leaseRow = null;
    if (payload.leaseId) {
      leaseRow = updateRowById('fact_Leases', 'Lease_ID', payload.leaseId, fields);
      if (!leaseRow) throw new Error('Lease not found.');
    } else {
      const leaseId = newId('LSE');
      leaseRow = Object.assign({ Lease_ID: leaseId }, fields);
      appendRecords('fact_Leases', [leaseRow]);
    }

    if (Array.isArray(payload.tenants) && payload.tenants.length > 0) {
      const newTenants = [];
      payload.tenants.forEach(function (t) {
        if (!t || !t.name) return;
        newTenants.push({
          Tenant_ID: newId('TNT'),
          Lease_ID: leaseRow.Lease_ID,
          Property_ID: prop.Property_ID,
          Property_Name: prop.Property_Name,
          Name: String(t.name).trim(),
          Room_Unit: String(t.roomUnit || ''),
          Email: String(t.email || ''),
          Phone: String(t.phone || ''),
          Rent_Share: round2(toNum(t.rentShare, 0)),
          Deposit_Required: round2(toNum(t.depositRequired, 0)),
          Deposit_Paid: round2(toNum(t.depositPaid, 0)),
          Deposit_Paid_Date: toYmd(t.depositPaidDate || ''),
          Lease_Override_Months_JSON: t.overridesJson || '{}',
          Move_In_Date: toYmd(t.moveInDate || fields.Start_Date),
          Move_Out_Date: toYmd(t.moveOutDate || ''),
          Guarantor_Info: String(t.guarantorInfo || ''),
          Status: 'Active'
        });
      });
      if (newTenants.length > 0) appendRecords('dim_Tenants', newTenants);
    }

    addAuditLog(user.email, payload.leaseId ? 'UPDATE_LEASE' : 'CREATE_LEASE', 'Lease', leaseRow.Lease_ID, prop.Property_Name);
    return sanitizeForWire(leaseRow);
  });
}

function renewLease(token, payload) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  return withScriptLock(function () {
    const oldLease = getTableData('fact_Leases').rows.find(function (l) { return String(l.Lease_ID) === String(payload.leaseId); });
    if (!oldLease) throw new Error('Source lease not found.');

    updateRowById('fact_Leases', 'Lease_ID', oldLease.Lease_ID, { Status: 'Renewed', Updated_At: nowIso() });
    const newLeaseId = newId('LSE');
    const newLease = {
      Lease_ID: newLeaseId,
      Property_ID: oldLease.Property_ID,
      Property_Name: oldLease.Property_Name,
      Unit_Label: oldLease.Unit_Label,
      Lease_Cycle: payload.leaseCycle || oldLease.Lease_Cycle,
      Cycle_Label: String(payload.newCycleLabel || ''),
      Start_Date: toYmd(payload.newStartDate),
      End_Date: toYmd(payload.newEndDate),
      Monthly_Rent_Total: round2(toNum(payload.newMonthlyRent, oldLease.Monthly_Rent_Total)),
      Deposit_Total: round2(toNum(payload.newDepositTotal, oldLease.Deposit_Total)),
      Grace_Day: toNum(oldLease.Grace_Day, 5),
      Late_Fee_Amount: toNum(oldLease.Late_Fee_Amount, 50),
      Proration_Notes: 'Renewed from ' + oldLease.Lease_ID,
      Status: 'Active',
      Updated_At: nowIso()
    };
    appendRecords('fact_Leases', [newLease]);

    if (payload.copyTenants !== false) {
      const oldTenants = getTableData('dim_Tenants').rows.filter(function (t) {
        return String(t.Lease_ID) === oldLease.Lease_ID && t.Status === 'Active';
      });
      const copied = oldTenants.map(function (t) {
        return {
          Tenant_ID: newId('TNT'),
          Lease_ID: newLeaseId,
          Property_ID: oldLease.Property_ID,
          Property_Name: oldLease.Property_Name,
          Name: t.Name,
          Room_Unit: t.Room_Unit,
          Email: t.Email,
          Phone: t.Phone,
          Rent_Share: toNum(t.Rent_Share, 0),
          Deposit_Required: toNum(t.Deposit_Required, 0),
          Deposit_Paid: toNum(t.Deposit_Paid, 0),
          Deposit_Paid_Date: t.Deposit_Paid_Date || '',
          Lease_Override_Months_JSON: '{}',
          Move_In_Date: newLease.Start_Date,
          Move_Out_Date: '',
          Guarantor_Info: t.Guarantor_Info || '',
          Status: 'Active'
        };
      });
      if (copied.length > 0) appendRecords('dim_Tenants', copied);
    }
    addAuditLog(user.email, 'RENEW_LEASE', 'Lease', newLeaseId, oldLease.Property_Name + ' -> ' + newLease.Cycle_Label);
    return sanitizeForWire(newLease);
  });
}

function saveTenant(token, payload) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  if (!payload || !payload.leaseId) throw new Error('leaseId is required.');
  return withScriptLock(function () {
    const lease = getTableData('fact_Leases').rows.find(function (l) { return String(l.Lease_ID) === String(payload.leaseId); });
    if (!lease) throw new Error('Lease not found.');
    const fields = {
      Lease_ID: lease.Lease_ID,
      Property_ID: lease.Property_ID,
      Property_Name: lease.Property_Name,
      Name: String(payload.name || '').trim(),
      Room_Unit: String(payload.roomUnit || ''),
      Email: String(payload.email || ''),
      Phone: String(payload.phone || ''),
      Rent_Share: round2(toNum(payload.rentShare, 0)),
      Deposit_Required: round2(toNum(payload.depositRequired, 0)),
      Deposit_Paid: round2(toNum(payload.depositPaid, 0)),
      Deposit_Paid_Date: toYmd(payload.depositPaidDate || ''),
      Lease_Override_Months_JSON: payload.overridesJson || '{}',
      Move_In_Date: toYmd(payload.moveInDate || lease.Start_Date),
      Move_Out_Date: toYmd(payload.moveOutDate || ''),
      Guarantor_Info: String(payload.guarantorInfo || ''),
      Status: String(payload.status || 'Active')
    };
    if (!fields.Name) throw new Error('Tenant name is required.');
    let saved = null;
    if (payload.tenantId) {
      saved = updateRowById('dim_Tenants', 'Tenant_ID', payload.tenantId, fields);
    } else {
      saved = Object.assign({ Tenant_ID: newId('TNT') }, fields);
      appendRecords('dim_Tenants', [saved]);
    }
    addAuditLog(user.email, payload.tenantId ? 'UPDATE_TENANT' : 'ADD_TENANT', 'Tenant', saved.Tenant_ID, fields.Name);
    return sanitizeForWire(saved);
  });
}

function setTenantMonthOverride(token, payload) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  return withScriptLock(function () {
    const t = getTableData('dim_Tenants').rows.find(function (r) { return String(r.Tenant_ID) === String(payload.tenantId); });
    if (!t) throw new Error('Tenant not found.');
    let overrides = {};
    try { if (t.Lease_Override_Months_JSON) overrides = JSON.parse(t.Lease_Override_Months_JSON); } catch (e) {}
    if (payload.clearOverride) {
      delete overrides[payload.monthKey];
    } else {
      overrides[payload.monthKey] = toNum(payload.overrideAmount, 0);
    }
    const updated = updateRowById('dim_Tenants', 'Tenant_ID', t.Tenant_ID, {
      Lease_Override_Months_JSON: JSON.stringify(overrides)
    });
    addAuditLog(user.email, 'TENANT_MONTH_OVERRIDE', 'Tenant', t.Tenant_ID, payload.monthKey + ' -> ' + (payload.clearOverride ? 'cleared' : payload.overrideAmount));
    return sanitizeForWire(updated);
  });
}

function recordRentLedgerEntry(token, payload) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  if (!payload || !payload.tenantId) throw new Error('tenantId is required.');
  return withScriptLock(function () {
    const t = getTableData('dim_Tenants').rows.find(function (r) { return String(r.Tenant_ID) === String(payload.tenantId); });
    if (!t) throw new Error('Tenant not found.');

    const entryType = ['Rent', 'Upfront Prepay', 'Late Fee', 'Deposit', 'Deposit Refund', 'Credit/Waiver', 'Utility/Misc'].indexOf(payload.entryType) !== -1
      ? payload.entryType
      : 'Rent';
    const amount = round2(toNum(payload.amount, 0));
    if (amount <= 0) throw new Error('Amount must be greater than $0.');

    const coveredMonths = Array.isArray(payload.coveredMonths) ? payload.coveredMonths : [];
    const rec = {
      Ledger_ID: newId('LDG'),
      Lease_ID: t.Lease_ID,
      Tenant_ID: t.Tenant_ID,
      Tenant_Name: t.Name,
      Property_ID: t.Property_ID,
      Property_Name: t.Property_Name,
      Entry_Date: toYmd(payload.entryDate) || todayYmd(),
      Period_Month: toMonthKey(payload.periodMonth || coveredMonths[0] || payload.entryDate || todayYmd()),
      Months_Covered: coveredMonths.length > 0 ? coveredMonths.length : toNum(payload.monthsCovered, 1),
      Covered_Months_JSON: coveredMonths.length > 0 ? JSON.stringify(coveredMonths) : '',
      Entry_Type: entryType,
      Amount: amount,
      Payment_Method: String(payload.paymentMethod || 'Zelle/Venmo'),
      Reference_Note: String(payload.referenceNote || ''),
      Recorded_By: user.email,
      Created_At: nowIso()
    };
    appendRecords('fact_Rent_Ledger', [rec]);

    if (entryType === 'Deposit') {
      updateRowById('dim_Tenants', 'Tenant_ID', t.Tenant_ID, {
        Deposit_Paid: round2(toNum(t.Deposit_Paid, 0) + amount),
        Deposit_Paid_Date: rec.Entry_Date
      });
    }
    addAuditLog(user.email, 'RECORD_RENT_LEDGER', 'RentLedger', rec.Ledger_ID, t.Name + ' ' + entryType + ' $' + amount);
    return sanitizeForWire(rec);
  });
}

function generateDepositSettlement(token, payload) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  if (!payload || !payload.tenantId) throw new Error('tenantId is required.');

  return withScriptLock(function () {
    const t = getTableData('dim_Tenants').rows.find(function (r) { return String(r.Tenant_ID) === String(payload.tenantId); });
    if (!t) throw new Error('Tenant not found.');
    const settings = getSettingsMap();
    const depositPaid = round2(toNum(payload.depositPaidOverride !== undefined && payload.depositPaidOverride !== '' ? payload.depositPaidOverride : t.Deposit_Paid, 0));
    const deductions = Array.isArray(payload.deductions) ? payload.deductions : [];
    let totalDeductions = 0;
    const dedRows = deductions.map(function (d) {
      const amt = round2(toNum(d.amount, 0));
      totalDeductions = round2(totalDeductions + amt);
      return '<tr><td style="padding:8px;border-bottom:1px solid #e2e8f0">' + escapeHtml(d.description || 'Item') +
        '</td><td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right">-$' + amt.toFixed(2) + '</td></tr>';
    }).join('');
    const netRefund = round2(depositPaid - totalDeductions);

    const html =
      '<!DOCTYPE html><html><head><meta charset="utf-8"><style>body{font-family:Helvetica,Arial,sans-serif;color:#1e293b;margin:28px;font-size:13px}</style></head><body>' +
      '<h1 style="margin:0;color:#1e3a8a">' + escapeHtml(settings.COMPANY_NAME) + '</h1>' +
      '<p style="color:#475569">' + escapeHtml(settings.COMPANY_ADDRESS) + ' &bull; ' + escapeHtml(settings.COMPANY_PHONE) + '</p>' +
      '<h2>Security Deposit Settlement Statement</h2>' +
      '<p><strong>Tenant:</strong> ' + escapeHtml(t.Name) + '<br><strong>Property:</strong> ' + escapeHtml(t.Property_Name) +
      '<br><strong>Statement Date:</strong> ' + todayYmd() + '</p>' +
      '<table style="width:100%;border-collapse:collapse;margin-top:14px">' +
      '<tr style="background:#f1f5f9"><th style="padding:8px;text-align:left">Description</th><th style="padding:8px;text-align:right">Amount</th></tr>' +
      '<tr><td style="padding:8px;border-bottom:1px solid #e2e8f0">Security Deposit Held</td><td style="padding:8px;border-bottom:1px solid #e2e8f0;text-align:right">$' + depositPaid.toFixed(2) + '</td></tr>' +
      dedRows +
      '<tr style="font-size:16px;font-weight:bold;background:#eff6ff"><td style="padding:10px">' +
      (netRefund >= 0 ? 'Net Security Deposit Refund Due to Tenant' : 'Remaining Balance Owed by Tenant') +
      '</td><td style="padding:10px;text-align:right">$' + Math.abs(netRefund).toFixed(2) + '</td></tr></table>' +
      (payload.notes ? '<p style="margin-top:16px"><strong>Notes:</strong> ' + escapeHtml(payload.notes) + '</p>' : '') +
      '</body></html>';

    const pdf = renderAndSavePdf('INVOICES_FOLDER_ID', 'Invoices & Settlements', 'Deposit_Settlement_' + t.Tenant_ID + '.pdf', html);

    if (payload.markMovedOut) {
      updateRowById('dim_Tenants', 'Tenant_ID', t.Tenant_ID, {
        Move_Out_Date: toYmd(payload.moveOutDate) || todayYmd(),
        Status: 'Moved Out'
      });
    }
    if (payload.recordRefundLedger && netRefund > 0) {
      appendRecords('fact_Rent_Ledger', [{
        Ledger_ID: newId('LDG'),
        Lease_ID: t.Lease_ID,
        Tenant_ID: t.Tenant_ID,
        Tenant_Name: t.Name,
        Property_ID: t.Property_ID,
        Property_Name: t.Property_Name,
        Entry_Date: todayYmd(),
        Period_Month: todayYmd().slice(0, 7),
        Months_Covered: 0,
        Covered_Months_JSON: '',
        Entry_Type: 'Deposit Refund',
        Amount: netRefund,
        Payment_Method: String(payload.paymentMethod || 'Check'),
        Reference_Note: 'Move-out deposit refund',
        Recorded_By: user.email,
        Created_At: nowIso()
      }]);
    }
    addAuditLog(user.email, 'DEPOSIT_SETTLEMENT', 'Tenant', t.Tenant_ID, 'Net refund $' + netRefund);
    return sanitizeForWire({
      ok: true,
      depositPaid: depositPaid,
      totalDeductions: totalDeductions,
      netRefund: netRefund,
      html: html,
      pdfUrl: pdf.pdfUrl,
      pdfBase64: pdf.pdfBase64,
      pdfFilename: 'Deposit_Settlement_' + t.Name.replace(/\s+/g, '_') + '.pdf'
    });
  });
}

// ============================================================================
// PAYROLL, SUBCONTRACTORS, OUT-OF-POCKET REIMBURSEMENTS & PAYSTUBS
// ============================================================================

function getPayrollSummary(token, params) {
  requireAuth(token, ['Owner', 'Admin']);
  const p = params || {};
  const startDate = toYmd(p.startDate || '');
  const endDate = toYmd(p.endDate || '');

  const emps = getTableData('dim_Employees').rows;
  const logs = getTableData('fact_Work_Logs').rows;
  const exps = getTableData('fact_Material_Expenses').rows;
  const history = getTableData('fact_Payroll_History').rows;

  const workers = emps.map(function (emp) {
    const approvedLogs = [];
    const pendingLogs = [];
    const approvedReimbursements = [];
    const pendingReimbursements = [];

    let approvedHours = 0;
    let approvedRegularPay = 0;
    let approvedReimbPay = 0;
    let pendingGrossPay = 0;

    logs.forEach(function (l) {
      if (String(l.Employee_ID) !== String(emp.Employee_ID)) return;
      if (startDate && l.Work_Date < startDate) return;
      if (endDate && l.Work_Date > endDate) return;
      if (l.Payroll_Status === 'Approved') {
        approvedLogs.push(l);
        approvedHours = round2(approvedHours + toNum(l.Hours, 0));
        approvedRegularPay = round2(approvedRegularPay + toNum(l.Gross_Pay, 0));
      } else if (l.Payroll_Status === 'Pending') {
        pendingLogs.push(l);
        pendingGrossPay = round2(pendingGrossPay + toNum(l.Gross_Pay, 0));
      }
    });

    exps.forEach(function (e) {
      if (String(e.Employee_ID) !== String(emp.Employee_ID)) return;
      if (e.Paid_By !== 'Employee Out-of-Pocket') return;
      if (startDate && e.Expense_Date < startDate) return;
      if (endDate && e.Expense_Date > endDate) return;
      if (e.Payroll_Status === 'Approved') {
        approvedReimbursements.push(e);
        approvedReimbPay = round2(approvedReimbPay + toNum(e.Total_Cost, 0));
      } else if (e.Payroll_Status === 'Pending') {
        pendingReimbursements.push(e);
        pendingGrossPay = round2(pendingGrossPay + toNum(e.Total_Cost, 0));
      }
    });

    return {
      employeeId: emp.Employee_ID,
      name: emp.Name,
      email: emp.Email,
      role: emp.Role,
      payType: emp.Pay_Type || 'Hourly',
      defaultPayRate: toNum(emp.Default_Pay_Rate, 0),
      approvedHours: approvedHours,
      approvedRegularPay: approvedRegularPay,
      approvedReimbursementsPay: approvedReimbPay,
      totalApprovedPayout: round2(approvedRegularPay + approvedReimbPay),
      pendingCount: pendingLogs.length + pendingReimbursements.length,
      pendingAmount: pendingGrossPay,
      approvedLogs: approvedLogs,
      approvedReimbursements: approvedReimbursements,
      pendingLogs: pendingLogs,
      pendingReimbursements: pendingReimbursements
    };
  });

  return sanitizeForWire({
    workers: workers,
    history: history.slice().reverse()
  });
}

/**
 * Pays strictly Approved labor logs and Approved out-of-pocket reimbursements for a single employee.
 * Pending items are NEVER touched or stamped.
 */
function runPayrollBatch(token, payload) {
  const actor = requireAuth(token, ['Owner', 'Admin']);
  if (!payload || !payload.employeeId) throw new Error('employeeId is required.');

  return withScriptLock(function () {
    const emp = findEmployeeById(payload.employeeId);
    if (!emp) throw new Error('Employee not found.');

    const selectedLogIds = Array.isArray(payload.logIds) ? payload.logIds : [];
    const selectedExpIds = Array.isArray(payload.expenseIds) ? payload.expenseIds : [];
    if (selectedLogIds.length === 0 && selectedExpIds.length === 0 && toNum(payload.adjustmentAmount, 0) === 0) {
      throw new Error('Select at least one Approved work log, reimbursement, or adjustment to run payroll.');
    }

    const logsTbl = getTableData('fact_Work_Logs');
    const expTbl = getTableData('fact_Material_Expenses');

    const payableLogs = logsTbl.rows.filter(function (l) {
      return selectedLogIds.indexOf(l.Log_ID) !== -1 &&
        String(l.Employee_ID) === String(emp.Employee_ID) &&
        l.Payroll_Status === 'Approved';
    });
    const payableExps = expTbl.rows.filter(function (e) {
      return selectedExpIds.indexOf(e.Expense_ID) !== -1 &&
        String(e.Employee_ID) === String(emp.Employee_ID) &&
        e.Paid_By === 'Employee Out-of-Pocket' &&
        e.Payroll_Status === 'Approved';
    });

    if (payableLogs.length === 0 && payableExps.length === 0 && toNum(payload.adjustmentAmount, 0) === 0) {
      throw new Error('Only Approved entries can be paid. Pending or already-paid entries were excluded.');
    }

    let totalHours = 0;
    let regularPay = 0;
    payableLogs.forEach(function (l) {
      totalHours = round2(totalHours + toNum(l.Hours, 0));
      regularPay = round2(regularPay + toNum(l.Gross_Pay, 0));
    });

    let reimbursementsPay = 0;
    payableExps.forEach(function (e) {
      reimbursementsPay = round2(reimbursementsPay + toNum(e.Total_Cost, 0));
    });

    const adjAmount = round2(toNum(payload.adjustmentAmount, 0));
    const adjNote = String(payload.adjustmentNote || '');
    const netPayout = round2(regularPay + reimbursementsPay + adjAmount);
    const payrollId = newId('PAY');
    const runDate = todayYmd();
    const settings = getSettingsMap();

    const paystubHtml = buildPaystubHtml({
      payrollId: payrollId,
      runDate: runDate,
      periodStart: toYmd(payload.periodStart || ''),
      periodEnd: toYmd(payload.periodEnd || runDate),
      emp: emp,
      logs: payableLogs,
      expenses: payableExps,
      totalHours: totalHours,
      regularPay: regularPay,
      reimbursementsPay: reimbursementsPay,
      adjAmount: adjAmount,
      adjNote: adjNote,
      netPayout: netPayout,
      paymentMethod: String(payload.paymentMethod || 'Direct Deposit'),
      referenceNumber: String(payload.referenceNumber || '')
    }, settings);

    const pdf = renderAndSavePdf('INVOICES_FOLDER_ID', 'Paystubs', payrollId + '_' + emp.Name.replace(/\s+/g, '_') + '.pdf', paystubHtml);

    const historyRec = {
      Payroll_ID: payrollId,
      Run_Date: runDate,
      Pay_Period_Start: toYmd(payload.periodStart || ''),
      Pay_Period_End: toYmd(payload.periodEnd || runDate),
      Employee_ID: emp.Employee_ID,
      Employee_Name: emp.Name,
      Pay_Type: emp.Pay_Type || 'Hourly',
      Total_Hours: totalHours,
      Regular_Pay: regularPay,
      Reimbursements_Pay: reimbursementsPay,
      Adjustment_Amount: adjAmount,
      Adjustment_Note: adjNote,
      Net_Payout: netPayout,
      Payment_Method: String(payload.paymentMethod || 'Direct Deposit'),
      Reference_Number: String(payload.referenceNumber || ''),
      Paystub_PDF_URL: pdf.pdfUrl,
      Work_Log_IDs_JSON: JSON.stringify(payableLogs.map(function (l) { return l.Log_ID; })),
      Expense_IDs_JSON: JSON.stringify(payableExps.map(function (e) { return e.Expense_ID; })),
      Paid_By: actor.email,
      Status: 'Paid'
    };
    appendRecords('fact_Payroll_History', [historyRec]);

    const logPatches = {};
    payableLogs.forEach(function (l) {
      logPatches[l.Log_ID] = { Payroll_Status: 'Paid', Payroll_ID: payrollId };
    });
    batchPatchRowsById('fact_Work_Logs', 'Log_ID', logPatches);

    const expPatches = {};
    payableExps.forEach(function (e) {
      expPatches[e.Expense_ID] = { Payroll_Status: 'Paid', Payroll_ID: payrollId };
    });
    batchPatchRowsById('fact_Material_Expenses', 'Expense_ID', expPatches);

    addAuditLog(actor.email, 'RUN_PAYROLL', 'Payroll', payrollId, emp.Name + ' Net $' + netPayout);
    return sanitizeForWire({
      ok: true,
      payroll: historyRec,
      html: paystubHtml,
      pdfBase64: pdf.pdfBase64,
      pdfFilename: payrollId + '_' + emp.Name.replace(/\s+/g, '_') + '.pdf'
    });
  });
}

function buildPaystubHtml(data, settings) {
  const logRows = data.logs.map(function (l) {
    return '<tr><td style="padding:6px;border-bottom:1px solid #e2e8f0">' + escapeHtml(l.Work_Date) + '</td>' +
      '<td style="padding:6px;border-bottom:1px solid #e2e8f0">' + escapeHtml(l.Property_Name) + ' — ' + escapeHtml(l.Task_Name) + '</td>' +
      '<td style="padding:6px;border-bottom:1px solid #e2e8f0;text-align:right">' + l.Hours + '</td>' +
      '<td style="padding:6px;border-bottom:1px solid #e2e8f0;text-align:right">$' + round2(l.Gross_Pay).toFixed(2) + '</td></tr>';
  }).join('');

  const expRows = data.expenses.map(function (e) {
    return '<tr><td style="padding:6px;border-bottom:1px solid #e2e8f0">' + escapeHtml(e.Expense_Date) + '</td>' +
      '<td style="padding:6px;border-bottom:1px solid #e2e8f0">Reimbursement: ' + escapeHtml(e.Property_Name) + ' — ' + escapeHtml(e.Item_Description) + '</td>' +
      '<td style="padding:6px;border-bottom:1px solid #e2e8f0;text-align:right">-</td>' +
      '<td style="padding:6px;border-bottom:1px solid #e2e8f0;text-align:right">$' + round2(e.Total_Cost).toFixed(2) + '</td></tr>';
  }).join('');

  return (
    '<!DOCTYPE html><html><head><meta charset="utf-8"><style>body{font-family:Helvetica,Arial,sans-serif;color:#1e293b;margin:28px;font-size:13px}</style></head><body>' +
    '<h1 style="margin:0;color:#1e3a8a">' + escapeHtml(settings.COMPANY_NAME) + '</h1>' +
    '<h2 style="margin:6px 0">Pay Summary / Remittance — ' + escapeHtml(data.payrollId) + '</h2>' +
    '<p><strong>Payee:</strong> ' + escapeHtml(data.emp.Name) + ' (' + escapeHtml(data.emp.Pay_Type || 'Hourly') + ')<br>' +
    '<strong>Pay Date:</strong> ' + escapeHtml(data.runDate) + ' &bull; <strong>Period:</strong> ' + escapeHtml((data.periodStart || 'Start') + ' to ' + data.periodEnd) + '<br>' +
    '<strong>Method:</strong> ' + escapeHtml(data.paymentMethod) + (data.referenceNumber ? ' #' + escapeHtml(data.referenceNumber) : '') + '</p>' +
    '<table style="width:100%;border-collapse:collapse;font-size:12px">' +
    '<thead><tr style="background:#f1f5f9"><th style="padding:6px;text-align:left">Date</th><th style="padding:6px;text-align:left">Work / Reimbursement</th>' +
    '<th style="padding:6px;text-align:right">Hours</th><th style="padding:6px;text-align:right">Pay</th></tr></thead>' +
    '<tbody>' + logRows + expRows + '</tbody></table>' +
    '<div style="width:300px;margin-left:auto;margin-top:16px;background:#f8fafc;padding:12px;border:1px solid #cbd5e1;border-radius:8px">' +
    '<table style="width:100%">' +
    '<tr><td>Regular / Contract Pay (' + data.totalHours + ' hrs)</td><td style="text-align:right">$' + round2(data.regularPay).toFixed(2) + '</td></tr>' +
    '<tr><td>Expense Reimbursements</td><td style="text-align:right">$' + round2(data.reimbursementsPay).toFixed(2) + '</td></tr>' +
    (data.adjAmount !== 0 ? '<tr><td>Adjustment (' + escapeHtml(data.adjNote) + ')</td><td style="text-align:right">$' + round2(data.adjAmount).toFixed(2) + '</td></tr>' : '') +
    '<tr style="font-size:16px;font-weight:bold;border-top:2px solid #94a3b8"><td style="padding-top:6px">Net Payout</td><td style="padding-top:6px;text-align:right">$' + round2(data.netPayout).toFixed(2) + '</td></tr>' +
    '</table></div></body></html>'
  );
}

function getMyPortalData(token) {
  const user = requireAuth(token);
  const wos = getWorkOrders(token, {});
  const workData = getWorkLogsAndExpenses(token, {});
  const payHistory = getTableData('fact_Payroll_History').rows.filter(function (p) {
    return String(p.Employee_ID) === user.employeeId;
  }).reverse();
  return sanitizeForWire({
    workOrders: wos,
    workLogs: workData.workLogs,
    expenses: workData.expenses,
    paystubs: payHistory
  });
}

// ============================================================================
// MASTER DATA CRUD, SETTINGS, BULK IMPORT & AUTOMATION TRIGGERS
// ============================================================================

function saveSettings(token, patch) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  if (!patch || typeof patch !== 'object') throw new Error('Settings patch required.');
  return withScriptLock(function () {
    Object.keys(patch).forEach(function (k) {
      saveSettingValue(k, patch[k], 'Updated in Settings');
    });
    addAuditLog(user.email, 'UPDATE_SETTINGS', 'Settings', '', Object.keys(patch).join(', '));
    return sanitizeForWire(getSettingsMap());
  });
}

function saveOwner(token, payload) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  return withScriptLock(function () {
    const rec = {
      Owner_Name: String(payload.ownerName || '').trim(),
      Company_Name: String(payload.companyName || ''),
      Email: String(payload.email || ''),
      CC_Emails: String(payload.ccEmails || ''),
      Phone: String(payload.phone || ''),
      Billing_Mode: String(payload.billingMode || 'Direct Invoice'),
      Mgmt_Fee_Pct: toNum(payload.mgmtFeePct, 8),
      Default_Rate_Override: payload.defaultRateOverride !== '' && payload.defaultRateOverride !== undefined ? toNum(payload.defaultRateOverride, '') : '',
      Material_Markup_Pct: toNum(payload.materialMarkupPct, 0),
      Payment_Terms_Days: toNum(payload.paymentTermsDays, 15),
      Notes: String(payload.notes || ''),
      Status: String(payload.status || 'Active')
    };
    if (!rec.Owner_Name) throw new Error('Owner name is required.');
    let saved;
    if (payload.ownerId) {
      saved = updateRowById('dim_Owners', 'Owner_ID', payload.ownerId, rec);
    } else {
      saved = Object.assign({ Owner_ID: newId('OWN') }, rec);
      appendRecords('dim_Owners', [saved]);
    }
    addAuditLog(user.email, payload.ownerId ? 'UPDATE_OWNER' : 'CREATE_OWNER', 'Owner', saved.Owner_ID, rec.Owner_Name);
    return sanitizeForWire(saved);
  });
}

function saveProperty(token, payload) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  return withScriptLock(function () {
    const owners = getTableData('dim_Owners').rows;
    const owner = owners.find(function (o) { return String(o.Owner_ID) === String(payload.ownerId); }) || {};
    const rec = {
      Property_Name: String(payload.propertyName || '').trim(),
      Owner_ID: String(payload.ownerId || owner.Owner_ID || ''),
      Owner_Name: String(payload.ownerName || owner.Owner_Name || ''),
      Address: String(payload.address || ''),
      Unit_Type: String(payload.unitType || 'Single Family'),
      Billing_Mode_Override: String(payload.billingModeOverride || ''),
      Mgmt_Fee_Override_Pct: payload.mgmtFeeOverridePct !== '' && payload.mgmtFeeOverridePct !== undefined ? toNum(payload.mgmtFeeOverridePct, '') : '',
      Default_Lease_Cycle: payload.defaultLeaseCycle === 'Calendar' ? 'Calendar' : 'Academic',
      Aliases: String(payload.aliases || ''),
      Gate_Lock_Notes: String(payload.gateLockNotes || ''),
      Status: String(payload.status || 'Active')
    };
    if (!rec.Property_Name) throw new Error('Property name is required.');
    let saved;
    if (payload.propertyId) {
      saved = updateRowById('Properties', 'Property_ID', payload.propertyId, rec);
    } else {
      saved = Object.assign({ Property_ID: newId('PROP') }, rec);
      appendRecords('Properties', [saved]);
    }
    addAuditLog(user.email, payload.propertyId ? 'UPDATE_PROPERTY' : 'CREATE_PROPERTY', 'Property', saved.Property_ID, rec.Property_Name);
    return sanitizeForWire(saved);
  });
}

function saveTask(token, payload) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  return withScriptLock(function () {
    const rec = {
      Task_Name: String(payload.taskName || '').trim(),
      Default_Billable_Rate: round2(toNum(payload.defaultBillableRate, 65)),
      Default_Billing_Unit: payload.defaultBillingUnit === 'Flat' ? 'Flat' : 'Hourly',
      Category: String(payload.category || 'Maintenance'),
      Description: String(payload.description || ''),
      Active: payload.active === false || payload.active === 'FALSE' ? 'FALSE' : 'TRUE'
    };
    if (!rec.Task_Name) throw new Error('Task name is required.');
    let saved;
    if (payload.taskId) {
      saved = updateRowById('Tasks', 'Task_ID', payload.taskId, rec);
    } else {
      saved = Object.assign({ Task_ID: newId('TSK') }, rec);
      appendRecords('Tasks', [saved]);
    }
    addAuditLog(user.email, payload.taskId ? 'UPDATE_TASK' : 'CREATE_TASK', 'Task', saved.Task_ID, rec.Task_Name);
    return sanitizeForWire(saved);
  });
}

function saveEmployee(token, payload) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  return withScriptLock(function () {
    const rec = {
      Name: String(payload.name || '').trim(),
      Email: String(payload.email || '').trim(),
      Phone: String(payload.phone || '').trim(),
      Pay_Type: ['Hourly', 'Flat/1099', 'Salary/Owner'].indexOf(payload.payType) !== -1 ? payload.payType : 'Hourly',
      Default_Pay_Rate: round2(toNum(payload.defaultPayRate, 25)),
      Role: ['Owner', 'Admin', 'Field Crew', 'Contractor'].indexOf(payload.role) !== -1 ? payload.role : 'Field Crew',
      Active: payload.active === false || payload.active === 'FALSE' ? 'FALSE' : 'TRUE',
      Default_Billable_Rate_Override: payload.defaultBillableRateOverride !== '' && payload.defaultBillableRateOverride !== undefined
        ? toNum(payload.defaultBillableRateOverride, '')
        : '',
      Notes: String(payload.notes || '')
    };
    if (!rec.Name || (!rec.Email && !rec.Phone)) throw new Error('Employee Name and either Email or Phone Number are required.');
    let saved;
    if (payload.employeeId) {
      if (payload.newPin && /^\d{4,8}$/.test(String(payload.newPin).trim())) {
        rec.Pin_Hash = sha256Hex(String(payload.employeeId) + ':' + String(payload.newPin).trim());
      }
      saved = updateRowById('dim_Employees', 'Employee_ID', payload.employeeId, rec);
    } else {
      const empId = newId('EMP');
      rec.Pin_Hash = (payload.newPin && /^\d{4,8}$/.test(String(payload.newPin).trim()))
        ? sha256Hex(empId + ':' + String(payload.newPin).trim())
        : '';
      saved = Object.assign({ Employee_ID: empId }, rec);
      appendRecords('dim_Employees', [saved]);
    }
    addAuditLog(user.email, payload.employeeId ? 'UPDATE_EMPLOYEE' : 'CREATE_EMPLOYEE', 'Employee', saved.Employee_ID, rec.Name);
    const wire = Object.assign({}, saved);
    wire.Has_Pin = Boolean(wire.Pin_Hash);
    delete wire.Pin_Hash;
    return sanitizeForWire(wire);
  });
}

function saveBillingRate(token, payload) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  return withScriptLock(function () {
    const rec = {
      Task_ID: String(payload.taskId || ''),
      Task_Name: String(payload.taskName || '').trim(),
      Scope_Type: ['Global', 'Owner', 'Property', 'Employee'].indexOf(payload.scopeType) !== -1 ? payload.scopeType : 'Global',
      Scope_ID: String(payload.scopeId || ''),
      Scope_Name: String(payload.scopeName || ''),
      Rate_Type: payload.rateType === 'Flat' ? 'Flat' : 'Hourly',
      Billable_Rate: round2(toNum(payload.billableRate, 0)),
      Effective_Date: toYmd(payload.effectiveDate) || todayYmd(),
      Notes: String(payload.notes || '')
    };
    if (!rec.Task_Name || rec.Billable_Rate <= 0) throw new Error('Task Name and Billable Rate > $0 are required.');
    let saved;
    if (payload.rateId) {
      saved = updateRowById('dim_Billing_Rates', 'Rate_ID', payload.rateId, rec);
    } else {
      saved = Object.assign({ Rate_ID: newId('RT') }, rec);
      appendRecords('dim_Billing_Rates', [saved]);
    }
    addAuditLog(user.email, payload.rateId ? 'UPDATE_RATE' : 'CREATE_RATE', 'BillingRate', saved.Rate_ID, rec.Scope_Type + ': ' + rec.Task_Name + ' $' + rec.Billable_Rate);
    return sanitizeForWire(saved);
  });
}

function deleteBillingRate(token, rateId) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  return withScriptLock(function () {
    const removed = deleteRowsWhere('dim_Billing_Rates', function (r) { return String(r.Rate_ID) === String(rateId); });
    addAuditLog(user.email, 'DELETE_RATE', 'BillingRate', rateId, 'Removed: ' + removed);
    return { ok: true, removed: removed };
  });
}

function importBulkRecords(token, payload) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  if (!payload || !payload.entityType || !Array.isArray(payload.rows)) {
    throw new Error('entityType and rows[] are required.');
  }
  return withScriptLock(function () {
    const type = payload.entityType;
    const rows = payload.rows;
    let imported = 0;

    if (type === 'Properties') {
      const existing = getTableData('Properties').rows;
      const seen = {};
      existing.forEach(function (p) { seen[String(p.Property_Name || '').toLowerCase()] = true; });
      const batch = [];
      rows.forEach(function (r) {
        const name = String(r.Property_Name || r.name || '').trim();
        if (!name || seen[name.toLowerCase()]) return;
        seen[name.toLowerCase()] = true;
        batch.push({
          Property_ID: newId('PROP'),
          Property_Name: name,
          Owner_ID: String(r.Owner_ID || ''),
          Owner_Name: String(r.Owner_Name || ''),
          Address: String(r.Address || ''),
          Unit_Type: String(r.Unit_Type || 'Single Family'),
          Billing_Mode_Override: '',
          Mgmt_Fee_Override_Pct: '',
          Default_Lease_Cycle: r.Default_Lease_Cycle === 'Calendar' ? 'Calendar' : 'Academic',
          Aliases: String(r.Aliases || ''),
          Gate_Lock_Notes: String(r.Gate_Lock_Notes || ''),
          Status: 'Active'
        });
      });
      appendRecords('Properties', batch);
      imported = batch.length;
    } else if (type === 'Tasks') {
      const existing = getTableData('Tasks').rows;
      const seen = {};
      existing.forEach(function (t) { seen[String(t.Task_Name || '').toLowerCase()] = true; });
      const batch = [];
      rows.forEach(function (r) {
        const name = String(r.Task_Name || r.name || '').trim();
        if (!name || seen[name.toLowerCase()]) return;
        seen[name.toLowerCase()] = true;
        batch.push({
          Task_ID: newId('TSK'),
          Task_Name: name,
          Default_Billable_Rate: round2(toNum(r.Default_Billable_Rate || r.rate, 65)),
          Default_Billing_Unit: r.Default_Billing_Unit === 'Flat' ? 'Flat' : 'Hourly',
          Category: String(r.Category || 'Maintenance'),
          Description: String(r.Description || ''),
          Active: 'TRUE'
        });
      });
      appendRecords('Tasks', batch);
      imported = batch.length;
    } else if (type === 'Owners') {
      const batch = rows.filter(function (r) { return String(r.Owner_Name || '').trim(); }).map(function (r) {
        return {
          Owner_ID: newId('OWN'),
          Owner_Name: String(r.Owner_Name).trim(),
          Company_Name: String(r.Company_Name || ''),
          Email: String(r.Email || ''),
          CC_Emails: String(r.CC_Emails || ''),
          Phone: String(r.Phone || ''),
          Billing_Mode: String(r.Billing_Mode || 'Direct Invoice'),
          Mgmt_Fee_Pct: toNum(r.Mgmt_Fee_Pct, 8),
          Default_Rate_Override: r.Default_Rate_Override || '',
          Material_Markup_Pct: toNum(r.Material_Markup_Pct, 0),
          Payment_Terms_Days: toNum(r.Payment_Terms_Days, 15),
          Notes: String(r.Notes || ''),
          Status: 'Active'
        };
      });
      appendRecords('dim_Owners', batch);
      imported = batch.length;
    } else if (type === 'Employees') {
      const batch = rows.filter(function (r) { return String(r.Name || '').trim() && (String(r.Email || '').trim() || String(r.Phone || '').trim()); }).map(function (r) {
        return {
          Employee_ID: newId('EMP'),
          Name: String(r.Name).trim(),
          Email: String(r.Email || '').trim(),
          Phone: String(r.Phone || '').trim(),
          Pay_Type: String(r.Pay_Type || 'Hourly'),
          Default_Pay_Rate: round2(toNum(r.Default_Pay_Rate, 25)),
          Role: String(r.Role || 'Field Crew'),
          Pin_Hash: '',
          Active: 'TRUE',
          Default_Billable_Rate_Override: r.Default_Billable_Rate_Override || '',
          Notes: String(r.Notes || '')
        };
      });
      appendRecords('dim_Employees', batch);
      imported = batch.length;
    } else {
      throw new Error('Unsupported import entityType: ' + type);
    }

    addAuditLog(user.email, 'BULK_IMPORT', type, '', 'Imported ' + imported + ' records');
    return { ok: true, imported: imported };
  });
}

// ============================================================================
// AUTOMATED DIGEST & ARCHIVAL JOBS
// ============================================================================

function installAutomationTriggers(token) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  const count = installTriggersInternal();
  addAuditLog(user.email, 'INSTALL_TRIGGERS', 'Automation', '', 'Installed ' + count + ' triggers');
  return { ok: true, triggersInstalled: count };
}

function installTriggersInternal() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (tr) {
    const fn = tr.getHandlerFunction();
    if (fn === 'runWeeklyDigestTrigger' || fn === 'archiveOldDataTrigger') {
      ScriptApp.deleteTrigger(tr);
    }
  });
  ScriptApp.newTrigger('runWeeklyDigestTrigger')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();
  ScriptApp.newTrigger('archiveOldDataTrigger')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(2)
    .create();
  return 2;
}

function sendTestWeeklyDigest(token) {
  const user = requireAuth(token, ['Owner', 'Admin']);
  const result = runWeeklyDigestTrigger(user.email);
  return sanitizeForWire(result);
}

/**
 * Monday "Nothing Missed" Digest emailing open requests, unbilled work, pending payroll, overdue invoices & past-due rent.
 */
function runWeeklyDigestTrigger(overrideRecipient) {
  const settings = getSettingsMap();
  const to = overrideRecipient || settings.DISPATCH_EMAIL || 'butlerpropmgmt@gmail.com';
  const today = todayYmd();

  const wos = getTableData('fact_Work_Orders').rows.filter(function (w) {
    return ['Completed', 'Billed', 'Cancelled'].indexOf(w.Status) === -1;
  });
  const logs = getTableData('fact_Work_Logs').rows;
  const exps = getTableData('fact_Material_Expenses').rows;
  const invoices = getTableData('fact_Invoices').rows;

  const unbilledLogs = logs.filter(function (l) { return l.Billing_Status === 'Unbilled' && l.Charge_Target !== 'BPM Internal'; });
  const unbilledExps = exps.filter(function (e) { return e.Billing_Status === 'Unbilled' && e.Charge_Target !== 'BPM Internal'; });
  let unbilledDollars = 0;
  unbilledLogs.forEach(function (l) { unbilledDollars = round2(unbilledDollars + getEffectiveLogBillable(l)); });
  unbilledExps.forEach(function (e) { unbilledDollars = round2(unbilledDollars + toNum(e.Billable_Amount, 0)); });

  const pendingApprovals = logs.filter(function (l) { return l.Payroll_Status === 'Pending'; }).length +
    exps.filter(function (e) { return e.Paid_By === 'Employee Out-of-Pocket' && e.Payroll_Status === 'Pending'; }).length;

  const overdueInvoices = invoices.filter(function (inv) {
    return computeInvoiceDisplayStatus(inv, today) === 'Overdue';
  });
  let overdueDollars = 0;
  overdueInvoices.forEach(function (inv) { overdueDollars = round2(overdueDollars + toNum(inv.Balance_Due, 0)); });

  const rentSnap = getRentRollDataInternal({ cycleType: 'Academic', cycleLabel: inferActiveAcademicCycle(today) }, settings);

  const html =
    '<div style="font-family:Arial,sans-serif;max-width:620px;margin:0 auto;padding:20px;border:1px solid #cbd5e1;border-radius:10px">' +
    '<h2 style="margin-top:0;color:#1e3a8a">' + escapeHtml(settings.COMPANY_NAME) + ' — Weekly "Nothing Missed" Digest</h2>' +
    '<p style="color:#475569">Snapshot as of ' + today + '</p>' +
    '<ul>' +
    '<li><strong>Open Work Orders:</strong> ' + wos.length + ' (' + wos.filter(function (w) { return w.Urgency === 'Emergency'; }).length + ' emergency)</li>' +
    '<li><strong>Unbilled Work Ready to Invoice:</strong> ' + unbilledLogs.length + ' labor + ' + unbilledExps.length + ' expenses = <strong>$' + unbilledDollars.toFixed(2) + '</strong></li>' +
    '<li><strong>Crew Entries Awaiting Approval:</strong> ' + pendingApprovals + '</li>' +
    '<li><strong>Overdue Invoices:</strong> ' + overdueInvoices.length + ' totaling <strong>$' + overdueDollars.toFixed(2) + '</strong></li>' +
    '<li><strong>Rent Past Grace Day:</strong> ' + rentSnap.summary.pastDueTenants + ' tenant(s) totaling <strong>$' + round2(rentSnap.summary.pastDueBalance).toFixed(2) + '</strong></li>' +
    '</ul>' +
    '<p style="color:#64748b;font-size:12px">Log into the BPM Operations Portal to clear unbilled items and send overdue reminders.</p></div>';

  MailApp.sendEmail({
    to: to,
    subject: settings.COMPANY_NAME + ' Weekly Digest (' + today + ') — $' + unbilledDollars.toFixed(2) + ' Unbilled',
    body: 'Open Work Orders: ' + wos.length + '\nUnbilled Ready: $' + unbilledDollars.toFixed(2) +
      '\nPending Approvals: ' + pendingApprovals + '\nOverdue Invoices: ' + overdueInvoices.length + ' ($' + overdueDollars.toFixed(2) + ')' +
      '\nRent Past Grace Day: ' + rentSnap.summary.pastDueTenants + ' ($' + round2(rentSnap.summary.pastDueBalance).toFixed(2) + ')',
    htmlBody: html
  });
  return { ok: true, sentTo: to, unbilledDollars: unbilledDollars, overdueDollars: overdueDollars };
}

/**
 * Safely archives closed Work Logs, Expenses & Work Orders older than 90 days using explicit header mapping.
 * Never archives Unbilled or Pending/Unpaid rows!
 */
function archiveOldDataTrigger() {
  return withScriptLock(function () {
    const cutoffYmd = addDaysYmd(todayYmd(), -90);
    let totalArchived = 0;

    function archiveTable(sourceName, archiveName, isClosedFn) {
      const tbl = getTableData(sourceName);
      const toArchive = [];
      const rowIndicesToDelete = [];
      tbl.rows.forEach(function (r) {
        if (isClosedFn(r)) {
          toArchive.push(r);
          rowIndicesToDelete.push(r._rowIndex);
        }
      });
      if (toArchive.length === 0) return 0;
      appendRecords(archiveName, toArchive);
      for (let i = rowIndicesToDelete.length - 1; i >= 0; i--) {
        tbl.sheet.deleteRow(rowIndicesToDelete[i]);
      }
      return toArchive.length;
    }

    totalArchived += archiveTable('fact_Work_Logs', 'archive_fact_Work_Logs', function (r) {
      return r.Work_Date && r.Work_Date < cutoffYmd &&
        (r.Payroll_Status === 'Paid' || r.Payroll_Status === 'N/A') &&
        (r.Billing_Status === 'Paid' || r.Billing_Status === 'Internal/Non-Billable');
    });

    totalArchived += archiveTable('fact_Material_Expenses', 'archive_fact_Material_Expenses', function (r) {
      return r.Expense_Date && r.Expense_Date < cutoffYmd &&
        (r.Payroll_Status === 'Paid' || r.Payroll_Status === 'N/A') &&
        (r.Billing_Status === 'Paid' || r.Billing_Status === 'Internal/Non-Billable');
    });

    totalArchived += archiveTable('fact_Work_Orders', 'archive_fact_Work_Orders', function (r) {
      return r.Completed_At && toYmd(r.Completed_At) < cutoffYmd &&
        (r.Status === 'Billed' || r.Status === 'Cancelled');
    });

    addAuditLog('system', 'ARCHIVE_OLD_DATA', 'Archive', '', 'Archived ' + totalArchived + ' closed records prior to ' + cutoffYmd);
    return { ok: true, archived: totalArchived };
  });
}
