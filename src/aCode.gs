/**
 * Financial System
 * PIN-Based Authentication System
 */

// Configuration
const CONFIG = {
  SPREADSHEET_NAME: "Financial System",
  SESSION_TIMEOUT: 30 * 60 * 1000, // 30 minutes in milliseconds
  SHEETS: {
    // PART 1: Core Transaction Database (3 sheets)
    DB_JOURNAL: "DB_JOURNAL",
    DB_BANK: "DB_BANK",
    DB_BUDGET: "DB_BUDGET",

    // PART 2: Unified Contacts - Suppliers, Customers, Staff (1 sheet)
    CONTACTS: "CONTACTS",

    // PART 3: Procurement (2 sheets)
    PURCHASE_ORDERS: "PURCHASE_ORDERS",
    GRN: "GRN",

    // PART 4: Unified Invoices - Payables & Receivables (1 sheet)
    INVOICES: "INVOICES",

    // PART 5: System (3 sheets)
    MASTER_DATA: "MASTER_DATA",
    SYS_USERS: "SYS_USERS",
    SYS_LOGS: "SYS_LOGS"
  }
  // TOTAL: 10 sheets
};

// Sheet header definitions - single source of truth
const SHEET_HEADERS = {
  DB_JOURNAL: ['UUID', 'Batch_ID', 'Date', 'Financial_Year', 'Account_Code', 'Payee', 'Ref_No',
               'Bank_Ref', 'Particulars', 'Sub_Category', 'Category', 'Description', 'Debit', 'Credit',
               'Account_Type', 'Report_Mapping', 'Recon_Status', 'Receipt_URL'],
  DB_BANK: ['Account_Code', 'Financial_Year', 'Txn_Date', 'Value_Date', 'Bank_Ref',
            'Description', 'Debit', 'Credit', 'Balance', 'Match_Status'],
  DB_BUDGET: ['Date', 'Financial_Year', 'Particulars', 'Sub_Category', 'Category', 'Account_Type',
              'Original_Budget', 'Reallocation', 'Supplementary', 'Final_Budget',
              'Actual_Amount', 'Variance', 'Auth_Ref', 'Description'],
  CONTACTS: ['Contact_ID', 'Contact_Type', 'Contact_Name', 'Contact_Person', 'Phone', 'Email',
             'Address', 'Tax_PIN', 'Bank_Name', 'Bank_Account', 'Category',
             'Credit_Limit', 'Payment_Terms', 'Status', 'Created_Date', 'Created_By'],
  PURCHASE_ORDERS: ['PO_ID', 'PO_Number', 'PO_Date', 'Financial_Year', 'Supplier_ID',
                    'Supplier_Name', 'Description', 'Total_Amount', 'Status',
                    'Requested_By', 'Requested_Date', 'Approved_By', 'Approved_Date',
                    'Delivery_Date', 'Notes', 'Line_Items'],
  GRN: ['GRN_ID', 'GRN_Number', 'GRN_Date', 'PO_ID', 'PO_Number',
        'Supplier_ID', 'Supplier_Name', 'Received_By', 'Status',
        'Invoice_Number', 'Invoice_Date', 'Invoice_Amount', 'Notes', 'Line_Items'],
  INVOICES: ['Invoice_ID', 'Invoice_Type', 'Invoice_Number', 'Invoice_Date', 'Due_Date', 'Financial_Year',
             'Contact_ID', 'Contact_Name', 'GRN_ID', 'GRN_Number', 'PO_ID', 'PO_Number',
             'Description', 'Amount', 'Paid_Amount', 'Balance', 'Status', 'Payment_Terms',
             'Created_Date', 'Created_By', 'Line_Items'],

  MASTER_DATA: ['Particulars', 'Sub_Category', 'Category', 'Account_Codes', 'Account_Type', 'Report_Mapping', 'Financial_Year'],
  SYS_USERS: ['Email', 'PIN', 'Name', 'Role', 'Status'],
  SYS_LOGS: ['Timestamp', 'User', 'Action', 'Target_ID', 'Details']
};

// Legacy sheets to remove (consolidated into CONTACTS and INVOICES)
const LEGACY_SHEETS = ['HOME', 'VIEW_LEDGER', 'VIEW_REPORTS', 'VIEW_RECON', 'PO_LINES', 'GRN_LINES',
                       'SUPPLIERS', 'CUSTOMERS', 'PAYABLES', 'RECEIVABLES'];

/**
 * Main entry point - serves the web app
 */
function doGet(e) {
  const userProperties = PropertiesService.getUserProperties();
  const sessionData = userProperties.getProperty('sessionData');

  if (sessionData) {
    const session = JSON.parse(sessionData);
    const now = new Date().getTime();

    // Check if session is still valid
    if (now - session.lastActivity < CONFIG.SESSION_TIMEOUT) {
      // Update last activity
      session.lastActivity = now;
      userProperties.setProperty('sessionData', JSON.stringify(session));

      // Return dashboard
      return HtmlService.createTemplateFromFile('cDashboard')
        .evaluate()
        .setTitle('Financial System - Dashboard')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } else {
      // Session expired
      userProperties.deleteProperty('sessionData');
    }
  }

  // Show login page
  return HtmlService.createTemplateFromFile('bLogin')
    .evaluate()
    .setTitle('Financial System - Login')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * PUBLIC: Initialize/Sync spreadsheet and sheets
 * Smart sync function that:
 * - Creates missing sheets
 * - Removes legacy/unused sheets
 * - Updates headers on every run
 */
function initializeSpreadsheet() {
  let ss = _getOrCreateSpreadsheet();
  const results = {
    created: [],
    updated: [],
    removed: []
  };

  // Step 1: Remove legacy sheets
  LEGACY_SHEETS.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.deleteSheet(sheet);
      results.removed.push(name);
      Logger.log('Removed legacy sheet: ' + name);
    }
  });

  // Step 2: Create/Update all required sheets
  Object.keys(CONFIG.SHEETS).forEach(key => {
    const sheetName = CONFIG.SHEETS[key];
    const headers = SHEET_HEADERS[key];
    if (!headers) return; // Skip if no headers defined

    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      // Create new sheet
      sheet = ss.insertSheet(sheetName);
      results.created.push(sheetName);
      Logger.log('Created sheet: ' + sheetName);
    } else {
      results.updated.push(sheetName);
    }

    // Always update headers (this ensures new columns are added)
    _syncSheetHeaders(sheet, headers);
  });

  // Step 3: Handle special cases (default data)
  _ensureDefaultUser(ss);

  const message = `Sync complete: ${results.created.length} created, ${results.updated.length} updated, ${results.removed.length} removed`;
  Logger.log(message);

  return {
    success: true,
    message: message,
    created: results.created,
    updated: results.updated,
    removed: results.removed,
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl()
  };
}

/**
 * PRIVATE: Sync sheet headers - updates headers to match definition
 */
function _syncSheetHeaders(sheet, headers) {
  const numCols = headers.length;

  // Set headers in row 1
  sheet.getRange(1, 1, 1, numCols).setValues([headers]);

  // Format header row
  _formatHeaderRow(sheet, numCols);
}

/**
 * PRIVATE: Ensure default admin user exists
 */
function _ensureDefaultUser(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SYS_USERS);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    // Add default admin user
    const defaultUser = [
      'cabdisirlam@gmail.com',
      '1234',
      'Admin User',
      'ADMIN',
      'Active'
    ];
    sheet.getRange(2, 1, 1, defaultUser.length).setValues([defaultUser]);
    Logger.log('Added default admin user');
  }
}

/**
 * PUBLIC: Alias for initializeSpreadsheet - for backwards compatibility
 * Use initializeSpreadsheet() instead
 */
function cleanupSheetHeaders() {
  return initializeSpreadsheet();
}

/**
 * PRIVATE: Get or create the main spreadsheet
 */
function _getOrCreateSpreadsheet() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let spreadsheetId = scriptProperties.getProperty('spreadsheetId');
  let ss;

  if (spreadsheetId) {
    try {
      ss = SpreadsheetApp.openById(spreadsheetId);
      return ss;
    } catch (e) {
      // Spreadsheet not found, create new one
      spreadsheetId = null;
    }
  }

  // Create new spreadsheet
  ss = SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);
  scriptProperties.setProperty('spreadsheetId', ss.getId());

  // Delete the default "Sheet1"
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) {
    ss.deleteSheet(defaultSheet);
  }

  return ss;
}

/**
 * PRIVATE: Helper function to format header row
 */
function _formatHeaderRow(sheet, numColumns) {
  sheet.getRange(1, 1, 1, numColumns)
    .setBackground('#4A90E2')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Auto-resize all columns
  for (let i = 1; i <= numColumns; i++) {
    sheet.autoResizeColumn(i);
  }

  // Freeze header row
  sheet.setFrozenRows(1);
}

/**
 * PRIVATE: Get or ensure a sheet exists (lightweight helper)
 */
function _ensureSheet(ss, sheetKey) {
  const sheetName = CONFIG.SHEETS[sheetKey];
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    const headers = SHEET_HEADERS[sheetKey];
    if (headers) {
      _syncSheetHeaders(sheet, headers);
    }
  }
  return sheet;
}

/**
 * Helper function to log system events
 */
function logSystemEvent(user, action, targetId, details) {
  try {
    const ss = _getOrCreateSpreadsheet();
    const logsSheet = _ensureSheet(ss, 'SYS_LOGS');

    const timestamp = new Date();
    const logEntry = [timestamp, user, action, targetId || '', details || ''];

    logsSheet.appendRow(logEntry);
  } catch (error) {
    Logger.log('Failed to log event: ' + error.toString());
  }
}

/**
 * Authenticate user with email and PIN
 */
function authenticateUser(email, pin) {
  try {
    // Initialize spreadsheet if needed
    const ss = _getOrCreateSpreadsheet();
    const usersSheet = _ensureSheet(ss, 'SYS_USERS');

    // Get all user data
    const data = usersSheet.getDataRange().getValues();

    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const userEmail = row[0];
      const userPin = row[1].toString();
      const userName = row[2];
      const userRole = row[3];
      const userStatus = row[4];

      // Check if email and PIN match
      if (userEmail.toLowerCase() === email.toLowerCase() && userPin === pin) {
        // Check if user is active
        if (userStatus.toUpperCase() !== 'ACTIVE') {
          // Log failed login attempt
          logSystemEvent(email, 'LOGIN_FAILED', '', 'Account inactive');

          return {
            success: false,
            message: 'Your account is inactive. Please contact the administrator.'
          };
        }

        // Create session
        const sessionData = {
          email: userEmail,
          name: userName,
          role: userRole,
          lastActivity: new Date().getTime()
        };

        PropertiesService.getUserProperties()
          .setProperty('sessionData', JSON.stringify(sessionData));

        // Log successful login
        logSystemEvent(userEmail, 'LOGIN_SUCCESS', '', 'User logged in successfully');

        return {
          success: true,
          user: {
            email: userEmail,
            name: userName,
            role: userRole
          }
        };
      }
    }

    // Log failed login attempt
    logSystemEvent(email, 'LOGIN_FAILED', '', 'Invalid credentials');

    return {
      success: false,
      message: 'Invalid email or PIN. Please try again.'
    };

  } catch (error) {
    Logger.log('Authentication error: ' + error.toString());
    logSystemEvent(email, 'LOGIN_ERROR', '', error.toString());

    return {
      success: false,
      message: 'An error occurred during authentication. Please try again.'
    };
  }
}

/**
 * Get current user session
 */
function getCurrentUser() {
  const userProperties = PropertiesService.getUserProperties();
  const sessionData = userProperties.getProperty('sessionData');

  if (!sessionData) {
    return { authenticated: false };
  }

  const session = JSON.parse(sessionData);
  const now = new Date().getTime();

  // Check if session is still valid
  if (now - session.lastActivity < CONFIG.SESSION_TIMEOUT) {
    // Update last activity
    session.lastActivity = now;
    userProperties.setProperty('sessionData', JSON.stringify(session));

    return {
      authenticated: true,
      email: session.email,
      name: session.name,
      role: session.role
    };
  } else {
    // Session expired
    logSystemEvent(session.email, 'SESSION_EXPIRED', '', 'Session timed out after 30 minutes of inactivity');
    userProperties.deleteProperty('sessionData');
    return {
      authenticated: false,
      expired: true
    };
  }
}

/**
 * Logout user
 */
function logout() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const sessionData = userProperties.getProperty('sessionData');

    if (sessionData) {
      const session = JSON.parse(sessionData);
      // Log logout event
      logSystemEvent(session.email, 'LOGOUT', '', 'User logged out');
    }

    userProperties.deleteProperty('sessionData');
    return { success: true };
  } catch (error) {
    Logger.log('Logout error: ' + error.toString());
    return { success: true }; // Still return success to clear session
  }
}

/**
 * Check session validity
 */
function checkSession() {
  return getCurrentUser();
}

/**
 * Admin: Users management
 */
function getUsers() {
  _requireAdmin();
  const ss = _getOrCreateSpreadsheet();
  const sheet = _ensureSheet(ss, 'SYS_USERS');
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) return [];

  return data.slice(1).map((row, index) => {
    return {
      rowId: index + 2,
      email: row[0],
      name: row[2],
      role: row[3],
      status: row[4]
    };
  });
}

function getUserByRowId(rowId) {
  _requireAdmin();
  const ss = _getOrCreateSpreadsheet();
  const sheet = _ensureSheet(ss, 'SYS_USERS');
  const row = Number(rowId);

  if (!row || row < 2) return null;

  const values = sheet.getRange(row, 1, 1, 5).getValues()[0];
  if (!values[0]) return null;

  return {
    rowId: row,
    email: values[0],
    name: values[2],
    role: values[3],
    status: values[4]
  };
}

function addUser(user) {
  _requireAdmin();
  const payload = _normalizeUserPayload(user);
  const ss = _getOrCreateSpreadsheet();
  const sheet = _ensureSheet(ss, 'SYS_USERS');

  const data = sheet.getDataRange().getValues();
  const emailLower = payload.email.toLowerCase();
  const existing = data.slice(1).some(row => String(row[0]).toLowerCase() === emailLower);
  if (existing) {
    throw new Error('Email already exists.');
  }

  sheet.appendRow([payload.email, payload.pin, payload.name, payload.role, payload.status]);
  return { success: true };
}

function updateUser(user) {
  _requireAdmin();
  if (!user) throw new Error('Missing user payload.');
  const row = Number(user.rowId);
  if (!row || row < 2) throw new Error('Invalid user row.');

  const ss = _getOrCreateSpreadsheet();
  const sheet = _ensureSheet(ss, 'SYS_USERS');
  const existingRow = sheet.getRange(row, 1, 1, 5).getValues()[0];
  const existingPin = String(existingRow[1] || '').trim();
  const providedPin = String(user.pin || '').trim();
  const payload = _normalizeUserPayload({
    rowId: row,
    email: user.email,
    pin: providedPin || existingPin,
    name: user.name,
    role: user.role,
    status: user.status
  });

  const data = sheet.getDataRange().getValues();
  const emailLower = payload.email.toLowerCase();
  const duplicate = data.slice(1).some((rowData, idx) => {
    const rowIndex = idx + 2;
    return rowIndex !== row && String(rowData[0]).toLowerCase() === emailLower;
  });
  if (duplicate) {
    throw new Error('Email already exists.');
  }

  sheet.getRange(row, 1, 1, 5).setValues([[
    payload.email,
    payload.pin,
    payload.name,
    payload.role,
    payload.status
  ]]);

  return { success: true };
}

function deactivateUser(rowId) {
  _requireAdmin();
  const ss = _getOrCreateSpreadsheet();
  const sheet = _ensureSheet(ss, 'SYS_USERS');
  const row = Number(rowId);

  if (!row || row < 2) {
    throw new Error('Invalid user row.');
  }

  sheet.getRange(row, 5).setValue('Inactive');
  return { success: true };
}

function _normalizeUserPayload(user) {
  if (!user) throw new Error('Missing user payload.');

  const email = String(user.email || '').trim();
  const pin = String(user.pin || '').trim();
  const name = String(user.name || '').trim();
  const role = String(user.role || '').trim().toUpperCase();
  const status = String(user.status || '').trim();

  if (!email) throw new Error('Email is required.');
  if (!/^\d{4}$/.test(pin)) throw new Error('PIN must be exactly 4 digits.');
  if (!name) throw new Error('Name is required.');
  if (!role) throw new Error('Role is required.');

  const normalizedStatus = status.toLowerCase() === 'inactive' ? 'Inactive' : 'Active';

  return {
    rowId: user.rowId,
    email: email,
    pin: pin,
    name: name,
    role: role,
    status: normalizedStatus
  };
}

function _requireAdmin() {
  const user = getCurrentUser();
  if (!user.authenticated || String(user.role).toUpperCase() !== 'ADMIN') {
    throw new Error('Unauthorized');
  }
}

/**
 * Admin: Settings
 */
function getSystemSettings() {
  _requireAdmin();
  const props = PropertiesService.getScriptProperties();
  return {
    systemName: props.getProperty('systemName') || 'Financial System',
    entityName: props.getProperty('entityName') || 'Main Entity',
    currency: props.getProperty('currency') || 'KSH',
    decimals: props.getProperty('decimals') || '0.00',
    bankName: props.getProperty('bankName') || '',
    bankBranch: props.getProperty('bankBranch') || '',
    bankAccountNumber: props.getProperty('bankAccountNumber') || ''
  };
}

function saveSystemSettings(settings) {
  _requireAdmin();
  if (!settings) throw new Error('Missing settings.');

  const systemName = String(settings.systemName || '').trim();
  const entityName = String(settings.entityName || '').trim();
  const currency = String(settings.currency || '').trim().toUpperCase();
  const decimals = String(settings.decimals || '').trim();
  const bankName = String(settings.bankName || '').trim();
  const bankBranch = String(settings.bankBranch || '').trim();
  const bankAccountNumber = String(settings.bankAccountNumber || '').trim();

  if (!systemName) throw new Error('System name is required.');
  if (!entityName) throw new Error('Entity name is required.');
  if (!currency) throw new Error('Currency is required.');
  if (!/^\d\.\d{2}$/.test(decimals)) throw new Error('Decimals must be in 0.00 format.');

  const props = PropertiesService.getScriptProperties();
  props.setProperty('systemName', systemName);
  props.setProperty('entityName', entityName);
  props.setProperty('currency', currency);
  props.setProperty('decimals', decimals);
  props.setProperty('bankName', bankName);
  props.setProperty('bankBranch', bankBranch);
  props.setProperty('bankAccountNumber', bankAccountNumber);

  return { success: true };
}

/**
 * MASTER DATA HELPER FUNCTIONS
 */

/**
 * Get all unique categories from MASTER_DATA
 * @returns {Array} Array of unique category names
 */
function getCategories() {
  try {
    const ss = _getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);

    if (!sheet) {
      Logger.log('MASTER_DATA sheet not found');
      return [];
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol < 1) return []; // No data

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(value => String(value || '').trim().toLowerCase().replace(/\s+/g, '_'));
    const categoryIndex = headers.indexOf('category');
    if (categoryIndex < 0) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const uniqueCategories = [...new Set(
      data.map(row => String(row[categoryIndex] || '').trim()).filter(cat => cat !== '')
    )];

    return uniqueCategories.sort();
  } catch (error) {
    Logger.log('Error in getCategories: ' + error.toString());
    return [];
  }
}

/**
 * Get sub-categories for a specific category
 * @param {string} category - The category to filter by
 * @returns {Array} Array of sub-categories for the given category
 */
function getSubCategories(category) {
  try {
    if (!category) return [];

    const ss = _getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);

    if (!sheet) {
      Logger.log('MASTER_DATA sheet not found');
      return [];
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol < 1) return []; // No data

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(value => String(value || '').trim().toLowerCase().replace(/\s+/g, '_'));
    const subIndex = headers.indexOf('sub_category');
    const categoryIndex = headers.indexOf('category');
    if (subIndex < 0 || categoryIndex < 0) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    const subCategories = data
      .filter(row => String(row[categoryIndex] || '').trim() === category && String(row[subIndex] || '').trim() !== '')
      .map(row => String(row[subIndex] || '').trim());

    return [...new Set(subCategories)].sort();
  } catch (error) {
    Logger.log('Error in getSubCategories: ' + error.toString());
    return [];
  }
}

/**
 * Get the category for a given sub-category
 * @param {string} subCategory - The sub-category to look up
 * @returns {string} The parent category name
 */
function getCategoryForSubCategory(subCategory) {
  try {
    if (!subCategory) return '';

    const ss = _getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);

    if (!sheet) {
      Logger.log('MASTER_DATA sheet not found');
      return '';
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol < 1) return ''; // No data

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(value => String(value || '').trim().toLowerCase().replace(/\s+/g, '_'));
    const subIndex = headers.indexOf('sub_category');
    const categoryIndex = headers.indexOf('category');
    if (subIndex < 0 || categoryIndex < 0) return '';

    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // Find the first matching row
    const matchingRow = data.find(row => String(row[subIndex] || '').trim() === subCategory);

    return matchingRow ? String(matchingRow[categoryIndex] || '').trim() : '';
  } catch (error) {
    Logger.log('Error in getCategoryForSubCategory: ' + error.toString());
    return '';
  }
}

/**
 * Get account type for a category
 * @param {string} category - The category to look up
 * @returns {string} The account type (Operating Expense, Operating Income, etc.)
 */
function getAccountType(category) {
  try {
    if (!category) return '';

    const ss = _getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);

    if (!sheet) {
      Logger.log('MASTER_DATA sheet not found');
      return '';
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol < 1) return ''; // No data

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(value => String(value || '').trim().toLowerCase().replace(/\s+/g, '_'));
    const categoryIndex = headers.indexOf('category');
    const accountTypeIndex = headers.indexOf('account_type');
    if (categoryIndex < 0 || accountTypeIndex < 0) return '';

    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // Find the first matching row
    const matchingRow = data.find(row => String(row[categoryIndex] || '').trim() === category);

    return matchingRow ? String(matchingRow[accountTypeIndex] || '').trim() : '';
  } catch (error) {
    Logger.log('Error in getAccountType: ' + error.toString());
    return '';
  }
}

/**
 * Get sub-category, category, and account type for a particulars entry.
 * @param {string} particulars - The particulars to look up
 * @returns {Object} { subCategory, category, accountType, reportMapping }
 */
function getDetailsForParticulars(particulars) {
  try {
    if (!particulars) return { subCategory: '', category: '', accountType: '', reportMapping: '' };

    const ss = _getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);

    if (!sheet) {
      Logger.log('MASTER_DATA sheet not found');
      return { subCategory: '', category: '', accountType: '', reportMapping: '' };
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol < 1) return { subCategory: '', category: '', accountType: '', reportMapping: '' };

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(value => String(value || '').trim().toLowerCase().replace(/\s+/g, '_'));
    const particularsIndex = headers.indexOf('particulars');
    const subIndex = headers.indexOf('sub_category');
    const categoryIndex = headers.indexOf('category');
    const accountTypeIndex = headers.indexOf('account_type');
    const reportIndex = headers.indexOf('report_mapping');
    if (particularsIndex < 0 || subIndex < 0 || categoryIndex < 0) {
      return { subCategory: '', category: '', accountType: '', reportMapping: '' };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const matchingRow = data.find(row => String(row[particularsIndex] || '').trim() === particulars);
    if (!matchingRow) {
      return { subCategory: '', category: '', accountType: '', reportMapping: '' };
    }

    return {
      subCategory: String(matchingRow[subIndex] || '').trim(),
      category: String(matchingRow[categoryIndex] || '').trim(),
      accountType: accountTypeIndex >= 0 ? String(matchingRow[accountTypeIndex] || '').trim() : '',
      reportMapping: reportIndex >= 0 ? String(matchingRow[reportIndex] || '').trim() : ''
    };
  } catch (error) {
    Logger.log('Error in getDetailsForParticulars: ' + error.toString());
    return { subCategory: '', category: '', accountType: '', reportMapping: '' };
  }
}

/**
 * Get all unique payees from MASTER_DATA
 * @returns {Array} Array of unique payee names
 */
function getPayees() {
  try {
    const ss = _getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);

    if (!sheet) {
      Logger.log('MASTER_DATA sheet not found');
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; // No data

    const payeeData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const uniquePayees = [...new Set(payeeData.map(row => row[0]).filter(payee => payee !== ''))];

    return uniquePayees.sort();
  } catch (error) {
    Logger.log('Error in getPayees: ' + error.toString());
    return [];
  }
}

function logSystemEventSafe(action, targetId, details) {
  try {
    const user = getCurrentUser();
    const actor = user && user.authenticated ? user.email : 'system';
    logSystemEvent(actor, action, targetId || '', details || '');
  } catch (error) {
    Logger.log('Log failure: ' + error.toString());
  }
}

/**
 * Get all unique sub-categories with category + account type for budget input.
 * @returns {Array} Array of { subCategory, category, accountType }
 */
function getBudgetCategoryCatalog() {
  try {
    const ss = _getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);

    if (!sheet) {
      Logger.log('MASTER_DATA sheet not found');
      return [];
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol < 1) return [];

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(value => String(value || '').trim().toLowerCase().replace(/\s+/g, '_'));
    const categoryIndex = headers.indexOf('category');
    const accountTypeIndex = headers.indexOf('account_type');
    if (categoryIndex < 0) {
      Logger.log('MASTER_DATA headers missing Category');
      return [];
    }

    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const catalogMap = new Map();

    data.forEach(row => {
      const category = String(row[categoryIndex] || '').trim();
      const accountType = accountTypeIndex >= 0 ? String(row[accountTypeIndex] || '').trim() : '';
      if (!category) return;
      if (!catalogMap.has(category)) {
        catalogMap.set(category, {
          category: category,
          accountType: accountType
        });
      }
    });

    const accountOrder = {
      Income: 1,
      Expense: 2,
      Asset: 3,
      'Capital Purchase': 3,
      Capital: 3,
      Liabilities: 4,
      Liability: 4,
      Equity: 5
    };

    const getOrder = (value) => {
      const key = String(value || '').trim();
      if (!key) return 99;
      if (accountOrder[key] != null) return accountOrder[key];
      const normalized = key.toLowerCase();
      if (normalized.includes('income')) return 1;
      if (normalized.includes('expense')) return 2;
      if (normalized.includes('asset') || normalized.includes('capital')) return 3;
      if (normalized.includes('liabil')) return 4;
      if (normalized.includes('equity')) return 5;
      return 99;
    };

    return Array.from(catalogMap.values()).sort((a, b) => {
      const orderA = getOrder(a.accountType);
      const orderB = getOrder(b.accountType);
      if (orderA !== orderB) return orderA - orderB;
      return a.category.localeCompare(b.category);
    });
  } catch (error) {
    Logger.log('Error in getBudgetCategoryCatalog: ' + error.toString());
    return [];
  }
}

// Keep old function for backward compatibility
function getBudgetSubCategoryCatalog() {
  return getBudgetCategoryCatalog();
}

/**
 * Get current budget amounts (Original + Reallocation + Supplementary) for particulars.
 * @param {string} financialYear
 * @param {Array} particularsList
 * @returns {Object} { year, amounts }
 */
function getBudgetCurrentAmounts(financialYear, particularsList) {
  try {
    const year = String(financialYear || '').trim();
    const targets = Array.isArray(particularsList) ? particularsList : [];
    const normalizedTargets = targets.map(item => String(item || '').trim()).filter(Boolean);
    if (!year || !normalizedTargets.length) return { year: year, amounts: {} };

    const ss = _getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_BUDGET);
    if (!sheet) return { year: year, amounts: {} };

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return { year: year, amounts: {} };

    const headerMap = _getBudgetHeaderMap(sheet);
    _ensureBudgetHeaders(headerMap);
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const targetSet = new Set(normalizedTargets);
    const amounts = {};

    normalizedTargets.forEach(item => {
      amounts[item] = 0;
    });

    data.forEach(row => {
      const rowYear = String(row[headerMap.Financial_Year] || '').trim();
      if (rowYear !== year) return;
      const particulars = String(row[headerMap.Particulars] || '').trim();
      if (!targetSet.has(particulars)) return;
      const original = Number(row[headerMap.Original_Budget] || 0);
      const reallocation = Number(row[headerMap.Reallocation] || 0);
      const supplementary = Number(row[headerMap.Supplementary] || 0);
      amounts[particulars] = (amounts[particulars] || 0) + original + reallocation + supplementary;
    });

    return { year: year, amounts: amounts };
  } catch (error) {
    Logger.log('Error in getBudgetCurrentAmounts: ' + error.toString());
    return { year: String(financialYear || '').trim(), amounts: {} };
  }
}

/**
 * Get financial years that already have original budgets.
 */
function getOriginalBudgetYears() {
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_BUDGET);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const headerMap = _getBudgetHeaderMap(sheet);
  _ensureBudgetHeaders(headerMap);
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const years = new Set();

  data.forEach(row => {
    const year = String(row[headerMap.Financial_Year] || '').trim();
    const original = row[headerMap.Original_Budget];
    const originalValue = String(original == null ? '' : original).trim();
    if (year && originalValue !== '') {
      years.add(year);
    }
  });

  return Array.from(years).sort();
}

/**
 * Get original budget amounts for a financial year.
 */
function getOriginalBudgetByYear(financialYear) {
  const year = String(financialYear || '').trim();
  if (!year) return { amounts: {} };

  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_BUDGET);
  if (!sheet) return { amounts: {} };

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { amounts: {} };

  const headerMap = _getBudgetHeaderMap(sheet);
  _ensureBudgetHeaders(headerMap);
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const amounts = {};
  let firstDate = '';

  data.forEach(row => {
    const rowYear = String(row[headerMap.Financial_Year] || '').trim();
    if (rowYear !== year) return;
    const originalValue = row[headerMap.Original_Budget];
    if (originalValue === '' || originalValue == null) return;
    const category = String(row[headerMap.Category] || '').trim();
    if (!category) return;
    amounts[category] = Number(originalValue);
    if (!firstDate) {
      const dateValue = row[headerMap.Date];
      if (dateValue instanceof Date) {
        firstDate = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
    }
  });

  return { year: year, date: firstDate, amounts: amounts };
}

/**
 * Save original budget via bulk loader. Blocks if original budget exists for year.
 */
function saveOriginalBudget(payload) {
  if (!payload) throw new Error('Missing payload.');
  const financialYear = String(payload.financialYear || '').trim();
  if (!financialYear) throw new Error('Financial year is required.');
  const dateValue = payload.date ? new Date(payload.date) : new Date();
  if (Number.isNaN(dateValue.getTime())) throw new Error('Invalid budget date.');

  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  if (!rows.length) throw new Error('No budget rows provided.');

  const ss = _getOrCreateSpreadsheet();
  const sheet = _ensureSheet(ss, 'DB_BUDGET');
  const headerMap = _getBudgetHeaderMap(sheet);
  _ensureBudgetHeaders(headerMap);

  const data = sheet.getDataRange().getValues();
  const yearIndex = headerMap.Financial_Year;
  const originalIndex = headerMap.Original_Budget;
  const headerCount = sheet.getLastColumn();
  let removedCount = 0;
  let existingCount = 0;
  const user = getCurrentUser();
  const isAdmin = user && user.authenticated && String(user.role || '').toUpperCase() === 'ADMIN';
  if (data.length > 1) {
    const kept = [data[0]];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowYear = String(row[yearIndex] || '').trim();
      const originalValue = String(row[originalIndex] || '').trim();
      if (rowYear === financialYear && originalValue !== '') {
        existingCount += 1;
        if (!isAdmin) {
          throw new Error('Original budget already exists for ' + financialYear + '. Admin required to overwrite.');
        }
        removedCount += 1;
        continue;
      }
      kept.push(row);
    }
    if (removedCount) {
      sheet.getRange(1, 1, sheet.getLastRow(), headerCount).clearContent();
      sheet.getRange(1, 1, kept.length, headerCount).setValues(kept);
      logSystemEventSafe('OVERWRITE_ORIGINAL_BUDGET', financialYear, 'Removed rows: ' + removedCount);
    }
  }

  const rowsToInsert = rows.map(item => {
    const category = String(item.category || '').trim();
    const accountType = String(item.accountType || '').trim();
    const amount = Number(item.amount);

    if (!category) throw new Error('Category is required.');
    if (!Number.isFinite(amount)) throw new Error('Invalid amount for ' + category + '.');
    if (amount < 0) throw new Error('Original budget must be positive for ' + category + '.');

    const row = new Array(headerCount).fill('');
    row[headerMap.Date] = dateValue;
    row[headerMap.Financial_Year] = financialYear;
    row[headerMap.Category] = category;
    row[headerMap.Account_Type] = accountType;
    row[headerMap.Original_Budget] = amount;
    return row;
  });

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rowsToInsert.length, headerCount).setValues(rowsToInsert);

  logSystemEventSafe(
    'CREATE_ORIGINAL_BUDGET',
    financialYear,
    'Rows: ' + rowsToInsert.length + ', Date: ' + _formatDate_(dateValue)
  );

  return { success: true, count: rowsToInsert.length };
}

/**
 * Save a budget adjustment (Supplementary or Reallocation).
 */
function saveBudgetAdjustment(payload) {
  if (!payload) throw new Error('Missing payload.');
  const dateValue = payload.date ? new Date(payload.date) : new Date();
  const financialYear = String(payload.financialYear || '').trim();
  const type = String(payload.type || '').trim();
  const amount = Number(payload.amount);
  const authRef = String(payload.authRef || '').trim();
  const description = String(payload.description || '').trim();

  if (!financialYear) throw new Error('Financial year is required.');
  if (!type) throw new Error('Adjustment type is required.');
  if (!Number.isFinite(amount) || amount === 0) throw new Error('Amount must be non-zero.');
  if (!authRef) throw new Error('Auth reference is required.');
  if (!description) throw new Error('Description is required.');

  const ss = _getOrCreateSpreadsheet();
  const sheet = _ensureSheet(ss, 'DB_BUDGET');
  const headerMap = _getBudgetHeaderMap(sheet);
  _ensureBudgetHeaders(headerMap);
  const headerCount = sheet.getLastColumn();
  const rowsToInsert = [];

  if (type === 'Supplementary') {
    const particulars = String(payload.particulars || '').trim();
    if (!particulars) throw new Error('Particulars is required.');
    const details = _getParticularDetails(particulars);

    const row = new Array(headerCount).fill('');
    row[headerMap.Date] = dateValue;
    row[headerMap.Financial_Year] = financialYear;
    row[headerMap.Particulars] = particulars;
    row[headerMap.Sub_Category] = details.subCategory;
    row[headerMap.Category] = details.category;
    row[headerMap.Account_Type] = details.accountType;
    row[headerMap.Supplementary] = amount;
    row[headerMap.Auth_Ref] = authRef;
    row[headerMap.Description] = description;
    rowsToInsert.push(row);
  } else if (type === 'Reallocation') {
    const fromParticular = String(payload.fromParticulars || '').trim();
    const toParticular = String(payload.toParticulars || '').trim();
    if (!fromParticular || !toParticular) throw new Error('From and To particulars are required.');
    if (fromParticular === toParticular) throw new Error('From and To particulars must be different.');

    const delta = Math.abs(amount);
    const fromDetails = _getParticularDetails(fromParticular);
    const toDetails = _getParticularDetails(toParticular);

    const fromRow = new Array(headerCount).fill('');
    fromRow[headerMap.Date] = dateValue;
    fromRow[headerMap.Financial_Year] = financialYear;
    fromRow[headerMap.Particulars] = fromParticular;
    fromRow[headerMap.Sub_Category] = fromDetails.subCategory;
    fromRow[headerMap.Category] = fromDetails.category;
    fromRow[headerMap.Account_Type] = fromDetails.accountType;
    fromRow[headerMap.Reallocation] = -delta;
    fromRow[headerMap.Auth_Ref] = authRef;
    fromRow[headerMap.Description] = description;

    const toRow = new Array(headerCount).fill('');
    toRow[headerMap.Date] = dateValue;
    toRow[headerMap.Financial_Year] = financialYear;
    toRow[headerMap.Particulars] = toParticular;
    toRow[headerMap.Sub_Category] = toDetails.subCategory;
    toRow[headerMap.Category] = toDetails.category;
    toRow[headerMap.Account_Type] = toDetails.accountType;
    toRow[headerMap.Reallocation] = delta;
    toRow[headerMap.Auth_Ref] = authRef;
    toRow[headerMap.Description] = description;

    rowsToInsert.push(fromRow, toRow);
  } else {
    throw new Error('Unknown adjustment type.');
  }

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rowsToInsert.length, headerCount).setValues(rowsToInsert);
  logSystemEventSafe(
    'CREATE_BUDGET_ADJUSTMENT',
    financialYear,
    'Type: ' + type + ', Rows: ' + rowsToInsert.length + ', Auth: ' + authRef
  );
  return { success: true, count: rowsToInsert.length };
}

function _getParticularDetails(particulars) {
  const details = getDetailsForParticulars(particulars);
  if (!details || !details.subCategory) {
    throw new Error('Particulars not found: ' + particulars + '.');
  }
  return {
    subCategory: details.subCategory,
    category: details.category,
    accountType: details.accountType || ''
  };
}

function _getBudgetHeaderMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((header, index) => {
    const raw = String(header || '').trim();
    map[raw] = index;
    const normalized = raw.toLowerCase().replace(/\s+/g, '_');
    if (normalized === 'date') map.Date = index;
    if (normalized === 'financial_year') map.Financial_Year = index;
    if (normalized === 'particulars') map.Particulars = index;
    if (normalized === 'sub_category') map.Sub_Category = index;
    if (normalized === 'category') map.Category = index;
    if (normalized === 'account_type') map.Account_Type = index;
    if (normalized === 'original_budget') map.Original_Budget = index;
    if (normalized === 'reallocation') map.Reallocation = index;
    if (normalized === 'supplementary') map.Supplementary = index;
    if (normalized === 'final_budget') map.Final_Budget = index;
    if (normalized === 'actual_amount') map.Actual_Amount = index;
    if (normalized === 'variance') map.Variance = index;
    if (normalized === 'auth_ref') map.Auth_Ref = index;
    if (normalized === 'description') map.Description = index;
  });
  return map;
}

function _ensureBudgetHeaders(map) {
  const required = [
    'Date',
    'Financial_Year',
    'Particulars',
    'Sub_Category',
    'Category',
    'Account_Type',
    'Original_Budget',
    'Reallocation',
    'Supplementary',
    'Final_Budget',
    'Actual_Amount',
    'Variance',
    'Auth_Ref',
    'Description'
  ];
  required.forEach(header => {
    if (map[header] === undefined) {
      throw new Error('DB_BUDGET headers are missing. Please run cleanupSheetHeaders().');
    }
  });
}

function getDashboardHtml() {
  const user = getCurrentUser();
  if (user.authenticated) {
    return HtmlService.createTemplateFromFile('cDashboard').evaluate().getContent();
  } else {
    // If session is not valid, return login page content.
    // This could happen if the user's session expires and they try to navigate.
    return HtmlService.createTemplateFromFile('bLogin').evaluate().getContent();
  }
}

/**
 * Include other HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// PROCUREMENT MODULE - Server Functions
// ============================================================

/**
 * Generate unique ID with prefix
 */
function _generateId(prefix) {
  const timestamp = new Date().getTime();
  const random = Math.floor(Math.random() * 1000);
  return `${prefix}-${timestamp}-${random}`;
}

/**
 * Generate sequential number (e.g., PO-2024-0001)
 */
function _generateSequentialNumber(prefix, sheet, columnIndex) {
  const year = new Date().getFullYear();
  const lastRow = sheet.getLastRow();
  let maxNum = 0;

  if (lastRow > 1) {
    const data = sheet.getRange(2, columnIndex, lastRow - 1, 1).getValues();
    data.forEach(row => {
      const num = String(row[0] || '');
      const match = num.match(new RegExp(`${prefix}-${year}-(\\d+)`));
      if (match) {
        const seq = parseInt(match[1], 10);
        if (seq > maxNum) maxNum = seq;
      }
    });
  }

  return `${prefix}-${year}-${String(maxNum + 1).padStart(4, '0')}`;
}


// ============================================================
// CONTACTS MODULE - Unified Contacts (Suppliers, Customers, Staff, Government Entities, Donors)
// ============================================================

const CONTACT_HEADERS = SHEET_HEADERS.CONTACTS;

/**
 * Get all contacts, optionally filtered by type
 */
function getContacts(contactType) {
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.CONTACTS);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, CONTACT_HEADERS.length).getValues();

  let results = data.map(row => {
    const obj = {};
    CONTACT_HEADERS.forEach((h, i) => {
      if (h.includes('Date') && row[i] instanceof Date) {
        obj[h] = Utilities.formatDate(row[i], Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        obj[h] = row[i];
      }
    });
    return obj;
  }).filter(c => c.Contact_ID);

  if (contactType) {
    results = results.filter(c => c.Contact_Type === contactType);
  }

  return results;
}

/**
 * Save a new contact
 */
function saveContact(contactType, data) {
  const ss = _getOrCreateSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.CONTACTS);
  if (!sheet) sheet = _ensureSheet(ss, 'CONTACTS');

  const user = getCurrentUser();
  const prefix = contactType === 'Supplier'
    ? 'SUP'
    : contactType === 'Customer'
      ? 'CUS'
      : contactType === 'Staff'
        ? 'STA'
        : contactType === 'Government Entity'
          ? 'GOV'
          : contactType === 'Donor'
            ? 'DON'
            : 'CON';
  const contactId = _generateId(prefix);
  const now = new Date();

  const row = [
    contactId,
    contactType,
    data.contactName || data.supplierName || data.customerName || '',
    data.contactPerson || '',
    data.phone || '',
    data.email || '',
    data.address || '',
    data.taxPin || data.kraPin || '',
    data.bankName || '',
    data.bankAccount || '',
    data.category || 'General',
    parseFloat(data.creditLimit) || 0,
    data.paymentTerms || 'Net 30',
    'Active',
    now,
    user.email || ''
  ];

  sheet.appendRow(row);
  logSystemEvent(user.email, 'CREATE_CONTACT', contactId, row[2]);
  CacheService.getScriptCache().remove('dropdownData');

  return { success: true, contactId: contactId };
}

/**
 * Update contact
 */
function updateContact(contactId, data) {
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.CONTACTS);
  if (!sheet) return { success: false, message: 'Contacts sheet not found' };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'Contact not found' };

  const idColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let rowIndex = -1;

  for (let i = 0; i < idColumn.length; i++) {
    if (idColumn[i][0] === contactId) {
      rowIndex = i + 2;
      break;
    }
  }

  if (rowIndex === -1) return { success: false, message: 'Contact not found' };

  // Update fields (columns 3-14, excluding ID, Type, Created_Date, Created_By)
  const updates = [
    data.contactName || data.supplierName || data.customerName || '',
    data.contactPerson || '',
    data.phone || '',
    data.email || '',
    data.address || '',
    data.taxPin || data.kraPin || '',
    data.bankName || '',
    data.bankAccount || '',
    data.category || '',
    parseFloat(data.creditLimit) || 0,
    data.paymentTerms || '',
    data.status || 'Active'
  ];

  sheet.getRange(rowIndex, 3, 1, 12).setValues([updates]);
  logSystemEvent(getCurrentUser().email, 'UPDATE_CONTACT', contactId, updates[0]);

  return { success: true };
}

// ============================================================
// SUPPLIERS - Wrapper functions for backward compatibility
// ============================================================

function getSuppliers() {
  return getContacts('Supplier').map(c => ({
    Supplier_ID: c.Contact_ID,
    Supplier_Name: c.Contact_Name,
    Contact_Person: c.Contact_Person,
    Phone: c.Phone,
    Email: c.Email,
    Address: c.Address,
    KRA_PIN: c.Tax_PIN,
    Bank_Name: c.Bank_Name,
    Bank_Account: c.Bank_Account,
    Category: c.Category,
    Payment_Terms: c.Payment_Terms,
    Status: c.Status,
    Created_Date: c.Created_Date,
    Created_By: c.Created_By
  }));
}

function getActiveSuppliers() {
  return getSuppliers().filter(s => s.Status === 'Active');
}

function saveSupplier(data) {
  data.contactName = data.supplierName;
  data.taxPin = data.kraPin;
  const result = saveContact('Supplier', data);
  return { success: result.success, supplierId: result.contactId };
}

function updateSupplier(supplierId, data) {
  data.contactName = data.supplierName;
  data.taxPin = data.kraPin;
  return updateContact(supplierId, data);
}

// ============================================================
// CUSTOMERS - Wrapper functions for backward compatibility
// ============================================================

function getCustomers(filters) {
  let results = getContacts('Customer').map(c => ({
    Customer_ID: c.Contact_ID,
    Customer_Name: c.Contact_Name,
    Contact_Person: c.Contact_Person,
    Phone: c.Phone,
    Email: c.Email,
    Address: c.Address,
    Customer_Type: c.Category,
    Credit_Limit: c.Credit_Limit,
    Payment_Terms: c.Payment_Terms,
    Status: c.Status,
    Created_Date: c.Created_Date,
    Created_By: c.Created_By
  }));

  if (filters) {
    if (filters.status) results = results.filter(c => c.Status === filters.status);
    if (filters.customerType) results = results.filter(c => c.Customer_Type === filters.customerType);
  }

  return results;
}

function getActiveCustomers() {
  return getCustomers().filter(c => c.Status === 'Active');
}

function saveCustomer(data) {
  data.contactName = data.customerName;
  data.category = data.customerType;
  const result = saveContact('Customer', data);
  return { success: result.success, customerId: result.contactId };
}

function updateCustomer(customerId, data) {
  data.contactName = data.customerName;
  data.category = data.customerType;
  return updateContact(customerId, data);
}

// ============================================================
// STAFF - Wrapper functions for contact type Staff
// ============================================================

function getStaff(filters) {
  let results = getContacts('Staff').map(c => ({
    Staff_ID: c.Contact_ID,
    Staff_Name: c.Contact_Name,
    Phone: c.Phone,
    Email: c.Email,
    Department: c.Category,
    Status: c.Status,
    Created_Date: c.Created_Date,
    Created_By: c.Created_By
  }));

  if (filters && filters.status) {
    results = results.filter(s => s.Status === filters.status);
  }
  return results;
}

function getActiveStaff() {
  return getStaff().filter(s => s.Status === 'Active');
}

function saveStaff(data) {
  data.contactName = data.staffName;
  data.category = data.department || 'General';
  const result = saveContact('Staff', data);
  return { success: result.success, staffId: result.contactId };
}

function updateStaff(staffId, data) {
  data.contactName = data.staffName;
  data.category = data.department;
  return updateContact(staffId, data);
}

// ============================================================
// PURCHASE ORDERS CRUD
// ============================================================

/**
 * Get all purchase orders
 */
function getPurchaseOrders(filters) {
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE_ORDERS);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const headers = SHEET_HEADERS.PURCHASE_ORDERS;
  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  let results = data.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      if (h.includes('Date') && row[i] instanceof Date) {
        obj[h] = Utilities.formatDate(row[i], Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else if (h === 'Line_Items') {
        // Parse JSON line items
        try {
          obj[h] = row[i] ? JSON.parse(row[i]) : [];
        } catch (e) {
          obj[h] = [];
        }
      } else {
        obj[h] = row[i];
      }
    });
    return obj;
  }).filter(po => po.PO_ID);

  // Apply filters
  if (filters) {
    if (filters.status) {
      results = results.filter(po => po.Status === filters.status);
    }
    if (filters.supplierId) {
      results = results.filter(po => po.Supplier_ID === filters.supplierId);
    }
    if (filters.financialYear) {
      results = results.filter(po => po.Financial_Year === filters.financialYear);
    }
  }

  return results;
}

/**
 * Get PO with lines (lines are stored as JSON in Line_Items column)
 */
function getPurchaseOrderWithLines(poId) {
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE_ORDERS);
  if (!sheet) return null;

  const headers = SHEET_HEADERS.PURCHASE_ORDERS;
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === poId) {
      const poRecord = {};
      headers.forEach((h, idx) => {
        if (h.includes('Date') && data[i][idx] instanceof Date) {
          poRecord[h] = Utilities.formatDate(data[i][idx], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else if (h === 'Line_Items') {
          try {
            poRecord.lines = data[i][idx] ? JSON.parse(data[i][idx]) : [];
          } catch (e) {
            poRecord.lines = [];
          }
        } else {
          poRecord[h] = data[i][idx];
        }
      });
      return poRecord;
    }
  }

  return null;
}

/**
 * Save a new purchase order with lines (lines stored as JSON)
 */
function savePurchaseOrder(data) {
  const ss = _getOrCreateSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE_ORDERS);
  if (!sheet) sheet = _ensureSheet(ss, 'PURCHASE_ORDERS');

  const user = getCurrentUser();
  const poId = _generateId('PO');
  const poNumber = _generateSequentialNumber('PO', sheet, 2);
  const now = new Date();

  // Process line items and calculate total
  let totalAmount = 0;
  const lineItems = (data.lines || []).map((line, index) => {
    const qty = parseFloat(line.quantity) || 0;
    const price = parseFloat(line.unitPrice) || 0;
    const lineTotal = qty * price;
    totalAmount += lineTotal;

    return {
      lineNo: index + 1,
      itemDescription: line.itemDescription || '',
      particulars: line.particulars || '',
      subCategory: line.subCategory || '',
      category: line.category || '',
      quantity: qty,
      unit: line.unit || 'Each',
      unitPrice: price,
      totalPrice: lineTotal,
      receivedQty: 0,
      status: 'Pending'
    };
  });

  // Save PO with lines as JSON
  const poRow = [
    poId,
    poNumber,
    data.poDate ? new Date(data.poDate) : now,
    data.financialYear || '',
    data.supplierId || '',
    data.supplierName || '',
    data.description || '',
    totalAmount,
    'Draft',
    user.email || '',
    now,
    '', // Approved_By
    '', // Approved_Date
    data.deliveryDate ? new Date(data.deliveryDate) : '',
    data.notes || '',
    JSON.stringify(lineItems) // Line_Items as JSON
  ];

  sheet.appendRow(poRow);
  logSystemEvent(user.email, 'CREATE_PO', poId, poNumber);

  return { success: true, poId: poId, poNumber: poNumber };
}

/**
 * Approve purchase order
 */
function approvePurchaseOrder(poId) {
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE_ORDERS);
  if (!sheet) return { success: false, message: 'PO sheet not found' };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'PO not found' };

  const data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  let rowIndex = -1;

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === poId) {
      rowIndex = i + 2;
      break;
    }
  }

  if (rowIndex === -1) return { success: false, message: 'PO not found' };

  const user = getCurrentUser();
  const now = new Date();

  // Update status, approved_by, approved_date (columns 9, 12, 13)
  sheet.getRange(rowIndex, 9).setValue('Approved');
  sheet.getRange(rowIndex, 12).setValue(user.email);
  sheet.getRange(rowIndex, 13).setValue(now);

  logSystemEvent(user.email, 'APPROVE_PO', poId, '');

  return { success: true };
}

// ============================================================
// GOODS RECEIVED NOTES (GRN) CRUD
// ============================================================

/**
 * Get all GRNs
 */
function getGoodsReceivedNotes(filters) {
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.GRN);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const headers = SHEET_HEADERS.GRN;
  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  let results = data.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      if (h.includes('Date') && row[i] instanceof Date) {
        obj[h] = Utilities.formatDate(row[i], Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else if (h === 'Line_Items') {
        try {
          obj[h] = row[i] ? JSON.parse(row[i]) : [];
        } catch (e) {
          obj[h] = [];
        }
      } else {
        obj[h] = row[i];
      }
    });
    return obj;
  }).filter(grn => grn.GRN_ID);

  if (filters && filters.poId) {
    results = results.filter(grn => grn.PO_ID === filters.poId);
  }

  return results;
}

/**
 * Save GRN (Goods Received Note) with lines as JSON
 */
function saveGoodsReceivedNote(data) {
  const ss = _getOrCreateSpreadsheet();
  let grnSheet = ss.getSheetByName(CONFIG.SHEETS.GRN);
  if (!grnSheet) grnSheet = _ensureSheet(ss, 'GRN');

  const user = getCurrentUser();
  const grnId = _generateId('GRN');
  const grnNumber = _generateSequentialNumber('GRN', grnSheet, 2);
  const now = new Date();

  // Process GRN line items
  const lineItems = (data.lines || []).map((line, index) => {
    const qtyReceived = parseFloat(line.qtyReceived) || 0;
    const qtyAccepted = parseFloat(line.qtyAccepted) || qtyReceived;
    const qtyRejected = parseFloat(line.qtyRejected) || 0;
    const unitPrice = parseFloat(line.unitPrice) || 0;

    return {
      lineNo: index + 1,
      poLineNo: line.poLineNo || (index + 1),
      itemDescription: line.itemDescription || '',
      particulars: line.particulars || '',
      subCategory: line.subCategory || '',
      category: line.category || '',
      qtyOrdered: parseFloat(line.qtyOrdered) || 0,
      qtyReceived: qtyReceived,
      qtyAccepted: qtyAccepted,
      qtyRejected: qtyRejected,
      unitPrice: unitPrice,
      totalPrice: qtyAccepted * unitPrice,
      rejectionReason: line.rejectionReason || ''
    };
  });

  // Save GRN with lines as JSON
  const grnRow = [
    grnId,
    grnNumber,
    data.grnDate ? new Date(data.grnDate) : now,
    data.poId || '',
    data.poNumber || '',
    data.supplierId || '',
    data.supplierName || '',
    user.email || '',
    'Received',
    data.invoiceNumber || '',
    data.invoiceDate ? new Date(data.invoiceDate) : '',
    parseFloat(data.invoiceAmount) || 0,
    data.notes || '',
    JSON.stringify(lineItems) // Line_Items as JSON
  ];

  grnSheet.appendRow(grnRow);

  // Update PO line items received quantities
  if (data.poId) {
    _updatePOLineItemsReceived(ss, data.poId, lineItems);
  }

  // Auto-create payable if invoice details provided
  let payableResult = null;
  if (data.invoiceNumber && parseFloat(data.invoiceAmount) > 0) {
    // Get supplier payment terms
    let paymentTerms = 'Net 30';
    if (data.supplierId) {
      const suppliers = getSuppliers();
      const supplier = suppliers.find(s => s.Supplier_ID === data.supplierId);
      if (supplier && supplier.Payment_Terms) {
        paymentTerms = supplier.Payment_Terms;
      }
    }

    payableResult = createPayableFromGRN({
      grnId: grnId,
      grnNumber: grnNumber,
      poId: data.poId || '',
      poNumber: data.poNumber || '',
      supplierId: data.supplierId || '',
      supplierName: data.supplierName || '',
      invoiceNumber: data.invoiceNumber,
      invoiceDate: data.invoiceDate,
      invoiceAmount: data.invoiceAmount,
      paymentTerms: paymentTerms,
      lineItems: lineItems
    });
  }

  logSystemEvent(user.email, 'CREATE_GRN', grnId, grnNumber);

  return {
    success: true,
    grnId: grnId,
    grnNumber: grnNumber,
    payableCreated: payableResult ? payableResult.success : false,
    payableId: payableResult ? payableResult.payableId : null
  };
}

/**
 * Update PO line items received quantities (JSON-based)
 */
function _updatePOLineItemsReceived(ss, poId, grnLines) {
  const poSheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE_ORDERS);
  if (!poSheet) return;

  const headers = SHEET_HEADERS.PURCHASE_ORDERS;
  const lineItemsIndex = headers.indexOf('Line_Items');
  const statusIndex = headers.indexOf('Status');
  const data = poSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === poId) {
      // Parse existing line items
      let poLines = [];
      try {
        poLines = data[i][lineItemsIndex] ? JSON.parse(data[i][lineItemsIndex]) : [];
      } catch (e) {
        poLines = [];
      }

      // Update received quantities
      let allReceived = true;
      let anyReceived = false;

      poLines.forEach(poLine => {
        // Find matching GRN line
        const grnLine = grnLines.find(g => g.lineNo === poLine.lineNo || g.poLineNo === poLine.lineNo);
        if (grnLine) {
          poLine.receivedQty = (poLine.receivedQty || 0) + grnLine.qtyAccepted;
          if (poLine.receivedQty >= poLine.quantity) {
            poLine.status = 'Received';
            anyReceived = true;
          } else if (poLine.receivedQty > 0) {
            poLine.status = 'Partial';
            anyReceived = true;
            allReceived = false;
          } else {
            allReceived = false;
          }
        } else {
          if (poLine.status !== 'Received') {
            allReceived = false;
          } else {
            anyReceived = true;
          }
        }
      });

      // Update PO line items JSON
      poSheet.getRange(i + 1, lineItemsIndex + 1).setValue(JSON.stringify(poLines));

      // Update PO status
      let newStatus = data[i][statusIndex];
      if (allReceived && anyReceived) {
        newStatus = 'Completed';
      } else if (anyReceived) {
        newStatus = 'Partial';
      }
      poSheet.getRange(i + 1, statusIndex + 1).setValue(newStatus);

      break;
    }
  }
}

/**
 * Get pending POs for GRN (approved but not fully received)
 */
function getPendingPOsForGRN() {
  return getPurchaseOrders({ status: 'Approved' }).concat(
    getPurchaseOrders({ status: 'Partial' })
  );
}

/**
 * Alias for initializeSpreadsheet - for backwards compatibility
 */
function cleanupLegacySheets() {
  return initializeSpreadsheet();
}


// ============================================================
// INVOICES MODULE - Unified AP/AR (Payables & Receivables)
// ============================================================

const INVOICE_HEADERS = SHEET_HEADERS.INVOICES;

function getInvoices(invoiceType, filters) {
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.INVOICES);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, INVOICE_HEADERS.length).getValues();
  let results = data.map(row => {
    const obj = {};
    INVOICE_HEADERS.forEach((h, i) => {
      if (h.includes('Date') && row[i] instanceof Date) {
        obj[h] = Utilities.formatDate(row[i], Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else if (h === 'Line_Items') {
        try { obj[h] = row[i] ? JSON.parse(row[i]) : []; } catch (e) { obj[h] = []; }
      } else { obj[h] = row[i]; }
    });
    if (obj.Due_Date && obj.Status !== 'Paid') {
      const dueDate = new Date(obj.Due_Date);
      const diffDays = Math.floor((new Date() - dueDate) / (1000 * 60 * 60 * 24));
      obj.Days_Overdue = diffDays > 0 ? diffDays : 0;
    } else { obj.Days_Overdue = 0; }
    return obj;
  }).filter(inv => inv.Invoice_ID);
  if (invoiceType) results = results.filter(inv => inv.Invoice_Type === invoiceType);
  if (filters) {
    if (filters.status) results = results.filter(inv => inv.Status === filters.status);
    if (filters.contactId) results = results.filter(inv => inv.Contact_ID === filters.contactId);
    if (filters.overdue) results = results.filter(inv => inv.Days_Overdue > 0);
  }
  return results;
}

function createInvoice(invoiceType, data) {
  const ss = _getOrCreateSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.INVOICES);
  if (!sheet) sheet = _ensureSheet(ss, 'INVOICES');
  const user = getCurrentUser();
  const prefix = invoiceType === 'AP' ? 'PAY' : 'REC';
  const invoiceId = _generateId(prefix);
  const now = new Date();
  const invoiceDate = data.invoiceDate ? new Date(data.invoiceDate) : now;
  const paymentTerms = data.paymentTerms || 'Net 30';
  const daysToAdd = parseInt(paymentTerms.replace(/\D/g, '')) || 30;
  const dueDate = new Date(invoiceDate);
  dueDate.setDate(dueDate.getDate() + daysToAdd);
  const amount = parseFloat(data.amount || data.invoiceAmount) || 0;
  const row = [
    invoiceId, invoiceType,
    data.invoiceNumber || _generateSequentialNumber(invoiceType === 'AP' ? 'BILL' : 'INV', sheet, 3),
    invoiceDate, dueDate, data.financialYear || '',
    data.contactId || data.supplierId || data.customerId || '',
    data.contactName || data.supplierName || data.customerName || '',
    data.grnId || '', data.grnNumber || '', data.poId || '', data.poNumber || '',
    data.description || '', amount, 0, amount, 'Pending', paymentTerms, now,
    user.email || '', JSON.stringify(data.lineItems || [])
  ];
  sheet.appendRow(row);
  logSystemEvent(user.email, 'CREATE_INVOICE', invoiceId, invoiceType + ': ' + row[2]);
  return { success: true, invoiceId: invoiceId, invoiceNumber: row[2] };
}

function recordInvoicePayment(invoiceId, paymentData) {
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.INVOICES);
  if (!sheet) return { success: false, message: 'Invoices sheet not found' };
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'Invoice not found' };
  const data = sheet.getRange(2, 1, lastRow - 1, INVOICE_HEADERS.length).getValues();
  let rowIndex = -1, invoice = null;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === invoiceId) {
      rowIndex = i + 2;
      invoice = {};
      INVOICE_HEADERS.forEach((h, j) => invoice[h] = data[i][j]);
      break;
    }
  }
  if (!invoice) return { success: false, message: 'Invoice not found' };
  const paymentAmount = parseFloat(paymentData.amount) || 0;
  const currentPaid = parseFloat(invoice.Paid_Amount) || 0;
  const totalAmount = parseFloat(invoice.Amount) || 0;
  const newPaidAmount = currentPaid + paymentAmount;
  const newBalance = totalAmount - newPaidAmount;
  const newStatus = newBalance <= 0 ? 'Paid' : 'Partial';
  const paidAmountCol = INVOICE_HEADERS.indexOf('Paid_Amount') + 1;
  const balanceCol = INVOICE_HEADERS.indexOf('Balance') + 1;
  const statusCol = INVOICE_HEADERS.indexOf('Status') + 1;
  sheet.getRange(rowIndex, paidAmountCol).setValue(newPaidAmount);
  sheet.getRange(rowIndex, balanceCol).setValue(Math.max(0, newBalance));
  sheet.getRange(rowIndex, statusCol).setValue(newStatus);
  _createPaymentJournalEntry(invoice, paymentData, invoice.Invoice_Type);
  logSystemEvent(getCurrentUser().email, 'RECORD_PAYMENT', invoiceId, 'Amount: ' + paymentAmount);
  return { success: true, newBalance: Math.max(0, newBalance), newStatus: newStatus };
}

function getInvoiceAgingSummary(invoiceType) {
  const invoices = getInvoices(invoiceType, { status: 'Pending' }).concat(getInvoices(invoiceType, { status: 'Partial' }));
  const summary = { current: 0, days_1_30: 0, days_31_60: 0, days_61_90: 0, days_over_90: 0, total: 0 };
  invoices.forEach(inv => {
    const balance = parseFloat(inv.Balance) || 0;
    summary.total += balance;
    if (inv.Days_Overdue <= 0) summary.current += balance;
    else if (inv.Days_Overdue <= 30) summary.days_1_30 += balance;
    else if (inv.Days_Overdue <= 60) summary.days_31_60 += balance;
    else if (inv.Days_Overdue <= 90) summary.days_61_90 += balance;
    else summary.days_over_90 += balance;
  });
  return summary;
}

// PAYABLES wrappers
function getPayables(filters) {
  const f = filters ? { status: filters.status, contactId: filters.supplierId, overdue: filters.overdue } : undefined;
  return getInvoices('AP', f).map(inv => ({
    Payable_ID: inv.Invoice_ID, Invoice_Number: inv.Invoice_Number, Invoice_Date: inv.Invoice_Date,
    Due_Date: inv.Due_Date, Financial_Year: inv.Financial_Year, Supplier_ID: inv.Contact_ID,
    Supplier_Name: inv.Contact_Name, GRN_ID: inv.GRN_ID, GRN_Number: inv.GRN_Number,
    PO_ID: inv.PO_ID, PO_Number: inv.PO_Number, Amount: inv.Amount, Paid_Amount: inv.Paid_Amount,
    Balance: inv.Balance, Status: inv.Status, Payment_Terms: inv.Payment_Terms,
    Created_Date: inv.Created_Date, Created_By: inv.Created_By, Line_Items: inv.Line_Items, Days_Overdue: inv.Days_Overdue
  }));
}
function getPayablesAgingSummary() { return getInvoiceAgingSummary('AP'); }
function createPayableFromGRN(grnData) {
  grnData.contactId = grnData.supplierId; grnData.contactName = grnData.supplierName; grnData.amount = grnData.invoiceAmount;
  return createInvoice('AP', grnData);
}
function recordPayablePayment(payableId, paymentData) { return recordInvoicePayment(payableId, paymentData); }
function getSupplierPayables(supplierId) { return getPayables({ supplierId: supplierId }).filter(p => p.Status !== 'Paid'); }

// RECEIVABLES wrappers
function getReceivables(filters) {
  const f = filters ? { status: filters.status, contactId: filters.customerId, overdue: filters.overdue } : undefined;
  return getInvoices('AR', f).map(inv => ({
    Receivable_ID: inv.Invoice_ID, Invoice_Number: inv.Invoice_Number, Invoice_Date: inv.Invoice_Date,
    Due_Date: inv.Due_Date, Financial_Year: inv.Financial_Year, Customer_ID: inv.Contact_ID,
    Customer_Name: inv.Contact_Name, Description: inv.Description, Amount: inv.Amount,
    Received_Amount: inv.Paid_Amount, Balance: inv.Balance, Status: inv.Status, Payment_Terms: inv.Payment_Terms,
    Created_Date: inv.Created_Date, Created_By: inv.Created_By, Line_Items: inv.Line_Items, Days_Overdue: inv.Days_Overdue
  }));
}
function getReceivablesAgingSummary() { return getInvoiceAgingSummary('AR'); }
function createReceivable(data) {
  data.contactId = data.customerId; data.contactName = data.customerName;
  const result = createInvoice('AR', data);
  return { success: result.success, receivableId: result.invoiceId, invoiceNumber: result.invoiceNumber };
}
function recordReceivablePayment(receivableId, paymentData) { return recordInvoicePayment(receivableId, paymentData); }
function getCustomerReceivables(customerId) { return getReceivables({ customerId: customerId }).filter(r => r.Status !== 'Paid'); }
