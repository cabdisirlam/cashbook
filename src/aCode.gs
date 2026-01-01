/**
 * Financial System
 * PIN-Based Authentication System
 */

// Configuration
const CONFIG = {
  SPREADSHEET_NAME: "Financial System",
  SESSION_TIMEOUT: 5 * 60 * 1000, // 5 minutes in milliseconds
  SHEETS: {
    // PART 1: Core Transaction Database
    DB_JOURNAL: "DB_JOURNAL",
    DB_BANK: "DB_BANK",
    DB_BUDGET: "DB_BUDGET",

    // PART 2: Procurement Module
    SUPPLIERS: "SUPPLIERS",
    PURCHASE_ORDERS: "PURCHASE_ORDERS",
    PO_LINES: "PO_LINES",
    GRN: "GRN",
    GRN_LINES: "GRN_LINES",

    // PART 3: Master Data & System
    MASTER_DATA: "MASTER_DATA",
    SYS_USERS: "SYS_USERS",
    SYS_LOGS: "SYS_LOGS"
  }
};

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
 * PUBLIC: Initialize spreadsheet and sheets
 * Run this once to set up your Financial System
 */
function initializeSpreadsheet() {
  let ss = _getOrCreateSpreadsheet();

  // Initialize Core Transaction sheets
  _initializeDbJournalSheet(ss);
  _initializeDbBankSheet(ss);
  _initializeDbBudgetSheet(ss);

  // Initialize Procurement sheets
  _initializeSuppliersSheet(ss);
  _initializePurchaseOrdersSheet(ss);
  _initializePoLinesSheet(ss);
  _initializeGrnSheet(ss);
  _initializeGrnLinesSheet(ss);

  // Initialize Master Data & System sheets
  _initializeMasterDataSheet(ss);
  _initializeSysUsersSheet(ss);
  _initializeSysLogsSheet(ss);

  return {
    success: true,
    message: "Spreadsheet initialized successfully with 11 sheets",
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl()
  };
}

/**
 * PUBLIC: Clean up sheet headers on existing sheets
 * Run this to fix headers without recreating the entire spreadsheet
 */
function cleanupSheetHeaders() {
  Logger.log('Starting manual sheet headers cleanup...');

  const ss = _getOrCreateSpreadsheet();
  let updatedSheets = [];

  // Clean up DB_JOURNAL headers
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (sheet) {
    const headers = ['UUID', 'Batch_ID', 'Date', 'Financial_Year', 'Account_Code', 'Payee', 'Ref_No',
                     'Bank_Ref', 'Particulars', 'Sub_Category', 'Category', 'Description', 'Debit', 'Credit',
                     'Account_Type', 'Report_Mapping', 'Recon_Status', 'Receipt_URL'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
    updatedSheets.push('DB_JOURNAL');
    Logger.log('✓ DB_JOURNAL headers cleaned');
  }

  // Clean up DB_BANK headers
  sheet = ss.getSheetByName(CONFIG.SHEETS.DB_BANK);
  if (sheet) {
    const headers = ['Account_Code', 'Financial_Year', 'Txn_Date', 'Value_Date', 'Bank_Ref',
                     'Description', 'Debit', 'Credit', 'Balance', 'Match_Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
    updatedSheets.push('DB_BANK');
    Logger.log('✓ DB_BANK headers cleaned');
  }

  // Clean up DB_BUDGET headers
  sheet = ss.getSheetByName(CONFIG.SHEETS.DB_BUDGET);
  if (sheet) {
    const headers = ['Date', 'Financial_Year', 'Particulars', 'Sub_Category', 'Category', 'Account_Type',
                     'Original_Budget', 'Reallocation', 'Supplementary', 'Final_Budget',
                     'Actual_Amount', 'Variance', 'Auth_Ref', 'Description'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
    updatedSheets.push('DB_BUDGET');
    Logger.log('✓ DB_BUDGET headers cleaned');
  }

  // Clean up MASTER_DATA headers
  sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);
  if (sheet) {
    const headers = ['Payees', 'Particulars', 'Sub_Category', 'Category', 'Account_Codes', 'Account_Type', 'Report_Mapping', 'Financial_Year'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
    updatedSheets.push('MASTER_DATA');
    Logger.log('✓ MASTER_DATA headers cleaned');
  }

  // Clean up SYS_USERS headers
  sheet = ss.getSheetByName(CONFIG.SHEETS.SYS_USERS);
  if (sheet) {
    const headers = ['Email', 'PIN', 'Name', 'Role', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
    updatedSheets.push('SYS_USERS');
    Logger.log('✓ SYS_USERS headers cleaned');
  }

  // Clean up SYS_LOGS headers
  sheet = ss.getSheetByName(CONFIG.SHEETS.SYS_LOGS);
  if (sheet) {
    const headers = ['Timestamp', 'User', 'Action', 'Target_ID', 'Details'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
    updatedSheets.push('SYS_LOGS');
    Logger.log('✓ SYS_LOGS headers cleaned');
  }

  const message = `Sheet headers cleanup completed!\nUpdated ${updatedSheets.length} sheets: ${updatedSheets.join(', ')}`;
  Logger.log(message);

  // Show result in a UI alert if running from editor
  try {
    SpreadsheetApp.getUi().alert('Cleanup Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    // If UI not available, just log
    Logger.log('UI not available, logged to console instead');
  }

  return {
    success: true,
    message: message,
    updatedSheets: updatedSheets
  };
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
 * PART 1: Core Transaction Database
 */

// 1. DB_JOURNAL Sheet
function _initializeDbJournalSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.DB_JOURNAL);
    const headers = ['UUID', 'Batch_ID', 'Date', 'Financial_Year', 'Account_Code', 'Payee', 'Ref_No',
                     'Bank_Ref', 'Particulars', 'Sub_Category', 'Category', 'Description', 'Debit', 'Credit',
                     'Account_Type', 'Report_Mapping', 'Recon_Status', 'Receipt_URL'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
  }
  return sheet;
}

// 2. DB_BANK Sheet
function _initializeDbBankSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DB_BANK);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.DB_BANK);
    const headers = ['Account_Code', 'Financial_Year', 'Txn_Date', 'Value_Date', 'Bank_Ref',
                     'Description', 'Debit', 'Credit', 'Balance', 'Match_Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
  }
  return sheet;
}

// 3. DB_BUDGET Sheet
function _initializeDbBudgetSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DB_BUDGET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.DB_BUDGET);
    const headers = ['Date', 'Financial_Year', 'Particulars', 'Sub_Category', 'Category', 'Account_Type',
                     'Original_Budget', 'Reallocation', 'Supplementary', 'Final_Budget',
                     'Actual_Amount', 'Variance', 'Auth_Ref', 'Description'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);

    // Add a note about never deleting rows
    sheet.getRange('A2').setNote('CRITICAL RULE: Never delete rows! For adjustments, add new rows with Supplementary or Reallocation amounts.');
  }
  return sheet;
}

/**
 * PART 2: Procurement Module
 */

// 4. SUPPLIERS Sheet
function _initializeSuppliersSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.SUPPLIERS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.SUPPLIERS);
    const headers = [
      'Supplier_ID', 'Supplier_Name', 'Contact_Person', 'Phone', 'Email',
      'Address', 'KRA_PIN', 'Bank_Name', 'Bank_Account', 'Category',
      'Payment_Terms', 'Status', 'Created_Date', 'Created_By'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
  }
  return sheet;
}

// 5. PURCHASE_ORDERS Sheet
function _initializePurchaseOrdersSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE_ORDERS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.PURCHASE_ORDERS);
    const headers = [
      'PO_ID', 'PO_Number', 'PO_Date', 'Financial_Year', 'Supplier_ID',
      'Supplier_Name', 'Description', 'Total_Amount', 'Status',
      'Requested_By', 'Requested_Date', 'Approved_By', 'Approved_Date',
      'Delivery_Date', 'Notes'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
  }
  return sheet;
}

// 6. PO_LINES Sheet
function _initializePoLinesSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.PO_LINES);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.PO_LINES);
    const headers = [
      'Line_ID', 'PO_ID', 'Line_No', 'Item_Description', 'Particulars',
      'Sub_Category', 'Category', 'Quantity', 'Unit', 'Unit_Price',
      'Total_Price', 'Received_Qty', 'Status'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
  }
  return sheet;
}

// 7. GRN Sheet (Goods Received Notes)
function _initializeGrnSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.GRN);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.GRN);
    const headers = [
      'GRN_ID', 'GRN_Number', 'GRN_Date', 'PO_ID', 'PO_Number',
      'Supplier_ID', 'Supplier_Name', 'Received_By', 'Status',
      'Invoice_Number', 'Invoice_Date', 'Invoice_Amount', 'Notes'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
  }
  return sheet;
}

// 8. GRN_LINES Sheet
function _initializeGrnLinesSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.GRN_LINES);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.GRN_LINES);
    const headers = [
      'Line_ID', 'GRN_ID', 'PO_Line_ID', 'Line_No', 'Item_Description',
      'Qty_Ordered', 'Qty_Received', 'Qty_Accepted', 'Qty_Rejected',
      'Unit_Price', 'Total_Price', 'Rejection_Reason'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
  }
  return sheet;
}

/**
 * PART 3: Master Data & System
 */

// 9. MASTER_DATA Sheet
function _initializeMasterDataSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.MASTER_DATA);

    // Create headers - All data will come from user input
    const headers = ['Payees', 'Particulars', 'Sub_Category', 'Category', 'Account_Codes', 'Account_Type', 'Report_Mapping', 'Financial_Year'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);

    // Auto-resize columns for better visibility
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

// 10. SYS_USERS Sheet
function _initializeSysUsersSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.SYS_USERS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.SYS_USERS);

    // Set headers
    const headers = ['Email', 'PIN', 'Name', 'Role', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);

    // Add default admin user
    const defaultUser = [
      'cabdisirlam@gmail.com',
      '1234',
      'Admin User',
      'ADMIN',
      'Active'
    ];
    sheet.getRange(2, 1, 1, defaultUser.length).setValues([defaultUser]);
  }
  return sheet;
}

// 11. SYS_LOGS Sheet
function _initializeSysLogsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.SYS_LOGS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.SYS_LOGS);
    const headers = ['Timestamp', 'User', 'Action', 'Target_ID', 'Details'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
  }
  return sheet;
}

/**
 * Helper function to log system events
 */
function logSystemEvent(user, action, targetId, details) {
  try {
    const ss = _getOrCreateSpreadsheet();
    const logsSheet = _initializeSysLogsSheet(ss);

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
    const usersSheet = _initializeSysUsersSheet(ss);

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
    logSystemEvent(session.email, 'SESSION_EXPIRED', '', 'Session timed out after 5 minutes of inactivity');
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
  const sheet = _initializeSysUsersSheet(ss);
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
  const sheet = _initializeSysUsersSheet(ss);
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
  const sheet = _initializeSysUsersSheet(ss);

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
  const sheet = _initializeSysUsersSheet(ss);
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
  const sheet = _initializeSysUsersSheet(ss);
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
function getBudgetSubCategoryCatalog() {
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
    const particularsIndex = headers.indexOf('particulars');
    const subIndex = headers.indexOf('sub_category');
    const categoryIndex = headers.indexOf('category');
    const accountTypeIndex = headers.indexOf('account_type');
    if (particularsIndex < 0 || subIndex < 0 || categoryIndex < 0) {
      Logger.log('MASTER_DATA headers missing Particulars, Sub_Category, or Category');
      return [];
    }

    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const catalogMap = new Map();

    const excludedParticulars = new Set([
      'advance',
      'advances',
      'account payable',
      'accounts payable',
      'revaluation reserve'
    ]);

    const normalizeParticulars = (value) => String(value || '')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, ' ')
      .trim();

    data.forEach(row => {
      const particulars = String(row[particularsIndex] || '').trim();
      const subCategory = String(row[subIndex] || '').trim();
      const category = String(row[categoryIndex] || '').trim();
      const accountType = accountTypeIndex >= 0 ? String(row[accountTypeIndex] || '').trim() : '';
      if (!particulars) return;
      const normalized = normalizeParticulars(particulars);
      if (excludedParticulars.has(normalized)) return;
      if (!catalogMap.has(particulars)) {
        catalogMap.set(particulars, {
          particulars: particulars,
          subCategory: subCategory,
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
      if (a.category === b.category) {
        if (a.subCategory === b.subCategory) return a.particulars.localeCompare(b.particulars);
        return a.subCategory.localeCompare(b.subCategory);
      }
      return a.category.localeCompare(b.category);
    });
  } catch (error) {
    Logger.log('Error in getBudgetSubCategoryCatalog: ' + error.toString());
    return [];
  }
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
    const particulars = String(row[headerMap.Particulars] || '').trim();
    if (!particulars) return;
    amounts[particulars] = Number(originalValue);
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
  const sheet = _initializeDbBudgetSheet(ss);
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
    const particulars = String(item.particulars || '').trim();
    const subCategory = String(item.subCategory || '').trim();
    const category = String(item.category || '').trim();
    const accountType = String(item.accountType || '').trim();
    const amount = Number(item.amount);

    if (!particulars) throw new Error('Particulars is required.');
    if (!subCategory) throw new Error('Sub-Category is required.');
    if (!Number.isFinite(amount)) throw new Error('Invalid amount for ' + subCategory + '.');
    if (amount < 0) throw new Error('Original budget must be positive for ' + subCategory + '.');

    const row = new Array(headerCount).fill('');
    row[headerMap.Date] = dateValue;
    row[headerMap.Financial_Year] = financialYear;
    row[headerMap.Particulars] = particulars;
    row[headerMap.Sub_Category] = subCategory;
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
  const sheet = _initializeDbBudgetSheet(ss);
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
// SUPPLIERS CRUD
// ============================================================

/**
 * Get all suppliers
 */
function getSuppliers() {
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SUPPLIERS);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  const headers = ['Supplier_ID', 'Supplier_Name', 'Contact_Person', 'Phone', 'Email',
                   'Address', 'KRA_PIN', 'Bank_Name', 'Bank_Account', 'Category',
                   'Payment_Terms', 'Status', 'Created_Date', 'Created_By'];

  return data.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).filter(s => s.Supplier_ID);
}

/**
 * Get active suppliers for dropdowns
 */
function getActiveSuppliers() {
  return getSuppliers().filter(s => s.Status === 'Active');
}

/**
 * Save a new supplier
 */
function saveSupplier(data) {
  const ss = _getOrCreateSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.SUPPLIERS);
  if (!sheet) {
    sheet = _initializeSuppliersSheet(ss);
  }

  const user = getCurrentUser();
  const supplierId = _generateId('SUP');
  const now = new Date();

  const row = [
    supplierId,
    data.supplierName || '',
    data.contactPerson || '',
    data.phone || '',
    data.email || '',
    data.address || '',
    data.kraPin || '',
    data.bankName || '',
    data.bankAccount || '',
    data.category || 'General',
    data.paymentTerms || 'Net 30',
    'Active',
    now,
    user.email || ''
  ];

  sheet.appendRow(row);
  logSystemEvent(user.email, 'CREATE_SUPPLIER', supplierId, data.supplierName);

  return { success: true, supplierId: supplierId };
}

/**
 * Update supplier
 */
function updateSupplier(supplierId, data) {
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SUPPLIERS);
  if (!sheet) return { success: false, message: 'Suppliers sheet not found' };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'Supplier not found' };

  const idColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let rowIndex = -1;

  for (let i = 0; i < idColumn.length; i++) {
    if (idColumn[i][0] === supplierId) {
      rowIndex = i + 2;
      break;
    }
  }

  if (rowIndex === -1) return { success: false, message: 'Supplier not found' };

  // Update fields (columns 2-11, excluding ID, Created_Date, Created_By)
  const updates = [
    data.supplierName || '',
    data.contactPerson || '',
    data.phone || '',
    data.email || '',
    data.address || '',
    data.kraPin || '',
    data.bankName || '',
    data.bankAccount || '',
    data.category || 'General',
    data.paymentTerms || 'Net 30',
    data.status || 'Active'
  ];

  sheet.getRange(rowIndex, 2, 1, updates.length).setValues([updates]);

  const user = getCurrentUser();
  logSystemEvent(user.email, 'UPDATE_SUPPLIER', supplierId, data.supplierName);

  return { success: true };
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

  const data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  const headers = ['PO_ID', 'PO_Number', 'PO_Date', 'Financial_Year', 'Supplier_ID',
                   'Supplier_Name', 'Description', 'Total_Amount', 'Status',
                   'Requested_By', 'Requested_Date', 'Approved_By', 'Approved_Date',
                   'Delivery_Date', 'Notes'];

  let results = data.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      if (h.includes('Date') && row[i] instanceof Date) {
        obj[h] = Utilities.formatDate(row[i], Session.getScriptTimeZone(), 'yyyy-MM-dd');
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
 * Get PO with lines
 */
function getPurchaseOrderWithLines(poId) {
  const ss = _getOrCreateSpreadsheet();

  // Get PO header
  const poSheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE_ORDERS);
  if (!poSheet) return null;

  const poData = poSheet.getDataRange().getValues();
  const poHeaders = poData[0];
  let poRecord = null;

  for (let i = 1; i < poData.length; i++) {
    if (poData[i][0] === poId) {
      poRecord = {};
      poHeaders.forEach((h, idx) => poRecord[h] = poData[i][idx]);
      break;
    }
  }

  if (!poRecord) return null;

  // Get PO lines
  const linesSheet = ss.getSheetByName(CONFIG.SHEETS.PO_LINES);
  const lines = [];

  if (linesSheet && linesSheet.getLastRow() > 1) {
    const linesData = linesSheet.getDataRange().getValues();
    const linesHeaders = linesData[0];

    for (let i = 1; i < linesData.length; i++) {
      if (linesData[i][1] === poId) { // PO_ID is column 2
        const line = {};
        linesHeaders.forEach((h, idx) => line[h] = linesData[i][idx]);
        lines.push(line);
      }
    }
  }

  poRecord.lines = lines;
  return poRecord;
}

/**
 * Save a new purchase order with lines
 */
function savePurchaseOrder(data) {
  const ss = _getOrCreateSpreadsheet();
  let poSheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE_ORDERS);
  let linesSheet = ss.getSheetByName(CONFIG.SHEETS.PO_LINES);

  if (!poSheet) poSheet = _initializePurchaseOrdersSheet(ss);
  if (!linesSheet) linesSheet = _initializePoLinesSheet(ss);

  const user = getCurrentUser();
  const poId = _generateId('PO');
  const poNumber = _generateSequentialNumber('PO', poSheet, 2);
  const now = new Date();

  // Calculate total
  let totalAmount = 0;
  (data.lines || []).forEach(line => {
    totalAmount += (parseFloat(line.quantity) || 0) * (parseFloat(line.unitPrice) || 0);
  });

  // Save PO header
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
    data.notes || ''
  ];

  poSheet.appendRow(poRow);

  // Save PO lines
  (data.lines || []).forEach((line, index) => {
    const lineId = _generateId('POL');
    const qty = parseFloat(line.quantity) || 0;
    const price = parseFloat(line.unitPrice) || 0;

    const lineRow = [
      lineId,
      poId,
      index + 1,
      line.itemDescription || '',
      line.particulars || '',
      line.subCategory || '',
      line.category || '',
      qty,
      line.unit || 'Each',
      price,
      qty * price,
      0, // Received_Qty
      'Pending'
    ];

    linesSheet.appendRow(lineRow);
  });

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

  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const headers = ['GRN_ID', 'GRN_Number', 'GRN_Date', 'PO_ID', 'PO_Number',
                   'Supplier_ID', 'Supplier_Name', 'Received_By', 'Status',
                   'Invoice_Number', 'Invoice_Date', 'Invoice_Amount', 'Notes'];

  let results = data.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      if (h.includes('Date') && row[i] instanceof Date) {
        obj[h] = Utilities.formatDate(row[i], Session.getScriptTimeZone(), 'yyyy-MM-dd');
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
 * Save GRN (Goods Received Note)
 */
function saveGoodsReceivedNote(data) {
  const ss = _getOrCreateSpreadsheet();
  let grnSheet = ss.getSheetByName(CONFIG.SHEETS.GRN);
  let grnLinesSheet = ss.getSheetByName(CONFIG.SHEETS.GRN_LINES);

  if (!grnSheet) grnSheet = _initializeGrnSheet(ss);
  if (!grnLinesSheet) grnLinesSheet = _initializeGrnLinesSheet(ss);

  const user = getCurrentUser();
  const grnId = _generateId('GRN');
  const grnNumber = _generateSequentialNumber('GRN', grnSheet, 2);
  const now = new Date();

  // Save GRN header
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
    data.notes || ''
  ];

  grnSheet.appendRow(grnRow);

  // Save GRN lines and update PO lines
  const poLinesSheet = ss.getSheetByName(CONFIG.SHEETS.PO_LINES);

  (data.lines || []).forEach((line, index) => {
    const lineId = _generateId('GRNL');
    const qtyReceived = parseFloat(line.qtyReceived) || 0;
    const qtyAccepted = parseFloat(line.qtyAccepted) || qtyReceived;
    const qtyRejected = parseFloat(line.qtyRejected) || 0;

    const lineRow = [
      lineId,
      grnId,
      line.poLineId || '',
      index + 1,
      line.itemDescription || '',
      parseFloat(line.qtyOrdered) || 0,
      qtyReceived,
      qtyAccepted,
      qtyRejected,
      parseFloat(line.unitPrice) || 0,
      qtyAccepted * (parseFloat(line.unitPrice) || 0),
      line.rejectionReason || ''
    ];

    grnLinesSheet.appendRow(lineRow);

    // Update PO line received qty
    if (line.poLineId && poLinesSheet) {
      const poLinesData = poLinesSheet.getDataRange().getValues();
      for (let i = 1; i < poLinesData.length; i++) {
        if (poLinesData[i][0] === line.poLineId) {
          const currentReceived = parseFloat(poLinesData[i][11]) || 0;
          poLinesSheet.getRange(i + 1, 12).setValue(currentReceived + qtyAccepted);

          // Update line status
          const ordered = parseFloat(poLinesData[i][7]) || 0;
          const newReceived = currentReceived + qtyAccepted;
          const status = newReceived >= ordered ? 'Received' : 'Partial';
          poLinesSheet.getRange(i + 1, 13).setValue(status);
          break;
        }
      }
    }
  });

  // Update PO status if all lines received
  if (data.poId) {
    _updatePoStatusFromLines(ss, data.poId);
  }

  logSystemEvent(user.email, 'CREATE_GRN', grnId, grnNumber);

  return { success: true, grnId: grnId, grnNumber: grnNumber };
}

/**
 * Update PO status based on line statuses
 */
function _updatePoStatusFromLines(ss, poId) {
  const linesSheet = ss.getSheetByName(CONFIG.SHEETS.PO_LINES);
  const poSheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE_ORDERS);

  if (!linesSheet || !poSheet) return;

  const linesData = linesSheet.getDataRange().getValues();
  let allReceived = true;
  let anyReceived = false;

  for (let i = 1; i < linesData.length; i++) {
    if (linesData[i][1] === poId) {
      const status = linesData[i][12];
      if (status === 'Received') {
        anyReceived = true;
      } else {
        allReceived = false;
      }
    }
  }

  // Update PO status
  const poData = poSheet.getDataRange().getValues();
  for (let i = 1; i < poData.length; i++) {
    if (poData[i][0] === poId) {
      let newStatus = poData[i][8]; // Current status
      if (allReceived && anyReceived) {
        newStatus = 'Completed';
      } else if (anyReceived) {
        newStatus = 'Partial';
      }
      poSheet.getRange(i + 1, 9).setValue(newStatus);
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
 * Clean up old unused sheets (HOME, VIEW_LEDGER, VIEW_REPORTS, VIEW_RECON)
 * Run this once to remove legacy sheets
 */
function cleanupLegacySheets() {
  const ss = _getOrCreateSpreadsheet();
  const legacySheets = ['HOME', 'VIEW_LEDGER', 'VIEW_REPORTS', 'VIEW_RECON'];
  const removed = [];

  legacySheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.deleteSheet(sheet);
      removed.push(name);
      Logger.log('Removed legacy sheet: ' + name);
    }
  });

  return {
    success: true,
    message: `Removed ${removed.length} legacy sheets`,
    removed: removed
  };
}
