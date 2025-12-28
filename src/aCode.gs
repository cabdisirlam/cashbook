/**
 * Financial System
 * PIN-Based Authentication System
 */

// Configuration
const CONFIG = {
  SPREADSHEET_NAME: "Financial System",
  SESSION_TIMEOUT: 5 * 60 * 1000, // 5 minutes in milliseconds
  SHEETS: {
    // PART 1: Views (Front-End Displays)
    HOME: "HOME",
    VIEW_LEDGER: "VIEW_LEDGER",
    VIEW_REPORTS: "VIEW_REPORTS",
    VIEW_RECON: "VIEW_RECON",

    // PART 2: Database (Storage)
    DB_JOURNAL: "DB_JOURNAL",
    DB_BANK: "DB_BANK",
    DB_BUDGET: "DB_BUDGET",

    // PART 3: System Config (Brains)
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

  // Initialize all 10 sheets in order
  _initializeHomeSheet(ss);
  _initializeViewLedgerSheet(ss);
  _initializeViewReportsSheet(ss);
  _initializeViewReconSheet(ss);
  _initializeDbJournalSheet(ss);
  _initializeDbBankSheet(ss);
  _initializeDbBudgetSheet(ss);
  _initializeMasterDataSheet(ss);
  _initializeSysUsersSheet(ss);
  _initializeSysLogsSheet(ss);

  return {
    success: true,
    message: "Spreadsheet initialized successfully with 10 sheets",
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

  // Clean up VIEW_LEDGER headers
  let sheet = ss.getSheetByName(CONFIG.SHEETS.VIEW_LEDGER);
  if (sheet) {
    const headers = ['UUID', 'Batch_ID', 'Date', 'Account_Code', 'Payee', 'Ref_No',
                     'Sub_Category', 'Category', 'Description', 'Debit', 'Credit',
                     'Account_Type', 'Report_Mapping', 'Recon_Status', 'Receipt_URL'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
    updatedSheets.push('VIEW_LEDGER');
    Logger.log('✓ VIEW_LEDGER headers cleaned');
  }

  // Clean up DB_JOURNAL headers
  sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (sheet) {
    const headers = ['UUID', 'Batch_ID', 'Date', 'Account_Code', 'Payee', 'Ref_No',
                     'Sub_Category', 'Category', 'Description', 'Debit', 'Credit',
                     'Account_Type', 'Report_Mapping', 'Recon_Status', 'Receipt_URL'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
    updatedSheets.push('DB_JOURNAL');
    Logger.log('✓ DB_JOURNAL headers cleaned');
  }

  // Clean up DB_BANK headers
  sheet = ss.getSheetByName(CONFIG.SHEETS.DB_BANK);
  if (sheet) {
    const headers = ['Account_Code', 'Txn_Date', 'Value_Date', 'Bank_Ref',
                     'Description', 'Debit', 'Credit', 'Balance', 'Match_Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
    updatedSheets.push('DB_BANK');
    Logger.log('✓ DB_BANK headers cleaned');
  }

  // Clean up DB_BUDGET headers
  sheet = ss.getSheetByName(CONFIG.SHEETS.DB_BUDGET);
  if (sheet) {
    const headers = ['Date', 'Type', 'Financial_Year', 'Sub_Category', 'Category',
                     'Account_Type', 'Amount', 'Auth_Ref', 'Description'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
    updatedSheets.push('DB_BUDGET');
    Logger.log('✓ DB_BUDGET headers cleaned');
  }

  // Clean up MASTER_DATA headers
  sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);
  if (sheet) {
    const headers = ['Payees', 'Sub_Category', 'Category', 'Account_Codes', 'Account_Type', 'Report_Mapping'];
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
 * PART 1: Views (Front-End Displays)
 */

// 1. HOME Sheet
function _initializeHomeSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.HOME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.HOME, 0); // First sheet
    sheet.getRange('A1').setValue('Welcome to Financial System');
    sheet.getRange('A1').setFontSize(24).setFontWeight('bold');
    sheet.getRange('A2').setValue('This is your landing page for dashboard widgets');
  }
  return sheet;
}

// 2. VIEW_LEDGER Sheet
function _initializeViewLedgerSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.VIEW_LEDGER);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.VIEW_LEDGER);
    const headers = ['UUID', 'Batch_ID', 'Date', 'Account_Code', 'Payee', 'Ref_No',
                     'Sub_Category', 'Category', 'Description', 'Debit', 'Credit',
                     'Account_Type', 'Report_Mapping', 'Recon_Status', 'Receipt_URL'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
  }
  return sheet;
}

// 3. VIEW_REPORTS Sheet
function _initializeViewReportsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.VIEW_REPORTS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.VIEW_REPORTS);
    sheet.getRange('A1').setValue('Financial Reports Canvas');
    sheet.getRange('A1').setFontSize(18).setFontWeight('bold');
    sheet.getRange('A2').setValue('Scripts will draw P&L, Balance Sheet, and Cash Flow here');
  }
  return sheet;
}

// 4. VIEW_RECON Sheet
function _initializeViewReconSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.VIEW_RECON);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.VIEW_RECON);
    sheet.getRange('A1').setValue('Bank Reconciliation View');
    sheet.getRange('A1').setFontSize(18).setFontWeight('bold');
    sheet.getRange('A2').setValue('Scripts will display Bank vs. Cashbook comparison here');
  }
  return sheet;
}

/**
 * PART 2: Database (Storage)
 */

// 5. DB_JOURNAL Sheet
function _initializeDbJournalSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.DB_JOURNAL);
    const headers = ['UUID', 'Batch_ID', 'Date', 'Account_Code', 'Payee', 'Ref_No',
                     'Sub_Category', 'Category', 'Description', 'Debit', 'Credit',
                     'Account_Type', 'Report_Mapping', 'Recon_Status', 'Receipt_URL'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
  }
  return sheet;
}

// 6. DB_BANK Sheet
function _initializeDbBankSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DB_BANK);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.DB_BANK);
    const headers = ['Account_Code', 'Txn_Date', 'Value_Date', 'Bank_Ref',
                     'Description', 'Debit', 'Credit', 'Balance', 'Match_Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);
  }
  return sheet;
}

// 7. DB_BUDGET Sheet
function _initializeDbBudgetSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DB_BUDGET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.DB_BUDGET);
    const headers = ['Date', 'Type', 'Financial_Year', 'Sub_Category', 'Category',
                     'Account_Type', 'Amount', 'Auth_Ref', 'Description'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);

    // Add a note about never deleting rows
    sheet.getRange('A2').setNote('CRITICAL RULE: Never delete rows! For budget adjustments, add new rows with negative amounts.');
  }
  return sheet;
}

/**
 * PART 3: System Config (Brains)
 */

// 8. MASTER_DATA Sheet
function _initializeMasterDataSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.MASTER_DATA);

    // Create headers - All data will come from user input
    const headers = ['Payees', 'Sub_Category', 'Category', 'Account_Codes', 'Account_Type', 'Report_Mapping'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _formatHeaderRow(sheet, headers.length);

    // Auto-resize columns for better visibility
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

// 9. SYS_USERS Sheet
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

// 10. SYS_LOGS Sheet
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
    decimals: props.getProperty('decimals') || '0.00'
  };
}

function saveSystemSettings(settings) {
  _requireAdmin();
  if (!settings) throw new Error('Missing settings.');

  const systemName = String(settings.systemName || '').trim();
  const entityName = String(settings.entityName || '').trim();
  const currency = String(settings.currency || '').trim().toUpperCase();
  const decimals = String(settings.decimals || '').trim();

  if (!systemName) throw new Error('System name is required.');
  if (!entityName) throw new Error('Entity name is required.');
  if (!currency) throw new Error('Currency is required.');
  if (!/^\d\.\d{2}$/.test(decimals)) throw new Error('Decimals must be in 0.00 format.');

  const props = PropertiesService.getScriptProperties();
  props.setProperty('systemName', systemName);
  props.setProperty('entityName', entityName);
  props.setProperty('currency', currency);
  props.setProperty('decimals', decimals);

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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);

    if (!sheet) {
      Logger.log('MASTER_DATA sheet not found');
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; // No data

    const categoryData = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
    const uniqueCategories = [...new Set(categoryData.map(row => row[0]).filter(cat => cat !== ''))];

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

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);

    if (!sheet) {
      Logger.log('MASTER_DATA sheet not found');
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; // No data

    // Get Sub_Category (col 2) and Category (col 3)
    const data = sheet.getRange(2, 2, lastRow - 1, 2).getValues();

    // Filter by category and get unique sub-categories
    const subCategories = data
      .filter(row => row[1] === category && row[0] !== '')
      .map(row => row[0]);

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

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);

    if (!sheet) {
      Logger.log('MASTER_DATA sheet not found');
      return '';
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return ''; // No data

    // Get Sub_Category (col 2) and Category (col 3)
    const data = sheet.getRange(2, 2, lastRow - 1, 2).getValues();

    // Find the first matching row
    const matchingRow = data.find(row => row[0] === subCategory);

    return matchingRow ? matchingRow[1] : '';
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

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);

    if (!sheet) {
      Logger.log('MASTER_DATA sheet not found');
      return '';
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return ''; // No data

    // Get Category (col 3) and Account_Type (col 5)
    const data = sheet.getRange(2, 3, lastRow - 1, 3).getValues();

    // Find the first matching row
    const matchingRow = data.find(row => row[0] === category);

    return matchingRow ? matchingRow[2] : '';
  } catch (error) {
    Logger.log('Error in getAccountType: ' + error.toString());
    return '';
  }
}

/**
 * Get all unique payees from MASTER_DATA
 * @returns {Array} Array of unique payee names
 */
function getPayees() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
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
