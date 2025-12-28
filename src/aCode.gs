/**
 * Smatika Kenya - Financial System
 * PIN-Based Authentication System with Complete Sheet Structure
 */

// Configuration
const CONFIG = {
  SPREADSHEET_NAME: "Smatika Kenya System",
  SESSION_TIMEOUT: 5 * 60 * 1000, // 5 minutes in milliseconds
  SHEETS: {
    // Part 1: Views (Front-End Displays)
    HOME: "HOME",
    VIEW_LEDGER: "VIEW_LEDGER",
    VIEW_REPORTS: "VIEW_REPORTS",
    VIEW_RECON: "VIEW_RECON",

    // Part 2: Database (Storage)
    DB_JOURNAL: "DB_JOURNAL",
    DB_BANK: "DB_BANK",
    DB_BUDGET: "DB_BUDGET",

    // Part 3: System Config (Brains)
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
        .setTitle('Smatika Kenya - Dashboard')
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
    .setTitle('Smatika Kenya - Login')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Initialize spreadsheet and all sheets
 */
function initializeSpreadsheet() {
  let ss = getOrCreateSpreadsheet();

  // Initialize all sheets in order
  initializeHomeSheet(ss);
  initializeViewLedgerSheet(ss);
  initializeViewReportsSheet(ss);
  initializeViewReconSheet(ss);
  initializeDbJournalSheet(ss);
  initializeDbBankSheet(ss);
  initializeDbBudgetSheet(ss);
  initializeMasterDataSheet(ss);
  initializeSysUsersSheet(ss);
  initializeSysLogsSheet(ss);

  return {
    success: true,
    message: "Spreadsheet initialized successfully with all 10 sheets",
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl()
  };
}

/**
 * Get or create the main spreadsheet
 */
function getOrCreateSpreadsheet() {
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
 * Helper function to create or get a sheet
 */
function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * Helper function to format header row
 */
function formatHeaderRow(sheet, headers) {
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange
    .setBackground('#4A90E2')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }

  // Freeze header row
  sheet.setFrozenRows(1);
}

// ===== PART 1: Views (Front-End Displays) =====

/**
 * Initialize HOME sheet (blank landing page)
 */
function initializeHomeSheet(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.HOME);

  // Only set this up once
  if (sheet.getLastRow() === 0) {
    sheet.getRange('A1').setValue('🏠 Welcome to Smatika Kenya System');
    sheet.getRange('A1')
      .setFontSize(18)
      .setFontWeight('bold')
      .setFontColor('#4A90E2');
  }

  return sheet;
}

/**
 * Initialize VIEW_LEDGER sheet (Search results)
 */
function initializeViewLedgerSheet(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.VIEW_LEDGER);

  if (sheet.getLastRow() === 0) {
    const headers = [
      'Date',
      'Account_Code',
      'Batch_ID',
      'Payee',
      'Category',
      'Sub_Category',
      'Description',
      'Ref_No',
      'Debit',
      'Credit',
      'Balance',
      'Receipt_Link'
    ];
    formatHeaderRow(sheet, headers);
  }

  return sheet;
}

/**
 * Initialize VIEW_REPORTS sheet (blank canvas for reports)
 */
function initializeViewReportsSheet(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.VIEW_REPORTS);
  // Keep blank - used as canvas for P&L, Balance Sheet, Cash Flow
  return sheet;
}

/**
 * Initialize VIEW_RECON sheet (blank canvas for reconciliation)
 */
function initializeViewReconSheet(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.VIEW_RECON);
  // Keep blank - used for Bank vs. Cashbook comparison
  return sheet;
}

// ===== PART 2: Database (Storage) =====

/**
 * Initialize DB_JOURNAL sheet (master transaction list)
 */
function initializeDbJournalSheet(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.DB_JOURNAL);

  if (sheet.getLastRow() === 0) {
    const headers = [
      'UUID',
      'Batch_ID',
      'Date',
      'Account_Code',
      'Payee',
      'Ref_No',
      'Type',
      'Category',
      'Sub_Category',
      'Description',
      'Debit',
      'Credit',
      'Recon_Status',
      'Receipt_URL'
    ];
    formatHeaderRow(sheet, headers);
  }

  return sheet;
}

/**
 * Initialize DB_BANK sheet (raw bank statement uploads)
 */
function initializeDbBankSheet(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.DB_BANK);

  if (sheet.getLastRow() === 0) {
    const headers = [
      'Upload_ID',
      'Account_Code',
      'Txn_Date',
      'Value_Date',
      'Bank_Ref',
      'Description',
      'Amount',
      'Balance',
      'Match_Status',
      'Matched_Batch_ID'
    ];
    formatHeaderRow(sheet, headers);
  }

  return sheet;
}

/**
 * Initialize DB_BUDGET sheet (budget audit trail)
 */
function initializeDbBudgetSheet(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.DB_BUDGET);

  if (sheet.getLastRow() === 0) {
    const headers = [
      'Entry_ID',
      'Date',
      'Type',
      'Financial_Year',
      'Category',
      'Sub_Category',
      'Amount',
      'Auth_Ref',
      'Description'
    ];
    formatHeaderRow(sheet, headers);
  }

  return sheet;
}

// ===== PART 3: System Config (Brains) =====

/**
 * Initialize MASTER_DATA sheet (dropdown lists)
 */
function initializeMasterDataSheet(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.MASTER_DATA);

  if (sheet.getLastRow() === 0) {
    const headers = [
      'Account_Codes',
      'Categories',
      'Sub_Categories',
      'Payees',
      'Projects',
      'Report_Mapping'
    ];
    formatHeaderRow(sheet, headers);

    // Add sample data
    const sampleData = [
      ['EQUITY_MAIN', 'Transport', 'Fuel', 'USAID', 'PROJECT-001', 'Revenue'],
      ['KCB_GRANT', 'Grants', 'USAID-Fund', 'Total Station', 'PROJECT-002', 'Expense'],
      ['NCBA_CURRENT', 'Salaries', 'Staff', 'Kenya Power', '', 'Asset'],
      ['MPESA_TILL', 'Utilities', 'Electricity', 'Safaricom', '', 'Liability']
    ];

    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  }

  return sheet;
}

/**
 * Initialize SYS_USERS sheet (user authentication)
 */
function initializeSysUsersSheet(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.SYS_USERS);

  if (sheet.getLastRow() === 0) {
    const headers = [
      'Email',
      'Role',
      'Status'
    ];
    formatHeaderRow(sheet, headers);

    // Add default admin user
    const defaultUser = [
      'cabdisirlam@gmail.com',
      'ADMIN',
      'ACTIVE'
    ];
    sheet.getRange(2, 1, 1, defaultUser.length).setValues([defaultUser]);

    // Add note about PIN
    sheet.getRange('D1')
      .setValue('Note: PIN is 1234 for default admin')
      .setFontColor('#718096')
      .setFontStyle('italic');
  }

  return sheet;
}

/**
 * Initialize SYS_LOGS sheet (audit trail)
 */
function initializeSysLogsSheet(ss) {
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.SYS_LOGS);

  if (sheet.getLastRow() === 0) {
    const headers = [
      'Timestamp',
      'User',
      'Action',
      'Target_ID',
      'Details'
    ];
    formatHeaderRow(sheet, headers);
  }

  return sheet;
}

/**
 * Log system action
 */
function logAction(user, action, targetId, details) {
  try {
    const ss = getOrCreateSpreadsheet();
    const logsSheet = initializeSysLogsSheet(ss);

    const timestamp = new Date();
    const logEntry = [timestamp, user, action, targetId, details];

    logsSheet.appendRow(logEntry);
  } catch (error) {
    Logger.log('Error logging action: ' + error.toString());
  }
}

/**
 * Authenticate user with email and PIN
 */
function authenticateUser(email, pin) {
  try {
    // Initialize spreadsheet if needed
    const ss = getOrCreateSpreadsheet();
    const usersSheet = initializeSysUsersSheet(ss);

    // Hardcoded PIN for now (will be enhanced later)
    const ADMIN_PIN = '1234';

    // Get all user data
    const data = usersSheet.getDataRange().getValues();

    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const userEmail = row[0];
      const userRole = row[1];
      const userStatus = row[2];

      // Check if email matches
      if (userEmail.toLowerCase() === email.toLowerCase()) {
        // Check if user is active
        if (userStatus.toUpperCase() !== 'ACTIVE') {
          logAction(email, 'LOGIN_FAILED', '', 'Account inactive');
          return {
            success: false,
            message: 'Your account is inactive. Please contact the administrator.'
          };
        }

        // Check PIN (hardcoded for now)
        if (pin !== ADMIN_PIN) {
          logAction(email, 'LOGIN_FAILED', '', 'Invalid PIN');
          return {
            success: false,
            message: 'Invalid email or PIN. Please try again.'
          };
        }

        // Create session
        const sessionData = {
          email: userEmail,
          name: userEmail.split('@')[0], // Use email prefix as name for now
          role: userRole,
          lastActivity: new Date().getTime()
        };

        PropertiesService.getUserProperties()
          .setProperty('sessionData', JSON.stringify(sessionData));

        logAction(email, 'LOGIN_SUCCESS', '', 'User logged in');

        return {
          success: true,
          user: {
            email: userEmail,
            name: sessionData.name,
            role: userRole
          }
        };
      }
    }

    logAction(email, 'LOGIN_FAILED', '', 'User not found');
    return {
      success: false,
      message: 'Invalid email or PIN. Please try again.'
    };

  } catch (error) {
    Logger.log('Authentication error: ' + error.toString());
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
    userProperties.deleteProperty('sessionData');
    logAction(session.email, 'SESSION_EXPIRED', '', 'Session timeout');
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
  const userProperties = PropertiesService.getUserProperties();
  const sessionData = userProperties.getProperty('sessionData');

  if (sessionData) {
    const session = JSON.parse(sessionData);
    logAction(session.email, 'LOGOUT', '', 'User logged out');
  }

  userProperties.deleteProperty('sessionData');
  return { success: true };
}

/**
 * Check session validity
 */
function checkSession() {
  return getCurrentUser();
}

/**
 * Include other HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Get spreadsheet URL for admin
 */
function getSpreadsheetUrl() {
  const ss = getOrCreateSpreadsheet();
  return ss.getUrl();
}
