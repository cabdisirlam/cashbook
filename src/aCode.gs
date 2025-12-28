/**
 * Smatika Kenya - Financial System
 * PIN-Based Authentication System
 */

// Configuration
const CONFIG = {
  SPREADSHEET_NAME: "Smatika Kenya System",
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
 * Initialize spreadsheet and sheets
 */
function initializeSpreadsheet() {
  let ss = getOrCreateSpreadsheet();

  // Initialize all 10 sheets in order
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
    message: "Spreadsheet initialized successfully with 10 sheets",
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
 * Helper function to format header row
 */
function formatHeaderRow(sheet, numColumns) {
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
function initializeHomeSheet(ss) {
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
function initializeViewLedgerSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.VIEW_LEDGER);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.VIEW_LEDGER);
    const headers = ['Date', 'Account_Code', 'Batch_ID', 'Payee', 'Category', 'Sub_Category',
                     'Description', 'Ref_No', 'Debit', 'Credit', 'Balance', 'Receipt_Link'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeaderRow(sheet, headers.length);
  }
  return sheet;
}

// 3. VIEW_REPORTS Sheet
function initializeViewReportsSheet(ss) {
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
function initializeViewReconSheet(ss) {
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
function initializeDbJournalSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.DB_JOURNAL);
    const headers = ['UUID', 'Batch_ID', 'Date', 'Account_Code', 'Payee', 'Ref_No', 'Type',
                     'Category', 'Sub_Category', 'Description', 'Debit', 'Credit',
                     'Recon_Status', 'Receipt_URL'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeaderRow(sheet, headers.length);
  }
  return sheet;
}

// 6. DB_BANK Sheet
function initializeDbBankSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DB_BANK);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.DB_BANK);
    const headers = ['Upload_ID', 'Account_Code', 'Txn_Date', 'Value_Date', 'Bank_Ref',
                     'Description', 'Amount', 'Balance', 'Match_Status', 'Matched_Batch_ID'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeaderRow(sheet, headers.length);
  }
  return sheet;
}

// 7. DB_BUDGET Sheet
function initializeDbBudgetSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DB_BUDGET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.DB_BUDGET);
    const headers = ['Entry_ID', 'Date', 'Type', 'Financial_Year', 'Category', 'Sub_Category',
                     'Amount', 'Auth_Ref', 'Description'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeaderRow(sheet, headers.length);

    // Add a note about never deleting rows
    sheet.getRange('A2').setNote('CRITICAL RULE: Never delete rows! For budget adjustments, add new rows with negative amounts.');
  }
  return sheet;
}

/**
 * PART 3: System Config (Brains)
 */

// 8. MASTER_DATA Sheet
function initializeMasterDataSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.MASTER_DATA);

    // Create headers for each section
    const headers = ['Account_Codes', 'Categories', 'Sub_Categories', 'Payees', 'Projects', 'Report_Mapping'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeaderRow(sheet, headers.length);

    // Add sample data
    const sampleData = [
      ['EQUITY_MAIN', 'Transport', 'Fuel', 'USAID', 'Project_A', 'Expense'],
      ['KCB_GRANT', 'Grants', 'Vehicle', 'Total Station', 'Project_B', 'Income'],
      ['NCBA_CURRENT', 'Salaries', 'Office Supplies', 'Kenya Power', 'Project_C', 'Expense'],
      ['MPESA_TILL', 'Utilities', 'Maintenance', 'Safaricom', '', 'Expense'],
      ['', 'Office Rent', '', '', '', ''],
      ['', 'Communications', '', '', '', '']
    ];

    sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
  }
  return sheet;
}

// 9. SYS_USERS Sheet
function initializeSysUsersSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.SYS_USERS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.SYS_USERS);

    // Set headers
    const headers = ['Email', 'PIN', 'Name', 'Role', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeaderRow(sheet, headers.length);

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
function initializeSysLogsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.SYS_LOGS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.SYS_LOGS);
    const headers = ['Timestamp', 'User', 'Action', 'Target_ID', 'Details'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeaderRow(sheet, headers.length);
  }
  return sheet;
}

/**
 * Helper function to log system events
 */
function logSystemEvent(user, action, targetId, details) {
  try {
    const ss = getOrCreateSpreadsheet();
    const logsSheet = initializeSysLogsSheet(ss);

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
    const ss = getOrCreateSpreadsheet();
    const usersSheet = initializeSysUsersSheet(ss);

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
 * Include other HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
