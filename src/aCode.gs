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
                     'Account_Type', 'Description', 'Ref_No', 'Debit', 'Credit', 'Balance', 'Receipt_Link'];
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
                     'Category', 'Sub_Category', 'Account_Type', 'Description', 'Debit', 'Credit',
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
                     'Account_Type', 'Amount', 'Auth_Ref', 'Description'];
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

    // Create headers - removed Projects column, renamed Report_Mapping to Account_Type
    const headers = ['Category', 'Sub_Category', 'Account_Type', 'Account_Codes', 'Payees'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeaderRow(sheet, headers.length);

    // Add comprehensive master data with proper category-subcategory relationships
    // Format: [Category, Sub_Category, Account_Type, Account_Code, Payee]
    const sampleData = [
      // Transport - Operating Expense
      ['Transport', 'Fuel', 'Operating Expense', 'EQUITY_MAIN', ''],
      ['Transport', 'Vehicle Maintenance', 'Operating Expense', '', ''],
      ['Transport', 'Vehicle Insurance', 'Operating Expense', '', ''],
      ['Transport', 'Vehicle Hire', 'Operating Expense', '', ''],

      // Salaries - Operating Expense
      ['Salaries', 'Permanent Staff', 'Operating Expense', 'NCBA_CURRENT', ''],
      ['Salaries', 'Temporary Staff', 'Operating Expense', '', ''],
      ['Salaries', 'Consultants', 'Operating Expense', '', ''],
      ['Salaries', 'Allowances', 'Operating Expense', '', ''],

      // Utilities - Operating Expense
      ['Utilities', 'Electricity', 'Operating Expense', 'MPESA_TILL', 'Kenya Power'],
      ['Utilities', 'Water', 'Operating Expense', '', ''],
      ['Utilities', 'Internet', 'Operating Expense', '', 'Safaricom'],
      ['Utilities', 'Telephone', 'Operating Expense', '', ''],

      // Office Rent - Operating Expense
      ['Office Rent', 'Monthly Rent', 'Operating Expense', '', ''],
      ['Office Rent', 'Service Charge', 'Operating Expense', '', ''],

      // Communications - Operating Expense
      ['Communications', 'Mobile Airtime', 'Operating Expense', '', ''],
      ['Communications', 'Internet', 'Operating Expense', '', ''],
      ['Communications', 'Postage', 'Operating Expense', '', ''],

      // Office Supplies - Operating Expense
      ['Office Supplies', 'Stationery', 'Operating Expense', '', ''],
      ['Office Supplies', 'Printing', 'Operating Expense', '', ''],
      ['Office Supplies', 'Office Equipment', 'Operating Expense', '', 'Total Station'],

      // Grants - Operating Income
      ['Grants', 'Donor Grants', 'Operating Income', 'KCB_GRANT', 'USAID'],
      ['Grants', 'Government Grants', 'Operating Income', '', ''],
      ['Grants', 'Other Income', 'Operating Income', '', ''],

      // Bank - Balance Sheet
      ['Bank', 'Bank Charges', 'Bank', '', ''],
      ['Bank', 'Bank Interest', 'Bank', '', ''],

      // Capital Expenses
      ['Capital Expenses', 'Equipment Purchase', 'Capital Expense', '', ''],
      ['Capital Expenses', 'Furniture', 'Capital Expense', '', ''],
      ['Capital Expenses', 'Vehicles', 'Capital Expense', '', ''],
      ['Capital Expenses', 'Buildings', 'Capital Expense', '', '']
    ];

    sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);

    // Auto-resize columns for better visibility
    sheet.autoResizeColumns(1, headers.length);
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

    const categoryData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
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

    // Get Category (col 1) and Sub_Category (col 2)
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

    // Filter by category and get unique sub-categories
    const subCategories = data
      .filter(row => row[0] === category && row[1] !== '')
      .map(row => row[1]);

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

    // Get Category (col 1) and Sub_Category (col 2)
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

    // Find the first matching row
    const matchingRow = data.find(row => row[1] === subCategory);

    return matchingRow ? matchingRow[0] : '';
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

    // Get Category (col 1) and Account_Type (col 3)
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

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

    const payeeData = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
    const uniquePayees = [...new Set(payeeData.map(row => row[0]).filter(payee => payee !== ''))];

    return uniquePayees.sort();
  } catch (error) {
    Logger.log('Error in getPayees: ' + error.toString());
    return [];
  }
}

/**
 * Include other HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
