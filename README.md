# Financial System

A secure PIN-based authentication system for managing financial operations with role-based access control and comprehensive database structure.

## 🌟 Features

### Authentication
- **PIN-based Login**: 4-digit PIN authentication (hardcoded as 1234 for now)
- **Email Verification**: User identification via email
- **Session Management**: 5-minute inactivity timeout
- **Auto-logout**: Automatic session expiration with warning
- **Audit Logging**: All login/logout events logged in SYS_LOGS

### Role-Based Access Control

Three user roles with different permissions:

#### 👤 ADMIN (Full Control)
- Full access to all features
- User management
- System settings
- All transactions and reports
- Banking and reconciliation

#### 👤 INPUTTER (Operational)
- Create and edit transactions
- Import bank statements
- Run reconciliations
- View all reports
- Cannot manage users or settings

#### 👤 VIEWER (Audit Only)
- View-only access to transactions
- Access to all reports
- System audit logs
- Reconciliation status
- Cannot create, edit, or import

### Complete Database Structure

The system automatically creates **10 sheets** organized into 3 categories:

#### PART 1: Views (Front-End Displays)
1. **HOME** - Landing page
2. **VIEW_LEDGER** - Search engine results
3. **VIEW_REPORTS** - Canvas for P&L, Balance Sheet, Cash Flow
4. **VIEW_RECON** - Bank vs. Cashbook comparison

#### PART 2: Database (Storage)
5. **DB_JOURNAL** - Master transaction list (UUID, Batch_ID, Date, Account, Payee, etc.)
6. **DB_BANK** - Raw bank statement uploads
7. **DB_BUDGET** - Budget audit trail (never delete, only add adjustments)

#### PART 3: System Config (Brains)
8. **MASTER_DATA** - Dropdown lists (Accounts, Categories, Payees, Projects)
9. **SYS_USERS** - User authentication (Email, Role, Status)
10. **SYS_LOGS** - System audit trail (Timestamp, User, Action, Details)

### Menu Structure

**📝 Transactions**
- New Payment (Expense) - Admin, Inputter
- New Receipt (Income) - Admin, Inputter
- Search / Edit Transaction - Admin, Inputter
- Search Transaction (View Only) - Viewer

**🏦 Banking & Recon** - Admin, Inputter
- Import Bank Statement
- Run Reconciliation

**📊 Financial Reports** - All Roles
- View Cashbook Report
- View P&L (Performance)
- View Balance Sheet
- View Cash Flow Statement
- View Budget Utilization
- Download as Excel
- Download as PDF

**🔎 Audit & Compliance** - Viewer
- View System Audit Logs
- View Reconciliation Status

**⚙️ Admin** - Admin Only
- User Management
- System Settings

## 📁 Project Structure

```
cashbook/
├── src/
│   ├── aCode.gs              # Backend server-side code
│   ├── bLogin.html           # PIN login page
│   ├── cDashboard.html       # Main dashboard
│   ├── dMenu.html            # Role-based sidebar menu
│   └── appsscript.json       # Apps Script configuration
├── .clasp.json.template      # Clasp configuration template
├── .gitignore                # Git ignore rules
└── README.md                 # This file
```

## 🚀 Setup Instructions

### Prerequisites
- Google Account
- [clasp](https://github.com/google/clasp) installed (`npm install -g @google/clasp`)

### Installation

1. **Clone this repository**
   ```bash
   git clone <repository-url>
   cd cashbook
   ```

2. **Login to clasp**
   ```bash
   clasp login
   ```

3. **Create a new Apps Script project**
   ```bash
   clasp create --type webapp --title "Financial System"
   ```
   This will create `.clasp.json` automatically.

4. **Push the code**
   ```bash
   clasp push
   ```

5. **Deploy as web app**
   ```bash
   clasp deploy --description "Initial deployment with 10 sheets"
   ```

6. **Open the project**
   ```bash
   clasp open
   ```

7. **Deploy the web app**
   - In the Apps Script editor, click **Deploy** > **New deployment**
   - Select type: **Web app**
   - Execute as: **User accessing the web app**
   - Who has access: **Anyone**
   - Click **Deploy**
   - Copy the web app URL

### First-Time Setup

When you first access the web app:
1. The system will automatically create a Google Sheet named **"Financial System"**
2. All **10 sheets** will be initialized with proper headers
3. Sample data will be added to **MASTER_DATA**
4. Default admin user will be created in **SYS_USERS**

**Default Admin Credentials:**
- Email: `cabdisirlam@gmail.com`
- PIN: `1234` (hardcoded for all users currently)

⚠️ **Note:** PIN functionality will be enhanced in future versions to support individual PINs per user.

## 📊 Sheet Details

### 1. HOME
- Landing page with welcome message
- Left blank for future dashboard widgets

### 2. VIEW_LEDGER (Search Results)
```
Date | Account_Code | Batch_ID | Payee | Category | Sub_Category |
Description | Ref_No | Debit | Credit | Balance | Receipt_Link
```

### 3. VIEW_REPORTS
- Blank canvas
- Scripts will draw P&L, Balance Sheet, and Cash Flow here

### 4. VIEW_RECON
- Blank canvas
- Scripts will display Bank vs. Cashbook comparison

### 5. DB_JOURNAL (Master Transactions)
```
UUID | Batch_ID | Date | Account_Code | Payee | Ref_No | Type |
Category | Sub_Category | Description | Debit | Credit |
Recon_Status | Receipt_URL
```

**Usage:**
- UUID: System-generated unique ID
- Batch_ID: Groups split transactions together
- Type: Income / Expense
- Recon_Status: Matched / Unreconciled

### 6. DB_BANK (Bank Statements)
```
Upload_ID | Account_Code | Txn_Date | Value_Date | Bank_Ref |
Description | Amount | Balance | Match_Status | Matched_Batch_ID
```

### 7. DB_BUDGET (Budget Audit Trail)
```
Entry_ID | Date | Type | Financial_Year | Category | Sub_Category |
Amount | Auth_Ref | Description
```

**Critical Rule:** Never delete rows! For budget adjustments:
```
// To move $500 from Travel to Food:
Row 1: Type=REALLOCATION, Category=Travel, Amount=-500
Row 2: Type=REALLOCATION, Category=Food, Amount=500
```

### 8. MASTER_DATA (Configuration)
```
Account_Codes | Categories | Sub_Categories | Payees | Projects | Report_Mapping
```

**Sample Data Included:**
- Account_Codes: EQUITY_MAIN, KCB_GRANT, NCBA_CURRENT, MPESA_TILL
- Categories: Transport, Grants, Salaries, Utilities
- Payees: USAID, Total Station, Kenya Power, Safaricom

### 9. SYS_USERS (Authentication)
```
Email | Role | Status | Note
```

**Roles:** ADMIN, INPUTTER, VIEWER
**Status:** ACTIVE, SUSPENDED

### 10. SYS_LOGS (Audit Trail)
```
Timestamp | User | Action | Target_ID | Details
```

**Logged Actions:**
- LOGIN_SUCCESS
- LOGIN_FAILED
- LOGOUT
- SESSION_EXPIRED
- (Future: CREATE_TRANSACTION, EDIT_TRANSACTION, DELETE_TRANSACTION, etc.)

## 🔐 Security Features

- **Session Timeout**: 5 minutes of inactivity
- **Session Warning**: 1-minute warning before expiration
- **Activity Tracking**: Monitors mouse, keyboard, scroll, and touch events
- **Auto-logout**: Redirects to login with expiration message
- **PIN Protection**: 4-digit numeric PIN (currently hardcoded as 1234)
- **Status Check**: Only "ACTIVE" users can login
- **Audit Logging**: All authentication events logged

## 🎨 User Interface

### Login Page
- Clean, modern design with professional branding
- Email and 4-digit PIN fields
- Session expiration warnings
- Error messages for failed login attempts

### Dashboard
- Welcome banner
- Placeholder statistics (Total Sales, Products, Customers, Revenue)
- Role-based sidebar menu
- User profile with avatar (first letter of email)
- Logout button
- Session timer (shows warning when < 1 minute remaining)

### Menu System
- Collapsible sidebar
- Role-based filtering
- Expandable submenus
- Icon-based navigation
- Placeholder pages for future features

## 🛠️ Development

### File Naming Convention

Files are prefixed with letters for clasp compatibility:
- `aCode.gs` - Backend code
- `bLogin.html` - Login page
- `cDashboard.html` - Dashboard
- `dMenu.html` - Menu component

All files are in the `src/` directory as specified by `.clasp.json`.

### Sheet Initialization Logic

The code uses a "get or create" pattern:
```javascript
function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}
```

This ensures sheets are created on first access and reused on subsequent calls.

### Testing

1. **Test Login**
   - Valid credentials (cabdisirlam@gmail.com / 1234)
   - Invalid credentials
   - Inactive user (change Status to SUSPENDED in SYS_USERS)

2. **Test Session**
   - Wait 4 minutes (should show warning)
   - Wait 5 minutes (should logout)
   - Perform activity (should reset timer)

3. **Test Roles**
   - Login as ADMIN (see all menu items)
   - Add INPUTTER user (no Admin section)
   - Add VIEWER user (read-only, no create/edit)

4. **Test Sheet Creation**
   - Delete the spreadsheet
   - Login again
   - Verify all 10 sheets are created
   - Check headers match specification

### Debugging

Enable logging in Apps Script:
```javascript
Logger.log('Debug message');
```

View logs:
```bash
clasp logs
```

Check SYS_LOGS sheet for audit trail:
- Login attempts
- Session expirations
- Logout events

## 📝 Next Steps

The current implementation provides the **complete skeleton** with:
- ✅ PIN-based authentication
- ✅ User management (SYS_USERS sheet)
- ✅ Role-based access control
- ✅ Session management with timeout
- ✅ Complete database structure (10 sheets)
- ✅ Audit logging (SYS_LOGS)
- ✅ Menu structure (all items are placeholders)
- ✅ Sample master data

**To implement actual features:**
1. **Transactions Module**
   - Build form for New Payment (Expense)
   - Build form for New Receipt (Income)
   - Implement transaction search and edit
   - Add UUID generation and Batch_ID grouping

2. **Banking Module**
   - Create bank statement upload parser
   - Implement reconciliation algorithm
   - Match bank transactions with journal entries

3. **Reports Module**
   - Build Cashbook Report generator
   - Create P&L Statement logic
   - Implement Balance Sheet calculation
   - Generate Cash Flow Statement
   - Add Excel/PDF export

4. **Budget Module**
   - Create budget entry form
   - Implement budget vs. actual comparison
   - Add budget utilization reports

5. **User Management**
   - Build user CRUD interface
   - Implement individual PIN management
   - Add role assignment UI

6. **Enhanced Logging**
   - Log all CRUD operations
   - Add detailed audit trail
   - Implement log viewer for admins

## 🔄 Data Entry Guidelines

### Donors vs Suppliers
- Use the **Payee** column for both
- Example Income: Payee = "USAID"
- Example Expense: Payee = "Total Station"

### Budget Adjustments
- **Never delete rows** in DB_BUDGET
- Always add new rows for changes
- Use negative amounts to decrease budget
- Example: Moving $500 from Travel to Food:
  ```
  Row: Type=REALLOCATION, Category=Travel, Amount=-500
  Row: Type=REALLOCATION, Category=Food, Amount=500
  ```

### Transaction Grouping
- Use **Batch_ID** to group related transactions
- Example: Salary payment split between Bank (-50000) and Payroll Tax (-5000)
- Both transactions share the same Batch_ID

## 🤝 Contributing

This is the initial skeleton with complete database structure. Future developers should:
- Follow the established naming convention
- Maintain role-based access control
- Keep the session management logic
- Update this README as features are added
- Never delete audit trail data (SYS_LOGS, DB_BUDGET)
- Test with all three roles (ADMIN, INPUTTER, VIEWER)

## 📄 License

Proprietary - Financial System

## 👤 Contact

For issues or questions, contact: cabdisirlam@gmail.com

---

**Version:** 2.0.0 (Complete Database Structure)
**Last Updated:** 2025-12-28
**Status:** Development - Skeleton Complete with 10 Sheets
