# Smatika Kenya - Financial System

A secure PIN-based authentication system for managing financial operations with role-based access control.

## 🌟 Features

### Authentication
- **PIN-based Login**: 4-digit PIN authentication
- **Email Verification**: User identification via email
- **Session Management**: 5-minute inactivity timeout
- **Auto-logout**: Automatic session expiration with warning

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
├── aCode.gs              # Backend server-side code
├── bLogin.html           # PIN login page
├── cDashboard.html       # Main dashboard
├── dMenu.html            # Role-based sidebar menu
├── appsscript.json       # Apps Script configuration
└── README.md             # This file
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
   clasp create --type webapp --title "Smatika Kenya System"
   ```

4. **Push the code**
   ```bash
   clasp push
   ```

5. **Deploy as web app**
   ```bash
   clasp deploy --description "Initial deployment"
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
1. The system will automatically create a Google Sheet named "Smatika Kenya System"
2. A "Users" sheet will be initialized with headers
3. A default admin user will be created

**Default Admin Credentials:**
- Email: `cabdisirlam@gmail.com`
- PIN: `1234`

⚠️ **Important:** Change the default admin PIN after first login!

## 📊 Users Sheet Structure

The Users sheet contains the following columns:

| Column | Type | Description |
|--------|------|-------------|
| Email | String | User's email address (login ID) |
| PIN | String | 4-digit PIN for authentication |
| Name | String | User's full name |
| Role | String | ADMIN / INPUTTER / VIEWER |
| Status | String | Active / Inactive |

### Adding New Users

1. Login as ADMIN
2. Navigate to **Admin** > **User Management** (placeholder)
3. Or manually add users to the Users sheet:
   - Open the "Smatika Kenya System" spreadsheet
   - Go to the "Users" sheet
   - Add a new row with: Email, PIN, Name, Role, Status

**Example:**
```
john@example.com | 5678 | John Doe | INPUTTER | Active
```

## 🔐 Security Features

- **Session Timeout**: 5 minutes of inactivity
- **Session Warning**: 1-minute warning before expiration
- **Activity Tracking**: Monitors mouse, keyboard, scroll, and touch events
- **Auto-logout**: Redirects to login with expiration message
- **PIN Protection**: 4-digit numeric PIN only
- **Status Check**: Only "Active" users can login

## 🎨 User Interface

### Login Page
- Clean, modern design
- Email and 4-digit PIN fields
- Session expiration warnings
- Error messages for failed login attempts

### Dashboard
- Welcome banner
- Placeholder statistics (Total Sales, Products, Customers, Revenue)
- Role-based sidebar menu
- User profile with avatar
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

### Testing

1. **Test Login**
   - Valid credentials
   - Invalid credentials
   - Inactive user

2. **Test Session**
   - Wait 4 minutes (should show warning)
   - Wait 5 minutes (should logout)
   - Perform activity (should reset timer)

3. **Test Roles**
   - Login as ADMIN (see all menu items)
   - Login as INPUTTER (no Admin section)
   - Login as VIEWER (read-only, no create/edit)

### Debugging

Enable logging in Apps Script:
```javascript
Logger.log('Debug message');
```

View logs:
```bash
clasp logs
```

## 📝 Next Steps

The current implementation provides the **skeleton** with:
- ✅ PIN-based authentication
- ✅ User management (sheet-based)
- ✅ Role-based access control
- ✅ Session management with timeout
- ✅ Menu structure (all items are placeholders)

**To implement actual features:**
1. Create additional sheets (Transactions, BankStatements, etc.)
2. Build forms for data entry (New Payment, New Receipt)
3. Implement search and edit functionality
4. Create report generation logic
5. Add Excel/PDF export functionality
6. Build reconciliation engine
7. Implement audit logging

## 🤝 Contributing

This is the initial skeleton. Future developers should:
- Follow the established naming convention
- Maintain role-based access control
- Keep the session management logic
- Update this README as features are added

## 📄 License

Proprietary - Smatika Kenya System

## 👤 Contact

For issues or questions, contact: cabdisirlam@gmail.com

---

**Version:** 1.0.0 (Skeleton)
**Last Updated:** 2025-12-28
**Status:** Development - Menu Structure Complete
