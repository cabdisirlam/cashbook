# Category System Implementation

## Overview
This document describes the implementation of the auto-populating category system in the Financial System cashbook application.

## Key Features

### 1. Master Data Restructuring
The `MASTER_DATA` sheet has been restructured to properly link categories with their sub-categories:

**New Structure:**
- **Category**: Main category (e.g., Transport, Salaries, Utilities)
- **Sub_Category**: Related sub-category (e.g., Fuel, Vehicle Maintenance)
- **Account_Type**: Classification for reporting (Operating Expense, Operating Income, Capital Expense, Bank)
- **Account_Codes**: Bank/account codes
- **Payees**: Vendor/donor names

**Removed:**
- Projects column (one project per entity, so not needed in master data)

### 2. Categories and Sub-Categories

#### Operating Expenses
- **Transport**: Fuel, Vehicle Maintenance, Vehicle Insurance, Vehicle Hire
- **Salaries**: Permanent Staff, Temporary Staff, Consultants, Allowances
- **Utilities**: Electricity, Water, Internet, Telephone
- **Office Rent**: Monthly Rent, Service Charge
- **Communications**: Mobile Airtime, Internet, Postage
- **Office Supplies**: Stationery, Printing, Office Equipment

#### Operating Income
- **Grants**: Donor Grants, Government Grants, Other Income

#### Bank Transactions
- **Bank**: Bank Charges, Bank Interest

#### Capital Expenses
- **Capital Expenses**: Equipment Purchase, Furniture, Vehicles, Buildings

### 3. Auto-Population Functionality

The system now includes intelligent dropdown behavior:

1. **Category Selection → Sub-Category Filtering**
   - When a user selects a category, only related sub-categories are shown
   - Example: Selecting "Transport" shows only: Fuel, Vehicle Maintenance, Vehicle Insurance, Vehicle Hire

2. **Sub-Category Selection → Auto-Populate Category**
   - When a user selects a sub-category directly, the main category auto-populates
   - Example: Selecting "Fuel" automatically sets Category to "Transport"

3. **Auto-Display Account Type**
   - When a category is selected, the account type is automatically displayed
   - Example: Selecting "Transport" shows "Operating Expense"

### 4. Helper Functions

New server-side functions in `aCode.gs`:

```javascript
getCategories()                    // Returns all unique categories
getSubCategories(category)         // Returns sub-categories for a category
getCategoryForSubCategory(subCat)  // Returns parent category for a sub-category
getAccountType(category)           // Returns account type for a category
getPayees()                        // Returns all unique payees
```

### 5. Transaction Entry Form

A new transaction entry form (`fTransaction.html`) demonstrates the dependent dropdown functionality:

**Features:**
- Date and transaction type selection
- Category dropdown (loads from MASTER_DATA)
- Sub-category dropdown (filtered by category)
- Auto-populated account type display
- Account code and payee selection
- Reference number and amount fields
- Description textarea

**Intelligent Behavior:**
- Selecting a category filters the sub-category dropdown
- Selecting a sub-category auto-fills the category if empty
- Account type displays automatically based on category
- Form validation ensures required fields are filled

### 6. Database Schema Updates

All transaction-related sheets now include the `Account_Type` column:

1. **DB_JOURNAL** (Master transactions)
   - Added: Account_Type column (position 10)

2. **VIEW_LEDGER** (Search/display)
   - Added: Account_Type column (position 7)

3. **DB_BUDGET** (Budget tracking)
   - Added: Account_Type column (position 7)

### 7. Menu Integration

The transaction menu items now open the new transaction form instead of showing placeholders:
- "New Payment (Expense)" → Opens transaction form
- "New Receipt (Income)" → Opens transaction form

## Benefits

1. **Data Integrity**: Categories and sub-categories are properly linked, preventing mismatched entries
2. **User Experience**: Auto-population reduces data entry errors and speeds up transaction recording
3. **Reporting**: Account_Type classification enables better financial reporting (P&L, Balance Sheet)
4. **Simplicity**: Removed project mapping since there's one project per entity
5. **Flexibility**: Easy to add new categories and sub-categories in MASTER_DATA sheet

## Usage Instructions

### Adding New Categories

1. Open the MASTER_DATA sheet
2. Add a new row with:
   - Category name
   - Sub-category name
   - Account_Type (choose from: Operating Expense, Operating Income, Capital Expense, Bank)
   - Optional: Account_Codes, Payees
3. Save the sheet
4. The new category/sub-category will automatically appear in all dropdowns

### Recording a Transaction

1. Click "Transactions" → "New Payment" or "New Receipt"
2. Select the **Category** (or **Sub-Category** first - category will auto-fill)
3. Notice the **Account Type** auto-displays
4. Fill in remaining fields (date, account, amount, etc.)
5. Click "Save Transaction"

## Technical Implementation

### Client-Side (HTML/JavaScript)
- Event listeners on category and sub-category dropdowns
- Calls server-side functions via `google.script.run`
- Updates UI dynamically based on selections

### Server-Side (Google Apps Script)
- Helper functions query MASTER_DATA sheet
- Return filtered data to client
- Validate relationships between categories and sub-categories

## Future Enhancements

Potential improvements:
- Add transaction saving functionality to DB_JOURNAL
- Implement search and edit transaction forms
- Add data validation rules in Google Sheets
- Create reports grouped by Account_Type
- Add budget tracking by category
