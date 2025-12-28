/**
 * Server-side transaction logic for split payment form.
 */
function getFormInitData() {
  const payload = {
    accounts: [],
    payees: [],
    subCats: [],
    subToCatMap: {},
    catToTypeMap: {}
  };

  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);
  if (!sheet) return payload;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return payload;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getMasterDataColumns_(headers);

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const accounts = [];
  const payees = [];
  const subCats = [];

  data.forEach(function(row) {
    const category = cols.category ? String(row[cols.category - 1]).trim() : '';
    const subCategory = cols.subCategory ? String(row[cols.subCategory - 1]).trim() : '';
    const accountType = cols.accountType ? String(row[cols.accountType - 1]).trim() : '';
    const accountCode = cols.accountCodes ? String(row[cols.accountCodes - 1]).trim() : '';
    const payee = cols.payees ? String(row[cols.payees - 1]).trim() : '';

    if (category && !payload.catToTypeMap[category]) {
      payload.catToTypeMap[category] = accountType || '';
    }

    if (subCategory) {
      subCats.push(subCategory);
      if (category) {
        payload.subToCatMap[subCategory] = category;
      }
    }

    if (accountCode) accounts.push(accountCode);
    if (payee) payees.push(payee);
  });

  payload.accounts = _uniqueSorted_(accounts);
  payload.payees = _uniqueSorted_(payees);
  payload.subCats = _uniqueSorted_(subCats);

  return payload;
}

function getPaymentFormHtml() {
  return HtmlService.createHtmlOutputFromFile('F_Transaction').getContent();
}

function addMasterItem(type, value, parent) {
  const trimmed = String(value || '').trim();
  if (!trimmed) throw new Error('Value is required.');

  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);
  if (!sheet) throw new Error('MASTER_DATA sheet not found.');

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getMasterDataColumns_(headers);
  const lastRow = sheet.getLastRow();
  const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  const typeKey = String(type || '').toLowerCase();

  if (typeKey === 'account') {
    _addUniqueToColumn_(sheet, data, cols.accountCodes, trimmed, lastCol);
    return { success: true };
  }

  if (typeKey === 'payee') {
    _addUniqueToColumn_(sheet, data, cols.payees, trimmed, lastCol);
    return { success: true };
  }

  if (typeKey === 'subcategory' || typeKey === 'sub_category') {
    const parentCategory = String(parent || '').trim();
    if (!parentCategory) throw new Error('Parent category is required.');

    const accountType = _lookupAccountType_(data, cols.category, cols.accountType, parentCategory);
    const row = new Array(lastCol).fill('');
    row[cols.category - 1] = parentCategory;
    row[cols.subCategory - 1] = trimmed;
    if (cols.accountType) {
      row[cols.accountType - 1] = accountType;
    }

    sheet.appendRow(row);
    return { success: true };
  }

  throw new Error('Unknown master data type.');
}

function saveTransaction(header, rows) {
  if (!header) throw new Error('Missing header.');
  if (!rows || !rows.length) throw new Error('Add at least one line item.');

  const dateValue = new Date(header.date);
  if (Number.isNaN(dateValue.getTime())) throw new Error('Invalid date.');

  const accountCode = String(header.accountCode || '').trim();
  if (!accountCode) throw new Error('Account code is required.');

  const payee = String(header.payee || '').trim();
  const refNo = String(header.refNo || '').trim();

  const cleanRows = rows.map(function(row) {
    const amount = Number(row.amount || 0);
    const subCategory = String(row.subCategory || '').trim();
    const category = String(row.category || '').trim();
    const description = String(row.description || '').trim();

    if (!subCategory) throw new Error('Each line needs a sub-category.');
    if (!category) throw new Error('Each line needs a category.');
    if (amount <= 0) throw new Error('Each line amount must be greater than zero.');

    return {
      subCategory: subCategory,
      category: category,
      description: description,
      amount: amount
    };
  });

  const total = cleanRows.reduce(function(sum, row) {
    return sum + row.amount;
  }, 0);

  if (total <= 0) throw new Error('Total must be greater than zero.');

  const ss = _getOrCreateSpreadsheet();
  const journal = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!journal) throw new Error('DB_JOURNAL sheet not found.');

  const batchId = 'TXN-' + new Date().getTime();
  const entries = cleanRows.map(function(row) {
    return [
      Utilities.getUuid(),
      batchId,
      dateValue,
      accountCode,
      payee,
      refNo,
      'Expense',
      row.category,
      row.subCategory,
      row.description,
      row.amount,
      0,
      'Unreconciled',
      ''
    ];
  });

  const startRow = journal.getLastRow() + 1;
  journal.getRange(startRow, 1, entries.length, entries[0].length).setValues(entries);

  try {
    const user = getCurrentUser();
    const actor = user && user.authenticated ? user.email : 'system';
    logSystemEvent(actor, 'CREATE_PAYMENT', batchId, 'Lines: ' + cleanRows.length + ', Total: ' + total);
  } catch (error) {
    Logger.log('Log failure: ' + error.toString());
  }

  return {
    success: true,
    batchId: batchId,
    lines: entries.length,
    total: total
  };
}

function _normalizeHeader_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_');
}

function _getMasterDataColumns_(headers) {
  return {
    category: _resolveColumn_(headers, ['category'], 1),
    subCategory: _resolveColumn_(headers, ['sub_category', 'subcategory'], 2),
    accountType: _resolveColumn_(headers, ['account_type', 'accounttype'], 3),
    accountCodes: _resolveColumn_(headers, ['account_codes', 'account_code', 'accountcodes'], 4),
    payees: _resolveColumn_(headers, ['payees', 'payee'], 5)
  };
}

function _resolveColumn_(headers, names, fallback) {
  for (let i = 0; i < names.length; i++) {
    const idx = headers.indexOf(names[i]);
    if (idx >= 0) return idx + 1;
  }
  return fallback;
}

function _uniqueSorted_(items) {
  return [...new Set(items.filter(Boolean))].sort();
}

function _addUniqueToColumn_(sheet, data, colIndex, value, lastCol) {
  if (!colIndex) throw new Error('Column not found.');

  const normalized = String(value || '').trim();
  if (!normalized) throw new Error('Value is required.');

  const existing = data.some(function(row) {
    return String(row[colIndex - 1]).trim().toLowerCase() === normalized.toLowerCase();
  });

  if (existing) return;

  let targetRow = -1;
  for (let i = 0; i < data.length; i++) {
    if (!String(data[i][colIndex - 1]).trim()) {
      targetRow = i + 2;
      break;
    }
  }

  if (targetRow > 0) {
    sheet.getRange(targetRow, colIndex).setValue(normalized);
  } else {
    const row = new Array(lastCol).fill('');
    row[colIndex - 1] = normalized;
    sheet.appendRow(row);
  }
}

function _lookupAccountType_(data, categoryCol, accountTypeCol, category) {
  if (!categoryCol || !accountTypeCol) return '';
  const match = data.find(function(row) {
    return String(row[categoryCol - 1]).trim() === category;
  });
  return match ? String(match[accountTypeCol - 1]).trim() : '';
}
