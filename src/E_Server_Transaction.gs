/**
 * Payment modal backend helpers.
 */
function getDropdownData() {
  const payload = {
    accounts: [],
    payees: [],
    subCats: [],
    subToCatMap: {},
    categories: []
  };

  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);
  if (!sheet) return payload;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return payload;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getMasterColumns_(headers);
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  data.forEach(function(row) {
    const account = cols.accountCodes ? String(row[cols.accountCodes - 1]).trim() : '';
    const payee = cols.payees ? String(row[cols.payees - 1]).trim() : '';
    const subCat = cols.subCategory ? String(row[cols.subCategory - 1]).trim() : '';
    const category = cols.category ? String(row[cols.category - 1]).trim() : '';

    if (account) payload.accounts.push(account);
    if (payee) payload.payees.push(payee);
    if (category) payload.categories.push(category);
    if (subCat) {
      payload.subCats.push(subCat);
      if (category) payload.subToCatMap[subCat] = category;
    }
  });

  payload.accounts = _uniqueSorted_(payload.accounts);
  payload.payees = _uniqueSorted_(payload.payees);
  payload.subCats = _uniqueSorted_(payload.subCats);
  payload.categories = _uniqueSorted_(payload.categories);

  return payload;
}

function saveTransaction(data) {
  if (!data || !data.header || !data.rows || !data.rows.length) {
    throw new Error('Missing transaction data.');
  }

  const header = data.header;
  const rows = data.rows;

  const dateValue = new Date(header.date);
  if (Number.isNaN(dateValue.getTime())) throw new Error('Invalid date.');

  const accountCode = String(header.accountCode || '').trim();
  if (!accountCode) throw new Error('Bank account is required.');

  const payee = String(header.payee || '').trim();
  const refNo = String(header.refNo || '').trim();

  const ss = _getOrCreateSpreadsheet();
  const journal = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!journal) throw new Error('DB_JOURNAL not found.');

  const master = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);
  if (!master) throw new Error('MASTER_DATA not found.');

  const masterLastRow = master.getLastRow();
  const masterLastCol = master.getLastColumn();
  const masterHeaders = masterLastRow >= 1
    ? master.getRange(1, 1, 1, masterLastCol).getValues()[0].map(_normalizeHeader_)
    : [];
  const masterCols = _getMasterColumns_(masterHeaders);
  const masterData = masterLastRow > 1
    ? master.getRange(2, 1, masterLastRow - 1, masterLastCol).getValues()
    : [];
  const subMeta = _buildSubCategoryMeta_(masterData, masterCols);

  const cleanedRows = rows.map(function(row) {
    const amount = Number(row.amount || 0);
    const subCategory = String(row.subCategory || '').trim();
    const category = String(row.category || '').trim();
    const description = String(row.description || '').trim();

    if (!subCategory) throw new Error('Each line needs a sub-category.');
    if (!category) throw new Error('Each line needs a category.');
    if (amount <= 0) throw new Error('Line amount must be greater than zero.');

    const meta = subMeta[subCategory] || {};
    const accountTypeValue = meta.accountType || '';
    const reportMappingValue = meta.reportMapping || '';

    return {
      subCategory,
      category,
      description,
      amount,
      accountType: accountTypeValue,
      reportMapping: reportMappingValue
    };
  });

  const total = cleanedRows.reduce(function(sum, row) {
    return sum + row.amount;
  }, 0);

  if (total <= 0) throw new Error('Total must be greater than zero.');

  const batchId = 'TXN-' + new Date().getTime();
  const entries = cleanedRows.map(function(row) {
    return [
      Utilities.getUuid(),
      batchId,
      dateValue,
      accountCode,
      payee,
      refNo,
      row.subCategory,
      row.category,
      row.description,
      row.amount,
      0,
      row.accountType,
      row.reportMapping,
      'Unreconciled',
      ''
    ];
  });

  const startRow = journal.getLastRow() + 1;
  journal.getRange(startRow, 1, entries.length, entries[0].length).setValues(entries);

  return 'Success';
}

function addMasterItem(type, value, parent, accountType, reportMapping) {
  const trimmed = String(value || '').trim();
  if (!trimmed) throw new Error('Value is required.');

  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);
  if (!sheet) throw new Error('MASTER_DATA not found.');

  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getMasterColumns_(headers);
  const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  const typeKey = String(type || '').toLowerCase();

  if (typeKey === 'payee') {
    if (_valueExistsInColumn_(data, cols.payees, trimmed)) return;
    const row = new Array(lastCol).fill('');
    row[cols.payees - 1] = trimmed;
    sheet.appendRow(row);
    return;
  }

  if (typeKey === 'account') {
    const accountTypeValue = String(accountType || '').trim();
    const reportValue = String(reportMapping || '').trim();
    if (!accountTypeValue || !reportValue) {
      throw new Error('Account type and report mapping are required.');
    }
    if (_valueExistsInColumn_(data, cols.accountCodes, trimmed)) return;
    const row = new Array(lastCol).fill('');
    row[cols.accountCodes - 1] = trimmed;
    row[cols.accountType - 1] = accountTypeValue;
    row[cols.reportMapping - 1] = reportValue;
    sheet.appendRow(row);
    return;
  }

  if (typeKey === 'subcategory') {
    const parentCategory = String(parent || '').trim();
    const accountTypeValue = String(accountType || '').trim();
    const reportValue = String(reportMapping || '').trim();
    if (!parentCategory) throw new Error('Parent category is required.');
    if (!accountTypeValue || !reportValue) {
      throw new Error('Account type and report mapping are required.');
    }
    if (_valueExistsInColumn_(data, cols.subCategory, trimmed)) return;
    const row = new Array(lastCol).fill('');
    row[cols.subCategory - 1] = trimmed;
    row[cols.category - 1] = parentCategory;
    row[cols.accountType - 1] = accountTypeValue;
    row[cols.reportMapping - 1] = reportValue;
    sheet.appendRow(row);
    return;
  }

  throw new Error('Unknown master data type.');
}

function _normalizeHeader_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_');
}

function _getMasterColumns_(headers) {
  return {
    payees: _resolveColumn_(headers, ['payees', 'payee'], 1),
    subCategory: _resolveColumn_(headers, ['sub_category', 'subcategory'], 2),
    category: _resolveColumn_(headers, ['category'], 3),
    accountCodes: _resolveColumn_(headers, ['account_codes', 'account_code'], 4),
    accountType: _resolveColumn_(headers, ['account_type', 'accounttype'], 5),
    reportMapping: _resolveColumn_(headers, ['report_mapping', 'reportmapping'], 6)
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

function _valueExistsInColumn_(data, colIndex, value) {
  if (!colIndex) throw new Error('Column not found.');
  const normalized = String(value || '').trim().toLowerCase();
  return data.some(function(row) {
    return String(row[colIndex - 1]).trim().toLowerCase() === normalized;
  });
}

function _buildSubCategoryMeta_(data, cols) {
  const meta = {};
  data.forEach(function(row) {
    const subCat = cols.subCategory ? String(row[cols.subCategory - 1]).trim() : '';
    if (!subCat) return;
    meta[subCat] = {
      accountType: cols.accountType ? String(row[cols.accountType - 1]).trim() : '',
      reportMapping: cols.reportMapping ? String(row[cols.reportMapping - 1]).trim() : ''
    };
  });
  return meta;
}
