/**
 * Payment modal backend helpers.
 */
function getDropdownData() {
  const payload = {
    accounts: [],
    payees: [],
    subCats: [],
    subToCatMap: {},
    categories: [],
    financialYears: []
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
    const financialYear = cols.financialYear ? String(row[cols.financialYear - 1]).trim() : '';

    if (account) payload.accounts.push(account);
    if (payee) payload.payees.push(payee);
    if (category) payload.categories.push(category);
    if (financialYear) payload.financialYears.push(financialYear);
    if (subCat) {
      payload.subCats.push(subCat);
      if (category) payload.subToCatMap[subCat] = category;
    }
  });

  payload.accounts = _uniqueSorted_(payload.accounts);
  payload.payees = _uniqueSorted_(payload.payees);
  payload.subCats = _uniqueSorted_(payload.subCats);
  payload.categories = _uniqueSorted_(payload.categories);
  payload.financialYears = _uniqueSorted_(payload.financialYears);

  return payload;
}

function saveTransaction(data) {
  if (!data || !data.header || !data.rows || !data.rows.length) {
    throw new Error('Missing transaction data.');
  }

  const header = data.header;
  const rows = data.rows;
  const type = String(data.type || 'Payment').trim();
  const isReceipt = type.toLowerCase() === 'receipt';

  const dateValue = new Date(header.date);
  if (Number.isNaN(dateValue.getTime())) throw new Error('Invalid date.');

  const financialYear = String(header.financialYear || '').trim();
  const accountCode = String(header.accountCode || '').trim();
  if (!financialYear) throw new Error('Financial year is required.');
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
    const debitValue = isReceipt ? 0 : row.amount;
    const creditValue = isReceipt ? row.amount : 0;
    return [
      Utilities.getUuid(),
      batchId,
      dateValue,
      financialYear,
      accountCode,
      payee,
      refNo,
      row.subCategory,
      row.category,
      row.description,
      debitValue,
      creditValue,
      row.accountType,
      row.reportMapping,
      'Unreconciled',
      ''
    ];
  });

  const startRow = journal.getLastRow() + 1;
  journal.getRange(startRow, 1, entries.length, entries[0].length).setValues(entries);

  logSystemEventSafe(
    'CREATE_JOURNAL',
    batchId,
    'Type: ' + type + ', Rows: ' + entries.length + ', Total: ' + total
  );

  return 'Success';
}

function searchJournal(criteria) {
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) throw new Error('DB_JOURNAL not found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getJournalColumns_(headers);
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const query = String(criteria && criteria.query || '').trim().toLowerCase();
  const refOrId = String(criteria && criteria.refOrId || '').trim().toLowerCase();
  const categoryFilter = String(criteria && criteria.category || '').trim();
  const payeeFilter = String(criteria && criteria.payee || '').trim().toLowerCase();
  const startDate = _parseDate_(criteria && criteria.startDate);
  const endDate = _parseDate_(criteria && criteria.endDate);
  const hasFilters = Boolean(query || refOrId || categoryFilter || payeeFilter || startDate || endDate);

  const results = [];

  data.forEach(function(row, index) {
    const rowDate = cols.date ? row[cols.date - 1] : null;
    const dateValue = rowDate instanceof Date ? rowDate : _parseDate_(rowDate);
    if (startDate && dateValue && dateValue < startDate) return;
    if (endDate && dateValue && dateValue > endDate) return;

    const category = cols.category ? String(row[cols.category - 1]).trim() : '';
    if (categoryFilter && category !== categoryFilter) return;

    const payee = cols.payee ? String(row[cols.payee - 1]).trim() : '';
    if (payeeFilter && !payee.toLowerCase().includes(payeeFilter)) return;

    const uuid = cols.uuid ? String(row[cols.uuid - 1]).trim() : '';
    const refNo = cols.refNo ? String(row[cols.refNo - 1]).trim() : '';
    if (refOrId && !(uuid.toLowerCase().includes(refOrId) || refNo.toLowerCase().includes(refOrId))) return;

    const haystack = [
      cols.accountCode ? row[cols.accountCode - 1] : '',
      payee,
      category,
      cols.subCategory ? row[cols.subCategory - 1] : '',
      cols.description ? row[cols.description - 1] : '',
      refNo
    ].map(String).join(' ').toLowerCase();

    if (query && !haystack.includes(query)) return;

    results.push({
      rowId: index + 2,
      uuid: uuid,
      batchId: cols.batchId ? row[cols.batchId - 1] : '',
      date: dateValue ? _formatDate_(dateValue) : '',
      accountCode: cols.accountCode ? row[cols.accountCode - 1] : '',
      payee: payee,
      refNo: refNo,
      subCategory: cols.subCategory ? row[cols.subCategory - 1] : '',
      category: category,
      description: cols.description ? row[cols.description - 1] : '',
      debit: cols.debit ? row[cols.debit - 1] : '',
      credit: cols.credit ? row[cols.credit - 1] : ''
    });
  });

  if (!hasFilters) {
    return results.slice(-5).reverse();
  }

  return results.slice(0, 200);
}

function getJournalRow(rowId) {
  const row = Number(rowId);
  if (!row || row < 2) throw new Error('Invalid row.');
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) throw new Error('DB_JOURNAL not found.');

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getJournalColumns_(headers);
  const values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  const dateValue = values[cols.date - 1];

  return {
    rowId: row,
    uuid: values[cols.uuid - 1] || '',
    batchId: values[cols.batchId - 1] || '',
    date: dateValue instanceof Date ? _formatDate_(dateValue) : '',
    financialYear: cols.financialYear ? values[cols.financialYear - 1] : '',
    accountCode: values[cols.accountCode - 1] || '',
    payee: values[cols.payee - 1] || '',
    refNo: values[cols.refNo - 1] || '',
    subCategory: values[cols.subCategory - 1] || '',
    category: values[cols.category - 1] || '',
    description: values[cols.description - 1] || '',
    debit: values[cols.debit - 1] || '',
    credit: values[cols.credit - 1] || '',
    accountType: values[cols.accountType - 1] || '',
    reportMapping: values[cols.reportMapping - 1] || '',
    reconStatus: values[cols.reconStatus - 1] || '',
    receiptUrl: values[cols.receiptUrl - 1] || ''
  };
}

function updateJournal(payload) {
  if (!payload) throw new Error('Missing payload.');
  const row = Number(payload.rowId);
  if (!row || row < 2) throw new Error('Invalid row.');

  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) throw new Error('DB_JOURNAL not found.');

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getJournalColumns_(headers);
  const existing = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  _storeUndoAction_('update', row, existing);

  const updated = existing.slice();
  updated[cols.date - 1] = _parseDate_(payload.date) || existing[cols.date - 1];
  if (cols.financialYear) {
    updated[cols.financialYear - 1] = payload.financialYear || '';
  }
  updated[cols.accountCode - 1] = payload.accountCode || '';
  updated[cols.payee - 1] = payload.payee || '';
  updated[cols.refNo - 1] = payload.refNo || '';
  updated[cols.subCategory - 1] = payload.subCategory || '';
  updated[cols.category - 1] = payload.category || '';
  updated[cols.description - 1] = payload.description || '';
  updated[cols.debit - 1] = Number(payload.debit || 0);
  updated[cols.credit - 1] = Number(payload.credit || 0);
  updated[cols.accountType - 1] = payload.accountType || '';
  updated[cols.reportMapping - 1] = payload.reportMapping || '';
  updated[cols.reconStatus - 1] = payload.reconStatus || '';
  updated[cols.receiptUrl - 1] = payload.receiptUrl || '';

  sheet.getRange(row, 1, 1, lastCol).setValues([updated]);

  try {
    const user = getCurrentUser();
    const actor = user && user.authenticated ? user.email : 'system';
    logSystemEvent(actor, 'UPDATE_JOURNAL', String(updated[cols.uuid - 1] || ''), 'Row ' + row);
  } catch (error) {
    Logger.log('Log failure: ' + error.toString());
  }

  return true;
}

function deleteJournal(rowId) {
  const row = Number(rowId);
  if (!row || row < 2) throw new Error('Invalid row.');

  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) throw new Error('DB_JOURNAL not found.');

  const lastCol = sheet.getLastColumn();
  const existing = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  _storeUndoAction_('delete', row, existing);

  sheet.deleteRow(row);

  try {
    const user = getCurrentUser();
    const actor = user && user.authenticated ? user.email : 'system';
    logSystemEvent(actor, 'DELETE_JOURNAL', String(existing[0] || ''), 'Row ' + row);
  } catch (error) {
    Logger.log('Log failure: ' + error.toString());
  }

  return true;
}

function undoLastJournalAction() {
  const props = PropertiesService.getUserProperties();
  const raw = props.getProperty('lastJournalAction');
  if (!raw) throw new Error('No action to undo.');
  const action = JSON.parse(raw);

  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) throw new Error('DB_JOURNAL not found.');

  if (action.type === 'delete') {
    const rowIndex = Math.min(action.rowId, sheet.getLastRow() + 1);
    sheet.insertRowsBefore(rowIndex, 1);
    sheet.getRange(rowIndex, 1, 1, action.values.length).setValues([action.values]);
  } else if (action.type === 'update') {
    sheet.getRange(action.rowId, 1, 1, action.values.length).setValues([action.values]);
  } else {
    throw new Error('Unknown action.');
  }

  logSystemEventSafe('UNDO_JOURNAL', '', 'Undo ' + action.type + ' on row ' + action.rowId);

  props.deleteProperty('lastJournalAction');
  return 'Undo completed.';
}

function exportJournal(criteria) {
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) throw new Error('DB_JOURNAL not found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { csv: '', filename: '' };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const startDate = _parseDate_(criteria && criteria.startDate);
  const endDate = _parseDate_(criteria && criteria.endDate);

  const headerMap = headers.map(_normalizeHeader_);
  const dateIndex = headerMap.indexOf('date');

  const filtered = data.filter(function(row) {
    if (dateIndex < 0) return true;
    const rowDate = row[dateIndex];
    const dateValue = rowDate instanceof Date ? rowDate : _parseDate_(rowDate);
    if (!dateValue) return true;
    if (startDate && dateValue < startDate) return false;
    if (endDate && dateValue > endDate) return false;
    return true;
  });

  if (!filtered.length) {
    logSystemEventSafe('EXPORT_JOURNAL', '', 'No data to export.');
    return { csv: '', filename: '' };
  }

  const output = [headers].concat(filtered).map(function(row) {
    return row.map(function(cell) {
      if (cell instanceof Date) {
        return _formatDate_(cell);
      }
      const value = String(cell == null ? '' : cell);
      return '"' + value.replace(/"/g, '""') + '"';
    }).join(',');
  }).join('\n');

  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  logSystemEventSafe('EXPORT_JOURNAL', '', 'Rows: ' + filtered.length);
  return {
    csv: output,
    filename: 'journal_export_' + stamp + '.csv'
  };
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
    const targetRow = _findRowForInsert_(data, cols.payees, []);
    _writeRowUpdate_(sheet, data, targetRow, lastCol, {
      [cols.payees]: trimmed
    });
    logSystemEventSafe('CREATE_MASTER_DATA', trimmed, 'Type: payee');
    return;
  }

  if (typeKey === 'financialyear') {
    if (_valueExistsInColumn_(data, cols.financialYear, trimmed)) return;
    const targetRow = _findRowForInsert_(data, cols.financialYear, [cols.subCategory, cols.accountCodes]);
    _writeRowUpdate_(sheet, data, targetRow, lastCol, {
      [cols.financialYear]: trimmed
    });
    logSystemEventSafe('CREATE_MASTER_DATA', trimmed, 'Type: financialYear');
    return;
  }

  if (typeKey === 'account') {
    const accountTypeValue = String(accountType || '').trim();
    const reportValue = String(reportMapping || '').trim();
    if (!accountTypeValue || !reportValue) {
      throw new Error('Account type and report mapping are required.');
    }
    if (_valueExistsInColumn_(data, cols.accountCodes, trimmed)) return;
    const targetRow = _findRowForInsert_(data, cols.accountCodes, [cols.subCategory]);
    _writeRowUpdate_(sheet, data, targetRow, lastCol, {
      [cols.accountCodes]: trimmed,
      [cols.accountType]: accountTypeValue,
      [cols.reportMapping]: reportValue
    });
    logSystemEventSafe('CREATE_MASTER_DATA', trimmed, 'Type: account');
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
    const targetRow = _findRowForInsert_(data, cols.subCategory, [cols.accountCodes]);
    _writeRowUpdate_(sheet, data, targetRow, lastCol, {
      [cols.subCategory]: trimmed,
      [cols.category]: parentCategory,
      [cols.accountType]: accountTypeValue,
      [cols.reportMapping]: reportValue
    });
    logSystemEventSafe('CREATE_MASTER_DATA', trimmed, 'Type: subCategory, Category: ' + parentCategory);
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
    reportMapping: _resolveColumn_(headers, ['report_mapping', 'reportmapping'], 6),
    financialYear: _resolveColumn_(headers, ['financial_year', 'financialyear'], 7)
  };
}

function _getJournalColumns_(headers) {
  return {
    uuid: _resolveColumn_(headers, ['uuid'], 1),
    batchId: _resolveColumn_(headers, ['batch_id'], 2),
    date: _resolveColumn_(headers, ['date'], 3),
    financialYear: _resolveColumn_(headers, ['financial_year', 'financialyear'], 4),
    accountCode: _resolveColumn_(headers, ['account_code'], 5),
    payee: _resolveColumn_(headers, ['payee'], 6),
    refNo: _resolveColumn_(headers, ['ref_no'], 7),
    subCategory: _resolveColumn_(headers, ['sub_category'], 8),
    category: _resolveColumn_(headers, ['category'], 9),
    description: _resolveColumn_(headers, ['description'], 10),
    debit: _resolveColumn_(headers, ['debit'], 11),
    credit: _resolveColumn_(headers, ['credit'], 12),
    accountType: _resolveColumn_(headers, ['account_type'], 13),
    reportMapping: _resolveColumn_(headers, ['report_mapping'], 14),
    reconStatus: _resolveColumn_(headers, ['recon_status'], 15),
    receiptUrl: _resolveColumn_(headers, ['receipt_url'], 16)
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

function _storeUndoAction_(type, rowId, values) {
  const props = PropertiesService.getUserProperties();
  props.setProperty('lastJournalAction', JSON.stringify({
    type: type,
    rowId: rowId,
    values: values,
    timestamp: new Date().toISOString()
  }));
}

function _parseDate_(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function _formatDate_(date) {
  const d = date instanceof Date ? date : _parseDate_(date);
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function _findRowForInsert_(data, colIndex, blockCols) {
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (String(row[colIndex - 1]).trim()) continue;
    let blocked = false;
    (blockCols || []).forEach(function(blockCol) {
      if (blockCol && String(row[blockCol - 1]).trim()) {
        blocked = true;
      }
    });
    if (!blocked) return i + 2;
  }
  return null;
}

function _writeRowUpdate_(sheet, data, targetRow, lastCol, updates) {
  let rowValues;
  if (targetRow) {
    const idx = targetRow - 2;
    rowValues = idx >= 0 && idx < data.length ? data[idx].slice() : new Array(lastCol).fill('');
  } else {
    rowValues = new Array(lastCol).fill('');
    targetRow = sheet.getLastRow() + 1;
  }

  Object.keys(updates).forEach(function(colKey) {
    const colIndex = Number(colKey);
    if (!colIndex) return;
    rowValues[colIndex - 1] = updates[colKey];
  });

  sheet.getRange(targetRow, 1, 1, lastCol).setValues([rowValues]);
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
