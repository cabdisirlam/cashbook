/**
 * Payment modal backend helpers.
 */
function getDropdownData() {
  const payload = {
    accounts: [],
    payees: [],
    particulars: [],
    subCats: [],
    subToCatMap: {},
    particularMeta: {},
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
    const particulars = cols.particulars ? String(row[cols.particulars - 1]).trim() : '';
    const subCat = cols.subCategory ? String(row[cols.subCategory - 1]).trim() : '';
    const category = cols.category ? String(row[cols.category - 1]).trim() : '';
    const financialYear = cols.financialYear ? String(row[cols.financialYear - 1]).trim() : '';
    const accountType = cols.accountType ? String(row[cols.accountType - 1]).trim() : '';
    const reportMapping = cols.reportMapping ? String(row[cols.reportMapping - 1]).trim() : '';

    if (account) payload.accounts.push(account);
    if (payee) payload.payees.push(payee);
    if (category) payload.categories.push(category);
    if (financialYear) payload.financialYears.push(financialYear);
    if (subCat) {
      payload.subCats.push(subCat);
      if (category) payload.subToCatMap[subCat] = category;
    }
    if (particulars) {
      payload.particulars.push(particulars);
      payload.particularMeta[particulars] = {
        subCategory: subCat,
        category: category,
        accountType: accountType,
        reportMapping: reportMapping
      };
    }
  });

  payload.accounts = _uniqueSorted_(payload.accounts);
  payload.payees = _uniqueSorted_(payload.payees);
  payload.particulars = _uniqueSorted_(payload.particulars);
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
  const particularMeta = _buildParticularMeta_(masterData, masterCols);

  const cleanedRows = rows.map(function(row) {
    const amount = Number(row.amount || 0);
    const particulars = String(row.particulars || '').trim();
    const description = String(row.description || '').trim();

    if (!particulars) throw new Error('Each line needs particulars.');
    const meta = particularMeta[particulars];
    if (!meta || !meta.subCategory || !meta.category) {
      throw new Error('Particulars not found: ' + particulars + '.');
    }
    const subCategory = meta.subCategory;
    const category = meta.category;
    if (amount <= 0) throw new Error('Line amount must be greater than zero.');

    const accountTypeValue = meta.accountType || '';
    const reportMappingValue = meta.reportMapping || '';

    return {
      particulars,
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
      row.particulars,
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

function getRecentTransactionsByType(type, limit) {
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) throw new Error('DB_JOURNAL not found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getJournalColumns_(headers);
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const requested = String(type || '').toLowerCase();
  const maxRows = Math.max(1, Number(limit) || 5);
  const results = [];

  for (let i = data.length - 1; i >= 0 && results.length < maxRows; i--) {
    const row = data[i];
    const debit = cols.debit ? _parseNumber_(row[cols.debit - 1]) : 0;
    const credit = cols.credit ? _parseNumber_(row[cols.credit - 1]) : 0;
    const isReceipt = credit > 0 && debit === 0;
    const isPayment = debit > 0 && credit === 0;
    if (requested === 'receipt' && !isReceipt) continue;
    if (requested === 'payment' && !isPayment) continue;

    const dateValue = cols.date ? row[cols.date - 1] : '';
    results.push({
      date: dateValue ? _formatDate_(dateValue) : '',
      payee: cols.payee ? row[cols.payee - 1] : '',
      refNo: cols.refNo ? row[cols.refNo - 1] : '',
      particulars: cols.particulars ? row[cols.particulars - 1] : '',
      subCategory: cols.subCategory ? row[cols.subCategory - 1] : '',
      category: cols.category ? row[cols.category - 1] : '',
      amount: requested === 'receipt' ? credit : debit
    });
  }

  return results;
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
  const particularsFilter = String(criteria && criteria.particulars || '').trim();
  const payeeFilter = String(criteria && criteria.payee || '').trim().toLowerCase();
  const startDate = _parseDate_(criteria && criteria.startDate);
  const endDate = _parseDate_(criteria && criteria.endDate);
  const hasFilters = Boolean(query || refOrId || categoryFilter || particularsFilter || payeeFilter || startDate || endDate);

  const results = [];

  data.forEach(function(row, index) {
    const rowDate = cols.date ? row[cols.date - 1] : null;
    const dateValue = rowDate instanceof Date ? rowDate : _parseDate_(rowDate);
    if (startDate && dateValue && dateValue < startDate) return;
    if (endDate && dateValue && dateValue > endDate) return;

    const category = cols.category ? String(row[cols.category - 1]).trim() : '';
    if (categoryFilter && category !== categoryFilter) return;

    const particulars = cols.particulars ? String(row[cols.particulars - 1]).trim() : '';
    if (particularsFilter && particulars !== particularsFilter) return;

    const payee = cols.payee ? String(row[cols.payee - 1]).trim() : '';
    if (payeeFilter && !payee.toLowerCase().includes(payeeFilter)) return;

    const uuid = cols.uuid ? String(row[cols.uuid - 1]).trim() : '';
    const refNo = cols.refNo ? String(row[cols.refNo - 1]).trim() : '';
    if (refOrId && !(uuid.toLowerCase().includes(refOrId) || refNo.toLowerCase().includes(refOrId))) return;

    const haystack = [
      cols.accountCode ? row[cols.accountCode - 1] : '',
      payee,
      category,
      particulars,
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
      particulars: cols.particulars ? row[cols.particulars - 1] : '',
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
    particulars: cols.particulars ? values[cols.particulars - 1] || '' : '',
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
  _setIfPresent_(updated, cols.date, _parseDate_(payload.date) || existing[cols.date - 1]);
  _setIfPresent_(updated, cols.financialYear, payload.financialYear || '');
  _setIfPresent_(updated, cols.accountCode, payload.accountCode || '');
  _setIfPresent_(updated, cols.payee, payload.payee || '');
  _setIfPresent_(updated, cols.refNo, payload.refNo || '');
  _setIfPresent_(updated, cols.particulars, payload.particulars || '');
  if (payload.particulars && (!payload.subCategory || !payload.category)) {
    const details = getDetailsForParticulars(String(payload.particulars || '').trim());
    _setIfPresent_(updated, cols.subCategory, details.subCategory || '');
    _setIfPresent_(updated, cols.category, details.category || '');
  } else {
    _setIfPresent_(updated, cols.subCategory, payload.subCategory || '');
    _setIfPresent_(updated, cols.category, payload.category || '');
  }
  _setIfPresent_(updated, cols.description, payload.description || '');
  _setIfPresent_(updated, cols.debit, Number(payload.debit || 0));
  _setIfPresent_(updated, cols.credit, Number(payload.credit || 0));
  _setIfPresent_(updated, cols.accountType, payload.accountType || '');
  _setIfPresent_(updated, cols.reportMapping, payload.reportMapping || '');
  _setIfPresent_(updated, cols.reconStatus, payload.reconStatus || '');
  _setIfPresent_(updated, cols.receiptUrl, payload.receiptUrl || '');

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

function getBudgetVsActual(financialYear) {
  const year = String(financialYear || '').trim();
  if (!year) throw new Error('Financial year is required.');

  const ss = _getOrCreateSpreadsheet();
  const budgetSheet = ss.getSheetByName(CONFIG.SHEETS.DB_BUDGET);
  if (!budgetSheet) throw new Error('DB_BUDGET not found.');

  const budgetLastRow = budgetSheet.getLastRow();
  const budgetLastCol = budgetSheet.getLastColumn();
  if (budgetLastCol < 1) return { rows: [], totals: {} };

  const budgetHeaders = budgetSheet.getRange(1, 1, 1, budgetLastCol).getValues()[0];
  const budgetMap = _getBudgetHeaderMap(budgetSheet);
  _ensureBudgetHeaders(budgetMap);
  const budgetData = budgetLastRow >= 2
    ? budgetSheet.getRange(2, 1, budgetLastRow - 1, budgetLastCol).getValues()
    : [];

  const budgetByParticular = {};
  budgetData.forEach(row => {
    const rowYear = String(row[budgetMap.Financial_Year] || '').trim();
    if (rowYear !== year) return;
    const particulars = String(row[budgetMap.Particulars] || '').trim();
    if (!particulars) return;
    const sub = String(row[budgetMap.Sub_Category] || '').trim();

    if (!budgetByParticular[particulars]) {
      budgetByParticular[particulars] = {
        particulars: particulars,
        subCategory: sub,
        category: String(row[budgetMap.Category] || '').trim(),
        accountType: String(row[budgetMap.Account_Type] || '').trim(),
        originalBudget: 0,
        reallocation: 0,
        supplementary: 0
      };
    }
    budgetByParticular[particulars].originalBudget += Number(row[budgetMap.Original_Budget] || 0);
    budgetByParticular[particulars].reallocation += Number(row[budgetMap.Reallocation] || 0);
    budgetByParticular[particulars].supplementary += Number(row[budgetMap.Supplementary] || 0);
  });

  const journal = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!journal) throw new Error('DB_JOURNAL not found.');

  const journalLastRow = journal.getLastRow();
  const journalLastCol = journal.getLastColumn();
  const actualByParticular = {};
  const debitByParticular = {};
  const creditByParticular = {};
  const journalMetaByParticular = {};
  let journalDebitTotal = 0;
  let journalCreditTotal = 0;
  if (journalLastRow >= 2) {
    const journalHeaders = journal.getRange(1, 1, 1, journalLastCol).getValues()[0].map(_normalizeHeader_);
    const cols = _getJournalColumns_(journalHeaders);
    const journalData = journal.getRange(2, 1, journalLastRow - 1, journalLastCol).getValues();

    journalData.forEach(row => {
      if (!cols.particulars) return;
      const rowYear = cols.financialYear ? String(row[cols.financialYear - 1] || '').trim() : '';
      if (rowYear !== year) return;
      const particulars = String(row[cols.particulars - 1] || '').trim();
      if (!particulars) return;
      const debit = cols.debit ? _parseNumber_(row[cols.debit - 1]) : 0;
      const credit = cols.credit ? _parseNumber_(row[cols.credit - 1]) : 0;
      debitByParticular[particulars] = (debitByParticular[particulars] || 0) + debit;
      creditByParticular[particulars] = (creditByParticular[particulars] || 0) + credit;
      journalDebitTotal += debit;
      journalCreditTotal += credit;
      if (!journalMetaByParticular[particulars]) {
        journalMetaByParticular[particulars] = { subCategory: '', category: '', accountType: '' };
      }
      if (cols.subCategory && !journalMetaByParticular[particulars].subCategory) {
        journalMetaByParticular[particulars].subCategory = String(row[cols.subCategory - 1] || '').trim();
      }
      if (cols.category && !journalMetaByParticular[particulars].category) {
        journalMetaByParticular[particulars].category = String(row[cols.category - 1] || '').trim();
      }
      if (cols.accountType && !journalMetaByParticular[particulars].accountType) {
        journalMetaByParticular[particulars].accountType = String(row[cols.accountType - 1] || '').trim();
      }
    });
  }

  const actualParticulars = new Set(Object.keys(debitByParticular).concat(Object.keys(creditByParticular)));
  const missingParticulars = [];
  actualParticulars.forEach(particulars => {
    if (budgetByParticular[particulars]) return;
    const meta = journalMetaByParticular[particulars] || {};
    let subCategory = String(meta.subCategory || '').trim();
    let category = String(meta.category || '').trim();
    let accountType = String(meta.accountType || '').trim();
    if (!subCategory || !category || !accountType) {
      const details = getDetailsForParticulars(particulars);
      subCategory = subCategory || details.subCategory || '';
      category = category || details.category || '';
      accountType = accountType || details.accountType || '';
    }
    budgetByParticular[particulars] = {
      particulars: particulars,
      subCategory: subCategory,
      category: category,
      accountType: accountType,
      originalBudget: 0,
      reallocation: 0,
      supplementary: 0
    };
    missingParticulars.push(particulars);
  });

  const results = Object.values(budgetByParticular).map(item => {
    const debitActual = debitByParticular[item.particulars] || 0;
    const creditActual = creditByParticular[item.particulars] || 0;
    const accountType = String(item.accountType || '').toLowerCase();
    const isReceipt = accountType.includes('income') || (!accountType && creditActual > debitActual);
    const actual = isReceipt ? creditActual : debitActual;
    actualByParticular[item.particulars] = actual;
    const finalBudget = item.originalBudget + item.reallocation + item.supplementary;
    const variance = finalBudget - actual;
    return {
      particulars: item.particulars,
      subCategory: item.subCategory,
      category: item.category,
      originalBudget: item.originalBudget,
      reallocation: item.reallocation,
      supplementary: item.supplementary,
      finalBudget: finalBudget,
      actualAmount: actual,
      variance: variance,
      section: isReceipt ? 'Receipts' : 'Payments'
    };
  }).sort((a, b) => {
    if (a.section !== b.section) return a.section.localeCompare(b.section);
    if (a.category === b.category) {
      if (a.subCategory === b.subCategory) return a.particulars.localeCompare(b.particulars);
      return a.subCategory.localeCompare(b.subCategory);
    }
    return a.category.localeCompare(b.category);
  });

  if (budgetData.length) {
    const updated = budgetData.map(row => {
      const rowYear = String(row[budgetMap.Financial_Year] || '').trim();
      if (rowYear !== year) return row;
      const particulars = String(row[budgetMap.Particulars] || '').trim();
      const summary = budgetByParticular[particulars];
      if (!summary) return row;
      const debitActual = debitByParticular[summary.particulars] || 0;
      const creditActual = creditByParticular[summary.particulars] || 0;
      const accountType = String(summary.accountType || '').toLowerCase();
      const isReceipt = accountType.includes('income') || (!accountType && creditActual > debitActual);
      const actual = isReceipt ? creditActual : debitActual;
      const finalBudget = summary.originalBudget + summary.reallocation + summary.supplementary;
      row[budgetMap.Final_Budget] = finalBudget;
      row[budgetMap.Actual_Amount] = actual;
      row[budgetMap.Variance] = finalBudget - actual;
      return row;
    });
    budgetSheet.getRange(2, 1, updated.length, budgetLastCol).setValues(updated);
  }

  if (missingParticulars.length) {
    const headerCount = budgetLastCol;
    const now = new Date();
    const rowsToInsert = missingParticulars.map(particulars => {
      const summary = budgetByParticular[particulars];
      const debitActual = debitByParticular[particulars] || 0;
      const creditActual = creditByParticular[particulars] || 0;
      const accountType = String(summary.accountType || '').toLowerCase();
      const isReceipt = accountType.includes('income') || (!accountType && creditActual > debitActual);
      const actual = isReceipt ? creditActual : debitActual;
      const row = new Array(headerCount).fill('');
      row[budgetMap.Date] = now;
      row[budgetMap.Financial_Year] = year;
      row[budgetMap.Particulars] = summary.particulars;
      row[budgetMap.Sub_Category] = summary.subCategory;
      row[budgetMap.Category] = summary.category;
      row[budgetMap.Account_Type] = summary.accountType;
      row[budgetMap.Original_Budget] = '';
      row[budgetMap.Reallocation] = '';
      row[budgetMap.Supplementary] = '';
      row[budgetMap.Final_Budget] = 0;
      row[budgetMap.Actual_Amount] = actual;
      row[budgetMap.Variance] = 0 - actual;
      row[budgetMap.Description] = 'Audit: actual without budget';
      return row;
    });
    const startRow = budgetSheet.getLastRow() + 1;
    budgetSheet.getRange(startRow, 1, rowsToInsert.length, headerCount).setValues(rowsToInsert);
    logSystemEventSafe('AUDIT_MISSING_BUDGET_SUBS', year, 'Rows: ' + rowsToInsert.length);
  }

  const totals = results.reduce((acc, row) => {
    acc.originalBudget += row.originalBudget || 0;
    acc.reallocation += row.reallocation || 0;
    acc.supplementary += row.supplementary || 0;
    acc.finalBudget += row.finalBudget || 0;
    acc.actualAmount += row.actualAmount || 0;
    acc.variance += row.variance || 0;
    return acc;
  }, { originalBudget: 0, reallocation: 0, supplementary: 0, finalBudget: 0, actualAmount: 0, variance: 0 });

  const sectionTotals = {
    receipts: { originalBudget: 0, reallocation: 0, supplementary: 0, finalBudget: 0, actualAmount: 0, variance: 0 },
    payments: { originalBudget: 0, reallocation: 0, supplementary: 0, finalBudget: 0, actualAmount: 0, variance: 0 }
  };

  results.forEach(row => {
    const key = row.section === 'Receipts' ? 'receipts' : 'payments';
    sectionTotals[key].originalBudget += row.originalBudget || 0;
    sectionTotals[key].reallocation += row.reallocation || 0;
    sectionTotals[key].supplementary += row.supplementary || 0;
    sectionTotals[key].finalBudget += row.finalBudget || 0;
    sectionTotals[key].actualAmount += row.actualAmount || 0;
    sectionTotals[key].variance += row.variance || 0;
  });

  const surplus = {
    originalBudget: sectionTotals.receipts.originalBudget - sectionTotals.payments.originalBudget,
    reallocation: sectionTotals.receipts.reallocation - sectionTotals.payments.reallocation,
    supplementary: sectionTotals.receipts.supplementary - sectionTotals.payments.supplementary,
    finalBudget: sectionTotals.receipts.finalBudget - sectionTotals.payments.finalBudget,
    actualAmount: sectionTotals.receipts.actualAmount - sectionTotals.payments.actualAmount,
    variance: sectionTotals.receipts.variance - sectionTotals.payments.variance
  };

  const journalActualTotal = Array.from(actualParticulars).reduce((sum, particulars) => {
    const debitActual = debitByParticular[particulars] || 0;
    const creditActual = creditByParticular[particulars] || 0;
    const summary = budgetByParticular[particulars] || {};
    const accountType = String(summary.accountType || '').toLowerCase();
    const isReceipt = accountType.includes('income') || (!accountType && creditActual > debitActual);
    const actual = isReceipt ? creditActual : debitActual;
    return sum + actual;
  }, 0);

  if (Math.abs(journalActualTotal - totals.actualAmount) > 0.01) {
    logSystemEventSafe(
      'AUDIT_ACTUAL_MISMATCH',
      year,
      'Budget total: ' + totals.actualAmount + ', Journal total: ' + journalActualTotal + ', Debit: ' + journalDebitTotal + ', Credit: ' + journalCreditTotal
    );
  }

  logSystemEventSafe('REFRESH_BUDGET_ACTUALS', year, 'Rows: ' + results.length);

  return {
    rows: results,
    totals: totals,
    sectionTotals: sectionTotals,
    surplus: surplus
  };
}

function exportBudgetVsActual(financialYear) {
  const result = getBudgetVsActual(financialYear);
  if (!result.rows || !result.rows.length) return { csv: '', filename: '' };

  const headers = ['Particulars', 'Sub_Category', 'Category', 'Original_Budget', 'Reallocation', 'Supplementary', 'Final_Budget', 'Actual_Amount', 'Variance'];
  const lines = [headers];

  const receipts = result.rows.filter(row => row.section === 'Receipts');
  const payments = result.rows.filter(row => row.section === 'Payments');

  lines.push(['Receipts', '', '', '', '', '', '', '', '']);
  receipts.forEach(row => {
    lines.push([
      row.particulars,
      row.subCategory,
      row.category,
      row.originalBudget,
      row.reallocation,
      row.supplementary,
      row.finalBudget,
      row.actualAmount,
      row.variance
    ]);
  });
  const receiptTotals = result.sectionTotals.receipts;
  lines.push([
    'Total Receipts',
    '',
    '',
    '',
    receiptTotals.originalBudget,
    receiptTotals.reallocation,
    receiptTotals.supplementary,
    receiptTotals.finalBudget,
    receiptTotals.actualAmount,
    receiptTotals.variance
  ]);

  lines.push(['Payments', '', '', '', '', '', '', '', '']);
  payments.forEach(row => {
    lines.push([
      row.particulars,
      row.subCategory,
      row.category,
      row.originalBudget,
      row.reallocation,
      row.supplementary,
      row.finalBudget,
      row.actualAmount,
      row.variance
    ]);
  });
  const paymentTotals = result.sectionTotals.payments;
  lines.push([
    'Total Payments',
    '',
    '',
    '',
    paymentTotals.originalBudget,
    paymentTotals.reallocation,
    paymentTotals.supplementary,
    paymentTotals.finalBudget,
    paymentTotals.actualAmount,
    paymentTotals.variance
  ]);

  const surplus = result.surplus;
  lines.push([
    'Surplus / Deficit',
    '',
    '',
    '',
    surplus.originalBudget,
    surplus.reallocation,
    surplus.supplementary,
    surplus.finalBudget,
    surplus.actualAmount,
    surplus.variance
  ]);

  const csv = lines.map(line => line.map(cell => {
    const value = String(cell == null ? '' : cell);
    return '"' + value.replace(/"/g, '""') + '"';
  }).join(',')).join('\n');

  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  return {
    csv: csv,
    filename: 'budget_vs_actual_' + stamp + '.csv'
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

  if (typeKey === 'particular') {
    const parentSubCategory = String(parent || '').trim();
    const categoryValue = String(accountType || '').trim();
    const accountTypeValue = String(reportMapping || '').trim();
    const reportValue = String(arguments[5] || '').trim();
    if (!parentSubCategory) throw new Error('Parent sub-category is required.');
    if (!categoryValue) throw new Error('Category is required.');
    if (!accountTypeValue) throw new Error('Account type is required.');
    if (!reportValue) throw new Error('Report mapping is required.');
    if (_valueExistsInColumn_(data, cols.particulars, trimmed)) return;

    const targetRow = _findRowForInsert_(data, cols.particulars, [cols.accountCodes]);
    _writeRowUpdate_(sheet, data, targetRow, lastCol, {
      [cols.particulars]: trimmed,
      [cols.subCategory]: parentSubCategory,
      [cols.category]: categoryValue,
      [cols.accountType]: accountTypeValue,
      [cols.reportMapping]: reportValue
    });
    logSystemEventSafe('CREATE_MASTER_DATA', trimmed, 'Type: particular, Sub-Category: ' + parentSubCategory);
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
    particulars: _resolveColumn_(headers, ['particulars', 'particular'], 2),
    subCategory: _resolveColumn_(headers, ['sub_category', 'subcategory'], 3),
    category: _resolveColumn_(headers, ['category'], 4),
    accountCodes: _resolveColumn_(headers, ['account_codes', 'account_code'], 5),
    accountType: _resolveColumn_(headers, ['account_type', 'accounttype'], 6),
    reportMapping: _resolveColumn_(headers, ['report_mapping', 'reportmapping'], 7),
    financialYear: _resolveColumn_(headers, ['financial_year', 'financialyear'], 8)
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
    particulars: _resolveColumn_(headers, ['particulars', 'particular'], 8),
    subCategory: _resolveColumn_(headers, ['sub_category'], 9),
    category: _resolveColumn_(headers, ['category'], 10),
    description: _resolveColumn_(headers, ['description'], 11),
    debit: _resolveColumn_(headers, ['debit'], 12),
    credit: _resolveColumn_(headers, ['credit'], 13),
    accountType: _resolveColumn_(headers, ['account_type'], 14),
    reportMapping: _resolveColumn_(headers, ['report_mapping'], 15),
    reconStatus: _resolveColumn_(headers, ['recon_status'], 16),
    receiptUrl: _resolveColumn_(headers, ['receipt_url'], 17)
  };
}

function _resolveColumn_(headers, names, fallback) {
  for (let i = 0; i < names.length; i++) {
    const idx = headers.indexOf(names[i]);
    if (idx >= 0) return idx + 1;
  }
  return fallback;
}

function _parseNumber_(value) {
  if (value == null || value === '') return 0;
  if (typeof value === 'number') return value;
  const cleaned = String(value).replace(/,/g, '').trim();
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function _setIfPresent_(row, colIndex, value) {
  if (!colIndex) return;
  if (colIndex > row.length) return;
  row[colIndex - 1] = value;
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

function _buildParticularMeta_(data, cols) {
  const meta = {};
  data.forEach(function(row) {
    const particulars = cols.particulars ? String(row[cols.particulars - 1]).trim() : '';
    if (!particulars) return;
    meta[particulars] = {
      subCategory: cols.subCategory ? String(row[cols.subCategory - 1]).trim() : '',
      category: cols.category ? String(row[cols.category - 1]).trim() : '',
      accountType: cols.accountType ? String(row[cols.accountType - 1]).trim() : '',
      reportMapping: cols.reportMapping ? String(row[cols.reportMapping - 1]).trim() : ''
    };
  });
  return meta;
}
