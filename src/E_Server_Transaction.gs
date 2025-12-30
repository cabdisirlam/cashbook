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
  const financialYearFilter = String(criteria && criteria.financialYear || '').trim();
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
  const yearIndex = headerMap.indexOf('financial_year');
  const accountIndex = headerMap.indexOf('account_code');

  const filtered = data.filter(function(row) {
    if (dateIndex >= 0) {
      const rowDate = row[dateIndex];
      const dateValue = rowDate instanceof Date ? rowDate : _parseDate_(rowDate);
      if (!dateValue) return true;
      if (startDate && dateValue < startDate) return false;
      if (endDate && dateValue > endDate) return false;
    }
    if (accountIndex >= 0 && criteria && criteria.accountCode) {
      const accountValue = String(row[accountIndex] || '').trim();
      if (accountValue !== String(criteria.accountCode || '').trim()) return false;
    }
    if (yearIndex >= 0 && criteria && criteria.financialYear) {
      const yearValue = String(row[yearIndex] || '').trim();
      if (yearValue !== String(criteria.financialYear || '').trim()) return false;
    }
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

function getReconciliationPreview(criteria) {
  const data = _collectReconciliationData_(criteria);
  const matchResult = _matchReconRows_(data.journalRows, data.bankRows);
  return _buildReconciliationResponse_(data.journalRows, data.bankRows, matchResult, true);
}

function autoReconcileBankStatements(criteria) {
  const data = _collectReconciliationData_(criteria);
  const matchResult = _matchReconRows_(data.journalRows, data.bankRows);
  if (matchResult.pairs.length) {
    const journalSheet = data.journalSheet;
    const bankSheet = data.bankSheet;
    const journalReconCol = data.journalCols.reconStatus;
    const bankMatchCol = data.bankCols.matchStatus;

    matchResult.pairs.forEach(function(pair) {
      if (journalReconCol) {
        journalSheet.getRange(pair.journal.rowIndex, journalReconCol).setValue('Reconciled');
      }
      if (bankMatchCol) {
        bankSheet.getRange(pair.bank.rowIndex, bankMatchCol).setValue('Reconciled');
      }
    });
  }

  const response = _buildReconciliationResponse_(data.journalRows, data.bankRows, matchResult);
  response.matchedCount = matchResult.pairs.length;
  return response;
}

function exportBankReconciliation(criteria) {
  const data = _collectReconciliationData_(criteria);
  const matchResult = _matchReconRows_(data.journalRows, data.bankRows);

  const cashbookReceipts = [];
  const cashbookPayments = [];
  const bankReceipts = [];
  const bankPayments = [];

  data.journalRows.forEach(function(row, index) {
    if (matchResult.matchedJournal.has(index)) return;
    if (row.type === 'receipt') {
      cashbookReceipts.push(row);
    } else {
      cashbookPayments.push(row);
    }
  });

  data.bankRows.forEach(function(row, index) {
    if (matchResult.matchedBank.has(index)) return;
    if (row.type === 'receipt') {
      bankReceipts.push(row);
    } else {
      bankPayments.push(row);
    }
  });

  const totals = {
    cashbookReceipts: _sumReconAmounts_(cashbookReceipts),
    cashbookPayments: _sumReconAmounts_(cashbookPayments),
    bankReceipts: _sumReconAmounts_(bankReceipts),
    bankPayments: _sumReconAmounts_(bankPayments)
  };

  const settings = _getReconReportSettings_();
  const bankBalance = _getLatestBankBalance_(data.bankSheet, data.bankCols, criteria);
  const cashbookBalance = bankBalance
    - totals.cashbookPayments
    - totals.bankReceipts
    + totals.bankPayments
    + totals.cashbookReceipts;

  const reportName = 'Bank Reconciliation';
  const report = SpreadsheetApp.create(reportName);
  const sheet = report.getActiveSheet();
  sheet.setName('Bank Reconciliation');

  _writeReconExportSheet_(sheet, {
    settings: settings,
    criteria: criteria,
    totals: totals,
    bankBalance: bankBalance,
    cashbookBalance: cashbookBalance,
    cashbookPayments: cashbookPayments,
    bankReceipts: bankReceipts,
    bankPayments: bankPayments,
    cashbookReceipts: cashbookReceipts,
    totals: totals
  });

  const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const exportUrl = 'https://www.googleapis.com/drive/v3/files/' + report.getId()
    + '/export?mimeType=' + encodeURIComponent(mimeType);
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error('Failed to export XLSX. Check Drive permissions.');
  }
  const base64 = Utilities.base64Encode(response.getContent());
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  DriveApp.getFileById(report.getId()).setTrashed(true);

  return {
    base64: base64,
    filename: 'bank_reconciliation_' + stamp + '.xlsx',
    mimeType: mimeType
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

function getNotesReport(currentYear, comparativeYear) {
  const year = String(currentYear || '').trim();
  const compare = String(comparativeYear || '').trim();
  if (!year) throw new Error('Current financial year is required.');
  if (!compare) throw new Error('Comparative financial year is required.');

  const ss = _getOrCreateSpreadsheet();
  const master = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);
  if (!master) throw new Error('MASTER_DATA not found.');

  const masterLastRow = master.getLastRow();
  const masterLastCol = master.getLastColumn();
  const categories = {};
  const categoryOrder = [];
  const particularsMeta = {};
  const categoryAccountTypes = {};
  const categoryReportMappings = {};

  if (masterLastRow >= 2) {
    const headers = master.getRange(1, 1, 1, masterLastCol).getValues()[0].map(_normalizeHeader_);
    const cols = _getMasterColumns_(headers);
    const data = master.getRange(2, 1, masterLastRow - 1, masterLastCol).getValues();

    data.forEach(row => {
      const particulars = cols.particulars ? String(row[cols.particulars - 1] || '').trim() : '';
      if (!particulars) return;
      const subCategory = cols.subCategory ? String(row[cols.subCategory - 1] || '').trim() : '';
      const category = cols.category ? String(row[cols.category - 1] || '').trim() : '';
      const accountType = cols.accountType ? String(row[cols.accountType - 1] || '').trim() : '';
      const reportMapping = cols.reportMapping ? String(row[cols.reportMapping - 1] || '').trim() : '';
      if (!category) return;

      particularsMeta[particulars] = {
        category: category,
        subCategory: subCategory,
        accountType: accountType,
        reportMapping: reportMapping
      };
      if (!categoryAccountTypes[category]) categoryAccountTypes[category] = {};
      if (accountType) categoryAccountTypes[category][accountType] = true;
      if (!categoryReportMappings[category]) categoryReportMappings[category] = {};
      if (reportMapping) categoryReportMappings[category][reportMapping] = true;

      if (!categories[category]) {
        categories[category] = { name: category, subcategories: {}, order: [] };
        categoryOrder.push(category);
      }
      const subKey = subCategory || 'Other';
      if (!categories[category].subcategories[subKey]) {
        categories[category].subcategories[subKey] = { name: subKey, particulars: [] };
        categories[category].order.push(subKey);
      }
      if (categories[category].subcategories[subKey].particulars.indexOf(particulars) === -1) {
        categories[category].subcategories[subKey].particulars.push(particulars);
      }
    });
  }

  const journal = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!journal) throw new Error('DB_JOURNAL not found.');

  const journalLastRow = journal.getLastRow();
  const journalLastCol = journal.getLastColumn();
  const current = { debit: {}, credit: {}, meta: {} };
  const comparative = { debit: {}, credit: {}, meta: {} };
  const bankNetCurrent = {};
  const bankNetComparative = {};

  if (journalLastRow >= 2) {
    const journalHeaders = journal.getRange(1, 1, 1, journalLastCol).getValues()[0].map(_normalizeHeader_);
    const cols = _getJournalColumns_(journalHeaders);
    const journalData = journal.getRange(2, 1, journalLastRow - 1, journalLastCol).getValues();

    journalData.forEach(row => {
      if (!cols.particulars || !cols.financialYear) return;
      const rowYear = String(row[cols.financialYear - 1] || '').trim();
      if (rowYear !== year && rowYear !== compare) return;
      const particulars = String(row[cols.particulars - 1] || '').trim();
      if (!particulars) return;

      const debit = cols.debit ? _parseNumber_(row[cols.debit - 1]) : 0;
      const credit = cols.credit ? _parseNumber_(row[cols.credit - 1]) : 0;
      const target = rowYear === year ? current : comparative;

      target.debit[particulars] = (target.debit[particulars] || 0) + debit;
      target.credit[particulars] = (target.credit[particulars] || 0) + credit;

      if (!target.meta[particulars]) {
        target.meta[particulars] = {
          subCategory: cols.subCategory ? String(row[cols.subCategory - 1] || '').trim() : '',
          category: cols.category ? String(row[cols.category - 1] || '').trim() : '',
          accountType: cols.accountType ? String(row[cols.accountType - 1] || '').trim() : '',
          reportMapping: cols.reportMapping ? String(row[cols.reportMapping - 1] || '').trim() : ''
        };
      }

      if (cols.accountCode) {
        const accountCode = String(row[cols.accountCode - 1] || '').trim();
        if (accountCode) {
          const net = credit - debit;
          if (rowYear === year) {
            bankNetCurrent[accountCode] = (bankNetCurrent[accountCode] || 0) + net;
          } else {
            bankNetComparative[accountCode] = (bankNetComparative[accountCode] || 0) + net;
          }
        }
      }
    });
  }

  const journalParticulars = new Set(
    Object.keys(current.debit)
      .concat(Object.keys(current.credit))
      .concat(Object.keys(comparative.debit))
      .concat(Object.keys(comparative.credit))
  );

  journalParticulars.forEach(particulars => {
    if (particularsMeta[particulars]) return;
    const fallback = current.meta[particulars] || comparative.meta[particulars] || {};
    let category = String(fallback.category || '').trim();
    let subCategory = String(fallback.subCategory || '').trim();
    let accountType = String(fallback.accountType || '').trim();
    let reportMapping = String(fallback.reportMapping || '').trim();

    if (!category || !subCategory || !accountType || !reportMapping) {
      const details = getDetailsForParticulars(particulars);
      category = category || details.category || '';
      subCategory = subCategory || details.subCategory || '';
      accountType = accountType || details.accountType || '';
      reportMapping = reportMapping || details.reportMapping || '';
    }

    if (!category) category = 'Uncategorized';
    if (!subCategory) subCategory = 'Other';

    particularsMeta[particulars] = {
      category: category,
      subCategory: subCategory,
      accountType: accountType,
      reportMapping: reportMapping
    };
    if (!categoryAccountTypes[category]) categoryAccountTypes[category] = {};
    if (accountType) categoryAccountTypes[category][accountType] = true;
    if (!categoryReportMappings[category]) categoryReportMappings[category] = {};
    if (reportMapping) categoryReportMappings[category][reportMapping] = true;

    if (!categories[category]) {
      categories[category] = { name: category, subcategories: {}, order: [] };
      categoryOrder.push(category);
    }
    if (!categories[category].subcategories[subCategory]) {
      categories[category].subcategories[subCategory] = { name: subCategory, particulars: [] };
      categories[category].order.push(subCategory);
    }
    if (categories[category].subcategories[subCategory].particulars.indexOf(particulars) === -1) {
      categories[category].subcategories[subCategory].particulars.push(particulars);
    }
  });

  const sortedCategories = categoryOrder.length ? categoryOrder.slice() : Object.keys(categories);
  sortedCategories.sort((a, b) => a.localeCompare(b));

  const results = [];
  sortedCategories.forEach(categoryName => {
    const category = categories[categoryName];
    if (!category) return;

    const subOrder = category.order.length ? category.order.slice() : Object.keys(category.subcategories);
    subOrder.sort((a, b) => a.localeCompare(b));

    let categoryTotalCurrent = 0;
    let categoryTotalComparative = 0;
    const subResults = [];

    subOrder.forEach(subName => {
      const sub = category.subcategories[subName];
      if (!sub) return;
      const particulars = sub.particulars.slice().sort((a, b) => a.localeCompare(b));
      const items = [];

      particulars.forEach(particularsName => {
        const meta = particularsMeta[particularsName] || {};
        const accountType = meta.accountType || '';
        const currentAmount = _resolveNotesAmount_(current, particularsName, accountType);
        const comparativeAmount = _resolveNotesAmount_(comparative, particularsName, accountType);
        if (!currentAmount && !comparativeAmount) return;

        categoryTotalCurrent += currentAmount;
        categoryTotalComparative += comparativeAmount;
        items.push({
          particulars: particularsName,
          currentAmount: currentAmount,
          comparativeAmount: comparativeAmount
        });
      });

      if (items.length) {
        subResults.push({
          subCategory: subName,
          items: items
        });
      }
    });

    if (!categoryTotalCurrent && !categoryTotalComparative) return;

    results.push({
      category: categoryName,
      totalCurrent: categoryTotalCurrent,
      totalComparative: categoryTotalComparative,
      subCategories: subResults,
      accountTypes: Object.keys(categoryAccountTypes[categoryName] || {}),
      reportMappings: Object.keys(categoryReportMappings[categoryName] || {})
    });
  });

  const bankAccounts = Array.from(
    new Set(Object.keys(bankNetCurrent).concat(Object.keys(bankNetComparative)))
  ).sort((a, b) => a.localeCompare(b));
  const bankItems = [];
  let bankTotalCurrent = 0;
  let bankTotalComparative = 0;
  bankAccounts.forEach(accountCode => {
    const currentAmount = bankNetCurrent[accountCode] || 0;
    const comparativeAmount = bankNetComparative[accountCode] || 0;
    if (!currentAmount && !comparativeAmount) return;
    bankTotalCurrent += currentAmount;
    bankTotalComparative += comparativeAmount;
    bankItems.push({
      particulars: accountCode,
      currentAmount: currentAmount,
      comparativeAmount: comparativeAmount
    });
  });

  if (bankItems.length) {
    results.push({
      category: 'Cash and Cash Equivalent',
      totalCurrent: bankTotalCurrent,
      totalComparative: bankTotalComparative,
      subCategories: [
        {
          subCategory: 'Banks',
          items: bankItems
        }
      ],
      accountTypes: ['asset']
    });
  }

  const getCategoryOrder = (item) => {
    const name = String(item.category || '').toLowerCase();
    const accountTypes = (item.accountTypes || []).map(value => String(value || '').toLowerCase());
    const mappings = (item.reportMappings || []).map(value => String(value || '').toLowerCase());

    if (name.includes('income') || name.includes('revenue') || accountTypes.some(v => v.includes('income')) || mappings.some(v => v.includes('operating income'))) {
      return 0;
    }
    if (name.includes('expense') || accountTypes.some(v => v.includes('expense')) || mappings.some(v => v.includes('operating expense'))) {
      return 1;
    }
    if (name.includes('cash and cash equivalent')) return 2;
    if (name.includes('non current asset') || name.includes('non-current asset') || mappings.some(v => v.includes('non-current asset') || v.includes('non current asset'))) {
      return 3;
    }
    if (name.includes('asset') || accountTypes.some(v => v.includes('asset')) || mappings.some(v => v.includes('current asset'))) {
      return 4;
    }
    if (name.includes('liabil') || accountTypes.some(v => v.includes('liabil')) || mappings.some(v => v.includes('liability'))) {
      return 5;
    }
    return 99;
  };

  results.sort((a, b) => {
    const aOrder = getCategoryOrder(a);
    const bOrder = getCategoryOrder(b);
    if (aOrder !== bOrder) return aOrder - bOrder;
    return String(a.category || '').localeCompare(String(b.category || ''));
  });

  return {
    currentYear: year,
    comparativeYear: compare,
    categories: results
  };
}

function getPerformanceReport(currentYear, comparativeYear) {
  const notes = getNotesReport(currentYear, comparativeYear);
  const categories = Array.isArray(notes.categories) ? notes.categories : [];
  const noteNumberByCategory = {};
  categories.forEach((category, index) => {
    if (category && category.category) {
      noteNumberByCategory[category.category] = index + 1;
    }
  });

  const revenueRows = [];
  const expenseRows = [];
  const otherRows = [];
  let revenueCurrent = 0;
  let revenueComparative = 0;
  let expenseCurrent = 0;
  let expenseComparative = 0;
  let otherCurrent = 0;
  let otherComparative = 0;

  categories.forEach(category => {
    if (!category || !category.category) return;
    const section = _classifyPerformanceCategory_(category.category, category.accountTypes || []);
    if (!section) return;
    const row = {
      description: category.category,
      note: noteNumberByCategory[category.category] || '',
      currentAmount: Number(category.totalCurrent || 0),
      comparativeAmount: Number(category.totalComparative || 0)
    };
    if (section === 'revenue') {
      revenueRows.push(row);
      revenueCurrent += row.currentAmount;
      revenueComparative += row.comparativeAmount;
    } else if (section === 'expense') {
      expenseRows.push(row);
      expenseCurrent += row.currentAmount;
      expenseComparative += row.comparativeAmount;
    } else if (section === 'other') {
      otherRows.push(row);
      otherCurrent += row.currentAmount;
      otherComparative += row.comparativeAmount;
    }
  });

  return {
    currentYear: notes.currentYear || String(currentYear || '').trim(),
    comparativeYear: notes.comparativeYear || String(comparativeYear || '').trim(),
    revenue: { rows: revenueRows, totalCurrent: revenueCurrent, totalComparative: revenueComparative },
    expenses: { rows: expenseRows, totalCurrent: expenseCurrent, totalComparative: expenseComparative },
    other: { rows: otherRows, totalCurrent: otherCurrent, totalComparative: otherComparative },
    surplus: {
      current: revenueCurrent - expenseCurrent + otherCurrent,
      comparative: revenueComparative - expenseComparative + otherComparative
    }
  };
}

function getCashFlowReport(currentYear, comparativeYear) {
  const notes = getNotesReport(currentYear, comparativeYear);
  const categories = Array.isArray(notes.categories) ? notes.categories : [];
  const noteNumberByCategory = {};
  categories.forEach((category, index) => {
    if (category && category.category) {
      noteNumberByCategory[category.category] = index + 1;
    }
  });

  const operatingReceipts = [];
  const operatingPayments = [];
  const investingRows = [];
  const financingRows = [];
  let receiptsCurrent = 0;
  let receiptsComparative = 0;
  let paymentsCurrent = 0;
  let paymentsComparative = 0;
  let investingCurrent = 0;
  let investingComparative = 0;
  let financingCurrent = 0;
  let financingComparative = 0;
  let cashNote = null;

  categories.forEach(category => {
    if (!category || !category.category) return;
    const section = _classifyCashFlowCategory_(
      category.category,
      category.accountTypes || [],
      category.reportMappings || []
    );
    if (!section) return;
    if (section === 'cash') {
      cashNote = category;
      return;
    }

    const row = {
      description: category.category,
      note: noteNumberByCategory[category.category] || '',
      currentAmount: Number(category.totalCurrent || 0),
      comparativeAmount: Number(category.totalComparative || 0)
    };

    if (section === 'operating_receipt') {
      operatingReceipts.push(row);
      receiptsCurrent += row.currentAmount;
      receiptsComparative += row.comparativeAmount;
      return;
    }

    if (section === 'operating_payment') {
      operatingPayments.push(row);
      paymentsCurrent += row.currentAmount;
      paymentsComparative += row.comparativeAmount;
      return;
    }

    if (section === 'investing') {
      const currentAmount = row.currentAmount === 0 ? 0 : -Math.abs(row.currentAmount);
      const comparativeAmount = row.comparativeAmount === 0 ? 0 : -Math.abs(row.comparativeAmount);
      investingRows.push({
        description: row.description,
        note: row.note,
        currentAmount: currentAmount,
        comparativeAmount: comparativeAmount
      });
      investingCurrent += currentAmount;
      investingComparative += comparativeAmount;
      return;
    }

    if (section === 'financing') {
      financingRows.push(row);
      financingCurrent += row.currentAmount;
      financingComparative += row.comparativeAmount;
    }
  });

  const netOperatingCurrent = receiptsCurrent - paymentsCurrent;
  const netOperatingComparative = receiptsComparative - paymentsComparative;
  const netIncreaseCurrent = netOperatingCurrent + investingCurrent + financingCurrent;
  const netIncreaseComparative = netOperatingComparative + investingComparative + financingComparative;
  const cashOpeningCurrent = cashNote ? Number(cashNote.totalComparative || 0) : 0;
  const cashClosingCurrent = cashNote ? Number(cashNote.totalCurrent || 0) : 0;
  const cashOpeningComparative = 0;
  const cashClosingComparative = cashNote ? Number(cashNote.totalComparative || 0) : 0;

  return {
    currentYear: notes.currentYear || String(currentYear || '').trim(),
    comparativeYear: notes.comparativeYear || String(comparativeYear || '').trim(),
    operating: {
      receipts: operatingReceipts,
      payments: operatingPayments,
      totalReceipts: receiptsCurrent,
      totalReceiptsComparative: receiptsComparative,
      totalPayments: paymentsCurrent,
      totalPaymentsComparative: paymentsComparative,
      netCurrent: netOperatingCurrent,
      netComparative: netOperatingComparative
    },
    investing: {
      rows: investingRows,
      netCurrent: investingCurrent,
      netComparative: investingComparative
    },
    financing: {
      rows: financingRows,
      netCurrent: financingCurrent,
      netComparative: financingComparative
    },
    netIncreaseCurrent: netIncreaseCurrent,
    netIncreaseComparative: netIncreaseComparative,
    cashNote: cashNote ? (noteNumberByCategory[cashNote.category] || '') : '',
    cashOpeningCurrent: cashOpeningCurrent,
    cashOpeningComparative: cashOpeningComparative,
    cashClosingCurrent: cashClosingCurrent,
    cashClosingComparative: cashClosingComparative
  };
}

function getPositionReport(currentYear, comparativeYear) {
  const notes = getNotesReport(currentYear, comparativeYear);
  const categories = Array.isArray(notes.categories) ? notes.categories : [];
  const noteNumberByCategory = {};
  categories.forEach((category, index) => {
    if (category && category.category) {
      noteNumberByCategory[category.category] = index + 1;
    }
  });

  const currentAssets = [];
  const nonCurrentAssets = [];
  const currentLiabilities = [];
  const nonCurrentLiabilities = [];
  const equityRows = [];
  let totalCurrentAssets = 0;
  let totalCurrentAssetsComparative = 0;
  let totalNonCurrentAssets = 0;
  let totalNonCurrentAssetsComparative = 0;
  let totalCurrentLiabilities = 0;
  let totalCurrentLiabilitiesComparative = 0;
  let totalNonCurrentLiabilities = 0;
  let totalNonCurrentLiabilitiesComparative = 0;
  let totalEquity = 0;
  let totalEquityComparative = 0;

  categories.forEach(category => {
    if (!category || !category.category) return;
    const section = _classifyPositionCategory_(
      category.category,
      category.accountTypes || [],
      category.reportMappings || []
    );
    if (!section) return;

    const row = {
      description: category.category,
      note: noteNumberByCategory[category.category] || '',
      currentAmount: Number(category.totalCurrent || 0),
      comparativeAmount: Number(category.totalComparative || 0)
    };

    if (section === 'current_asset') {
      currentAssets.push(row);
      totalCurrentAssets += row.currentAmount;
      totalCurrentAssetsComparative += row.comparativeAmount;
      return;
    }

    if (section === 'non_current_asset') {
      nonCurrentAssets.push(row);
      totalNonCurrentAssets += row.currentAmount;
      totalNonCurrentAssetsComparative += row.comparativeAmount;
      return;
    }

    if (section === 'current_liability') {
      currentLiabilities.push(row);
      totalCurrentLiabilities += row.currentAmount;
      totalCurrentLiabilitiesComparative += row.comparativeAmount;
      return;
    }

    if (section === 'non_current_liability') {
      nonCurrentLiabilities.push(row);
      totalNonCurrentLiabilities += row.currentAmount;
      totalNonCurrentLiabilitiesComparative += row.comparativeAmount;
      return;
    }

    if (section === 'equity') {
      equityRows.push(row);
      totalEquity += row.currentAmount;
      totalEquityComparative += row.comparativeAmount;
    }
  });

  const totalAssetsCurrent = totalCurrentAssets + totalNonCurrentAssets;
  const totalAssetsComparative = totalCurrentAssetsComparative + totalNonCurrentAssetsComparative;
  const totalLiabilitiesCurrent = totalCurrentLiabilities + totalNonCurrentLiabilities;
  const totalLiabilitiesComparative = totalCurrentLiabilitiesComparative + totalNonCurrentLiabilitiesComparative;
  const netAssetsCurrent = totalAssetsCurrent - totalLiabilitiesCurrent;
  const netAssetsComparative = totalAssetsComparative - totalLiabilitiesComparative;
  if (!equityRows.length) {
    totalEquity = netAssetsCurrent;
    totalEquityComparative = netAssetsComparative;
  }

  return {
    currentYear: notes.currentYear || String(currentYear || '').trim(),
    comparativeYear: notes.comparativeYear || String(comparativeYear || '').trim(),
    assets: {
      current: currentAssets,
      nonCurrent: nonCurrentAssets,
      totalCurrent: totalCurrentAssets,
      totalCurrentComparative: totalCurrentAssetsComparative,
      totalNonCurrent: totalNonCurrentAssets,
      totalNonCurrentComparative: totalNonCurrentAssetsComparative,
      total: totalAssetsCurrent,
      totalComparative: totalAssetsComparative
    },
    liabilities: {
      current: currentLiabilities,
      nonCurrent: nonCurrentLiabilities,
      totalCurrent: totalCurrentLiabilities,
      totalCurrentComparative: totalCurrentLiabilitiesComparative,
      totalNonCurrent: totalNonCurrentLiabilities,
      totalNonCurrentComparative: totalNonCurrentLiabilitiesComparative,
      total: totalLiabilitiesCurrent,
      totalComparative: totalLiabilitiesComparative
    },
    equity: {
      rows: equityRows,
      total: totalEquity,
      totalComparative: totalEquityComparative
    },
    netAssets: {
      current: netAssetsCurrent,
      comparative: netAssetsComparative
    }
  };
}

function saveBankStatementUpload(payload) {
  if (!payload) throw new Error('Missing payload.');
  const accountCode = String(payload.accountCode || '').trim();
  if (!accountCode) throw new Error('Bank account is required.');
  const financialYear = String(payload.financialYear || '').trim();
  if (!financialYear) throw new Error('Financial year is required.');
  const rows = payload.rows || [];
  if (!rows.length) throw new Error('No bank rows provided.');

  const ss = _getOrCreateSpreadsheet();
  const sheet = _initializeDbBankSheet(ss);
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  _requireBankHeaders_(headers);
  const headerMap = _getBankHeaderMap_(headers);
  const lastRow = sheet.getLastRow();
  const existingRefs = new Set();
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    data.forEach(function(row) {
      const rowAccount = headerMap.accountCode ? String(row[headerMap.accountCode - 1] || '').trim() : '';
      const rowYear = headerMap.financialYear ? String(row[headerMap.financialYear - 1] || '').trim() : '';
      if (rowAccount !== accountCode || rowYear !== financialYear) return;
      const refValue = headerMap.bankRef ? String(row[headerMap.bankRef - 1] || '').trim() : '';
      if (!refValue) return;
      existingRefs.add(refValue.toLowerCase());
    });
  }

  const output = [];
  let duplicateCount = 0;
  const incomingRefs = new Set();
  rows.forEach(function(item, index) {
    const rowIndex = index + 1;
    const rowAccount = String(item.accountCode || '').trim() || accountCode;
    const rowYear = String(item.financialYear || '').trim() || financialYear;
    if (!rowAccount) throw new Error('Account code missing on row ' + rowIndex + '.');
    if (rowAccount !== accountCode) throw new Error('Account code mismatch on row ' + rowIndex + '.');
    if (!rowYear) throw new Error('Financial year missing on row ' + rowIndex + '.');
    if (rowYear !== financialYear) throw new Error('Financial year mismatch on row ' + rowIndex + '.');

    const description = String(item.description || '').trim();
    if (!description) throw new Error('Description required on row ' + rowIndex + '.');

    const txnDate = _parseBankDate_(item.txnDate, true, rowIndex);
    const valueDate = _parseBankDate_(item.valueDate, false, rowIndex);
    const debit = _parseBankNumber_(item.debit, rowIndex, 'Debit');
    const credit = _parseBankNumber_(item.credit, rowIndex, 'Credit');
    const balance = _parseBankNumber_(item.balance, rowIndex, 'Balance');
    if (debit === '' && credit === '' && balance === '') {
      throw new Error('Debit, credit, or balance is required on row ' + rowIndex + '.');
    }

    const row = new Array(lastCol).fill('');
    const bankRefValue = String(item.bankRef || '').trim();
    const normalizedRef = bankRefValue ? bankRefValue.toLowerCase() : '';
    if (normalizedRef && (existingRefs.has(normalizedRef) || incomingRefs.has(normalizedRef))) {
      duplicateCount += 1;
      return;
    }
    if (normalizedRef) incomingRefs.add(normalizedRef);
    _setIfPresent_(row, headerMap.accountCode, rowAccount);
    _setIfPresent_(row, headerMap.financialYear, rowYear);
    _setIfPresent_(row, headerMap.txnDate, txnDate);
    _setIfPresent_(row, headerMap.valueDate, valueDate);
    _setIfPresent_(row, headerMap.bankRef, bankRefValue);
    _setIfPresent_(row, headerMap.description, description);
    _setIfPresent_(row, headerMap.debit, debit);
    _setIfPresent_(row, headerMap.credit, credit);
    _setIfPresent_(row, headerMap.balance, balance);
    _setIfPresent_(row, headerMap.matchStatus, String(item.matchStatus || '').trim());
    output.push(row);
  });

  if (!output.length && duplicateCount) {
    return { count: 0, duplicateCount: duplicateCount };
  }
  if (!output.length) throw new Error('No valid rows to upload.');

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, output.length, lastCol).setValues(output);
  return { count: output.length, duplicateCount: duplicateCount };
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

function _requireBankHeaders_(headers) {
  const required = [
    'account_code',
    'financial_year',
    'txn_date',
    'value_date',
    'bank_ref',
    'description',
    'debit',
    'credit',
    'balance',
    'match_status'
  ];
  const missing = required.filter(name => headers.indexOf(name) < 0);
  if (missing.length) {
    throw new Error('DB_BANK missing headers: ' + missing.join(', ') + '.');
  }
}

function _getBankHeaderMap_(headers) {
  return {
    accountCode: _resolveColumn_(headers, ['account_code'], 1),
    financialYear: _resolveColumn_(headers, ['financial_year', 'financialyear'], 2),
    txnDate: _resolveColumn_(headers, ['txn_date', 'transaction_date'], 3),
    valueDate: _resolveColumn_(headers, ['value_date'], 4),
    bankRef: _resolveColumn_(headers, ['bank_ref', 'reference'], 5),
    description: _resolveColumn_(headers, ['description'], 6),
    debit: _resolveColumn_(headers, ['debit'], 7),
    credit: _resolveColumn_(headers, ['credit'], 8),
    balance: _resolveColumn_(headers, ['balance'], 9),
    matchStatus: _resolveColumn_(headers, ['match_status'], 10)
  };
}

function _parseBankDate_(value, required, rowIndex) {
  if (value == null || value === '') {
    if (required) throw new Error('Transaction date required on row ' + rowIndex + '.');
    return '';
  }

  if (value instanceof Date) return value;

  if (typeof value === 'number') {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const date = new Date(epoch.getTime() + value * 24 * 60 * 60 * 1000);
    if (!Number.isNaN(date.getTime())) return date;
  }

  const parsed = new Date(value);
  if (!Number.isNaN(parsed.getTime())) return parsed;

  if (required) throw new Error('Invalid date on row ' + rowIndex + '.');
  return '';
}

function _parseBankNumber_(value, rowIndex, label) {
  if (value == null || value === '') return '';
  if (typeof value === 'number') return value;
  const cleaned = String(value).replace(/,/g, '').trim();
  if (!cleaned) return '';
  const parsed = Number(cleaned);
  if (!Number.isFinite(parsed)) {
    throw new Error(label + ' invalid on row ' + rowIndex + '.');
  }
  return parsed;
}

function _getBankColumns_(headers) {
  return {
    accountCode: _resolveColumn_(headers, ['account_code'], 1),
    financialYear: _resolveColumn_(headers, ['financial_year', 'financialyear'], 2),
    txnDate: _resolveColumn_(headers, ['txn_date', 'transaction_date'], 3),
    valueDate: _resolveColumn_(headers, ['value_date'], 4),
    bankRef: _resolveColumn_(headers, ['bank_ref', 'reference'], 5),
    description: _resolveColumn_(headers, ['description'], 6),
    debit: _resolveColumn_(headers, ['debit'], 7),
    credit: _resolveColumn_(headers, ['credit'], 8),
    balance: _resolveColumn_(headers, ['balance'], 9),
    matchStatus: _resolveColumn_(headers, ['match_status'], 10)
  };
}

function _collectReconciliationData_(criteria) {
  const ss = _getOrCreateSpreadsheet();
  const journalSheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  const bankSheet = ss.getSheetByName(CONFIG.SHEETS.DB_BANK);
  if (!journalSheet) throw new Error('DB_JOURNAL not found.');
  if (!bankSheet) throw new Error('DB_BANK not found.');

  const accountCodeFilter = String(criteria && criteria.accountCode || '').trim();
  const financialYearFilter = String(criteria && criteria.financialYear || '').trim();
  const startDate = _parseDate_(criteria && criteria.startDate);
  const endDate = _parseDate_(criteria && criteria.endDate);

  const journalLastRow = journalSheet.getLastRow();
  const journalLastCol = journalSheet.getLastColumn();
  const bankLastRow = bankSheet.getLastRow();
  const bankLastCol = bankSheet.getLastColumn();

  const journalHeaders = journalLastCol
    ? journalSheet.getRange(1, 1, 1, journalLastCol).getValues()[0].map(_normalizeHeader_)
    : [];
  const bankHeaders = bankLastCol
    ? bankSheet.getRange(1, 1, 1, bankLastCol).getValues()[0].map(_normalizeHeader_)
    : [];

  const journalCols = _getJournalColumns_(journalHeaders);
  const bankCols = _getBankColumns_(bankHeaders);

  const journalData = journalLastRow > 1
    ? journalSheet.getRange(2, 1, journalLastRow - 1, journalLastCol).getValues()
    : [];
  const bankData = bankLastRow > 1
    ? bankSheet.getRange(2, 1, bankLastRow - 1, bankLastCol).getValues()
    : [];

  const journalRows = [];
  journalData.forEach(function(row, index) {
    const rowIndex = index + 2;
    const accountCode = journalCols.accountCode ? String(row[journalCols.accountCode - 1] || '').trim() : '';
    if (accountCodeFilter && accountCode !== accountCodeFilter) return;
    const financialYear = journalCols.financialYear ? String(row[journalCols.financialYear - 1] || '').trim() : '';
    if (financialYearFilter && financialYear !== financialYearFilter) return;
    const dateCell = journalCols.date ? row[journalCols.date - 1] : '';
    if (!_isWithinRange_(dateCell, startDate, endDate)) return;

    const reconStatus = journalCols.reconStatus ? row[journalCols.reconStatus - 1] : '';
    if (_isReconciled_(reconStatus)) return;

    const debit = journalCols.debit ? _parseNumber_(row[journalCols.debit - 1]) : 0;
    const credit = journalCols.credit ? _parseNumber_(row[journalCols.credit - 1]) : 0;
    const typeInfo = _getReconType_(debit, credit);
    if (!typeInfo) return;

    const particulars = journalCols.particulars ? String(row[journalCols.particulars - 1] || '').trim() : '';
    if (String(particulars).toLowerCase() === 'accumulated fund') return;

    journalRows.push({
      rowIndex: rowIndex,
      accountCode: accountCode,
      date: _formatDate_(dateCell),
      financialYear: financialYear,
      ref: journalCols.refNo ? String(row[journalCols.refNo - 1] || '').trim() : '',
      payee: journalCols.payee ? String(row[journalCols.payee - 1] || '').trim() : '',
      description: journalCols.description ? String(row[journalCols.description - 1] || '').trim() : '',
      debit: debit,
      credit: credit,
      type: typeInfo.type,
      amount: typeInfo.amount
    });
  });

  const bankRows = [];
  bankData.forEach(function(row, index) {
    const rowIndex = index + 2;
    const accountCode = bankCols.accountCode ? String(row[bankCols.accountCode - 1] || '').trim() : '';
    if (accountCodeFilter && accountCode !== accountCodeFilter) return;
    const dateCell = bankCols.txnDate ? row[bankCols.txnDate - 1] : '';
    if (!_isWithinRange_(dateCell, startDate, endDate)) return;
    const financialYear = bankCols.financialYear ? String(row[bankCols.financialYear - 1] || '').trim() : '';
    if (financialYearFilter && financialYear !== financialYearFilter) return;

    const matchStatus = bankCols.matchStatus ? row[bankCols.matchStatus - 1] : '';
    if (_isReconciled_(matchStatus)) return;

    const debit = bankCols.debit ? _parseNumber_(row[bankCols.debit - 1]) : 0;
    const credit = bankCols.credit ? _parseNumber_(row[bankCols.credit - 1]) : 0;
    const typeInfo = _getReconType_(debit, credit);
    if (!typeInfo) return;

    bankRows.push({
      rowIndex: rowIndex,
      accountCode: accountCode,
      txnDate: _formatDate_(dateCell),
      valueDate: bankCols.valueDate ? _formatDate_(row[bankCols.valueDate - 1]) : '',
      financialYear: financialYear,
      ref: bankCols.bankRef ? String(row[bankCols.bankRef - 1] || '').trim() : '',
      description: bankCols.description ? String(row[bankCols.description - 1] || '').trim() : '',
      debit: debit,
      credit: credit,
      type: typeInfo.type,
      amount: typeInfo.amount
    });
  });

  return {
    journalSheet: journalSheet,
    bankSheet: bankSheet,
    journalCols: journalCols,
    bankCols: bankCols,
    journalRows: journalRows,
    bankRows: bankRows
  };
}

function _matchReconRows_(journalRows, bankRows) {
  const bankMap = {};
  bankRows.forEach(function(row, index) {
    const key = _buildReconKey_(row);
    if (!key) return;
    if (!bankMap[key]) bankMap[key] = [];
    bankMap[key].push(index);
  });

  const matchedJournal = new Set();
  const matchedBank = new Set();
  const pairs = [];

  journalRows.forEach(function(row, index) {
    const key = _buildReconKey_(row);
    if (!key) return;
    const list = bankMap[key];
    if (!list || !list.length) return;
    const bankIndex = list.shift();
    matchedJournal.add(index);
    matchedBank.add(bankIndex);
    pairs.push({ journal: row, bank: bankRows[bankIndex] });
  });

  return { matchedJournal: matchedJournal, matchedBank: matchedBank, pairs: pairs };
}

function _buildReconciliationResponse_(journalRows, bankRows, matchResult, ignoreMatches) {
  const ignore = !!ignoreMatches;
  const cashbookReceipts = [];
  const cashbookPayments = [];
  const bankReceipts = [];
  const bankPayments = [];

  journalRows.forEach(function(row, index) {
    if (!ignore && matchResult.matchedJournal.has(index)) return;
    if (row.type === 'receipt') {
      cashbookReceipts.push(_reconRowSummary_(row));
    } else {
      cashbookPayments.push(_reconRowSummary_(row));
    }
  });

  bankRows.forEach(function(row, index) {
    if (!ignore && matchResult.matchedBank.has(index)) return;
    if (row.type === 'receipt') {
      bankReceipts.push(_reconRowSummary_(row));
    } else {
      bankPayments.push(_reconRowSummary_(row));
    }
  });

  return {
    cashbookReceiptsNotInBank: cashbookReceipts,
    cashbookPaymentsNotInBank: cashbookPayments,
    bankReceiptsNotInCashbook: bankReceipts,
    bankPaymentsNotInCashbook: bankPayments,
    summary: {
      totalCashbook: journalRows.length,
      totalBank: bankRows.length,
      matched: matchResult.pairs.length
    }
  };
}

function _reconRowSummary_(row) {
  return {
    accountCode: row.accountCode || '',
    date: row.date || row.txnDate || '',
    financialYear: row.financialYear || '',
    description: row.description || '',
    ref: row.ref || '',
    debit: row.debit || 0,
    credit: row.credit || 0
  };
}

function _buildReconKey_(row) {
  const ref = _normalizeRef_(row.ref);
  if (!ref) return '';
  const amount = _roundAmount_(row.amount);
  if (!amount) return '';
  const accountCode = String(row.accountCode || '').trim();
  const type = String(row.type || '').trim();
  return [accountCode, type, ref, amount.toFixed(2)].join('|');
}

function _normalizeRef_(value) {
  return String(value || '').trim().toLowerCase();
}

function _roundAmount_(value) {
  const number = Number(value || 0);
  if (!Number.isFinite(number)) return 0;
  return Math.round(number * 100) / 100;
}

function _getReconType_(debit, credit) {
  const debitValue = Number(debit || 0);
  const creditValue = Number(credit || 0);
  if (creditValue > 0 && debitValue === 0) {
    return { type: 'receipt', amount: creditValue };
  }
  if (debitValue > 0 && creditValue === 0) {
    return { type: 'payment', amount: debitValue };
  }
  return null;
}

function _isReconciled_(value) {
  return String(value || '').trim().toLowerCase() === 'reconciled';
}

function _isWithinRange_(value, startDate, endDate) {
  if (!startDate && !endDate) return true;
  if (!value) return true;
  const dateValue = value instanceof Date ? value : _parseDate_(value);
  if (!dateValue) return true;
  if (startDate && dateValue < startDate) return false;
  if (endDate && dateValue > endDate) return false;
  return true;
}

function _sumReconAmounts_(rows) {
  return (rows || []).reduce(function(total, row) {
    const amount = Number(row && row.amount || 0);
    return total + (Number.isFinite(amount) ? amount : 0);
  }, 0);
}

function _sumReconDebitCredit_(rows) {
  return (rows || []).reduce(function(totals, row) {
    const debit = Number(row && row.debit || 0);
    const credit = Number(row && row.credit || 0);
    totals.debit += Number.isFinite(debit) ? debit : 0;
    totals.credit += Number.isFinite(credit) ? credit : 0;
    return totals;
  }, { debit: 0, credit: 0 });
}

function _getReconReportSettings_() {
  const props = PropertiesService.getScriptProperties();
  return {
    entityName: props.getProperty('entityName') || '',
    bankName: props.getProperty('bankName') || '',
    bankBranch: props.getProperty('bankBranch') || '',
    bankAccountNumber: props.getProperty('bankAccountNumber') || ''
  };
}

function _getLatestBankBalance_(sheet, bankCols, criteria) {
  if (!sheet || !bankCols || !bankCols.balance) return 0;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return 0;

  const accountCodeFilter = String(criteria && criteria.accountCode || '').trim();
  const startDate = _parseDate_(criteria && criteria.startDate);
  const endDate = _parseDate_(criteria && criteria.endDate);

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  let latestDate = null;
  let latestRowIndex = -1;
  let latestBalance = 0;

  data.forEach(function(row, index) {
    const accountCode = bankCols.accountCode ? String(row[bankCols.accountCode - 1] || '').trim() : '';
    if (accountCodeFilter && accountCode !== accountCodeFilter) return;
    const dateCell = bankCols.txnDate ? row[bankCols.txnDate - 1] : '';
    if (!_isWithinRange_(dateCell, startDate, endDate)) return;

    const balanceCell = row[bankCols.balance - 1];
    if (balanceCell === '' || balanceCell == null) return;
    const balanceValue = _parseNumber_(balanceCell);
    const dateValue = dateCell instanceof Date ? dateCell : _parseDate_(dateCell);

    if (dateValue) {
      if (!latestDate || dateValue > latestDate || (dateValue.getTime() === latestDate.getTime() && index > latestRowIndex)) {
        latestDate = dateValue;
        latestRowIndex = index;
        latestBalance = balanceValue;
      }
    } else if (!latestDate && index > latestRowIndex) {
      latestRowIndex = index;
      latestBalance = balanceValue;
    }
  });

  return Number(latestBalance || 0);
}

function _writeReconExportSheet_(sheet, payload) {
  const settings = payload.settings || {};
  const criteria = payload.criteria || {};
  const totals = payload.totals || {};
  const bankBalance = Number(payload.bankBalance || 0);
  const cashbookBalance = Number(payload.cashbookBalance || 0);

  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 240);
  sheet.setColumnWidth(6, 110);
  sheet.setColumnWidth(7, 110);

  sheet.getRange('A1').setValue('F.O. 30').setFontWeight('bold');
  _mergeAndSet_(sheet, 'A2:G2', 'REPUBLIC OF KENYA', { bold: true, align: 'center', merge: true });
  _mergeAndSet_(sheet, 'A3:G3', 'BANK RECONCILIATION', { bold: true, align: 'center', merge: true });

  const fromDate = _formatDate_(criteria.startDate);
  const toDate = _formatDate_(criteria.endDate);
  _mergeAndSet_(
    sheet,
    'A4:D4',
    'From Date : ' + (fromDate || '') + ' To : ' + (toDate || ''),
    { merge: true }
  );
  _mergeAndSet_(
    sheet,
    'E4:G4',
    settings.entityName || '',
    { merge: true, align: 'right' }
  );

  const bankLabel = criteria.accountCode || settings.bankName || '';
  _mergeAndSet_(sheet, 'A5:G5', 'Bank : ' + bankLabel, { merge: true });

  _mergeAndSet_(sheet, 'A7:F7', 'Balance as per bank certificate', { merge: true, border: true, bold: true });
  _mergeAndSet_(sheet, 'G7:G7', bankBalance, { merge: true, border: true, align: 'right' })
    .setNumberFormat('#,##0.00');

  sheet.getRange('A8').setValue('Less --').setFontWeight('bold');
  _mergeAndSet_(
    sheet,
    'B9:F9',
    '1. Payments in Cash Book not yet recorded in Bank Statement (Unpresented Cheques)',
    { merge: true, border: true }
  );
  _mergeAndSet_(sheet, 'G9:G9', totals.cashbookPayments || 0, { merge: true, border: true, align: 'right' })
    .setNumberFormat('#,##0.00');

  _mergeAndSet_(
    sheet,
    'B10:F10',
    '2. Receipts in Bank Statement not yet recorded in Cash Book',
    { merge: true, border: true }
  );
  _mergeAndSet_(sheet, 'G10:G10', totals.bankReceipts || 0, { merge: true, border: true, align: 'right' })
    .setNumberFormat('#,##0.00');

  sheet.getRange('A11').setValue('Add --').setFontWeight('bold');
  _mergeAndSet_(
    sheet,
    'B12:F12',
    '3. Payments in Bank Statement not yet recorded in Cash Book',
    { merge: true, border: true }
  );
  _mergeAndSet_(sheet, 'G12:G12', totals.bankPayments || 0, { merge: true, border: true, align: 'right' })
    .setNumberFormat('#,##0.00');

  _mergeAndSet_(
    sheet,
    'B13:F13',
    '4. Receipts in Cash Book not yet recorded in Bank Statement',
    { merge: true, border: true }
  );
  _mergeAndSet_(sheet, 'G13:G13', totals.cashbookReceipts || 0, { merge: true, border: true, align: 'right' })
    .setNumberFormat('#,##0.00');

  _mergeAndSet_(sheet, 'A14:F14', 'Bank Balance as per Cash Book', { merge: true, border: true, bold: true });
  _mergeAndSet_(sheet, 'G14:G14', cashbookBalance, { merge: true, border: true, align: 'right' })
    .setNumberFormat('#,##0.00');

  let row = 17;
  row = _writeReconSection_(
    sheet,
    row,
    '1. PAYMENTS IN CASH BOOK NOT YET RECORDED IN BANK STATEMENT (UNPRESENTED CHEQUES)',
    'Ref_No',
    payload.cashbookPayments || []
  );
  row += 1;
  row = _writeReconSection_(
    sheet,
    row,
    '2. RECEIPTS IN BANK STATEMENT NOT YET RECORDED IN CASH BOOK',
    'Bank_Ref',
    payload.bankReceipts || []
  );
  row += 1;
  row = _writeReconSection_(
    sheet,
    row,
    '3. PAYMENTS IN BANK STATEMENT NOT YET RECORDED IN CASH BOOK',
    'Bank_Ref',
    payload.bankPayments || []
  );
  row += 1;
  row = _writeReconSection_(
    sheet,
    row,
    '4. RECEIPTS IN CASH BOOK NOT YET RECORDED IN BANK STATEMENT',
    'Ref_No',
    payload.cashbookReceipts || []
  );

  const signaturesRow = row + 2;
  _mergeAndSet_(
    sheet,
    'A' + signaturesRow + ':G' + signaturesRow,
    'Reconciled by : ............................ Signature: ............................ Date: ............................',
    { merge: true }
  );
  _mergeAndSet_(
    sheet,
    'A' + (signaturesRow + 2) + ':G' + (signaturesRow + 2),
    'Reviewed by : .............................. Signature: ............................ Date: ............................',
    { merge: true }
  );
  _mergeAndSet_(
    sheet,
    'A' + (signaturesRow + 4) + ':G' + (signaturesRow + 4),
    'Approved by : .............................. Signature: ............................ Date: ............................',
    { merge: true }
  );
}


function _writeReconSection_(sheet, startRow, title, refLabel, rows) {
  const titleRange = sheet.getRange(startRow, 1, 1, 7);
  titleRange.merge();
  titleRange.setValue(title).setFontWeight('bold');
  titleRange.setBorder(true, true, true, true, true, true);

  const headerRow = startRow + 1;
  sheet.getRange(headerRow, 1).setValue('Account_Code').setFontWeight('bold');
  sheet.getRange(headerRow, 2).setValue('Financial_Year').setFontWeight('bold');
  sheet.getRange(headerRow, 3).setValue('Date').setFontWeight('bold');
  sheet.getRange(headerRow, 4).setValue(refLabel).setFontWeight('bold');
  sheet.getRange(headerRow, 5).setValue('Description').setFontWeight('bold');
  sheet.getRange(headerRow, 6).setValue('Debit').setFontWeight('bold');
  sheet.getRange(headerRow, 7).setValue('Credit').setFontWeight('bold');
  sheet.getRange(headerRow, 1, 1, 7).setBorder(true, true, true, true, true, true);

  let rowIndex = startRow + 2;
  if (!rows || !rows.length) {
    sheet.getRange(rowIndex, 1, 1, 5).merge().setValue('No unreconciled data').setHorizontalAlignment('left');
    sheet.getRange(rowIndex, 6).setValue(0).setNumberFormat('#,##0.00');
    sheet.getRange(rowIndex, 7).setValue(0).setNumberFormat('#,##0.00');
    sheet.getRange(rowIndex, 1, 1, 7).setBorder(true, true, true, true, true, true);
    rowIndex += 1;
  } else {
    rows.forEach(function(row) {
      sheet.getRange(rowIndex, 1).setValue(row.accountCode || '');
      sheet.getRange(rowIndex, 2).setValue(row.financialYear || '');
      sheet.getRange(rowIndex, 3).setValue(row.date || row.txnDate || '');
      sheet.getRange(rowIndex, 4).setValue(row.ref || '');
      sheet.getRange(rowIndex, 5).setValue(row.description || '');
      sheet.getRange(rowIndex, 6).setValue(Number(row.debit || 0));
      sheet.getRange(rowIndex, 7).setValue(Number(row.credit || 0));
      sheet.getRange(rowIndex, 1, 1, 7).setBorder(true, true, true, true, true, true);
      sheet.getRange(rowIndex, 6, 1, 2).setNumberFormat('#,##0.00');
      rowIndex += 1;
    });
  }

  const totalRow = rowIndex;
  const totals = _sumReconDebitCredit_(rows);
  sheet.getRange(totalRow, 1, 1, 5).merge().setValue('Total :').setHorizontalAlignment('right').setFontWeight('bold');
  sheet.getRange(totalRow, 6).setValue(Number(totals.debit || 0)).setFontWeight('bold').setNumberFormat('#,##0.00');
  sheet.getRange(totalRow, 7).setValue(Number(totals.credit || 0)).setFontWeight('bold').setNumberFormat('#,##0.00');
  sheet.getRange(totalRow, 1, 1, 7).setBorder(true, true, true, true, true, true);

  return totalRow + 1;
}

function _mergeAndSet_(sheet, rangeA1, value, options) {
  const range = sheet.getRange(rangeA1);
  if (options && options.merge) range.merge();
  range.setValue(value);
  if (options && options.bold) range.setFontWeight('bold');
  if (options && options.align) range.setHorizontalAlignment(options.align);
  if (options && options.border) range.setBorder(true, true, true, true, true, true);
  return range;
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

function _resolveNotesAmount_(bucket, particulars, accountType) {
  const debit = bucket.debit[particulars] || 0;
  const credit = bucket.credit[particulars] || 0;
  const accountTypeValue = String(accountType || '').toLowerCase();
  const isReceipt = accountTypeValue.includes('income') || (!accountTypeValue && credit > debit);
  return isReceipt ? credit : debit;
}

function _classifyPerformanceCategory_(categoryName, accountTypes) {
  const name = String(categoryName || '').toLowerCase();
  const normalizedTypes = (accountTypes || []).map(type => String(type || '').toLowerCase());

  if (name.includes('gain') || name.includes('loss')) return 'other';
  if (normalizedTypes.some(type => type.includes('income') || type.includes('revenue'))) return 'revenue';
  if (normalizedTypes.some(type => type.includes('expense'))) return 'expense';
  if (normalizedTypes.some(type => type.includes('gain') || type.includes('loss'))) return 'other';
  if (name.includes('income') || name.includes('revenue')) return 'revenue';
  if (name.includes('expense')) return 'expense';
  return null;
}

function _classifyCashFlowCategory_(categoryName, accountTypes, reportMappings) {
  const name = String(categoryName || '').toLowerCase();
  const typeValues = (accountTypes || []).map(type => String(type || '').toLowerCase());
  const mappingValues = (reportMappings || []).map(value => String(value || '').toLowerCase());

  if (name.includes('cash and cash equivalent')) return 'cash';

  if (mappingValues.some(value => value.includes('operating income'))) return 'operating_receipt';
  if (mappingValues.some(value => value.includes('operating expense'))) return 'operating_payment';

  if (typeValues.some(value => value.includes('income')) || name.includes('income') || name.includes('revenue')) {
    return 'operating_receipt';
  }

  if (typeValues.some(value => value.includes('expense')) || name.includes('expense')) {
    return 'operating_payment';
  }

  if (mappingValues.some(value => value.includes('non-current asset') || value.includes('non current asset'))) {
    return 'investing';
  }

  if (mappingValues.some(value => value.includes('current liability') || value.includes('non-current liability') || value.includes('non current liability'))) {
    return 'financing';
  }

  if (name.includes('liabilit')) return 'financing';
  return null;
}

function _classifyPositionCategory_(categoryName, accountTypes, reportMappings) {
  const name = String(categoryName || '').toLowerCase();
  const typeValues = (accountTypes || []).map(type => String(type || '').toLowerCase());
  const mappingValues = (reportMappings || []).map(value => String(value || '').toLowerCase());

  if (name.includes('cash and cash equivalent')) return 'current_asset';
  if (mappingValues.some(value => value.includes('current asset'))) return 'current_asset';
  if (mappingValues.some(value => value.includes('non-current asset') || value.includes('non current asset'))) {
    return 'non_current_asset';
  }
  if (mappingValues.some(value => value.includes('current liability'))) return 'current_liability';
  if (mappingValues.some(value => value.includes('non-current liability') || value.includes('non current liability'))) {
    return 'non_current_liability';
  }

  if (typeValues.some(value => value.includes('asset'))) return 'current_asset';
  if (typeValues.some(value => value.includes('liabil'))) return 'current_liability';
  if (typeValues.some(value => value.includes('equity'))) return 'equity';

  if (name.includes('equity') || name.includes('net asset') || name.includes('reserve') || name.includes('surplus') || name.includes('deficit')) {
    return 'equity';
  }

  return null;
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
