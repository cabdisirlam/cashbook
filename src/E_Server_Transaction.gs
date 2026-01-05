/**
 * Payment modal backend helpers.
 */
function getDropdownData() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'dropdownData_v3';
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (error) {
      Logger.log('Cache parse error: ' + error.toString());
    }
  }

  const payload = {
    accounts: [],
    payees: [],
    payeeTypes: [],
    payeesByType: {},
    particulars: [],
    subCats: [],
    subToCatMap: {},
    particularMeta: {},
    categories: [],
    financialYears: []
  };

  const ss = _getOrCreateSpreadsheet();

  // Get master data for accounts, particulars, categories, etc.
  const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);
  if (sheet) {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow >= 2) {
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
      const cols = _getMasterColumns_(headers);
      const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

      data.forEach(function(row) {
        const account = cols.accountCodes ? String(row[cols.accountCodes - 1]).trim() : '';
        const particulars = cols.particulars ? String(row[cols.particulars - 1]).trim() : '';
        const subCat = cols.subCategory ? String(row[cols.subCategory - 1]).trim() : '';
        const category = cols.category ? String(row[cols.category - 1]).trim() : '';
        const financialYear = cols.financialYear ? String(row[cols.financialYear - 1]).trim() : '';
        const accountType = cols.accountType ? String(row[cols.accountType - 1]).trim() : '';
        const reportMapping = cols.reportMapping ? String(row[cols.reportMapping - 1]).trim() : '';

        if (account) payload.accounts.push(account);
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
    }
  }

  // Get payees from CONTACTS sheet (Suppliers, Customers, Staff, Government Entities, Donors)
  const contactsSheet = ss.getSheetByName(CONFIG.SHEETS.CONTACTS);
  if (contactsSheet) {
    const contactsLastRow = contactsSheet.getLastRow();
    if (contactsLastRow >= 2) {
      // CONTACTS columns: Contact_ID, Contact_Type, Contact_Name, ...
      const contactsData = contactsSheet.getRange(2, 1, contactsLastRow - 1, 3).getValues();
      contactsData.forEach(function(row) {
        const contactId = String(row[0] || '').trim();
        const contactType = _normalizeContactType_(row[1]);
        const contactName = String(row[2] || '').trim();
        // Only add active contacts with valid names
        if (contactId && contactName) {
          payload.payees.push(contactName);
          if (contactType) {
            payload.payeeTypes.push(contactType);
            if (!payload.payeesByType[contactType]) payload.payeesByType[contactType] = [];
            payload.payeesByType[contactType].push(contactName);
          }
        }
      });
    }
  }

  payload.accounts = _uniqueSorted_(payload.accounts);
  payload.payees = _uniqueSorted_(payload.payees);
  payload.payeeTypes = _uniqueSorted_(payload.payeeTypes);
  Object.keys(payload.payeesByType).forEach(function(key) {
    payload.payeesByType[key] = _uniqueSorted_(payload.payeesByType[key]);
  });
  payload.particulars = _uniqueSorted_(payload.particulars);
  payload.subCats = _uniqueSorted_(payload.subCats);
  payload.categories = _uniqueSorted_(payload.categories);
  payload.financialYears = _uniqueSorted_(payload.financialYears);

  cache.put(cacheKey, JSON.stringify(payload), 300);
  return payload;
}

function getFinancialYears() {
  const dropdowns = getDropdownData();
  return dropdowns && dropdowns.financialYears ? dropdowns.financialYears : [];
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
  const bankRef = String(header.bankRef || '').trim();
  const bankParticulars = String(header.bankParticulars || '').trim();
  const bankDescriptionInput = String(header.bankDescription || '').trim();

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
  const accountMeta = _buildAccountMeta_(masterData, masterCols);

  const contactsSheet = ss.getSheetByName(CONFIG.SHEETS.CONTACTS);
  const contactMapByType = {};
  const contactMapByName = {};
  if (contactsSheet) {
    const contactsLastRow = contactsSheet.getLastRow();
    if (contactsLastRow >= 2) {
      const contactsData = contactsSheet.getRange(2, 1, contactsLastRow - 1, 3).getValues();
      contactsData.forEach(function(row) {
        const contactId = String(row[0] || '').trim();
        const contactType = _normalizeContactType_(row[1]);
        const contactName = String(row[2] || '').trim();
        if (!contactId || !contactName) return;
        if (contactType) {
          if (!contactMapByType[contactType]) contactMapByType[contactType] = {};
          contactMapByType[contactType][contactName.toLowerCase()] = contactId;
        }
        const nameKey = contactName.toLowerCase();
        if (!contactMapByName[nameKey]) {
          contactMapByName[nameKey] = contactId;
        }
      });
    }
  }

  if (!bankParticulars) throw new Error('Bank particulars are required.');
  const bankParticularMeta = particularMeta[bankParticulars];
  if (!bankParticularMeta || !bankParticularMeta.subCategory || !bankParticularMeta.category) {
    throw new Error('Bank particulars not found: ' + bankParticulars + '.');
  }

  const cleanedRows = rows.map(function(row) {
    const amount = Number(row.amount || 0);
    const particulars = String(row.particulars || '').trim();
    const rowPayee = String(row.payee || '').trim();
    const rowPayeeType = String(row.payeeType || '').trim();
    const description = String(row.description || '').trim();

    if (!particulars) throw new Error('Each line needs particulars.');
    if (!rowPayee) throw new Error('Each line needs a payee.');
    const meta = particularMeta[particulars];
    if (!meta || !meta.subCategory || !meta.category) {
      throw new Error('Particulars not found: ' + particulars + '.');
    }
    const subCategory = meta.subCategory;
    const category = meta.category;
    if (amount <= 0) throw new Error('Line amount must be greater than zero.');

    const accountTypeValue = meta.accountType || '';
    const reportMappingValue = meta.reportMapping || '';

    const payeeTypeKey = _normalizeContactType_(rowPayeeType);
    const contactId = payeeTypeKey && contactMapByType[payeeTypeKey]
      ? (contactMapByType[payeeTypeKey][rowPayee.toLowerCase()] || '')
      : (contactMapByName[rowPayee.toLowerCase()] || '');

    return {
      particulars,
      subCategory,
      category,
      payee: rowPayee,
      contactId: contactId,
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
  // For staff advance payments, the batchId IS the advanceId
  const advanceId = batchId;

  const entries = cleanedRows.map(function(row) {
    const debitValue = isReceipt ? 0 : row.amount;
    const creditValue = isReceipt ? row.amount : 0;
    return [
      Utilities.getUuid(),
      batchId,
      dateValue,
      financialYear,
      '',
      row.contactId || '',
      row.payee,
      refNo,
      bankRef,
      row.particulars,
      row.subCategory,
      row.category,
      row.description,
      debitValue,
      creditValue,
      row.accountType,
      row.reportMapping,
      '',
      '',
      advanceId
    ];
  });

  const bankSubCategory = bankParticularMeta.subCategory || '';
  const bankCategory = bankParticularMeta.category || '';
  const bankDescription = bankDescriptionInput || 'Bank entry';
  const bankAccountType = bankParticularMeta.accountType || '';
  const bankReportMapping = bankParticularMeta.reportMapping || '';
  const bankDebit = isReceipt ? total : 0;
  const bankCredit = isReceipt ? 0 : total;
  const bankPayee = _resolveBatchPayee_(cleanedRows, payee);
  const bankContactId = cleanedRows[0] ? (cleanedRows[0].contactId || '') : '';
  entries.push([
    Utilities.getUuid(),
    batchId,
    dateValue,
    financialYear,
    accountCode,
    bankContactId,
    bankPayee,
    refNo,
    bankRef,
    bankParticulars,
    bankSubCategory,
    bankCategory,
    bankDescription,
    bankDebit,
    bankCredit,
    bankAccountType,
    bankReportMapping,
    'Unreconciled',
    '',
    advanceId
  ]);

  const startRow = journal.getLastRow() + 1;
  journal.getRange(startRow, 1, entries.length, entries[0].length).setValues(entries);

  logSystemEventSafe(
    'CREATE_JOURNAL',
    batchId,
    'Type: ' + type + ', Rows: ' + entries.length + ', Total: ' + total
  );

  return 'Success';
}

function _resolveBatchPayee_(rows, headerPayee) {
  const unique = {};
  (rows || []).forEach(function(row) {
    const value = String(row.payee || '').trim();
    if (value) unique[value] = true;
  });
  const values = Object.keys(unique);
  if (values.length === 1) return values[0];
  if (headerPayee) return headerPayee;
  return 'Multiple Payees';
}

function getNextJournalRef() {
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) throw new Error('DB_JOURNAL not found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return 'J1000001';

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getJournalColumns_(headers);
  if (!cols.refNo) return 'J1000001';

  const values = sheet.getRange(2, cols.refNo, lastRow - 1, 1).getValues();
  let maxNumber = 1000000;
  values.forEach(function(row) {
    const value = String(row[0] || '').trim().toUpperCase();
    const match = /^J(\d{7})$/.exec(value);
    if (!match) return;
    const numeric = Number(match[1]);
    if (Number.isFinite(numeric) && numeric > maxNumber) {
      maxNumber = numeric;
    }
  });

  return 'J' + String(maxNumber + 1).padStart(7, '0');
}

function saveJournalEntry(payload) {
  if (!payload || !payload.header || !payload.rows || !payload.rows.length) {
    throw new Error('Missing journal data.');
  }

  const header = payload.header;
  const rows = payload.rows;
  const dateValue = header.date ? new Date(header.date) : new Date();
  if (Number.isNaN(dateValue.getTime())) throw new Error('Invalid date.');

  const financialYear = String(header.financialYear || '').trim();
  const payee = String(header.payee || '').trim();
  const linkedAdvanceId = String(header.advanceId || '').trim();
  if (!financialYear) throw new Error('Financial year is required.');

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
  const accountMeta = _buildAccountMeta_(masterData, masterCols);

  const cleanedRows = rows.map(function(row) {
    const lineType = String(row.lineType || '').trim().toLowerCase();
    const particulars = String(row.particulars || '').trim();
    const accountCode = String(row.accountCode || '').trim();
    const description = String(row.description || '').trim();
    const debitValue = Number(row.debit || 0);
    const creditValue = Number(row.credit || 0);

    if (debitValue < 0 || creditValue < 0) throw new Error('Debit and Credit must be positive.');
    if (debitValue === 0 && creditValue === 0) {
      throw new Error('Each line needs a debit or credit value.');
    }
    if (debitValue > 0 && creditValue > 0) {
      throw new Error('Each line can only have a debit or credit amount.');
    }

    if (lineType === 'bank' || accountCode) {
      if (!accountCode) throw new Error('Bank account is required for bank lines.');
      const bankMeta = accountMeta[accountCode] || {};
      return {
        accountCode: accountCode,
        particulars: '',
        subCategory: '',
        category: '',
        description: description,
        debit: debitValue,
        credit: creditValue,
        accountType: bankMeta.accountType || '',
        reportMapping: bankMeta.reportMapping || '',
        reconStatus: 'Unreconciled'
      };
    }

    if (!particulars) throw new Error('Each line needs particulars.');
    const meta = particularMeta[particulars];
    if (!meta || !meta.subCategory || !meta.category) {
      throw new Error('Particulars not found: ' + particulars + '.');
    }

    return {
      accountCode: '',
      particulars: particulars,
      subCategory: meta.subCategory,
      category: meta.category,
      description: description,
      debit: debitValue,
      credit: creditValue,
      accountType: meta.accountType || '',
      reportMapping: meta.reportMapping || '',
      reconStatus: ''
    };
  });

  const totals = cleanedRows.reduce(function(acc, row) {
    acc.debit += row.debit;
    acc.credit += row.credit;
    return acc;
  }, { debit: 0, credit: 0 });

  if (totals.debit <= 0 || totals.credit <= 0) {
    throw new Error('Journal entries must have both debit and credit totals.');
  }
  if (Math.abs(totals.debit - totals.credit) > 0.01) {
    throw new Error('Journal entry is out of balance.');
  }

  const fallbackDetail = cleanedRows.find(row => !row.accountCode) || null;

  let refNo = String(header.refNo || '').trim();
  if (!refNo) {
    refNo = getNextJournalRef();
  }

  const contactId = String(header.contactId || '').trim() || (payee ? (contactMapByName[payee.toLowerCase()] || '') : '');
  const batchId = 'JRN-' + new Date().getTime();
  const entries = cleanedRows.map(function(row) {
    const useFallback = row.accountCode && fallbackDetail;
    const entryParticulars = useFallback ? fallbackDetail.particulars : row.particulars;
    const entrySubCategory = useFallback ? fallbackDetail.subCategory : row.subCategory;
    const entryCategory = useFallback ? fallbackDetail.category : row.category;
    const entryDescription = useFallback ? (fallbackDetail.description || row.description) : row.description;
    const entryAccountType = useFallback ? (fallbackDetail.accountType || '') : row.accountType;
    const entryReportMapping = useFallback ? (fallbackDetail.reportMapping || '') : row.reportMapping;
    return [
      Utilities.getUuid(),
      batchId,
      dateValue,
      financialYear,
      row.accountCode || '',
      contactId,
      payee,
      refNo,
      '',
      entryParticulars,
      entrySubCategory,
      entryCategory,
      entryDescription,
      row.debit,
      row.credit,
      entryAccountType,
      entryReportMapping,
      row.reconStatus || '',
      '',
      linkedAdvanceId
    ];
  });

  const startRow = journal.getLastRow() + 1;
  journal.getRange(startRow, 1, entries.length, entries[0].length).setValues(entries);

  logSystemEventSafe(
    'CREATE_JOURNAL',
    batchId,
    'Type: Journal, Ref: ' + refNo + ', Rows: ' + entries.length
  );

  return { success: true, refNo: refNo };
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
  const added = {};
  const batchInfo = {};

  for (let i = data.length - 1; i >= 0 && results.length < maxRows; i--) {
    const row = data[i];
    const batchId = cols.batchId ? String(row[cols.batchId - 1] || '').trim() : '';
    if (!batchId.startsWith('TXN-')) continue;

    const info = batchInfo[batchId] || {
      batchId: batchId,
      date: '',
      payee: '',
      refNo: '',
      particulars: '',
      subCategory: '',
      category: '',
      bankDebit: 0,
      bankCredit: 0,
      hasBank: false
    };

    const accountCode = cols.accountCode ? String(row[cols.accountCode - 1] || '').trim() : '';
    const particulars = cols.particulars ? String(row[cols.particulars - 1] || '').trim() : '';
    if (accountCode) {
      const debit = cols.debit ? _parseNumber_(row[cols.debit - 1]) : 0;
      const credit = cols.credit ? _parseNumber_(row[cols.credit - 1]) : 0;
      if (!info.hasBank) {
        const dateValue = cols.date ? row[cols.date - 1] : '';
        info.date = dateValue ? _formatDate_(dateValue) : '';
        info.payee = cols.payee ? row[cols.payee - 1] : info.payee;
        info.refNo = cols.refNo ? row[cols.refNo - 1] : '';
        info.bankDebit = debit;
        info.bankCredit = credit;
        info.hasBank = true;
      }
    } else if (particulars && !info.particulars) {
      info.particulars = particulars;
      info.subCategory = cols.subCategory ? row[cols.subCategory - 1] : '';
      info.category = cols.category ? row[cols.category - 1] : '';
      if (!info.payee) info.payee = cols.payee ? row[cols.payee - 1] : '';
    }

    batchInfo[batchId] = info;

    if (!info.hasBank || !info.particulars || added[batchId]) continue;
    const direction = info.bankDebit > 0 ? 'receipt' : (info.bankCredit > 0 ? 'payment' : '');
    if (!direction) continue;
    if (requested && requested !== direction) continue;

    results.push({
      batchId: info.batchId,
      date: info.date,
      financialYear: cols.financialYear ? row[cols.financialYear - 1] : '',
      payee: info.payee,
      refNo: info.refNo,
      particulars: info.particulars,
      subCategory: info.subCategory,
      category: info.category,
      amount: direction === 'receipt' ? info.bankDebit : info.bankCredit
    });
    added[batchId] = true;
  }

  return results;
}

function getRecentJournalEntries(limit) {
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) throw new Error('DB_JOURNAL not found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getJournalColumns_(headers);
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const maxRows = Math.max(1, Number(limit) || 5);
  const results = [];

  for (let i = data.length - 1; i >= 0 && results.length < maxRows; i--) {
    const row = data[i];
    const batchId = cols.batchId ? String(row[cols.batchId - 1] || '').trim() : '';
    if (!batchId.startsWith('JRN-')) continue;
    const particulars = cols.particulars ? String(row[cols.particulars - 1] || '').trim() : '';
    if (!particulars) continue;
    const dateValue = cols.date ? row[cols.date - 1] : '';
    results.push({
      batchId: batchId,
      date: dateValue ? _formatDate_(dateValue) : '',
      payee: cols.payee ? row[cols.payee - 1] : '',
      refNo: cols.refNo ? row[cols.refNo - 1] : '',
      particulars: particulars,
      subCategory: cols.subCategory ? row[cols.subCategory - 1] : '',
      category: cols.category ? row[cols.category - 1] : '',
      debit: cols.debit ? row[cols.debit - 1] : '',
      credit: cols.credit ? row[cols.credit - 1] : ''
    });
  }

  return results;
}

function getAdvanceSurrenderTotals(batchIds) {
  const ids = Array.isArray(batchIds) ? batchIds.map(String) : [];
  const targets = {};
  ids.forEach(function(id) {
    const trimmed = String(id || '').trim();
    if (trimmed) targets[trimmed] = true;
  });
  if (!Object.keys(targets).length) return {};

  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) throw new Error('DB_JOURNAL not found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return {};

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getJournalColumns_(headers);
  if (!cols.batchId) return {};

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const taggedBatches = {};

  data.forEach(function(row) {
    const journalBatch = String(row[cols.batchId - 1] || '').trim();
    if (!journalBatch || !journalBatch.startsWith('JRN-')) return;

    // Check Advance_ID column first (new method)
    let advanceId = cols.advanceId ? String(row[cols.advanceId - 1] || '').trim() : '';

    // Fallback to description pattern (old method) for backwards compatibility
    if (!advanceId && cols.description) {
      const description = String(row[cols.description - 1] || '');
      const match = /ADV:(TXN-\d+)/.exec(description);
      if (match) advanceId = match[1];
    }

    if (!advanceId || !targets[advanceId]) return;

    if (!taggedBatches[advanceId]) taggedBatches[advanceId] = {};
    if (!taggedBatches[advanceId][journalBatch]) taggedBatches[advanceId][journalBatch] = { debit: 0, credit: 0 };
    const debit = cols.debit ? _parseNumber_(row[cols.debit - 1]) : 0;
    const credit = cols.credit ? _parseNumber_(row[cols.credit - 1]) : 0;
    taggedBatches[advanceId][journalBatch].debit += debit;
    taggedBatches[advanceId][journalBatch].credit += credit;
  });

  const totals = {};
  Object.keys(taggedBatches).forEach(function(advanceId) {
    const batches = taggedBatches[advanceId];
    let sum = 0;
    Object.keys(batches).forEach(function(batchId) {
      const totalsForBatch = batches[batchId];
      const amount = Math.max(totalsForBatch.debit, totalsForBatch.credit);
      sum += amount;
    });
    totals[advanceId] = sum;
  });

  return totals;
}

function getAdvanceBankClassification(advanceId) {
  const target = String(advanceId || '').trim();
  if (!target) return {};

  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) throw new Error('DB_JOURNAL not found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return {};

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getJournalColumns_(headers);
  if (!cols.batchId) return {};

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const batchId = String(row[cols.batchId - 1] || '').trim();
    if (batchId !== target) continue;
    const accountCode = cols.accountCode ? String(row[cols.accountCode - 1] || '').trim() : '';
    if (!accountCode) continue;
    return {
      particulars: cols.particulars ? String(row[cols.particulars - 1] || '').trim() : '',
      subCategory: cols.subCategory ? String(row[cols.subCategory - 1] || '').trim() : '',
      category: cols.category ? String(row[cols.category - 1] || '').trim() : '',
      description: cols.description ? String(row[cols.description - 1] || '').trim() : ''
    };
  }

  return {};
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
    const bankRef = cols.bankRef ? String(row[cols.bankRef - 1]).trim() : '';
    if (refOrId && !(uuid.toLowerCase().includes(refOrId) || refNo.toLowerCase().includes(refOrId) || bankRef.toLowerCase().includes(refOrId))) return;

    const haystack = [
      cols.accountCode ? row[cols.accountCode - 1] : '',
      payee,
      category,
      particulars,
      cols.subCategory ? row[cols.subCategory - 1] : '',
      cols.description ? row[cols.description - 1] : '',
      refNo,
      bankRef
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
      bankRef: cols.bankRef ? row[cols.bankRef - 1] : '',
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
    bankRef: cols.bankRef ? values[cols.bankRef - 1] || '' : '',
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

function getJournalBatch(batchId) {
  const batch = String(batchId || '').trim();
  if (!batch) throw new Error('Invalid batch.');
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) throw new Error('DB_JOURNAL not found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getJournalColumns_(headers);
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const results = [];

  data.forEach(function(row, index) {
    const rowBatch = cols.batchId ? String(row[cols.batchId - 1] || '').trim() : '';
    if (rowBatch !== batch) return;

    const rowIndex = index + 2;
    const dateValue = row[cols.date - 1];

    results.push({
      rowId: rowIndex,
      uuid: row[cols.uuid - 1] || '',
      batchId: rowBatch,
      date: dateValue instanceof Date ? _formatDate_(dateValue) : '',
      financialYear: cols.financialYear ? row[cols.financialYear - 1] : '',
      accountCode: cols.accountCode ? row[cols.accountCode - 1] : '',
      payee: cols.payee ? row[cols.payee - 1] : '',
      refNo: cols.refNo ? row[cols.refNo - 1] : '',
      bankRef: cols.bankRef ? row[cols.bankRef - 1] || '' : '',
      particulars: cols.particulars ? row[cols.particulars - 1] || '' : '',
      subCategory: cols.subCategory ? row[cols.subCategory - 1] || '' : '',
      category: cols.category ? row[cols.category - 1] || '' : '',
      description: cols.description ? row[cols.description - 1] || '' : '',
      debit: cols.debit ? row[cols.debit - 1] || '' : '',
      credit: cols.credit ? row[cols.credit - 1] || '' : '',
      accountType: cols.accountType ? row[cols.accountType - 1] || '' : '',
      reportMapping: cols.reportMapping ? row[cols.reportMapping - 1] || '' : '',
      reconStatus: cols.reconStatus ? row[cols.reconStatus - 1] || '' : '',
      receiptUrl: cols.receiptUrl ? row[cols.receiptUrl - 1] || '' : ''
    });
  });

  return results;
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
  _setIfPresent_(updated, cols.bankRef, payload.bankRef || '');
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

  const applyToBatch = !!payload.applyToBatch;
  const batchId = String(payload.batchId || '').trim();
  if (applyToBatch && batchId && cols.batchId) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      let touched = false;
      const newDate = _parseDate_(payload.date);
      data.forEach(function(rowValues) {
        const rowBatch = String(rowValues[cols.batchId - 1] || '').trim();
        if (rowBatch !== batchId) return;
        _setIfPresent_(rowValues, cols.date, newDate || rowValues[cols.date - 1]);
        _setIfPresent_(rowValues, cols.financialYear, payload.financialYear || rowValues[cols.financialYear - 1] || '');
        _setIfPresent_(rowValues, cols.payee, payload.payee || rowValues[cols.payee - 1] || '');
        _setIfPresent_(rowValues, cols.refNo, payload.refNo || rowValues[cols.refNo - 1] || '');
        _setIfPresent_(rowValues, cols.bankRef, payload.bankRef || rowValues[cols.bankRef - 1] || '');
        touched = true;
      });
      if (touched) {
        sheet.getRange(2, 1, lastRow - 1, lastCol).setValues(data);
      }
    }
  }

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

function deleteJournalBatch(batchId) {
  const batch = String(batchId || '').trim();
  if (!batch) throw new Error('Invalid batch.');
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) throw new Error('DB_JOURNAL not found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return 0;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getJournalColumns_(headers);
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rowsToDelete = [];

  data.forEach(function(row, index) {
    const rowBatch = cols.batchId ? String(row[cols.batchId - 1] || '').trim() : '';
    if (rowBatch === batch) {
      rowsToDelete.push({ rowIndex: index + 2, values: row });
    }
  });

  if (!rowsToDelete.length) return 0;

  rowsToDelete.sort((a, b) => b.rowIndex - a.rowIndex).forEach(function(entry) {
    _storeUndoAction_('delete', entry.rowIndex, entry.values);
    sheet.deleteRow(entry.rowIndex);
  });

  try {
    const user = getCurrentUser();
    const actor = user && user.authenticated ? user.email : 'system';
    logSystemEvent(actor, 'DELETE_JOURNAL_BATCH', batch, 'Rows ' + rowsToDelete.length);
  } catch (error) {
    Logger.log('Log failure: ' + error.toString());
  }

  return rowsToDelete.length;
}

function reverseJournalRow(rowId) {
  const row = Number(rowId);
  if (!row || row < 2) throw new Error('Invalid row.');
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) throw new Error('DB_JOURNAL not found.');

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getJournalColumns_(headers);
  const existing = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  const reversalBatchId = 'REV-' + new Date().getTime();
  const dateValue = new Date();

  const reversed = _buildReversalRow_(existing, cols, reversalBatchId, dateValue);
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, lastCol).setValues([reversed]);

  const actor = _resolveActorEmail_();
  logSystemEvent(actor, 'REVERSE_JOURNAL_ROW', String(reversed[cols.uuid - 1] || ''), 'Row ' + row);
  return { count: 1, batchId: reversalBatchId };
}

function reverseJournalBatch(batchId) {
  const batch = String(batchId || '').trim();
  if (!batch) throw new Error('Invalid batch.');
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) throw new Error('DB_JOURNAL not found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) throw new Error('No rows to reverse.');

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getJournalColumns_(headers);
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const matching = data.filter(row => String(row[cols.batchId - 1] || '').trim() === batch);
  if (!matching.length) throw new Error('Batch not found.');

  const reversalBatchId = 'REV-' + new Date().getTime();
  const dateValue = new Date();
  const reversedRows = matching.map(row => _buildReversalRow_(row, cols, reversalBatchId, dateValue));

  sheet.getRange(sheet.getLastRow() + 1, 1, reversedRows.length, lastCol).setValues(reversedRows);

  const actor = _resolveActorEmail_();
  logSystemEvent(actor, 'REVERSE_JOURNAL_BATCH', reversalBatchId, 'Source ' + batch + ', Rows ' + reversedRows.length);
  return { count: reversedRows.length, batchId: reversalBatchId };
}

function _buildReversalRow_(row, cols, reversalBatchId, dateValue) {
  const reversed = new Array(row.length).fill('');
  if (cols.uuid) reversed[cols.uuid - 1] = Utilities.getUuid();
  if (cols.batchId) reversed[cols.batchId - 1] = reversalBatchId;
  if (cols.date) reversed[cols.date - 1] = dateValue;
  if (cols.financialYear) reversed[cols.financialYear - 1] = row[cols.financialYear - 1] || '';
  if (cols.accountCode) reversed[cols.accountCode - 1] = row[cols.accountCode - 1] || '';
  if (cols.payee) reversed[cols.payee - 1] = row[cols.payee - 1] || '';
  if (cols.refNo) {
    const refNo = String(row[cols.refNo - 1] || '').trim();
    reversed[cols.refNo - 1] = refNo ? 'REV-' + refNo : 'REV-' + reversalBatchId;
  }
  if (cols.bankRef) {
    const bankRef = String(row[cols.bankRef - 1] || '').trim();
    reversed[cols.bankRef - 1] = bankRef ? 'REV-' + bankRef : '';
  }
  if (cols.particulars) reversed[cols.particulars - 1] = row[cols.particulars - 1] || '';
  if (cols.subCategory) reversed[cols.subCategory - 1] = row[cols.subCategory - 1] || '';
  if (cols.category) reversed[cols.category - 1] = row[cols.category - 1] || '';
  if (cols.description) {
    const desc = String(row[cols.description - 1] || '').trim();
    reversed[cols.description - 1] = desc ? 'Reversal: ' + desc : 'Reversal entry';
  }
  if (cols.debit) reversed[cols.debit - 1] = Number(row[cols.credit - 1] || 0);
  if (cols.credit) reversed[cols.credit - 1] = Number(row[cols.debit - 1] || 0);
  if (cols.accountType) reversed[cols.accountType - 1] = row[cols.accountType - 1] || '';
  if (cols.reportMapping) reversed[cols.reportMapping - 1] = row[cols.reportMapping - 1] || '';
  if (cols.reconStatus) {
    const hasAccount = cols.accountCode ? String(row[cols.accountCode - 1] || '').trim() : '';
    reversed[cols.reconStatus - 1] = hasAccount ? 'Pending' : '';
  }
  if (cols.receiptUrl) reversed[cols.receiptUrl - 1] = row[cols.receiptUrl - 1] || '';
  return reversed;
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
  _requireReconCriteria_(criteria);
  const data = _collectReconciliationData_(criteria);
  const matchResult = _matchReconRows_(data.journalRows, data.bankRows);
  return _buildReconciliationResponse_(data.journalRows, data.bankRows, matchResult, true);
}

function autoReconcileBankStatements(criteria) {
  _requireReconCriteria_(criteria);
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
  _requireReconCriteria_(criteria);
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
  SpreadsheetApp.flush();
  Utilities.sleep(100);

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
    const cashBatchIds = new Set();

    journalData.forEach(row => {
      if (!cols.financialYear || !cols.accountCode || !cols.batchId) return;
      const rowYear = String(row[cols.financialYear - 1] || '').trim();
      if (rowYear !== year) return;
      const accountCode = String(row[cols.accountCode - 1] || '').trim();
      const batchId = String(row[cols.batchId - 1] || '').trim();
      if (accountCode && batchId) cashBatchIds.add(batchId);
    });

    journalData.forEach(row => {
      if (!cols.particulars) return;
      const rowYear = cols.financialYear ? String(row[cols.financialYear - 1] || '').trim() : '';
      if (rowYear !== year) return;
      const batchId = cols.batchId ? String(row[cols.batchId - 1] || '').trim() : '';
      if (batchId && !cashBatchIds.has(batchId)) return;
      const accountCode = cols.accountCode ? String(row[cols.accountCode - 1] || '').trim() : '';
      if (accountCode) return;

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

function getNotesReport(currentYear, comparativeYear, options) {
  const year = String(currentYear || '').trim();
  const compare = String(comparativeYear || '').trim();
  if (!year) throw new Error('Current financial year is required.');
  if (!compare) throw new Error('Comparative financial year is required.');
  const basis = options && options.basis ? String(options.basis).trim().toLowerCase() : 'accrual';
  const useCashBasis = basis === 'cash';

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
  const isNetAssetMapping = (value) => String(value || '').toLowerCase().includes('net asset');

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
      if (isNetAssetMapping(reportMapping)) return;

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
  const cashBatchCurrent = new Set();
  const cashBatchComparative = new Set();

  if (journalLastRow >= 2) {
    const journalHeaders = journal.getRange(1, 1, 1, journalLastCol).getValues()[0].map(_normalizeHeader_);
    const cols = _getJournalColumns_(journalHeaders);
    const journalData = journal.getRange(2, 1, journalLastRow - 1, journalLastCol).getValues();

    journalData.forEach(row => {
      if (!cols.financialYear) return;
      const rowYear = String(row[cols.financialYear - 1] || '').trim();
      if (rowYear !== year && rowYear !== compare) return;
      const debit = cols.debit ? _parseNumber_(row[cols.debit - 1]) : 0;
      const credit = cols.credit ? _parseNumber_(row[cols.credit - 1]) : 0;
      const accountCode = cols.accountCode ? String(row[cols.accountCode - 1] || '').trim() : '';
      const batchId = cols.batchId ? String(row[cols.batchId - 1] || '').trim() : '';

      if (accountCode && batchId) {
        if (rowYear === year) {
          cashBatchCurrent.add(batchId);
          bankNetCurrent[accountCode] = (bankNetCurrent[accountCode] || 0) + (debit - credit);
        } else {
          cashBatchComparative.add(batchId);
          bankNetComparative[accountCode] = (bankNetComparative[accountCode] || 0) + (debit - credit);
        }
      }
    });

    journalData.forEach(row => {
      if (!cols.particulars || !cols.financialYear) return;
      const rowYear = String(row[cols.financialYear - 1] || '').trim();
      if (rowYear !== year && rowYear !== compare) return;
      const accountCode = cols.accountCode ? String(row[cols.accountCode - 1] || '').trim() : '';
      if (useCashBasis) {
        if (!accountCode) return;
        const batchId = cols.batchId ? String(row[cols.batchId - 1] || '').trim() : '';
        if (rowYear === year && !cashBatchCurrent.has(batchId)) return;
        if (rowYear === compare && !cashBatchComparative.has(batchId)) return;
      } else {
        if (accountCode) return;
      }
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

    if (isNetAssetMapping(reportMapping)) return;
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
        const currentAmount = useCashBasis
          ? _resolveCashNotesAmount_(current, particularsName, meta)
          : _resolveNotesAmount_(current, particularsName, accountType);
        const comparativeAmount = useCashBasis
          ? _resolveCashNotesAmount_(comparative, particularsName, meta)
          : _resolveNotesAmount_(comparative, particularsName, accountType);
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

    const categoryLower = String(categoryName || '').toLowerCase();
    const categoryMappings = Object.keys(categoryReportMappings[categoryName] || {}).map(value => String(value || '').toLowerCase());
    if (categoryLower === 'accumulated fund' || categoryLower === 'revaluation reserve' || categoryLower === 'revaluation surplus') {
      return;
    }
    if (categoryMappings.some(value => value.includes('net asset'))) {
      return;
    }

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

  // Sort categories in government-required order for financial statements
  const getCategoryOrder = (item) => {
    const name = String(item.category || '').toLowerCase();
    const accountTypes = (item.accountTypes || []).map(value => String(value || '').toLowerCase());
    const mappings = (item.reportMappings || []).map(value => String(value || '').toLowerCase());
    const isIncome = accountTypes.some(v => v.includes('income')) || mappings.some(v => v.includes('operating income'));
    const isExpense = accountTypes.some(v => v.includes('expense')) || mappings.some(v => v.includes('operating expense'));
    const isCurrentAsset = mappings.some(v => v.includes('current asset') && !v.includes('non-current') && !v.includes('non current'));
    const isNonCurrentAsset = mappings.some(v => v.includes('non-current asset') || v.includes('non current asset'));
    const isCurrentLiability = mappings.some(v => v.includes('current liab') && !v.includes('non-current') && !v.includes('non current'));
    const isNonCurrentLiability = mappings.some(v => v.includes('non-current liab') || v.includes('non current liab'));

    // Revenue categories (Notes 6-10)
    if (name.includes('transfer') && (name.includes('ppf') || name.includes('political parties'))) return 100;
    if (name.includes('membership fee')) return 101;
    if (name.includes('public contribution') || name.includes('donation')) return 102;
    if (name.includes('investment income')) return 103;
    if (name.includes('miscellaneous revenue') || name.includes('other revenue') || name.includes('other income')) return 104;
    if (isIncome || name.includes('income') || name.includes('revenue')) return 105;

    // Expense categories (Notes 11-14)
    if (name.includes('administrative expense') || name.includes('admin expense')) return 200;
    if (name.includes('special interest') || name.includes('interest group')) return 201;
    if (name.includes('advocacy') || name.includes('electoral')) return 202;
    if (name.includes('finance cost') || name.includes('interest expense')) return 203;
    if (isExpense || name.includes('expense')) return 204;

    // Current Assets (Notes 19-22)
    if (name.includes('cash and cash equivalent') || name.includes('cash & cash equivalent')) return 300;
    if (name.includes('receivable') || name.includes('debtors') || name.includes('advance')) return 301;
    if (name.includes('inventor') || name.includes('stock')) return 302;
    if ((name.includes('investment') && isCurrentAsset) || name.includes('current investment')) return 303;
    if (isCurrentAsset) return 304;

    // Non-Current Assets (Notes 23-25)
    if (name.includes('investment') && (isNonCurrentAsset || (!isCurrentAsset && !name.includes('current')))) return 400;
    if (name.includes('property') && (name.includes('plant') || name.includes('equipment')) || name.includes('ppe')) return 401;
    if (name.includes('intangible')) return 402;
    if (name.includes('investment property')) return 403;
    if (isNonCurrentAsset || name.includes('non-current asset') || name.includes('non current asset')) return 404;

    // Current Liabilities (Notes 26-31)
    if (name.includes('trade') && name.includes('payable') || name.includes('other payable') || name.includes('accounts payable')) return 500;
    if (name.includes('refundable deposit') || name.includes('customer deposit')) return 501;
    if (name.includes('current provision') || (name.includes('provision') && isCurrentLiability)) return 502;
    if (name.includes('finance lease') || name.includes('lease obligation')) return 503;
    if (name.includes('deferred income') || name.includes('unearned revenue')) return 504;
    if (name.includes('current portion') && name.includes('borrowing')) return 505;
    if (isCurrentLiability) return 506;

    // Non-Current Liabilities
    if (isNonCurrentLiability || name.includes('non-current liab') || name.includes('non current liab')) return 600;
    if (name.includes('liabil')) return 601;

    return 999;
  };

  results.sort((a, b) => {
    const aOrder = getCategoryOrder(a);
    const bOrder = getCategoryOrder(b);
    if (aOrder !== bOrder) return aOrder - bOrder;
    return String(a.category || '').localeCompare(String(b.category || ''));
  });

  // Assign note numbers to each category (starting from 6)
  const noteNumbers = _buildNoteNumberByCategory(results);
  results.forEach(category => {
    if (category && category.category) {
      category.noteNumber = noteNumbers[category.category] || '';
    }
  });

  return {
    currentYear: year,
    comparativeYear: compare,
    categories: results
  };
}

/**
 * Helper to assign note numbers starting from 6, following government reporting order
 * Categories are already sorted in correct order, just assign sequential numbers starting from 6
 */
function _buildNoteNumberByCategory(categories) {
  const noteNumberByCategory = {};
  if (!categories || !categories.length) return noteNumberByCategory;

  // Assign sequential note numbers starting from 6
  let noteNumber = 6;
  categories.forEach(category => {
    if (!category || !category.category) return;
    noteNumberByCategory[category.category] = noteNumber;
    noteNumber++;
  });

  return noteNumberByCategory;
}

function _buildNoteMapsForReports_(performanceCategories, cashCategories) {
  const baseMap = _buildNoteNumberByCategory(performanceCategories || []);
  const perfNames = new Set((performanceCategories || []).map(item => item && item.category).filter(Boolean));
  const cashNames = new Set((cashCategories || []).map(item => item && item.category).filter(Boolean));
  const shared = new Set();
  perfNames.forEach(name => {
    if (cashNames.has(name)) shared.add(name);
  });

  const performanceNotesByCategory = {};
  (performanceCategories || []).forEach(category => {
    if (!category || !category.category) return;
    const base = baseMap[category.category];
    performanceNotesByCategory[category.category] = shared.has(category.category) ? String(base) + 'a' : String(base || '');
  });

  let nextBase = Object.values(baseMap).reduce((maxValue, value) => {
    const numeric = Number(value);
    return Number.isFinite(numeric) ? Math.max(maxValue, numeric) : maxValue;
  }, 0);
  const cashNotesByCategory = {};
  (cashCategories || []).forEach(category => {
    if (!category || !category.category) return;
    const name = category.category;
    if (shared.has(name)) {
      cashNotesByCategory[name] = String(baseMap[name]) + 'b';
    } else {
      nextBase += 1;
      cashNotesByCategory[name] = String(nextBase);
    }
  });

  return {
    performanceNotesByCategory,
    cashNotesByCategory
  };
}

function getPerformanceReport(currentYear, comparativeYear) {
  const notes = getNotesReport(currentYear, comparativeYear);
  const categories = Array.isArray(notes.categories) ? notes.categories : [];
  const cashNotes = getNotesReport(currentYear, comparativeYear, { basis: 'cash' });
  const cashCategories = Array.isArray(cashNotes.categories) ? cashNotes.categories : [];
  const noteMaps = _buildNoteMapsForReports_(categories, cashCategories);
  const noteNumberByCategory = noteMaps.performanceNotesByCategory || {};

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

function getNotesReportCombined(currentYear, comparativeYear) {
  const performance = getNotesReport(currentYear, comparativeYear);
  const cash = getNotesReport(currentYear, comparativeYear, { basis: 'cash' });
  const performanceCategories = Array.isArray(performance.categories) ? performance.categories : [];
  const cashCategories = Array.isArray(cash.categories) ? cash.categories : [];
  const noteMaps = _buildNoteMapsForReports_(performanceCategories, cashCategories);

  performanceCategories.forEach(category => {
    if (!category || !category.category) return;
    category.noteNumber = noteMaps.performanceNotesByCategory[category.category] || '';
  });
  cashCategories.forEach(category => {
    if (!category || !category.category) return;
    category.noteNumber = noteMaps.cashNotesByCategory[category.category] || '';
  });

  return {
    currentYear: performance.currentYear || String(currentYear || '').trim(),
    comparativeYear: performance.comparativeYear || String(comparativeYear || '').trim(),
    performanceCategories: performanceCategories,
    cashCategories: cashCategories
  };
}

function getCashFlowReport(currentYear, comparativeYear) {
  const notes = getNotesReport(currentYear, comparativeYear, { basis: 'cash' });
  const categories = Array.isArray(notes.categories) ? notes.categories : [];
  const performanceNotes = getNotesReport(currentYear, comparativeYear);
  const performanceCategories = Array.isArray(performanceNotes.categories) ? performanceNotes.categories : [];
  const noteMaps = _buildNoteMapsForReports_(performanceCategories, categories);
  const cashFlowNoteNumberByCategory = noteMaps.cashNotesByCategory || {};
  const performanceNoteNumberByCategory = noteMaps.performanceNotesByCategory || {};

  const ss = _getOrCreateSpreadsheet();
  const journal = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!journal) throw new Error('DB_JOURNAL not found.');
  _ensureJournalContactColumn_(journal);
  _ensureJournalContactColumn_(journal);
  const master = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);
  if (!master) throw new Error('MASTER_DATA not found.');

  const journalLastRow = journal.getLastRow();
  const journalLastCol = journal.getLastColumn();

  const operatingReceipts = {};
  const operatingPayments = {};
  const investing = {};
  const financing = {};
  let receiptsCurrent = 0;
  let receiptsComparative = 0;
  let paymentsCurrent = 0;
  let paymentsComparative = 0;
  let investingCurrent = 0;
  let investingComparative = 0;
  let financingCurrent = 0;
  let financingComparative = 0;
  const subCategoryLookup = {};
  const nonReceivableEntries = [];

  if (journalLastRow >= 2) {
    const journalHeaders = journal.getRange(1, 1, 1, journalLastCol).getValues()[0].map(_normalizeHeader_);
    const cols = _getJournalColumns_(journalHeaders);
    const journalData = journal.getRange(2, 1, journalLastRow - 1, journalLastCol).getValues();
    const cashBatchCurrent = new Set();
    const cashBatchComparative = new Set();
    const bankDirectionByBatch = {};
    const bankOverrideByBatch = {};
    const bankOverrideProcessed = {};

    const isReceivablePayableEntry = function(mapping, accountType, category, particulars) {
      const mappingLower = String(mapping || '').toLowerCase();
      const accountTypeLower = String(accountType || '').toLowerCase();
      const categoryLower = String(category || '').toLowerCase();
      const particularsLower = String(particulars || '').toLowerCase();
      const hasReceivablePayable = mappingLower.includes('receivable') || mappingLower.includes('payable') ||
        accountTypeLower.includes('receivable') || accountTypeLower.includes('payable') ||
        categoryLower.includes('receivable') || categoryLower.includes('payable') ||
        particularsLower.includes('receivable') || particularsLower.includes('payable');
      if (hasReceivablePayable) return true;
      const isAdvanceLike = categoryLower.includes('advance') || particularsLower.includes('advance') ||
        categoryLower.includes('prepaid') || particularsLower.includes('prepaid');
      if (isAdvanceLike && (mappingLower.includes('current asset') || accountTypeLower.includes('asset'))) {
        return true;
      }
      return false;
    };

    const masterLastRow = master.getLastRow();
    const masterLastCol = master.getLastColumn();
    if (masterLastRow >= 2) {
      const masterHeaders = master.getRange(1, 1, 1, masterLastCol).getValues()[0].map(_normalizeHeader_);
      const masterCols = _getMasterColumns_(masterHeaders);
      const masterData = master.getRange(2, 1, masterLastRow - 1, masterLastCol).getValues();

      masterData.forEach(row => {
        const particulars = masterCols.particulars ? String(row[masterCols.particulars - 1] || '').trim() : '';
        const subCategory = masterCols.subCategory ? String(row[masterCols.subCategory - 1] || '').trim() : '';
        const category = masterCols.category ? String(row[masterCols.category - 1] || '').trim() : '';
        const accountType = masterCols.accountType ? String(row[masterCols.accountType - 1] || '').trim() : '';
        const reportMapping = masterCols.reportMapping ? String(row[masterCols.reportMapping - 1] || '').trim() : '';
        if (!particulars) return;
        if (isReceivablePayableEntry(reportMapping, accountType, category, particulars)) return;
        const entry = { particulars, subCategory, category, accountType, reportMapping };
        nonReceivableEntries.push(entry);
        if (subCategory) {
          if (!subCategoryLookup[subCategory]) subCategoryLookup[subCategory] = [];
          subCategoryLookup[subCategory].push(entry);
        }
      });
    }

    const normalizeKeywords = function(value) {
      const cleaned = String(value || '')
        .toLowerCase()
        .replace(/receivable|payable|advance|prepaid/g, ' ')
        .replace(/[^a-z0-9]+/g, ' ')
        .trim();
      if (!cleaned) return [];
      return cleaned.split(/\s+/).filter(Boolean);
    };

    const resolveReceivableFallback = function(particulars, subCategory) {
      const keywords = normalizeKeywords(particulars);
      const candidates = (subCategory && subCategoryLookup[subCategory]) ? subCategoryLookup[subCategory] : nonReceivableEntries;
      if (!candidates.length) return null;
      if (keywords.length) {
        const match = candidates.find(entry => {
          const entryText = String(entry.particulars || '').toLowerCase();
          return keywords.some(keyword => entryText.includes(keyword));
        });
        if (match) return match;
      }
      return candidates[0] || null;
    };

    // Build map of original entries by Advance_ID for tracing receivable/payable settlements
    const originalEntriesByAdvanceId = {};
    journalData.forEach(row => {
      const advanceId = cols.advanceId ? String(row[cols.advanceId - 1] || '').trim() : '';
      if (!advanceId) return;
      const accountCode = cols.accountCode ? String(row[cols.accountCode - 1] || '').trim() : '';
      if (accountCode) return; // Skip bank entries
      const particulars = cols.particulars ? String(row[cols.particulars - 1] || '').trim() : '';
      const category = cols.category ? String(row[cols.category - 1] || '').trim() : '';
      const reportMapping = cols.reportMapping ? String(row[cols.reportMapping - 1] || '').trim() : '';
      const accountType = cols.accountType ? String(row[cols.accountType - 1] || '').trim() : '';
      const batchId = cols.batchId ? String(row[cols.batchId - 1] || '').trim() : '';
      const debit = cols.debit ? _parseNumber_(row[cols.debit - 1]) : 0;
      const credit = cols.credit ? _parseNumber_(row[cols.credit - 1]) : 0;
      const amount = Math.abs(debit - credit);
      if (!amount) return;

      if (!originalEntriesByAdvanceId[advanceId]) {
        originalEntriesByAdvanceId[advanceId] = [];
      }
      originalEntriesByAdvanceId[advanceId].push({
        particulars, category, reportMapping, accountType, debit, credit, amount, batchId
      });
    });

    journalData.forEach(row => {
      if (!cols.financialYear || !cols.accountCode || !cols.batchId) return;
      const rowYear = String(row[cols.financialYear - 1] || '').trim();
      if (rowYear !== notes.currentYear && rowYear !== notes.comparativeYear) return;
      const accountCode = String(row[cols.accountCode - 1] || '').trim();
      const batchId = String(row[cols.batchId - 1] || '').trim();
      if (!accountCode || !batchId) return;
      const debit = cols.debit ? _parseNumber_(row[cols.debit - 1]) : 0;
      const credit = cols.credit ? _parseNumber_(row[cols.credit - 1]) : 0;
      const direction = debit > 0 ? 'receipt' : (credit > 0 ? 'payment' : '');
      if (direction) bankDirectionByBatch[batchId] = direction;

      const particulars = cols.particulars ? String(row[cols.particulars - 1] || '').trim() : '';
      if (particulars) {
        const category = cols.category ? String(row[cols.category - 1] || '').trim() : '';
        const subCategory = cols.subCategory ? String(row[cols.subCategory - 1] || '').trim() : '';
        const reportMapping = cols.reportMapping ? String(row[cols.reportMapping - 1] || '').trim() : '';
        const accountType = cols.accountType ? String(row[cols.accountType - 1] || '').trim() : '';
        const description = cols.description ? String(row[cols.description - 1] || '').trim() : '';
        const amount = Math.abs(debit - credit);
        if (amount > 0) {
          bankOverrideByBatch[batchId] = {
            particulars,
            category,
            subCategory,
            reportMapping,
            accountType,
            description,
            amount,
            year: rowYear
          };
        }
      }

      if (rowYear === notes.currentYear) {
        cashBatchCurrent.add(batchId);
      } else {
        cashBatchComparative.add(batchId);
      }
    });

    // Helper function to add amounts to cash flow sections
    function addToCashFlow(lineCategory, amount, classification, targetYear, noteOverride) {
      const lineNote = noteOverride !== undefined ? noteOverride : '';
      if (classification.section === 'operating') {
        if (classification.direction === 'receipt') {
          operatingReceipts[lineCategory] = operatingReceipts[lineCategory] || { description: lineCategory, note: lineNote, currentAmount: 0, comparativeAmount: 0 };
          operatingReceipts[lineCategory][targetYear + 'Amount'] += amount;
          if (targetYear === 'current') receiptsCurrent += amount;
          else receiptsComparative += amount;
        } else {
          operatingPayments[lineCategory] = operatingPayments[lineCategory] || { description: lineCategory, note: lineNote, currentAmount: 0, comparativeAmount: 0 };
          operatingPayments[lineCategory][targetYear + 'Amount'] += amount;
          if (targetYear === 'current') paymentsCurrent += amount;
          else paymentsComparative += amount;
        }
      } else if (classification.section === 'investing') {
        investing[lineCategory] = investing[lineCategory] || { description: lineCategory, note: lineNote, currentAmount: 0, comparativeAmount: 0 };
        const signedAmount = classification.direction === 'receipt' ? amount : -Math.abs(amount);
        investing[lineCategory][targetYear + 'Amount'] += signedAmount;
        if (targetYear === 'current') investingCurrent += signedAmount;
        else investingComparative += signedAmount;
      } else if (classification.section === 'financing') {
        financing[lineCategory] = financing[lineCategory] || { description: lineCategory, note: lineNote, currentAmount: 0, comparativeAmount: 0 };
        const signedAmount = classification.direction === 'receipt' ? amount : -Math.abs(amount);
        financing[lineCategory][targetYear + 'Amount'] += signedAmount;
        if (targetYear === 'current') financingCurrent += signedAmount;
        else financingComparative += signedAmount;
      }
    }

    journalData.forEach(row => {
      if (!cols.financialYear || !cols.particulars || !cols.batchId) return;
      const rowYear = String(row[cols.financialYear - 1] || '').trim();
      if (rowYear !== notes.currentYear && rowYear !== notes.comparativeYear) return;
      const batchId = String(row[cols.batchId - 1] || '').trim();
      if (rowYear === notes.currentYear && !cashBatchCurrent.has(batchId)) return;
      if (rowYear === notes.comparativeYear && !cashBatchComparative.has(batchId)) return;

      const accountCode = cols.accountCode ? String(row[cols.accountCode - 1] || '').trim() : '';
      if (accountCode) return;
      const bankDirection = bankDirectionByBatch[batchId];
      if (!bankDirection) return;

      const bankOverride = bankOverrideByBatch[batchId];
      if (bankOverride) {
        if (!bankOverrideProcessed[batchId]) {
          bankOverrideProcessed[batchId] = true;
          const overrideYear = bankOverride.year;
          const targetYear = overrideYear === notes.currentYear ? 'current' : 'comparative';
          const lineCategory = bankOverride.category || 'Uncategorized';
          const classDebit = bankDirection === 'payment' ? bankOverride.amount : 0;
          const classCredit = bankDirection === 'receipt' ? bankOverride.amount : 0;
          const classification = _classifyCashFlowLine_(bankOverride.reportMapping, bankOverride.accountType, classDebit, classCredit);
          if (classification) {
            addToCashFlow(lineCategory, bankOverride.amount, classification, targetYear);
          }
        }
        return;
      }

      const particulars = String(row[cols.particulars - 1] || '').trim();
      if (!particulars) return;
      const category = cols.category ? String(row[cols.category - 1] || '').trim() : 'Uncategorized';
      const subCategory = cols.subCategory ? String(row[cols.subCategory - 1] || '').trim() : '';
      const reportMapping = cols.reportMapping ? String(row[cols.reportMapping - 1] || '').trim() : '';
      const accountType = cols.accountType ? String(row[cols.accountType - 1] || '').trim() : '';
      const debit = cols.debit ? _parseNumber_(row[cols.debit - 1]) : 0;
      const credit = cols.credit ? _parseNumber_(row[cols.credit - 1]) : 0;
      const amount = credit > 0 ? credit : debit;
      if (!amount) return;

      const targetYear = rowYear === notes.currentYear ? 'current' : 'comparative';
      const advanceId = cols.advanceId ? String(row[cols.advanceId - 1] || '').trim() : '';

      // Check if this is a receivable/payable settlement that needs tracing
      const mappingLower = String(reportMapping || '').toLowerCase();
      const isReceivablePayable = isReceivablePayableEntry(reportMapping, accountType, category, particulars);

      // If this is a receivable/payable entry with an Advance_ID, trace to original entries
      if (isReceivablePayable && advanceId && originalEntriesByAdvanceId[advanceId]) {
        const originalEntries = originalEntriesByAdvanceId[advanceId];

        // Filter to get only income/expense entries (not the receivable/payable counter-entries)
        const underlyingEntries = originalEntries.filter(e => {
          return !isReceivablePayableEntry(e.reportMapping, e.accountType, e.category, e.particulars);
        });

        if (underlyingEntries.length > 0) {
          // Calculate total of underlying entries for proportional allocation
          const totalUnderlying = underlyingEntries.reduce((sum, e) => sum + e.amount, 0);

          // Allocate the cash amount proportionally across underlying categories
          underlyingEntries.forEach(entry => {
            const proportion = totalUnderlying > 0 ? entry.amount / totalUnderlying : 1 / underlyingEntries.length;
            const allocatedAmount = amount * proportion;

            const lineCategory = entry.category || 'Uncategorized';

            const classDebit = bankDirection === 'payment' ? allocatedAmount : 0;
            const classCredit = bankDirection === 'receipt' ? allocatedAmount : 0;
            const classification = _classifyCashFlowLine_(entry.reportMapping, entry.accountType, classDebit, classCredit);

            if (classification) {
              addToCashFlow(lineCategory, allocatedAmount, classification, targetYear);
            }
          });
          return; // Skip normal processing since we handled via tracing
        }
      }

      if (isReceivablePayable) {
        const fallbackEntry = resolveReceivableFallback(particulars, subCategory);
        if (fallbackEntry) {
          const lineCategory = fallbackEntry.category || 'Uncategorized';
          const classDebit = bankDirection === 'payment' ? amount : 0;
          const classCredit = bankDirection === 'receipt' ? amount : 0;
          const classification = _classifyCashFlowLine_(fallbackEntry.reportMapping, fallbackEntry.accountType, classDebit, classCredit);
          if (classification) {
            addToCashFlow(lineCategory, amount, classification, targetYear);
          }
        }
        return;
      }

      // Normal processing for non-traced entries
      const isNetAssets = mappingLower.includes('net asset');
      let classification = null;
      let lineCategory = category;
      if (isNetAssets) {
        classification = { section: 'financing', direction: bankDirection };
        lineCategory = 'Prior year adjustment';
      } else {
        const classDebit = bankDirection === 'payment' ? amount : 0;
        const classCredit = bankDirection === 'receipt' ? amount : 0;
        classification = _classifyCashFlowLine_(reportMapping, accountType, classDebit, classCredit);
      }
      if (!classification) return;

      const noteOverride = isNetAssets ? '' : undefined;
      addToCashFlow(lineCategory, amount, classification, targetYear, noteOverride);
    });
  }

  const operatingReceiptsRows = Object.values(operatingReceipts);
  const operatingPaymentsRows = Object.values(operatingPayments);
  const investingRows = Object.values(investing);
  const financingRows = Object.values(financing);
  const cashFlowOrder = [operatingReceiptsRows, operatingPaymentsRows, investingRows, financingRows];
  cashFlowOrder.forEach(rows => {
    rows.forEach(row => {
      if (!row || !row.description) return;
      if (row.description === 'Prior year adjustment') {
        row.note = '';
        return;
      }
      row.note = cashFlowNoteNumberByCategory[row.description] || '';
    });
  });

  const cashNote = categories.find(category => String(category.category || '').toLowerCase().includes('cash and cash equivalent')) || null;
  const netOperatingCurrent = receiptsCurrent - paymentsCurrent;
  const netOperatingComparative = receiptsComparative - paymentsComparative;
  const netIncreaseCurrent = netOperatingCurrent + investingCurrent + financingCurrent;
  const netIncreaseComparative = netOperatingComparative + investingComparative + financingComparative;
  const cashOpeningCurrent = cashNote ? Number(cashNote.totalComparative || 0) : 0;
  const cashClosingCurrent = cashOpeningCurrent + netIncreaseCurrent;
  const cashOpeningComparative = 0;
  const cashClosingComparative = cashOpeningComparative + netIncreaseComparative;

  return {
    currentYear: notes.currentYear || String(currentYear || '').trim(),
    comparativeYear: notes.comparativeYear || String(comparativeYear || '').trim(),
    operating: {
      receipts: operatingReceiptsRows,
      payments: operatingPaymentsRows,
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
    cashNote: cashNote ? (performanceNoteNumberByCategory[cashNote.category] || '') : '',
    cashOpeningCurrent: cashOpeningCurrent,
    cashOpeningComparative: cashOpeningComparative,
    cashClosingCurrent: cashClosingCurrent,
    cashClosingComparative: cashClosingComparative
  };
}

function getPositionReport(currentYear, comparativeYear) {
  const notes = getNotesReport(currentYear, comparativeYear);
  const categories = Array.isArray(notes.categories) ? notes.categories : [];
  const noteNumberByCategory = _buildNoteNumberByCategory(categories);

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

  const ss = _getOrCreateSpreadsheet();
  const journal = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!journal) throw new Error('DB_JOURNAL not found.');

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

  const performance = getPerformanceReport(currentYear, comparativeYear);
  const surplusCurrent = performance && performance.surplus ? Number(performance.surplus.current || 0) : 0;
  const surplusComparative = performance && performance.surplus ? Number(performance.surplus.comparative || 0) : 0;
  const journalLastRow = journal.getLastRow();
  const journalLastCol = journal.getLastColumn();
  let movementsCurrent = { accumulated: 0, revaluation: 0 };
  let movementsComparative = { accumulated: 0, revaluation: 0 };
  if (journalLastRow >= 2) {
    const journalHeaders = journal.getRange(1, 1, 1, journalLastCol).getValues()[0].map(_normalizeHeader_);
    const cols = _getJournalColumns_(journalHeaders);
    const data = journal.getRange(2, 1, journalLastRow - 1, journalLastCol).getValues();

    const sumMovements = function(targetYear) {
      const totals = { accumulated: 0, revaluation: 0 };
      data.forEach(function(row) {
        if (!cols.financialYear || !cols.particulars) return;
        const rowYear = String(row[cols.financialYear - 1] || '').trim();
        if (rowYear !== targetYear) return;
        const accountCode = cols.accountCode ? String(row[cols.accountCode - 1] || '').trim() : '';
        if (accountCode) return;
        const particulars = String(row[cols.particulars - 1] || '').trim().toLowerCase();
        if (!particulars) return;
        const debit = cols.debit ? _parseNumber_(row[cols.debit - 1]) : 0;
        const credit = cols.credit ? _parseNumber_(row[cols.credit - 1]) : 0;
        const movement = credit - debit;
        if (particulars === 'accumulated fund') {
          totals.accumulated += movement;
        }
        if (particulars === 'revaluation reserve' || particulars === 'revaluation surplus') {
          totals.revaluation += movement;
        }
      });
      return totals;
    };

    movementsCurrent = sumMovements(notes.currentYear || String(currentYear || '').trim());
    movementsComparative = sumMovements(notes.comparativeYear || String(comparativeYear || '').trim());
  }

  const previousOpeningAccumulated = 0;
  const previousOpeningRevaluation = 0;
  const previousOtherAccumulated = movementsComparative.accumulated;
  const previousOtherRevaluation = movementsComparative.revaluation;
  const previousClosingAccumulated = previousOpeningAccumulated + previousOtherAccumulated + surplusComparative;
  const previousClosingRevaluation = previousOpeningRevaluation + previousOtherRevaluation;
  const currentOpeningAccumulated = previousClosingAccumulated;
  const currentOpeningRevaluation = previousClosingRevaluation;
  const currentOtherAccumulated = movementsCurrent.accumulated;
  const currentOtherRevaluation = movementsCurrent.revaluation;
  const currentClosingAccumulated = currentOpeningAccumulated + currentOtherAccumulated + surplusCurrent;
  const currentClosingRevaluation = currentOpeningRevaluation + currentOtherRevaluation;
  const changesInNetAssets = {
    titleYear: notes.currentYear || String(currentYear || '').trim(),
    previousYear: notes.comparativeYear || String(comparativeYear || '').trim(),
    previous: {
      openingAccumulated: previousOpeningAccumulated,
      openingRevaluation: previousOpeningRevaluation,
      revaluationGain: 0,
      transfer: 0,
      otherChangesAccumulated: previousOtherAccumulated,
      otherChangesRevaluation: previousOtherRevaluation,
      surplus: surplusComparative,
      closingAccumulated: previousClosingAccumulated,
      closingRevaluation: previousClosingRevaluation
    },
    current: {
      openingAccumulated: currentOpeningAccumulated,
      openingRevaluation: currentOpeningRevaluation,
      revaluationGain: 0,
      transfer: 0,
      otherChangesAccumulated: currentOtherAccumulated,
      otherChangesRevaluation: currentOtherRevaluation,
      surplus: surplusCurrent,
      closingAccumulated: currentClosingAccumulated,
      closingRevaluation: currentClosingRevaluation
    }
  };

  const accumulatedCurrent = changesInNetAssets.current.closingAccumulated;
  const accumulatedComparative = changesInNetAssets.previous.closingAccumulated;
  const revaluationCurrent = changesInNetAssets.current.closingRevaluation;
  const revaluationComparative = changesInNetAssets.previous.closingRevaluation;
  const revaluationReserveRow = {
    description: 'Revaluation Reserve',
    note: '',
    currentAmount: revaluationCurrent,
    comparativeAmount: revaluationComparative
  };
  const accumulatedRow = {
    description: 'Accumulated Fund',
    note: '',
    currentAmount: accumulatedCurrent,
    comparativeAmount: accumulatedComparative
  };
  equityRows.length = 0;
  equityRows.push(revaluationReserveRow, accumulatedRow);
  const equityTotalsCurrent = revaluationReserveRow.currentAmount + accumulatedCurrent;
  const equityTotalsComparative = revaluationReserveRow.comparativeAmount + accumulatedComparative;

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
      rows: [accumulatedRow],
      total: equityTotalsCurrent,
      totalComparative: equityTotalsComparative
    },
    netAssets: {
      current: netAssetsCurrent,
      comparative: netAssetsComparative
    },
    changesInNetAssets: changesInNetAssets
  };
}

function getReceivablePayableSummary(type, criteria) {
  const kind = String(type || '').toLowerCase();
  const isReceivable = kind.includes('receivable');
  const targetContactType = isReceivable ? 'Customer' : 'Supplier';
  const financialYear = String(criteria && criteria.financialYear || '').trim();
  const payeeFilter = String(criteria && criteria.payee || '').trim().toLowerCase();
  const contactFilter = String(criteria && criteria.contactId || '').trim();
  const startDate = _parseDate_(criteria && criteria.startDate);
  const endDate = _parseDate_(criteria && criteria.endDate);
  const hasDateFilter = Boolean(startDate || endDate);

  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) throw new Error('DB_JOURNAL not found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { rows: [] };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getJournalColumns_(headers);
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const summaries = {};

  // Build contact type lookup map
  const contactTypes = {};
  const contactSheet = ss.getSheetByName(CONFIG.SHEETS.CONTACTS);
  if (contactSheet) {
    const contactLastRow = contactSheet.getLastRow();
    if (contactLastRow >= 2) {
      const contactData = contactSheet.getRange(2, 1, contactLastRow - 1, 3).getValues();
      contactData.forEach(function(row) {
        const cid = String(row[0] || '').trim();
        const ctype = String(row[1] || '').trim();
        if (cid) contactTypes[cid] = ctype;
      });
    }
  }

  const isAdvanceLike = function(category, reportMapping, accountType, particulars) {
    const categoryLower = String(category || '').toLowerCase();
    const mappingLower = String(reportMapping || '').toLowerCase();
    const accountTypeLower = String(accountType || '').toLowerCase();
    const particularsLower = String(particulars || '').toLowerCase();
    if (categoryLower.includes('advance') || particularsLower.includes('advance')) return true;
    if (categoryLower.includes('deposit') || particularsLower.includes('deposit')) return true;
    if (categoryLower.includes('unearned') || particularsLower.includes('unearned')) return true;
    if (categoryLower.includes('deferred') || particularsLower.includes('deferred')) return true;
    if (mappingLower.includes('liabil') || accountTypeLower.includes('liabil')) return true;
    return false;
  };

  data.forEach(function(row) {
    if (!cols.payee || !cols.particulars) return;
    const payee = String(row[cols.payee - 1] || '').trim();
    if (!payee) return;
    const contactId = cols.contactId ? String(row[cols.contactId - 1] || '').trim() : '';
    if (contactFilter && contactId !== contactFilter) return;
    if (payeeFilter && !payee.toLowerCase().includes(payeeFilter)) return;
    const accountCode = cols.accountCode ? String(row[cols.accountCode - 1] || '').trim() : '';
    const advanceId = cols.advanceId ? String(row[cols.advanceId - 1] || '').trim() : '';
    const hasAdvanceBankLine = Boolean(accountCode && advanceId && contactId);
    if (accountCode && !hasAdvanceBankLine) return;

    // Filter by Contact_Type instead of Report_Mapping
    const contactType = contactId ? (contactTypes[contactId] || '') : '';
    if (contactId && contactType !== targetContactType) return;

    const rowYear = cols.financialYear ? String(row[cols.financialYear - 1] || '').trim() : '';
    const dateCell = cols.date ? row[cols.date - 1] : '';
    const rowDate = dateCell instanceof Date ? dateCell : _parseDate_(dateCell);

    const category = cols.category ? String(row[cols.category - 1] || '').trim() : '';
    const reportMapping = cols.reportMapping ? String(row[cols.reportMapping - 1] || '').trim() : '';
    const accountType = cols.accountType ? String(row[cols.accountType - 1] || '').trim() : '';
    const particulars = cols.particulars ? String(row[cols.particulars - 1] || '').trim() : '';
    const mappingLower = reportMapping.toLowerCase();
    const categoryLower = category.toLowerCase();

    // Check transaction nature - only include receivable/payable/advance entries
    const matchesReceivable = mappingLower.includes('receivable') || categoryLower.includes('receivable');
    const matchesPayable = mappingLower.includes('payable') || categoryLower.includes('payable');
    const matchesAdvance = isAdvanceLike(category, reportMapping, accountType, particulars) || hasAdvanceBankLine;

    // Filter: only include relevant transaction types for the statement
    // Customer statement: receivables and advances only
    // Supplier statement: payables and advances only
    if (isReceivable) {
      if (!matchesReceivable && !matchesAdvance) return;
    } else {
      if (!matchesPayable && !matchesAdvance) return;
    }

    const debit = cols.debit ? _parseNumber_(row[cols.debit - 1]) : 0;
    const credit = cols.credit ? _parseNumber_(row[cols.credit - 1]) : 0;

    // Calculate increase/decrease based on transaction type
    let increase = 0;
    let decrease = 0;
    if (isReceivable) {
      // Customer statement: debit increases balance (invoice), credit decreases (payment/advance)
      if (matchesAdvance) {
        // Advance from customer: receivable entry uses credit; bank entry uses debit.
        decrease = hasAdvanceBankLine ? debit : credit;
      } else {
        increase = debit;
        decrease = credit;
      }
    } else {
      // Supplier statement: credit increases balance (invoice), debit decreases (payment)
      if (matchesAdvance) {
        // Advance to supplier: payable entry uses debit; bank entry uses credit.
        decrease = hasAdvanceBankLine ? credit : debit;
      } else {
        increase = credit;
        decrease = debit;
      }
    }

    const isYearMatch = !financialYear || (rowYear && rowYear === financialYear);

    const key = contactId || payee;
    if (!summaries[key]) {
      summaries[key] = { payee: payee, contactId: contactId, opening: 0, additions: 0, payments: 0, closing: 0 };
    }

    if (hasDateFilter) {
      if (financialYear && !isYearMatch) return;
      if (!rowDate) return;
      if (startDate && rowDate < startDate) {
        summaries[key].opening += increase - decrease;
        return;
      }
      if (endDate && rowDate > endDate) return;
      summaries[key].additions += increase;
      summaries[key].payments += decrease;
      return;
    }

    if (financialYear && !isYearMatch) {
      summaries[key].opening += increase - decrease;
      return;
    }

    summaries[key].additions += increase;
    summaries[key].payments += decrease;
  });

  const rows = Object.values(summaries).map(function(item) {
    item.closing = item.opening + item.additions - item.payments;
    return item;
  }).filter(item => Math.abs(item.closing) > 0.01)
    .sort((a, b) => String(a.payee || '').localeCompare(String(b.payee || '')));

  return { rows: rows };
}

function getNettingSummary(criteria) {
  const receivable = getReceivablePayableSummary('receivable', criteria).rows || [];
  const payable = getReceivablePayableSummary('payable', criteria).rows || [];
  const summary = {};

  const ensure = function(row) {
    const key = row.contactId || row.payee;
    if (!key) return null;
    if (!summary[key]) {
      summary[key] = { contactId: row.contactId || '', payee: row.payee || '', receivable: 0, payable: 0, net: 0 };
    }
    if (!summary[key].payee && row.payee) summary[key].payee = row.payee;
    if (!summary[key].contactId && row.contactId) summary[key].contactId = row.contactId;
    return summary[key];
  };

  receivable.forEach(function(row) {
    const item = ensure(row);
    if (!item) return;
    item.receivable += Number(row.closing || 0);
  });

  payable.forEach(function(row) {
    const item = ensure(row);
    if (!item) return;
    item.payable += Number(row.closing || 0);
  });

  const rows = Object.values(summary).map(function(item) {
    item.net = (Number(item.receivable || 0) - Number(item.payable || 0));
    return item;
  }).filter(item => Math.abs(item.net) > 0.01 || Math.abs(item.receivable) > 0.01 || Math.abs(item.payable) > 0.01)
    .sort((a, b) => String(a.payee || '').localeCompare(String(b.payee || '')));

  return { rows: rows };
}

function getReceivablePayableStatement(type, payeeName, criteria) {
  const kind = String(type || '').toLowerCase();
  const isReceivable = kind.includes('receivable');
  const targetContactType = isReceivable ? 'Customer' : 'Supplier';
  const payee = String(payeeName || '').trim();
  const contactFilter = String(criteria && criteria.contactId || '').trim();
  if (!payee && !contactFilter) throw new Error('Payee is required.');
  const financialYear = String(criteria && criteria.financialYear || '').trim();
  const startDate = _parseDate_(criteria && criteria.startDate);
  const endDate = _parseDate_(criteria && criteria.endDate);
  const hasDateFilter = Boolean(startDate || endDate);

  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) throw new Error('DB_JOURNAL not found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { payee: payee, opening: 0, rows: [] };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getJournalColumns_(headers);
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Build contact type lookup map
  const contactTypes = {};
  const contactSheet = ss.getSheetByName(CONFIG.SHEETS.CONTACTS);
  if (contactSheet) {
    const contactLastRow = contactSheet.getLastRow();
    if (contactLastRow >= 2) {
      const contactData = contactSheet.getRange(2, 1, contactLastRow - 1, 3).getValues();
      contactData.forEach(function(row) {
        const cid = String(row[0] || '').trim();
        const ctype = String(row[1] || '').trim();
        if (cid) contactTypes[cid] = ctype;
      });
    }
  }

  const isAdvanceLike = function(category, reportMapping, accountType, particulars) {
    const categoryLower = String(category || '').toLowerCase();
    const mappingLower = String(reportMapping || '').toLowerCase();
    const accountTypeLower = String(accountType || '').toLowerCase();
    const particularsLower = String(particulars || '').toLowerCase();
    if (categoryLower.includes('advance') || particularsLower.includes('advance')) return true;
    if (categoryLower.includes('deposit') || particularsLower.includes('deposit')) return true;
    if (categoryLower.includes('unearned') || particularsLower.includes('unearned')) return true;
    if (categoryLower.includes('deferred') || particularsLower.includes('deferred')) return true;
    if (mappingLower.includes('liabil') || accountTypeLower.includes('liabil')) return true;
    return false;
  };

  let opening = 0;
  const rows = [];

  data.forEach(function(row, index) {
    if (!cols.payee || !cols.particulars) return;
    const rowPayee = String(row[cols.payee - 1] || '').trim();
    const contactId = cols.contactId ? String(row[cols.contactId - 1] || '').trim() : '';
    if (contactFilter) {
      if (contactId !== contactFilter) return;
    } else if (rowPayee !== payee) {
      return;
    }
    const accountCode = cols.accountCode ? String(row[cols.accountCode - 1] || '').trim() : '';
    const advanceId = cols.advanceId ? String(row[cols.advanceId - 1] || '').trim() : '';
    const hasAdvanceBankLine = Boolean(accountCode && advanceId && contactId);
    if (accountCode && !hasAdvanceBankLine) return;

    // Filter by Contact_Type instead of Report_Mapping
    const contactType = contactId ? (contactTypes[contactId] || '') : '';
    if (contactId && contactType !== targetContactType) return;

    const rowYear = cols.financialYear ? String(row[cols.financialYear - 1] || '').trim() : '';
    const category = cols.category ? String(row[cols.category - 1] || '').trim() : '';
    const reportMapping = cols.reportMapping ? String(row[cols.reportMapping - 1] || '').trim() : '';
    const accountType = cols.accountType ? String(row[cols.accountType - 1] || '').trim() : '';
    const particulars = cols.particulars ? String(row[cols.particulars - 1] || '').trim() : '';
    const mappingLower = reportMapping.toLowerCase();
    const categoryLower = category.toLowerCase();

    // Check transaction nature - only include receivable/payable/advance entries
    const matchesReceivable = mappingLower.includes('receivable') || categoryLower.includes('receivable');
    const matchesPayable = mappingLower.includes('payable') || categoryLower.includes('payable');
    const matchesAdvance = isAdvanceLike(category, reportMapping, accountType, particulars) || hasAdvanceBankLine;

    // Filter: only include relevant transaction types for the statement
    // Customer statement: receivables and advances only
    // Supplier statement: payables and advances only
    if (isReceivable) {
      if (!matchesReceivable && !matchesAdvance) return;
    } else {
      if (!matchesPayable && !matchesAdvance) return;
    }

    const debit = cols.debit ? _parseNumber_(row[cols.debit - 1]) : 0;
    const credit = cols.credit ? _parseNumber_(row[cols.credit - 1]) : 0;

    // Calculate increase/decrease based on transaction type
    let increase = 0;
    let decrease = 0;
    if (isReceivable) {
      // Customer statement: debit increases balance (invoice), credit decreases (payment/advance)
      if (matchesAdvance) {
        decrease = hasAdvanceBankLine ? debit : credit;
      } else {
        increase = debit;
        decrease = credit;
      }
    } else {
      // Supplier statement: credit increases balance (invoice), debit decreases (payment)
      if (matchesAdvance) {
        decrease = hasAdvanceBankLine ? credit : debit;
      } else {
        increase = credit;
        decrease = debit;
      }
    }

    const dateCell = cols.date ? row[cols.date - 1] : '';
    const rowDate = dateCell instanceof Date ? dateCell : _parseDate_(dateCell);
    const description = cols.description ? String(row[cols.description - 1] || '').trim() : '';

    if (hasDateFilter) {
      if (financialYear && rowYear !== financialYear) return;
      if (!rowDate) return;
      if (startDate && rowDate < startDate) {
        opening += increase - decrease;
        return;
      }
      if (endDate && rowDate > endDate) return;
    } else if (financialYear && rowYear !== financialYear) {
      opening += increase - decrease;
      return;
    }

    const displayDebit = isReceivable && hasAdvanceBankLine && matchesAdvance ? 0 : debit;
    const displayCredit = isReceivable && hasAdvanceBankLine && matchesAdvance ? debit : credit;

    rows.push({
      index: index,
      dateValue: rowDate,
      date: rowDate ? _formatDate_(rowDate) : '',
      particulars: particulars,
      narration: description,
      debit: displayDebit,
      credit: displayCredit,
      increase: increase,
      decrease: decrease,
      advanceId: advanceId
    });
  });

  rows.sort((a, b) => {
    const aTime = a.dateValue instanceof Date ? a.dateValue.getTime() : 0;
    const bTime = b.dateValue instanceof Date ? b.dateValue.getTime() : 0;
    if (aTime !== bTime) return aTime - bTime;
    return a.index - b.index;
  });

  let balance = opening;
  const finalRows = rows.map(row => {
    balance += row.increase - row.decrease;
    return {
      date: row.date,
      particulars: row.particulars,
      narration: row.narration,
      debit: row.debit,
      credit: row.credit,
      balance: balance,
      advanceId: row.advanceId
    };
  });

  return {
    payee: payee || '',
    contactId: contactFilter || '',
    financialYear: financialYear,
    opening: opening,
    rows: finalRows
  };
}

function getAssetPurchasesSummary(criteria) {
  const financialYear = String(criteria && criteria.financialYear || '').trim();
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) throw new Error('DB_JOURNAL not found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { rows: [] };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getJournalColumns_(headers);
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const openingsBySection = {};
  const additionsBySection = {};
  let totalOpening = 0;
  let totalAdditions = 0;

  data.forEach(function(row) {
    if (!cols.financialYear || !cols.reportMapping || !cols.debit || !cols.credit) return;
    const rowYear = String(row[cols.financialYear - 1] || '').trim();
    if (!rowYear) return;
    const accountCode = cols.accountCode ? String(row[cols.accountCode - 1] || '').trim() : '';
    if (accountCode) return;
    const mapping = String(row[cols.reportMapping - 1] || '');
    const category = cols.category ? String(row[cols.category - 1] || '') : '';
    const section = _classifyAssetPurchaseSection_(mapping, category);
    if (!section) return;
    const debit = _parseNumber_(row[cols.debit - 1]);
    const credit = _parseNumber_(row[cols.credit - 1]);
    const net = debit - credit;
    if (!financialYear) {
      additionsBySection[section] = (additionsBySection[section] || 0) + net;
      totalAdditions += net;
      return;
    }
    const comparison = _compareFinancialYears_(rowYear, financialYear);
    if (comparison < 0) {
      openingsBySection[section] = (openingsBySection[section] || 0) + net;
      totalOpening += net;
      return;
    }
    if (comparison === 0) {
      additionsBySection[section] = (additionsBySection[section] || 0) + net;
      totalAdditions += net;
    }
  });

  const sectionOrder = [
    'Property, Plant & Equipment',
    'Investment Property',
    'Intangible Assets',
    'Biological Assets',
    'Other Non-Current Assets'
  ];
  const sections = sectionOrder.map(function(name) {
    const opening = openingsBySection[name] || 0;
    const additions = additionsBySection[name] || 0;
    return {
      section: name,
      opening: opening,
      additions: additions,
      closing: opening + additions
    };
  });

  return {
    financialYear: financialYear,
    sections: sections,
    totalOpening: totalOpening,
    totalAdditions: totalAdditions,
    totalClosing: totalOpening + totalAdditions
  };
}

function _classifyAssetPurchaseSection_(reportMapping, category) {
  const mapping = String(reportMapping || '').toLowerCase();
  const categoryText = String(category || '').toLowerCase();
  const text = mapping + ' ' + categoryText;

  if (text.includes('investment property')) return 'Investment Property';
  if (text.includes('intangible')) return 'Intangible Assets';
  if (text.includes('biological')) return 'Biological Assets';
  if (text.includes('property, plant') || text.includes('property plant') || text.includes('ppe') || text.includes('property and equipment')) {
    return 'Property, Plant & Equipment';
  }
  if (text.includes('non-current asset') || text.includes('non current asset')) return 'Other Non-Current Assets';
  return '';
}

function getDashboardSummary() {
  const dropdowns = getDropdownData();
  const years = dropdowns && dropdowns.financialYears ? dropdowns.financialYears : [];
  const financialYear = _resolveCurrentFinancialYear_(years);

  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!sheet) throw new Error('DB_JOURNAL not found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  let totalReceipts = 0;
  let totalPayments = 0;
  if (lastRow > 1) {
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
    const cols = _getJournalColumns_(headers);
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    data.forEach(function(row) {
      if (!cols.financialYear || !cols.accountCode) return;
      const rowYear = String(row[cols.financialYear - 1] || '').trim();
      if (financialYear && rowYear !== financialYear) return;
      const accountCode = String(row[cols.accountCode - 1] || '').trim();
      if (!accountCode) return;
      const debit = cols.debit ? _parseNumber_(row[cols.debit - 1]) : 0;
      const credit = cols.credit ? _parseNumber_(row[cols.credit - 1]) : 0;
      if (debit > 0) totalReceipts += debit;
      if (credit > 0) totalPayments += credit;
    });
  }

  const receivableSummary = getReceivablePayableSummary('receivable', { financialYear: financialYear });
  const payableSummary = getReceivablePayableSummary('payable', { financialYear: financialYear });
  const totalReceivables = (receivableSummary.rows || []).reduce((sum, row) => sum + Number(row.closing || 0), 0);
  const totalPayables = (payableSummary.rows || []).reduce((sum, row) => sum + Number(row.closing || 0), 0);

  return {
    financialYear: financialYear,
    totalReceipts: totalReceipts,
    totalPayments: totalPayments,
    totalReceivables: totalReceivables,
    totalPayables: totalPayables
  };
}

function getBankAccountSummaries() {
  const dropdowns = getDropdownData();
  const years = dropdowns && dropdowns.financialYears ? dropdowns.financialYears : [];
  const financialYear = _resolveCurrentFinancialYear_(years);
  const bounds = _resolveFinancialYearBounds_();

  const ss = _getOrCreateSpreadsheet();
  const master = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);
  if (!master) throw new Error('MASTER_DATA not found.');

  const masterLastRow = master.getLastRow();
  const masterLastCol = master.getLastColumn();
  let accountCodes = [];
  if (masterLastRow >= 2) {
    const headers = master.getRange(1, 1, 1, masterLastCol).getValues()[0].map(_normalizeHeader_);
    const cols = _getMasterColumns_(headers);
    const data = master.getRange(2, 1, masterLastRow - 1, masterLastCol).getValues();
    data.forEach(function(row) {
      const accountCode = cols.accountCodes ? String(row[cols.accountCodes - 1]).trim() : '';
      if (accountCode) accountCodes.push(accountCode);
    });
  }

  accountCodes = _uniqueSorted_(accountCodes);
  if (!accountCodes.length) {
    return { financialYear: financialYear, accounts: [] };
  }

  const summaries = {};
  accountCodes.forEach(function(code) {
    summaries[code] = { accountCode: code, opening: 0, receipts: 0, payments: 0, closing: 0 };
  });

  const journal = ss.getSheetByName(CONFIG.SHEETS.DB_JOURNAL);
  if (!journal) throw new Error('DB_JOURNAL not found.');

  const lastRow = journal.getLastRow();
  const lastCol = journal.getLastColumn();
  if (lastRow > 1) {
    const headers = journal.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
    const cols = _getJournalColumns_(headers);
    const data = journal.getRange(2, 1, lastRow - 1, lastCol).getValues();
    data.forEach(function(row) {
      if (!cols.accountCode) return;
      const accountCode = String(row[cols.accountCode - 1] || '').trim();
      if (!accountCode || !summaries[accountCode]) return;
      const debit = cols.debit ? _parseNumber_(row[cols.debit - 1]) : 0;
      const credit = cols.credit ? _parseNumber_(row[cols.credit - 1]) : 0;
      const rowDate = cols.date ? _parseDate_(row[cols.date - 1]) : null;

      if (rowDate) {
        if (rowDate < bounds.startDate) {
          summaries[accountCode].opening += debit - credit;
          return;
        }
        if (rowDate <= bounds.endDate) {
          summaries[accountCode].receipts += debit;
          summaries[accountCode].payments += credit;
        }
        return;
      }

      if (cols.financialYear) {
        const rowYear = String(row[cols.financialYear - 1] || '').trim();
        if (financialYear && rowYear === financialYear) {
          summaries[accountCode].receipts += debit;
          summaries[accountCode].payments += credit;
        }
      }
    });
  }

  const accounts = accountCodes.map(function(code) {
    const summary = summaries[code];
    summary.closing = summary.opening + summary.receipts - summary.payments;
    return summary;
  });

  return { financialYear: financialYear, accounts: accounts };
}

function _resolveCurrentFinancialYear_(years) {
  const now = new Date();
  const currentYear = now.getFullYear();
  const month = now.getMonth() + 1;
  const startYear = month >= 7 ? currentYear : currentYear - 1;
  const endYear = startYear + 1;
  const yearList = Array.isArray(years) ? years.slice() : [];
  const targetLabel = startYear + '/' + endYear;
  const match = yearList.find(value => String(value || '').includes(targetLabel));
  if (match) return match;
  return targetLabel;
}

function _resolveFinancialYearBounds_(referenceDate) {
  const now = referenceDate instanceof Date ? referenceDate : new Date();
  const currentYear = now.getFullYear();
  const month = now.getMonth() + 1;
  const startYear = month >= 7 ? currentYear : currentYear - 1;
  const endYear = startYear + 1;
  return {
    startYear: startYear,
    endYear: endYear,
    startDate: new Date(startYear, 6, 1),
    endDate: new Date(endYear, 5, 30, 23, 59, 59, 999)
  };
}

function _compareFinancialYears_(a, b) {
  const yearA = _financialYearEnd_(a);
  const yearB = _financialYearEnd_(b);
  if (yearA !== yearB) return yearA - yearB;
  return String(a || '').localeCompare(String(b || ''));
}

function _financialYearEnd_(label) {
  const text = String(label || '').trim();
  if (!text) return 0;
  const match = text.match(/(\d{4})\s*[\/-]\s*(\d{2,4})/);
  if (match) {
    const start = parseInt(match[1], 10);
    const endRaw = match[2];
    const end = endRaw.length === 2 ? parseInt(String(start).slice(0, 2) + endRaw, 10) : parseInt(endRaw, 10);
    return Number.isFinite(end) ? end : 0;
  }
  const years = text.match(/\d{4}/g) || [];
  if (years.length) {
    return parseInt(years[years.length - 1], 10) || 0;
  }
  return 0;
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

  if (typeKey === 'category') {
    if (_valueExistsInColumn_(data, cols.category, trimmed)) return;
    const targetRow = _findRowForInsert_(data, cols.category, [cols.subCategory, cols.accountCodes]);
    _writeRowUpdate_(sheet, data, targetRow, lastCol, {
      [cols.category]: trimmed
    });
    logSystemEventSafe('CREATE_MASTER_DATA', trimmed, 'Type: category');
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

function saveMasterDataRows(payload) {
  if (!payload || !Array.isArray(payload.rows) || !payload.rows.length) {
    throw new Error('No master data rows provided.');
  }

  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);
  if (!sheet) throw new Error('MASTER_DATA not found.');

  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  const cols = _getMasterColumns_(headers);
  const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  const existing = {
    particulars: new Set(),
    subCategory: new Set(),
    category: new Set(),
    accountCodes: new Set(),
    financialYear: new Set()
  };

  data.forEach(function(row) {
    if (cols.particulars) existing.particulars.add(String(row[cols.particulars - 1] || '').trim().toLowerCase());
    if (cols.subCategory) existing.subCategory.add(String(row[cols.subCategory - 1] || '').trim().toLowerCase());
    if (cols.category) existing.category.add(String(row[cols.category - 1] || '').trim().toLowerCase());
    if (cols.accountCodes) existing.accountCodes.add(String(row[cols.accountCodes - 1] || '').trim().toLowerCase());
    if (cols.financialYear) existing.financialYear.add(String(row[cols.financialYear - 1] || '').trim().toLowerCase());
  });

  let addedCount = 0;
  let duplicateCount = 0;
  const errors = [];

  payload.rows.forEach(function(raw, index) {
    const rowIndex = index + 2;
    const particulars = String(raw.Particulars || raw.particulars || '').trim();
    const subCategory = String(raw.Sub_Category || raw.subCategory || '').trim();
    const category = String(raw.Category || raw.category || '').trim();
    const accountCodes = String(raw.Account_Codes || raw.accountCodes || '').trim();
    const accountType = String(raw.Account_Type || raw.accountType || '').trim();
    const reportMapping = String(raw.Report_Mapping || raw.reportMapping || '').trim();
    const financialYear = String(raw.Financial_Year || raw.financialYear || '').trim();

    const hasAny = particulars || subCategory || category || accountCodes || accountType || reportMapping || financialYear;
    if (!hasAny) return;

    const hasCategoryFields = Boolean(particulars || subCategory || category);
    let typeKey = '';
    if (accountCodes) {
      typeKey = 'account';
    } else if (financialYear && !hasCategoryFields) {
      typeKey = 'financialyear';
    } else if (hasCategoryFields) {
      typeKey = 'category_group';
    }

    if (!typeKey) {
      errors.push('Row ' + rowIndex + ': could not determine item type.');
      return;
    }

    if (typeKey === 'account') {
      if (!accountType || !reportMapping) {
        errors.push('Row ' + rowIndex + ': account codes need account type and report mapping.');
        return;
      }
      if (hasCategoryFields || financialYear) {
        errors.push('Row ' + rowIndex + ': banks only use account codes, account type, report mapping.');
        return;
      }
      const key = accountCodes.toLowerCase();
      if (existing.accountCodes.has(key)) {
        duplicateCount += 1;
        return;
      }
      const targetRow = _findRowForInsert_(data, cols.accountCodes, [cols.subCategory]);
      _writeRowUpdate_(sheet, data, targetRow, lastCol, {
        [cols.accountCodes]: accountCodes,
        [cols.accountType]: accountType,
        [cols.reportMapping]: reportMapping
      });
      existing.accountCodes.add(key);
      addedCount += 1;
      return;
    }

    if (typeKey === 'financialyear') {
      const key = financialYear.toLowerCase();
      if (existing.financialYear.has(key)) {
        duplicateCount += 1;
        return;
      }
      if (hasCategoryFields || accountCodes || accountType || reportMapping) {
        errors.push('Row ' + rowIndex + ': financial year entries only use Financial_Year.');
        return;
      }
      const targetRow = _findRowForInsert_(data, cols.financialYear, [cols.subCategory, cols.accountCodes]);
      _writeRowUpdate_(sheet, data, targetRow, lastCol, {
        [cols.financialYear]: financialYear
      });
      existing.financialYear.add(key);
      addedCount += 1;
      return;
    }

    if (typeKey === 'category_group') {
      if (accountCodes || financialYear) {
        errors.push('Row ' + rowIndex + ': categories do not use account codes or financial year.');
        return;
      }
      if (!accountType || !reportMapping) {
        errors.push('Row ' + rowIndex + ': categories need account type and report mapping.');
        return;
      }
      if (particulars) {
        if (!subCategory || !category) {
          errors.push('Row ' + rowIndex + ': particulars need sub-category and category.');
          return;
        }
        const key = particulars.toLowerCase();
        if (existing.particulars.has(key)) {
          duplicateCount += 1;
          return;
        }
        const targetRow = _findRowForInsert_(data, cols.particulars, [cols.accountCodes]);
        _writeRowUpdate_(sheet, data, targetRow, lastCol, {
          [cols.particulars]: particulars,
          [cols.subCategory]: subCategory,
          [cols.category]: category,
          [cols.accountType]: accountType,
          [cols.reportMapping]: reportMapping
        });
        existing.particulars.add(key);
        addedCount += 1;
        return;
      }

      if (subCategory) {
        if (!category) {
          errors.push('Row ' + rowIndex + ': sub-category needs category.');
          return;
        }
        const key = subCategory.toLowerCase();
        if (existing.subCategory.has(key)) {
          duplicateCount += 1;
          return;
        }
        const targetRow = _findRowForInsert_(data, cols.subCategory, [cols.accountCodes]);
        _writeRowUpdate_(sheet, data, targetRow, lastCol, {
          [cols.subCategory]: subCategory,
          [cols.category]: category,
          [cols.accountType]: accountType,
          [cols.reportMapping]: reportMapping
        });
        existing.subCategory.add(key);
        addedCount += 1;
        return;
      }

      if (category) {
        const key = category.toLowerCase();
        if (existing.category.has(key)) {
          duplicateCount += 1;
          return;
        }
        const targetRow = _findRowForInsert_(data, cols.category, [cols.subCategory, cols.accountCodes]);
        _writeRowUpdate_(sheet, data, targetRow, lastCol, {
          [cols.category]: category,
          [cols.accountType]: accountType,
          [cols.reportMapping]: reportMapping
        });
        existing.category.add(key);
        addedCount += 1;
      }
    }
  });

  return { addedCount: addedCount, duplicateCount: duplicateCount, errors: errors };
}

function getMasterDataRows() {
  const ss = _getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_DATA);
  if (!sheet) throw new Error('MASTER_DATA not found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1) return { headers: [], rows: [] };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  return { headers: headers, rows: rows };
}

function _normalizeHeader_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_');
}

function _normalizeContactType_(value) {
  const text = String(value || '').trim().toLowerCase();
  if (!text) return '';
  if (text === 'supplier' || text === 'suppliers') return 'Supplier';
  if (text === 'customer' || text === 'customers') return 'Customer';
  if (text === 'staff' || text === 'employee' || text === 'employees') return 'Staff';
  if (text === 'government entity' || text === 'government entities') return 'Government Entity';
  if (text === 'donor' || text === 'donors') return 'Donor';
  return text.charAt(0).toUpperCase() + text.slice(1);
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
  if (debitValue > 0 && creditValue === 0) {
    return { type: 'receipt', amount: debitValue };
  }
  if (creditValue > 0 && debitValue === 0) {
    return { type: 'payment', amount: creditValue };
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

function _requireReconCriteria_(criteria) {
  if (!criteria) throw new Error('Reconciliation filters are required.');
  const accountCode = String(criteria.accountCode || '').trim();
  const financialYear = String(criteria.financialYear || '').trim();
  const startDate = String(criteria.startDate || '').trim();
  const endDate = String(criteria.endDate || '').trim();
  if (!accountCode || !financialYear || !startDate || !endDate) {
    throw new Error('Select bank account, financial year, start date, and end date.');
  }
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
  // Column positions are resolved by header name only - no fallback numbers
  return {
    payees: _resolveColumn_(headers, ['payees', 'payee'], 0),
    particulars: _resolveColumn_(headers, ['particulars', 'particular'], 0),
    subCategory: _resolveColumn_(headers, ['sub_category', 'subcategory'], 0),
    category: _resolveColumn_(headers, ['category'], 0),
    accountCodes: _resolveColumn_(headers, ['account_codes', 'account_code'], 0),
    accountType: _resolveColumn_(headers, ['account_type', 'accounttype'], 0),
    reportMapping: _resolveColumn_(headers, ['report_mapping', 'reportmapping'], 0),
    financialYear: _resolveColumn_(headers, ['financial_year', 'financialyear'], 0)
  };
}

function _getJournalColumns_(headers) {
  // Column positions are resolved by header name only - no fallback numbers
  // This ensures columns are always found by their exact header name
  return {
    uuid: _resolveColumn_(headers, ['uuid'], 0),
    batchId: _resolveColumn_(headers, ['batch_id', 'batchid'], 0),
    date: _resolveColumn_(headers, ['date'], 0),
    financialYear: _resolveColumn_(headers, ['financial_year', 'financialyear'], 0),
    accountCode: _resolveColumn_(headers, ['account_code', 'accountcode'], 0),
    contactId: _resolveColumn_(headers, ['contact_id', 'contactid'], 0),
    payee: _resolveColumn_(headers, ['payee'], 0),
    refNo: _resolveColumn_(headers, ['ref_no', 'refno'], 0),
    bankRef: _resolveColumn_(headers, ['bank_ref', 'bankref'], 0),
    particulars: _resolveColumn_(headers, ['particulars', 'particular'], 0),
    subCategory: _resolveColumn_(headers, ['sub_category', 'subcategory'], 0),
    category: _resolveColumn_(headers, ['category'], 0),
    description: _resolveColumn_(headers, ['description'], 0),
    debit: _resolveColumn_(headers, ['debit'], 0),
    credit: _resolveColumn_(headers, ['credit'], 0),
    accountType: _resolveColumn_(headers, ['account_type', 'accounttype'], 0),
    reportMapping: _resolveColumn_(headers, ['report_mapping', 'reportmapping'], 0),
    reconStatus: _resolveColumn_(headers, ['recon_status', 'reconcstatus'], 0),
    receiptUrl: _resolveColumn_(headers, ['receipt_url', 'receipturl'], 0),
    advanceId: _resolveColumn_(headers, ['advance_id', 'advanceid'], 0)
  };
}

function _ensureJournalContactColumn_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(_normalizeHeader_);
  if (headers.indexOf('contact_id') >= 0) return;
  const accountIdx = headers.indexOf('account_code');
  const insertAt = accountIdx >= 0 ? accountIdx + 2 : lastCol + 1;
  sheet.insertColumnBefore(insertAt);
  sheet.getRange(1, insertAt).setValue('Contact_ID');
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
  if (accountTypeValue.includes('income') || accountTypeValue.includes('revenue')) return credit - debit;
  if (accountTypeValue.includes('liabil')) return credit - debit;
  if (accountTypeValue.includes('equity') || accountTypeValue.includes('capital')) return credit - debit;
  if (accountTypeValue.includes('expense')) return debit - credit;
  if (accountTypeValue.includes('asset')) return debit - credit;
  return credit - debit;
}

function _resolveCashNotesAmount_(bucket, particulars, meta) {
  const debit = bucket.debit[particulars] || 0;
  const credit = bucket.credit[particulars] || 0;
  const net = debit - credit;
  const classification = _classifyCashFlowLine_(
    meta && meta.reportMapping || '',
    meta && meta.accountType || '',
    debit,
    credit
  );
  if (!classification || classification.section === 'operating') {
    return Math.abs(net);
  }
  return net;
}

function _classifyCashFlowLine_(reportMapping, accountType, debit, credit) {
  const mapping = String(reportMapping || '').toLowerCase();
  const type = String(accountType || '').toLowerCase();
  const direction = credit > 0 ? 'receipt' : 'payment';

  if (mapping.includes('loan') || mapping.includes('borrow') || mapping.includes('debt') || mapping.includes('overdraft') ||
      type.includes('loan') || type.includes('borrow') || type.includes('debt')) {
    return { section: 'financing', direction: direction };
  }

  if (mapping.includes('operating income') || mapping.includes('operating revenue')) {
    return { section: 'operating', direction: 'receipt' };
  }
  if (mapping.includes('operating expense')) {
    return { section: 'operating', direction: 'payment' };
  }

  if (mapping.includes('investment') || mapping.includes('investments')) {
    return { section: 'investing', direction: direction };
  }

  if (mapping.includes('current asset') || mapping.includes('current liability')) {
    return { section: 'operating', direction: direction };
  }

  if (mapping.includes('non-current asset') || mapping.includes('non current asset') || mapping.includes('ppe') || mapping.includes('property, plant') || mapping.includes('property plant') || mapping.includes('property and equipment')) {
    return { section: 'investing', direction: direction };
  }

  if (mapping.includes('non-current liability') || mapping.includes('non current liability')) {
    return { section: 'financing', direction: direction };
  }

  if (mapping.includes('equity')) {
    return { section: 'financing', direction: direction };
  }

  if (type.includes('income') || type.includes('revenue')) {
    return { section: 'operating', direction: 'receipt' };
  }
  if (type.includes('expense')) {
    return { section: 'operating', direction: 'payment' };
  }
  if (type.includes('liabil')) {
    return { section: 'financing', direction: direction };
  }
  if (type.includes('equity') || type.includes('capital')) {
    return { section: 'financing', direction: direction };
  }
  if (type.includes('asset')) {
    return { section: 'operating', direction: direction };
  }

  return { section: 'operating', direction: direction };
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

  if (mappingValues.some(value => value.includes('loan') || value.includes('borrow') || value.includes('debt') || value.includes('overdraft'))) {
    return 'financing';
  }

  if (mappingValues.some(value => value.includes('operating income'))) return 'operating_receipt';
  if (mappingValues.some(value => value.includes('operating expense'))) return 'operating_payment';

  if (typeValues.some(value => value.includes('income')) || name.includes('income') || name.includes('revenue')) {
    return 'operating_receipt';
  }

  if (typeValues.some(value => value.includes('expense')) || name.includes('expense')) {
    return 'operating_payment';
  }

  if (mappingValues.some(value => value.includes('non-current asset') || value.includes('non current asset') || value.includes('ppe') || value.includes('property, plant') || value.includes('property plant') || value.includes('property and equipment'))) {
    return 'investing';
  }

  if (mappingValues.some(value => value.includes('current liability') || value.includes('non-current liability') || value.includes('non current liability'))) {
    return 'financing';
  }

  if (typeValues.some(value => value.includes('loan') || value.includes('borrow') || value.includes('debt'))) return 'financing';
  if (name.includes('liabilit')) return 'financing';
  return null;
}

function _classifyPositionCategory_(categoryName, accountTypes, reportMappings) {
  const name = String(categoryName || '').toLowerCase();
  const typeValues = (accountTypes || []).map(type => String(type || '').toLowerCase());
  const mappingValues = (reportMappings || []).map(value => String(value || '').toLowerCase());

  if (name.includes('cash and cash equivalent')) return 'current_asset';
  if (name.includes('ppe') || name.includes('property, plant') || name.includes('property plant') || name.includes('property and equipment')) {
    return 'non_current_asset';
  }
  if (name.includes('receivable') || name.includes('debtors')) return 'current_asset';
  if (name.includes('payable') || name.includes('creditors')) return 'current_liability';
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

function _buildAccountMeta_(data, cols) {
  const meta = {};
  data.forEach(function(row) {
    const accountCode = cols.accountCodes ? String(row[cols.accountCodes - 1]).trim() : '';
    if (!accountCode) return;
    meta[accountCode] = {
      accountType: cols.accountType ? String(row[cols.accountType - 1]).trim() : '',
      reportMapping: cols.reportMapping ? String(row[cols.reportMapping - 1]).trim() : ''
    };
  });
  return meta;
}
