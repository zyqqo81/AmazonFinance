/***************
 * AMAZON SETTLEMENT IMPORTER (Deposit Date basis) — BALANCED + COGS + AUDIT PACK
 * Google Sheets + Apps Script
 ***************/

const CONFIG = {
  SUMMARY_SHEET: 'ІМПОРТ – ЗВЕДЕННЯ',
  MONTHLY_SHEET: 'МІСЯЦІ',
  PURCHASES_SHEET: 'Закупки',
  TZ: 'Europe/Rome',
  TOTAL_FILE_ID: '__TOTAL__',
  SETTLEMENT_FOLDER_ID: '1K9AuTAmNr5AXHmlTuOIlUdYDRKlKrOAj',
  FOLDER_LIST_LIMIT: 15,
  BULK_IMPORT_CONFIRM_LIMIT: 200,
  MONTHLY_REPORT_FOLDER_ID: 'PUT_MONTHLY_REPORT_FOLDER_ID_HERE',
  MONTHLY_REPORT_SHEET: 'ПДВ ЗВІТ',

  AUDIT: {
    ENABLED: true,
    FOLDER_ID: '1ALCVcKM_3QlEeCedr6DE1YNOI5HKzE2s',
    RAW_LINES_LIMIT: 50000,
    ORDER_ITEMS_LIMIT: 100000,
    SKU_AGG_LIMIT: 100000
  },

  HEADERS: {
    depositDate: 'Deposit Date',
    month: 'Month',
    settlementId: 'Settlement ID',
    marketplace: 'Країна',
    units: 'Units',

    salesNet: 'Продажі (ItemPrice)',
    vatDebito: 'ПДВ',
    feesCost: 'Комісії (ItemFees/Fees)',
    otherNet: 'Інші Комісії Other (net)',
    transfer: 'Виплата на банк (Transfer)',
    payoutExReimbursements: 'Виплата Amazon за продажі (без reimbursement)',

    cogs: 'COGS (Last)',
    netProfit: 'Net Profit',
    amazonReimbursements: 'Amazon Reimbursements',
    soldProfit: 'Sold Profit',
    profitExReimbursements: 'Profit Ex-Reimbursements',
    companyProfit: 'Чистий прибуток компанії',

    unitsWithCost: 'Units With Cost',
    missingUnits: 'Missing Units',
    cogsCoverage: 'COGS Coverage %',
    cogsStatus: 'COGS Status',
    missingSkus: 'Missing SKUs (COGS)',

    fileName: 'Файл',
    fileId: 'File ID',
    importedAt: 'Імпортовано',
    rowCheck: 'Row Check',

    auditUrl: 'Audit URL',
    auditStatus: 'Audit Status'
  },

  PURCHASES: {
    skuHeader: 'SKU',
    unitCostHeader: 'Unit Cost'
  },

  REIMBURSEMENTS: {
    transactionTypeKeywords: ['reimbursement', 'compensation', 'adjustment', 'other-transaction'],
    amountTypeKeywords: ['other', 'misc', 'adjustment', 'missing_from_inbound', 'inventory'],
    amountDescriptionKeywords: [
      'reimbursement',
      'compensation',
      'safe-t',
      'fba inventory reimbursement',
      'lost',
      'damaged',
      'warehouse',
      'inventory',
      'reimbursed',
      'missing_from_inbound'
    ],
    excludedAmountTypes: ['itemprice', 'itemfees', 'fees', 'shipping']
  }
};

/* =========================
 * MENU
 * ========================= */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Фінанси Amazon')
    .addItem('Імпорт місячного звіту (останній із папки)', 'uiImportLatestMonthlyVatReport_')
    .addItem('Імпорт Settlement (останній із папки)', 'uiImportLatestFromFolder_')
    .addItem('Імпорт Settlement (вибір зі списку файлів)', 'uiImportChooseFromFolderList_')
    .addItem('Оновити всі Settlement з папки', 'uiImportAllFromFolder_')
    .addItem('Створити/оновити аудит для вибраного рядка', 'uiCreateAuditForSelectedRow_')
    .addItem('Імпорт Settlement (TSV)', 'uiImportByFileId_')
    .addItem('Налагодження аудиту для вибраного рядка', 'uiDebugAuditForSelectedRow_')
    .addSeparator()
    .addItem('Перебудувати місячну агрегацію', 'rebuildMonthly_')
    .addItem('Перевірити заголовки', 'validateSummarySheet_')
    .addToUi();
}

/* =========================
 * UI
 * ========================= */

function uiImportLatestFromFolder_() {
  const ui = SpreadsheetApp.getUi();
  const warnings = [];
  try {
    validateSummarySheet_();
    const candidates = getSettlementFileCandidatesFromFolder_(CONFIG.SETTLEMENT_FOLDER_ID, CONFIG.FOLDER_LIST_LIMIT);
    if (!candidates.length) {
      ui.alert('Не знайдено settlement-файлів у папці: ' + CONFIG.SETTLEMENT_FOLDER_ID);
      return;
    }

    const fileMeta = candidates[0];
    const msg = importSettlementTxtFile_(fileMeta.id, { warnings: warnings });
    ui.alert(buildUiResultMessage_('Імпорт Settlement (latest) завершено.', msg, warnings));
  } catch (e) {
    handleFatal_('uiImportLatestFromFolder_', e);
    ui.alert(buildUiResultMessage_('Імпорт Settlement (latest) завершився з помилкою.', '', warnings, [toErrorMessage_(e)]));
  }
}

function uiImportLatestMonthlyVatReport_() {
  const ui = SpreadsheetApp.getUi();
  try {
    if (!CONFIG.MONTHLY_REPORT_FOLDER_ID || CONFIG.MONTHLY_REPORT_FOLDER_ID === 'PUT_MONTHLY_REPORT_FOLDER_ID_HERE') {
      ui.alert(
        'Налаштування не завершено',
        'Заповніть CONFIG.MONTHLY_REPORT_FOLDER_ID (ID папки Google Drive з місячними Amazon звітами).',
        ui.ButtonSet.OK
      );
      return;
    }

    const candidates = getSettlementFileCandidatesFromFolder_(CONFIG.MONTHLY_REPORT_FOLDER_ID, 1);
    if (!candidates.length) {
      ui.alert('У папці не знайдено TXT/TSV/CSV файлів: ' + CONFIG.MONTHLY_REPORT_FOLDER_ID);
      return;
    }

    const latest = candidates[0];
    const result = importMonthlyVatReportFile_(latest.id);
    ui.alert(
      'Місячний звіт оброблено',
      [
        'Файл: ' + latest.name,
        'Період: ' + result.monthLabel,
        'Сума продажів: ' + result.sales.toFixed(2),
        'ПДВ до сплати: ' + result.vat.toFixed(2)
      ].join('\n'),
      ui.ButtonSet.OK
    );
  } catch (e) {
    handleFatal_('uiImportLatestMonthlyVatReport_', e);
    ui.alert('Помилка імпорту місячного звіту: ' + toErrorMessage_(e));
  }
}

function uiImportChooseFromFolderList_() {
  const ui = SpreadsheetApp.getUi();
  const warnings = [];
  try {
    validateSummarySheet_();
    const candidates = getSettlementFileCandidatesFromFolder_(CONFIG.SETTLEMENT_FOLDER_ID, 20);
    if (!candidates.length) {
      ui.alert('Не знайдено settlement-файлів у папці: ' + CONFIG.SETTLEMENT_FOLDER_ID);
      return;
    }

    const lines = ['Оберіть номер файлу для імпорту (1..' + candidates.length + '):'];
    for (let i = 0; i < candidates.length; i++) {
      const f = candidates[i];
      lines.push(
        (i + 1) + ') ' + f.name + ' — ' + Utilities.formatDate(f.updatedAt, CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss') + ' — ' + f.size + ' B'
      );
    }

    const res = ui.prompt('Settlement files from folder', lines.join('\n'), ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return;

    const n = Number(String(res.getResponseText() || '').trim());
    if (!isFinite(n) || n < 1 || n > candidates.length) {
      ui.alert('Невірний номер. Введіть число від 1 до ' + candidates.length + '.');
      return;
    }

    const chosen = candidates[n - 1];
    const msg = importSettlementTxtFile_(chosen.id, { warnings: warnings });
    ui.alert(buildUiResultMessage_('Імпорт Settlement (selected) завершено.', msg, warnings));
  } catch (e) {
    handleFatal_('uiImportChooseFromFolderList_', e);
    ui.alert(buildUiResultMessage_('Імпорт Settlement (selected) завершився з помилкою.', '', warnings, [toErrorMessage_(e)]));
  }
}

function uiImportByFileId_() {
  const ui = SpreadsheetApp.getUi();
  const warnings = [];
  try {
    validateSummarySheet_();
    const res = ui.prompt('File ID', 'Встав Drive File ID settlement .txt:', ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return;

    const fileId = String(res.getResponseText() || '').trim();
    if (!fileId) return;

    const msg = importSettlementTxtFile_(fileId, { warnings: warnings });
    ui.alert(buildUiResultMessage_('Імпорт Settlement (manual File ID) завершено.', msg, warnings));
  } catch (e) {
    handleFatal_('uiImportByFileId_', e);
    ui.alert(buildUiResultMessage_('Імпорт Settlement завершився з помилкою.', '', warnings, [toErrorMessage_(e)]));
  }
}

function uiImportAllFromFolder_() {
  const ui = SpreadsheetApp.getUi();
  const warnings = [];
  try {
    validateSummarySheet_();

    const candidates = getSettlementFileCandidatesFromFolder_(CONFIG.SETTLEMENT_FOLDER_ID, Number.MAX_SAFE_INTEGER);
    if (!candidates.length) {
      ui.alert('Не знайдено settlement-файлів у папці: ' + CONFIG.SETTLEMENT_FOLDER_ID);
      return;
    }

    const question = [
      'Знайдено файлів для оновлення: ' + candidates.length,
      'Продовжити масове оновлення всіх settlement?'
    ].join('\n');
    const confirm = ui.alert('Підтвердження масового оновлення', question, ui.ButtonSet.OK_CANCEL);
    if (confirm !== ui.Button.OK) return;

    const result = importAllSettlementsFromFolder_({
      warnings: warnings,
      candidates: candidates,
      progressEvery: 10
    });

    const lines = [
      'Знайдено файлів: ' + result.total,
      'Успішно оновлено: ' + result.imported
    ];

    if (result.failed > 0) {
      lines.push('З помилками: ' + result.failed);
      const limitedErrors = result.errors.slice(0, 10);
      for (let i = 0; i < limitedErrors.length; i++) {
        lines.push('- ' + limitedErrors[i]);
      }
      if (result.errors.length > limitedErrors.length) {
        lines.push('... ще ' + (result.errors.length - limitedErrors.length) + ' помилок. Деталі в Logger.');
      }
    }

    ui.alert(buildUiResultMessage_('Масове оновлення settlement завершено.', lines.join('\n'), warnings));
  } catch (e) {
    handleFatal_('uiImportAllFromFolder_', e);
    ui.alert(buildUiResultMessage_('Масове оновлення settlement завершилося з помилкою.', '', warnings, [toErrorMessage_(e)]));
  }
}


function uiCreateAuditForSelectedRow_() {
  const ui = SpreadsheetApp.getUi();
  const warnings = [];
  try {
    validateSummarySheet_();
    const ctx = buildImportContext_();
    const sh = ctx.sh;
    const active = sh.getActiveRange();
    const r = active ? active.getRow() : 0;
    if (r < 2) throw new Error('Оберіть рядок з даними settlement.');

    const fileId = String(sh.getRange(r, ctx.headerMap[CONFIG.HEADERS.fileId]).getValue() || '').trim();
    if (!fileId || fileId === CONFIG.TOTAL_FILE_ID) throw new Error('У вибраному рядку немає валідного File ID settlement.');

    const msg = importSettlementTxtFile_(fileId, { auditOnly: true, warnings: warnings });
    ui.alert(buildUiResultMessage_('Create/Update Audit виконано.', msg, warnings));
  } catch (e) {
    handleFatal_('uiCreateAuditForSelectedRow_', e);
    ui.alert(buildUiResultMessage_('Create/Update Audit завершився з помилкою.', '', warnings, [toErrorMessage_(e)]));
  }
}


function uiDebugAuditForSelectedRow_() {
  const ui = SpreadsheetApp.getUi();
  const warnings = [];
  try {
    validateSummarySheet_();
    const ctx = buildImportContext_();
    const sh = ctx.sh;
    const active = sh.getActiveRange();
    const r = active ? active.getRow() : 0;
    if (r < 2) throw new Error('Оберіть рядок з даними settlement.');

    const fileId = String(sh.getRange(r, ctx.headerMap[CONFIG.HEADERS.fileId]).getValue() || '').trim();
    if (!fileId || fileId === CONFIG.TOTAL_FILE_ID) throw new Error('У вибраному рядку немає валідного File ID settlement.');

    const msg = importSettlementTxtFile_(fileId, { forceReimport: true, debugAudit: true, warnings: warnings });
    ui.alert(buildUiResultMessage_('Debug Audit виконано.', msg, warnings));
  } catch (e) {
    handleFatal_('uiDebugAuditForSelectedRow_', e);
    ui.alert(buildUiResultMessage_('Debug Audit завершився з помилкою.', '', warnings, [toErrorMessage_(e)]));
  }
}


function getSettlementFileCandidatesFromFolder_(folderId, limit) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const out = [];

  while (files.hasNext()) {
    const f = files.next();
    const name = String(f.getName() || '');
    const lname = name.toLowerCase();
    const mime = String(f.getMimeType() || '').toLowerCase();

    const mimeOk = mime === 'text/plain' || mime === 'application/octet-stream' || mime === 'text/tab-separated-values';
    const nameOk = lname.indexOf('settlement') !== -1 || /\.(txt|tsv)$/i.test(name);
    if (!mimeOk && !nameOk) continue;

    out.push({
      id: f.getId(),
      name: name,
      mimeType: f.getMimeType(),
      size: Number(f.getSize() || 0),
      updatedAt: f.getLastUpdated(),
      createdAt: f.getDateCreated()
    });
  }

  out.sort(function(a, b) {
    const ta = (a.updatedAt && a.updatedAt.getTime()) || (a.createdAt && a.createdAt.getTime()) || 0;
    const tb = (b.updatedAt && b.updatedAt.getTime()) || (b.createdAt && b.createdAt.getTime()) || 0;
    return tb - ta;
  });

  return out.slice(0, Math.max(1, Number(limit) || CONFIG.FOLDER_LIST_LIMIT));
}

function importAllSettlementsFromFolder_(options) {
  options = options || {};
  const warnings = Array.isArray(options.warnings) ? options.warnings : [];
  const candidates = Array.isArray(options.candidates)
    ? options.candidates
    : getSettlementFileCandidatesFromFolder_(CONFIG.SETTLEMENT_FOLDER_ID, Number.MAX_SAFE_INTEGER);
  if (!candidates.length) {
    return { total: 0, imported: 0, failed: 0, errors: [] };
  }

  if (candidates.length > CONFIG.BULK_IMPORT_CONFIRM_LIMIT) {
    warnings.push(
      'Великий обсяг імпорту (' + candidates.length + ' файлів). Apps Script може зупинити виконання через time limit.'
    );
  }

  const errors = [];
  let imported = 0;
  const progressEvery = Math.max(1, Number(options.progressEvery) || 0);
  const rebuildEvery = Math.max(1, Number(options.rebuildEvery) || progressEvery || 10);
  let importedSinceRebuild = 0;

  function runBulkCheckpointRebuild_(reason) {
    if (importedSinceRebuild < 1) return;
    runNonCritical_('ensureMonthAndTotals_ [' + reason + ']', function() {
      ensureMonthAndTotals_(warnings);
    }, warnings);
    runNonCritical_('rebuildMonthly_ [' + reason + ']', function() {
      rebuildMonthly_(warnings);
    }, warnings);
    importedSinceRebuild = 0;
  }

  for (let i = 0; i < candidates.length; i++) {
    const fileMeta = candidates[i];
    try {
      importSettlementTxtFile_(fileMeta.id, {
        warnings: warnings,
        skipPostImportRebuild: true
      });
      imported++;
      importedSinceRebuild++;
    } catch (e) {
      const emsg = '[' + fileMeta.name + '] ' + toErrorMessage_(e);
      errors.push(emsg);
      Logger.log('[BULK IMPORT ERROR] ' + emsg);
    }

    if ((i + 1) % rebuildEvery === 0) {
      runBulkCheckpointRebuild_('checkpoint ' + (i + 1) + '/' + candidates.length);
    }

    if (progressEvery > 0 && ((i + 1) % progressEvery === 0 || i === candidates.length - 1)) {
      safeToast_('Settlement bulk update: ' + (i + 1) + '/' + candidates.length);
    }
  }

  runBulkCheckpointRebuild_('final pass');

  return {
    total: candidates.length,
    imported: imported,
    failed: errors.length,
    errors: errors
  };
}


/* =========================
 * IMPORT CORE
 * ========================= */

function importSettlementTxtFile_(fileId, options) {
  options = options || {};
  const warnings = Array.isArray(options.warnings) ? options.warnings : [];

  const ctx = buildImportContext_();
  const sh = ctx.sh;
  const hm = ctx.headerMap;
  const costMap = ctx.costMap;

  const file = DriveApp.getFileById(fileId);
  const fileName = file.getName();
  const fileMeta = {
    id: file.getId(),
    name: fileName,
    size: Number(file.getSize() || 0),
    updatedAt: file.getLastUpdated(),
    createdAt: file.getDateCreated(),
    mimeType: file.getMimeType()
  };

  const content = file.getBlob().getDataAsString('UTF-8');

  const parsed = parseSettlementTsv_(content, costMap, fileMeta, warnings);
  const rowData = parsed.rowData;

  rowData[CONFIG.HEADERS.fileName] = fileName;
  rowData[CONFIG.HEADERS.fileId] = fileId;
  rowData[CONFIG.HEADERS.importedAt] = new Date();
  rowData[CONFIG.HEADERS.auditStatus] = 'START';
  rowData[CONFIG.HEADERS.auditUrl] = '';

  const rowValues = buildRowFromHeaderMap_(hm, rowData);
  const rowIndex = findOrCreateRowByFileId_(sh, hm, ctx.fileIdRowMap, fileId);

  if (!options.auditOnly) {
    sh.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    runNonCritical_('formatSummaryRow_', function() {
      formatSummaryRow_(sh, hm, rowIndex, warnings);
    }, warnings);
  }

  const auditPayload = buildAuditPayload_(parsed, fileId, fileName);
  let auditResult;

  try {
    auditResult = createSettlementAuditFile_(auditPayload);
  } catch (e) {
    const emsg = toErrorMessage_(e);
    Logger.log('[AUDIT ERR] ' + emsg);
    safeToast_('Audit creation failed: ' + emsg);
    auditResult = { url: '', status: 'ERR:' + emsg, errors: [emsg] };
  }

  const auditStatus = auditResult.url ? ('CREATED:' + auditResult.url) : String(auditResult.status || 'ERR:UNKNOWN');
  writeAuditMetaToSummary_(sh, hm, rowIndex, auditResult.url || '', auditStatus);

  if (!options.auditOnly) {
    runNonCritical_('applyRowCheckAtRow_', function() {
      applyRowCheckAtRow_(sh, hm, rowIndex);
    }, warnings);

    if (!options.skipPostImportRebuild) {
      runNonCritical_('ensureMonthAndTotals_', function() {
        ensureMonthAndTotals_(warnings);
      }, warnings);
      runNonCritical_('rebuildMonthly_', function() {
        rebuildMonthly_(warnings);
      }, warnings);
    }
  }

  const msg = [
    'Імпортовано: ' + fileName,
    'Settlement ID: ' + parsed.settlementId,
    'Deposit Date: ' + Utilities.formatDate(parsed.depositDate, CONFIG.TZ, 'yyyy-MM-dd'),
    'Sales: ' + fromCents_(parsed.salesC).toFixed(2),
    'VAT: ' + fromCents_(parsed.vatC).toFixed(2),
    'Fees: ' + fromCents_(parsed.feesExpenseC).toFixed(2),
    'Other: ' + fromCents_(parsed.otherC).toFixed(2),
    'Transfer: ' + fromCents_(parsed.transferC).toFixed(2),
    'Payout Ex-Reimbursements: ' + fromCents_(parsed.payoutExReimbC).toFixed(2),
    'COGS: ' + fromCents_(parsed.cogsRes.cogsC).toFixed(2),
    'Amazon Reimbursements: ' + fromCents_(parsed.reimbursementsC).toFixed(2),
    'Company Profit: ' + fromCents_(parsed.companyProfitC).toFixed(2),
    'Sold Profit: ' + fromCents_(parsed.soldProfitC).toFixed(2),
    'Row Check: ' + parsed.rowCheck,
    'Audit Status: ' + auditStatus
  ].join('\n');

  return msg;
}

function importMonthlyVatReportFile_(fileId) {
  const file = DriveApp.getFileById(fileId);
  const content = file.getBlob().getDataAsString('UTF-8');
  const report = parseMonthlyVatReport_(content, file.getName());

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(CONFIG.MONTHLY_REPORT_SHEET) || ss.insertSheet(CONFIG.MONTHLY_REPORT_SHEET);
  const headers = ['Month', 'Sales Total', 'VAT To Pay', 'Rows', 'Файл', 'File ID', 'Імпортовано'];
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  upsertMonthlyVatRow_(sheet, {
    monthDate: report.monthDate,
    sales: report.sales,
    vat: report.vat,
    rows: report.rows,
    fileName: file.getName(),
    fileId: fileId,
    importedAt: new Date()
  });

  if (sheet.getLastRow() > 1) {
    safeSetNumberFormat_(sheet.getRange(2, 1, sheet.getLastRow() - 1, 1), 'yyyy-MM', [], 'monthlyVat.month');
    safeSetNumberFormat_(sheet.getRange(2, 2, sheet.getLastRow() - 1, 2), '#,##0.00', [], 'monthlyVat.money');
  }

  return {
    monthLabel: Utilities.formatDate(report.monthDate, 'UTC', 'yyyy-MM'),
    sales: report.sales,
    vat: report.vat,
    rows: report.rows
  };
}

function parseMonthlyVatReport_(content, fileName) {
  const lines = splitLines_(content).filter(function(line) { return line.trim() !== ''; });
  if (lines.length < 2) throw new Error('Файл порожній або не містить даних: ' + fileName);

  const sep = detectSep_(lines[0]);
  const headers = splitTsvLine_(lines[0], sep).map(normalizeHeader_);

  const salesIdx = findHeaderIdx_(headers, ['item-price', 'itemprice', 'principal', 'sales', 'product sales']);
  const vatIdx = findHeaderIdx_(headers, ['item-related-fee-tax', 'tax', 'vat', 'vat amount', 'itemtax']);
  const dateIdx = findHeaderIdx_(headers, ['posted-date', 'date/time', 'settlement-start-date', 'transaction-date', 'date']);

  if (salesIdx < 0) throw new Error('Не знайдено колонку продажів (ItemPrice/Principal/Sales).');
  if (vatIdx < 0) throw new Error('Не знайдено колонку ПДВ (VAT/Tax).');

  let salesC = 0;
  let vatC = 0;
  let rows = 0;
  let firstDate = null;

  for (let i = 1; i < lines.length; i++) {
    const cols = splitTsvLine_(lines[i], sep);
    if (!cols.length) continue;

    salesC += toCents_(parseSmartNumber_(cols[salesIdx]));
    vatC += toCents_(parseSmartNumber_(cols[vatIdx]));
    rows++;

    if (!firstDate && dateIdx >= 0) {
      const d = parseDateMaybe_(cols[dateIdx]);
      if (d) firstDate = d;
    }
  }

  const monthDate = firstDate
    ? new Date(Date.UTC(firstDate.getUTCFullYear(), firstDate.getUTCMonth(), 1))
    : parseMonthFromFilename_(fileName);

  if (!monthDate) throw new Error('Не вдалося визначити місяць: додайте дату у файл або у назву файлу (наприклад 2025-01).');

  return {
    monthDate: monthDate,
    sales: fromCents_(salesC),
    vat: fromCents_(vatC),
    rows: rows
  };
}

function upsertMonthlyVatRow_(sheet, entry) {
  const key = Utilities.formatDate(entry.monthDate, 'UTC', 'yyyy-MM');
  const lastRow = sheet.getLastRow();
  const rowValues = [entry.monthDate, entry.sales, entry.vat, entry.rows, entry.fileName, entry.fileId, entry.importedAt];

  if (lastRow < 2) {
    sheet.getRange(2, 1, 1, rowValues.length).setValues([rowValues]);
    return;
  }

  const existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < existing.length; i++) {
    const d = existing[i][0];
    if (d instanceof Date && !isNaN(d.getTime()) && Utilities.formatDate(d, 'UTC', 'yyyy-MM') === key) {
      sheet.getRange(i + 2, 1, 1, rowValues.length).setValues([rowValues]);
      return;
    }
  }

  sheet.getRange(lastRow + 1, 1, 1, rowValues.length).setValues([rowValues]);
}

function parseSettlementTsv_(content, costMap, fileMeta, warnings) {
  const parsedTsv = parseTsv_(content, fileMeta);
  const header = parsedTsv.headers;
  const dataRows = parsedTsv.rows;
  const idx = indexMapFlexible_(header);

  const firstRow = dataRows[0];
  if (!firstRow) {
    throw buildFileDiagnosticError_('TSV не містить data rows після header.', fileMeta, content, header, []);
  }

  const settlementId = cellByHeader_(firstRow, idx, 'settlement-id');
  const depositDateRaw = cellByHeader_(firstRow, idx, 'deposit-date');
  const transferC = detectTransferC_(dataRows, idx, warnings || []);
  let marketplaceName = cellByHeader_(firstRow, idx, 'marketplace-name');

  const depositDate = parseDateFlexible_(depositDateRaw, CONFIG.TZ);
  if (!(depositDate instanceof Date) || isNaN(depositDate.getTime())) {
    throw buildFileDiagnosticError_('Не вдалося розпізнати Deposit Date: "' + depositDateRaw + '"', fileMeta, content, header, [
      'Підтримуються формати: YYYY-MM-DD, DD/MM/YYYY, MM/DD/YYYY, з часом/таймзоною.'
    ]);
  }

  const monthDate = new Date(Date.UTC(depositDate.getUTCFullYear(), depositDate.getUTCMonth(), 1));

  let salesC = 0;
  let vatC = 0;
  let itemFeesSignedSumC = 0;
  let feeNeg = 0;
  let feePos = 0;
  let reimbursementsC = 0;

  const skuQtyMap = Object.create(null);
  const orderAgg = Object.create(null);
  const rawRows = [];

  const idxOrderId = idx['order-id'];
  const idxSku = idx['sku'];
  const idxQty = idx['quantity-purchased'];

  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];

    const amountType = cellByHeader_(row, idx, 'amount-type');
    const amountDesc = cellByHeader_(row, idx, 'amount-description');
    const transactionType = cellByHeader_(row, idx, 'transaction-type');
    const amountC = toCents_(parseNumberLoose_(cellByHeader_(row, idx, 'amount')));

    if (!marketplaceName) {
      const mp = cellByHeader_(row, idx, 'marketplace-name');
      if (mp) marketplaceName = mp;
    }

    if (rawRows.length < CONFIG.AUDIT.RAW_LINES_LIMIT) rawRows.push(row);

    if (isReimbursementLine_(transactionType, amountType, amountDesc)) reimbursementsC += amountC;

    const t = String(amountType || '').trim();
    const d = String(amountDesc || '').trim();

    if (t === 'ItemPrice') {
      if (d === 'Principal' || d === 'Shipping' || d === 'GiftWrap') salesC += amountC;
      else if (d === 'Tax' || d === 'ShippingTax' || d === 'GiftWrapTax') vatC += amountC;
    }

    if (t === 'ItemFees' || t === 'Fees') {
      itemFeesSignedSumC += amountC;
      if (amountC < 0) feeNeg++;
      if (amountC > 0) feePos++;
    }

    const sku = idxSku !== undefined ? normalizeSku_(row[idxSku]) : '';
    const qty = idxQty !== undefined ? Math.max(0, Math.round(parseNumberLoose_(row[idxQty]))) : 0;
    const orderId = idxOrderId !== undefined ? String(row[idxOrderId] || '').trim() : '';

    if (t === 'ItemPrice' && d === 'Principal' && sku && qty > 0) skuQtyMap[sku] = (skuQtyMap[sku] || 0) + qty;

    if (sku || orderId) {
      const k = orderId + '||' + sku;
      if (!orderAgg[k]) {
        orderAgg[k] = { orderId: orderId, sku: sku, qty: 0, principalC: 0, taxC: 0, feesSignedC: 0 };
      }
      if (t === 'ItemPrice' && d === 'Principal') {
        orderAgg[k].qty += qty;
        orderAgg[k].principalC += amountC;
      }
      if (t === 'ItemPrice' && (d === 'Tax' || d === 'ShippingTax' || d === 'GiftWrapTax')) orderAgg[k].taxC += amountC;
      if (t === 'ItemFees' || t === 'Fees') orderAgg[k].feesSignedC += amountC;
    }
  }

  const feesNorm = normalizeFeesExpenseC_(itemFeesSignedSumC, feeNeg, feePos);
  const feesExpenseC = feesNorm.feesExpenseC;

  const units = Object.keys(skuQtyMap).reduce(function(acc, sku) { return acc + Number(skuQtyMap[sku] || 0); }, 0);
  const cogsRes = calcCogsFromCostMap_(skuQtyMap, costMap);

  const otherC = transferC - (salesC + vatC - feesExpenseC);
  const payoutExReimbC = transferC - reimbursementsC;
  const netCashC = transferC - cogsRes.cogsC;
  const soldProfitC = payoutExReimbC - cogsRes.cogsC;
  const profitExReimbC = soldProfitC;
  const companyProfitC = netCashC;

  const diffC = (salesC + vatC + otherC - feesExpenseC) - transferC;
  const rowCheck = Math.abs(diffC) <= 1 ? 'OK' : ('ERR diff ' + fromCents_(diffC).toFixed(2));

  const rowData = {};
  rowData[CONFIG.HEADERS.depositDate] = depositDate;
  rowData[CONFIG.HEADERS.month] = monthDate;
  rowData[CONFIG.HEADERS.settlementId] = settlementId;
  rowData[CONFIG.HEADERS.marketplace] = marketplaceName ? mapMarketplace_(marketplaceName) : '';
  rowData[CONFIG.HEADERS.units] = units;

  rowData[CONFIG.HEADERS.salesNet] = fromCents_(salesC);
  rowData[CONFIG.HEADERS.vatDebito] = fromCents_(vatC);
  rowData[CONFIG.HEADERS.feesCost] = fromCents_(feesExpenseC);
  rowData[CONFIG.HEADERS.otherNet] = fromCents_(otherC);
  rowData[CONFIG.HEADERS.transfer] = fromCents_(transferC);
  rowData[CONFIG.HEADERS.payoutExReimbursements] = fromCents_(payoutExReimbC);

  rowData[CONFIG.HEADERS.cogs] = fromCents_(cogsRes.cogsC);
  rowData[CONFIG.HEADERS.netProfit] = fromCents_(netCashC);
  rowData[CONFIG.HEADERS.amazonReimbursements] = fromCents_(reimbursementsC);
  rowData[CONFIG.HEADERS.soldProfit] = fromCents_(soldProfitC);
  rowData[CONFIG.HEADERS.profitExReimbursements] = fromCents_(profitExReimbC);
  rowData[CONFIG.HEADERS.companyProfit] = fromCents_(companyProfitC);

  rowData[CONFIG.HEADERS.unitsWithCost] = cogsRes.unitsWithCost;
  rowData[CONFIG.HEADERS.missingUnits] = cogsRes.missingUnits;
  rowData[CONFIG.HEADERS.cogsCoverage] = cogsRes.coveragePct;
  rowData[CONFIG.HEADERS.cogsStatus] = cogsRes.missingUnits > 0 ? 'MISSING_COST' : 'OK';
  rowData[CONFIG.HEADERS.missingSkus] = cogsRes.missingSkusText;
  rowData[CONFIG.HEADERS.rowCheck] = rowCheck;

  return {
    header: header,
    rawRows: rawRows,
    idx: idx,
    settlementId: settlementId,
    depositDate: depositDate,
    monthDate: monthDate,
    marketplaceName: marketplaceName,
    transferC: transferC,
    payoutExReimbC: payoutExReimbC,
    salesC: salesC,
    vatC: vatC,
    feesExpenseC: feesExpenseC,
    feesRule: feesNorm.feesRule,
    reimbursementsC: reimbursementsC,
    otherC: otherC,
    soldProfitC: soldProfitC,
    profitExReimbC: profitExReimbC,
    companyProfitC: companyProfitC,
    rowCheck: rowCheck,
    orderAgg: orderAgg,
    skuQtyMap: skuQtyMap,
    cogsRes: cogsRes,
    units: units,
    rowData: rowData
  };
}

/* =========================
 * MONTHLY
 * ========================= */

function rebuildMonthly_(warnings) {
  warnings = warnings || [];
  const ss = SpreadsheetApp.getActive();
  const summary = ss.getSheetByName(CONFIG.SUMMARY_SHEET);
  if (!summary) throw new Error('Не знайдено вкладку "' + CONFIG.SUMMARY_SHEET + '"');

  ensureSummaryHeaders_(summary);
  const hm = getHeaderMap_(summary);

  const required = [
    CONFIG.HEADERS.month,
    CONFIG.HEADERS.fileId,
    CONFIG.HEADERS.salesNet,
    CONFIG.HEADERS.vatDebito,
    CONFIG.HEADERS.feesCost,
    CONFIG.HEADERS.otherNet,
    CONFIG.HEADERS.transfer,
    CONFIG.HEADERS.payoutExReimbursements,
    CONFIG.HEADERS.units,
    CONFIG.HEADERS.cogs,
    CONFIG.HEADERS.netProfit,
    CONFIG.HEADERS.amazonReimbursements,
    CONFIG.HEADERS.soldProfit,
    CONFIG.HEADERS.profitExReimbursements,
    CONFIG.HEADERS.companyProfit
  ];

  const missing = required.filter(function(h) { return !hm[h]; });
  if (missing.length) throw new Error('Не вистачає заголовків: ' + missing.join(', '));

  const monthly = ss.getSheetByName(CONFIG.MONTHLY_SHEET) || ss.insertSheet(CONFIG.MONTHLY_SHEET);
  monthly.clearContents();

  const headers = [
    'Month',
    'Sales',
    'VAT',
    'Fees',
    'Other',
    'Transfer',
    'Payout Ex-Reimbursements',
    'Amazon Reimbursements',
    'Units',
    'COGS',
    'Net Profit (cash)',
    'Sold Profit',
    'Profit Ex-Reimbursements',
    'Company Profit',
    'Reconcile'
  ];
  monthly.getRange(1, 1, 1, headers.length).setValues([headers]);

  const lastRow = summary.getLastRow();
  if (lastRow < 2) return;

  const values = summary.getRange(2, 1, lastRow - 1, summary.getLastColumn()).getValues();

  const c = {
    month: hm[CONFIG.HEADERS.month] - 1,
    fileId: hm[CONFIG.HEADERS.fileId] - 1,
    sales: hm[CONFIG.HEADERS.salesNet] - 1,
    vat: hm[CONFIG.HEADERS.vatDebito] - 1,
    fees: hm[CONFIG.HEADERS.feesCost] - 1,
    other: hm[CONFIG.HEADERS.otherNet] - 1,
    transfer: hm[CONFIG.HEADERS.transfer] - 1,
    payoutExReimb: hm[CONFIG.HEADERS.payoutExReimbursements] - 1,
    reimb: hm[CONFIG.HEADERS.amazonReimbursements] - 1,
    units: hm[CONFIG.HEADERS.units] - 1,
    cogs: hm[CONFIG.HEADERS.cogs] - 1,
    net: hm[CONFIG.HEADERS.netProfit] - 1,
    sold: hm[CONFIG.HEADERS.soldProfit] - 1,
    ex: hm[CONFIG.HEADERS.profitExReimbursements] - 1,
    company: hm[CONFIG.HEADERS.companyProfit] - 1
  };

  const bucket = Object.create(null);

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const fid = String(row[c.fileId] || '').trim();
    if (!fid || fid === CONFIG.TOTAL_FILE_ID) continue;

    const m = row[c.month];
    if (!(m instanceof Date) || isNaN(m.getTime())) continue;

    const key = Utilities.formatDate(m, 'UTC', 'yyyy-MM');
    if (!bucket[key]) {
      bucket[key] = {
        monthDate: new Date(Date.UTC(m.getUTCFullYear(), m.getUTCMonth(), 1)),
        salesC: 0,
        vatC: 0,
        feesC: 0,
        otherC: 0,
        transferC: 0,
        payoutExReimbC: 0,
        reimbC: 0,
        units: 0,
        cogsC: 0,
        netC: 0,
        soldC: 0,
        exC: 0,
        companyC: 0
      };
    }

    const b = bucket[key];
    b.salesC += toCents_(row[c.sales]);
    b.vatC += toCents_(row[c.vat]);
    b.feesC += toCents_(row[c.fees]);
    b.otherC += toCents_(row[c.other]);
    b.transferC += toCents_(row[c.transfer]);
    b.payoutExReimbC += toCents_(row[c.payoutExReimb]);
    b.reimbC += toCents_(row[c.reimb]);
    b.units += Math.round(Number(row[c.units]) || 0);
    b.cogsC += toCents_(row[c.cogs]);
    b.netC += toCents_(row[c.net]);
    b.soldC += toCents_(row[c.sold]);
    b.exC += toCents_(row[c.ex]);
    b.companyC += toCents_(row[c.company]);
  }

  const keys = Object.keys(bucket).sort();
  if (!keys.length) return;

  const out = keys.map(function(k) {
    const b = bucket[k];
    const diffC = (b.salesC + b.vatC + b.otherC - b.feesC) - b.transferC;
    return [
      b.monthDate,
      fromCents_(b.salesC),
      fromCents_(b.vatC),
      fromCents_(b.feesC),
      fromCents_(b.otherC),
      fromCents_(b.transferC),
      fromCents_(b.payoutExReimbC),
      fromCents_(b.reimbC),
      b.units,
      fromCents_(b.cogsC),
      fromCents_(b.netC),
      fromCents_(b.soldC),
      fromCents_(b.exC),
      fromCents_(b.companyC),
      Math.abs(diffC) <= 1 ? 'OK' : ('ERR ' + fromCents_(diffC).toFixed(2))
    ];
  });

  monthly.getRange(2, 1, out.length, headers.length).setValues(out);
  applyMonthlyFormats_(monthly, out.length, warnings);
}

function rebuildMonthlySheet_() {
  rebuildMonthly_([]);
}

/* =========================
 * SUMMARY MAINTENANCE
 * ========================= */

function validateSummarySheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONFIG.SUMMARY_SHEET);
  if (!sh) throw new Error('Не знайдено вкладку "' + CONFIG.SUMMARY_SHEET + '"');

  ensureSummaryHeaders_(sh);
  const hm = getHeaderMap_(sh);

  const required = Object.keys(CONFIG.HEADERS).map(function(k) { return CONFIG.HEADERS[k]; });
  const missing = required.filter(function(h) { return !hm[h]; });
  if (missing.length) throw new Error('Не вистачає заголовків: ' + missing.join(', '));

  const p = ss.getSheetByName(CONFIG.PURCHASES_SHEET);
  if (!p) throw new Error('Не знайдено вкладку "' + CONFIG.PURCHASES_SHEET + '"');
  const phm = getHeaderMap_(p);

  if (!phm[CONFIG.PURCHASES.skuHeader] || !phm[CONFIG.PURCHASES.unitCostHeader]) {
    throw new Error('У вкладці "' + CONFIG.PURCHASES_SHEET + '" потрібні колонки: ' + CONFIG.PURCHASES.skuHeader + ', ' + CONFIG.PURCHASES.unitCostHeader);
  }
}

function ensureMonthAndTotals_(warnings) {
  warnings = warnings || [];
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SUMMARY_SHEET);
  if (!sh) throw new Error('Не знайдено вкладку "' + CONFIG.SUMMARY_SHEET + '"');

  ensureSummaryHeaders_(sh);
  const hm = getHeaderMap_(sh);

  const colFileId = hm[CONFIG.HEADERS.fileId];
  const colMonth = hm[CONFIG.HEADERS.month];
  const colDeposit = hm[CONFIG.HEADERS.depositDate];
  const colSettlId = hm[CONFIG.HEADERS.settlementId];
  if (!colFileId || !colMonth || !colDeposit) return;

  let lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  for (let r = lastRow; r >= 2; r--) {
    const fid = String(sh.getRange(r, colFileId).getValue() || '').trim();
    if (fid === CONFIG.TOTAL_FILE_ID) sh.deleteRow(r);
  }

  lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const ids = sh.getRange(2, colFileId, lastRow - 1, 1).getValues();
  const deposits = sh.getRange(2, colDeposit, lastRow - 1, 1).getValues();

  let lastDataRow = 1;
  const monthVals = [];

  for (let i = 0; i < ids.length; i++) {
    const fid = String(ids[i][0] || '').trim();
    if (!fid || fid === CONFIG.TOTAL_FILE_ID) {
      monthVals.push(['']);
      continue;
    }

    const d = deposits[i][0];
    if (d instanceof Date && !isNaN(d.getTime())) {
      monthVals.push([new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), 1))]);
      lastDataRow = 2 + i;
    } else {
      monthVals.push(['']);
    }
  }

  sh.getRange(2, colMonth, monthVals.length, 1).setValues(monthVals);
  safeSetNumberFormat_(sh.getRange(2, colMonth, monthVals.length, 1), 'yyyy-MM', warnings, 'summary.month');

  if (lastDataRow < 2) return;

  const spacerRow = lastDataRow + 1;
  const totalRow = lastDataRow + 2;

  sh.getRange(spacerRow, 1, 1, sh.getLastColumn()).clearContent();
  sh.getRange(totalRow, colFileId).setValue(CONFIG.TOTAL_FILE_ID);
  if (colSettlId) sh.getRange(totalRow, colSettlId).setValue('TOTAL (filtered)');

  const moneyCols = [
    hm[CONFIG.HEADERS.salesNet],
    hm[CONFIG.HEADERS.vatDebito],
    hm[CONFIG.HEADERS.feesCost],
    hm[CONFIG.HEADERS.otherNet],
    hm[CONFIG.HEADERS.transfer],
    hm[CONFIG.HEADERS.payoutExReimbursements],
    hm[CONFIG.HEADERS.cogs],
    hm[CONFIG.HEADERS.netProfit],
    hm[CONFIG.HEADERS.amazonReimbursements],
    hm[CONFIG.HEADERS.soldProfit],
    hm[CONFIG.HEADERS.profitExReimbursements],
    hm[CONFIG.HEADERS.companyProfit]
  ].filter(Boolean);

  const unitCols = [
    hm[CONFIG.HEADERS.units],
    hm[CONFIG.HEADERS.unitsWithCost],
    hm[CONFIG.HEADERS.missingUnits]
  ].filter(Boolean);

  const sep = getFormulaArgSeparator_();
  const allSumCols = moneyCols.concat(unitCols);

  for (let i = 0; i < allSumCols.length; i++) {
    const c = allSumCols[i];
    const rangeA1 = colToA1_(c) + '2:' + colToA1_(c) + lastDataRow;
    sh.getRange(totalRow, c).setFormula('=SUBTOTAL(109' + sep + rangeA1 + ')');
  }

  for (let i = 0; i < moneyCols.length; i++) {
    safeSetNumberFormat_(sh.getRange(totalRow, moneyCols[i], 1, 1), '#,##0.00', warnings, 'summary.total.money');
  }
  for (let i = 0; i < unitCols.length; i++) {
    safeSetNumberFormat_(sh.getRange(totalRow, unitCols[i], 1, 1), '0', warnings, 'summary.total.units');
  }

  applyRowCheckAll_();
}

function applyRowCheckAll_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SUMMARY_SHEET);
  if (!sh) return;

  ensureSummaryHeaders_(sh);
  const hm = getHeaderMap_(sh);

  const req = [
    hm[CONFIG.HEADERS.salesNet],
    hm[CONFIG.HEADERS.vatDebito],
    hm[CONFIG.HEADERS.otherNet],
    hm[CONFIG.HEADERS.feesCost],
    hm[CONFIG.HEADERS.transfer],
    hm[CONFIG.HEADERS.rowCheck],
    hm[CONFIG.HEADERS.fileId]
  ];

  if (req.some(function(c) { return !c; })) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  for (let r = 2; r <= lastRow; r++) {
    applyRowCheckAtRow_(sh, hm, r);
  }
}

function applyRowCheckAtRow_(sh, hm, r) {
  const fid = String(sh.getRange(r, hm[CONFIG.HEADERS.fileId]).getValue() || '').trim();
  if (!fid || fid === CONFIG.TOTAL_FILE_ID) {
    sh.getRange(r, hm[CONFIG.HEADERS.rowCheck]).clearContent();
    return;
  }

  const salesC = toCents_(sh.getRange(r, hm[CONFIG.HEADERS.salesNet]).getValue());
  const vatC = toCents_(sh.getRange(r, hm[CONFIG.HEADERS.vatDebito]).getValue());
  const otherC = toCents_(sh.getRange(r, hm[CONFIG.HEADERS.otherNet]).getValue());
  const feesC = toCents_(sh.getRange(r, hm[CONFIG.HEADERS.feesCost]).getValue());
  const transferC = toCents_(sh.getRange(r, hm[CONFIG.HEADERS.transfer]).getValue());

  const diffC = (salesC + vatC + otherC - feesC) - transferC;
  const status = Math.abs(diffC) <= 1 ? 'OK' : ('ERR diff ' + fromCents_(diffC).toFixed(2));
  sh.getRange(r, hm[CONFIG.HEADERS.rowCheck]).setValue(status);
}

/* =========================
 * AUDIT
 * ========================= */

function buildAuditPayload_(parsed, fileId, fileName) {
  const orderRows = buildOrderRows_(parsed.orderAgg, parsed.cogsRes.costBySku || {});
  const skuRows = buildSkuRows_(parsed.skuQtyMap, parsed.orderAgg, parsed.cogsRes.costBySku || {}, parsed.cogsRes.missingSkus || []);

  return {
    fileId: fileId,
    fileName: fileName,
    settlementId: parsed.settlementId,
    depositDate: parsed.depositDate,
    marketplace: parsed.marketplaceName,

    sales: fromCents_(parsed.salesC),
    vat: fromCents_(parsed.vatC),
    fees: fromCents_(parsed.feesExpenseC),
    other: fromCents_(parsed.otherC),
    transfer: fromCents_(parsed.transferC),
    payoutExReimb: fromCents_(parsed.payoutExReimbC),
    cogs: fromCents_(parsed.cogsRes.cogsC),
    reimbursements: fromCents_(parsed.reimbursementsC),
    soldProfit: fromCents_(parsed.soldProfitC),
    profitExReimb: fromCents_(parsed.profitExReimbC),
    companyProfit: fromCents_(parsed.companyProfitC),
    netCash: fromCents_(parsed.transferC - parsed.cogsRes.cogsC),

    units: parsed.units,
    unitsWithCost: parsed.cogsRes.unitsWithCost,
    missingUnits: parsed.cogsRes.missingUnits,
    coverage: parsed.cogsRes.coveragePct,
    missingSkusText: parsed.cogsRes.missingSkusText,
    rowCheck: parsed.rowCheck,

    tsvHeader: parsed.header,
    rawRows: parsed.rawRows,
    orderRows: orderRows,
    skuRows: skuRows
  };
}

function createSettlementAuditFile_(audit) {
  if (!CONFIG.AUDIT.ENABLED) return { url: '', status: 'DISABLED', errors: [] };

  const folder = DriveApp.getFolderById(CONFIG.AUDIT.FOLDER_ID);
  const dep = Utilities.formatDate(audit.depositDate, CONFIG.TZ, 'yyyy-MM-dd');
  const shortId = String(audit.fileId || '').slice(0, 8);
  const safeSettlement = String(audit.settlementId || 'UNKNOWN').replace(/[^a-zA-Z0-9_-]/g, '_');
  const name = 'SETTLEMENT_AUDIT_' + dep + '__' + safeSettlement + '__' + shortId;

  const aSs = SpreadsheetApp.create(name);
  const aFile = DriveApp.getFileById(aSs.getId());
  aFile.moveTo(folder);

  const errors = [];
  const warnings = [];

  try {
    writeAuditSummaryTab_(aSs.getSheets()[0], audit, errors, warnings);
  } catch (e) {
    errors.push('SUMMARY:' + toErrorMessage_(e));
  }

  try {
    writeAuditOrderItemsTab_(aSs, audit, warnings);
  } catch (e) {
    errors.push('ORDER_ITEMS:' + toErrorMessage_(e));
  }

  try {
    writeAuditSkuAggTab_(aSs, audit, warnings);
  } catch (e) {
    errors.push('SKU_AGG:' + toErrorMessage_(e));
  }

  try {
    writeAuditRawLinesTab_(aSs, audit, warnings);
  } catch (e) {
    errors.push('RAW_LINES:' + toErrorMessage_(e));
  }

  if (errors.length) {
    Logger.log('[AUDIT PARTIAL] ' + errors.join(' | '));
    safeToast_('Audit partial: ' + errors.join(' | '));
    return { url: aSs.getUrl(), status: 'ERR:' + errors.join(' | '), errors: errors };
  }

  if (warnings.length) {
    Logger.log('[AUDIT WARN] ' + warnings.join(' | '));
    return { url: aSs.getUrl(), status: 'OK_WARN:' + warnings.join(' | '), errors: [] };
  }

  return { url: aSs.getUrl(), status: 'OK', errors: [] };
}

function writeAuditSummaryTab_(sheet, audit, errors, warnings) {
  sheet.setName('SUMMARY');

  const rows = [
    ['Settlement ID', audit.settlementId],
    ['File ID', audit.fileId],
    ['File Name', audit.fileName],
    ['Deposit Date', Utilities.formatDate(audit.depositDate, CONFIG.TZ, 'yyyy-MM-dd')],
    ['Marketplace', audit.marketplace],
    ['', ''],
    ['Sales', audit.sales],
    ['VAT', audit.vat],
    ['Fees', audit.fees],
    ['Other', audit.other],
    ['Transfer', audit.transfer],
    ['Payout Ex-Reimbursements', audit.payoutExReimb],
    ['COGS (Last)', audit.cogs],
    ['Net Profit (cash)', audit.netCash],
    ['Amazon Reimbursements', audit.reimbursements],
    ['Sold Profit', audit.soldProfit],
    ['Profit Ex-Reimbursements', audit.profitExReimb],
    ['Company Profit', audit.companyProfit],
    ['', ''],
    ['Units', audit.units],
    ['Units With Cost', audit.unitsWithCost],
    ['Missing Units', audit.missingUnits],
    ['Coverage', audit.coverage],
    ['Missing SKUs', audit.missingSkusText],
    ['Row Check', audit.rowCheck],
    ['', ''],
    ['Tab Errors', errors && errors.length ? errors.join(' | ') : ''],
    ['Tab Warnings', warnings && warnings.length ? warnings.join(' | ') : '']
  ];

  const rect = normalize2D_(rows, 2);
  sheet.getRange(1, 1, rect.length, 2).setValues(rect);

  applyAuditFormats_(sheet, 7, 12, [{ col: 2, pattern: '#,##0.00', context: 'audit.summary.money' }], warnings || []);
  applyAuditFormats_(sheet, 20, 3, [{ col: 2, pattern: '0', context: 'audit.summary.units' }], warnings || []);
  applyAuditFormats_(sheet, 23, 1, [{ col: 2, pattern: '0.00%', context: 'audit.summary.coverage' }], warnings || []);
  sheet.autoResizeColumns(1, 2);
}

function writeAuditOrderItemsTab_(ss, audit, warnings) {
  const sh = ss.insertSheet('ORDER_ITEMS');
  const hdr = ['order-id', 'sku', 'qty', 'principal', 'tax', 'fees_signed', 'fees_expense', 'last_unit_cost', 'cogs', 'profit'];
  sh.getRange(1, 1, 1, hdr.length).setValues([hdr]);

  const rows = normalize2D_(audit.orderRows.slice(0, CONFIG.AUDIT.ORDER_ITEMS_LIMIT), hdr.length);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, hdr.length).setValues(rows);
    applyAuditFormats_(sh, 2, rows.length, [
      { col: 3, pattern: '0', context: 'audit.order.qty' },
      { col: 4, pattern: '#,##0.00', context: 'audit.order.principal' },
      { col: 5, pattern: '#,##0.00', context: 'audit.order.tax' },
      { col: 6, pattern: '#,##0.00', context: 'audit.order.fees_signed' },
      { col: 7, pattern: '#,##0.00', context: 'audit.order.fees_expense' },
      { col: 8, pattern: '#,##0.00', context: 'audit.order.last_cost' },
      { col: 9, pattern: '#,##0.00', context: 'audit.order.cogs' },
      { col: 10, pattern: '#,##0.00', context: 'audit.order.profit' }
    ], warnings || []);
  }

  sh.autoResizeColumns(1, hdr.length);
}

function writeAuditSkuAggTab_(ss, audit, warnings) {
  const sh = ss.insertSheet('SKU_AGG');
  const hdr = ['sku', 'qty', 'principal', 'tax', 'fees_expense', 'last_unit_cost', 'cogs', 'profit', 'has_cost'];
  sh.getRange(1, 1, 1, hdr.length).setValues([hdr]);

  const rows = normalize2D_(audit.skuRows.slice(0, CONFIG.AUDIT.SKU_AGG_LIMIT), hdr.length);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, hdr.length).setValues(rows);
    applyAuditFormats_(sh, 2, rows.length, [
      { col: 2, pattern: '0', context: 'audit.sku.qty' },
      { col: 3, pattern: '#,##0.00', context: 'audit.sku.principal' },
      { col: 4, pattern: '#,##0.00', context: 'audit.sku.tax' },
      { col: 5, pattern: '#,##0.00', context: 'audit.sku.fees_expense' },
      { col: 6, pattern: '#,##0.00', context: 'audit.sku.last_cost' },
      { col: 7, pattern: '#,##0.00', context: 'audit.sku.cogs' },
      { col: 8, pattern: '#,##0.00', context: 'audit.sku.profit' }
    ], warnings || []);
  }

  sh.autoResizeColumns(1, hdr.length);
}

function writeAuditRawLinesTab_(ss, audit, warnings) {
  const sh = ss.insertSheet('RAW_LINES');
  const header = audit.tsvHeader || [];
  if (!header.length) return;

  sh.getRange(1, 1, 1, header.length).setValues([header]);
  const rows = normalize2D_(audit.rawRows.slice(0, CONFIG.AUDIT.RAW_LINES_LIMIT), header.length);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);
    const idx = indexMapFlexible_(header);
    const fmts = [];
    if (idx['amount'] !== undefined) fmts.push({ col: idx['amount'] + 1, pattern: '#,##0.00', context: 'audit.raw.amount' });
    if (idx['quantity-purchased'] !== undefined) fmts.push({ col: idx['quantity-purchased'] + 1, pattern: '0', context: 'audit.raw.qty' });
    if (fmts.length) applyAuditFormats_(sh, 2, rows.length, fmts, warnings || []);
  }
  sh.autoResizeColumns(1, header.length);
}

function writeAuditMetaToSummary_(sh, hm, row, url, status) {
  const colUrl = hm[CONFIG.HEADERS.auditUrl];
  const colStatus = hm[CONFIG.HEADERS.auditStatus];

  if (colUrl) {
    if (url) {
      sh.getRange(row, colUrl).setFormula('=HYPERLINK("' + escapeForFormula_(url) + '","Open Audit")');
    } else {
      sh.getRange(row, colUrl).clearContent();
    }
  }

  if (colStatus) sh.getRange(row, colStatus).setValue(status || '');
}

/* =========================
 * CONTEXT / HEADER / FORMAT
 * ========================= */

function buildImportContext_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONFIG.SUMMARY_SHEET);
  if (!sh) throw new Error('Не знайдено вкладку "' + CONFIG.SUMMARY_SHEET + '"');

  ensureSummaryHeaders_(sh);
  const hm = getHeaderMap_(sh);
  const fileIdRowMap = buildFileIdRowIndexMap_(sh, hm[CONFIG.HEADERS.fileId]);
  const costMap = getLastUnitCostMap_();

  return { ss: ss, sh: sh, headerMap: hm, fileIdRowMap: fileIdRowMap, costMap: costMap };
}

function ensureSummaryHeaders_(sheet) {
  const lastCol = sheet.getLastColumn();
  const existing = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(v) { return String(v || '').trim(); }) : [];

  const set = new Set(existing.filter(function(x) { return !!x; }));
  const needed = Object.keys(CONFIG.HEADERS).map(function(k) { return CONFIG.HEADERS[k]; });
  const toAdd = needed.filter(function(h) { return !set.has(h); });

  if (toAdd.length) sheet.getRange(1, lastCol + 1, 1, toAdd.length).setValues([toAdd]);

  const hm = getHeaderMap_(sheet);
  applySummaryFormats_(sheet, hm, 2, Math.max(0, sheet.getLastRow() - 1), []);
}

function formatSummaryRow_(sh, hm, row, warnings) {
  warnings = warnings || [];
  applySummaryFormats_(sh, hm, row, 1, warnings);
}

function findOrCreateRowByFileId_(sh, hm, fileIdRowMap, fileId) {
  const key = String(fileId || '').trim();
  let row = fileIdRowMap.get(key);
  if (row) return row;

  row = findFirstEmptyRow_(sh, hm[CONFIG.HEADERS.fileId], 2);
  fileIdRowMap.set(key, row);
  return row;
}

function buildFileIdRowIndexMap_(sheet, colFileId) {
  const map = new Map();
  if (!colFileId) return map;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;

  const ids = sheet.getRange(2, colFileId, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    const fid = String(ids[i][0] || '').trim();
    if (!fid || fid === CONFIG.TOTAL_FILE_ID) continue;
    map.set(fid, 2 + i);
  }

  return map;
}

/* =========================
 * COSTS / COGS
 * ========================= */

function getLastUnitCostMap_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.PURCHASES_SHEET);
  if (!sh) throw new Error('Не знайдено вкладку "' + CONFIG.PURCHASES_SHEET + '"');

  const hm = getHeaderMap_(sh);
  const colSku = hm[CONFIG.PURCHASES.skuHeader];
  const colCost = hm[CONFIG.PURCHASES.unitCostHeader];
  if (!colSku || !colCost) throw new Error('У вкладці Закупки потрібні колонки SKU і Unit Cost');

  const lastRow = sh.getLastRow();
  const map = new Map();
  if (lastRow < 2) return map;

  const minCol = Math.min(colSku, colCost);
  const maxCol = Math.max(colSku, colCost);
  const vals = sh.getRange(2, minCol, lastRow - 1, maxCol - minCol + 1).getValues();

  const iSku = colSku - minCol;
  const iCost = colCost - minCol;

  for (let i = vals.length - 1; i >= 0; i--) {
    const sku = normalizeSku_(vals[i][iSku]);
    if (!sku || map.has(sku)) continue;

    const costC = toCents_(parseNumberLoose_(vals[i][iCost]));
    if (costC > 0) map.set(sku, costC);
  }

  return map;
}

function calcCogsFromCostMap_(skuQtyMap, costMap) {
  const skus = Object.keys(skuQtyMap || {});
  if (!skus.length) {
    return {
      cogsC: 0,
      unitsWithCost: 0,
      missingUnits: 0,
      coveragePct: 1,
      missingSkus: [],
      missingSkusText: '',
      costBySku: {}
    };
  }

  let cogsC = 0;
  let unitsTotal = 0;
  let unitsWithCost = 0;
  const missingSkus = [];
  const missingPairs = [];
  const costBySku = {};

  for (let i = 0; i < skus.length; i++) {
    const sku = skus[i];
    const qty = Number(skuQtyMap[sku] || 0);
    if (qty <= 0) continue;

    unitsTotal += qty;

    const costC = costMap.get(sku);
    if (costC === undefined) {
      missingSkus.push(sku);
      missingPairs.push(sku + '(' + qty + ')');
    } else {
      unitsWithCost += qty;
      cogsC += costC * qty;
      costBySku[sku] = fromCents_(costC);
    }
  }

  const missingUnits = Math.max(0, unitsTotal - unitsWithCost);
  const coveragePct = unitsTotal > 0 ? (unitsWithCost / unitsTotal) : 1;
  let missingSkusText = missingPairs.join(', ');
  if (missingSkusText.length > 400) missingSkusText = missingSkusText.slice(0, 380) + '…';

  return {
    cogsC: cogsC,
    unitsWithCost: unitsWithCost,
    missingUnits: missingUnits,
    coveragePct: coveragePct,
    missingSkus: missingSkus,
    missingSkusText: missingSkusText,
    costBySku: costBySku
  };
}

function normalizeFeesExpenseC_(sumSignedC, negCount, posCount) {
  if (sumSignedC === 0) return { feesExpenseC: 0, feesRule: 'expense = 0' };

  let expenseC;
  let rule;

  if (negCount > posCount) {
    expenseC = -sumSignedC;
    rule = 'expense = -SUM(signed) [neg-dominant]';
  } else if (posCount > negCount) {
    expenseC = sumSignedC;
    rule = 'expense = +SUM(signed) [pos-dominant]';
  } else {
    expenseC = Math.abs(sumSignedC);
    rule = 'expense = ABS(SUM(signed)) [tie-safe]';
  }

  if (expenseC < 0) {
    expenseC = Math.abs(expenseC);
    rule += ' + auto-flip';
  }

  return { feesExpenseC: expenseC, feesRule: rule };
}

/* =========================
 * REIMBURSEMENT CLASSIFIER
 * ========================= */

function isReimbursementLine_(transactionType, amountType, amountDesc) {
  const tx = String(transactionType || '').toLowerCase();
  const t = String(amountType || '').toLowerCase();
  const d = String(amountDesc || '').toLowerCase();

  if (containsAnyKeyword_(t, CONFIG.REIMBURSEMENTS.excludedAmountTypes)) return false;

  const txSig = containsAnyKeyword_(tx, CONFIG.REIMBURSEMENTS.transactionTypeKeywords);
  const typeSig = containsAnyKeyword_(t, CONFIG.REIMBURSEMENTS.amountTypeKeywords);
  const descSig = containsAnyKeyword_(d, CONFIG.REIMBURSEMENTS.amountDescriptionKeywords);

  if (txSig && (typeSig || descSig)) return true;
  if (descSig && typeSig) return true;
  if (tx.indexOf('other-transaction') !== -1 && descSig) return true;
  return false;
}

function containsAnyKeyword_(text, keywords) {
  const s = String(text || '').toLowerCase();
  if (!s) return false;

  for (let i = 0; i < (keywords || []).length; i++) {
    const k = String(keywords[i] || '').toLowerCase().trim();
    if (!k) continue;
    if (s.indexOf(k) !== -1) return true;
  }

  return false;
}

/* =========================
 * AUDIT ROW BUILDERS
 * ========================= */

function buildOrderRows_(orderAgg, costBySku) {
  const keys = Object.keys(orderAgg || {});
  const out = [];

  for (let i = 0; i < keys.length; i++) {
    const r = orderAgg[keys[i]];
    const feeNorm = normalizeFeesExpenseC_(toCents_(fromCents_(r.feesSignedC)), r.feesSignedC < 0 ? 1 : 0, r.feesSignedC > 0 ? 1 : 0);
    const lastUnitCost = costBySku[r.sku];
    const cogs = (lastUnitCost === undefined) ? '' : Number(lastUnitCost) * Number(r.qty || 0);
    const profit = fromCents_(r.principalC + r.taxC - feeNorm.feesExpenseC) - (cogs === '' ? 0 : cogs);

    out.push([
      r.orderId,
      r.sku,
      Number(r.qty || 0),
      fromCents_(r.principalC),
      fromCents_(r.taxC),
      fromCents_(r.feesSignedC),
      fromCents_(feeNorm.feesExpenseC),
      lastUnitCost === undefined ? '' : Number(lastUnitCost),
      cogs === '' ? '' : cogs,
      profit
    ]);
  }

  out.sort(function(a, b) { return Number(b[2] || 0) - Number(a[2] || 0); });
  return out;
}

function buildSkuRows_(skuQtyMap, orderAgg, costBySku, missingSkus) {
  const skus = Object.keys(skuQtyMap || {});
  const missingSet = new Set(missingSkus || []);
  const out = [];
  const bySku = Object.create(null);

  const orderKeys = Object.keys(orderAgg || {});
  for (let i = 0; i < orderKeys.length; i++) {
    const r = orderAgg[orderKeys[i]];
    if (!r || !r.sku) continue;
    if (!bySku[r.sku]) bySku[r.sku] = { principalC: 0, taxC: 0, feesExpenseC: 0 };
    const feeNorm = normalizeFeesExpenseC_(toCents_(fromCents_(r.feesSignedC)), r.feesSignedC < 0 ? 1 : 0, r.feesSignedC > 0 ? 1 : 0);
    bySku[r.sku].principalC += Number(r.principalC || 0);
    bySku[r.sku].taxC += Number(r.taxC || 0);
    bySku[r.sku].feesExpenseC += Number(feeNorm.feesExpenseC || 0);
  }

  for (let i = 0; i < skus.length; i++) {
    const sku = skus[i];
    const qty = Number(skuQtyMap[sku] || 0);
    const sale = bySku[sku] || { principalC: 0, taxC: 0, feesExpenseC: 0 };
    const cost = costBySku[sku];
    const cogs = (cost === undefined) ? '' : (cost * qty);
    const profit = fromCents_(sale.principalC + sale.taxC - sale.feesExpenseC) - (cogs === '' ? 0 : cogs);

    out.push([
      sku,
      qty,
      fromCents_(sale.principalC),
      fromCents_(sale.taxC),
      fromCents_(sale.feesExpenseC),
      cost === undefined ? '' : cost,
      cogs === '' ? '' : cogs,
      profit,
      missingSet.has(sku) ? 'NO' : 'YES'
    ]);
  }

  out.sort(function(a, b) { return Number(b[1] || 0) - Number(a[1] || 0); });
  return out;
}

/* =========================
 * HELPERS
 * ========================= */

function getHeaderMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) throw new Error('Sheet has no columns: ' + sheet.getName());

  const vals = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  const seen = new Set();

  for (let i = 0; i < vals.length; i++) {
    const key = String(vals[i] || '').trim();
    if (!key) continue;
    if (seen.has(key)) throw new Error('Duplicate header "' + key + '" in ' + sheet.getName());
    seen.add(key);
    map[key] = i + 1;
  }

  return map;
}

function buildRowFromHeaderMap_(headerMap, outObj) {
  const maxCol = Math.max.apply(null, Object.keys(headerMap).map(function(k) { return headerMap[k]; }));
  const row = Array(maxCol).fill('');

  Object.keys(outObj).forEach(function(h) {
    const col = headerMap[h];
    if (col) row[col - 1] = outObj[h];
  });

  return row;
}

function indexMap_(headers) {
  const map = {};
  headers.forEach(function(h, i) { map[String(h || '').trim()] = i; });

  ['settlement-id', 'deposit-date', 'total-amount', 'amount-type', 'amount-description', 'amount'].forEach(function(k) {
    if (map[k] === undefined) throw new Error('Missing column in settlement header: ' + k);
  });

  return map;
}


function splitLines_(content) {
  return String(content || '').replace(/^\uFEFF/, '').split(/\r?\n/);
}

function detectSep_(headerLine) {
  const line = String(headerLine || '');
  const tab = (line.match(/	/g) || []).length;
  const semicolon = (line.match(/;/g) || []).length;
  const comma = (line.match(/,/g) || []).length;
  if (tab >= semicolon && tab >= comma) return '	';
  if (semicolon >= comma) return ';';
  return ',';
}

function splitTsvLine_(line, sep) {
  return String(line || '').split(sep || '	');
}

function findHeaderIdx_(normalizedHeaders, variants) {
  const target = variants.map(function(v) { return normalizeHeader_(v); });
  for (let i = 0; i < normalizedHeaders.length; i++) {
    if (target.indexOf(normalizedHeaders[i]) !== -1) return i;
  }
  return -1;
}

function parseSmartNumber_(value) {
  return parseNumberLoose_(value);
}

function parseDateMaybe_(value) {
  return parseDateFlexible_(value, CONFIG.TZ);
}

function parseMonthFromFilename_(fileName) {
  const raw = String(fileName || '');
  let m = raw.match(/(20\d{2})[-_.](0[1-9]|1[0-2])/);
  if (m) return new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, 1));

  m = raw.match(/(0[1-9]|1[0-2])[-_.](20\d{2})/);
  if (m) return new Date(Date.UTC(Number(m[2]), Number(m[1]) - 1, 1));

  return null;
}

function safeSplitTsv_(line, expectedLen) {
  const parts = String(line || '').split('\t');
  if (!expectedLen) return parts;

  const out = parts.slice(0, expectedLen);
  while (out.length < expectedLen) out.push('');
  return out;
}

function cellByHeader_(row, idx, key) {
  const i = idx[key];
  if (i === undefined) return '';
  return String(row[i] || '').trim();
}

function normalizeSku_(v) {
  const s = String(v || '').trim();
  return s ? s.toUpperCase() : '';
}

function parseNumberLoose_(s) {
  if (s === null || s === undefined || s === '') return 0;
  let t = String(s).trim();
  if (!t) return 0;

  t = t.replace(/\s+/g, '');
  t = t.replace(/[€$£₴]/g, '');

  const hasComma = t.indexOf(',') !== -1;
  const hasDot = t.indexOf('.') !== -1;

  if (hasComma && hasDot) {
    if (t.lastIndexOf(',') > t.lastIndexOf('.')) {
      t = t.replace(/\./g, '').replace(',', '.');
    } else {
      t = t.replace(/,/g, '');
    }
  } else if (hasComma) {
    t = t.replace(',', '.');
  }

  const n = Number(t);
  return isNaN(n) ? 0 : n;
}

function parseDateFlexible_(value, tz) {
  const raw = String(value || '').trim();
  if (!raw) return null;

  const direct = new Date(raw);
  if (!isNaN(direct.getTime())) return direct;

  let m = raw.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?(?:\s*(Z|UTC|[+-]\d{2}:?\d{2})?)?$/i);
  if (m) {
    const y=Number(m[1]), mo=Number(m[2])-1, d=Number(m[3]), hh=Number(m[4]||0), mm=Number(m[5]||0), ss=Number(m[6]||0);
    return new Date(Date.UTC(y, mo, d, hh, mm, ss));
  }

  m = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    const a=Number(m[1]), b=Number(m[2]), y=Number(m[3]), hh=Number(m[4]||0), mm=Number(m[5]||0), ss=Number(m[6]||0);
    const dayFirst = a > 12 || (a <= 12 && b <= 12);
    const d = dayFirst ? a : b;
    const mo = (dayFirst ? b : a) - 1;
    return new Date(Date.UTC(y, mo, d, hh, mm, ss));
  }

  m = raw.match(/^(\d{2})\.(\d{2})\.(\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?(?:\s*UTC)?$/i);
  if (m) return new Date(Date.UTC(Number(m[3]), Number(m[2]) - 1, Number(m[1]), Number(m[4]||0), Number(m[5]||0), Number(m[6]||0)));

  return null;
}

function parseTsv_(text, fileMeta) {
  const raw = String(text || '').replace(/^\uFEFF/, '');
  const lines = raw.split(/\r?\n/);

  if (fileMeta && Number(fileMeta.size || 0) === 0) {
    throw buildFileDiagnosticError_('Файл порожній (size=0).', fileMeta, raw, [], []);
  }

  if (raw.indexOf('	') === -1) {
    throw buildFileDiagnosticError_('Файл не TSV (no tab separators).', fileMeta, raw, [], []);
  }

  const headerRowIndex = findHeaderRowIndex_(lines);
  if (headerRowIndex < 0) {
    throw buildFileDiagnosticError_('Не знайдено header row з required headers.', fileMeta, raw, [], []);
  }

  const headers = safeSplitTsv_(lines[headerRowIndex]);
  const idx = indexMapFlexible_(headers, true);
  const essential = ['settlement-id', 'deposit-date', 'amount-type', 'amount-description', 'transaction-type', 'amount'];
  const missing = essential.filter(function(k) { return idx[k] === undefined; });
  if (missing.length) {
    throw buildFileDiagnosticError_('Не знайдені required headers: ' + missing.join(', '), fileMeta, raw, headers, []);
  }

  const rows = [];
  for (let i = headerRowIndex + 1; i < lines.length; i++) {
    const line = lines[i];
    if (!String(line || '').trim()) continue;
    const row = safeSplitTsv_(line, headers.length);
    const nonEmpty = row.some(function(v) { return String(v || '').trim() !== ''; });
    if (nonEmpty) rows.push(row);
  }

  if (!rows.length) {
    throw buildFileDiagnosticError_('TSV містить header, але не містить жодного data row.', fileMeta, raw, headers, []);
  }

  return { headers: headers, rows: rows, headerRowIndex: headerRowIndex };
}

function findHeaderRowIndex_(lines) {
  let best = { idx: -1, score: -1 };
  const maxScan = Math.min(lines.length, 40);
  for (let i = 0; i < maxScan; i++) {
    const line = String(lines[i] || '');
    if (!line.trim() || line.indexOf('	') === -1) continue;

    const headers = safeSplitTsv_(line);
    const idx = indexMapFlexible_(headers, true);
    const scoreKeys = ['settlement-id', 'deposit-date', 'amount-type', 'amount-description', 'transaction-type', 'amount', 'sku', 'quantity-purchased'];
    let score = 0;
    for (let k = 0; k < scoreKeys.length; k++) if (idx[scoreKeys[k]] !== undefined) score++;

    if (score > best.score) best = { idx: i, score: score };
  }
  return best.score >= 4 ? best.idx : -1;
}

function normalizeHeader_(h) {
  return String(h || '')
    .trim()
    .toLowerCase()
    .replace(/[_\s]+/g, '-')
    .replace(/[^a-z0-9-]/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '');
}

function indexMapFlexible_(headers, skipRequiredCheck) {
  const aliases = {
    'settlement-id': ['settlement-id', 'settlementid'],
    'deposit-date': ['deposit-date', 'deposit-date-time', 'depositdate', 'depositdatetime'],
    'total-amount': ['total-amount', 'totalamount', 'total amount'],
    'amount-type': ['amount-type', 'amounttype'],
    'amount-description': ['amount-description', 'amountdescription'],
    'transaction-type': ['transaction-type', 'transactiontype'],
    'amount': ['amount'],
    'sku': ['sku'],
    'quantity-purchased': ['quantity-purchased', 'quantitypurchased', 'quantity-purchase', 'quantity'],
    'order-id': ['order-id', 'orderid'],
    'marketplace-name': ['marketplace-name', 'marketplacename'],
    'principal': ['principal'],
    'tax': ['tax'],
    'shipping-tax': ['shipping-tax', 'shippingtax'],
    'gift-wrap': ['gift-wrap', 'giftwrap'],
    'gift-wrap-tax': ['gift-wrap-tax', 'giftwraptax'],
    'item-fees': ['item-fees', 'itemfees']
  };

  const normHeaders = headers.map(function(h) { return normalizeHeader_(h); });
  const map = {};

  Object.keys(aliases).forEach(function(key) {
    const variants = aliases[key].map(function(v) { return normalizeHeader_(v); });
    for (let i = 0; i < normHeaders.length; i++) {
      if (variants.indexOf(normHeaders[i]) !== -1) {
        map[key] = i;
        break;
      }
    }
  });

  if (!skipRequiredCheck) {
    ['settlement-id', 'deposit-date', 'amount-type', 'amount-description', 'amount'].forEach(function(k) {
      if (map[k] === undefined) throw new Error('Missing column in settlement header: ' + k);
    });
  }

  return map;
}

function detectTransferC_(rows, idx, warnings) {
  warnings = warnings || [];
  if (idx['total-amount'] !== undefined) {
    for (let i = 0; i < rows.length; i++) {
      const v = rows[i][idx['total-amount']];
      if (String(v || '').trim() !== '') return toCents_(parseNumberLoose_(v));
    }
  }

  const iAmount = idx['amount'];
  const iType = idx['amount-type'];
  const iDesc = idx['amount-description'];
  if (iAmount !== undefined) {
    for (let i = 0; i < rows.length; i++) {
      const t = String((iType !== undefined ? rows[i][iType] : '') || '').toLowerCase();
      const d = String((iDesc !== undefined ? rows[i][iDesc] : '') || '').toLowerCase();
      if (t.indexOf('total') !== -1 || d.indexOf('total') !== -1) {
        const c = toCents_(parseNumberLoose_(rows[i][iAmount]));
        warnings.push('[WARN] transfer determined from amount row containing "total".');
        return c;
      }
    }
  }

  throw new Error('Не вдалося визначити Transfer: відсутній total-amount і не знайдено альтернативний total рядок.');
}

function buildFileDiagnosticError_(reason, fileMeta, text, headers, extraLines) {
  const meta = fileMeta || {};
  const preview = String(text || '').split(/\r?\n/).slice(0, 3).map(function(l) {
    return String(l || '').slice(0, 280);
  });
  const hdrs = (headers || []).slice(0, 30);

  const lines = [
    reason,
    'fileId: ' + String(meta.id || ''),
    'fileName: ' + String(meta.name || ''),
    'fileSize: ' + String(meta.size || 0),
    'lastUpdated: ' + (meta.updatedAt ? Utilities.formatDate(meta.updatedAt, CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss') : ''),
    'mimeType: ' + String(meta.mimeType || '')
  ];

  if (preview.length) {
    lines.push('preview:');
    for (let i = 0; i < preview.length; i++) lines.push('  ' + (i + 1) + ') ' + preview[i]);
  }

  if (hdrs.length) lines.push('headers(found up to 30): ' + hdrs.join(', '));
  if (extraLines && extraLines.length) Array.prototype.push.apply(lines, extraLines);

  return new Error(lines.join('\n'));
}

function toCents_(x) {
  return Math.round((Number(x) || 0) * 100);
}

function fromCents_(c) {
  return (Number(c) || 0) / 100;
}

function assertValidDate_(d, label) {
  if (!(d instanceof Date) || isNaN(d.getTime())) throw new Error('Invalid Date for ' + label + ': ' + d);
}

function mapMarketplace_(marketplaceName) {
  const s = String(marketplaceName || '').toLowerCase();
  if (s.indexOf('amazon.it') !== -1 || s.indexOf('italy') !== -1) return 'Amazon IT';
  if (s.indexOf('amazon.de') !== -1 || s.indexOf('germany') !== -1) return 'Amazon DE';
  if (s.indexOf('amazon.fr') !== -1 || s.indexOf('france') !== -1) return 'Amazon FR';
  if (s.indexOf('amazon.es') !== -1 || s.indexOf('spain') !== -1) return 'Amazon ES';
  if (s.indexOf('amazon.nl') !== -1 || s.indexOf('netherlands') !== -1) return 'Amazon NL';
  if (s.indexOf('amazon.be') !== -1 || s.indexOf('belgium') !== -1) return 'Amazon BE';
  return marketplaceName || 'Amazon (невідомо)';
}

function findFirstEmptyRow_(sheet, colIndex, startRow) {
  const from = startRow || 2;
  const last = Math.max(sheet.getLastRow(), from);
  const vals = sheet.getRange(from, colIndex, last - from + 1, 1).getValues();

  for (let i = 0; i < vals.length; i++) {
    if (!String(vals[i][0] || '').trim()) return from + i;
  }

  return last + 1;
}

function normalize2D_(rows, cols) {
  const out = [];
  const n = rows ? rows.length : 0;
  for (let i = 0; i < n; i++) {
    const src = Array.isArray(rows[i]) ? rows[i] : [rows[i]];
    const row = src.slice(0, cols);
    while (row.length < cols) row.push('');
    out.push(row);
  }
  return out;
}

function colToA1_(colIndex) {
  let n = colIndex;
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function getFormulaArgSeparator_() {
  const locale = SpreadsheetApp.getActive().getSpreadsheetLocale() || '';
  return /^en_/i.test(locale) ? ',' : ';';
}

function escapeForFormula_(s) {
  return String(s || '').replace(/"/g, '""');
}

function safeSetNumberFormat_(range, pattern, warnings, context) {
  const warnList = Array.isArray(warnings) ? warnings : [];
  try {
    if (!range || !pattern) return true;
    const vals = range.getValues();
    const hasValue = vals.some(function(row) {
      return row.some(function(v) { return v !== '' && v !== null; });
    });
    if (!hasValue) return true;

    range.setNumberFormat(pattern);
    return true;
  } catch (e) {
    const msg = '[FORMAT WARN] ' + (context || 'unknown') + ': ' + toErrorMessage_(e);
    Logger.log(msg);
    warnList.push(msg);

    try {
      const vals = range.getValues();
      const r0 = range.getRow();
      const c0 = range.getColumn();
      const isDateFmt = /[dmyhHsS]/i.test(pattern);

      for (let r = 0; r < vals.length; r++) {
        for (let c = 0; c < vals[r].length; c++) {
          const v = vals[r][c];
          if (v === '' || v === null) continue;
          if (isDateFmt ? (v instanceof Date && !isNaN(v.getTime())) : (typeof v === 'number' && isFinite(v))) {
            try {
              range.getSheet().getRange(r0 + r, c0 + c, 1, 1).setNumberFormat(pattern);
            } catch (cellErr) {
              const cmsg = '[FORMAT WARN CELL] ' + (context || 'unknown') + ' R' + (r0 + r) + 'C' + (c0 + c) + ': ' + toErrorMessage_(cellErr);
              Logger.log(cmsg);
              warnList.push(cmsg);
            }
          }
        }
      }
    } catch (fallbackErr) {
      const fmsg = '[FORMAT WARN FALLBACK] ' + (context || 'unknown') + ': ' + toErrorMessage_(fallbackErr);
      Logger.log(fmsg);
      warnList.push(fmsg);
    }

    return false;
  }
}

function applySummaryFormats_(sheet, hm, startRow, rowCount, warnings) {
  warnings = warnings || [];
  if (!sheet || !hm || !startRow || rowCount <= 0) return;

  const dateCols = [hm[CONFIG.HEADERS.depositDate], hm[CONFIG.HEADERS.month], hm[CONFIG.HEADERS.importedAt]].filter(Boolean);
  for (let i = 0; i < dateCols.length; i++) {
    const col = dateCols[i];
    const fmt = col === hm[CONFIG.HEADERS.month] ? 'yyyy-MM' : (col === hm[CONFIG.HEADERS.importedAt] ? 'dd.mm.yyyy HH:mm:ss' : 'dd.mm.yyyy');
    safeSetNumberFormat_(sheet.getRange(startRow, col, rowCount, 1), fmt, warnings, 'summary.date.col' + col);
  }

  const moneyCols = [
    hm[CONFIG.HEADERS.salesNet], hm[CONFIG.HEADERS.vatDebito], hm[CONFIG.HEADERS.feesCost], hm[CONFIG.HEADERS.otherNet],
    hm[CONFIG.HEADERS.transfer], hm[CONFIG.HEADERS.payoutExReimbursements], hm[CONFIG.HEADERS.cogs], hm[CONFIG.HEADERS.netProfit], hm[CONFIG.HEADERS.amazonReimbursements],
    hm[CONFIG.HEADERS.soldProfit], hm[CONFIG.HEADERS.profitExReimbursements], hm[CONFIG.HEADERS.companyProfit]
  ].filter(Boolean);
  for (let i = 0; i < moneyCols.length; i++) {
    safeSetNumberFormat_(sheet.getRange(startRow, moneyCols[i], rowCount, 1), '#,##0.00', warnings, 'summary.money.col' + moneyCols[i]);
  }

  const intCols = [hm[CONFIG.HEADERS.units], hm[CONFIG.HEADERS.unitsWithCost], hm[CONFIG.HEADERS.missingUnits]].filter(Boolean);
  for (let i = 0; i < intCols.length; i++) {
    safeSetNumberFormat_(sheet.getRange(startRow, intCols[i], rowCount, 1), '0', warnings, 'summary.int.col' + intCols[i]);
  }

  if (hm[CONFIG.HEADERS.cogsCoverage]) {
    safeSetNumberFormat_(sheet.getRange(startRow, hm[CONFIG.HEADERS.cogsCoverage], rowCount, 1), '0.00%', warnings, 'summary.coverage');
  }
}

function applyAuditFormats_(sheet, dataStartRow, rowCount, formats, warnings) {
  warnings = warnings || [];
  if (!sheet || rowCount <= 0 || !formats || !formats.length) return;
  for (let i = 0; i < formats.length; i++) {
    const f = formats[i];
    if (!f || !f.col || !f.pattern) continue;
    safeSetNumberFormat_(sheet.getRange(dataStartRow, f.col, rowCount, 1), f.pattern, warnings, (f.context || 'audit') + '.col' + f.col);
  }
}

function applyMonthlyFormats_(sheet, rowCount, warnings) {
  warnings = warnings || [];
  if (!sheet || rowCount <= 0) return;
  safeSetNumberFormat_(sheet.getRange(2, 1, rowCount, 1), 'yyyy-MM', warnings, 'monthly.month');
  safeSetNumberFormat_(sheet.getRange(2, 2, rowCount, 7), '#,##0.00', warnings, 'monthly.money.left');
  safeSetNumberFormat_(sheet.getRange(2, 9, rowCount, 1), '0', warnings, 'monthly.units');
  safeSetNumberFormat_(sheet.getRange(2, 10, rowCount, 5), '#,##0.00', warnings, 'monthly.money.right');
}

function buildUiResultMessage_(title, successDetails, warnings, errors) {
  const parts = [String(title || '')];
  if (successDetails) parts.push('', 'SUCCESS:', String(successDetails));

  const ws = (warnings || []).filter(function(w) { return !!String(w || '').trim(); });
  if (ws.length) {
    parts.push('', 'WARNINGS (' + ws.length + '):');
    for (let i = 0; i < ws.length; i++) parts.push('- ' + ws[i]);
  }

  const es = (errors || []).filter(function(e) { return !!String(e || '').trim(); });
  if (es.length) {
    parts.push('', 'ERRORS (' + es.length + '):');
    for (let i = 0; i < es.length; i++) parts.push('- ' + es[i]);
  }

  return parts.join('\n');
}


function runNonCritical_(label, fn, warnings) {
  const warnList = Array.isArray(warnings) ? warnings : [];
  try {
    return fn();
  } catch (e) {
    const msg = '[NON-CRITICAL WARN] ' + String(label || 'unknown') + ': ' + toErrorMessage_(e);
    Logger.log(msg);
    warnList.push(msg);
    return null;
  }
}

function safeToast_(msg) {
  try {
    SpreadsheetApp.getActive().toast(String(msg || ''), 'Amazon Finance', 8);
  } catch (e) {
    Logger.log('[TOAST WARN] ' + toErrorMessage_(e));
  }
}

function toErrorMessage_(e) {
  if (!e) return 'Unknown error';
  return e && e.message ? e.message : String(e);
}

function handleFatal_(where, e) {
  const msg = '[' + where + '] ' + toErrorMessage_(e);
  Logger.log(msg + '\n' + (e && e.stack ? e.stack : ''));
  safeToast_(msg);
}
