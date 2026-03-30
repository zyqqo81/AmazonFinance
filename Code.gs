/***************
 * AMAZON SETTLEMENT IMPORTER (Posted Date month basis) — BALANCED + COGS + AUDIT PACK
 * Google Sheets + Apps Script
 ***************/

const CONFIG = {
  SUMMARY_SHEET: 'ІМПОРТ – ЗВЕДЕННЯ',
  MONTHLY_SHEET: 'МІСЯЧНИЙ_ЗВІТ',
  DASHBOARD_SHEET: 'ДЕШБОРД',
  PURCHASES_SHEET: 'Закупки',
  MANUAL_EXPENSES_SHEET: 'РУЧНІ_ВИТРАТИ',
  MANUAL_OPERATIONS_SHEET: 'РУЧНІ_ОПЕРАЦІЇ',
  WORKING_CAPITAL_SHEET: 'ОБОРОТНИЙ_КАПІТАЛ',
  LEGACY_BUSINESS_EXPENSES_SHEET: 'ВИТРАТИ_БІЗНЕСУ',
  TZ: 'Europe/Rome',
  TOTAL_FILE_ID: '__TOTAL__',
  SETTLEMENT_FOLDER_ID: '1K9AuTAmNr5AXHmlTuOIlUdYDRKlKrOAj',
  FOLDER_LIST_LIMIT: 15,
  BULK_IMPORT_CONFIRM_LIMIT: 200,
  MONTHLY_REPORT_FOLDER_ID: '1k4fDrE_XYoZ0ukOByEz9A-053dIKsXSo',
  MONTHLY_REPORT_SHEET: 'ПДВ ЗВІТ',

  TAX_REPORT_FOLDER_ID: '1k4fDrE_XYoZ0ukOByEz9A-053dIKsXSo',
  TAX_RAW_SHEET: 'TAX_REPORT_RAW',
  VAT_SUMMARY_SHEET: 'VAT_SALES_SUMMARY',
  VAT_PIVOT_SHEET: 'VAT_PIVOT',
  CURRENT_MONTH_SNAPSHOT_SHEET: 'VAT_CURRENT_MONTH',
  DEFAULT_GROUP_DATE_FIELD: 'Tax Calculation Date',
  ALT_GROUP_DATE_FIELD: 'Order Date',
  SHIPMENT_DATE_FIELD: 'Shipment Date',
  CURRENCY: 'EUR',

  SALES_TAX_REPORT_FOLDER_ID: '1k4fDrE_XYoZ0ukOByEz9A-053dIKsXSo',
  SALES_TAX_RAW_SHEET: 'SALES_TAX_RAW',
  MONTHLY_VAT_PAYOUT_SUMMARY_SHEET: 'МІСЯЧНИЙ_ЗВІТ',
  LEGACY_MONTHLY_SHEET: 'МІСЯЦІ',
  LEGACY_MONTHLY_NOTE: 'DEPRECATED: не використовується кодом. Джерело правди перенесено у МІСЯЧНИЙ_ЗВІТ.',
  DIAGNOSTICS_SHEET: 'ДІАГНОСТИКА',

  AUDIT: {
    ENABLED: true,
    FOLDER_ID: '1ALCVcKM_3QlEeCedr6DE1YNOI5HKzE2s',
    RAW_LINES_LIMIT: 50000,
    ORDER_ITEMS_LIMIT: 100000,
    SKU_AGG_LIMIT: 100000
  },

  HEADERS: {
    depositDate: 'Deposit Date',
    postedDate: 'Posted Date',
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
    unitCostHeader: 'Unit Cost',
    sourceHeader: 'Джерело',
    sourceIdHeader: 'ID ручної операції',
    syncUpdatedAtHeader: 'Оновлено синхронізацією'
  },

  MANUAL_OPERATION_TYPES: {
    BUSINESS_EXPENSE: 'Бізнес-витрата',
    PURCHASE: 'Закупка товару'
  },

  MANUAL_EXPENSE_FUND_CATEGORIES: [
    'Реінвест (75%)',
    'Бізнес витрати (12%)',
    'Зарплата (7%)',
    'Інше (6%)'
  ],

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

function onOpenLegacyMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('Фінанси Amazon')
    .addItem('Завантажити всі settlement файли', 'uiImportAllFromFolder_')
    .addItem('Завантажити всі sales звіти', 'menuImportAllSalesTaxReports_')
    .addItem('Перерахувати фінансовий звіт', 'menuRebuildMonthlyVatPayoutSummary_')
    .addItem('Перерахувати лише останній місяць', 'menuRebuildLatestMonthOnly_')
    .addItem('Показати діагностику', 'menuRunVatDiagnostics_')
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

function uiImportAllMonthlyVatReportsFromFolder_() {
  const ui = SpreadsheetApp.getUi();
  try {
    if (!CONFIG.MONTHLY_REPORT_FOLDER_ID) {
      ui.alert('Не задано CONFIG.MONTHLY_REPORT_FOLDER_ID');
      return;
    }

    const candidates = getMonthlyReportFileCandidatesFromFolder_(CONFIG.MONTHLY_REPORT_FOLDER_ID, Number.MAX_SAFE_INTEGER);
    if (!candidates.length) {
      ui.alert('У папці не знайдено TXT/TSV/CSV файлів: ' + CONFIG.MONTHLY_REPORT_FOLDER_ID);
      return;
    }

    const confirm = ui.alert(
      'Підтвердження імпорту місячних звітів',
      ['Знайдено файлів: ' + candidates.length, 'Перерахувати всі місячні звіти?'].join('\n'),
      ui.ButtonSet.OK_CANCEL
    );
    if (confirm !== ui.Button.OK) return;

    const result = importAllMonthlyVatReportsFromFolder_({ candidates: candidates, resetSheet: true, progressEvery: 10 });
    const lines = [
      'Знайдено файлів: ' + result.total,
      'Успішно оброблено: ' + result.imported,
      'Оновлено місяців: ' + result.monthsTouched
    ];
    if (result.failed > 0) {
      lines.push('Помилки: ' + result.failed);
      const shown = result.errors.slice(0, 10);
      for (let i = 0; i < shown.length; i++) lines.push('- ' + shown[i]);
      if (result.errors.length > shown.length) lines.push('... ще ' + (result.errors.length - shown.length) + ' помилок.');
    }

    ui.alert('Імпорт усіх місячних звітів завершено', lines.join('\n'), ui.ButtonSet.OK);
  } catch (e) {
    handleFatal_('uiImportAllMonthlyVatReportsFromFolder_', e);
    ui.alert('Помилка масового імпорту місячних звітів: ' + toErrorMessage_(e));
  }
}

function uiRecalculateAllSettlementsAndMonthlyReports_() {
  const ui = SpreadsheetApp.getUi();
  const warnings = [];
  try {
    validateSummarySheet_();

    const settlementCandidates = getSettlementFileCandidatesFromFolder_(CONFIG.SETTLEMENT_FOLDER_ID, Number.MAX_SAFE_INTEGER);
    const monthlyCandidates = getMonthlyReportFileCandidatesFromFolder_(CONFIG.MONTHLY_REPORT_FOLDER_ID, Number.MAX_SAFE_INTEGER);

    const confirm = ui.alert(
      'Повний перерахунок',
      [
        'Settlement файлів: ' + settlementCandidates.length,
        'Місячних звітів: ' + monthlyCandidates.length,
        'Запустити ПОВНИЙ перерахунок всіх даних?'
      ].join('\n'),
      ui.ButtonSet.OK_CANCEL
    );
    if (confirm !== ui.Button.OK) return;

    const settlements = importAllSettlementsFromFolder_({
      warnings: warnings,
      candidates: settlementCandidates,
      progressEvery: 10,
      rebuildEvery: 10,
      forceReimport: true
    });

    const monthly = importAllMonthlyVatReportsFromFolder_({
      candidates: monthlyCandidates,
      resetSheet: true,
      progressEvery: 10
    });

    const lines = [
      'Settlement: ' + settlements.imported + '/' + settlements.total,
      'Monthly reports: ' + monthly.imported + '/' + monthly.total,
      'Оновлено місяців у ПДВ звіті: ' + monthly.monthsTouched
    ];
    if (settlements.failed > 0 || monthly.failed > 0) {
      lines.push('Помилки settlement: ' + settlements.failed);
      lines.push('Помилки monthly: ' + monthly.failed);
    }

    ui.alert(buildUiResultMessage_('Повний перерахунок завершено.', lines.join('\n'), warnings));
  } catch (e) {
    handleFatal_('uiRecalculateAllSettlementsAndMonthlyReports_', e);
    ui.alert(buildUiResultMessage_('Повний перерахунок завершився з помилкою.', '', warnings, [toErrorMessage_(e)]));
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



function getMonthlyReportFileCandidatesFromFolder_(folderId, limit) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const out = [];

  while (files.hasNext()) {
    const f = files.next();
    const name = String(f.getName() || '');
    const lname = name.toLowerCase();
    const mime = String(f.getMimeType() || '').toLowerCase();

    const mimeOk = mime === 'text/plain' || mime === 'application/octet-stream' || mime === 'text/tab-separated-values' || mime === 'text/csv' || mime === 'application/vnd.ms-excel';
    const nameOk = /\.(txt|tsv|csv)$/i.test(name) || lname.indexOf('vat') !== -1 || lname.indexOf('report') !== -1;
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

function importAllMonthlyVatReportsFromFolder_(options) {
  options = options || {};
  const candidates = Array.isArray(options.candidates)
    ? options.candidates
    : getMonthlyReportFileCandidatesFromFolder_(CONFIG.MONTHLY_REPORT_FOLDER_ID, Number.MAX_SAFE_INTEGER);

  if (!candidates.length) return { total: 0, imported: 0, failed: 0, monthsTouched: 0, errors: [] };

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(CONFIG.MONTHLY_REPORT_SHEET) || ss.insertSheet(CONFIG.MONTHLY_REPORT_SHEET);
  if (options.resetSheet) {
    const headers = ['Month', 'Sales Total', 'VAT To Pay', 'Rows', 'Файл', 'File ID', 'Імпортовано'];
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const progressEvery = Math.max(1, Number(options.progressEvery) || 0);
  const errors = [];
  const months = Object.create(null);
  let imported = 0;

  for (let i = 0; i < candidates.length; i++) {
    const f = candidates[i];
    try {
      const res = importMonthlyVatReportFile_(f.id);
      months[res.monthLabel] = true;
      imported++;
    } catch (e) {
      const emsg = '[' + f.name + '] ' + toErrorMessage_(e);
      errors.push(emsg);
      Logger.log('[MONTHLY VAT BULK ERROR] ' + emsg);
    }

    if (progressEvery > 0 && ((i + 1) % progressEvery === 0 || i === candidates.length - 1)) {
      safeToast_('Monthly VAT bulk update: ' + (i + 1) + '/' + candidates.length);
    }
  }

  if (sheet.getLastRow() > 1) {
    safeSetNumberFormat_(sheet.getRange(2, 1, sheet.getLastRow() - 1, 1), 'yyyy-MM', [], 'monthlyVat.month.bulk');
    safeSetNumberFormat_(sheet.getRange(2, 2, sheet.getLastRow() - 1, 2), '#,##0.00', [], 'monthlyVat.money.bulk');
  }

  return {
    total: candidates.length,
    imported: imported,
    failed: errors.length,
    monthsTouched: Object.keys(months).length,
    errors: errors
  };
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
        skipPostImportRebuild: true,
        forceReimport: !!options.forceReimport
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
  const importedAt = new Date();
  const rowIndexes = [];

  if (!options.auditOnly) {
    for (let b = 0; b < parsed.summaryRows.length; b++) {
      const rowData = parsed.summaryRows[b].rowData;
      rowData[CONFIG.HEADERS.fileName] = fileName;
      rowData[CONFIG.HEADERS.fileId] = fileId;
      rowData[CONFIG.HEADERS.importedAt] = importedAt;
      rowData[CONFIG.HEADERS.auditStatus] = 'START';
      rowData[CONFIG.HEADERS.auditUrl] = '';

      const rowValues = buildRowFromHeaderMap_(hm, rowData);
      const rowIndex = findOrCreateRowBySettlementMonth_(sh, hm, fileId, parsed.settlementId, parsed.summaryRows[b].monthKey);
      sh.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
      rowIndexes.push(rowIndex);
    }

    runNonCritical_('formatSummaryRows_', function() {
      for (let i = 0; i < rowIndexes.length; i++) formatSummaryRow_(sh, hm, rowIndexes[i], warnings);
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
  if (!options.auditOnly) {
    for (let i = 0; i < rowIndexes.length; i++) {
      writeAuditMetaToSummary_(sh, hm, rowIndexes[i], auditResult.url || '', auditStatus);
    }
  }

  if (!options.auditOnly) {
    runNonCritical_('applyRowCheckAtRow_', function() {
      for (let i = 0; i < rowIndexes.length; i++) applyRowCheckAtRow_(sh, hm, rowIndexes[i]);
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
    'Settlement split mode: ROW-LEVEL BY POSTED DATE',
    'Bucket Months Found: ' + parsed.bucketMonths.join(', '),
    'Assigned via POSTED_DATE: ' + parsed.assignmentStats.postedDateRows + ', DEPOSIT_FALLBACK: ' + parsed.assignmentStats.depositFallbackRows,
    'Split Integrity Check: ' + parsed.splitIntegrity.status + ' (diff ' + fromCents_(parsed.splitIntegrity.diffC).toFixed(2) + ')',
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
    'Summary rows upserted: ' + parsed.summaryRows.length,
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

  const dateReference = extractDateFromFileName_(fileMeta && fileMeta.name);
  const depositDate = parseDateFlexible_(depositDateRaw, CONFIG.TZ, { referenceDate: dateReference });
  if (!(depositDate instanceof Date) || isNaN(depositDate.getTime())) {
    throw buildFileDiagnosticError_('Не вдалося розпізнати Deposit Date: "' + depositDateRaw + '"', fileMeta, content, header, [
      'Підтримуються формати: YYYY-MM-DD, DD/MM/YYYY, MM/DD/YYYY, з часом/таймзоною.'
    ]);
  }

  const effectivePosted = resolveSettlementEffectiveDate_(dataRows, idx, depositDate, dateReference);

  const orderAgg = Object.create(null);
  const rawRows = [];
  const bucketBuild = buildSettlementMonthlyBuckets_(dataRows, idx, depositDate, dateReference);
  const bucketMonths = Object.keys(bucketBuild.buckets).sort();
  if (!bucketMonths.length) {
    const fallbackMonth = Utilities.formatDate(new Date(Date.UTC(depositDate.getUTCFullYear(), depositDate.getUTCMonth(), 1)), CONFIG.TZ, 'yyyy-MM');
    bucketBuild.buckets[fallbackMonth] = { monthKey: fallbackMonth, rows: [] };
    bucketMonths.push(fallbackMonth);
  }

  let salesC = 0;
  let vatC = 0;
  let feesExpenseC = 0;
  let reimbursementsC = 0;
  let otherC = 0;
  let transferBucketsSumC = 0;
  let cogsTotalC = 0;
  let unitsTotal = 0;
  let unitsWithCostTotal = 0;
  let missingUnitsTotal = 0;
  const missingSkuSet = new Set();
  const summaryRows = [];

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

    const t = String(amountType || '').trim();
    const d = String(amountDesc || '').trim();
    const idxOrderId = idx['order-id'];
    const idxSku = idx['sku'];
    const idxQty = idx['quantity-purchased'];
    const sku = idxSku !== undefined ? normalizeSku_(row[idxSku]) : '';
    const qty = idxQty !== undefined ? Math.max(0, Math.round(parseNumberLoose_(row[idxQty]))) : 0;
    const orderId = idxOrderId !== undefined ? String(row[idxOrderId] || '').trim() : '';

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

  for (let m = 0; m < bucketMonths.length; m++) {
    const monthKey = bucketMonths[m];
    const bucket = bucketBuild.buckets[monthKey];
    const agg = aggregateSettlementBucket_(bucket, idx, costMap, settlementId, depositDate, marketplaceName);
    summaryRows.push({ monthKey: monthKey, rowData: agg.rowData, stats: agg.stats });

    salesC += agg.salesC;
    vatC += agg.vatC;
    feesExpenseC += agg.feesExpenseC;
    reimbursementsC += agg.reimbursementsC;
    otherC += agg.otherC;
    transferBucketsSumC += agg.transferC;
    cogsTotalC += agg.cogsRes.cogsC;
    unitsTotal += agg.units;
    unitsWithCostTotal += agg.cogsRes.unitsWithCost;
    missingUnitsTotal += agg.cogsRes.missingUnits;
    (agg.cogsRes.missingSkus || []).forEach(function(sku) { if (sku) missingSkuSet.add(sku); });
  }

  const transferDiffC = transferC - transferBucketsSumC;
  if (transferDiffC !== 0 && summaryRows.length) {
    const first = summaryRows[0].rowData;
    const firstTransfer = toCents_(parseNumberFlexible_(first[CONFIG.HEADERS.transfer]));
    const firstOther = toCents_(parseNumberFlexible_(first[CONFIG.HEADERS.otherNet]));
    first[CONFIG.HEADERS.transfer] = fromCents_(firstTransfer + transferDiffC);
    first[CONFIG.HEADERS.otherNet] = fromCents_(firstOther + transferDiffC);
    transferBucketsSumC += transferDiffC;
    otherC += transferDiffC;
  }

  const payoutExReimbC = transferBucketsSumC - reimbursementsC;
  const netCashC = transferBucketsSumC - cogsTotalC;
  const soldProfitC = payoutExReimbC - cogsTotalC;
  const profitExReimbC = soldProfitC;
  const companyProfitC = netCashC;

  const diffC = (salesC + vatC + otherC - feesExpenseC) - transferBucketsSumC;
  const rowCheck = Math.abs(diffC) <= 1 ? 'OK' : ('ERR diff ' + fromCents_(diffC).toFixed(2));
  const splitIntegrityDiffC = transferC - transferBucketsSumC;

  return {
    header: header,
    rawRows: rawRows,
    idx: idx,
    settlementId: settlementId,
    depositDate: depositDate,
    postedDate: effectivePosted.date,
    postedDateSource: effectivePosted.source,
    postedDateCandidates: effectivePosted.candidates,
    monthDate: new Date(Date.UTC(effectivePosted.date.getUTCFullYear(), effectivePosted.date.getUTCMonth(), 1)),
    bucketMonths: bucketMonths,
    assignmentStats: bucketBuild.stats,
    splitIntegrity: { status: Math.abs(splitIntegrityDiffC) <= 1 ? 'OK' : 'FAIL', diffC: splitIntegrityDiffC },
    marketplaceName: marketplaceName,
    transferC: transferBucketsSumC,
    payoutExReimbC: payoutExReimbC,
    salesC: salesC,
    vatC: vatC,
    feesExpenseC: feesExpenseC,
    feesRule: 'split-by-month bucket aggregation',
    reimbursementsC: reimbursementsC,
    otherC: otherC,
    soldProfitC: soldProfitC,
    profitExReimbC: profitExReimbC,
    companyProfitC: companyProfitC,
    rowCheck: rowCheck,
    orderAgg: orderAgg,
    skuQtyMap: {},
    cogsRes: {
      cogsC: cogsTotalC,
      unitsWithCost: unitsWithCostTotal,
      missingUnits: missingUnitsTotal,
      coveragePct: unitsTotal > 0 ? (unitsWithCostTotal / unitsTotal) : 1,
      missingSkus: Array.from(missingSkuSet),
      missingSkusText: Array.from(missingSkuSet).join(', ')
    },
    units: unitsTotal,
    summaryRows: summaryRows,
    rowData: summaryRows.length ? summaryRows[0].rowData : {}
  };
}

function resolveRowMonthKey_(row, idx, fallbackDepositDate, referenceDate, stats) {
  const postedRaw = cellByHeader_(row, idx, 'posted-date');
  const posted = parseDateFlexible_(postedRaw, CONFIG.TZ, { referenceDate: referenceDate });
  if (posted instanceof Date && !isNaN(posted.getTime())) {
    if (stats) stats.postedDateRows += 1;
    return { monthKey: Utilities.formatDate(posted, CONFIG.TZ, 'yyyy-MM'), postedDate: posted, source: 'POSTED_DATE' };
  }
  if (stats) stats.depositFallbackRows += 1;
  return {
    monthKey: Utilities.formatDate(fallbackDepositDate, CONFIG.TZ, 'yyyy-MM'),
    postedDate: fallbackDepositDate,
    source: 'DEPOSIT_FALLBACK'
  };
}

function buildSettlementMonthlyBuckets_(rows, idx, fallbackDepositDate, referenceDate) {
  const out = {};
  const stats = { postedDateRows: 0, depositFallbackRows: 0 };
  for (let i = 0; i < (rows || []).length; i++) {
    const row = rows[i];
    const info = resolveRowMonthKey_(row, idx, fallbackDepositDate, referenceDate, stats);
    if (!out[info.monthKey]) out[info.monthKey] = { monthKey: info.monthKey, rows: [], postedDateRows: 0, depositFallbackRows: 0 };
    out[info.monthKey].rows.push(row);
    if (info.source === 'POSTED_DATE') out[info.monthKey].postedDateRows += 1;
    else out[info.monthKey].depositFallbackRows += 1;
  }
  return { buckets: out, stats: stats };
}

function aggregateSettlementBucket_(bucket, idx, costMap, settlementId, depositDate, marketplaceName) {
  const rows = bucket && bucket.rows ? bucket.rows : [];
  const monthParts = String(bucket.monthKey || '').split('-');
  const monthDate = new Date(Date.UTC(Number(monthParts[0] || 1970), Number(monthParts[1] || 1) - 1, 1));

  let transferC = 0;
  let salesC = 0;
  let vatC = 0;
  let itemFeesSignedSumC = 0;
  let feeNeg = 0;
  let feePos = 0;
  let reimbursementsC = 0;
  const skuQtyMap = Object.create(null);

  const idxSku = idx['sku'];
  const idxQty = idx['quantity-purchased'];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const amountType = cellByHeader_(row, idx, 'amount-type');
    const amountDesc = cellByHeader_(row, idx, 'amount-description');
    const transactionType = cellByHeader_(row, idx, 'transaction-type');
    const amountC = toCents_(parseNumberLoose_(cellByHeader_(row, idx, 'amount')));
    transferC += amountC;

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
    if (t === 'ItemPrice' && d === 'Principal' && sku && qty > 0) skuQtyMap[sku] = (skuQtyMap[sku] || 0) + qty;
  }

  const feesNorm = normalizeFeesExpenseC_(itemFeesSignedSumC, feeNeg, feePos);
  const feesExpenseC = feesNorm.feesExpenseC;
  const cogsRes = calcCogsFromCostMap_(skuQtyMap, costMap);
  const units = Object.keys(skuQtyMap).reduce(function(acc, sku) { return acc + Number(skuQtyMap[sku] || 0); }, 0);

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
  rowData[CONFIG.HEADERS.postedDate] = monthDate;
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
  rowData[CONFIG.HEADERS.rowCheck] = rowCheck + ' | split ' + bucket.monthKey + ' rows:' + rows.length + ' posted:' + (bucket.postedDateRows || 0) + ' fallback:' + (bucket.depositFallbackRows || 0);

  return {
    monthKey: bucket.monthKey,
    transferC: transferC,
    salesC: salesC,
    vatC: vatC,
    feesExpenseC: feesExpenseC,
    reimbursementsC: reimbursementsC,
    otherC: otherC,
    payoutExReimbC: payoutExReimbC,
    cogsRes: cogsRes,
    units: units,
    rowData: rowData,
    stats: { rows: rows.length, postedDateRows: bucket.postedDateRows || 0, depositFallbackRows: bucket.depositFallbackRows || 0 }
  };
}

function resolveSettlementEffectiveDate_(rows, idx, depositDate, referenceDate) {
  const postedIdx = idx && idx['posted-date'] !== undefined ? idx['posted-date'] : -1;
  let earliest = null;
  let candidates = 0;

  if (postedIdx >= 0) {
    for (let i = 0; i < (rows || []).length; i++) {
      const rawPosted = rows[i] && rows[i][postedIdx] !== undefined ? rows[i][postedIdx] : '';
      const parsed = parseDateFlexible_(rawPosted, CONFIG.TZ, { referenceDate: referenceDate });
      if (!(parsed instanceof Date) || isNaN(parsed.getTime())) continue;
      candidates += 1;
      if (!earliest || parsed.getTime() < earliest.getTime()) earliest = parsed;
    }
  }

  if (earliest) {
    return { date: earliest, source: 'posted-date rows (earliest)', candidates: candidates };
  }

  return { date: depositDate, source: 'deposit-date fallback', candidates: 0 };
}

/* =========================
 * MONTHLY
 * ========================= */

function rebuildMonthly_(warnings) {
  warnings = warnings || [];
  runNonCritical_('rebuildMonthlyVatPayoutSummary_ [migration]', function() {
    rebuildMonthlyVatPayoutSummary_();
  }, warnings);

  migrateLegacyMonthlySheetToMain_();
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

function migrateLegacyMonthlySheetToMain_() {
  const main = getMonthlyVatPayoutSheet_();
  const legacy = getLegacyMonthlySheet_();
  if (!legacy || !main || legacy.getSheetId() === main.getSheetId()) return { migratedRows: 0, legacyRows: 0 };

  const migratedRows = migrateLegacyMonthlyRowsToMain_(legacy, main);
  markLegacyMonthlySheetDeprecated_(legacy, main, migratedRows);
  return { migratedRows: migratedRows, legacyRows: Math.max(0, legacy.getLastRow() - 1) };
}

function ensureMonthAndTotals_(warnings) {
  warnings = warnings || [];
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SUMMARY_SHEET);
  if (!sh) throw new Error('Не знайдено вкладку "' + CONFIG.SUMMARY_SHEET + '"');

  ensureSummaryHeaders_(sh);
  const hm = getHeaderMap_(sh);

  const colFileId = hm[CONFIG.HEADERS.fileId];
  const colMonth = hm[CONFIG.HEADERS.month];
  const colPosted = hm[CONFIG.HEADERS.postedDate];
  const colDeposit = hm[CONFIG.HEADERS.depositDate];
  const colSettlId = hm[CONFIG.HEADERS.settlementId];
  if (!colFileId || !colMonth || (!colPosted && !colDeposit)) return;

  let lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  for (let r = lastRow; r >= 2; r--) {
    const fid = String(sh.getRange(r, colFileId).getValue() || '').trim();
    if (fid === CONFIG.TOTAL_FILE_ID) sh.deleteRow(r);
  }

  lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const ids = sh.getRange(2, colFileId, lastRow - 1, 1).getValues();
  const postedVals = colPosted ? sh.getRange(2, colPosted, lastRow - 1, 1).getValues() : [];
  const deposits = colDeposit ? sh.getRange(2, colDeposit, lastRow - 1, 1).getValues() : [];

  let lastDataRow = 1;
  const monthVals = [];

  for (let i = 0; i < ids.length; i++) {
    const fid = String(ids[i][0] || '').trim();
    if (!fid || fid === CONFIG.TOTAL_FILE_ID) {
      monthVals.push(['']);
      continue;
    }

    const p = colPosted ? postedVals[i][0] : '';
    const d = colDeposit ? deposits[i][0] : '';
    const basis = (p instanceof Date && !isNaN(p.getTime())) ? p : d;
    if (basis instanceof Date && !isNaN(basis.getTime())) {
      monthVals.push([new Date(Date.UTC(basis.getUTCFullYear(), basis.getUTCMonth(), 1))]);
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

function findOrCreateRowBySettlementMonth_(sh, hm, fileId, settlementId, monthKey) {
  const colFileId = hm[CONFIG.HEADERS.fileId];
  const colSettlementId = hm[CONFIG.HEADERS.settlementId];
  const colMonth = hm[CONFIG.HEADERS.month];
  if (!colFileId || !colSettlementId || !colMonth) return findFirstEmptyRow_(sh, hm[CONFIG.HEADERS.fileId], 2);

  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const existingFileId = String(row[colFileId - 1] || '').trim();
      if (!existingFileId || existingFileId === CONFIG.TOTAL_FILE_ID) continue;
      const existingSettlementId = String(row[colSettlementId - 1] || '').trim();
      const existingMonth = toMonthText_(row[colMonth - 1]);
      if (existingFileId === String(fileId || '').trim() && existingSettlementId === String(settlementId || '').trim() && existingMonth === monthKey) {
        return 2 + i;
      }
    }
  }

  return findFirstEmptyRow_(sh, hm[CONFIG.HEADERS.fileId], 2);
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

function parseDateFlexible_(value, tz, options) {
  options = options || {};
  const referenceDate = options.referenceDate instanceof Date && !isNaN(options.referenceDate.getTime())
    ? options.referenceDate
    : null;

  const raw = String(value || '').trim();
  if (!raw) return null;

  let m = raw.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?(?:\s*(Z|UTC|[+-]\d{2}:?\d{2})?)?$/i);
  if (m) {
    const y=Number(m[1]), mo=Number(m[2])-1, d=Number(m[3]), hh=Number(m[4]||0), mm=Number(m[5]||0), ss=Number(m[6]||0);
    return new Date(Date.UTC(y, mo, d, hh, mm, ss));
  }

  m = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    const a=Number(m[1]), b=Number(m[2]), y=Number(m[3]), hh=Number(m[4]||0), mm=Number(m[5]||0), ss=Number(m[6]||0);
    if (a > 12) return new Date(Date.UTC(y, b - 1, a, hh, mm, ss));
    if (b > 12) return new Date(Date.UTC(y, a - 1, b, hh, mm, ss));

    const dmy = new Date(Date.UTC(y, b - 1, a, hh, mm, ss));
    const mdy = new Date(Date.UTC(y, a - 1, b, hh, mm, ss));
    return chooseDateByReference_(dmy, mdy, referenceDate) || dmy;
  }

  m = raw.match(/^(\d{2})\.(\d{2})\.(\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?(?:\s*UTC)?$/i);
  if (m) return new Date(Date.UTC(Number(m[3]), Number(m[2]) - 1, Number(m[1]), Number(m[4]||0), Number(m[5]||0), Number(m[6]||0)));

  const direct = new Date(raw);
  if (!isNaN(direct.getTime())) return direct;

  return null;
}

function chooseDateByReference_(a, b, referenceDate) {
  if (!(a instanceof Date) || isNaN(a.getTime())) return b;
  if (!(b instanceof Date) || isNaN(b.getTime())) return a;
  if (!(referenceDate instanceof Date) || isNaN(referenceDate.getTime())) return a;

  const diffA = Math.abs(a.getTime() - referenceDate.getTime());
  const diffB = Math.abs(b.getTime() - referenceDate.getTime());
  return diffA <= diffB ? a : b;
}

function extractDateFromFileName_(fileName) {
  const name = String(fileName || '');
  let m = name.match(/(\d{2})_(\d{2})_(\d{4})/);
  if (m) {
    const d = Number(m[1]);
    const mo = Number(m[2]) - 1;
    const y = Number(m[3]);
    return new Date(Date.UTC(y, mo, d));
  }

  m = name.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (m) {
    return new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
  }

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
  const essential = ['settlement-id', 'amount-type', 'amount-description', 'transaction-type', 'amount'];
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
    'posted-date': ['posted-date', 'posted-date-time', 'posteddate', 'posteddatetime'],
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
    ['settlement-id', 'amount-type', 'amount-description', 'amount'].forEach(function(k) {
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

  const dateCols = [hm[CONFIG.HEADERS.depositDate], hm[CONFIG.HEADERS.postedDate], hm[CONFIG.HEADERS.month], hm[CONFIG.HEADERS.importedAt]].filter(Boolean);
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
    SpreadsheetApp.getActive().toast(String(msg || ''), 'Фінанси Amazon', 8);
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


/* =========================
 * VAT / SALES FROM AMAZON TAX REPORT
 * ========================= */

const TAX_REQUIRED_HEADERS = [
  'Order Date',
  'Order ID',
  'Transaction Type',
  'Quantity',
  'Tax Calculation Date',
  'Tax Rate',
  'Tax Collection Responsibility',
  'Ship To Country',
  'OUR_PRICE Tax Inclusive Selling Price',
  'OUR_PRICE Tax Amount',
  'OUR_PRICE Tax Exclusive Selling Price',
  'OUR_PRICE Tax Inclusive Promo Amount',
  'OUR_PRICE Tax Amount Promo',
  'OUR_PRICE Tax Exclusive Promo Amount',
  'SHIPPING Tax Inclusive Selling Price',
  'SHIPPING Tax Amount',
  'SHIPPING Tax Exclusive Selling Price',
  'SHIPPING Tax Inclusive Promo Amount',
  'SHIPPING Tax Amount Promo',
  'SHIPPING Tax Exclusive Promo Amount',
  'GIFTWRAP Tax Inclusive Selling Price',
  'GIFTWRAP Tax Amount',
  'GIFTWRAP Tax Exclusive Selling Price',
  'GIFTWRAP Tax Inclusive Promo Amount',
  'GIFTWRAP Tax Amount Promo',
  'GIFTWRAP Tax Exclusive Promo Amount'
];

const TAX_COMPUTED_HEADERS = [
  'Period (YYYY-MM)',
  'Net Product',
  'VAT Product',
  'Gross Product',
  'Net Shipping',
  'VAT Shipping',
  'Gross Shipping',
  'Net Giftwrap',
  'VAT Giftwrap',
  'Gross Giftwrap',
  'Net Total',
  'VAT Total',
  'Gross Total'
];

const VAT_SUMMARY_HEADERS = [
  'Period (YYYY-MM)',
  'Ship To Country',
  'Tax Rate',
  'Tax Collection Responsibility',
  'Transaction Type',
  'Orders (distinct Order ID count)',
  'Units (sum Quantity)',
  'Net Sales (Total)',
  'VAT (Total)',
  'Gross Sales (Total)',
  'Net Product',
  'VAT Product',
  'Net Shipping',
  'VAT Shipping',
  'Net Giftwrap',
  'VAT Giftwrap',
  'Notes'
];

function uiImportTaxReportByFileId_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const prompt = ui.prompt(
      'Імпорт Tax Report (CSV) — File ID або Folder ID',
      'Введіть Google Drive File ID (або Folder ID).\nЯкщо залишити порожнім — автоматично візьметься останній CSV з CONFIG.TAX_REPORT_FOLDER_ID.',
      ui.ButtonSet.OK_CANCEL
    );
    if (prompt.getSelectedButton() !== ui.Button.OK) return;
    const driveIdOrEmpty = String(prompt.getResponseText() || '').trim();

    if (!driveIdOrEmpty) {
      const latest = importLatestTaxReportFromConfiguredFolder_(CONFIG.DEFAULT_GROUP_DATE_FIELD);
      ui.alert('Імпорт Tax Report завершено автоматично.\nФайл: ' + latest.fileName + '\nІмпортовано рядків: ' + latest.rows + '\nАркуш: ' + CONFIG.TAX_RAW_SHEET);
      return;
    }

    const result = importTaxReportByFileOrFolderId_(driveIdOrEmpty, CONFIG.DEFAULT_GROUP_DATE_FIELD);
    ui.alert('Імпорт Tax Report завершено.\nФайл: ' + result.fileName + '\nІмпортовано рядків: ' + result.rows + '\nАркуш: ' + CONFIG.TAX_RAW_SHEET);
  } catch (e) {
    ui.alert('Помилка імпорту Tax Report: ' + toErrorMessage_(e));
  }
}

function uiImportTaxReportLatestFromFolder_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const latest = importLatestTaxReportFromConfiguredFolder_(CONFIG.DEFAULT_GROUP_DATE_FIELD);
    ui.alert('Імпортовано останній Tax Report: ' + latest.fileName + '\nРядків: ' + latest.rows);
  } catch (e) {
    ui.alert('Помилка імпорту останнього Tax Report: ' + toErrorMessage_(e));
  }
}

function importLatestTaxReportFromConfiguredFolder_(groupDateField) {
  if (!CONFIG.TAX_REPORT_FOLDER_ID) {
    throw new Error('CONFIG.TAX_REPORT_FOLDER_ID порожній. Вкажіть ID папки у CONFIG.');
  }
  const files = getTaxCsvCandidatesFromFolder_(CONFIG.TAX_REPORT_FOLDER_ID, 1);
  if (!files.length) {
    throw new Error('У папці не знайдено CSV файлів: ' + CONFIG.TAX_REPORT_FOLDER_ID);
  }
  const result = importTaxReportCsvFromFileId_(files[0].id, groupDateField);
  return { rows: result.rows, fileName: files[0].name, fileId: files[0].id };
}

function importTaxReportByFileOrFolderId_(driveId, groupDateField) {
  const id = String(driveId || '').trim();
  if (!id) throw new Error('Передайте Drive ID файлу або папки.');

  try {
    const file = DriveApp.getFileById(id);
    const resultByFile = importTaxReportCsvFromFileId_(id, groupDateField);
    return { rows: resultByFile.rows, fileName: file.getName(), fileId: id };
  } catch (fileErr) {
    try {
      const files = getTaxCsvCandidatesFromFolder_(id, 1);
      if (!files.length) throw new Error('У папці не знайдено CSV файлів: ' + id);
      const resultByFolder = importTaxReportCsvFromFileId_(files[0].id, groupDateField);
      return { rows: resultByFolder.rows, fileName: files[0].name, fileId: files[0].id };
    } catch (folderErr) {
      throw new Error(
        'Не вдалося імпортувати за переданим ID. Очікується File ID CSV або Folder ID з CSV. ' +
        'Деталі File ID: ' + toErrorMessage_(fileErr) + '. Деталі Folder ID: ' + toErrorMessage_(folderErr)
      );
    }
  }
}

function uiBuildVatSalesSummary_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = buildVatSalesSummary_(CONFIG.DEFAULT_GROUP_DATE_FIELD);
    ui.alert('VAT/Sales зведення побудовано.\nРядків: ' + res.rows + '\nVAT до сплати (Seller): ' + res.vatPayableSeller.toFixed(2) + ' ' + CONFIG.CURRENCY + '\nVAT зібрано Marketplace/Amazon: ' + res.vatCollectedMarketplace.toFixed(2) + ' ' + CONFIG.CURRENCY);
  } catch (e) {
    ui.alert('Помилка побудови VAT/Sales зведення: ' + toErrorMessage_(e));
  }
}

function uiBuildVatSalesSummaryByOrderDate_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = buildVatSalesSummary_(CONFIG.ALT_GROUP_DATE_FIELD);
    ui.alert('VAT/Sales зведення (Order Date) побудовано.\nРядків: ' + res.rows + '\nVAT до сплати (Seller): ' + res.vatPayableSeller.toFixed(2) + ' ' + CONFIG.CURRENCY + '\nVAT зібрано Marketplace/Amazon: ' + res.vatCollectedMarketplace.toFixed(2) + ' ' + CONFIG.CURRENCY);
  } catch (e) {
    ui.alert('Помилка побудови VAT/Sales зведення за Order Date: ' + toErrorMessage_(e));
  }
}


function uiBuildCurrentMonthVatSalesByShipmentDate_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = buildCurrentMonthVatSalesSnapshot_();
    const lines = [
      'Поточний місяць: ' + res.period,
      'Поле дати: ' + res.groupingDateField,
      'Рядків враховано: ' + res.rows,
      'Замовлень: ' + res.orders,
      'Одиниць: ' + res.units,
      'Net Sales: ' + res.netSales.toFixed(2) + ' ' + CONFIG.CURRENCY,
      'VAT (всього): ' + res.vatTotal.toFixed(2) + ' ' + CONFIG.CURRENCY,
      'Gross Sales: ' + res.grossSales.toFixed(2) + ' ' + CONFIG.CURRENCY,
      'VAT до сплати (Seller): ' + res.vatPayableSeller.toFixed(2) + ' ' + CONFIG.CURRENCY,
      'VAT зібрано Marketplace/Amazon: ' + res.vatCollectedMarketplace.toFixed(2) + ' ' + CONFIG.CURRENCY
    ];
    if (res.warning) lines.push('Увага: ' + res.warning);
    ui.alert('Поточний місяць (Shipment Date) — підрахунок завершено', lines.join('\n'), ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Помилка підрахунку поточного місяця за Shipment Date: ' + toErrorMessage_(e));
  }
}


function uiValidateTaxReportHeaders_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const report = validateTaxReportHeaders_();
    ui.alert(report);
  } catch (e) {
    ui.alert('Помилка діагностики: ' + toErrorMessage_(e));
  }
}

function importTaxReportCsvFromFileId_(fileId, groupDateField) {
  const file = DriveApp.getFileById(fileId);
  const text = file.getBlob().getDataAsString();
  const parsed = parseTaxReportTable_(text);
  const csv = parsed.rows;
  if (!csv || csv.length < 2) {
    throw new Error('CSV has no data rows. Delimiter detected: ' + parsed.delimiterLabel + '.');
  }

  const headers = (csv[0] || []).map(function(h) { return String(h || '').trim(); });
  const hm = buildHeaderMapCaseInsensitive_(headers);
  const missing = findMissingHeaders_(hm, TAX_REQUIRED_HEADERS);
  if (missing.length) throw new Error('Missing required headers: ' + missing.join(', '));

  const dateField = normalizeHeaderKey_(groupDateField) === normalizeHeaderKey_(CONFIG.ALT_GROUP_DATE_FIELD)
    ? CONFIG.ALT_GROUP_DATE_FIELD
    : CONFIG.DEFAULT_GROUP_DATE_FIELD;

  const rawRows = [];
  for (let i = 1; i < csv.length; i++) {
    const row = csv[i];
    if (!row || isEmptyRow_(row)) continue;
    const ext = buildTaxComputedColumns_(headers, row, dateField);
    rawRows.push(headers.map(function(_, idx) { return row[idx] !== undefined ? row[idx] : ''; }).concat(ext));
  }

  const sheet = getOrCreateSheet_(CONFIG.TAX_RAW_SHEET);
  sheet.clear();
  const finalHeaders = headers.concat(TAX_COMPUTED_HEADERS);
  sheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
  if (rawRows.length) {
    sheet.getRange(2, 1, rawRows.length, finalHeaders.length).setValues(rawRows);
    applyTaxRawFormats_(sheet, rawRows.length, headers.length, TAX_COMPUTED_HEADERS.length);
  }
  return { rows: rawRows.length };
}

function applyTaxRawFormats_(sheet, rowCount, sourceColsCount, computedColsCount) {
  if (!sheet || rowCount <= 0) return;

  if (sourceColsCount > 0) {
    // Keep imported CSV values as plain text so decimal values like "01.09" are not auto-converted to dates.
    sheet.getRange(2, 1, rowCount, sourceColsCount).setNumberFormat('@');
  }

  if (computedColsCount > 0) {
    sheet.getRange(2, sourceColsCount + 1, rowCount, 1).setNumberFormat('yyyy-MM');
    if (computedColsCount > 1) {
      sheet.getRange(2, sourceColsCount + 2, rowCount, computedColsCount - 1).setNumberFormat('#,##0.00');
    }
  }
}

function parseTaxReportTable_(text) {
  const cleaned = String(text || '').replace(/^\uFEFF/, '');
  const lines = cleaned.split(/\r?\n/).filter(function(line) { return String(line || '').trim() !== ''; });
  if (!lines.length) return { rows: [], delimiterLabel: 'none' };

  const headerLine = lines[0];
  const delimiterCandidates = [
    { char: ',', label: 'comma' },
    { char: ';', label: 'semicolon' },
    { char: '\t', label: 'tab' }
  ];

  let best = delimiterCandidates[0];
  let bestCount = -1;
  for (let i = 0; i < delimiterCandidates.length; i++) {
    const d = delimiterCandidates[i];
    const count = headerLine.split(d.char).length;
    if (count > bestCount) {
      best = d;
      bestCount = count;
    }
  }

  return {
    rows: Utilities.parseCsv(cleaned, best.char),
    delimiterLabel: best.label
  };
}

function buildTaxComputedColumns_(headers, row, groupDateField) {
  const hm = buildHeaderMapCaseInsensitive_(headers);
  function n(name) {
    const idx = hm[normalizeHeaderKey_(name)];
    return parseNumberFlexible_(idx === undefined ? '' : row[idx]);
  }

  const periodDate = parseAmazonUtcDate_(valueByHeader_(row, hm, groupDateField));
  const period = periodDate ? Utilities.formatDate(periodDate, CONFIG.TZ, 'yyyy-MM') : '';

  const netProduct = n('OUR_PRICE Tax Exclusive Selling Price') - n('OUR_PRICE Tax Exclusive Promo Amount');
  const vatProduct = n('OUR_PRICE Tax Amount') - n('OUR_PRICE Tax Amount Promo');
  const grossProduct = n('OUR_PRICE Tax Inclusive Selling Price') - n('OUR_PRICE Tax Inclusive Promo Amount');

  const netShipping = n('SHIPPING Tax Exclusive Selling Price') - n('SHIPPING Tax Exclusive Promo Amount');
  const vatShipping = n('SHIPPING Tax Amount') - n('SHIPPING Tax Amount Promo');
  const grossShipping = n('SHIPPING Tax Inclusive Selling Price') - n('SHIPPING Tax Inclusive Promo Amount');

  const netGiftwrap = n('GIFTWRAP Tax Exclusive Selling Price') - n('GIFTWRAP Tax Exclusive Promo Amount');
  const vatGiftwrap = n('GIFTWRAP Tax Amount') - n('GIFTWRAP Tax Amount Promo');
  const grossGiftwrap = n('GIFTWRAP Tax Inclusive Selling Price') - n('GIFTWRAP Tax Inclusive Promo Amount');

  const netTotal = netProduct + netShipping + netGiftwrap;
  const vatTotal = vatProduct + vatShipping + vatGiftwrap;
  const grossTotal = grossProduct + grossShipping + grossGiftwrap;

  return [
    period,
    netProduct, vatProduct, grossProduct,
    netShipping, vatShipping, grossShipping,
    netGiftwrap, vatGiftwrap, grossGiftwrap,
    netTotal, vatTotal, grossTotal
  ];
}

function buildVatSalesSummary_(groupDateField) {
  const raw = getOrCreateSheet_(CONFIG.TAX_RAW_SHEET);
  const lastRow = raw.getLastRow();
  const lastCol = raw.getLastColumn();
  if (lastRow < 2 || lastCol < 1) throw new Error('TAX_REPORT_RAW is empty. Import CSV first.');

  const all = raw.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = all[0].map(function(h) { return String(h || '').trim(); });
  const hm = buildHeaderMapCaseInsensitive_(headers);

  const requiredRaw = [
    'Order ID', 'Quantity', 'Tax Rate', 'Ship To Country', 'Tax Collection Responsibility', 'Transaction Type',
    'Net Total', 'VAT Total', 'Gross Total', 'Net Product', 'VAT Product', 'Net Shipping', 'VAT Shipping', 'Net Giftwrap', 'VAT Giftwrap',
    CONFIG.DEFAULT_GROUP_DATE_FIELD, CONFIG.ALT_GROUP_DATE_FIELD
  ];
  const missing = findMissingHeaders_(hm, requiredRaw);
  if (missing.length) throw new Error('Missing columns in TAX_REPORT_RAW: ' + missing.join(', '));

  const useDateField = normalizeHeaderKey_(groupDateField) === normalizeHeaderKey_(CONFIG.ALT_GROUP_DATE_FIELD)
    ? CONFIG.ALT_GROUP_DATE_FIELD
    : CONFIG.DEFAULT_GROUP_DATE_FIELD;

  const grouped = {};
  const uniqResponsibility = {};
  for (let i = 1; i < all.length; i++) {
    const row = all[i];
    if (isEmptyRow_(row)) continue;

    const d = parseAmazonUtcDate_(valueByHeader_(row, hm, useDateField));
    const period = d ? Utilities.formatDate(d, CONFIG.TZ, 'yyyy-MM') : '';
    const shipToCountry = String(valueByHeader_(row, hm, 'Ship To Country') || '').trim();
    const taxRate = parseNumberFlexible_(valueByHeader_(row, hm, 'Tax Rate'));
    const responsibilityRaw = String(valueByHeader_(row, hm, 'Tax Collection Responsibility') || '').trim();
    const responsibilityNorm = responsibilityRaw.toLowerCase();
    uniqResponsibility[responsibilityNorm || '(empty)'] = responsibilityRaw || '(empty)';
    const transactionType = String(valueByHeader_(row, hm, 'Transaction Type') || '').trim() || 'ALL';

    const key = [period, shipToCountry, taxRate.toFixed(6), responsibilityNorm, transactionType].join('||');
    if (!grouped[key]) {
      grouped[key] = {
        period: period,
        shipToCountry: shipToCountry,
        taxRate: taxRate,
        responsibility: responsibilityRaw,
        transactionType: transactionType,
        orderIds: {},
        units: 0,
        netTotal: 0,
        vatTotal: 0,
        grossTotal: 0,
        netProduct: 0,
        vatProduct: 0,
        netShipping: 0,
        vatShipping: 0,
        netGiftwrap: 0,
        vatGiftwrap: 0
      };
    }

    const g = grouped[key];
    const orderId = String(valueByHeader_(row, hm, 'Order ID') || '').trim();
    if (orderId) g.orderIds[orderId] = true;
    g.units += parseNumberFlexible_(valueByHeader_(row, hm, 'Quantity'));
    g.netTotal += parseNumberFlexible_(valueByHeader_(row, hm, 'Net Total'));
    g.vatTotal += parseNumberFlexible_(valueByHeader_(row, hm, 'VAT Total'));
    g.grossTotal += parseNumberFlexible_(valueByHeader_(row, hm, 'Gross Total'));
    g.netProduct += parseNumberFlexible_(valueByHeader_(row, hm, 'Net Product'));
    g.vatProduct += parseNumberFlexible_(valueByHeader_(row, hm, 'VAT Product'));
    g.netShipping += parseNumberFlexible_(valueByHeader_(row, hm, 'Net Shipping'));
    g.vatShipping += parseNumberFlexible_(valueByHeader_(row, hm, 'VAT Shipping'));
    g.netGiftwrap += parseNumberFlexible_(valueByHeader_(row, hm, 'Net Giftwrap'));
    g.vatGiftwrap += parseNumberFlexible_(valueByHeader_(row, hm, 'VAT Giftwrap'));
  }

  const keys = Object.keys(grouped).sort();
  const out = [];
  let vatPayableSeller = 0;
  let vatCollectedMarketplace = 0;

  for (let i = 0; i < keys.length; i++) {
    const g = grouped[keys[i]];
    const isSeller = String(g.responsibility || '').trim().toLowerCase() === 'seller';
    if (isSeller) vatPayableSeller += g.vatTotal;
    else vatCollectedMarketplace += g.vatTotal;
    out.push([
      g.period,
      g.shipToCountry,
      g.taxRate,
      g.responsibility,
      g.transactionType || 'ALL',
      Object.keys(g.orderIds).length,
      g.units,
      g.netTotal,
      g.vatTotal,
      g.grossTotal,
      g.netProduct,
      g.vatProduct,
      g.netShipping,
      g.vatShipping,
      g.netGiftwrap,
      g.vatGiftwrap,
      isSeller ? 'Seller-responsible VAT' : 'Marketplace collected'
    ]);
  }

  const summary = getOrCreateSheet_(CONFIG.VAT_SUMMARY_SHEET);
  summary.clearContents();
  summary.getRange(1, 1, 1, VAT_SUMMARY_HEADERS.length).setValues([VAT_SUMMARY_HEADERS]);
  if (out.length) summary.getRange(2, 1, out.length, VAT_SUMMARY_HEADERS.length).setValues(out);

  const metaStart = out.length + 4;
  const uniqVals = Object.keys(uniqResponsibility).map(function(k) { return uniqResponsibility[k]; }).sort();
  const meta = [
    ['Metric', 'Value'],
    ['Grouping Date Field', useDateField],
    ['VAT Payable (Seller)', vatPayableSeller],
    ['VAT Collected by Marketplace/Amazon', vatCollectedMarketplace],
    ['Unique Tax Collection Responsibility', uniqVals.join(', ')]
  ];
  summary.getRange(metaStart, 1, meta.length, 2).setValues(meta);

  applyVatSummaryFormats_(summary, out.length, metaStart + 2);

  return { rows: out.length, vatPayableSeller: vatPayableSeller, vatCollectedMarketplace: vatCollectedMarketplace };
}


function buildCurrentMonthVatSalesSnapshot_() {
  const raw = getOrCreateSheet_(CONFIG.TAX_RAW_SHEET);
  const lastRow = raw.getLastRow();
  const lastCol = raw.getLastColumn();
  if (lastRow < 2 || lastCol < 1) throw new Error('TAX_REPORT_RAW is empty. Import CSV first.');

  const all = raw.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = all[0].map(function(h) { return String(h || '').trim(); });
  const hm = buildHeaderMapCaseInsensitive_(headers);

  const shipmentHeader = resolveHeaderName_(hm, CONFIG.SHIPMENT_DATE_FIELD);
  const defaultDateHeader = resolveHeaderName_(hm, CONFIG.DEFAULT_GROUP_DATE_FIELD);
  const dateHeader = shipmentHeader || defaultDateHeader;
  if (!dateHeader) {
    throw new Error('Не знайдено колонки для дати. Очікується "' + CONFIG.SHIPMENT_DATE_FIELD + '" або "' + CONFIG.DEFAULT_GROUP_DATE_FIELD + '".');
  }

  const required = [
    'Order ID', 'Quantity', 'Tax Collection Responsibility',
    'Net Total', 'VAT Total', 'Gross Total',
    'Net Product', 'VAT Product', 'Net Shipping', 'VAT Shipping', 'Net Giftwrap', 'VAT Giftwrap'
  ];
  const missing = findMissingHeaders_(hm, required);
  if (missing.length) throw new Error('Missing columns in TAX_REPORT_RAW: ' + missing.join(', '));

  const now = new Date();
  const currentPeriod = Utilities.formatDate(now, CONFIG.TZ, 'yyyy-MM');

  const rowsOut = [];
  const orders = {};
  let units = 0;
  let netSales = 0;
  let vatTotal = 0;
  let grossSales = 0;
  let vatPayableSeller = 0;
  let vatCollectedMarketplace = 0;
  let netProduct = 0, vatProduct = 0, netShipping = 0, vatShipping = 0, netGiftwrap = 0, vatGiftwrap = 0;

  for (let i = 1; i < all.length; i++) {
    const row = all[i];
    if (isEmptyRow_(row)) continue;

    const d = parseAmazonUtcDate_(valueByHeader_(row, hm, dateHeader));
    if (!d) continue;
    const period = Utilities.formatDate(d, CONFIG.TZ, 'yyyy-MM');
    if (period !== currentPeriod) continue;

    rowsOut.push(row);
    const orderId = String(valueByHeader_(row, hm, 'Order ID') || '').trim();
    if (orderId) orders[orderId] = true;

    const rowUnits = parseNumberFlexible_(valueByHeader_(row, hm, 'Quantity'));
    const rowNet = parseNumberFlexible_(valueByHeader_(row, hm, 'Net Total'));
    const rowVat = parseNumberFlexible_(valueByHeader_(row, hm, 'VAT Total'));
    const rowGross = parseNumberFlexible_(valueByHeader_(row, hm, 'Gross Total'));

    units += rowUnits;
    netSales += rowNet;
    vatTotal += rowVat;
    grossSales += rowGross;

    netProduct += parseNumberFlexible_(valueByHeader_(row, hm, 'Net Product'));
    vatProduct += parseNumberFlexible_(valueByHeader_(row, hm, 'VAT Product'));
    netShipping += parseNumberFlexible_(valueByHeader_(row, hm, 'Net Shipping'));
    vatShipping += parseNumberFlexible_(valueByHeader_(row, hm, 'VAT Shipping'));
    netGiftwrap += parseNumberFlexible_(valueByHeader_(row, hm, 'Net Giftwrap'));
    vatGiftwrap += parseNumberFlexible_(valueByHeader_(row, hm, 'VAT Giftwrap'));

    const responsibility = String(valueByHeader_(row, hm, 'Tax Collection Responsibility') || '').trim().toLowerCase();
    if (responsibility === 'seller') vatPayableSeller += rowVat;
    else vatCollectedMarketplace += rowVat;
  }

  writeCurrentMonthSnapshotSheet_({
    period: currentPeriod,
    groupingDateField: dateHeader,
    rows: rowsOut.length,
    orders: Object.keys(orders).length,
    units: units,
    netSales: netSales,
    vatTotal: vatTotal,
    grossSales: grossSales,
    netProduct: netProduct,
    vatProduct: vatProduct,
    netShipping: netShipping,
    vatShipping: vatShipping,
    netGiftwrap: netGiftwrap,
    vatGiftwrap: vatGiftwrap,
    vatPayableSeller: vatPayableSeller,
    vatCollectedMarketplace: vatCollectedMarketplace,
    warning: shipmentHeader ? '' : 'Колонку Shipment Date не знайдено, використано Tax Calculation Date.'
  });

  return {
    period: currentPeriod,
    groupingDateField: dateHeader,
    rows: rowsOut.length,
    orders: Object.keys(orders).length,
    units: units,
    netSales: netSales,
    vatTotal: vatTotal,
    grossSales: grossSales,
    vatPayableSeller: vatPayableSeller,
    vatCollectedMarketplace: vatCollectedMarketplace,
    warning: shipmentHeader ? '' : 'Колонку Shipment Date не знайдено, використано Tax Calculation Date.'
  };
}

function writeCurrentMonthSnapshotSheet_(snapshot) {
  const sheet = getOrCreateSheet_(CONFIG.CURRENT_MONTH_SNAPSHOT_SHEET);
  sheet.clearContents();

  const headers = ['Metric', 'Value'];
  const rows = [
    ['Period (YYYY-MM)', snapshot.period],
    ['Grouping Date Field', snapshot.groupingDateField],
    ['Rows Included', snapshot.rows],
    ['Orders (distinct)', snapshot.orders],
    ['Units', snapshot.units],
    ['Net Sales (Total)', snapshot.netSales],
    ['VAT (Total)', snapshot.vatTotal],
    ['Gross Sales (Total)', snapshot.grossSales],
    ['Net Product', snapshot.netProduct],
    ['VAT Product', snapshot.vatProduct],
    ['Net Shipping', snapshot.netShipping],
    ['VAT Shipping', snapshot.vatShipping],
    ['Net Giftwrap', snapshot.netGiftwrap],
    ['VAT Giftwrap', snapshot.vatGiftwrap],
    ['VAT Payable (Seller)', snapshot.vatPayableSeller],
    ['VAT Collected by Marketplace/Amazon', snapshot.vatCollectedMarketplace],
    ['Warning', snapshot.warning || '']
  ];

  sheet.getRange(1, 1, 1, 2).setValues([headers]);
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);

  sheet.getRange(2, 2, 4, 1).setNumberFormat('0');
  sheet.getRange(6, 2, 11, 1).setNumberFormat('#,##0.00');
}

function resolveHeaderName_(hm, expectedHeader) {
  const key = normalizeHeaderKey_(expectedHeader);
  const idx = hm[key];
  return idx === undefined ? '' : expectedHeader;
}

function validateTaxReportHeaders_() {
  const raw = getOrCreateSheet_(CONFIG.TAX_RAW_SHEET);
  const lastRow = raw.getLastRow();
  const lastCol = raw.getLastColumn();
  if (lastRow < 1 || lastCol < 1) throw new Error('TAX_REPORT_RAW sheet is empty.');

  const headers = raw.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || '').trim(); });
  const hm = buildHeaderMapCaseInsensitive_(headers);
  const missing = findMissingHeaders_(hm, TAX_REQUIRED_HEADERS.concat(TAX_COMPUTED_HEADERS));

  const data = lastRow > 1 ? raw.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  const stat = buildTaxDiagnosticsStats_(data, hm);

  const lines = [];
  lines.push(missing.length ? ('Missing headers: ' + missing.join(', ')) : 'All required headers are present.');
  lines.push('Rows count: ' + data.length);
  lines.push('Found headers (first 40): ' + headers.slice(0, 40).join(' | '));
  lines.push('Order Date min/max: ' + (stat.orderMin || '-') + ' / ' + (stat.orderMax || '-'));
  lines.push('Tax Calculation Date min/max: ' + (stat.taxMin || '-') + ' / ' + (stat.taxMax || '-'));
  lines.push('Unique Tax Rate: ' + stat.taxRates.join(', '));
  lines.push('Unique Tax Collection Responsibility: ' + stat.responsibilities.join(', '));
  lines.push('Top Ship To Country: ' + stat.shipTop.join(', '));

  return lines.join('\n');
}

function buildTaxDiagnosticsStats_(rows, hm) {
  const taxRates = {};
  const resp = {};
  const ship = {};
  let orderMin = null, orderMax = null, taxMin = null, taxMax = null;

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const od = parseAmazonUtcDate_(valueByHeader_(r, hm, 'Order Date'));
    const td = parseAmazonUtcDate_(valueByHeader_(r, hm, 'Tax Calculation Date'));
    if (od) { orderMin = !orderMin || od < orderMin ? od : orderMin; orderMax = !orderMax || od > orderMax ? od : orderMax; }
    if (td) { taxMin = !taxMin || td < taxMin ? td : taxMin; taxMax = !taxMax || td > taxMax ? td : taxMax; }

    taxRates[String(parseNumberFlexible_(valueByHeader_(r, hm, 'Tax Rate')))] = true;
    const rr = String(valueByHeader_(r, hm, 'Tax Collection Responsibility') || '').trim();
    if (rr) resp[rr] = true;
    const c = String(valueByHeader_(r, hm, 'Ship To Country') || '').trim() || '(empty)';
    ship[c] = (ship[c] || 0) + 1;
  }

  const shipTop = Object.keys(ship).sort(function(a, b) { return ship[b] - ship[a]; }).slice(0, 10).map(function(c) { return c + ': ' + ship[c]; });

  return {
    orderMin: orderMin ? Utilities.formatDate(orderMin, CONFIG.TZ, 'yyyy-MM-dd') : '',
    orderMax: orderMax ? Utilities.formatDate(orderMax, CONFIG.TZ, 'yyyy-MM-dd') : '',
    taxMin: taxMin ? Utilities.formatDate(taxMin, CONFIG.TZ, 'yyyy-MM-dd') : '',
    taxMax: taxMax ? Utilities.formatDate(taxMax, CONFIG.TZ, 'yyyy-MM-dd') : '',
    taxRates: Object.keys(taxRates).sort(),
    responsibilities: Object.keys(resp).sort(),
    shipTop: shipTop
  };
}

function applyVatSummaryFormats_(sheet, dataRows, metaMoneyRow) {
  if (!sheet || dataRows <= 0) return;
  safeSetNumberFormat_(sheet.getRange(2, 3, dataRows, 1), '0.00%');
  safeSetNumberFormat_(sheet.getRange(2, 6, dataRows, 1), '0');
  safeSetNumberFormat_(sheet.getRange(2, 7, dataRows, 1), '0');
  safeSetNumberFormat_(sheet.getRange(2, 8, dataRows, 9), '#,##0.00');
  safeSetNumberFormat_(sheet.getRange(metaMoneyRow, 2, 2, 1), '#,##0.00');
}

function getTaxCsvCandidatesFromFolder_(folderId, limit) {
  const folder = DriveApp.getFolderById(folderId);
  const it = folder.getFiles();
  const out = [];
  while (it.hasNext()) {
    const f = it.next();
    const n = String(f.getName() || '').toLowerCase();
    if (n.endsWith('.csv') || n.indexOf('tax') !== -1) {
      out.push({ id: f.getId(), name: f.getName(), updated: f.getLastUpdated() });
    }
  }
  out.sort(function(a, b) { return b.updated.getTime() - a.updated.getTime(); });
  return out.slice(0, Math.max(0, limit || 1));
}

function getOrCreateSheet_(name) {
  const requestedName = String(name || '').trim();
  if (requestedName === CONFIG.LEGACY_MONTHLY_SHEET) {
    logLegacyMonthlyAccess_('getOrCreateSheet_', requestedName);
    return getMonthlyVatPayoutSheet_();
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(requestedName);
  if (!sh) sh = ss.insertSheet(requestedName);
  return sh;
}

function getLegacyMonthlySheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.LEGACY_MONTHLY_SHEET);
}

function getMonthlyVatPayoutSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.MONTHLY_VAT_PAYOUT_SUMMARY_SHEET);
  if (!sh) sh = ss.insertSheet(CONFIG.MONTHLY_VAT_PAYOUT_SUMMARY_SHEET);

  const legacy = getLegacyMonthlySheet_();
  if (legacy && legacy.getSheetId() !== sh.getSheetId()) {
    migrateLegacyMonthlyRowsToMain_(legacy, sh);
    markLegacyMonthlySheetDeprecated_(legacy, sh, null);
  }

  return sh;
}

function migrateLegacyMonthlyRowsToMain_(legacySheet, mainSheet) {
  if (!legacySheet || !mainSheet || legacySheet.getSheetId() === mainSheet.getSheetId()) return 0;
  if (legacySheet.getLastRow() < 2) return 0;

  const legacyLastColumn = Math.max(legacySheet.getLastColumn(), 1);
  const legacyData = legacySheet.getRange(1, 1, legacySheet.getLastRow(), legacyLastColumn).getValues();
  if (legacyData.length < 2) return 0;

  const legacyHeaders = legacyData[0].map(function(v) { return String(v || '').trim(); });
  const legacyHeaderMap = buildHeaderMapCaseInsensitive_(legacyHeaders);
  const mainLastColumn = Math.max(mainSheet.getLastColumn(), 1);
  const mainHasHeaders = mainSheet.getLastRow() >= 1 && mainLastColumn > 0;
  const mainHeaders = mainHasHeaders
    ? mainSheet.getRange(1, 1, 1, mainLastColumn).getValues()[0].map(function(v) { return String(v || '').trim(); })
    : legacyHeaders.slice();

  if (!mainHeaders.some(function(v) { return !!v; })) {
    mainSheet.getRange(1, 1, 1, legacyHeaders.length).setValues([legacyHeaders]);
  }

  const effectiveMainHeaders = mainSheet.getRange(1, 1, 1, Math.max(mainSheet.getLastColumn(), legacyHeaders.length)).getValues()[0]
    .map(function(v) { return String(v || '').trim(); });
  const mainHeaderMap = buildHeaderMapCaseInsensitive_(effectiveMainHeaders);

  const existingKeys = {};
  if (mainSheet.getLastRow() >= 2) {
    const existingRows = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, effectiveMainHeaders.length).getValues();
    for (let i = 0; i < existingRows.length; i++) {
      const key = buildMonthlySheetRowKey_(existingRows[i], effectiveMainHeaders, mainHeaderMap);
      if (key) existingKeys[key] = true;
    }
  }

  const rowsToAppend = [];
  for (let r = 1; r < legacyData.length; r++) {
    const legacyRow = legacyData[r];
    if (isEmptyRow_(legacyRow)) continue;

    const mappedRow = effectiveMainHeaders.map(function(header) {
      const idx = legacyHeaderMap[normalizeHeaderKey_(header)];
      return idx === undefined ? '' : legacyRow[idx];
    });
    const key = buildMonthlySheetRowKey_(mappedRow, effectiveMainHeaders, mainHeaderMap);
    if (!key || existingKeys[key]) continue;

    existingKeys[key] = true;
    rowsToAppend.push(mappedRow);
  }

  if (rowsToAppend.length) {
    mainSheet.getRange(mainSheet.getLastRow() + 1, 1, rowsToAppend.length, effectiveMainHeaders.length).setValues(rowsToAppend);
  }

  return rowsToAppend.length;
}

function buildMonthlySheetRowKey_(row, headers, headerMap) {
  if (!row || !row.length) return '';
  const monthIdx = headerMap[normalizeHeaderKey_('Місяць')];
  const payoutIdx = headerMap[normalizeHeaderKey_('Виплата Amazon')];
  const vatIdx = headerMap[normalizeHeaderKey_('НДС до оплати')];
  const salesIdx = headerMap[normalizeHeaderKey_('Продажі без НДС')];

  const monthValue = monthIdx === undefined ? row[0] : row[monthIdx];
  const monthKey = toMonthText_(monthValue) || String(monthValue || '').trim();
  if (!monthKey) return '';

  const payout = payoutIdx === undefined ? '' : parseNumberFlexible_(row[payoutIdx]);
  const vat = vatIdx === undefined ? '' : parseNumberFlexible_(row[vatIdx]);
  const sales = salesIdx === undefined ? '' : parseNumberFlexible_(row[salesIdx]);
  return [monthKey, payout, vat, sales].join('||');
}

function markLegacyMonthlySheetDeprecated_(legacySheet, mainSheet, migratedRows) {
  if (!legacySheet || !mainSheet || legacySheet.getSheetId() === mainSheet.getSheetId()) return;

  const message = [
    CONFIG.LEGACY_MONTHLY_NOTE,
    'Основна вкладка: ' + CONFIG.MONTHLY_VAT_PAYOUT_SUMMARY_SHEET,
    migratedRows === null || migratedRows === undefined ? '' : ('Міговано рядків: ' + migratedRows)
  ].filter(function(v) { return !!v; }).join('\n');

  try {
    legacySheet.getRange(1, 1).setNote(message);
  } catch (e) {
    Logger.log('[LEGACY MONTH NOTE WARN] ' + toErrorMessage_(e));
  }

  try {
    if (legacySheet.isSheetHidden && !legacySheet.isSheetHidden()) legacySheet.hideSheet();
  } catch (e) {
    Logger.log('[LEGACY MONTH HIDE WARN] ' + toErrorMessage_(e));
  }
}

function logLegacyMonthlyAccess_(source, requestedName) {
  const msg = '[LEGACY MONTHLY SHEET REDIRECT] ' + String(source || 'unknown') + ' requested "' + String(requestedName || '') + '" -> "' + CONFIG.MONTHLY_VAT_PAYOUT_SUMMARY_SHEET + '"';
  Logger.log(msg);
  safeToast_(msg);
}

function buildHeaderMapCaseInsensitive_(headers) {
  const hm = {};
  for (let i = 0; i < headers.length; i++) hm[normalizeHeaderKey_(headers[i])] = i;
  return hm;
}

function normalizeHeaderKey_(s) {
  return String(s || '').trim().toLowerCase();
}

function findMissingHeaders_(hm, required) {
  const missing = [];
  for (let i = 0; i < required.length; i++) {
    if (hm[normalizeHeaderKey_(required[i])] === undefined) missing.push(required[i]);
  }
  return missing;
}



/* =========================
 * EXTENSION: SALES/TAX VAT + PAYOUT MONTHLY SUMMARY (SAFE LAYER)
 * ========================= */

const SALES_TAX_REQUIRED_HEADERS = [
  'Order Date',
  'Tax Calculation Date',
  'Order ID',
  'SKU',
  'Quantity',
  'Tax Rate',
  'Tax Collection Responsibility',
  'Ship To Country',
  'OUR_PRICE Tax Inclusive Selling Price',
  'OUR_PRICE Tax Amount',
  'OUR_PRICE Tax Exclusive Selling Price',
  'OUR_PRICE Tax Inclusive Promo Amount',
  'OUR_PRICE Tax Amount Promo',
  'OUR_PRICE Tax Exclusive Promo Amount',
  'SHIPPING Tax Inclusive Selling Price',
  'SHIPPING Tax Amount',
  'SHIPPING Tax Exclusive Selling Price',
  'SHIPPING Tax Inclusive Promo Amount',
  'SHIPPING Tax Amount Promo',
  'SHIPPING Tax Exclusive Promo Amount',
  'GIFTWRAP Tax Inclusive Selling Price',
  'GIFTWRAP Tax Amount',
  'GIFTWRAP Tax Exclusive Selling Price',
  'GIFTWRAP Tax Inclusive Promo Amount',
  'GIFTWRAP Tax Amount Promo',
  'GIFTWRAP Tax Exclusive Promo Amount'
];

const SALES_TAX_COMPUTED_HEADERS = [
  'Month',
  'Period',
  'Period YYYY-MM',
  'Source Month',
  'Net Product', 'VAT Product', 'Gross Product',
  'Net Shipping', 'VAT Shipping', 'Gross Shipping',
  'Net Giftwrap', 'VAT Giftwrap', 'Gross Giftwrap',
  'Net Sales Total', 'VAT Total', 'Gross Sales Total',
  'Sales Amount', 'VAT Payable', 'Gross Sales',
  'Import File ID', 'Import File Name', 'Imported At',
  'Source File ID', 'Source File Name', 'Source Type',
  'Row Hash'
];

function menuImportAllSalesTaxReports_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = importAllSalesTaxFilesFromFolder_();
    ui.alert([
      'Import all sales/tax reports complete.',
      'Found files: ' + res.total,
      'Imported: ' + res.imported,
      'Skipped: ' + res.skipped,
      'Updated: ' + res.updated,
      'Errors: ' + res.errors.length
    ].join('\n'));
  } catch (e) {
    ui.alert('Import all failed: ' + toErrorMessage_(e));
  }
}

function menuImportOnlyNewSalesTaxReports_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = importOnlyNewSalesTaxFilesFromFolder_();
    ui.alert([
      'Import only new sales/tax reports complete.',
      'Found files: ' + res.total,
      'Imported: ' + res.imported,
      'Skipped: ' + res.skipped,
      'Errors: ' + res.errors.length
    ].join('\n'));
  } catch (e) {
    ui.alert('Import only new failed: ' + toErrorMessage_(e));
  }
}

function menuReimportAllSalesTaxReports_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = reimportAllSalesTaxFilesFromFolder_();
    ui.alert([
      'Reimport all sales/tax reports complete.',
      'Found files: ' + res.total,
      'Imported: ' + res.imported,
      'Updated: ' + res.updated,
      'Errors: ' + res.errors.length,
      'Summary months: ' + (res.summary ? res.summary.months : 0)
    ].join('\n'));
  } catch (e) {
    ui.alert('Reimport failed: ' + toErrorMessage_(e));
  }
}

function menuRebuildMonthlyVatPayoutSummary_() {
  return uiRebuildMonthlyVatPayoutSummary_();
}

function menuRunVatDiagnostics_() {
  return uiRunVatDiagnostics_();
}

function menuRebuildLatestMonthOnly_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = rebuildLatestMonthOnly_();
    ui.alert([
      'Перерахунок останнього місяця завершено.',
      'Місяць: ' + (res.month || '-'),
      'Виплата Amazon: ' + res.paidOut.toFixed(2),
      'Продажі: ' + res.salesAmount.toFixed(2),
      'НДС до сплати: ' + res.vatPayable.toFixed(2)
    ].join('\n'));
  } catch (e) {
    ui.alert('Помилка перерахунку останнього місяця: ' + toErrorMessage_(e));
  }
}

function uiImportLatestSalesTaxReportFromFolder_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = importLatestSalesTaxReportFromFolder_();
    ui.alert('Sales/Tax report imported.\nFile: ' + res.fileName + '\nRows imported: ' + res.rows);
  } catch (e) {
    ui.alert('Import failed: ' + toErrorMessage_(e));
  }
}

function uiRebuildMonthlyVatPayoutSummary_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = rebuildMonthlyVatPayoutSummary_();
    ui.alert('MONTHLY_VAT_PAYOUT_SUMMARY rebuilt.\nMonths: ' + res.months + '\nAmazon Paid Out total: ' + res.totalPayout.toFixed(2));
  } catch (e) {
    ui.alert('Rebuild failed: ' + toErrorMessage_(e));
  }
}

function uiRunVatDiagnostics_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = runVatDiagnostics_();
    ui.alert('VAT diagnostics complete.\nRows in diagnostics: ' + res.rows);
  } catch (e) {
    ui.alert('Diagnostics failed: ' + toErrorMessage_(e));
  }
}

function importLatestSalesTaxReportFromFolder_() {
  const files = getSalesTaxCsvCandidatesFromFolder_(CONFIG.SALES_TAX_REPORT_FOLDER_ID || CONFIG.TAX_REPORT_FOLDER_ID);
  if (!files.length) throw new Error('No sales/tax CSV found in folder.');
  return importSingleSalesTaxFile_(files[0], { mode: 'replace' });
}

function importAllSalesTaxFilesFromFolder_() {
  return importSalesTaxFilesFromFolderByMode_('all');
}

function importOnlyNewSalesTaxFilesFromFolder_() {
  return importSalesTaxFilesFromFolderByMode_('only_new');
}

function reimportAllSalesTaxFilesFromFolder_() {
  const result = importSalesTaxFilesFromFolderByMode_('reimport_all');
  result.summary = rebuildMonthlyVatPayoutSummary_();
  return result;
}

function importSalesTaxFilesFromFolderByMode_(mode) {
  const files = getSalesTaxCsvCandidatesFromFolder_(CONFIG.SALES_TAX_REPORT_FOLDER_ID || CONFIG.TAX_REPORT_FOLDER_ID);
  const importedFileIds = getImportedSalesTaxFileIds_();
  const report = { total: files.length, imported: 0, skipped: 0, updated: 0, errors: [] };

  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    const alreadyImported = importedFileIds[f.id] === true;
    try {
      if (mode === 'only_new' && alreadyImported) {
        report.skipped++;
        continue;
      }
      if (mode === 'all' && alreadyImported) {
        report.skipped++;
        continue;
      }

      if (mode === 'reimport_all' || alreadyImported) {
        const removed = deleteSalesTaxRowsByFileId_(f.id);
        if (removed > 0) report.updated++;
      }

      importSingleSalesTaxFile_(f, { mode: mode });
      report.imported++;
      importedFileIds[f.id] = true;
    } catch (e) {
      report.errors.push(f.name + ': ' + toErrorMessage_(e));
    }
  }

  rebuildMonthlyVatPayoutSummary_();
  return report;
}

function importSingleSalesTaxFile_(file, options) {
  const opts = options || {};
  const fileId = typeof file === 'string' ? file : file.id;
  const driveFile = DriveApp.getFileById(fileId);
  const fileName = (typeof file === 'object' && file.name) ? file.name : driveFile.getName();

  if (opts.mode === 'replace') deleteSalesTaxRowsByFileId_(fileId);

  const parsed = parseTaxReportTable_(driveFile.getBlob().getDataAsString());
  const table = parsed.rows;
  if (!table || table.length < 2) throw new Error('Sales/Tax CSV has no data rows.');

  const sourceHeaders = (table[0] || []).map(function(h) { return String(h || '').trim(); });
  const hm = buildHeaderMapCaseInsensitive_(sourceHeaders);
  const missing = findMissingHeaders_(hm, SALES_TAX_REQUIRED_HEADERS);
  if (missing.length) throw new Error('Missing required headers: ' + missing.join(', '));

  const importedAt = new Date();
  const importedAtText = Utilities.formatDate(importedAt, CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');
  const rows = [];

  for (let i = 1; i < table.length; i++) {
    const srcRow = table[i];
    if (!srcRow || isEmptyRow_(srcRow)) continue;
    const ext = buildSalesTaxComputedColumns_(srcRow, hm, fileId, fileName, importedAtText);
    rows.push(sourceHeaders.map(function(_, idx) { return srcRow[idx] === undefined ? '' : srcRow[idx]; }).concat(ext));
  }

  appendSalesTaxRawRows_(sourceHeaders, rows);
  return { fileId: fileId, fileName: fileName, rows: rows.length };
}

function buildSalesTaxComputedColumns_(row, hm, fileId, fileName, importedAtText) {
  function n(name) { return parseNumberFlexible_(valueByHeader_(row, hm, name)); }

  const period = resolveSalesTaxPeriod_(row, hm);

  const netProduct = n('OUR_PRICE Tax Exclusive Selling Price') - n('OUR_PRICE Tax Exclusive Promo Amount');
  const vatProduct = n('OUR_PRICE Tax Amount') - n('OUR_PRICE Tax Amount Promo');
  const grossProduct = n('OUR_PRICE Tax Inclusive Selling Price') - n('OUR_PRICE Tax Inclusive Promo Amount');

  const netShipping = n('SHIPPING Tax Exclusive Selling Price') - n('SHIPPING Tax Exclusive Promo Amount');
  const vatShipping = n('SHIPPING Tax Amount') - n('SHIPPING Tax Amount Promo');
  const grossShipping = n('SHIPPING Tax Inclusive Selling Price') - n('SHIPPING Tax Inclusive Promo Amount');

  const netGiftwrap = n('GIFTWRAP Tax Exclusive Selling Price') - n('GIFTWRAP Tax Exclusive Promo Amount');
  const vatGiftwrap = n('GIFTWRAP Tax Amount') - n('GIFTWRAP Tax Amount Promo');
  const grossGiftwrap = n('GIFTWRAP Tax Inclusive Selling Price') - n('GIFTWRAP Tax Inclusive Promo Amount');

  const netSalesTotal = netProduct + netShipping + netGiftwrap;
  const vatTotal = vatProduct + vatShipping + vatGiftwrap;
  const grossSalesTotal = grossProduct + grossShipping + grossGiftwrap;

  const rowHash = buildSalesTaxRowHash_(row, hm, period, netSalesTotal, vatTotal);

  return [
    period,
    period,
    period,
    period,
    netProduct, vatProduct, grossProduct,
    netShipping, vatShipping, grossShipping,
    netGiftwrap, vatGiftwrap, grossGiftwrap,
    netSalesTotal, vatTotal, grossSalesTotal,
    netSalesTotal, vatTotal, grossSalesTotal,
    fileId, fileName, importedAtText,
    fileId, fileName, 'Sales/Tax CSV',
    rowHash
  ];
}

function resolveSalesTaxPeriod_(row, hm) {
  const taxDate = parseAmazonUtcDate_(valueByHeader_(row, hm, 'Tax Calculation Date'));
  const orderDate = parseAmazonUtcDate_(valueByHeader_(row, hm, 'Order Date'));
  const d = taxDate || orderDate;
  return d ? Utilities.formatDate(d, CONFIG.TZ, 'yyyy-MM') : '';
}

function buildSalesTaxRowHash_(row, hm, period, netSalesTotal, vatTotal) {
  const keyParts = [
    valueByHeader_(row, hm, 'Order ID'),
    valueByHeader_(row, hm, 'Shipment ID'),
    valueByHeader_(row, hm, 'Transaction ID'),
    valueByHeader_(row, hm, 'SKU'),
    valueByHeader_(row, hm, 'Quantity'),
    valueByHeader_(row, hm, 'Tax Calculation Date'),
    period,
    netSalesTotal.toFixed(6),
    vatTotal.toFixed(6)
  ].map(function(v) { return String(v || '').trim(); }).join('|');

  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, keyParts);
  let out = '';
  for (let i = 0; i < bytes.length; i++) {
    let b = bytes[i];
    if (b < 0) b += 256;
    out += (b < 16 ? '0' : '') + b.toString(16);
  }
  return out;
}

function appendSalesTaxRawRows_(sourceHeaders, rows) {
  const sh = getOrCreateSheet_(CONFIG.SALES_TAX_RAW_SHEET);
  const finalHeaders = sourceHeaders.concat(SALES_TAX_COMPUTED_HEADERS);
  const lastRow = sh.getLastRow();

  if (lastRow < 1) {
    sh.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
  } else {
    const existingHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h) { return String(h || '').trim(); });
    const missing = [];
    for (let i = 0; i < finalHeaders.length; i++) {
      if (existingHeaders.indexOf(finalHeaders[i]) === -1) missing.push(finalHeaders[i]);
    }
    if (missing.length) {
      const merged = existingHeaders.concat(missing);
      const data = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues() : [];
      sh.clearContents();
      sh.getRange(1, 1, 1, merged.length).setValues([merged]);
      if (data.length) {
        const migrated = data.map(function(r) { return realignRowByHeaders_(existingHeaders, r, merged); });
        sh.getRange(2, 1, migrated.length, merged.length).setValues(migrated);
      }
    }
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h) { return String(h || '').trim(); });
  const alignedRows = rows.map(function(r) { return realignRowByHeaders_(finalHeaders, r, headers); });
  if (alignedRows.length) {
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, alignedRows.length, headers.length).setValues(alignedRows);
    applySalesTaxRawFormats_(sh, sh.getLastRow() - 1, sourceHeaders.length, SALES_TAX_COMPUTED_HEADERS.length);
  }
}

function deleteSalesTaxRowsByFileId_(fileId) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SALES_TAX_RAW_SHEET);
  if (!sh || sh.getLastRow() < 2) return 0;

  const all = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const headers = all[0].map(function(h) { return String(h || '').trim(); });
  const hm = buildHeaderMapCaseInsensitive_(headers);
  const col = hm[normalizeHeaderKey_('Import File ID')];
  if (col === undefined) return 0;

  const kept = [];
  let removed = 0;
  for (let i = 1; i < all.length; i++) {
    const fid = String(all[i][col] || '').trim();
    if (fid && fid === String(fileId)) {
      removed++;
      continue;
    }
    kept.push(all[i]);
  }

  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (kept.length) sh.getRange(2, 1, kept.length, headers.length).setValues(kept);
  return removed;
}

function getImportedSalesTaxFileIds_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SALES_TAX_RAW_SHEET);
  if (!sh || sh.getLastRow() < 2) return {};

  const all = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const headers = all[0].map(function(h) { return String(h || '').trim(); });
  const hm = buildHeaderMapCaseInsensitive_(headers);
  const col = hm[normalizeHeaderKey_('Import File ID')] !== undefined
    ? hm[normalizeHeaderKey_('Import File ID')]
    : hm[normalizeHeaderKey_('Source File ID')];
  if (col === undefined) return {};

  const out = {};
  for (let i = 1; i < all.length; i++) {
    const fid = String(all[i][col] || '').trim();
    if (fid) out[fid] = true;
  }
  return out;
}

function getSalesTaxCsvCandidatesFromFolder_(folderId) {
  if (!folderId) throw new Error('CONFIG.SALES_TAX_REPORT_FOLDER_ID is empty.');
  const files = getTaxCsvCandidatesFromFolder_(folderId, Number.MAX_SAFE_INTEGER);
  files.sort(function(a, b) { return a.updated.getTime() - b.updated.getTime(); });
  return files;
}

function importSalesTaxReportCsvFile_(fileId, fileNameHint) {
  return importSingleSalesTaxFile_({ id: fileId, name: fileNameHint || '' }, { mode: 'replace' });
}

function upsertSalesTaxRawRows_(sourceHeaders, newRows, fileId) {
  deleteSalesTaxRowsByFileId_(fileId);
  appendSalesTaxRawRows_(sourceHeaders, newRows);
}

function realignRowByHeaders_(existingHeaders, row, targetHeaders) {
  const hm = buildHeaderMapCaseInsensitive_(existingHeaders);
  const out = new Array(targetHeaders.length);
  for (let i = 0; i < targetHeaders.length; i++) {
    const idx = hm[normalizeHeaderKey_(targetHeaders[i])];
    out[i] = idx === undefined ? '' : (row[idx] === undefined ? '' : row[idx]);
  }
  return out;
}

function applySalesTaxRawFormats_(sheet, rowCount, sourceColsCount, computedColsCount) {
  if (!sheet || rowCount <= 0) return;
  safeSetNumberFormat_(sheet.getRange(2, 1, rowCount, sourceColsCount), '@', [], 'salesTax.raw.sourceText');
  if (computedColsCount > 0) {
    safeSetNumberFormat_(sheet.getRange(2, sourceColsCount + 1, rowCount, 3), '@', [], 'salesTax.raw.month');
    safeSetNumberFormat_(sheet.getRange(2, sourceColsCount + 4, rowCount, 15), '#,##0.00', [], 'salesTax.raw.money');
  }
}


function rebuildLatestMonthOnly_() {
  const payoutByMonth = buildSettlementPayoutByMonth_();
  const salesAgg = buildSalesTaxMonthlyAgg_();
  const months = Object.keys(salesAgg.byMonth || {}).sort();
  if (!months.length) throw new Error('У SALES_TAX_RAW немає даних для розрахунку останнього місяця.');

  const month = months[months.length - 1];
  const p = payoutByMonth[month] || { paidOut: 0, fileIds: {} };
  const s = salesAgg.byMonth[month] || { salesAmount: 0, vatPayable: 0, fileIds: {}, rows: 0 };

  const paidOut = p.paidOut;
  const salesAmount = s.salesAmount;
  const vatPayable = s.vatPayable;
  const remaining = paidOut - vatPayable;
  const settlementCount = Object.keys(p.fileIds || {}).length;
  const salesFileCount = Object.keys(s.fileIds || {}).length;
  const notes = buildMonthlyVatPayoutNote_(paidOut, salesAmount, vatPayable, settlementCount, salesFileCount, s.rows);

  const sh = getMonthlyVatPayoutSheet_();
  const headers = [
    'Місяць',
    'Виплата Amazon',
    'Продажі',
    'НДС до сплати',
    'Залишок після НДС',
    'Кількість settlement файлів',
    'Кількість sales файлів',
    'Кількість sales рядків',
    'Нотатки / Діагностика'
  ];

  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(2, 1, 1, headers.length).setValues([[month, paidOut, salesAmount, vatPayable, remaining, settlementCount, salesFileCount, s.rows, notes]]);
  safeSetNumberFormat_(sh.getRange(2, 1, 1, 1), '@', [], 'monthly.last.month');
  safeSetNumberFormat_(sh.getRange(2, 2, 1, 4), '#,##0.00', [], 'monthly.last.money');
  safeSetNumberFormat_(sh.getRange(2, 6, 1, 3), '0', [], 'monthly.last.counts');

  return {
    month: month,
    paidOut: paidOut,
    salesAmount: salesAmount,
    vatPayable: vatPayable,
    remaining: remaining,
    settlementCount: settlementCount,
    salesFileCount: salesFileCount
  };
}

function rebuildMonthlyVatPayoutSummary_() {
  const payoutByMonth = buildSettlementPayoutByMonth_();
  const salesAgg = buildSalesTaxMonthlyAgg_();
  const months = mergeMonthKeys_(Object.keys(payoutByMonth), Object.keys(salesAgg.byMonth));

  const headers = [
    'Місяць',
    'Виплата Amazon',
    'Продажі',
    'НДС до сплати',
    'Залишок після НДС',
    'Кількість settlement файлів',
    'Кількість sales файлів',
    'Кількість sales рядків',
    'Нотатки / Діагностика'
  ];

  const rows = [];
  let totalPayout = 0;
  for (let i = 0; i < months.length; i++) {
    const m = months[i];
    const p = payoutByMonth[m] || { paidOut: 0, fileIds: {} };
    const s = salesAgg.byMonth[m] || { salesAmount: 0, vatPayable: 0, fileIds: {}, rows: 0 };

    const paidOut = p.paidOut;
    const salesAmount = s.salesAmount;
    const vatPayable = s.vatPayable;
    const remaining = paidOut - vatPayable;
    totalPayout += paidOut;

    const settlementCount = Object.keys(p.fileIds || {}).length;
    const salesFileCount = Object.keys(s.fileIds || {}).length;
    const notes = buildMonthlyVatPayoutNote_(paidOut, salesAmount, vatPayable, settlementCount, salesFileCount, s.rows);

    rows.push([m, paidOut, salesAmount, vatPayable, remaining, settlementCount, salesFileCount, s.rows, notes]);
  }

  const sh = getMonthlyVatPayoutSheet_();
  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    safeSetNumberFormat_(sh.getRange(2, 1, rows.length, 1), '@', [], 'vatPayout.month');
    safeSetNumberFormat_(sh.getRange(2, 2, rows.length, 4), '#,##0.00', [], 'vatPayout.money');
    safeSetNumberFormat_(sh.getRange(2, 6, rows.length, 3), '0', [], 'vatPayout.counts');
  }

  return { months: rows.length, totalPayout: totalPayout };
}

function buildSettlementPayoutByMonth_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SUMMARY_SHEET);
  if (!sh) throw new Error('Summary sheet not found: ' + CONFIG.SUMMARY_SHEET);
  if (sh.getLastRow() < 2) return {};

  const all = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const headers = all[0].map(function(h) { return String(h || '').trim(); });
  const hm = buildHeaderMapCaseInsensitive_(headers);

  const colMonth = hm[normalizeHeaderKey_(CONFIG.HEADERS.month)] + 1 || 0;
  const colPosted = hm[normalizeHeaderKey_(CONFIG.HEADERS.postedDate)] + 1 || 0;
  const colDeposit = hm[normalizeHeaderKey_(CONFIG.HEADERS.depositDate)] + 1 || 0;
  const colTransfer = hm[normalizeHeaderKey_(CONFIG.HEADERS.transfer)] + 1 || 0;
  const colFileId = hm[normalizeHeaderKey_(CONFIG.HEADERS.fileId)] + 1 || 0;

  if (!colTransfer) throw new Error('Transfer column not found in ' + CONFIG.SUMMARY_SHEET);

  const out = {};
  for (let i = 1; i < all.length; i++) {
    const r = all[i];
    const fid = String(colFileId ? r[colFileId - 1] : '').trim();
    if (!fid || fid === CONFIG.TOTAL_FILE_ID) continue;

    const month = monthFromSummaryRow_(r, colMonth, colPosted, colDeposit);
    if (!month) continue;

    const transfer = parseNumberFlexible_(r[colTransfer - 1]);
    if (!out[month]) out[month] = { paidOut: 0, fileIds: {} };
    out[month].paidOut += transfer;
    if (fid) out[month].fileIds[fid] = true;
  }
  return out;
}

function monthFromSummaryRow_(row, colMonth, colPosted, colDeposit) {
  const monthVal = colMonth ? row[colMonth - 1] : '';
  const monthText = toMonthText_(monthVal);
  if (monthText) return monthText;

  const postedVal = colPosted ? row[colPosted - 1] : '';
  const postedMonth = toMonthText_(postedVal);
  if (postedMonth) return postedMonth;

  const depVal = colDeposit ? row[colDeposit - 1] : '';
  return toMonthText_(depVal);
}

function toMonthText_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, CONFIG.TZ, 'yyyy-MM');
  const s = String(value || '').trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}$/.test(s)) return s;
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 7);
  const d = parseAmazonUtcDate_(s);
  if (d) return Utilities.formatDate(d, CONFIG.TZ, 'yyyy-MM');
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, CONFIG.TZ, 'yyyy-MM');
  return '';
}

function buildSalesTaxMonthlyAgg_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SALES_TAX_RAW_SHEET);
  if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < 1) return { byMonth: {} };

  const all = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const headers = all[0].map(function(h) { return String(h || '').trim(); });
  const hm = buildHeaderMapCaseInsensitive_(headers);

  const required = ['Net Sales Total', 'VAT Total', 'Import File ID'];
  const missing = findMissingHeaders_(hm, required);
  if (missing.length) throw new Error('Missing columns in ' + CONFIG.SALES_TAX_RAW_SHEET + ': ' + missing.join(', '));

  const out = {};
  for (let i = 1; i < all.length; i++) {
    const r = all[i];
    const month = resolveSalesTaxAggMonth_(r, hm);
    if (!month) continue;

    if (!out[month]) out[month] = { salesAmount: 0, vatPayable: 0, fileIds: {}, rows: 0 };
    out[month].salesAmount += parseNumberFlexible_(valueByHeader_(r, hm, 'Net Sales Total') || valueByHeader_(r, hm, 'Sales Amount'));
    out[month].vatPayable += parseNumberFlexible_(valueByHeader_(r, hm, 'VAT Total') || valueByHeader_(r, hm, 'VAT Payable'));
    out[month].rows += 1;

    const fid = String(valueByHeader_(r, hm, 'Import File ID') || valueByHeader_(r, hm, 'Source File ID') || '').trim();
    if (fid) out[month].fileIds[fid] = true;
  }

  return { byMonth: out };
}

function resolveSalesTaxAggMonth_(row, hm) {
  const period = toMonthText_(valueByHeader_(row, hm, 'Period YYYY-MM') || valueByHeader_(row, hm, 'Period') || valueByHeader_(row, hm, 'Month'));
  if (period) return period;

  const taxDate = toMonthText_(valueByHeader_(row, hm, 'Tax Calculation Date'));
  if (taxDate) return taxDate;

  return toMonthText_(valueByHeader_(row, hm, 'Order Date'));
}

function mergeMonthKeys_(a, b) {
  const set = {};
  for (let i = 0; i < a.length; i++) if (a[i]) set[a[i]] = true;
  for (let j = 0; j < b.length; j++) if (b[j]) set[b[j]] = true;
  return Object.keys(set).sort();
}

function buildMonthlyVatPayoutNote_(paidOut, salesAmount, vatPayable, settlementCount, salesFileCount, salesRowsCount) {
  const notes = [];
  if (settlementCount === 0) notes.push('Немає даних по виплатах settlement');
  if (salesFileCount === 0) notes.push('Немає даних зі звітів продажів');
  if (salesRowsCount > 0 && salesAmount === 0 && vatPayable === 0) notes.push('Є рядки продажів, але підсумки дорівнюють нулю');
  return notes.join('; ');
}

function runVatDiagnosticsLegacy_() {
  const diagnosticsRows = writeDiagnosticsLegacy_();
  return { rows: diagnosticsRows.length };
}

function writeDiagnosticsLegacy_() {
  const diagnosticsRows = [];
  const salesSheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SALES_TAX_RAW_SHEET);
  if (!salesSheet || salesSheet.getLastRow() < 2) {
    diagnosticsRows.push(['ERROR', 'General', 'Missing raw sheet or no data', CONFIG.SALES_TAX_RAW_SHEET]);
  } else {
    const all = salesSheet.getRange(1, 1, salesSheet.getLastRow(), salesSheet.getLastColumn()).getValues();
    const headers = all[0].map(function(h) { return String(h || '').trim(); });
    const hm = buildHeaderMapCaseInsensitive_(headers);
    const missing = findMissingHeaders_(hm, SALES_TAX_REQUIRED_HEADERS.concat(SALES_TAX_COMPUTED_HEADERS));

    diagnosticsRows.push(['INFO', 'General', 'Рядків у SALES_TAX_RAW', String(all.length - 1)]);
    diagnosticsRows.push(['INFO', 'General', 'Missing headers', missing.join(', ')]);

    const importedRegistry = {};
    const monthStats = {};
    const rowHashSeen = {};
    const comboMap = {};

    for (let i = 1; i < all.length; i++) {
      const r = all[i];
      const fid = String(valueByHeader_(r, hm, 'Import File ID') || valueByHeader_(r, hm, 'Source File ID') || '').trim() || '(empty)';
      const fname = String(valueByHeader_(r, hm, 'Import File Name') || valueByHeader_(r, hm, 'Source File Name') || '').trim();
      const importedAt = String(valueByHeader_(r, hm, 'Imported At') || '').trim();
      const period = String(valueByHeader_(r, hm, 'Period') || valueByHeader_(r, hm, 'Month') || '').trim() || '(empty)';
      const orderId = String(valueByHeader_(r, hm, 'Order ID') || '').trim();
      const transactionId = String(valueByHeader_(r, hm, 'Transaction ID') || '').trim();
      const shipmentId = String(valueByHeader_(r, hm, 'Shipment ID') || '').trim();
      const taxRate = String(valueByHeader_(r, hm, 'Tax Rate') || '').trim();
      const shipTo = String(valueByHeader_(r, hm, 'Ship To Country') || '').trim() || '(empty)';
      const resp = String(valueByHeader_(r, hm, 'Tax Collection Responsibility') || '').trim() || '(empty)';
      const rowHash = String(valueByHeader_(r, hm, 'Row Hash') || '').trim();

      if (!importedRegistry[fid]) importedRegistry[fid] = { fileName: fname, rows: 0, firstMonth: period, lastMonth: period, importedAtSet: {} };
      const reg = importedRegistry[fid];
      reg.rows += 1;
      if (importedAt) reg.importedAtSet[importedAt] = true;
      if (period && (reg.firstMonth === '(empty)' || period < reg.firstMonth)) reg.firstMonth = period;
      if (period && period > reg.lastMonth) reg.lastMonth = period;

      if (!monthStats[period]) monthStats[period] = { rows: 0, files: {}, orders: {}, taxRates: {}, countries: {}, responsibilities: {} };
      const ms = monthStats[period];
      ms.rows += 1;
      ms.files[fid] = true;
      if (orderId) ms.orders[orderId] = true;
      if (taxRate) ms.taxRates[taxRate] = true;
      ms.countries[shipTo] = true;
      ms.responsibilities[resp] = true;

      if (rowHash) {
        if (!rowHashSeen[rowHash]) rowHashSeen[rowHash] = 0;
        rowHashSeen[rowHash] += 1;
      }

      const combo = [orderId, transactionId, shipmentId].join('|');
      if (combo !== '||') {
        if (!comboMap[combo]) comboMap[combo] = { files: {}, count: 0 };
        comboMap[combo].files[fid] = true;
        comboMap[combo].count += 1;
      }
    }

    diagnosticsRows.push(['INFO', 'Imported files registry', 'File ID', 'File Name', 'Rows Imported', 'First Month', 'Last Month', 'Imported At', 'Duplicate Status / Notes']);
    const fileIds = Object.keys(importedRegistry).sort();
    for (let j = 0; j < fileIds.length; j++) {
      const fid = fileIds[j];
      const reg = importedRegistry[fid];
      const importEvents = Object.keys(reg.importedAtSet).sort();
      const duplicateStatus = importEvents.length > 1 ? 'same File ID imported multiple times' : '';
      diagnosticsRows.push(['DATA', 'Imported files registry', fid, reg.fileName, reg.rows, reg.firstMonth, reg.lastMonth, importEvents.join(' | '), duplicateStatus]);
    }

    diagnosticsRows.push(['INFO', 'By month diagnostics', 'Month', 'Raw rows count', 'File count', 'Order count', 'Tax rates', 'Ship-to countries', 'Tax collection responsibilities']);
    const months = Object.keys(monthStats).sort();
    for (let k = 0; k < months.length; k++) {
      const month = months[k];
      const ms = monthStats[month];
      diagnosticsRows.push([
        'DATA',
        'By month diagnostics',
        month,
        ms.rows,
        Object.keys(ms.files).length,
        Object.keys(ms.orders).length,
        Object.keys(ms.taxRates).sort().join(', '),
        Object.keys(ms.countries).sort().join(', '),
        Object.keys(ms.responsibilities).sort().join(', ')
      ]);
    }

    let rowHashDuplicates = 0;
    const rowHashes = Object.keys(rowHashSeen);
    for (let x = 0; x < rowHashes.length; x++) if (rowHashSeen[rowHashes[x]] > 1) rowHashDuplicates += 1;

    let crossFileComboDuplicates = 0;
    const combos = Object.keys(comboMap);
    for (let y = 0; y < combos.length; y++) {
      if (Object.keys(comboMap[combos[y]].files).length > 1) crossFileComboDuplicates += 1;
    }

    let fileIdReimportWarnings = 0;
    for (let z = 0; z < fileIds.length; z++) {
      const reg = importedRegistry[fileIds[z]];
      if (Object.keys(reg.importedAtSet).length > 1) fileIdReimportWarnings += 1;
    }

    diagnosticsRows.push(['WARN', 'Duplicate detection', 'File IDs imported multiple times', String(fileIdReimportWarnings)]);
    diagnosticsRows.push(['WARN', 'Duplicate detection', 'Duplicate row hashes', String(rowHashDuplicates)]);
    diagnosticsRows.push(['WARN', 'Duplicate detection', 'Cross-file duplicate order/transaction/shipment combos', String(crossFileComboDuplicates)]);
  }

  const summarySheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.MONTHLY_VAT_PAYOUT_SUMMARY_SHEET);
  diagnosticsRows.push(['INFO', 'General', 'Рядків у МІСЯЧНИЙ_ЗВІТ', String(summarySheet ? Math.max(0, summarySheet.getLastRow() - 1) : 0)]);

  const sh = getOrCreateSheet_(CONFIG.DIAGNOSTICS_SHEET);
  sh.clearContents();
  sh.getRange(1, 1, 1, 9).setValues([['Level', 'Section', 'Metric/Col1', 'Value/Col2', 'Col3', 'Col4', 'Col5', 'Col6', 'Col7']]);
  if (diagnosticsRows.length) {
    const norm = diagnosticsRows.map(function(r) {
      const out = new Array(9).fill('');
      for (let i = 0; i < Math.min(9, r.length); i++) out[i] = r[i];
      return out;
    });
    sh.getRange(2, 1, norm.length, 9).setValues(norm);
  }
  return diagnosticsRows;
}

function valueByHeader_(row, hm, header) {
  const idx = hm[normalizeHeaderKey_(header)];
  return idx === undefined ? '' : row[idx];
}

function parseNumberFlexible_(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  let s = String(value).trim();
  if (!s) return 0;
  s = s.replace(/\s/g, '');
  const hasComma = s.indexOf(',') >= 0;
  const hasDot = s.indexOf('.') >= 0;
  if (hasComma && hasDot) {
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.');
    else s = s.replace(/,/g, '');
  } else if (hasComma) {
    s = s.replace(',', '.');
  }
  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function parseAmazonUtcDate_(s) {
  if (!s && s !== 0) return null;
  let v = String(s).trim();
  if (!v) return null;
  v = v.replace(/\s+UTC$/i, '').trim();
  const m = v.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
  if (!m) return null;
  const day = Number(m[1]);
  const monTxt = m[2].toLowerCase();
  const year = Number(m[3]);
  const mons = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
  if (mons[monTxt] === undefined) return null;
  return new Date(Date.UTC(year, mons[monTxt], day));
}

function isEmptyRow_(row) {
  for (let i = 0; i < row.length; i++) {
    if (String(row[i] || '').trim() !== '') return false;
  }
  return true;
}

/* =========================
 * UI/REPORT EXTENSION: РУЧНІ ВИТРАТИ + МІСЯЧНИЙ ЗВІТ (БЕЗ ЗМІНИ COGS)
 * ========================= */

function onOpen() {
  ensureUkrainianLabels_();
  const menu = getLocalizedMenuConfig_();
  const uiMenu = SpreadsheetApp.getUi().createMenu(menu.title);
  for (let i = 0; i < menu.items.length; i++) {
    const item = menu.items[i];
    if (item.separator) uiMenu.addSeparator();
    else uiMenu.addItem(item.label, item.functionName);
  }
  uiMenu.addToUi();
}

function ensureUkrainianLabels_() {
  return {
    menuTitle: 'Фінанси Amazon',
    manualExpensesSheet: getManualExpensesSheetName_(),
    monthlySheet: CONFIG.MONTHLY_SHEET,
    diagnosticsSheet: CONFIG.DIAGNOSTICS_SHEET
  };
}

function getLocalizedMenuConfig_() {
  return {
    title: 'Фінанси Amazon',
    items: [
      { label: 'Імпортувати всі settlement', functionName: 'menuImportAllSettlements_' },
      { label: 'Імпортувати останній settlement', functionName: 'menuImportLatestSettlement_' },
      { label: 'Імпортувати всі звіти продажів', functionName: 'menuImportAllSalesTaxReports_' },
      { label: 'Імпортувати останній звіт продажів', functionName: 'menuImportLatestSalesTaxReport_' },
      { separator: true },
      { label: 'Ініціалізувати лист ручних витрат', functionName: 'menuEnsureManualExpensesSheet_' },
      { label: 'Підготувати лист оборотного капіталу', functionName: 'menuEnsureWorkingCapitalSheet_' },
      { label: 'Оновити лист введення НДС', functionName: 'menuEnsureManualVatSheet_' },
      { label: 'Оновити лист введення комісій', functionName: 'menuEnsureManualFeesSheet_' },
      { separator: true },
      { label: 'Перерахувати весь місячний звіт', functionName: 'menuRebuildMonthlyReport_' },
      { label: 'Перерахувати останній місяць', functionName: 'menuRebuildLatestMonthOnly_' },
      { label: 'Діагностика', functionName: 'menuRunDiagnostics_' },
      { separator: true },
      { label: 'Оновити дешборд', functionName: 'menuRebuildDashboard_' }
    ]
  };
}

function menuImportAllSettlements_() { return uiImportAllFromFolder_(); }
function menuImportLatestSettlement_() { return uiImportLatestFromFolder_(); }

function menuImportLatestSalesTaxReport_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = importLatestSalesTaxFileFromFolder_();
    ui.alert(['Імпорт останнього звіту продажів завершено.', 'Файл: ' + res.fileName, 'Рядків імпортовано: ' + res.rows].join('\n'));
  } catch (e) {
    ui.alert('Помилка імпорту останнього звіту продажів: ' + toErrorMessage_(e));
  }
}

function menuImportAllSalesTaxReports_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = importAllSalesTaxFilesFromFolder_();
    ui.alert([
      'Імпорт усіх звітів продажів завершено.',
      'Знайдено файлів: ' + res.total,
      'Імпортовано: ' + res.imported,
      'Пропущено як вже імпортовані: ' + res.skipped,
      'Помилок: ' + res.errors.length
    ].join('\n'));
  } catch (e) {
    ui.alert('Помилка імпорту звітів продажів: ' + toErrorMessage_(e));
  }
}

function menuEnsureManualExpensesSheet_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = ensureManualExpensesSheet_();
    ui.alert([
      'Лист ручних витрат готовий.',
      'Лист: ' + res.sheetName,
      'Створено: ' + (res.created ? 'так' : 'ні'),
      'Додано заголовків: ' + res.addedHeaders
    ].join('\n'));
  } catch (e) {
    ui.alert('Помилка підготовки листа ручних витрат: ' + toErrorMessage_(e));
  }
}

function menuEnsureManualOperationsSheet_() {
  return menuEnsureManualExpensesSheet_();
}
function menuEnsureWorkingCapitalSheet_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = ensureWorkingCapitalSheet_();
    ui.alert([
      'Лист оборотного капіталу готовий.',
      'Лист: ' + res.sheetName,
      'Створено: ' + (res.created ? 'так' : 'ні'),
      'Додано заголовків: ' + res.addedHeaders
    ].join('\n'));
  } catch (e) {
    ui.alert('Помилка підготовки листа оборотного капіталу: ' + toErrorMessage_(e));
  }
}


function menuSyncManualPurchases_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = syncManualPurchasesToZakupky_();
    ui.alert([
      'Синхронізацію в "Закупки" вимкнено для безпечної міграції.',
      'Опрацьовано рядків: ' + res.processed,
      'Створено: ' + res.inserted,
      'Оновлено: ' + res.updated,
      'Пропущено: ' + res.skipped,
      'Коментар: ' + res.message
    ].join('\n'));
  } catch (e) {
    ui.alert('Помилка безпечної перевірки синхронізації: ' + toErrorMessage_(e));
  }
}

function menuEnsureManualVatSheet_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = ensureManualVatSheet_();
    ui.alert(['Лист ручного введення НДС готовий.', 'Лист: ' + res.sheetName, 'Створено: ' + (res.created ? 'так' : 'ні')].join('\n'));
  } catch (e) {
    ui.alert('Помилка підготовки листа НДС: ' + toErrorMessage_(e));
  }
}

function menuEnsureManualFeesSheet_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = ensureManualFeesSheet_();
    ui.alert(['Лист ручного введення комісій готовий.', 'Лист: ' + res.sheetName, 'Створено: ' + (res.created ? 'так' : 'ні')].join('\n'));
  } catch (e) {
    ui.alert('Помилка підготовки листа комісій: ' + toErrorMessage_(e));
  }
}

function menuRebuildMonthlyReport_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = rebuildMonthlyVatPayoutSummary_();
    ui.alert(['Місячний звіт перераховано.', 'Місяців: ' + res.months, 'Сума виплат Amazon: ' + res.totalPayout.toFixed(2)].join('\n'));
  } catch (e) {
    ui.alert('Помилка перерахунку місячного звіту: ' + toErrorMessage_(e));
  }
}

function menuRebuildLatestMonthOnly_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = rebuildLatestMonthOnly_();
    ui.alert(['Останній місяць перераховано.', 'Місяць: ' + res.month, 'Виплата: ' + res.paidOut.toFixed(2), 'НДС до оплати: ' + res.vatToPay.toFixed(2)].join('\n'));
  } catch (e) {
    ui.alert('Помилка перерахунку останнього місяця: ' + toErrorMessage_(e));
  }
}

function menuRebuildDashboard_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = rebuildDashboard_();
    if (res.empty) {
      ui.alert('Дешборд оновлено. Даних у ' + CONFIG.MONTHLY_SHEET + ' поки немає.');
      return;
    }
    ui.alert(['Дешборд оновлено.', 'Місяців: ' + res.months, 'Останній місяць: ' + res.latestMonth].join('\n'));
  } catch (e) {
    ui.alert('Помилка оновлення дешборду: ' + toErrorMessage_(e));
  }
}

function menuRunDiagnostics_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = runVatDiagnostics_();
    ui.alert('Діагностику оновлено. Рядків: ' + res.rows);
  } catch (e) {
    ui.alert('Помилка діагностики: ' + toErrorMessage_(e));
  }
}

function getManualExpensesSheetName_() {
  return CONFIG.MANUAL_EXPENSES_SHEET || 'РУЧНІ_ВИТРАТИ';
}

function getManualExpenseTypes_() {
  return [
    'Бізнес-витрата',
    'Закуп у постачальника',
    'Пакування',
    'Сервіси та підписки',
    'Логістика',
    'Інше'
  ];
}

function getManualExpensePaymentMethods_() {
  return ['Картка', 'Банк', 'Готівка', 'PayPal', 'Інше'];
}

function getManualExpenseDocumentTypes_() {
  return ['Рахунок', 'Фактура', 'Чек', 'Без документа'];
}

function getManualExpenseFundCategories_() {
  return (CONFIG.MANUAL_EXPENSE_FUND_CATEGORIES || [
    'Реінвест (75%)',
    'Бізнес витрати (12%)',
    'Зарплата (7%)',
    'Інше (6%)'
  ]).slice();
}

function getManualExpensesHeaders_() {
  return [
    'ID',
    'Тип витрати',
    'Дата',
    'Місяць',
    'Категорія',
    'Тип документа',
    'Сума без НДС',
    'Сума НДС',
    'Сума з НДС',
    'Опис',
    'Враховувати у прибутку',
    'Активно',
    'Категорія коштів'
  ];
}

function getManualExpensesReadOnlyHeaders_() {
  return getManualExpensesHeaders_().slice();
}

function getManualOperationsHeaders_() {
  return getManualExpensesHeaders_();
}

function ensureManualExpensesSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(getManualExpensesSheetName_());
  const created = !sh;
  if (sh) return { sheetName: getManualExpensesSheetName_(), created: false, addedHeaders: 0, readOnlyProtected: true };
  sh = ss.insertSheet(getManualExpensesSheetName_());
  const headers = getManualExpensesHeaders_();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  return { sheetName: getManualExpensesSheetName_(), created: created, addedHeaders: headers.length, readOnlyProtected: true };
}

function ensureManualOperationsSheet_() {
  return ensureManualExpensesSheet_();
}

function getWorkingCapitalSheetName_() {
  return CONFIG.WORKING_CAPITAL_SHEET || 'ОБОРОТНИЙ_КАПІТАЛ';
}

function getWorkingCapitalHeaders_() {
  return [
    'ID',
    'Дата',
    'Місяць',
    'Собівартість товарів на Amazon',
    'Кошти доступні на вивід Amazon',
    'Товари готові до відправки',
    'Кошти на руках',
    'Можливість кредитування',
    'Оборотний капітал',
    'Примітки',
    'Активно'
  ];
}

function ensureWorkingCapitalSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(getWorkingCapitalSheetName_());
  const created = !sh;
  if (!sh) sh = ss.insertSheet(getWorkingCapitalSheetName_());

  const headers = getWorkingCapitalHeaders_();
  let addedHeaders = 0;
  const maxCols = Math.max(sh.getLastColumn() || 1, headers.length);
  const row1 = sh.getRange(1, 1, 1, maxCols).getValues()[0].map(function(v) { return String(v || '').trim(); });

  if (sh.getLastRow() === 0 || row1.join('') === '') {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    addedHeaders = headers.length;
  } else {
    let appendAt = sh.getLastColumn();
    for (let i = 0; i < headers.length; i++) {
      if (row1.indexOf(headers[i]) === -1) {
        appendAt += 1;
        sh.getRange(1, appendAt).setValue(headers[i]);
        addedHeaders += 1;
      }
    }
  }

  const hm = getHeaderMap_(sh);
  if (created) {
    const dataRows = Math.max(1, sh.getMaxRows() - 1);
    if (hm['Активно']) sh.getRange(2, hm['Активно'], dataRows, 1).insertCheckboxes();
    if (hm['Дата']) safeSetNumberFormat_(sh.getRange(2, hm['Дата'], dataRows, 1), 'yyyy-mm-dd', [], 'workingCapital.date');
    if (hm['Місяць']) {
      safeSetNumberFormat_(sh.getRange(2, hm['Місяць'], dataRows, 1), '@', [], 'workingCapital.month');
      sh.getRange(1, hm['Місяць']).setNote('Рекомендований формат: yyyy-MM. Перевірка виконується під час читання даних.');
    }

    [
      'Собівартість товарів на Amazon',
      'Кошти доступні на вивід Amazon',
      'Товари готові до відправки',
      'Кошти на руках',
      'Можливість кредитування',
      'Оборотний капітал'
    ].forEach(function(header) {
      if (hm[header]) safeSetNumberFormat_(sh.getRange(2, hm[header], dataRows, 1), '€#,##0.00', [], 'workingCapital.' + header);
    });
  }

  return { sheetName: getWorkingCapitalSheetName_(), created: created, addedHeaders: addedHeaders };
}

function normalizeWorkingCapitalMonth_(monthValue, dateValue) {
  const direct = String(monthValue || '').trim();
  if (/^\d{4}-\d{2}$/.test(direct)) return direct;

  const fromDate = parseDateFlexible_(dateValue, CONFIG.TZ, {});
  if (fromDate instanceof Date && !isNaN(fromDate.getTime())) {
    return Utilities.formatDate(fromDate, CONFIG.TZ, 'yyyy-MM');
  }

  const normalized = toMonthText_(monthValue);
  if (normalized && /^\d{4}-\d{2}$/.test(normalized)) return normalized;
  return '';
}

function readWorkingCapitalRows_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(getWorkingCapitalSheetName_());
  if (!sh || sh.getLastRow() < 2) return { rows: [], warnings: sh ? [] : ['Лист ' + getWorkingCapitalSheetName_() + ' не знайдено.'] };

  const all = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const hm = buildHeaderMapCaseInsensitive_(all[0]);

  const required = [
    'Дата',
    'Місяць',
    'Собівартість товарів на Amazon',
    'Кошти доступні на вивід Amazon',
    'Товари готові до відправки',
    'Кошти на руках',
    'Можливість кредитування'
  ];
  const missing = [];
  for (let i = 0; i < required.length; i++) {
    if (hm[normalizeHeaderKey_(required[i])] === undefined) missing.push(required[i]);
  }
  if (missing.length) throw new Error('У ' + getWorkingCapitalSheetName_() + ' відсутні колонки: ' + missing.join(', '));

  function val(row, header) {
    const idx = hm[normalizeHeaderKey_(header)];
    return idx === undefined ? '' : row[idx];
  }

  const rows = [];
  const warnings = [];
  const monthStats = {};
  let invalidMonths = 0;
  let missingMonths = 0;
  let nonNumericValues = 0;

  function parseWorkingCapitalNumber_(rawValue, rowIndex, headerName) {
    if (rawValue === null || rawValue === undefined || rawValue === '') return 0;
    if (typeof rawValue === 'number') return isFinite(rawValue) ? rawValue : 0;
    const parsed = parseNumberFlexible_(rawValue);
    if (parsed === 0) {
      const normalized = String(rawValue).replace(/\s/g, '').replace(',', '.');
      if (normalized !== '0' && normalized !== '0.0' && normalized !== '0.00') {
        nonNumericValues += 1;
        warnings.push('ОБОРОТНИЙ_КАПІТАЛ рядок ' + rowIndex + ': нечислове значення у колонці "' + headerName + '" (' + rawValue + ').');
      }
    }
    return parsed;
  }

  for (let i = 1; i < all.length; i++) {
    const raw = all[i];
    if (isEmptyRow_(raw)) continue;

    const active = manualExpenseBoolean_(val(raw, 'Активно'), true);
    const month = normalizeWorkingCapitalMonth_(val(raw, 'Місяць'), val(raw, 'Дата'));
    const cogsAmazon = parseWorkingCapitalNumber_(val(raw, 'Собівартість товарів на Amazon'), i + 1, 'Собівартість товарів на Amazon');
    const amazonPayoutAvailable = parseWorkingCapitalNumber_(val(raw, 'Кошти доступні на вивід Amazon'), i + 1, 'Кошти доступні на вивід Amazon');
    const goodsReadyToShip = parseWorkingCapitalNumber_(val(raw, 'Товари готові до відправки'), i + 1, 'Товари готові до відправки');
    const cashOnHand = parseWorkingCapitalNumber_(val(raw, 'Кошти на руках'), i + 1, 'Кошти на руках');
    const creditCapacity = parseWorkingCapitalNumber_(val(raw, 'Можливість кредитування'), i + 1, 'Можливість кредитування');
    const total = roundMoney_(cogsAmazon + amazonPayoutAvailable + goodsReadyToShip + cashOnHand + creditCapacity);

    if (active && !month) {
      const rawMonth = String(val(raw, 'Місяць') || '').trim();
      if (!rawMonth) missingMonths += 1;
      else invalidMonths += 1;
      warnings.push('ОБОРОТНИЙ_КАПІТАЛ рядок ' + (i + 1) + ': невалідний місяць (очікується yyyy-MM).');
    }

    rows.push({
      rowNumber: i + 1,
      id: String(val(raw, 'ID') || '').trim(),
      date: val(raw, 'Дата'),
      month: month,
      cogsAmazon: cogsAmazon,
      amazonPayoutAvailable: amazonPayoutAvailable,
      goodsReadyToShip: goodsReadyToShip,
      cashOnHand: cashOnHand,
      creditCapacity: creditCapacity,
      workingCapital: total,
      note: val(raw, 'Примітки'),
      active: active
    });

    if (active && month) {
      if (!monthStats[month]) monthStats[month] = { rows: 0, total: 0 };
      monthStats[month].rows += 1;
      monthStats[month].total = roundMoney_(monthStats[month].total + total);
    }
  }

  return {
    rows: rows,
    warnings: warnings,
    stats: {
      invalidMonths: invalidMonths,
      missingMonths: missingMonths,
      nonNumericValues: nonNumericValues,
      totalsByMonth: monthStats
    }
  };
}

function collectWorkingCapitalByMonth_() {
  const parsed = readWorkingCapitalRows_();
  const byMonth = {};
  let activeRows = 0;

  for (let i = 0; i < parsed.rows.length; i++) {
    const row = parsed.rows[i];
    if (!row.active || !row.month) continue;
    activeRows += 1;

    if (!byMonth[row.month]) {
      byMonth[row.month] = {
        month: row.month,
        cogsAmazon: 0,
        amazonPayoutAvailable: 0,
        goodsReadyToShip: 0,
        cashOnHand: 0,
        creditCapacity: 0,
        workingCapital: 0,
        rows: 0
      };
    }

    byMonth[row.month].cogsAmazon = roundMoney_(byMonth[row.month].cogsAmazon + row.cogsAmazon);
    byMonth[row.month].amazonPayoutAvailable = roundMoney_(byMonth[row.month].amazonPayoutAvailable + row.amazonPayoutAvailable);
    byMonth[row.month].goodsReadyToShip = roundMoney_(byMonth[row.month].goodsReadyToShip + row.goodsReadyToShip);
    byMonth[row.month].cashOnHand = roundMoney_(byMonth[row.month].cashOnHand + row.cashOnHand);
    byMonth[row.month].creditCapacity = roundMoney_(byMonth[row.month].creditCapacity + row.creditCapacity);
    byMonth[row.month].workingCapital = roundMoney_(byMonth[row.month].workingCapital + row.workingCapital);
    byMonth[row.month].rows += 1;
  }

  const months = Object.keys(byMonth).sort();
  const list = months.map(function(month) { return byMonth[month]; });
  return {
    byMonth: byMonth,
    months: months,
    rows: list,
    activeRows: activeRows,
    warnings: parsed.warnings || [],
    diagnostics: parsed.stats || { invalidMonths: 0, missingMonths: 0, nonNumericValues: 0, totalsByMonth: {} }
  };
}

function getLatestWorkingCapitalMonth_(agg) {
  const months = (agg && agg.months ? agg.months : Object.keys((agg && agg.byMonth) || {})).slice().sort();
  return months.length ? months[months.length - 1] : '';
}

function getWorkingCapitalTotals_(agg) {
  const rows = (agg && agg.rows) ? agg.rows : [];
  const latestMonth = getLatestWorkingCapitalMonth_(agg);
  const latest = latestMonth ? (agg.byMonth || {})[latestMonth] : null;
  const previous = latestMonth && agg.months && agg.months.length > 1 ? agg.byMonth[agg.months[agg.months.length - 2]] : null;

  let total = 0;
  for (let i = 0; i < rows.length; i++) total += rows[i].workingCapital;

  const average = rows.length ? roundMoney_(total / rows.length) : 0;
  const current = latest ? roundMoney_(latest.workingCapital) : 0;
  const changeVsPrev = roundMoney_(current - (previous ? previous.workingCapital : 0));

  return {
    latestMonth: latestMonth,
    currentWorkingCapital: current,
    averageWorkingCapital: average,
    changeVsPreviousMonth: previous ? changeVsPrev : current,
    previousMonth: previous ? previous.month : '',
    latestBreakdown: latest || null
  };
}

function rowArrayToObject_(row, headers) {
  const out = {};
  for (let i = 0; i < headers.length; i++) out[headers[i]] = row[i];
  return out;
}

function normalizeManualExpenseType_(value) {
  const raw = String(value || '').trim().toLowerCase();
  if (!raw) return '';

  const known = getManualExpenseTypes_();
  for (let i = 0; i < known.length; i++) {
    if (raw === known[i].toLowerCase()) return known[i];
  }
  if (raw.indexOf('закуп') !== -1 || raw.indexOf('постачаль') !== -1 || raw.indexOf('supplier') !== -1) return 'Закуп у постачальника';
  if (raw.indexOf('пакув') !== -1) return 'Пакування';
  if (raw.indexOf('сервіс') !== -1 || raw.indexOf('підпис') !== -1 || raw.indexOf('subscription') !== -1) return 'Сервіси та підписки';
  if (raw.indexOf('логіст') !== -1 || raw.indexOf('достав') !== -1 || raw.indexOf('transport') !== -1) return 'Логістика';
  if (raw.indexOf('бізнес') !== -1 || raw.indexOf('витрат') !== -1 || raw.indexOf('expense') !== -1) return 'Бізнес-витрата';
  if (raw.indexOf('інше') !== -1 || raw.indexOf('other') !== -1) return 'Інше';
  return String(value || '').trim();
}

function normalizeManualExpenseFundCategory_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const known = getManualExpenseFundCategories_();
  const normalizedRaw = raw.toLowerCase();
  for (let i = 0; i < known.length; i++) {
    if (known[i].toLowerCase() === normalizedRaw) return known[i];
  }

  if (normalizedRaw.indexOf('реінвест') !== -1 || normalizedRaw.indexOf('реинвест') !== -1 || normalizedRaw.indexOf('reinvest') !== -1) return 'Реінвест (75%)';
  if (normalizedRaw.indexOf('бізнес') !== -1 || normalizedRaw.indexOf('бизнес') !== -1 || normalizedRaw.indexOf('business') !== -1) return 'Бізнес витрати (12%)';
  if (normalizedRaw.indexOf('зарплат') !== -1 || normalizedRaw.indexOf('salary') !== -1 || normalizedRaw.indexOf('payroll') !== -1) return 'Зарплата (7%)';
  if (normalizedRaw.indexOf('інше') !== -1 || normalizedRaw.indexOf('иное') !== -1 || normalizedRaw.indexOf('other') !== -1) return 'Інше (6%)';
  return raw;
}

function isKnownManualExpenseFundCategory_(value) {
  if (!String(value || '').trim()) return false;
  const normalized = normalizeManualExpenseFundCategory_(value);
  const known = getManualExpenseFundCategories_();
  for (let i = 0; i < known.length; i++) {
    if (known[i] === normalized) return true;
  }
  return false;
}

function normalizeManualOperationType_(value) {
  return normalizeManualExpenseType_(value);
}

function manualExpenseBoolean_(value, defaultValue) {
  if (value === true || value === false) return value;
  const raw = String(value === null || value === undefined ? '' : value).trim().toLowerCase();
  if (!raw) return defaultValue;
  if (['true', '1', 'так', 'yes', 'y'].indexOf(raw) !== -1) return true;
  if (['false', '0', 'ні', 'no', 'n'].indexOf(raw) !== -1) return false;
  return defaultValue;
}

function isManualOperationActive_(value) {
  return manualExpenseBoolean_(value, true);
}

function isManualExpenseIncludedInProfit_(value) {
  return manualExpenseBoolean_(value, false);
}

function roundMoney_(value) {
  const num = parseNumberFlexible_(value);
  return Math.round(num * 100) / 100;
}

function deriveManualExpenseMonthKey_(value) {
  const parsedDate = parseDateFlexible_(value, CONFIG.TZ, {});
  if (parsedDate instanceof Date && !isNaN(parsedDate.getTime())) {
    return Utilities.formatDate(parsedDate, CONFIG.TZ, 'yyyy-MM');
  }
  return toMonthText_(value);
}

function normalizeManualExpenseRow_(row, headerMap, rowNumber) {
  const headers = getManualExpensesReadOnlyHeaders_();
  const out = headers.map(function(header) {
    const col = headerMap[header];
    return col ? row[col - 1] : '';
  });
  const warnings = [];

  function idx(header) { return headers.indexOf(header); }
  function get(header) {
    const index = idx(header);
    return index === -1 ? '' : out[index];
  }
  function set(header, value) {
    const index = idx(header);
    if (index !== -1) out[index] = value;
  }

  const originalDate = get('Дата');
  const parsedDate = parseDateFlexible_(originalDate, CONFIG.TZ, {});
  if (parsedDate instanceof Date && !isNaN(parsedDate.getTime())) {
    set('Місяць', Utilities.formatDate(parsedDate, CONFIG.TZ, 'yyyy-MM'));
  } else {
    const monthKey = deriveManualExpenseMonthKey_(get('Місяць'));
    if (monthKey) set('Місяць', monthKey);
    if (String(originalDate || '').trim() !== '') warnings.push('Рядок ' + rowNumber + ': невалідна дата.');
  }

  if (get('Активно') === '' || get('Активно') === null) set('Активно', true);
  set('Тип витрати', normalizeManualExpenseType_(get('Тип витрати')));
  set('Категорія коштів', normalizeManualExpenseFundCategory_(get('Категорія коштів')));

  const netAmount = parseNumberFlexible_(get('Сума без НДС'));
  const vatAmount = parseNumberFlexible_(get('Сума НДС'));
  const grossAmount = parseNumberFlexible_(get('Сума з НДС'));
  if (!grossAmount && (netAmount || vatAmount)) set('Сума з НДС', roundMoney_(netAmount + vatAmount));

  const effectiveAmount = netAmount || parseNumberFlexible_(get('Сума з НДС'));
  if (!String(get('Тип витрати') || '').trim()) warnings.push('Рядок ' + rowNumber + ': не вказано тип витрати.');
  if (!effectiveAmount) warnings.push('Рядок ' + rowNumber + ': не вказано суму без НДС або суму з НДС.');

  return { row: out, warnings: warnings };
}

function normalizeManualOperationRow_(row, headerMap, rowNumber) {
  return normalizeManualExpenseRow_(row, headerMap, rowNumber);
}

function getManualExpenseSheetHeaderMap_() {
  return getManualExpensesHeaderMapReadOnly_();
}

function getManualExpensesHeaderMapReadOnly_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(getManualExpensesSheetName_());
  if (!sh) return { sheet: null, headerMap: {} };
  const width = sh.getLastColumn();
  if (width < 1) return { sheet: sh, headerMap: {} };
  const headerRow = sh.getRange(1, 1, 1, width).getValues()[0];
  const headerMap = {};
  for (let i = 0; i < headerRow.length; i++) {
    const header = String(headerRow[i] || '').trim();
    if (header && headerMap[header] === undefined) headerMap[header] = i + 1;
  }
  return { sheet: sh, headerMap: headerMap };
}

function readManualExpensesRowsReadOnly_() {
  const sheetInfo = getManualExpensesHeaderMapReadOnly_();
  const sh = sheetInfo.sheet;
  const hm = sheetInfo.headerMap;
  const headers = getManualExpensesReadOnlyHeaders_();
  if (!sh || sh.getLastRow() < 2) return { sheet: sh, headerMap: hm, rows: [] };

  const width = sh.getLastColumn();
  const all = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();
  const objects = [];
  for (let i = 0; i < all.length; i++) {
    if (isEmptyRow_(all[i])) continue;
    const normalized = normalizeManualExpenseRow_(all[i], hm, i + 2);
    const obj = rowArrayToObject_(normalized.row, headers);
    obj.__rowNumber = i + 2;
    obj.__warnings = normalized.warnings;
    obj.__active = isManualOperationActive_(obj['Активно']);
    obj.__includeInProfit = isManualExpenseIncludedInProfit_(obj['Враховувати у прибутку']);
    obj.__month = deriveManualExpenseMonthKey_(obj['Місяць']) || deriveManualExpenseMonthKey_(obj['Дата']) || '';
    obj.__netAmount = parseNumberFlexible_(obj['Сума без НДС']);
    obj.__vatAmount = parseNumberFlexible_(obj['Сума НДС']);
    obj.__grossAmount = parseNumberFlexible_(obj['Сума з НДС']);
    obj.__effectiveAmount = obj.__netAmount || obj.__grossAmount || 0;
    obj.__fundCategory = normalizeManualExpenseFundCategory_(obj['Категорія коштів']);
    obj.__hasKnownFundCategory = isKnownManualExpenseFundCategory_(obj.__fundCategory);
    objects.push(obj);
  }

  return { sheet: sh, headerMap: hm, rows: objects };
}

function getManualExpensesData_() {
  return readManualExpensesRowsReadOnly_();
}

function getManualOperationsData_() {
  return getManualExpensesData_();
}

function readManualExpensesRowsForDashboard_() {
  return readManualExpensesRowsReadOnly_().rows;
}

function collectManualExpenseCategoryTotalsReadOnly_() {
  const rows = readManualExpensesRowsForDashboard_();
  const totalsByCategory = buildManualExpenseFundCategoryTotalsSkeleton_();
  const byMonth = {};
  const warnings = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row.__active || !row.__month || !row.__effectiveAmount) continue;
    const category = readManualExpenseFundCategory_(row);
    if (!category || !isKnownManualExpenseFundCategory_(category)) continue;

    if (!byMonth[row.__month]) byMonth[row.__month] = { total: 0, categories: buildManualExpenseFundCategoryTotalsSkeleton_(), rows: 0 };
    byMonth[row.__month].categories[category] = roundMoney_((byMonth[row.__month].categories[category] || 0) + row.__effectiveAmount);
    byMonth[row.__month].total = roundMoney_(byMonth[row.__month].total + row.__effectiveAmount);
    byMonth[row.__month].rows += 1;
    totalsByCategory[category] = roundMoney_((totalsByCategory[category] || 0) + row.__effectiveAmount);
  }

  const months = Object.keys(byMonth).sort();
  for (let m = 0; m < months.length; m++) {
    const monthKey = months[m];
    const monthBucket = byMonth[monthKey];
    const monthTotal = monthBucket.total || 0;
    monthBucket.shares = {};
    const categories = getManualExpenseFundCategories_();
    for (let c = 0; c < categories.length; c++) {
      const cat = categories[c];
      monthBucket.shares[cat] = monthTotal > 0 ? roundMoney_((monthBucket.categories[cat] || 0) / monthTotal * 100) : 0;
    }
  }

  return { byMonth: byMonth, totalsByCategory: totalsByCategory, warnings: warnings };
}

function collectManualExpenseCategoryTotalsForDashboard_() {
  return collectManualExpenseCategoryTotalsReadOnly_();
}


function validateManualExpenseRows_() {
  const data = getManualExpensesData_();
  const stats = {
    totalRows: data.rows.length,
    activeRows: 0,
    includedInProfitRows: 0,
    inactiveRows: 0,
    invalidDateRows: [],
    missingAmountRows: [],
    missingTypeRows: [],
    supplierPurchaseRows: 0,
    businessExpenseRows: 0,
    missingFundCategoryRows: [],
    unknownFundCategoryRows: [],
    totalsByMonth: {},
    vatTotalsByMonth: {},
    totalsByFundCategory: buildManualExpenseFundCategoryTotalsSkeleton_(),
    warnings: []
  };

  for (let i = 0; i < data.rows.length; i++) {
    const row = data.rows[i];
    if (!row.__active) {
      stats.inactiveRows += 1;
      continue;
    }

    stats.activeRows += 1;
    if (!row.__month) stats.invalidDateRows.push(row.__rowNumber);
    if (!row['Тип витрати']) stats.missingTypeRows.push(row.__rowNumber);
    if (!row.__effectiveAmount) stats.missingAmountRows.push(row.__rowNumber);
    if (row['Тип витрати'] === 'Закуп у постачальника') stats.supplierPurchaseRows += 1;
    if (row['Тип витрати'] === 'Бізнес-витрата') stats.businessExpenseRows += 1;
    if (!String(row.__fundCategory || '').trim()) stats.missingFundCategoryRows.push(row.__rowNumber);
    else if (!row.__hasKnownFundCategory) stats.unknownFundCategoryRows.push(row.__rowNumber);
    else stats.totalsByFundCategory[row.__fundCategory] = roundMoney_((stats.totalsByFundCategory[row.__fundCategory] || 0) + row.__effectiveAmount);

    if (row.__includeInProfit && row.__month && row.__effectiveAmount) {
      stats.includedInProfitRows += 1;
      stats.totalsByMonth[row.__month] = roundMoney_((stats.totalsByMonth[row.__month] || 0) + row.__effectiveAmount);
      stats.vatTotalsByMonth[row.__month] = roundMoney_((stats.vatTotalsByMonth[row.__month] || 0) + row.__vatAmount);
    }

    if (row.__warnings && row.__warnings.length) stats.warnings = stats.warnings.concat(row.__warnings);
  }

  return stats;
}

function validateManualOperationRows_() {
  return validateManualExpenseRows_();
}

function collectManualExpensesByMonth_() {
  const data = getManualExpensesData_();
  const out = {};
  const warnings = [];

  for (let i = 0; i < data.rows.length; i++) {
    const row = data.rows[i];
    if (!row.__active) continue;
    if (!row.__month) {
      warnings.push('Рядок ' + row.__rowNumber + ' пропущено: невалідна дата або місяць.');
      continue;
    }
    if (!row.__effectiveAmount) {
      warnings.push('Рядок ' + row.__rowNumber + ' пропущено: не вказано суму без НДС або суму з НДС.');
      continue;
    }
    if (!out[row.__month]) {
      out[row.__month] = {
        amount: 0,
        vat: 0,
        rows: 0,
        includedRows: 0,
        allActiveRows: 0,
        types: {}
      };
    }

    out[row.__month].allActiveRows += 1;
    const typeKey = String(row['Тип витрати'] || 'Без типу');
    out[row.__month].types[typeKey] = (out[row.__month].types[typeKey] || 0) + 1;

    if (!row.__includeInProfit) continue;

    out[row.__month].amount = roundMoney_(out[row.__month].amount + row.__effectiveAmount);
    out[row.__month].vat = roundMoney_(out[row.__month].vat + row.__vatAmount);
    out[row.__month].rows += 1;
    out[row.__month].includedRows += 1;
  }

  const legacy = readLegacyBusinessExpensesByMonth_();
  const legacyMonths = Object.keys(legacy.byMonth || {});
  for (let j = 0; j < legacyMonths.length; j++) {
    const legacyMonth = legacyMonths[j];
    if (!out[legacyMonth]) {
      out[legacyMonth] = {
        amount: 0,
        vat: 0,
        rows: 0,
        includedRows: 0,
        allActiveRows: 0,
        types: {}
      };
    }
    out[legacyMonth].amount = roundMoney_(out[legacyMonth].amount + legacy.byMonth[legacyMonth].amount);
    out[legacyMonth].vat = roundMoney_(out[legacyMonth].vat + legacy.byMonth[legacyMonth].vat);
    out[legacyMonth].rows += legacy.byMonth[legacyMonth].rows;
    out[legacyMonth].includedRows += legacy.byMonth[legacyMonth].rows;
    out[legacyMonth].allActiveRows += legacy.byMonth[legacyMonth].rows;
    out[legacyMonth].types['Застарілий лист витрат'] = (out[legacyMonth].types['Застарілий лист витрат'] || 0) + legacy.byMonth[legacyMonth].rows;
  }

  return { byMonth: out, warnings: warnings.concat(legacy.warnings || []) };
}

function readManualExpenseFundCategory_(rowObject) {
  if (!rowObject) return '';
  return normalizeManualExpenseFundCategory_(rowObject['Категорія коштів'] || rowObject.__fundCategory || '');
}

function buildManualExpenseFundCategoryTotalsSkeleton_() {
  const out = {};
  const categories = getManualExpenseFundCategories_();
  for (let i = 0; i < categories.length; i++) out[categories[i]] = 0;
  return out;
}

function collectManualExpensesByFundCategoryByMonth_() {
  const data = getManualExpensesData_();
  const byMonth = {};
  const totalsByCategory = buildManualExpenseFundCategoryTotalsSkeleton_();
  const warnings = [];

  for (let i = 0; i < data.rows.length; i++) {
    const row = data.rows[i];
    if (!row.__active) continue;
    if (!row.__month || !row.__effectiveAmount) continue;

    const category = readManualExpenseFundCategory_(row);
    if (!category || !isKnownManualExpenseFundCategory_(category)) continue;

    if (!byMonth[row.__month]) byMonth[row.__month] = { total: 0, categories: buildManualExpenseFundCategoryTotalsSkeleton_(), rows: 0 };
    byMonth[row.__month].categories[category] = roundMoney_((byMonth[row.__month].categories[category] || 0) + row.__effectiveAmount);
    byMonth[row.__month].total = roundMoney_(byMonth[row.__month].total + row.__effectiveAmount);
    byMonth[row.__month].rows += 1;
    totalsByCategory[category] = roundMoney_((totalsByCategory[category] || 0) + row.__effectiveAmount);
  }

  const months = Object.keys(byMonth).sort();
  for (let m = 0; m < months.length; m++) {
    const monthKey = months[m];
    const monthBucket = byMonth[monthKey];
    const monthTotal = monthBucket.total || 0;
    monthBucket.shares = {};
    const categories = getManualExpenseFundCategories_();
    for (let c = 0; c < categories.length; c++) {
      const cat = categories[c];
      monthBucket.shares[cat] = monthTotal > 0 ? roundMoney_((monthBucket.categories[cat] || 0) / monthTotal * 100) : 0;
    }
  }

  return { byMonth: byMonth, totalsByCategory: totalsByCategory, warnings: warnings };
}

function collectManualOperationsByMonth_() {
  return collectManualExpensesByMonth_();
}

function validateManualExpenseVatRows_() {
  return collectManualExpenseVatByMonthReadOnly_().stats;
}

function collectManualExpenseVatByMonthReadOnly_() {
  const data = getManualExpensesData_();
  const out = {};
  const warnings = [];
  const stats = {
    totalRows: data.rows.length,
    activeRows: 0,
    includedRows: 0,
    skippedInactive: 0,
    skippedInvalidDate: 0,
    skippedMissingVat: 0,
    totalsByMonth: {}
  };

  for (let i = 0; i < data.rows.length; i++) {
    const row = data.rows[i];
    if (!row.__active) {
      stats.skippedInactive += 1;
      continue;
    }
    stats.activeRows += 1;

    if (!row.__month) {
      stats.skippedInvalidDate += 1;
      warnings.push('РУЧНІ_ВИТРАТИ рядок ' + row.__rowNumber + ': пропущено НДС через невалідну дату або місяць.');
      continue;
    }

    const vat = parseNumberFlexible_(row['Сума НДС']);
    if (!(vat > 0)) {
      stats.skippedMissingVat += 1;
      continue;
    }

    if (!out[row.__month]) out[row.__month] = { vat: 0, rows: 0 };
    out[row.__month].vat = roundMoney_(out[row.__month].vat + vat);
    out[row.__month].rows += 1;

    stats.includedRows += 1;
    stats.totalsByMonth[row.__month] = roundMoney_((stats.totalsByMonth[row.__month] || 0) + vat);
  }

  return { byMonth: out, warnings: warnings, stats: stats };
}

function collectManualExpenseVatByMonth_() {
  return collectManualExpenseVatByMonthReadOnly_();
}


function getManualExpenseVatTotalForMonth_(month, manualExpenseVatByMonth) {
  return roundMoney_(((manualExpenseVatByMonth || {})[month] || { vat: 0 }).vat || 0);
}

function getPaidVatBreakdownForMonth_(month, manualVatByMonth, manualExpenseVatByMonth) {
  const manualInputVat = roundMoney_(((manualVatByMonth || {})[month] || { paidVat: 0 }).paidVat || 0);
  const manualExpenseVat = getManualExpenseVatTotalForMonth_(month, manualExpenseVatByMonth || {});
  const totalPaidVat = roundMoney_(manualInputVat + manualExpenseVat);
  return {
    manualInputVat: manualInputVat,
    manualExpenseVat: manualExpenseVat,
    totalPaidVat: totalPaidVat
  };
}

function rebuildPaidVatSection_(row, paidVatBreakdown) {
  const breakdown = paidVatBreakdown || { manualInputVat: 0, manualExpenseVat: 0, totalPaidVat: 0 };
  row[8] = roundMoney_(breakdown.manualInputVat);
  row[9] = roundMoney_(breakdown.manualExpenseVat);
  row[10] = roundMoney_(breakdown.totalPaidVat);
  row[11] = roundMoney_(breakdown.totalPaidVat);
  return row;
}

function readLegacyBusinessExpensesByMonth_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.LEGACY_BUSINESS_EXPENSES_SHEET);
  if (!sh || sh.getLastRow() < 2) return { byMonth: {}, warnings: [] };

  const all = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const hm = buildHeaderMapCaseInsensitive_(all[0]);
  const out = {};
  for (let i = 1; i < all.length; i++) {
    const row = all[i];
    const month = toMonthText_(valueByHeader_(row, hm, 'Місяць'));
    if (!month) continue;
    if (!out[month]) out[month] = { amount: 0, vat: 0, rows: 0 };
    out[month].amount += parseNumberFlexible_(valueByHeader_(row, hm, 'Сума без НДС')) || parseNumberFlexible_(valueByHeader_(row, hm, 'Сума')) || 0;
    out[month].vat += parseNumberFlexible_(valueByHeader_(row, hm, 'Сума НДС')) || parseNumberFlexible_(valueByHeader_(row, hm, 'НДС')) || 0;
    out[month].rows += 1;
  }
  return { byMonth: out, warnings: ['Використано застарілий лист "' + CONFIG.LEGACY_BUSINESS_EXPENSES_SHEET + '" як резервне джерело ручних витрат.'] };
}

function collectBusinessExpensesByMonth_() {
  return collectManualExpensesByMonth_();
}

function collectPurchaseRowsForSync_() {
  const data = getManualExpensesData_();
  const rows = [];
  for (let i = 0; i < data.rows.length; i++) {
    const row = data.rows[i];
    if (!row.__active) continue;
    if (row['Тип витрати'] === 'Закуп у постачальника') rows.push(row);
  }
  return {
    rows: rows,
    warnings: rows.length ? ['Синхронізацію з листом "Закупки" вимкнено: агреговані закупи не можна використовувати як SKU-level COGS.'] : []
  };
}

function ensurePurchasesSyncHeaders_(sh) {
  return getHeaderMap_(sh);
}

function syncManualPurchasesToZakupky_() {
  const preview = collectPurchaseRowsForSync_();
  return {
    processed: preview.rows.length,
    inserted: 0,
    updated: 0,
    skipped: preview.rows.length,
    errors: [],
    message: 'Лист "РУЧНІ_ВИТРАТИ" не синхронізується з "Закупки". Існуюча COGS логіка залишена без змін.'
  };
}

function applyManualExpensesToMonthlyReport_(row, month, manualExpensesByMonth) {
  const profitBeforeVat = parseNumberFlexible_(row[7]);
  const paidVat = parseNumberFlexible_(row[11]);
  const vatToPay = roundMoney_(row[3] - paidVat);

  row[12] = vatToPay;
  row[13] = roundMoney_(profitBeforeVat - vatToPay);
  return row;
}

function applyBusinessExpensesToMonthlyReport_(row, month, businessExpensesByMonth) {
  return applyManualExpensesToMonthlyReport_(row, month, businessExpensesByMonth);
}

function importLatestSalesTaxFileFromFolder_() {
  const folderId = CONFIG.SALES_TAX_REPORT_FOLDER_ID || CONFIG.TAX_REPORT_FOLDER_ID;
  const files = getSalesTaxCsvCandidatesFromFolder_(folderId);
  if (!files.length) throw new Error('Не знайдено CSV звітів продажів у папці.');

  const latest = files[0];
  const imported = getImportedSalesTaxFileIds_();
  if (imported[latest.id] === true) {
    return { fileId: latest.id, fileName: latest.name, rows: 0, skipped: true };
  }

  const res = importSingleSalesTaxFile_(latest, { mode: 'all' });
  rebuildMonthlyVatPayoutSummary_();
  return res;
}

function ensureManualVatSheet_() {
  return ensureInputSheetWithHeaders_('ВВЕДЕННЯ_НДС', ['Місяць', 'Сплачений НДС', 'Коментар'], 'manualVat');
}

function ensureManualFeesSheet_() {
  return ensureInputSheetWithHeaders_('ВВЕДЕННЯ_КОМІСІЙ', ['Місяць', 'Комісії Amazon', 'Коментар'], 'manualFees');
}

function ensureInputSheetWithHeaders_(sheetName, requiredHeaders, formatPrefix) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(sheetName);
  const created = !sh;
  if (!sh) sh = ss.insertSheet(sheetName);

  const maxCols = Math.max(requiredHeaders.length, sh.getLastColumn() || 1);
  const row1 = sh.getRange(1, 1, 1, maxCols).getValues()[0];
  const normalized = row1.map(function(h) { return String(h || '').trim(); });

  if (sh.getLastRow() === 0 || normalized.join('') === '') {
    sh.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
  } else {
    let appendAt = normalized.length;
    for (let i = 0; i < requiredHeaders.length; i++) {
      const header = requiredHeaders[i];
      if (normalized.indexOf(header) === -1) {
        appendAt += 1;
        sh.getRange(1, appendAt).setValue(header);
      }
    }
  }

  const dataRows = Math.max(0, sh.getLastRow() - 1);
  if (dataRows > 0) {
    safeSetNumberFormat_(sh.getRange(2, 1, dataRows, 1), '@', [], formatPrefix + '.month');
    safeSetNumberFormat_(sh.getRange(2, 2, dataRows, 1), '#,##0.00', [], formatPrefix + '.money');
  }

  return { sheetName: sheetName, created: created };
}

function readManualVatByMonth_() {
  ensureManualVatSheet_();
  const sh = SpreadsheetApp.getActive().getSheetByName('ВВЕДЕННЯ_НДС');
  if (!sh || sh.getLastRow() < 2) return { byMonth: {}, warnings: [] };

  const all = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const hm = buildHeaderMapCaseInsensitive_(all[0]);
  const out = {};
  const warnings = [];

  for (let i = 1; i < all.length; i++) {
    const r = all[i];
    const month = toMonthText_(valueByHeader_(r, hm, 'Місяць'));
    if (!month) continue;
    const raw = valueByHeader_(r, hm, 'Сплачений НДС');
    const paidVat = parseNumberFlexible_(raw);
    if (!out[month]) out[month] = { paidVat: 0, rows: 0 };
    out[month].paidVat += paidVat;
    out[month].rows += 1;
    if (String(raw === null || raw === undefined ? '' : raw).trim() !== '' && !isFinite(Number(String(raw).replace(',', '.')))) {
      warnings.push('ВВЕДЕННЯ_НДС row ' + (i + 1) + ': невалідне число "' + raw + '" -> 0');
    }
  }

  return { byMonth: out, warnings: warnings };
}

function readManualFeesByMonth_() {
  ensureManualFeesSheet_();
  const sh = SpreadsheetApp.getActive().getSheetByName('ВВЕДЕННЯ_КОМІСІЙ');
  if (!sh || sh.getLastRow() < 2) return { byMonth: {}, warnings: [] };

  const all = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const hm = buildHeaderMapCaseInsensitive_(all[0]);
  const out = {};
  const warnings = [];

  for (let i = 1; i < all.length; i++) {
    const r = all[i];
    const month = toMonthText_(valueByHeader_(r, hm, 'Місяць'));
    if (!month) continue;
    const rawMain = valueByHeader_(r, hm, 'Комісії Amazon');
    const rawGross = valueByHeader_(r, hm, 'Комісії Amazon з НДС');
    const rawExVat = valueByHeader_(r, hm, 'Комісії Amazon без НДС');
    const rawVat = valueByHeader_(r, hm, 'НДС на комісії Amazon');
    const hasMain = String(rawMain === null || rawMain === undefined ? '' : rawMain).trim() !== '';
    const hasGross = String(rawGross === null || rawGross === undefined ? '' : rawGross).trim() !== '';
    const manualFees = hasMain
      ? parseNumberFlexible_(rawMain)
      : (hasGross ? parseNumberFlexible_(rawGross) : (parseNumberFlexible_(rawExVat) + parseNumberFlexible_(rawVat)));

    if (!out[month]) out[month] = { fees: 0, rows: 0 };
    out[month].fees += manualFees;
    out[month].rows += 1;

    const rawForValidation = hasMain ? rawMain : (hasGross ? rawGross : '');
    if (String(rawForValidation === null || rawForValidation === undefined ? '' : rawForValidation).trim() !== '' && !isFinite(Number(String(rawForValidation).replace(',', '.')))) {
      warnings.push('ВВЕДЕННЯ_КОМІСІЙ row ' + (i + 1) + ': невалідне число "' + rawForValidation + '" -> 0');
    }
  }

  return { byMonth: out, warnings: warnings };
}

function rebuildLatestMonthOnly_() {

  const payoutByMonth = buildSettlementPayoutByMonth_();
  const settlementFees = buildSettlementFeesByMonth_();
  const cogsByMonth = buildSettlementCogsByMonth_();
  const salesAgg = buildSalesTaxMonthlyAgg_().byMonth || {};
  const manualVat = readManualVatByMonth_();
  const manualFees = readManualFeesByMonth_();
  const manualExpenseVat = collectManualExpenseVatByMonthReadOnly_();

  const months = mergeMonthKeys_(
    mergeMonthKeys_(Object.keys(payoutByMonth), Object.keys(salesAgg)),
    mergeMonthKeys_(Object.keys(cogsByMonth), mergeMonthKeys_(Object.keys(manualVat.byMonth), mergeMonthKeys_(Object.keys(manualFees.byMonth), Object.keys(manualExpenseVat.byMonth))))
  );
  if (!months.length) throw new Error('Немає даних для перерахунку останнього місяця.');

  const month = months[months.length - 1];
  const row = buildMonthlyVatPayoutRow_(month, payoutByMonth, settlementFees, cogsByMonth, salesAgg, manualVat.byMonth, manualFees.byMonth, manualExpenseVat.byMonth);

  const headers = monthlyVatPayoutHeaders_();
  const sh = getMonthlyVatPayoutSheet_();
  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(2, 1, 1, headers.length).setValues([row]);
  applyMonthlyVatPayoutFormats_(sh, 1);

  return { month: month, paidOut: row[1], vatToPay: row[12], salesFileCount: row[15], settlementCount: row[14] };
}

function rebuildMonthlyVatPayoutSummary_() {

  const payoutByMonth = buildSettlementPayoutByMonth_();
  const settlementFees = buildSettlementFeesByMonth_();
  const cogsByMonth = buildSettlementCogsByMonth_();
  const salesAgg = buildSalesTaxMonthlyAgg_().byMonth || {};
  const manualVat = readManualVatByMonth_();
  const manualFees = readManualFeesByMonth_();
  const manualExpenseVat = collectManualExpenseVatByMonthReadOnly_();

  const months = mergeMonthKeys_(
    mergeMonthKeys_(Object.keys(payoutByMonth), Object.keys(salesAgg)),
    mergeMonthKeys_(Object.keys(cogsByMonth), mergeMonthKeys_(Object.keys(manualVat.byMonth), mergeMonthKeys_(Object.keys(manualFees.byMonth), Object.keys(manualExpenseVat.byMonth))))
  );

  const headers = monthlyVatPayoutHeaders_();
  const rows = months.map(function(m) {
    return buildMonthlyVatPayoutRow_(m, payoutByMonth, settlementFees, cogsByMonth, salesAgg, manualVat.byMonth, manualFees.byMonth, manualExpenseVat.byMonth);
  });

  const sh = getMonthlyVatPayoutSheet_();
  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    applyMonthlyVatPayoutFormats_(sh, rows.length);
  }

  let totalPayout = 0;
  for (let i = 0; i < rows.length; i++) totalPayout += parseNumberFlexible_(rows[i][1]);
  return { months: rows.length, totalPayout: totalPayout };
}

function monthlyVatPayoutHeaders_() {
  return [
    'Місяць',
    'Виплата Amazon',
    'Продажі без НДС',
    'НДС з продажів',
    'Продажі з НДС',
    'Комісії Amazon',
    'Собівартість',
    'Прибуток до НДС',
    'Вже сплачений НДС (введення)',
    'Вже сплачений НДС (з ручних витрат)',
    'Вже сплачений НДС (разом)',
    'Вже сплачений НДС',
    'НДС до оплати',
    'Залишок після НДС',
    'К-сть settlement файлів',
    'К-сть sales файлів',
    'Примітки'
  ];
}

function buildMonthlyVatPayoutRow_(month, payoutByMonth, settlementFeesByMonth, cogsByMonth, salesAgg, manualVatByMonth, manualFeesByMonth, manualExpenseVatByMonth) {
  const p = payoutByMonth[month] || { paidOut: 0, fileIds: {} };
  const s = salesAgg[month] || { salesAmount: 0, vatPayable: 0, grossSales: 0, fileIds: {}, rows: 0 };
  const cogs = (cogsByMonth[month] || { cogs: 0 }).cogs;
  const paidVatBreakdown = getPaidVatBreakdownForMonth_(month, manualVatByMonth, manualExpenseVatByMonth);
  const paidVat = paidVatBreakdown.totalPaidVat;
  const settlementFees = (settlementFeesByMonth[month] || { fees: 0 }).fees;
  const hasManualFees = !!manualFeesByMonth[month];
  const fees = hasManualFees ? manualFeesByMonth[month].fees : settlementFees;

  const profitBeforeVat = roundMoney_(p.paidOut - cogs);
  const settlementCount = Object.keys(p.fileIds || {}).length;
  const salesFileCount = Object.keys(s.fileIds || {}).length;
  const notes = (hasManualFees ? 'Комісії: ручне перевизначення. ' : 'Комісії: із settlement. ') +
    'Ручні витрати не впливають на COGS. ' +
    buildMonthlyVatPayoutNote_(p.paidOut, s.salesAmount, roundMoney_(s.vatPayable - paidVat), settlementCount, salesFileCount, s.rows);

  const row = [
    month,
    p.paidOut,
    s.salesAmount,
    s.vatPayable,
    s.grossSales,
    fees,
    cogs,
    profitBeforeVat,
    0,
    0,
    0,
    paidVat,
    roundMoney_(s.vatPayable - paidVat),
    roundMoney_(profitBeforeVat - roundMoney_(s.vatPayable - paidVat)),
    settlementCount,
    salesFileCount,
    notes
  ];

  rebuildPaidVatSection_(row, paidVatBreakdown);
  return applyManualExpensesToMonthlyReport_(row, month, {});
}

function applyMonthlyVatPayoutFormats_(sheet, rowCount) {
  if (!sheet || rowCount <= 0) return;
  safeSetNumberFormat_(sheet.getRange(2, 1, rowCount, 1), '@', [], 'monthly.month');
  safeSetNumberFormat_(sheet.getRange(2, 2, rowCount, 13), '#,##0.00', [], 'monthly.money');
  safeSetNumberFormat_(sheet.getRange(2, 15, rowCount, 2), '0', [], 'monthly.counts');
}

function buildSettlementFeesByMonth_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SUMMARY_SHEET);
  if (!sh || sh.getLastRow() < 2) return {};

  const all = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const hm = buildHeaderMapCaseInsensitive_(all[0]);
  const feeCol = hm[normalizeHeaderKey_(CONFIG.HEADERS.feesCost)];
  if (feeCol === undefined) return {};

  const out = {};
  for (let i = 1; i < all.length; i++) {
    const r = all[i];
    const fid = String(valueByHeader_(r, hm, CONFIG.HEADERS.fileId) || '').trim();
    if (!fid || fid === CONFIG.TOTAL_FILE_ID) continue;
    const month = monthFromSummaryRow_(r, (hm[normalizeHeaderKey_(CONFIG.HEADERS.month)] || -1) + 1, (hm[normalizeHeaderKey_(CONFIG.HEADERS.postedDate)] || -1) + 1, (hm[normalizeHeaderKey_(CONFIG.HEADERS.depositDate)] || -1) + 1);
    if (!month) continue;
    if (!out[month]) out[month] = { fees: 0 };
    out[month].fees += parseNumberFlexible_(r[feeCol]);
  }

  return out;
}

function buildSalesTaxMonthlyAgg_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SALES_TAX_RAW_SHEET);
  if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < 1) return { byMonth: {} };

  const all = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const headers = all[0].map(function(h) { return String(h || '').trim(); });
  const hm = buildHeaderMapCaseInsensitive_(headers);

  const required = ['Net Sales Total', 'VAT Total', 'Gross Sales Total', 'Import File ID'];
  const missing = findMissingHeaders_(hm, required);
  if (missing.length) throw new Error('Missing columns in ' + CONFIG.SALES_TAX_RAW_SHEET + ': ' + missing.join(', '));

  const out = {};
  for (let i = 1; i < all.length; i++) {
    const r = all[i];
    const month = resolveSalesTaxAggMonth_(r, hm);
    if (!month) continue;

    if (!out[month]) out[month] = { salesAmount: 0, vatPayable: 0, grossSales: 0, fileIds: {}, rows: 0 };
    out[month].salesAmount += parseNumberFlexible_(valueByHeader_(r, hm, 'Net Sales Total'));
    out[month].vatPayable += parseNumberFlexible_(valueByHeader_(r, hm, 'VAT Total'));
    out[month].grossSales += parseNumberFlexible_(valueByHeader_(r, hm, 'Gross Sales Total'));
    out[month].rows += 1;

    const fid = String(valueByHeader_(r, hm, 'Import File ID') || '').trim();
    if (fid) out[month].fileIds[fid] = true;
  }

  return { byMonth: out };
}

function buildSettlementCogsByMonth_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SUMMARY_SHEET);
  if (!sh || sh.getLastRow() < 2) return {};

  const all = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const hm = buildHeaderMapCaseInsensitive_(all[0]);

  const cogsHeader = CONFIG.HEADERS.cogs;
  const colCogs = hm[normalizeHeaderKey_(cogsHeader)];
  if (colCogs === undefined) return {};

  const out = {};
  for (let i = 1; i < all.length; i++) {
    const r = all[i];
    const fid = String(valueByHeader_(r, hm, CONFIG.HEADERS.fileId) || '').trim();
    if (!fid || fid === CONFIG.TOTAL_FILE_ID) continue;
    const month = monthFromSummaryRow_(r, (hm[normalizeHeaderKey_(CONFIG.HEADERS.month)] || -1) + 1, (hm[normalizeHeaderKey_(CONFIG.HEADERS.postedDate)] || -1) + 1, (hm[normalizeHeaderKey_(CONFIG.HEADERS.depositDate)] || -1) + 1);
    if (!month) continue;

    if (!out[month]) out[month] = { cogs: 0 };
    out[month].cogs += parseNumberFlexible_(r[colCogs]);
  }

  return out;
}

function runVatDiagnostics_() {
  const diagnosticsRows = writeDiagnostics_();
  return { rows: diagnosticsRows.length };
}

function writeDiagnostics_() {
  const rows = [];

  const salesAgg = buildSalesTaxMonthlyAgg_().byMonth || {};
  const payoutByMonth = buildSettlementPayoutByMonth_();
  const cogsByMonth = buildSettlementCogsByMonth_();
  const settlementFees = buildSettlementFeesByMonth_();
  const manualVat = readManualVatByMonth_();
  const manualFees = readManualFeesByMonth_();
  const manualExpenses = collectManualExpensesByMonth_();
  const manualByFundCategory = collectManualExpensesByFundCategoryByMonth_();
  const manualExpenseVat = collectManualExpenseVatByMonthReadOnly_();
  const manualValidation = validateManualExpenseRows_();
  const manualVatValidation = validateManualExpenseVatRows_();
  const purchasePreview = collectPurchaseRowsForSync_();
  const purchaseSync = syncManualPurchasesToZakupky_();
  const workingCapitalAgg = collectWorkingCapitalByMonth_();
  const workingCapitalTotals = getWorkingCapitalTotals_(workingCapitalAgg);
  const legacyStatus = inspectLegacyMonthlySheetUsage_();

  rows.push(['INFO', 'A. Імпортовані settlement файли', 'File ID', 'Назва файлу', 'Дата виплати', 'Ефективна posted date', 'Призначений місяць', 'Виплата', 'COGS | Комісії']);
  const settlementFiles = buildSettlementFilesRegistry_();
  const settlementIds = Object.keys(settlementFiles).sort();
  for (let i = 0; i < settlementIds.length; i++) {
    const fid = settlementIds[i];
    const it = settlementFiles[fid];
    rows.push(['DATA', 'A. Імпортовані settlement файли', fid, it.fileName, it.depositDate, it.postedDate, it.month, it.payout, it.cogs + ' | ' + it.fees]);
  }

  rows.push(['INFO', 'B. Імпортовані звіти продажів', 'File ID', 'Назва файлу', 'Імпортовано рядків', 'Перший місяць', 'Останній місяць', 'Імпортовано', '']);
  const salesFiles = buildSalesFilesRegistry_();
  const salesFileIds = Object.keys(salesFiles).sort();
  for (let j = 0; j < salesFileIds.length; j++) {
    const sfid = salesFileIds[j];
    const sit = salesFiles[sfid];
    rows.push(['DATA', 'B. Імпортовані звіти продажів', sfid, sit.fileName, sit.rows, sit.firstMonth, sit.lastMonth, Object.keys(sit.importedAtSet).sort().join(' | '), '']);
  }

  rows.push(['INFO', 'C. Підсумок по місяцях', 'Місяць', 'Виплата', 'COGS', 'Комісії settlement', 'НДС з продажів', 'Сплачений НДС', 'Ручні витрати']);
  const months = mergeMonthKeys_(Object.keys(payoutByMonth), mergeMonthKeys_(Object.keys(salesAgg), mergeMonthKeys_(Object.keys(cogsByMonth), mergeMonthKeys_(Object.keys(manualVat.byMonth), mergeMonthKeys_(Object.keys(manualFees.byMonth), mergeMonthKeys_(Object.keys(manualExpenses.byMonth), Object.keys(manualExpenseVat.byMonth)))))));
  for (let k = 0; k < months.length; k++) {
    const m = months[k];
    const paidVatBreakdown = getPaidVatBreakdownForMonth_(m, manualVat.byMonth, manualExpenseVat.byMonth);
    rows.push(['DATA', 'C. Підсумок по місяцях', m,
      (payoutByMonth[m] || { paidOut: 0 }).paidOut,
      (cogsByMonth[m] || { cogs: 0 }).cogs,
      (settlementFees[m] || { fees: 0 }).fees,
      (salesAgg[m] || { vatPayable: 0 }).vatPayable,
      paidVatBreakdown.totalPaidVat,
      'ручні витрати: ' + (manualExpenses.byMonth[m] || { amount: 0 }).amount
    ]);
  }

  rows.push(['INFO', 'C2. Cashflow totals by month', 'Місяць', 'Виплата Amazon', 'Повернення собівартості', 'Прибуток до НДС', 'НДС до оплати', 'Чистий кеш після НДС', 'Реально вільний кеш']);
  for (let k2 = 0; k2 < months.length; k2++) {
    const month = months[k2];
    const payout = (payoutByMonth[month] || { paidOut: 0 }).paidOut;
    const cogs = (cogsByMonth[month] || { cogs: 0 }).cogs;
    const paidVatBreakdown = getPaidVatBreakdownForMonth_(month, manualVat.byMonth, manualExpenseVat.byMonth);
    const vatToPay = roundMoney_(((salesAgg[month] || { vatPayable: 0 }).vatPayable || 0) - paidVatBreakdown.totalPaidVat);
    const profitBeforeVat = roundMoney_(payout - cogs);
    const cashAfterVat = roundMoney_(profitBeforeVat - vatToPay);
    rows.push(['DATA', 'C2. Cashflow totals by month', month, payout, cogs, profitBeforeVat, vatToPay, cashAfterVat, cashAfterVat]);
  }

  rows.push(['INFO', 'C3. VAT reserve by month', 'Місяць', 'НДС нараховано', 'Вже сплачений НДС', 'НДС до оплати', 'Тимчасово в обороті', 'Ризик', '']);
  for (let k3 = 0; k3 < months.length; k3++) {
    const monthKey = months[k3];
    const paidVatBreakdown = getPaidVatBreakdownForMonth_(monthKey, manualVat.byMonth, manualExpenseVat.byMonth);
    const vatAccrued = roundMoney_((salesAgg[monthKey] || { vatPayable: 0 }).vatPayable || 0);
    const vatToPay = roundMoney_(vatAccrued - paidVatBreakdown.totalPaidVat);
    rows.push(['DATA', 'C3. VAT reserve by month', monthKey, vatAccrued, paidVatBreakdown.totalPaidVat, vatToPay, vatToPay, vatToPay > 0 ? 'ризик: є резерв' : 'ok', '']);
  }

  rows.push(['INFO', 'D. Ручні витрати', 'Усього рядків', manualValidation.totalRows, 'Активні рядки', manualValidation.activeRows, 'У прибутку', manualValidation.includedInProfitRows, 'Неактивні: ' + manualValidation.inactiveRows]);
  rows.push(['INFO', 'D. Ручні витрати', 'Невалідні дати', manualValidation.invalidDateRows.join(', '), 'Рядки без суми', manualValidation.missingAmountRows.join(', '), 'Рядки без типу', manualValidation.missingTypeRows.join(', '), 'Закуп у постачальника: ' + manualValidation.supplierPurchaseRows]);
  rows.push(['INFO', 'D. Ручні витрати', 'Бізнес-витрата', manualValidation.businessExpenseRows, 'Попередній перегляд рядків типу закуп у постачальника', purchasePreview.rows.length, 'Синхронізація в Закупки', 'вимкнено: ' + purchaseSync.skipped]);
  rows.push(['INFO', 'D3. Категорія коштів', 'Порожні категорії', manualValidation.missingFundCategoryRows.join(', '), 'Невідомі категорії', manualValidation.unknownFundCategoryRows.join(', '), 'Категорій у довіднику', getManualExpenseFundCategories_().length, '']);
  const fundCategories = getManualExpenseFundCategories_();
  for (let fc = 0; fc < fundCategories.length; fc++) {
    const fundCategory = fundCategories[fc];
    rows.push(['DATA', 'D3. Категорія коштів', fundCategory, manualValidation.totalsByFundCategory[fundCategory] || 0, 'Сума активних витрат', '', '', '', '']);
  }
  rows.push(['INFO', 'D2. НДС з ручних витрат', 'Усього рядків', manualVatValidation.totalRows, 'Активні рядки', manualVatValidation.activeRows, 'Враховано НДС рядків', manualVatValidation.includedRows, 'Пропущено: неактивні=' + manualVatValidation.skippedInactive + ', невалідна дата=' + manualVatValidation.skippedInvalidDate + ', без НДС=' + manualVatValidation.skippedMissingVat]);

  const manualExpenseMonths = Object.keys(manualValidation.totalsByMonth).sort();
  rows.push(['INFO', 'E. Ручні витрати по місяцях', 'Місяць', 'Сума для прибутку', 'НДС', 'Коментар', '', '', '']);
  for (let x = 0; x < manualExpenseMonths.length; x++) {
    const monthKey = manualExpenseMonths[x];
    rows.push(['DATA', 'E. Ручні витрати по місяцях', monthKey, manualValidation.totalsByMonth[monthKey], manualValidation.vatTotalsByMonth[monthKey] || 0, 'Активні та з прапорцем "Враховувати у прибутку"', '', '', '']);
  }

  rows.push(['INFO', 'F. Ручний НДС по місяцях', 'Місяць', 'Сплачений НДС', 'Рядків', 'Попередження', '', '', '']);
  const vatMonths = Object.keys(manualVat.byMonth).sort();
  for (let y = 0; y < vatMonths.length; y++) {
    const vm = vatMonths[y];
    const v = manualVat.byMonth[vm];
    rows.push(['DATA', 'F. Ручний НДС по місяцях', vm, v.paidVat, v.rows, '', '', '', '']);
  }

  rows.push(['INFO', 'F2. НДС з ручних витрат по місяцях', 'Місяць', 'НДС з ручних витрат', 'Рядків', 'Коментар', '', '', '']);
  const vatExpenseMonths = Object.keys(manualExpenseVat.byMonth).sort();
  for (let y2 = 0; y2 < vatExpenseMonths.length; y2++) {
    const vem = vatExpenseMonths[y2];
    const ve = manualExpenseVat.byMonth[vem];
    rows.push(['DATA', 'F2. НДС з ручних витрат по місяцях', vem, ve.vat, ve.rows, 'Активні рядки з валідною датою і НДС > 0', '', '', '']);
  }

  rows.push(['INFO', 'F3. Комбінований сплачений НДС', 'Місяць', 'НДС введення', 'НДС ручні витрати', 'Разом сплачений НДС', '', '', '']);
  const combinedVatMonths = mergeMonthKeys_(Object.keys(manualVat.byMonth), Object.keys(manualExpenseVat.byMonth));
  for (let y3 = 0; y3 < combinedVatMonths.length; y3++) {
    const cvm = combinedVatMonths[y3];
    const breakdown = getPaidVatBreakdownForMonth_(cvm, manualVat.byMonth, manualExpenseVat.byMonth);
    rows.push(['DATA', 'F3. Комбінований сплачений НДС', cvm, breakdown.manualInputVat, breakdown.manualExpenseVat, breakdown.totalPaidVat, '', '', '']);
  }

  rows.push(['INFO', 'G. Ручні комісії по місяцях', 'Місяць', 'Комісії', 'Рядків', 'Попередження', '', '', '']);
  const feeMonths = Object.keys(manualFees.byMonth).sort();
  for (let z = 0; z < feeMonths.length; z++) {
    const fm = feeMonths[z];
    const f = manualFees.byMonth[fm];
    rows.push(['DATA', 'G. Ручні комісії по місяцях', fm, f.fees, f.rows, '', '', '', '']);
  }

  rows.push(['INFO', 'G2. Ручні витрати за категоріями (місяць)', 'Місяць', 'Категорія коштів', 'Сума', 'Рядків', 'Коментар', '', '']);
  const fundCategoryMonths = Object.keys(manualByFundCategory.byMonth || {}).sort();
  for (let fcm = 0; fcm < fundCategoryMonths.length; fcm++) {
    const fmKey = fundCategoryMonths[fcm];
    const fmBucket = manualByFundCategory.byMonth[fmKey] || { categories: {}, rows: 0 };
    for (let fcc = 0; fcc < fundCategories.length; fcc++) {
      const fcName = fundCategories[fcc];
      rows.push(['DATA', 'G2. Ручні витрати за категоріями (місяць)', fmKey, fcName, roundMoney_((fmBucket.categories || {})[fcName] || 0), fmBucket.rows || 0, '', '', '']);
    }
  }


  rows.push(['INFO', 'G3. Оборотний капітал', 'Активних рядків', workingCapitalAgg.activeRows || 0, 'Активних місяців', (workingCapitalAgg.months || []).length, 'Поточний місяць', workingCapitalTotals.latestMonth || '', '']);
  rows.push(['INFO', 'G3. Оборотний капітал', 'Поточний оборотний капітал', roundMoney_(workingCapitalTotals.currentWorkingCapital || 0), 'Середній оборотний капітал', roundMoney_(workingCapitalTotals.averageWorkingCapital || 0), 'Зміна до попереднього місяця', roundMoney_(workingCapitalTotals.changeVsPreviousMonth || 0), '']);
  const wcDiag = (workingCapitalAgg && workingCapitalAgg.diagnostics) ? workingCapitalAgg.diagnostics : {};
  rows.push(['INFO', 'G3. Оборотний капітал', 'Невалідні місяці', wcDiag.invalidMonths || 0, 'Пропущені місяці', wcDiag.missingMonths || 0, 'Нечислові значення', wcDiag.nonNumericValues || 0, '']);
  rows.push(['INFO', 'G3. Оборотний капітал по місяцях', 'Місяць', 'Собівартість товарів на Amazon', 'Кошти доступні на вивід Amazon', 'Товари готові до відправки', 'Кошти на руках', 'Можливість кредитування', 'Оборотний капітал']);
  for (let wc = 0; wc < (workingCapitalAgg.months || []).length; wc++) {
    const wcMonth = workingCapitalAgg.months[wc];
    const wcData = (workingCapitalAgg.byMonth || {})[wcMonth] || {};
    rows.push(['DATA', 'G3. Оборотний капітал по місяцях', wcMonth, wcData.cogsAmazon || 0, wcData.amazonPayoutAvailable || 0, wcData.goodsReadyToShip || 0, wcData.cashOnHand || 0, wcData.creditCapacity || 0, wcData.workingCapital || 0]);
  }
  rows.push(['INFO', 'G3. Оборотний капітал totals by month', 'Місяць', 'Активних рядків', 'Разом оборотний капітал', '', '', '', '']);
  const wcTotalsByMonth = (wcDiag && wcDiag.totalsByMonth) ? wcDiag.totalsByMonth : {};
  const wcDiagMonths = Object.keys(wcTotalsByMonth).sort();
  for (let wdm = 0; wdm < wcDiagMonths.length; wdm++) {
    const m = wcDiagMonths[wdm];
    const md = wcTotalsByMonth[m] || { rows: 0, total: 0 };
    rows.push(['DATA', 'G3. Оборотний капітал totals by month', m, md.rows || 0, roundMoney_(md.total || 0), '', '', '', '']);
  }

  const unique = buildSalesTaxUniques_();
  rows.push(['INFO', 'H. Унікальні значення', 'Ставки податку', unique.taxRates.join(', '), '', '', '', '', '']);
  rows.push(['INFO', 'H. Унікальні значення', 'Країни доставки', unique.shipToCountries.join(', '), '', '', '', '', '']);
  rows.push(['INFO', 'H. Унікальні значення', 'Відповідальний за збір податку', unique.taxCollectionResponsibility.join(', '), '', '', '', '', '']);

  rows.push(['INFO', 'I. Legacy monthly sheet', 'Legacy sheet', legacyStatus.legacyExists ? CONFIG.LEGACY_MONTHLY_SHEET : 'відсутній', legacyStatus.hidden ? 'прихований' : 'видимий', 'rows=' + legacyStatus.legacyRows, 'migrated unique=' + legacyStatus.uniqueLegacyRows, 'sourceRefs=' + legacyStatus.sourceReferences, legacyStatus.redirectStatus]);
  if (legacyStatus.sourceReferences > 0) rows.push(['WARN', 'I. Legacy monthly sheet', 'У коді знайдено legacy references', String(legacyStatus.sourceReferences), '', '', '', '', '']);
  if (legacyStatus.legacyExists && legacyStatus.uniqueLegacyRows > 0) rows.push(['WARN', 'I. Legacy monthly sheet', 'Legacy sheet містить рядки поза main', String(legacyStatus.uniqueLegacyRows), '', '', '', '', '']);

  const allWarnings = []
    .concat(manualVat.warnings || [], manualFees.warnings || [], manualExpenses.warnings || [], manualByFundCategory.warnings || [], manualExpenseVat.warnings || [], manualValidation.warnings || [], purchasePreview.warnings || [], purchaseSync.errors || [], workingCapitalAgg.warnings || []);
  for (let q = 0; q < allWarnings.length; q++) {
    rows.push(['WARN', 'J. Попередження', allWarnings[q], '', '', '', '', '', '']);
  }

  const sh = getOrCreateSheet_(CONFIG.DIAGNOSTICS_SHEET);
  sh.clearContents();
  sh.getRange(1, 1, 1, 9).setValues([['Рівень', 'Розділ', 'Колонка 1', 'Колонка 2', 'Колонка 3', 'Колонка 4', 'Колонка 5', 'Колонка 6', 'Колонка 7']]);
  if (rows.length) sh.getRange(2, 1, rows.length, 9).setValues(rows);

  return rows;
}

function buildSettlementFilesRegistry_() {
  const out = {};
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SUMMARY_SHEET);
  if (!sh || sh.getLastRow() < 2) return out;

  const all = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const hm = buildHeaderMapCaseInsensitive_(all[0]);

  for (let i = 1; i < all.length; i++) {
    const r = all[i];
    const fid = String(valueByHeader_(r, hm, CONFIG.HEADERS.fileId) || '').trim();
    if (!fid || fid === CONFIG.TOTAL_FILE_ID) continue;
    if (!out[fid]) {
      out[fid] = {
        fileName: String(valueByHeader_(r, hm, CONFIG.HEADERS.fileName) || ''),
        depositDate: formatDateForDiagnostics_(valueByHeader_(r, hm, CONFIG.HEADERS.depositDate)),
        postedDate: formatDateForDiagnostics_(valueByHeader_(r, hm, CONFIG.HEADERS.postedDate)),
        month: '',
        payout: 0,
        cogs: 0,
        fees: 0,
        months: {}
      };
    }

    const month = monthFromSummaryRow_(r, (hm[normalizeHeaderKey_(CONFIG.HEADERS.month)] || -1) + 1, (hm[normalizeHeaderKey_(CONFIG.HEADERS.postedDate)] || -1) + 1, (hm[normalizeHeaderKey_(CONFIG.HEADERS.depositDate)] || -1) + 1);
    out[fid].payout += parseNumberFlexible_(valueByHeader_(r, hm, CONFIG.HEADERS.transfer));
    out[fid].cogs += parseNumberFlexible_(valueByHeader_(r, hm, CONFIG.HEADERS.cogs));
    out[fid].fees += parseNumberFlexible_(valueByHeader_(r, hm, CONFIG.HEADERS.feesCost));
    if (month) out[fid].months[month] = true;

    const monthList = Object.keys(out[fid].months).sort();
    out[fid].month = monthList.join(', ');
  }

  return out;
}

function buildSalesFilesRegistry_() {
  const out = {};
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SALES_TAX_RAW_SHEET);
  if (!sh || sh.getLastRow() < 2) return out;

  const all = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const hm = buildHeaderMapCaseInsensitive_(all[0]);

  for (let i = 1; i < all.length; i++) {
    const r = all[i];
    const fid = String(valueByHeader_(r, hm, 'Import File ID') || '').trim();
    if (!fid) continue;

    const fileName = String(valueByHeader_(r, hm, 'Import File Name') || '').trim();
    const importedAt = String(valueByHeader_(r, hm, 'Imported At') || '').trim();
    const month = resolveSalesTaxAggMonth_(r, hm) || '';

    if (!out[fid]) out[fid] = { fileName: fileName, rows: 0, firstMonth: month, lastMonth: month, importedAtSet: {} };
    out[fid].rows += 1;
    if (importedAt) out[fid].importedAtSet[importedAt] = true;
    if (month && (!out[fid].firstMonth || month < out[fid].firstMonth)) out[fid].firstMonth = month;
    if (month && (!out[fid].lastMonth || month > out[fid].lastMonth)) out[fid].lastMonth = month;
  }

  return out;
}

function inspectLegacyMonthlySheetUsage_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const legacy = ss.getSheetByName(CONFIG.LEGACY_MONTHLY_SHEET);
  const main = ss.getSheetByName(CONFIG.MONTHLY_VAT_PAYOUT_SUMMARY_SHEET);
  const uniqueLegacyRows = countUniqueLegacyMonthlyRows_(legacy, main);

  return {
    legacyExists: !!legacy,
    hidden: !!(legacy && legacy.isSheetHidden && legacy.isSheetHidden()),
    legacyRows: legacy ? Math.max(0, legacy.getLastRow() - 1) : 0,
    uniqueLegacyRows: uniqueLegacyRows,
    sourceReferences: 0,
    redirectStatus: legacy ? 'redirect ready' : 'clean'
  };
}

function countUniqueLegacyMonthlyRows_(legacySheet, mainSheet) {
  if (!legacySheet || !mainSheet || legacySheet.getSheetId() === mainSheet.getSheetId()) return 0;
  if (legacySheet.getLastRow() < 2) return 0;

  const mainLastColumn = Math.max(mainSheet.getLastColumn(), 1);
  const mainHeaders = mainSheet.getRange(1, 1, 1, mainLastColumn).getValues()[0].map(function(v) { return String(v || '').trim(); });
  const mainHeaderMap = buildHeaderMapCaseInsensitive_(mainHeaders);
  const existingKeys = {};
  if (mainSheet.getLastRow() >= 2) {
    const mainRows = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainLastColumn).getValues();
    for (let i = 0; i < mainRows.length; i++) {
      const key = buildMonthlySheetRowKey_(mainRows[i], mainHeaders, mainHeaderMap);
      if (key) existingKeys[key] = true;
    }
  }

  const legacyLastColumn = Math.max(legacySheet.getLastColumn(), 1);
  const legacyRows = legacySheet.getRange(1, 1, legacySheet.getLastRow(), legacyLastColumn).getValues();
  const legacyHeaders = legacyRows[0].map(function(v) { return String(v || '').trim(); });
  const legacyHeaderMap = buildHeaderMapCaseInsensitive_(legacyHeaders);

  let unique = 0;
  for (let r = 1; r < legacyRows.length; r++) {
    const legacyRow = legacyRows[r];
    if (isEmptyRow_(legacyRow)) continue;
    const mapped = mainHeaders.map(function(header) {
      const idx = legacyHeaderMap[normalizeHeaderKey_(header)];
      return idx === undefined ? '' : legacyRow[idx];
    });
    const key = buildMonthlySheetRowKey_(mapped, mainHeaders, mainHeaderMap);
    if (key && !existingKeys[key]) unique += 1;
  }
  return unique;
}

function buildSalesTaxUniques_() {
  const out = { taxRates: [], shipToCountries: [], taxCollectionResponsibility: [] };
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SALES_TAX_RAW_SHEET);
  if (!sh || sh.getLastRow() < 2) return out;

  const all = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const hm = buildHeaderMapCaseInsensitive_(all[0]);
  const taxRates = {};
  const countries = {};
  const responsibilities = {};

  for (let i = 1; i < all.length; i++) {
    const r = all[i];
    const taxRate = String(valueByHeader_(r, hm, 'Tax Rate') || '').trim();
    const country = String(valueByHeader_(r, hm, 'Ship To Country') || '').trim();
    const resp = String(valueByHeader_(r, hm, 'Tax Collection Responsibility') || '').trim();
    if (taxRate) taxRates[taxRate] = true;
    if (country) countries[country] = true;
    if (resp) responsibilities[resp] = true;
  }

  out.taxRates = Object.keys(taxRates).sort();
  out.shipToCountries = Object.keys(countries).sort();
  out.taxCollectionResponsibility = Object.keys(responsibilities).sort();
  return out;
}

function formatDateForDiagnostics_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, CONFIG.TZ, 'yyyy-MM-dd');
  const parsed = parseDateFlexible_(value, CONFIG.TZ);
  if (parsed instanceof Date && !isNaN(parsed.getTime())) return Utilities.formatDate(parsed, CONFIG.TZ, 'yyyy-MM-dd');
  return String(value || '').trim();
}

function arraysEqualByTrim_(a, b) {
  if (!a || !b || a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (String(a[i] || '').trim() !== String(b[i] || '').trim()) return false;
  }
  return true;
}

function rebuildDashboard_() {
  const ss = SpreadsheetApp.getActive();
  const source = ss.getSheetByName(CONFIG.MONTHLY_SHEET);
  if (!source) throw new Error('Не знайдено лист: ' + CONFIG.MONTHLY_SHEET);

  const dashboard = ensureDashboardSheet_();
  clearDashboardSheet_(dashboard);

  const monthlyData = readMonthlyDashboardRows_(source);
  const monthlyRows = monthlyData.rows || [];
  const monthlyWarnings = monthlyData.warnings || [];
  if (!monthlyRows.length) {
    renderEmptyDashboard_(dashboard);
    return { empty: true, months: 0, latestMonth: '' };
  }

  monthlyRows.sort(function(a, b) {
    if (a.month < b.month) return -1;
    if (a.month > b.month) return 1;
    return 0;
  });

  const manualByCategory = collectManualExpenseCategoryTotalsReadOnly_();
  const workingCapitalAgg = collectWorkingCapitalByMonth_();
  const workingCapitalTotals = getWorkingCapitalTotals_(workingCapitalAgg);
  const dashboardWarnings = []
    .concat(monthlyWarnings)
    .concat((manualByCategory && manualByCategory.warnings) || [])
    .concat((workingCapitalAgg && workingCapitalAgg.warnings) || []);
  renderDashboardBlocks_(dashboard, monthlyRows, manualByCategory, workingCapitalAgg, workingCapitalTotals, dashboardWarnings);
  renderDashboardCharts_(dashboard, monthlyRows, workingCapitalAgg);
  dashboard.autoResizeColumns(1, 8);
  dashboard.setFrozenRows(2);

  return { empty: false, months: monthlyRows.length, latestMonth: monthlyRows[monthlyRows.length - 1].month };
}

function ensureDashboardSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG.DASHBOARD_SHEET);
  if (!sh) sh = ss.insertSheet(CONFIG.DASHBOARD_SHEET);
  return sh;
}

function clearDashboardSheet_(sheet) {
  if (!sheet) return;
  sheet.clear();
  const charts = sheet.getCharts();
  for (let i = 0; i < charts.length; i++) sheet.removeChart(charts[i]);
}

function readMonthlyDashboardRows_(sheet) {
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return { rows: [], warnings: [] };
  const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const hm = buildHeaderMapCaseInsensitive_(values[0]);

  function pickHeaderIndex_(candidates) {
    for (let i = 0; i < candidates.length; i++) {
      const idx = hm[normalizeHeaderKey_(candidates[i])];
      if (idx !== undefined) return idx;
    }
    return undefined;
  }

  const idxMonth = pickHeaderIndex_(['Місяць']);
  const idxPayout = pickHeaderIndex_(['Виплата Amazon']);
  const idxCogs = pickHeaderIndex_(['Собівартість']);
  const idxProfitBeforeVat = pickHeaderIndex_(['Прибуток до НДС']);
  const idxVatToPay = pickHeaderIndex_(['НДС до оплати']);
  const idxVatPaidTotal = pickHeaderIndex_(['Вже сплачений НДС (разом)', 'Вже сплачений НДС']);
  const idxCashAfterVat = pickHeaderIndex_(['Залишок після НДС']);

  const missing = [];
  if (idxMonth === undefined) missing.push('Місяць');
  if (idxPayout === undefined) missing.push('Виплата Amazon');
  if (idxCogs === undefined) missing.push('Собівартість');
  if (idxProfitBeforeVat === undefined) missing.push('Прибуток до НДС');
  if (idxVatToPay === undefined) missing.push('НДС до оплати');
  if (idxCashAfterVat === undefined) missing.push('Залишок після НДС');
  if (missing.length) throw new Error('У ' + CONFIG.MONTHLY_SHEET + ' відсутні колонки: ' + missing.join(', '));

  const out = [];
  const warnings = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const month = toMonthText_(row[idxMonth]);
    if (!month) {
      if (!isEmptyRow_(row)) warnings.push('МІСЯЧНИЙ_ЗВІТ рядок ' + (i + 1) + ': невалідний місяць.');
      continue;
    }

    const rawPayout = row[idxPayout];
    const rawCogs = row[idxCogs];
    const rawProfitBeforeVat = row[idxProfitBeforeVat];
    const rawVatToPay = row[idxVatToPay];
    const rawVatPaid = idxVatPaidTotal === undefined ? 0 : row[idxVatPaidTotal];
    const rawCashAfterVat = row[idxCashAfterVat];
    const numericChecks = [
      { header: 'Виплата Amazon', value: rawPayout },
      { header: 'Собівартість', value: rawCogs },
      { header: 'Прибуток до НДС', value: rawProfitBeforeVat },
      { header: 'НДС до оплати', value: rawVatToPay },
      { header: 'Залишок після НДС', value: rawCashAfterVat }
    ];
    if (idxVatPaidTotal !== undefined) numericChecks.push({ header: 'Вже сплачений НДС', value: rawVatPaid });
    for (let q = 0; q < numericChecks.length; q++) {
      if (!isDashboardNumberLike_(numericChecks[q].value)) {
        warnings.push('МІСЯЧНИЙ_ЗВІТ рядок ' + (i + 1) + ': нечислове значення у колонці "' + numericChecks[q].header + '".');
      }
    }

    out.push({
      month: month,
      payout: roundMoney_(parseNumberFlexible_(rawPayout)),
      cogs: roundMoney_(parseNumberFlexible_(rawCogs)),
      profitBeforeVat: roundMoney_(parseNumberFlexible_(rawProfitBeforeVat)),
      vatToPay: roundMoney_(parseNumberFlexible_(rawVatToPay)),
      vatPaid: roundMoney_(parseNumberFlexible_(rawVatPaid)),
      cashAfterVat: roundMoney_(parseNumberFlexible_(rawCashAfterVat)),
      freeCash: roundMoney_(parseNumberFlexible_(rawCashAfterVat))
    });
  }

  const byMonth = {};
  for (let j = 0; j < out.length; j++) byMonth[out[j].month] = out[j];
  return {
    rows: Object.keys(byMonth).map(function(key) { return byMonth[key]; }),
    warnings: warnings
  };
}

function isDashboardNumberLike_(value) {
  if (value === null || value === undefined || value === '') return true;
  if (typeof value === 'number') return isFinite(value);
  const parsed = parseNumberFlexible_(value);
  const text = String(value).replace(/\s/g, '').replace(',', '.').trim();
  if (!text) return true;
  const simple = text.replace(/^[+-]/, '');
  if (!/^\d*(\.\d+)?$/.test(simple)) return false;
  return isFinite(parsed);
}

function renderEmptyDashboard_(sheet) {
  sheet.getRange('A1').setValue('ДЕШБОРД AMAZON FBA').setFontWeight('bold').setFontSize(18);
  sheet.getRange('A3').setValue('Немає даних у листі ' + CONFIG.MONTHLY_SHEET + '. Спочатку сформуйте місячний звіт.');
  sheet.getRange('A3').setFontColor('#b71c1c').setFontWeight('bold');
  sheet.autoResizeColumns(1, 4);
}

function getDashboardFundAllocationRules_() {
  return [
    { category: 'Реінвест (75%)', share: 0.75 },
    { category: 'Бізнес витрати (12%)', share: 0.12 },
    { category: 'Зарплата (7%)', share: 0.07 },
    { category: 'Інше (6%)', share: 0.06 }
  ];
}

function calculateDashboardFundAllocatedByCategory_(totalProfitAfterVat) {
  const rules = getDashboardFundAllocationRules_();
  const out = {};
  for (let i = 0; i < rules.length; i++) {
    out[rules[i].category] = roundMoney_(totalProfitAfterVat * rules[i].share);
  }
  return out;
}

function aggregateManualExpensesForDashboardByCategory_(manualByCategoryAgg) {
  const categories = getManualExpenseFundCategories_();
  const out = buildManualExpenseFundCategoryTotalsSkeleton_();
  const totals = (manualByCategoryAgg && manualByCategoryAgg.totalsByCategory) ? manualByCategoryAgg.totalsByCategory : {};
  for (let i = 0; i < categories.length; i++) {
    const category = categories[i];
    out[category] = roundMoney_(totals[category] || 0);
  }
  return out;
}

function getDashboardProfitOnlyCategories_() {
  return [
    'Бізнес витрати (12%)',
    'Зарплата (7%)',
    'Інше (6%)'
  ];
}

function calculateDashboardProfitDistributionRows_(allocatedByCategory) {
  const rules = getDashboardFundAllocationRules_();
  const out = [];
  for (let i = 0; i < rules.length; i++) {
    const category = rules[i].category;
    const label = category === 'Реінвест (75%)' ? 'Реінвест із прибутку (75%)' : category;
    out.push([label, roundMoney_((allocatedByCategory || {})[category] || 0)]);
  }
  return out;
}

function calculateDashboardReinvestFundRows_(totals, allocatedByCategory, spentByCategory, latest) {
  const cogsReturnTotal = roundMoney_((totals && totals.cogs) || 0);
  const reinvestFromProfitTotal = roundMoney_((allocatedByCategory || {})['Реінвест (75%)'] || 0);
  const fullFundTotal = roundMoney_(cogsReturnTotal + reinvestFromProfitTotal);
  const spentReinvestTotal = roundMoney_((spentByCategory || {})['Реінвест (75%)'] || 0);
  const remainingTotal = roundMoney_(fullFundTotal - spentReinvestTotal);

  const latestRow = latest || {};
  const latestCashAfterVat = roundMoney_(latestRow.cashAfterVat || 0);
  const latestCogs = roundMoney_(latestRow.cogs || 0);
  const latestReinvestFromProfit = roundMoney_(latestCashAfterVat * 0.75);
  const latestFullFund = roundMoney_(latestCogs + latestReinvestFromProfit);
  const latestSpentReinvest = roundMoney_((latestRow.manualSpentReinvest || 0));
  const latestRemaining = roundMoney_(latestFullFund - latestSpentReinvest);

  return [
    ['Собівартість до повернення (весь період)', cogsReturnTotal],
    ['Реінвест із прибутку (весь період)', reinvestFromProfitTotal],
    ['Повний фонд реінвесту (весь період)', fullFundTotal],
    ['Фактично витрачено з фонду реінвесту (весь період)', spentReinvestTotal],
    ['Залишок фонду реінвесту (весь період)', remainingTotal],
    ['Собівартість до повернення (останній місяць)', latestCogs],
    ['Реінвест із прибутку (останній місяць)', latestReinvestFromProfit],
    ['Повний фонд реінвесту (останній місяць)', latestFullFund],
    ['Фактично витрачено з фонду реінвесту (останній місяць)', latestSpentReinvest],
    ['Залишок фонду реінвесту (останній місяць)', latestRemaining]
  ];
}

function calculateDashboardProfitBucketExpenseRows_(allocatedByCategory, spentByCategory) {
  const categories = getDashboardProfitOnlyCategories_();
  const out = [];
  for (let i = 0; i < categories.length; i++) {
    const category = categories[i];
    const allocated = roundMoney_((allocatedByCategory || {})[category] || 0);
    const spent = roundMoney_((spentByCategory || {})[category] || 0);
    const remaining = roundMoney_(allocated - spent);
    out.push([category, allocated, spent, remaining]);
  }
  return out;
}

function getDashboardReinvestNegativeRows_(reinvestFundRows) {
  const negativeRows = [];
  for (let i = 0; i < reinvestFundRows.length; i++) {
    const label = String(reinvestFundRows[i][0] || '');
    const value = roundMoney_(reinvestFundRows[i][1] || 0);
    if (label.indexOf('Залишок фонду реінвесту') !== -1 && value < 0) negativeRows.push(i);
  }
  return negativeRows;
}

function getDashboardProfitBucketNegativeRows_(profitBucketRows) {
  const out = [];
  for (let i = 0; i < profitBucketRows.length; i++) {
    if (roundMoney_(profitBucketRows[i][3] || 0) < 0) out.push(i);
  }
  return out;
}

function percentOf_(value, total) {
  return total ? roundMoney_((value / total) * 100) : 0;
}

function renderDashboardBlocks_(sheet, rows, manualByCategoryAgg, workingCapitalAgg, workingCapitalTotals, dashboardWarnings) {
  const totals = rows.reduce(function(acc, row) {
    acc.payout += row.payout;
    acc.cogs += row.cogs;
    acc.profitBeforeVat += row.profitBeforeVat;
    acc.vatToPay += row.vatToPay;
    acc.vatPaid += row.vatPaid;
    acc.cashAfterVat += row.cashAfterVat;
    acc.freeCash += row.freeCash;
    return acc;
  }, { payout: 0, cogs: 0, profitBeforeVat: 0, vatToPay: 0, vatPaid: 0, cashAfterVat: 0, freeCash: 0 });

  const latest = rows[rows.length - 1];
  const wcTotals = workingCapitalTotals || { currentWorkingCapital: 0, averageWorkingCapital: 0, changeVsPreviousMonth: 0 };
  sheet.getRange('A1').setValue('ДЕШБОРД AMAZON FBA').setFontWeight('bold').setFontSize(18);
  sheet.getRange('A2').setValue('Оновлено: ' + Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss'));

  const kpiHeaders = [
    'Загальний чистий прибуток після НДС',
    'Загальний прибуток до НДС',
    'Загальна виплата Amazon',
    'Загальна собівартість',
    'Загальний НДС до оплати',
    'Реально вільний кеш',
    'Поточний оборотний капітал',
    'Середній оборотний капітал'
  ];
  const kpiValues = [
    roundMoney_(totals.cashAfterVat),
    roundMoney_(totals.profitBeforeVat),
    roundMoney_(totals.payout),
    roundMoney_(totals.cogs),
    roundMoney_(totals.vatToPay),
    roundMoney_(totals.freeCash),
    roundMoney_(wcTotals.currentWorkingCapital || 0),
    roundMoney_(wcTotals.averageWorkingCapital || 0)
  ];

  for (let i = 0; i < kpiHeaders.length; i++) {
    const col = 1 + (i % 4) * 2;
    const row = 4 + Math.floor(i / 4) * 3;
    sheet.getRange(row, col, 1, 2).merge().setValue(kpiHeaders[i]).setFontWeight('bold').setBackground('#e8f0fe');
    sheet.getRange(row + 1, col, 1, 2).merge().setValue(kpiValues[i]).setFontSize(14).setFontWeight('bold').setNumberFormat('€#,##0.00');
  }

  let cursor = 11;
  sheet.getRange(cursor, 1).setValue('Поточний місяць: ' + latest.month).setFontWeight('bold').setFontSize(13);
  sheet.getRange(cursor + 1, 1, 1, 6).setValues([['Виплата Amazon', 'Повернення собівартості', 'Прибуток до НДС', 'НДС до оплати', 'Чистий прибуток після НДС', 'Реально вільний кеш']]).setFontWeight('bold').setBackground('#f3f3f3');
  sheet.getRange(cursor + 2, 1, 1, 6).setValues([[latest.payout, latest.cogs, latest.profitBeforeVat, latest.vatToPay, latest.cashAfterVat, latest.freeCash]]).setNumberFormat('€#,##0.00').setFontWeight('bold');

  cursor += 4;
  sheet.getRange(cursor, 1).setValue('СТРУКТУРА ВИПЛАТИ AMAZON').setFontWeight('bold').setFontSize(13);
  sheet.getRange(cursor + 1, 1, 1, 3).setValues([['Показник', 'Сума', '% у виплаті']]).setFontWeight('bold').setBackground('#d9ead3');
  const structureRows = [
    ['Виплата Amazon', latest.payout, latest.payout ? 100 : 0],
    ['Повернення собівартості', latest.cogs, percentOf_(latest.cogs, latest.payout)],
    ['Прибуток до НДС', latest.profitBeforeVat, percentOf_(latest.profitBeforeVat, latest.payout)],
    ['НДС до оплати', latest.vatToPay, percentOf_(latest.vatToPay, latest.payout)],
    ['Чистий прибуток після НДС', latest.cashAfterVat, percentOf_(latest.cashAfterVat, latest.payout)],
    ['Реально вільний кеш', latest.freeCash, percentOf_(latest.freeCash, latest.payout)]
  ];
  sheet.getRange(cursor + 2, 1, structureRows.length, 3).setValues(structureRows);
  sheet.getRange(cursor + 2, 2, structureRows.length, 1).setNumberFormat('€#,##0.00');
  sheet.getRange(cursor + 2, 3, structureRows.length, 1).setNumberFormat('0.00"%"');

  cursor += structureRows.length + 3;
  sheet.getRange(cursor, 1).setValue('КЕШФЛОУ ТА РЕЗЕРВИ').setFontWeight('bold').setFontSize(13);
  const cashflowRows = [
    ['Загальна виплата Amazon', totals.payout],
    ['Повернення собівартості', totals.cogs],
    ['Загальний прибуток до НДС', totals.profitBeforeVat],
    ['НДС до оплати', totals.vatToPay],
    ['Чистий прибуток після НДС', totals.cashAfterVat],
    ['Реально вільний кеш', totals.freeCash]
  ];
  sheet.getRange(cursor + 1, 1, 1, 2).setValues([['Показник', 'Сума']]).setFontWeight('bold').setBackground('#fff2cc');
  sheet.getRange(cursor + 2, 1, cashflowRows.length, 2).setValues(cashflowRows);
  sheet.getRange(cursor + 2, 2, cashflowRows.length, 1).setNumberFormat('€#,##0.00');

  cursor += cashflowRows.length + 3;
  sheet.getRange(cursor, 1).setValue('ПОВЕРНЕННЯ В ОБОРОТ').setFontWeight('bold').setFontSize(13);
  const returnRows = [
    ['Собівартість за весь період', totals.cogs],
    ['Собівартість за останній місяць', latest.cogs],
    ['Частка собівартості у виплаті', percentOf_(latest.cogs, latest.payout)],
    ['Скільки потрібно повернути в товар', latest.cogs],
    ['Частка прибутку у виплаті', percentOf_(latest.profitBeforeVat, latest.payout)]
  ];
  sheet.getRange(cursor + 1, 1, 1, 2).setValues([['Показник', 'Значення']]).setFontWeight('bold').setBackground('#fce5cd');
  sheet.getRange(cursor + 2, 1, returnRows.length, 2).setValues(returnRows);
  sheet.getRange(cursor + 2, 2, 2, 1).setNumberFormat('€#,##0.00');
  sheet.getRange(cursor + 4, 2, 2, 1).setNumberFormat('0.00"%"');

  cursor += returnRows.length + 3;
  sheet.getRange(cursor, 1).setValue('НДС РЕЗЕРВ').setFontWeight('bold').setFontSize(13);
  const vatReserveRows = [
    ['НДС нараховано', roundMoney_(totals.vatToPay + totals.vatPaid)],
    ['Вже сплачений НДС', totals.vatPaid],
    ['НДС до оплати', totals.vatToPay],
    ['Тимчасово використовується в обороті', totals.vatToPay],
    ['Безпечний кеш після резерву', totals.cashAfterVat]
  ];
  sheet.getRange(cursor + 1, 1, 1, 2).setValues([['Показник', 'Сума']]).setFontWeight('bold').setBackground('#ead1dc');
  sheet.getRange(cursor + 2, 1, vatReserveRows.length, 2).setValues(vatReserveRows);
  sheet.getRange(cursor + 2, 2, vatReserveRows.length, 1).setNumberFormat('€#,##0.00');
  const vatRiskCell = sheet.getRange(cursor + 4, 2);
  if (totals.vatToPay > 0) vatRiskCell.setBackground('#f4cccc').setFontColor('#b71c1c').setFontWeight('bold');
  else vatRiskCell.setBackground('#d9ead3').setFontColor('#1b5e20').setFontWeight('bold');

  cursor += vatReserveRows.length + 3;
  const allocatedByCategory = calculateDashboardFundAllocatedByCategory_(roundMoney_(totals.cashAfterVat));
  const spentByCategory = aggregateManualExpensesForDashboardByCategory_(manualByCategoryAgg);
  const latestMonthSpentByCategory = (manualByCategoryAgg && manualByCategoryAgg.byMonth && manualByCategoryAgg.byMonth[latest.month] && manualByCategoryAgg.byMonth[latest.month].categories)
    ? manualByCategoryAgg.byMonth[latest.month].categories
    : buildManualExpenseFundCategoryTotalsSkeleton_();
  latest.manualSpentReinvest = roundMoney_(latestMonthSpentByCategory['Реінвест (75%)'] || 0);

  sheet.getRange(cursor, 1).setValue('РОЗПОДІЛ ЧИСТОГО ПРИБУТКУ').setFontWeight('bold').setFontSize(13);
  sheet.getRange(cursor + 1, 1, 1, 2).setValues([['Категорія', 'Виділено з прибутку']]).setFontWeight('bold').setBackground('#fff2cc');
  const profitDistributionRows = calculateDashboardProfitDistributionRows_(allocatedByCategory);
  sheet.getRange(cursor + 2, 1, profitDistributionRows.length, 2).setValues(profitDistributionRows);
  sheet.getRange(cursor + 2, 2, profitDistributionRows.length, 1).setNumberFormat('€#,##0.00');

  cursor += profitDistributionRows.length + 3;
  sheet.getRange(cursor, 1).setValue('ФОНД РЕІНВЕСТУ').setFontWeight('bold').setFontSize(13);
  sheet.getRange(cursor + 1, 1, 1, 2).setValues([['Показник', 'Сума']]).setFontWeight('bold').setBackground('#d9ead3');
  const reinvestFundRows = calculateDashboardReinvestFundRows_(totals, allocatedByCategory, spentByCategory, latest);
  sheet.getRange(cursor + 2, 1, reinvestFundRows.length, 2).setValues(reinvestFundRows);
  sheet.getRange(cursor + 2, 2, reinvestFundRows.length, 1).setNumberFormat('€#,##0.00');
  const reinvestNegativeRows = getDashboardReinvestNegativeRows_(reinvestFundRows);
  for (let i = 0; i < reinvestNegativeRows.length; i++) {
    sheet.getRange(cursor + 2 + reinvestNegativeRows[i], 2).setBackground('#f4cccc').setFontColor('#b71c1c').setFontWeight('bold');
  }

  cursor += reinvestFundRows.length + 3;
  sheet.getRange(cursor, 1).setValue('ПРИБУТКОВІ ФОНДИ (БЕЗ РЕІНВЕСТУ)').setFontWeight('bold').setFontSize(13);
  sheet.getRange(cursor + 1, 1, 1, 4).setValues([['Категорія', 'Виділено', 'Витрачено', 'Залишилось']]).setFontWeight('bold').setBackground('#fce5cd');
  const profitBucketRows = calculateDashboardProfitBucketExpenseRows_(allocatedByCategory, spentByCategory);
  sheet.getRange(cursor + 2, 1, profitBucketRows.length, 4).setValues(profitBucketRows);
  sheet.getRange(cursor + 2, 2, profitBucketRows.length, 3).setNumberFormat('€#,##0.00');
  const profitBucketNegativeRows = getDashboardProfitBucketNegativeRows_(profitBucketRows);
  for (let i = 0; i < profitBucketNegativeRows.length; i++) {
    sheet.getRange(cursor + 2 + profitBucketNegativeRows[i], 4).setBackground('#f4cccc').setFontColor('#b71c1c').setFontWeight('bold');
  }

  cursor += profitBucketRows.length + 3;
  sheet.getRange(cursor, 1).setValue('КЕШФЛОУ ПО МІСЯЦЯХ').setFontWeight('bold').setFontSize(13);
  sheet.getRange(cursor + 1, 1, 1, 7).setValues([['Місяць', 'Виплата Amazon', 'Повернення собівартості', 'Прибуток до НДС', 'НДС до оплати', 'Чистий прибуток після НДС', 'Реально вільний кеш']]).setFontWeight('bold').setBackground('#cfe2f3');
  const monthRows = rows.map(function(row) {
    return [row.month, row.payout, row.cogs, row.profitBeforeVat, row.vatToPay, row.cashAfterVat, row.freeCash];
  });
  sheet.getRange(cursor + 2, 1, monthRows.length, 7).setValues(monthRows);
  sheet.getRange(cursor + 2, 1, monthRows.length, 1).setNumberFormat('@');
  sheet.getRange(cursor + 2, 2, monthRows.length, 6).setNumberFormat('€#,##0.00');

  cursor += monthRows.length + 3;
  const workingCapitalRows = (workingCapitalAgg && workingCapitalAgg.rows) ? workingCapitalAgg.rows.slice() : [];
  sheet.getRange(cursor, 1).setValue('Оборотний капітал').setFontWeight('bold').setFontSize(13);
  sheet.getRange(cursor + 1, 1, 1, 7).setValues([[
    'Місяць',
    'Собівартість товарів на Amazon',
    'Кошти доступні на вивід Amazon',
    'Товари готові до відправки',
    'Кошти на руках',
    'Можливість кредитування',
    'Оборотний капітал'
  ]]).setFontWeight('bold').setBackground('#cfe2f3');

  if (workingCapitalRows.length) {
    const wcTable = workingCapitalRows.map(function(row) {
      return [
        row.month,
        row.cogsAmazon,
        row.amazonPayoutAvailable,
        row.goodsReadyToShip,
        row.cashOnHand,
        row.creditCapacity,
        row.workingCapital
      ];
    });
    sheet.getRange(cursor + 2, 1, wcTable.length, 7).setValues(wcTable);
    sheet.getRange(cursor + 2, 1, wcTable.length, 1).setNumberFormat('@');
    sheet.getRange(cursor + 2, 2, wcTable.length, 6).setNumberFormat('€#,##0.00');
  } else {
    sheet.getRange(cursor + 2, 1).setValue('Немає активних даних у листі ' + getWorkingCapitalSheetName_() + '.').setFontColor('#777777');
  }

  if (dashboardWarnings && dashboardWarnings.length) {
    const warnStart = cursor + 2 + Math.max(workingCapitalRows.length, 1) + 2;
    sheet.getRange(warnStart, 1).setValue('Попередження').setFontWeight('bold').setFontSize(13);
    for (let i = 0; i < dashboardWarnings.length; i++) {
      sheet.getRange(warnStart + 1 + i, 1).setValue('• ' + dashboardWarnings[i]).setFontColor('#b71c1c');
    }
  }
}

function renderDashboardCharts_(sheet, rows, workingCapitalAgg) {
  if (!rows.length) return;
  const chartStartCol = 12;
  const cashflowHeaderRow = 2;
  sheet.getRange(cashflowHeaderRow, chartStartCol, 1, 4).setValues([['Місяць', 'Виплата Amazon', 'Повернення собівартості', 'Чистий прибуток після НДС']]);
  const cashflowRows = rows.map(function(row) { return [row.month, row.payout, row.cogs, row.cashAfterVat]; });
  sheet.getRange(cashflowHeaderRow + 1, chartStartCol, cashflowRows.length, 4).setValues(cashflowRows);

  const cashflowChart = sheet.newChart()
    .asLineChart()
    .addRange(sheet.getRange(cashflowHeaderRow, chartStartCol, cashflowRows.length + 1, 4))
    .setOption('title', 'Тенденція кешфлоу')
    .setOption('legend', { position: 'bottom' })
    .setPosition(2, 9, 0, 0)
    .build();
  sheet.insertChart(cashflowChart);

  const vatHeaderRow = cashflowHeaderRow + cashflowRows.length + 3;
  sheet.getRange(vatHeaderRow, chartStartCol, 1, 3).setValues([['Місяць', 'НДС до оплати', 'Вже сплачений НДС']]);
  const vatRows = rows.map(function(row) { return [row.month, row.vatToPay, row.vatPaid]; });
  sheet.getRange(vatHeaderRow + 1, chartStartCol, vatRows.length, 3).setValues(vatRows);

  const vatChart = sheet.newChart()
    .asLineChart()
    .addRange(sheet.getRange(vatHeaderRow, chartStartCol, vatRows.length + 1, 3))
    .setOption('title', 'Тенденція НДС резерву')
    .setOption('legend', { position: 'bottom' })
    .setPosition(18, 9, 0, 0)
    .build();
  sheet.insertChart(vatChart);

  const wcRows = (workingCapitalAgg && workingCapitalAgg.rows) ? workingCapitalAgg.rows : [];
  if (wcRows.length) {
    const startRow = vatHeaderRow + vatRows.length + 3;
    sheet.getRange(startRow, chartStartCol, 1, 2).setValues([['Місяць', 'Оборотний капітал']]);
    const wcChartRows = wcRows.map(function(row) { return [row.month, row.workingCapital]; });
    sheet.getRange(startRow + 1, chartStartCol, wcChartRows.length, 2).setValues(wcChartRows);
    sheet.getRange(startRow + 1, chartStartCol, wcChartRows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow + 1, chartStartCol + 1, wcChartRows.length, 1).setNumberFormat('€#,##0.00');

    const workingCapitalChart = sheet.newChart()
      .asLineChart()
      .addRange(sheet.getRange(startRow, chartStartCol, wcChartRows.length + 1, 2))
      .setOption('title', 'Тенденція оборотного капіталу')
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { title: 'Місяць' })
      .setOption('vAxis', { title: 'Оборотний капітал' })
      .setPosition(36, 9, 0, 0)
      .build();
    sheet.insertChart(workingCapitalChart);
  }
}
