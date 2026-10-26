/**
 * Warranty recheck automation for Google Sheets.
 * Paste this file into Extensions -> Apps Script and run WR_setupWarrantyAutomation_v1() once.
 */

const WR_WarrantyRecheckAutomation_v1 = (() => {

const WARRANTY_CFG = Object.freeze({
  headerRow: 1,
  mainPrefix: 'Морда',
  warrantyPrefix: 'Перечек по гарантії',
  headers: {
    brief: ['ТЗ замовника'],
    check: ['Перевірка'],
    deadline: ['Дедлайн', 'Deadline'],
    mainComment: ['Коментар Аня'],
    endDate: ['End data', 'End date'],
    number: ['№', 'No', 'Номер'],
    warrantyNumber: ['№', 'No', 'Номер', 'l']
  },
  eligibleMainStatuses: ['прийнято', 'перевірено'],
  needsReplacementStatus: 'Потрібно зробити заміни',
  readyStatus: 'Перечек готово',
  warrantyDays: 14,
  replacementDays: 7,
  tzHeaderScanRows: 10,
  tzHeaders: {
    donor: ['Donor'],
    textOrDate: ['Text', 'Date']
  },
  logSheet: '_Warranty log'
});

const WARRANTY_MONTH_NAMES = Object.freeze([
  'січень', 'лютий', 'березень', 'квітень', 'травень', 'червень',
  'липень', 'серпень', 'вересень', 'жовтень', 'листопад', 'грудень'
]);

function setup() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['WR_runDailyTransfer_v1', 'WR_installedOnEdit_v1'].includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('WR_runDailyTransfer_v1').timeBased().everyDays(1).atHour(6).create();
  ScriptApp.newTrigger('WR_installedOnEdit_v1').forSpreadsheet(ss).onEdit().create();
  ensureLogSheet_(ss);
  SpreadsheetApp.getUi().alert('Готово: створено щоденний та onEdit тригери.');
}

function runDailyTransfer() {
  withDocumentLock_(() => {
    const ss = SpreadsheetApp.getActive();
    const today = startOfDay_(new Date(), ss.getSpreadsheetTimeZone());
    const mainSheets = activeMainSheets_(ss, ss.getSpreadsheetTimeZone());
    const warrantyContexts = Object.create(null);
    mainSheets.forEach(main => transferEligibleRows_(ss, main, today, warrantyContexts));
  });
}

function installedOnEdit(e) {
  if (!e || !e.range || e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
  const sheet = e.range.getSheet();
  if (!parseSheetMonth_(sheet.getName(), 'warranty')) return;
  const map = headerMap_(sheet);
  const checkCol = requiredColumn_(map, WARRANTY_CFG.headers.check, sheet.getName());
  if (e.range.getRow() <= WARRANTY_CFG.headerRow || e.range.getColumn() !== checkCol) return;
  if (norm_(e.value) !== norm_(WARRANTY_CFG.readyStatus)) return;

  try {
    withDocumentLock_(() => finishWarrantyRow_(e.source, sheet, e.range.getRow()));
    e.range.clearNote();
  } catch (err) {
    e.range.setValue(e.oldValue || WARRANTY_CFG.needsReplacementStatus);
    e.range.setNote('Автоматизація не завершена: ' + err.message);
    log_(e.source, 'ERROR', sheet.getName(), e.range.getRow(), err.stack || err.message);
  }
}

function transferEligibleRows_(ss, main, today, warrantyContexts) {
  const sourceMap = headerMap_(main);
  const briefCol = requiredColumn_(sourceMap, WARRANTY_CFG.headers.brief, main.getName());
  const checkCol = requiredColumn_(sourceMap, WARRANTY_CFG.headers.check, main.getName());
  const deadlineCol = requiredColumn_(sourceMap, WARRANTY_CFG.headers.deadline, main.getName());
  const numberCol = requiredColumn_(sourceMap, WARRANTY_CFG.headers.number, main.getName());
  const mainCommentCol = requiredColumn_(sourceMap, WARRANTY_CFG.headers.mainComment, main.getName());
  const lastRow = main.getLastRow();
  if (lastRow <= WARRANTY_CFG.headerRow) return;

  const count = lastRow - WARRANTY_CFG.headerRow;
  const deadlines = main.getRange(WARRANTY_CFG.headerRow + 1, deadlineCol, count, 1).getValues();
  const statuses = main.getRange(WARRANTY_CFG.headerRow + 1, checkCol, count, 1).getDisplayValues();
  const projectNumbers = main.getRange(WARRANTY_CFG.headerRow + 1, numberCol, count, 1).getDisplayValues();
  const briefKeys = briefKeysFromRange_(main.getRange(WARRANTY_CFG.headerRow + 1, briefCol, count, 1));
  const candidates = [];

  for (let i = 0; i < count; i++) {
    const row = WARRANTY_CFG.headerRow + 1 + i;
    const deadline = asDate_(deadlines[i][0]);
    if (!deadline) continue;
    const due = addDays_(startOfDay_(deadline, ss.getSpreadsheetTimeZone()), WARRANTY_CFG.warrantyDays);
    if (today.getTime() < due.getTime()) continue;
    if (!WARRANTY_CFG.eligibleMainStatuses.includes(norm_(statuses[i][0]))) continue;

    const projectNumber = String(projectNumbers[i][0] == null ? '' : projectNumbers[i][0]).trim();
    if (!projectNumber) {
      log_(ss, 'SKIP', main.getName(), row, 'Порожній номер проєкту');
      continue;
    }
    // Номер із цифровим суфіксом /1, /2, ... означає частину батьківського проєкту.
    if (isSubprojectNumber_(projectNumber)) continue;

    const key = briefKeys[i];
    if (!key) {
      log_(ss, 'SKIP', main.getName(), row, 'Порожнє або нерозпізнане ТЗ замовника');
      continue;
    }
    candidates.push({
      row,
      key,
      deadline,
      due,
      projectNumber,
      targetMonth: monthNameFromDate_(deadline, ss.getSpreadsheetTimeZone())
    });
  }

  candidates.forEach(candidate => {
    try {
      const context = warrantyContext_(ss, candidate.targetMonth, warrantyContexts);
      if (context.existingKeys.has(candidate.key)) return;

      const destRow = context.nextRow;
      copyMappedRow_(main, candidate.row, sourceMap, context.sheet, destRow, context.destMap, [
        { sourceCol: numberCol, destCol: context.destNumberCol },
        { sourceCol: mainCommentCol, destCol: context.destMainCommentCol }
      ]);
      const statusCell = context.sheet.getRange(destRow, context.destCheckCol);
      ensureListValidationValues_(statusCell, [WARRANTY_CFG.needsReplacementStatus, WARRANTY_CFG.readyStatus]);
      statusCell.setValue(WARRANTY_CFG.needsReplacementStatus);
      const endDate = addDays_(candidate.due, WARRANTY_CFG.replacementDays);
      context.sheet.getRange(destRow, context.endDateCol).setValue(endDate).setNumberFormat('dd.MM.yyyy');
      context.existingKeys.add(candidate.key);
      context.nextRow = destRow + 1;
      log_(ss, 'TRANSFER', main.getName(), candidate.row, 'Перенесено в ' + context.sheet.getName() + ', рядок ' + destRow);
    } catch (err) {
      log_(ss, 'ERROR', main.getName(), candidate.row, err.stack || err.message);
    }
  });
}

function finishWarrantyRow_(ss, warranty, row) {
  const warrantyMap = headerMap_(warranty);
  const briefCol = requiredColumn_(warrantyMap, WARRANTY_CFG.headers.brief, warranty.getName());
  const briefCell = warranty.getRange(row, briefCol);
  const key = briefKeyFromCell_(briefCell);
  const url = urlFromCell_(briefCell);
  if (!key || !url) throw new Error('У рядку ' + row + ' не знайдено URL у колонці ТЗ замовника.');

  const matches = findMainRowsByKey_(ss, key);
  if (matches.length !== 1) {
    throw new Error('Для ТЗ знайдено ' + matches.length + ' рядків у основних листах; статус не змінено, щоб не зачепити неправильний рядок.');
  }

  stampExternalBrief_(url, startOfDay_(new Date(), ss.getSpreadsheetTimeZone()));
  const match = matches[0];
  const mainMap = headerMap_(match.sheet);
  const mainCheckCol = requiredColumn_(mainMap, WARRANTY_CFG.headers.check, match.sheet.getName());
  match.sheet.getRange(match.row, mainCheckCol).setValue(WARRANTY_CFG.readyStatus);
  log_(ss, 'READY', warranty.getName(), row, 'ТЗ оновлено; статус синхронізовано з ' + match.sheet.getName());
}

function stampExternalBrief_(url, date) {
  const tzSS = SpreadsheetApp.openByUrl(url);
  const sheet = tzSS.getSheets()[0];
  const scanRows = Math.min(WARRANTY_CFG.tzHeaderScanRows, Math.max(sheet.getLastRow(), 1));
  const scanCols = Math.max(sheet.getLastColumn(), 1);
  const values = sheet.getRange(1, 1, scanRows, scanCols).getDisplayValues();
  let headerRow = -1, donorCol = -1, dateCol = -1;
  for (let r = 0; r < values.length; r++) {
    const rowMap = mapFromHeaders_(values[r]);
    donorCol = optionalColumn_(rowMap, WARRANTY_CFG.tzHeaders.donor);
    dateCol = optionalColumn_(rowMap, WARRANTY_CFG.tzHeaders.textOrDate);
    if (donorCol && dateCol) { headerRow = r + 1; break; }
  }
  if (headerRow < 0) throw new Error('У ТЗ не знайдено в одному рядку заголовки Donor та Text/Date.');

  sheet.getRange(headerRow, dateCol).setValue('Date');
  const dataRows = sheet.getLastRow() - headerRow;
  if (dataRows <= 0) return;
  const dateRange = sheet.getRange(headerRow + 1, dateCol, dataRows, 1);
  dateRange.clearContent();
  const colors = sheet.getRange(headerRow + 1, donorCol, dataRows, 1).getBackgrounds();
  const output = colors.map(([hex]) => [isGreen_(hex) ? date : '']);
  dateRange.setValues(output).setNumberFormat('dd.MM.yyyy');
}

function copyMappedRow_(src, srcRow, srcMap, dst, dstRow, dstMap, explicitMappings) {
  const mappings = explicitMappings || [];
  const explicitDestCols = new Set(mappings.map(m => m.destCol));
  const copyPlan = [];
  Object.keys(dstMap).forEach(headerKey => {
    const dstCol = dstMap[headerKey];
    if (explicitDestCols.has(dstCol) || isAlias_(headerKey, WARRANTY_CFG.headers.endDate)) return;
    const srcCol = srcMap[headerKey];
    if (srcCol) copyPlan.push({ sourceCol: srcCol, destCol: dstCol });
  });
  mappings.forEach(m => copyPlan.push(m));
  copyPlan.sort((a, b) => a.destCol - b.destCol);

  const groups = [];
  copyPlan.forEach(m => {
    const last = groups[groups.length - 1];
    if (last && m.sourceCol === last.sourceCol + last.width && m.destCol === last.destCol + last.width) {
      last.width++;
    } else {
      groups.push({ sourceCol: m.sourceCol, destCol: m.destCol, width: 1 });
    }
  });
  groups.forEach(g => {
    src.getRange(srcRow, g.sourceCol, 1, g.width)
      .copyTo(dst.getRange(dstRow, g.destCol, 1, g.width), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  });
  dst.setRowHeight(dstRow, src.getRowHeight(srcRow));
}

function headerMap_(sheet) {
  return mapFromHeaders_(sheet.getRange(WARRANTY_CFG.headerRow, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0]);
}

function mapFromHeaders_(headers) {
  const map = {};
  headers.forEach((h, i) => { if (norm_(h)) map[norm_(h)] = i + 1; });
  return map;
}

function requiredColumn_(map, aliases, context) {
  const col = optionalColumn_(map, aliases);
  if (!col) throw new Error('Не знайдено колонку [' + aliases.join(' / ') + '] у ' + context);
  return col;
}

function optionalColumn_(map, aliases) {
  for (const alias of aliases) if (map[norm_(alias)]) return map[norm_(alias)];
  return 0;
}

function isAlias_(normalizedHeader, aliases) {
  return aliases.some(a => norm_(a) === normalizedHeader);
}

function collectKeys_(sheet, col) {
  const set = new Set();
  const count = sheet.getLastRow() - WARRANTY_CFG.headerRow;
  if (count <= 0) return set;
  briefKeysFromRange_(sheet.getRange(WARRANTY_CFG.headerRow + 1, col, count, 1))
    .forEach(key => { if (key) set.add(key); });
  return set;
}

function findRowsByKey_(sheet, col, key) {
  const rows = [];
  const count = sheet.getLastRow() - WARRANTY_CFG.headerRow;
  if (count <= 0) return rows;
  briefKeysFromRange_(sheet.getRange(WARRANTY_CFG.headerRow + 1, col, count, 1))
    .forEach((candidate, i) => {
      if (candidate === key) rows.push(WARRANTY_CFG.headerRow + 1 + i);
    });
  return rows;
}

function briefKeysFromRange_(range) {
  const richValues = range.getRichTextValues();
  const formulas = range.getFormulas();
  const displayValues = range.getDisplayValues();
  return displayValues.map((row, i) => normalizeUrl_(
    urlFromParts_(richValues[i][0], formulas[i][0], row[0]) || row[0]
  ));
}

function briefKeyFromCell_(cell) {
  const url = urlFromCell_(cell);
  return normalizeUrl_(url || cell.getDisplayValue());
}

function urlFromCell_(cell) {
  return urlFromParts_(cell.getRichTextValue(), cell.getFormula(), cell.getDisplayValue());
}

function urlFromParts_(rich, formula, displayValue) {
  if (rich) {
    if (rich.getLinkUrl()) return rich.getLinkUrl();
    for (const run of rich.getRuns()) if (run.getLinkUrl()) return run.getLinkUrl();
  }
  const mFormula = formula && formula.match(/^=HYPERLINK\(\s*[\"']([^\"']+)/i);
  if (mFormula) return mFormula[1];
  const mText = String(displayValue || '').match(/https?:\/\/[^\s]+/i);
  return mText ? mText[0] : '';
}

function normalizeUrl_(value) {
  const s = String(value || '').trim();
  const id = s.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return id ? id[1] : s.replace(/[?#].*$/, '').replace(/\/$/, '').toLowerCase();
}

function parseSheetMonth_(name, kind) {
  const prefix = kind === 'main' ? 'морда' : 'перечек по гарантії';
  const normalized = norm_(name).replace(/[–—]/g, '-');
  const match = normalized.match(new RegExp('^' + prefix + '\\s*(?:-\\s*)?(.+)$', 'i'));
  return match ? match[1].trim() : '';
}

function findMonthSheet_(ss, kind, month) {
  const wanted = norm_(month);
  return ss.getSheets().find(s => norm_(parseSheetMonth_(s.getName(), kind)) === wanted) || null;
}

function warrantyContext_(ss, month, contexts) {
  const key = norm_(month);
  if (contexts[key]) return contexts[key];

  const sheet = ensureWarrantySheet_(ss, month);
  const destMap = headerMap_(sheet);
  const context = {
    sheet,
    destMap,
    destBriefCol: requiredColumn_(destMap, WARRANTY_CFG.headers.brief, sheet.getName()),
    destCheckCol: requiredColumn_(destMap, WARRANTY_CFG.headers.check, sheet.getName()),
    endDateCol: requiredColumn_(destMap, WARRANTY_CFG.headers.endDate, sheet.getName()),
    destNumberCol: requiredColumn_(destMap, WARRANTY_CFG.headers.warrantyNumber, sheet.getName()),
    destMainCommentCol: requiredColumn_(destMap, WARRANTY_CFG.headers.mainComment, sheet.getName()),
    existingKeys: null,
    nextRow: Math.max(sheet.getLastRow() + 1, WARRANTY_CFG.headerRow + 1)
  };
  context.existingKeys = collectKeys_(sheet, context.destBriefCol);
  contexts[key] = context;
  return context;
}

function ensureWarrantySheet_(ss, month) {
  const existing = findMonthSheet_(ss, 'warranty', month);
  if (existing) return existing;

  const template = findWarrantyTemplateSheet_(ss, month);
  if (!template) {
    throw new Error('Не знайдено лист-шаблон для створення перечеку: ' + month);
  }

  const name = WARRANTY_CFG.warrantyPrefix + ' - ' + monthLabel_(month);
  const created = template.copyTo(ss);
  created.setName(name);
  created.showSheet();
  clearWarrantyData_(created);
  log_(ss, 'SHEET_CREATE', created.getName(), '', 'Створено автоматично з шаблону ' + template.getName());
  return created;
}

function findWarrantyTemplateSheet_(ss, targetMonth) {
  const sheets = ss.getSheets().filter(s => parseSheetMonth_(s.getName(), 'warranty'));
  if (!sheets.length) return null;

  const targetNumber = monthNumber_(targetMonth);
  if (targetNumber) {
    const previous = sheets
      .filter(s => {
        const number = monthNumber_(parseSheetMonth_(s.getName(), 'warranty'));
        return number && number <= targetNumber;
      })
      .sort((a, b) => monthNumber_(parseSheetMonth_(a.getName(), 'warranty')) - monthNumber_(parseSheetMonth_(b.getName(), 'warranty')) || a.getIndex() - b.getIndex());
    if (previous.length) return previous[previous.length - 1];
  }
  return sheets[sheets.length - 1];
}

function clearWarrantyData_(sheet) {
  const rowCount = Math.max(sheet.getLastRow() - WARRANTY_CFG.headerRow, 1);
  const dataRange = sheet.getRange(WARRANTY_CFG.headerRow + 1, 1, rowCount, Math.max(sheet.getMaxColumns(), 1));
  dataRange.clearContent();
  dataRange.clearNote();
}

function activeMainSheets_(ss, tz) {
  const currentMonth = Number(Utilities.formatDate(new Date(), tz, 'M'));
  return ss.getSheets()
    .filter(s => parseSheetMonth_(s.getName(), 'main') && !s.isSheetHidden())
    .sort((a, b) => {
      const aMonth = monthNumber_(parseSheetMonth_(a.getName(), 'main'));
      const bMonth = monthNumber_(parseSheetMonth_(b.getName(), 'main'));
      return relativeMonthDistance_(aMonth, currentMonth) - relativeMonthDistance_(bMonth, currentMonth) || a.getIndex() - b.getIndex();
    });
}

function relativeMonthDistance_(month, currentMonth) {
  if (!month || !currentMonth) return 99;
  let distance = month - currentMonth;
  while (distance > 6) distance -= 12;
  while (distance < -6) distance += 12;
  return distance;
}

function findMainRowsByKey_(ss, key) {
  const matches = [];
  ss.getSheets().filter(s => parseSheetMonth_(s.getName(), 'main')).forEach(main => {
    const map = headerMap_(main);
    const briefCol = optionalColumn_(map, WARRANTY_CFG.headers.brief);
    if (!briefCol) return;
    findRowsByKey_(main, briefCol, key).forEach(row => matches.push({ sheet: main, row }));
  });
  return matches;
}

function monthNameFromDate_(date, tz) {
  const month = Number(Utilities.formatDate(new Date(date), tz, 'M'));
  return WARRANTY_MONTH_NAMES[month - 1] || '';
}

function monthLabel_(month) {
  const value = norm_(month);
  return value ? value.charAt(0).toUpperCase() + value.slice(1) : '';
}

function monthNumber_(month) {
  const index = WARRANTY_MONTH_NAMES.indexOf(norm_(month));
  return index < 0 ? 0 : index + 1;
}

function isGreen_(hex) {
  const m = String(hex || '').match(/^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i);
  if (!m) return false;
  const r = parseInt(m[1], 16), g = parseInt(m[2], 16), b = parseInt(m[3], 16);
  return g >= 90 && g >= r + 10 && g >= b + 10;
}

function ensureListValidationValues_(cell, requiredValues) {
  const rule = cell.getDataValidation();
  if (!rule || rule.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return;
  const args = rule.getCriteriaValues();
  const values = (args[0] || []).slice();
  const known = new Set(values.map(norm_));
  requiredValues.forEach(v => { if (!known.has(norm_(v))) values.push(v); });
  cell.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build());
}

function norm_(v) { return String(v == null ? '' : v).trim().toLowerCase().replace(/\s+/g, ' '); }
function isSubprojectNumber_(v) { return /\/\s*\d+\s*$/.test(String(v == null ? '' : v).trim()); }
function asDate_(v) { const d = v instanceof Date ? new Date(v) : new Date(v); return isNaN(d.getTime()) ? null : d; }
function addDays_(d, days) { const out = new Date(d); out.setDate(out.getDate() + days); return out; }
function startOfDay_(d, tz) {
  const parts = Utilities.formatDate(new Date(d), tz, 'yyyy,MM,dd').split(',').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function withDocumentLock_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function ensureLogSheet_(ss) {
  let s = ss.getSheetByName(WARRANTY_CFG.logSheet);
  if (!s) s = ss.insertSheet(WARRANTY_CFG.logSheet);
  if (s.getLastRow() === 0) s.appendRow(['Timestamp', 'Type', 'Sheet', 'Row', 'Message']);
  s.hideSheet();
  return s;
}

function log_(ss, type, sheet, row, message) {
  ensureLogSheet_(ss).appendRow([new Date(), type, sheet, row, message]);
}

return Object.freeze({ setup, runDailyTransfer, installedOnEdit });
})();

/** Unique public entry points. Do not rename: installed triggers refer to them. */
function WR_setupWarrantyAutomation_v1() {
  return WR_WarrantyRecheckAutomation_v1.setup();
}

function WR_runDailyTransfer_v1() {
  return WR_WarrantyRecheckAutomation_v1.runDailyTransfer();
}

function WR_installedOnEdit_v1(e) {
  return WR_WarrantyRecheckAutomation_v1.installedOnEdit(e);
}
