/**
 * Warranty recheck automation for Google Sheets.
 * Paste this file into Extensions -> Apps Script and run setupWarrantyAutomation() once.
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
    number: ['№', 'No', 'Номер']
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
    const mainSheets = ss.getSheets().filter(s => parseSheetMonth_(s.getName(), 'main'));
    mainSheets.forEach(main => transferEligibleRows_(ss, main, today));
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

function transferEligibleRows_(ss, main, today) {
  const suffix = parseSheetMonth_(main.getName(), 'main');
  const warranty = findMonthSheet_(ss, 'warranty', suffix);
  if (!warranty) {
    log_(ss, 'ERROR', main.getName(), '', 'Немає гарантійного листа для місяця: ' + suffix);
    return;
  }

  const sourceMap = headerMap_(main);
  const destMap = headerMap_(warranty);
  const briefCol = requiredColumn_(sourceMap, WARRANTY_CFG.headers.brief, main.getName());
  const checkCol = requiredColumn_(sourceMap, WARRANTY_CFG.headers.check, main.getName());
  const deadlineCol = requiredColumn_(sourceMap, WARRANTY_CFG.headers.deadline, main.getName());
  const destBriefCol = requiredColumn_(destMap, WARRANTY_CFG.headers.brief, warranty.getName());
  const destCheckCol = requiredColumn_(destMap, WARRANTY_CFG.headers.check, warranty.getName());
  const endDateCol = requiredColumn_(destMap, WARRANTY_CFG.headers.endDate, warranty.getName());
  const lastRow = main.getLastRow();
  if (lastRow <= WARRANTY_CFG.headerRow) return;

  const count = lastRow - WARRANTY_CFG.headerRow;
  const deadlines = main.getRange(WARRANTY_CFG.headerRow + 1, deadlineCol, count, 1).getValues();
  const statuses = main.getRange(WARRANTY_CFG.headerRow + 1, checkCol, count, 1).getDisplayValues();
  const existingKeys = collectKeys_(warranty, destBriefCol);

  for (let i = 0; i < count; i++) {
    const row = WARRANTY_CFG.headerRow + 1 + i;
    const deadline = asDate_(deadlines[i][0]);
    if (!deadline) continue;
    const due = addDays_(startOfDay_(deadline, ss.getSpreadsheetTimeZone()), WARRANTY_CFG.warrantyDays);
    if (today.getTime() < due.getTime()) continue;
    if (!WARRANTY_CFG.eligibleMainStatuses.includes(norm_(statuses[i][0]))) continue;

    const key = briefKeyFromCell_(main.getRange(row, briefCol));
    if (!key) {
      log_(ss, 'SKIP', main.getName(), row, 'Порожнє або нерозпізнане ТЗ замовника');
      continue;
    }
    if (existingKeys.has(key)) continue;

    const destRow = Math.max(warranty.getLastRow() + 1, WARRANTY_CFG.headerRow + 1);
    copyMappedRow_(main, row, sourceMap, warranty, destRow, destMap);
    const statusCell = warranty.getRange(destRow, destCheckCol);
    ensureListValidationValues_(statusCell, [WARRANTY_CFG.needsReplacementStatus, WARRANTY_CFG.readyStatus]);
    statusCell.setValue(WARRANTY_CFG.needsReplacementStatus);
    warranty.getRange(destRow, endDateCol).setValue(addDays_(today, WARRANTY_CFG.replacementDays)).setNumberFormat('dd.MM.yyyy');
    existingKeys.add(key);
    log_(ss, 'TRANSFER', main.getName(), row, 'Перенесено в ' + warranty.getName() + ', рядок ' + destRow);
  }
}

function finishWarrantyRow_(ss, warranty, row) {
  const warrantyMap = headerMap_(warranty);
  const briefCol = requiredColumn_(warrantyMap, WARRANTY_CFG.headers.brief, warranty.getName());
  const briefCell = warranty.getRange(row, briefCol);
  const key = briefKeyFromCell_(briefCell);
  const url = urlFromCell_(briefCell);
  if (!key || !url) throw new Error('У рядку ' + row + ' не знайдено URL у колонці ТЗ замовника.');

  const suffix = parseSheetMonth_(warranty.getName(), 'warranty');
  const main = findMonthSheet_(ss, 'main', suffix);
  if (!main) throw new Error('Не знайдено основний лист для місяця: ' + suffix);
  const mainMap = headerMap_(main);
  const mainBriefCol = requiredColumn_(mainMap, WARRANTY_CFG.headers.brief, main.getName());
  const mainCheckCol = requiredColumn_(mainMap, WARRANTY_CFG.headers.check, main.getName());
  const matches = findRowsByKey_(main, mainBriefCol, key);
  if (matches.length !== 1) {
    throw new Error('Для ТЗ знайдено ' + matches.length + ' рядків у ' + main.getName() + '; статус не змінено, щоб не зачепити неправильний рядок.');
  }

  stampExternalBrief_(url, startOfDay_(new Date(), ss.getSpreadsheetTimeZone()));
  main.getRange(matches[0], mainCheckCol).setValue(WARRANTY_CFG.readyStatus);
  log_(ss, 'READY', warranty.getName(), row, 'ТЗ оновлено; статус синхронізовано з ' + main.getName());
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

function copyMappedRow_(src, srcRow, srcMap, dst, dstRow, dstMap) {
  Object.keys(dstMap).forEach(headerKey => {
    const dstCol = dstMap[headerKey];
    let srcCol = srcMap[headerKey];
    if (!srcCol && isAlias_(headerKey, WARRANTY_CFG.headers.endDate)) {
      srcCol = optionalColumn_(srcMap, WARRANTY_CFG.headers.mainComment);
    }
    if (srcCol) src.getRange(srcRow, srcCol).copyTo(dst.getRange(dstRow, dstCol), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
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
  for (let row = WARRANTY_CFG.headerRow + 1; row <= sheet.getLastRow(); row++) {
    const key = briefKeyFromCell_(sheet.getRange(row, col));
    if (key) set.add(key);
  }
  return set;
}

function findRowsByKey_(sheet, col, key) {
  const rows = [];
  for (let row = WARRANTY_CFG.headerRow + 1; row <= sheet.getLastRow(); row++) {
    if (briefKeyFromCell_(sheet.getRange(row, col)) === key) rows.push(row);
  }
  return rows;
}

function briefKeyFromCell_(cell) {
  const url = urlFromCell_(cell);
  return normalizeUrl_(url || cell.getDisplayValue());
}

function urlFromCell_(cell) {
  const rich = cell.getRichTextValue();
  if (rich) {
    if (rich.getLinkUrl()) return rich.getLinkUrl();
    for (const run of rich.getRuns()) if (run.getLinkUrl()) return run.getLinkUrl();
  }
  const formula = cell.getFormula();
  const mFormula = formula && formula.match(/^=HYPERLINK\(\s*[\"']([^\"']+)/i);
  if (mFormula) return mFormula[1];
  const mText = String(cell.getDisplayValue()).match(/https?:\/\/[^\s]+/i);
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
