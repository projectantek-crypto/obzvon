/**
 * База статусов для страницы «Обзвон ЕГРЗ».
 * Вставьте этот код в Расширения → Apps Script вашей Google Таблицы,
 * поменяйте пароль ниже и разверните как веб-приложение (инструкция в README).
 */
const PASSWORD = 'поменяйте-этот-пароль';

// Адрес вашей Google Таблицы (скопируйте из адресной строки браузера).
// Нужен, если скрипт создан не через меню таблицы «Расширения → Apps Script».
const SPREADSHEET_URL = '';

const SHEET = 'Статусы';
const HEAD = ['ИНН', 'Компания', 'Телефон', 'ЛПР', 'Статус', 'Следующее касание', 'Заметка', 'Обновлено', 'Менеджер', 'История'];
const TZ = 'Europe/Moscow';

function sheet_() {
  const ss = SPREADSHEET_URL ? SpreadsheetApp.openByUrl(SPREADSHEET_URL) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Скрипт не видит таблицу: впишите SPREADSHEET_URL в начале кода');
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange('A:J').setNumberFormat('@');           // всё храним как текст, чтобы ИНН и даты не портились
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 1, 110); sh.setColumnWidth(2, 320); sh.setColumnWidth(7, 320); sh.setColumnWidth(10, 360);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try { return doGet_(e); } catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); }
}

function doGet_(e) {
  if ((e.parameter.key || '') !== PASSWORD) return json_({ ok: false, error: 'bad_key' });
  const sh = sheet_();
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, HEAD.length).getDisplayValues() : [];
  return json_({ ok: true, rows: rows });
}

function doPost(e) {
  try { return doPost_(e); } catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); }
}

function doPost_(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  if ((b.key || '') !== PASSWORD) return json_({ ok: false, error: 'bad_key' });
  const inn = String(b.inn || '').replace(/\D/g, '');
  if (!inn) return json_({ ok: false, error: 'no_inn' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
    const found = sh.getRange('A:A').createTextFinder(inn).matchEntireCell(true).findNext();
    let rowIdx, old;
    if (found) {
      rowIdx = found.getRow();
      old = sh.getRange(rowIdx, 1, 1, HEAD.length).getDisplayValues()[0];
    } else {
      rowIdx = sh.getLastRow() + 1;
      old = ['', '', '', '', 'Новый', '', '', '', '', ''];
    }
    const status = String(b.status || old[4] || 'Новый');
    let history = old[9];
    if (status !== old[4]) history = (history ? history + '\n' : '') + now + ' · ' + (b.by || '') + ' · ' + status;
    const row = [inn, b.company || old[1], b.phone || old[2], b.lpr || old[3], status,
                 String(b.next || ''), String(b.note || ''), now, String(b.by || ''), history];
    sh.getRange(rowIdx, 1, 1, HEAD.length).setNumberFormat('@').setValues([row]);
    return json_({ ok: true, row: row });
  } finally {
    lock.releaseLock();
  }
}
