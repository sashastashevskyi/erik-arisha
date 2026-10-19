/**
 * СКРИПТ ДЛЯ ЗБОРУ RSVP У GOOGLE ТАБЛИЦЮ · Эрик & Арина
 * ------------------------------------------------------------
 * Куди вставляти: Google Таблиця → Розширення → Apps Script
 *
 * Що робить:
 *  - аркуш «RSVP»: кольорова шапка, чергування рядків, підсвітка
 *    відповідей (зелена — прийде, теракотова — не зможе), час Каліфорнії
 *  - аркуш «Итог»: автоматичний підрахунок гостей (формули)
 *  - лист на пошту про кожну нову відповідь + поточні підсумки
 *    (якщо NOTIFY_EMAIL порожній — листи не надсилаються)
 *
 * Після вставки:
 *  1. Вкажіть NOTIFY_EMAIL (пошта, куди приходитимуть відповіді)
 *  2. У редакторі оберіть функцію setup і натисніть «Виконати» (Run) —
 *     обидва аркуші створяться одразу, з оформленням (дозвольте доступ)
 *  3. Deploy → New deployment → тип Web app
 *       Execute as: Me · Who has access: Anyone → Deploy
 *  4. Скопіюйте URL (…/exec) у index.html → RSVP_CONFIG.googleSheetsWebhook
 *  5. Після будь-яких змін у скрипті: Manage deployments → ✏️ → New version
 */

const SHEET_NAME = 'RSVP';
const SUMMARY_NAME = 'Итог';
const NOTIFY_EMAIL = 'erik.trufik@gmail.com';

// Палітра в стилі запрошення
const C_HEADER_BG   = '#5F6438'; // олива
const C_HEADER_TEXT = '#ffffff';
const C_BAND_1      = '#F7F4EC'; // кістяний
const C_BAND_2      = '#ECE6D6';
const C_YES         = '#e3edd5'; // м'яка зелень — прийде
const C_NO          = '#f5e0d7'; // м'який теракот — не зможе

/** Разова ініціалізація: створює й оформлює обидва аркуші. Запустіть вручну. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('America/Los_Angeles');
  if (!ss.getSheetByName(SHEET_NAME)) createRsvpSheet(ss);
  if (!ss.getSheetByName(SUMMARY_NAME)) createSummarySheet(ss);
}

function doPost(e) {
  // блокування, щоб одночасні відповіді гостей не перезаписували одна одну
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || createRsvpSheet(ss);
    if (!ss.getSheetByName(SUMMARY_NAME)) createSummarySheet(ss);

    const data = JSON.parse(e.postData.contents);
    const isComing = String(data.attendance || '').indexOf('Придёт') === 0;
    const adults = Number(data.adults) || 0;
    const children = Number(data.children) || 0;

    sheet.appendRow([
      new Date(),
      String(data.name || '').trim(),
      data.attendance || '',
      isComing ? adults : '',
      isComing ? children : '',
      String(data.wishes || '').trim()
    ]);

    // поточні підсумки — для листа
    const t = totals(sheet);
    if (NOTIFY_EMAIL) MailApp.sendEmail(
      NOTIFY_EMAIL,
      (isComing ? '💌 ' : '🙁 ') + 'Новый ответ: ' + data.name,
      'Имя: ' + data.name + '\n' +
      'Присутствие: ' + data.attendance + '\n' +
      (isComing ? 'Взрослых: ' + adults + ', детей: ' + children + '\n' : '') +
      (data.wishes ? 'Пожелание: ' + data.wishes + '\n' : '') +
      '\n— Всего на данный момент —\n' +
      '✅ Подтвердили: ' + t.yes + ' (взрослых ' + t.adults + ', детей ' + t.children + ')\n' +
      '🙁 Не смогут: ' + t.no
    );

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/** Створює й оформлює аркуш RSVP */
function createRsvpSheet(ss) {
  const sheet = ss.insertSheet(SHEET_NAME, 0);
  const headers = ['Дата ответа', 'Имя', 'Присутствие', 'Взрослых', 'Детей', 'Пожелание'];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(C_HEADER_BG)
    .setFontColor(C_HEADER_TEXT)
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 42);
  sheet.setFrozenRows(1);

  // ширини колонок
  [150, 210, 180, 95, 75, 430].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // формати даних
  sheet.getRange('A2:A').setNumberFormat('dd.mm.yyyy  hh:mm');
  sheet.getRange('C2:E').setHorizontalAlignment('center');
  sheet.getRange('F2:F').setWrap(true);
  sheet.getRange('A2:F').setVerticalAlignment('middle');

  // чергування кольорів рядків
  const banding = sheet.getRange('A2:F1000').applyRowBanding();
  banding.setFirstRowColor(C_BAND_1).setSecondRowColor(C_BAND_2);

  // підсвітка присутності
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Придёт')
      .setBackground(C_YES)
      .setRanges([sheet.getRange('C2:C1000')])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Не сможет')
      .setBackground(C_NO)
      .setRanges([sheet.getRange('C2:C1000')])
      .build()
  ]);

  // прибрати зайві колонки та сітку — чистий вигляд
  if (sheet.getMaxColumns() > headers.length) {
    sheet.deleteColumns(headers.length + 1, sheet.getMaxColumns() - headers.length);
  }
  sheet.setHiddenGridlines(true);

  return sheet;
}

/** Створює аркуш «Итог» з формулами — оновлюється сам */
function createSummarySheet(ss) {
  const s = ss.insertSheet(SUMMARY_NAME, 1);

  s.getRange('B2:C2').merge();
  s.getRange('B2').setValue('💍 Итог ответов')
    .setFontSize(14).setFontWeight('bold')
    .setFontColor(C_HEADER_BG).setHorizontalAlignment('center');

  // формули без роздільників аргументів — працюють у будь-якій локалі.
  // Колонка D заповнюється лише коли гість прийде, тому COUNT(D) = к-сть «так»
  const rows = [
    ['Всего ответов',            '=COUNTA(RSVP!B2:B)'],
    ['✅ Придут (ответов)',      '=COUNT(RSVP!D2:D)'],
    ['🙁 Не смогут',             '=COUNTA(RSVP!B2:B)-COUNT(RSVP!D2:D)'],
    ['Взрослых',                 '=SUM(RSVP!D2:D)'],
    ['Детей',                    '=SUM(RSVP!E2:E)'],
    ['🥂 Всего гостей',          '=SUM(RSVP!D2:E)']
  ];
  rows.forEach((r, i) => {
    s.getRange(4 + i, 2).setValue(r[0]);
    s.getRange(4 + i, 3).setFormula(r[1]);
  });

  const table = s.getRange(4, 2, rows.length, 2);
  table.setVerticalAlignment('middle');
  s.getRange(4, 2, rows.length, 1).setFontWeight('bold');
  s.getRange(4, 3, rows.length, 1).setHorizontalAlignment('center');
  for (let i = 0; i < rows.length; i++) {
    s.setRowHeight(4 + i, 32);
    s.getRange(4 + i, 2, 1, 2).setBackground(i % 2 ? C_BAND_2 : C_BAND_1);
  }
  // головний рядок — разом гостей
  s.getRange(4 + rows.length - 1, 2, 1, 2)
    .setBackground(C_HEADER_BG).setFontColor(C_HEADER_TEXT).setFontWeight('bold');

  s.setColumnWidth(1, 30);
  s.setColumnWidth(2, 230);
  s.setColumnWidth(3, 110);
  if (s.getMaxColumns() > 4) s.deleteColumns(5, s.getMaxColumns() - 4);
  s.setHiddenGridlines(true);

  return s;
}

/** Рахує поточні підсумки по аркушу RSVP */
function totals(sheet) {
  const last = sheet.getLastRow();
  const t = { yes: 0, no: 0, adults: 0, children: 0 };
  if (last < 2) return t;
  const vals = sheet.getRange(2, 3, last - 1, 3).getValues();
  vals.forEach(v => {
    if (String(v[0]).indexOf('Придёт') === 0) {
      t.yes++;
      t.adults += Number(v[1]) || 0;
      t.children += Number(v[2]) || 0;
    } else if (v[0]) {
      t.no++;
    }
  });
  return t;
}
