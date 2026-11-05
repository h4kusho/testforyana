const SHEET_NAME = 'Ответы';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Тест на нас')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Запусти эту функцию один раз из редактора Apps Script,
// чтобы создать лист и заголовки.
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Сначала открой Apps Script из нужной Google Таблицы.');
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Дата завершения', 'Имя', 'Ответ 1', 'Ответ 2', 'Ответ 3',
      'Ответ 4', 'Ответ 5', 'Ответ 6 (сердечки)', 'Ответ 7',
      'Ответ 8', 'Ответ 9 (карта)', 'Выбранный приз',
      'Активированные купоны', 'Неактивированные купоны',
      'Тип пары', 'Финальный вопрос', 'Секретная концовка',
      'Счёт A', 'Счёт B', 'Счёт C', 'Начало теста'
    ]);
    sheet.setFrozenRows(1);
  }
  return 'Готово';
}

function submitTestResult(data) {
  if (!data || typeof data !== 'object') throw new Error('Некорректные данные теста.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Не найдена связанная Google Таблица.');
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    setup();
    sheet = ss.getSheetByName(SHEET_NAME);
  }

  const answers = data.answers || {};
  const answerText = (n) => {
    const item = answers[n];
    if (!item) return '';
    return String(item.answer == null ? '' : item.answer);
  };
  const coupons = Array.isArray(data.coupons) ? data.coupons : [];
  const activated = coupons.filter(c => c && c.activated).map(c => c.name).join(' | ');
  const notActivated = coupons.filter(c => c && !c.activated).map(c => c.name).join(' | ');
  const tally = data.tally || {};

  const row = [
    data.finishedAt ? new Date(data.finishedAt) : new Date(),
    String(data.name || 'Не указано').slice(0, 100),
    answerText(1), answerText(2), answerText(3), answerText(4),
    answerText(5), answerText(6), answerText(7), answerText(8),
    answerText(9), String(data.prize || ''),
    activated, notActivated, String(data.pairType || ''),
    answerText('final'), data.secret ? 'Да' : 'Нет',
    Number(tally.a || 0), Number(tally.b || 0), Number(tally.c || 0),
    data.startedAt ? new Date(data.startedAt) : ''
  ];

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}
