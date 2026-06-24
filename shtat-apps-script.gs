/* ══════════════════════════════════════════════════════════════
   Apps Script Web App — Штатное расписание для shtat.html

   ЧТО ДЕЛАЕТ
   - Хранит всё штатное расписание (отделы, сотрудники, направления бизнеса)
     как ОДИН JSON-документ в листе "Штат_данные" (ячейка A2), плюс
     "Штат_история" — последние версии для отката (на случай ошибки).
   - GET  ?action=load                          -> вернуть текущий документ
   - POST {action:'save', token, data:{...}}     -> сохранить документ целиком

   КАК РАЗВЕРНУТЬ
   1. Откройте таблицу https://docs.google.com/spreadsheets/d/1LCJeCeJLQ05sPWLz_7MoO5Wab4nRJhO4PlXrxy1fQ7U
   2. Расширения → Apps Script
   3. Создайте новый файл (или используйте Code.gs), вставьте этот файл целиком
   4. Замените TOKEN ниже НА ТОТ ЖЕ ПАРОЛЬ, что стоит на странице shtat.html (переменная PWD
      в самом начале <head> shtat.html). Токена-константы в shtat.html больше нет —
      на запись отправляется пароль, который ввёл пользователь при входе. Если потребуется
      сменить пароль — меняйте его в ДВУХ местах: PWD в shtat.html и TOKEN здесь.
   5. Деплой → Новый деплой → тип "Веб-приложение"
        - Выполнять от имени: Я (ваш аккаунт)
        - У кого есть доступ: Все (Anyone)
   6. Скопируйте URL веб-приложения и вставьте в shtat.html (APPS_SCRIPT_URL)
   7. При обновлении кода — Деплой → Управление деплоями → Изменить → Новая версия
══════════════════════════════════════════════════════════════ */

const SHEET_ID = '1LCJeCeJLQ05sPWLz_7MoO5Wab4nRJhO4PlXrxy1fQ7U';
const DATA_SHEET = 'Штат_данные';
const HISTORY_SHEET = 'Штат_история';
const TOKEN = 'CHANGE_ME_SECRET'; // <-- поменяйте на пароль страницы (PWD из shtat.html) перед деплоем
const MAX_HISTORY = 30;

function doGet(e)  { return handle(e); }
function doPost(e) { return handle(e); }

function handle(e) {
  try {
    const params = (e && e.parameter) || {};
    let body = {};
    if (e && e.postData && e.postData.contents) {
      try { body = JSON.parse(e.postData.contents); } catch (err) {}
    }
    const action = params.action || body.action;

    if (action === 'load') {
      const sheet = getDataSheet();
      const raw = sheet.getRange('A2').getValue();
      const data = raw ? JSON.parse(raw) : null;
      const updatedAt = sheet.getRange('B2').getValue() || '';
      return json({ ok: true, data, updatedAt });
    }

    if (action === 'save') {
      if ((params.token || body.token) !== TOKEN) return json({ ok: false, error: 'bad token' });
      const data = body.data;
      if (!data) return json({ ok: false, error: 'no data' });
      const sheet = getDataSheet();
      const now = new Date();
      pushHistory(sheet.getRange('A2').getValue());
      sheet.getRange('A2').setValue(JSON.stringify(data));
      sheet.getRange('B2').setValue(now);
      return json({ ok: true, updatedAt: now.toISOString() });
    }

    return json({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function getDataSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DATA_SHEET);
    sheet.getRange('A1').setValue('JSON документ');
    sheet.getRange('B1').setValue('Обновлено');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 600);
  }
  return sheet;
}

// сохраняет предыдущую версию документа в "Штат_история" перед перезаписью
function pushHistory(prevValue) {
  if (!prevValue) return;
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(HISTORY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(HISTORY_SHEET);
    sheet.appendRow(['Дата', 'JSON документ']);
    sheet.setFrozenRows(1);
  }
  sheet.insertRowAfter(1);
  sheet.getRange(2, 1).setValue(new Date());
  sheet.getRange(2, 2).setValue(prevValue);
  const last = sheet.getLastRow();
  if (last > MAX_HISTORY + 1) {
    sheet.deleteRows(MAX_HISTORY + 2, last - MAX_HISTORY - 1);
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
