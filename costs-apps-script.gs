/* ══════════════════════════════════════════════════════════════
   Apps Script Web App — база затрат и ФОТ для costs.html

   ЧТО ДЕЛАЕТ
   - Хранит данные в двух листах таблицы 1tUsNmeRK64OAdj1lU72l7miFz6Oxpfs4aO1dLcBtty4:
       "Затраты_v2" — журнал затрат (создаётся автоматически при первом обращении)
       "ФОТ_v2"     — журнал ЗП по сотрудникам/направлениям (создаётся автоматически)
   - GET  ?action=list&sheet=costs|payroll            -> вернуть все строки
   - POST {action:'add', sheet:'costs'|'payroll', token, row:{...}}    -> добавить строку
   - POST {action:'delete', sheet:'costs'|'payroll', token, row:N}     -> удалить строку N (1-based, без заголовка)

   КАК РАЗВЕРНУТЬ
   1. Откройте таблицу https://docs.google.com/spreadsheets/d/1tUsNmeRK64OAdj1lU72l7miFz6Oxpfs4aO1dLcBtty4
   2. Расширения → Apps Script
   3. Удалите содержимое Code.gs, вставьте этот файл целиком
   4. Замените TOKEN ниже на свой произвольный секрет
   5. Деплой → Новый деплой → тип "Веб-приложение"
        - Выполнять от имени: Я (ваш аккаунт)
        - У кого есть доступ: Все (Anyone)
   6. Скопируйте URL веб-приложения и вставьте в costs.html (APPS_SCRIPT_URL),
      туда же впишите тот же TOKEN (APPS_SCRIPT_TOKEN)
   7. При обновлении кода — Деплой → Управление деплоями → Изменить → Новая версия
══════════════════════════════════════════════════════════════ */

const SHEET_ID = '1tUsNmeRK64OAdj1lU72l7miFz6Oxpfs4aO1dLcBtty4';
const COSTS_SHEET = 'Затраты_v2';
const PAYROLL_SHEET = 'ФОТ_v2';
const TOKEN = 'CHANGE_ME_SECRET'; // <-- поменяйте перед деплоем

const COSTS_HEADERS   = ['Дата','Направление','Поднаправление','Статья затрат','Описание','Сумма $','Сумма грн','Курс','Счёт','Примечание','Добавлено'];
const PAYROLL_HEADERS = ['Период','Сотрудник','Направление','Поднаправление','Сумма $','Примечание','Добавлено'];

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
    const sheetName = (params.sheet || body.sheet) === 'payroll' ? PAYROLL_SHEET : COSTS_SHEET;
    const headers = sheetName === PAYROLL_SHEET ? PAYROLL_HEADERS : COSTS_HEADERS;

    if (action === 'list') {
      const sheet = getSheet(sheetName, headers);
      const values = sheet.getDataRange().getValues();
      const rows = [];
      for (let i = 1; i < values.length; i++) {
        const r = values[i];
        if (r.every(c => c === '' || c === null)) continue;
        const obj = { _row: i }; // 1-based, без заголовка — для удаления
        headers.forEach((h, j) => obj[h] = formatCell(r[j]));
        rows.push(obj);
      }
      return json({ ok: true, rows });
    }

    if (action === 'add') {
      if ((params.token || body.token) !== TOKEN) return json({ ok: false, error: 'bad token' });
      const sheet = getSheet(sheetName, headers);
      const data = body.row || {};
      const row = headers.map(h => h === 'Добавлено' ? new Date() : (data[h] != null ? data[h] : ''));
      sheet.appendRow(row);
      return json({ ok: true });
    }

    if (action === 'delete') {
      if ((params.token || body.token) !== TOKEN) return json({ ok: false, error: 'bad token' });
      const sheet = getSheet(sheetName, headers);
      const rowIndex = Number(params.row || body.row);
      if (!rowIndex || rowIndex < 1) return json({ ok: false, error: 'bad row' });
      sheet.deleteRow(rowIndex + 1); // +1 за заголовок
      return json({ ok: true });
    }

    return json({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function getSheet(name, headers) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// даты приходят из getValues() как объекты Date — приводим к "YYYY-MM-DD" / "YYYY-MM"
function formatCell(v) {
  if (v instanceof Date) {
    const y = v.getFullYear(), m = String(v.getMonth() + 1).padStart(2, '0'), d = String(v.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return v;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
