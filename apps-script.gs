// RONI × Сопрано — приёмник ответов анкеты (Google Таблица).
// Установка, 5 минут:
// 1. Создайте пустую Google Таблицу → Расширения → Apps Script → удалите всё и вставьте этот код.
// 2. В строке KEY впишите ключ доступа к дашборду (его даёт Claude, в репозиторий его не кладём).
// 3. Развернуть → Новое развёртывание → шестерёнка → «Веб-приложение».
//    Выполнять от имени: «Я». У кого есть доступ: «Все». → Развернуть → разрешить доступ.
// 4. Скопируйте URL веб-приложения (заканчивается на /exec) и пришлите Claude.

const KEY = 'ВСТАВЬТЕ_КЛЮЧ';

// Анкета присылает сюда ответы (POST)
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const row = JSON.parse(e.postData.contents);
    const head = sh.getLastRow() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    Object.keys(row).forEach(k => {
      if (head.indexOf(k) < 0) { head.push(k); sh.getRange(1, head.length).setValue(k); }
    });
    // апостроф = хранить как текст (иначе +998… станет числом или формулой)
    sh.appendRow(head.map(k => row[k] == null || row[k] === '' ? '' : "'" + String(row[k]).slice(0, 2000)));
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

// Дашборд забирает все ответы (GET ?key=…)
function doGet(e) {
  const out = obj => ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
  if (!e.parameter || e.parameter.key !== KEY) return out({ error: 'key' });
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sh.getLastRow() < 2) return out({ rows: [] });
  const v = sh.getDataRange().getDisplayValues();
  const head = v[0];
  const rows = v.slice(1).map(r => Object.fromEntries(head.map((h, i) => [h, r[i]])));
  return out({ rows });
}
