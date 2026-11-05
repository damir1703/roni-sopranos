// Подключение анкеты к Google Таблице (5 минут):
// 1. Создайте Google Таблицу → Расширения → Apps Script → вставьте этот код.
// 2. Развернуть → Новое развёртывание → Тип «Веб-приложение»,
//    «Выполнять от имени: я», «Доступ: все». Скопируйте URL.
// 3. Вставьте URL в anketa.html в строку  const ENDPOINT = '...';
function doPost(e) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const row = JSON.parse(e.postData.contents);
  let head = sh.getLastRow() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  Object.keys(row).forEach(k => { if (head.indexOf(k) < 0) { head.push(k); sh.getRange(1, head.length).setValue(k); } });
  sh.appendRow(head.map(k => row[k] ?? ''));
  return ContentService.createTextOutput('ok');
}
