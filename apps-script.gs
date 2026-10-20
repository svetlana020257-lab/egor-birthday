// Приём ответов с сайта-приглашения в таблицу «ДР Егора 27.10 — ответы гостей».
// Если гость отвечает повторно с тем же именем, его строка обновляется.
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActive().getSheetByName('Ответы');
    var clean = function (v, n) { return String(v || '').replace(/^[=+\-@]/, "'$&").slice(0, n); };
    var name = clean(d.name, 80).trim();
    if (!name) return out({ ok: false });
    var row = [
      new Date(),
      name,
      d.going ? 'Да' : 'Нет',
      d.going && d.plus ? 'Да' : 'Нет',
      clean(d.plusName, 80),
      clean(d.adventure, 200),
      clean(d.msg, 500)
    ];
    var last = sheet.getLastRow();
    var target = last + 1;
    if (last > 1) {
      var names = sheet.getRange(2, 2, last - 1, 1).getValues();
      for (var i = 0; i < names.length; i++) {
        if (String(names[i][0]).trim().toLowerCase() === name.toLowerCase()) { target = i + 2; break; }
      }
    }
    sheet.getRange(target, 1, 1, row.length).setValues([row]);
    return out({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
