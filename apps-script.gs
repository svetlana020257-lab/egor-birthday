// Приём ответов с сайта-приглашения в таблицу «ДР Егора 27.10 — ответы гостей».
// Если гость отвечает повторно с тем же именем, его строка обновляется.
// Фото гостей сохраняются в папку на Google Диске и показываются на сайте,
// если в колонке «Показывать на сайте» стоит «Да».

var FOLDER_NAME = 'ДР Егора 27.10 — фото гостей';

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var d = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActive().getSheetByName('Ответы');
    var clean = function (v, n) { return String(v || '').replace(/^[=+\-@]/, "'$&").slice(0, n); };
    var name = clean(d.name, 80).trim();
    if (!name) return out({ ok: false });

    var last = sheet.getLastRow();
    var target = last + 1;
    if (last > 1) {
      var names = sheet.getRange(2, 2, last - 1, 1).getValues();
      for (var i = 0; i < names.length; i++) {
        if (String(names[i][0]).trim().toLowerCase() === name.toLowerCase()) { target = i + 2; break; }
      }
    }

    var row = [
      new Date(),
      name,
      d.going ? 'Да' : 'Нет',
      d.going && d.plus ? 'Да' : 'Нет',
      clean(d.plusName, 80),
      clean(d.adventure, 300),
      clean(d.msg, 500)
    ];
    sheet.getRange(target, 1, 1, row.length).setValues([row]);

    if (d.photo && String(d.photo).indexOf('data:image/') === 0) {
      var m = String(d.photo).match(/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/);
      if (m) {
        var bytes = Utilities.base64Decode(m[2]);
        if (bytes.length <= 8 * 1024 * 1024) {
          var blob = Utilities.newBlob(bytes, m[1], name + '.jpg');
          var file = folder().createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          sheet.getRange(target, 8, 1, 2).setValues([[file.getUrl(), 'Да']]);
        }
      }
    }
    return out({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Список фото гостей для сайта
function doGet() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Ответы');
  var last = sheet.getLastRow();
  var list = [];
  if (last > 1) {
    var rows = sheet.getRange(2, 1, last - 1, 9).getValues();
    rows.forEach(function (r) {
      var url = String(r[7] || '');
      var show = String(r[8] || '').trim().toLowerCase() === 'да';
      var id = (url.match(/[-\w]{25,}/) || [])[0];
      if (id && show) list.push({ id: id, name: String(r[1]), story: String(r[5]), at: r[0] instanceof Date ? r[0].getTime() : 0 });
    });
  }
  list.sort(function (a, b) { return a.at - b.at; });
  return out({ photos: list });
}

function folder() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (err) {} }
  var f = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
