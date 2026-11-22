// Dán toàn bộ file này vào Extensions > Apps Script của bảng tính RSVP.
// Cột: A Thời gian | B Họ tên | C SĐT | D Tham dự | E Số người đi kèm | F Lời chúc | G Hiển thị lời chúc | H Thiệp (Nhà gái / Nhà trai)

var SHEET_ID = '1qFeV8LaljGGtwrhwOJkmTbCrm6RNzDI1YF7NmY0dTL8';

function sheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Cắt độ dài và chặn công thức (=, +, -, @ ở đầu ô)
function clean_(v, max) {
  var t = String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.website) return out_({ ok: true }); // ô ẩn chống spam: người thật không điền
    var name = clean_(d.name, 80);
    var phone = String(d.phone || '').replace(/[^\d+]/g, '').slice(0, 15);
    var attend = d.attend === 'Có tham dự' ? 'Có tham dự' : (d.attend === 'Không tham dự' ? 'Không tham dự' : '');
    var guests = Math.max(0, Math.min(4, parseInt(d.guests, 10) || 0));
    var msg = clean_(d.message, 500);
    var side = d.side === 'Nhà gái' ? 'Nhà gái' : (d.side === 'Nhà trai' ? 'Nhà trai' : '');
    if (!name || !phone || !attend) return out_({ error: 'Vui lòng điền đủ họ tên, số điện thoại và xác nhận tham dự.' });

    lock.waitLock(10000);
    var s = sheet_();
    var r = s.getLastRow() + 1;
    s.getRange(r, 3).setNumberFormat('@'); // giữ số 0 đầu của SĐT
    if (!s.getRange(1, 8).getValue()) s.getRange(1, 8).setValue('Thiệp');
    s.getRange(r, 1, 1, 8).setValues([[new Date(), name, phone, attend, guests, msg, msg ? 'Có' : '', side]]);
    return out_({ ok: true });
  } catch (err) {
    return out_({ error: 'Không thể gửi xác nhận, vui lòng thử lại.' });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'wishes') {
    var rows = sheet_().getDataRange().getValues().slice(1);
    var side = String((e.parameter && e.parameter.side) || '');
    var wishes = rows
      .filter(function (r) { return r[5] && r[6] === 'Có' && (!side || !r[7] || r[7] === side); })
      .map(function (r) {
        return { name: String(r[1]), message: String(r[5]), time: new Date(r[0]).getTime() };
      })
      .reverse()
      .slice(0, 100);
    return out_(wishes);
  }
  return out_({ ok: true });
}
