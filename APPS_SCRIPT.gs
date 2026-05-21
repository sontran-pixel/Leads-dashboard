// ════════════════════════════════════════════════════════════
//  BookEasy – Google Apps Script
//  Dán toàn bộ code này vào Apps Script, xoá code cũ
//  Deploy → New deployment → Web app → Anyone
// ════════════════════════════════════════════════════════════

const SHEET_ID    = '1y8IuLW8XCqY838yAwJ_zmZMe_LGT3cvenPCliKNGINg';
const SHEET_NAME  = 'Trang tính1';
const NOTIFY_EMAIL = 'sontran81@gmail.com';

// ── POST: Ghi lịch mới hoặc cập nhật trạng thái ─────────────
function doPost(e) {
  try {
    const data  = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);

    // Cập nhật trạng thái (từ dashboard / admin)
    if (data.action === 'updateStatus') {
      const row = parseInt(data.rowIndex);
      sheet.getRange(row, 8).setValue(data.status); // cột H = Trạng thái
      return ok({ updated: row, status: data.status });
    }

    // Ghi lịch mới (từ booking-customer)
    sheet.appendRow([
      data.shopId        || '',
      data.customerName  || '',
      data.customerPhone || '',
      data.service       || '',
      data.date          || '',
      data.time          || '',
      data.note          || '',
      'Chờ xác nhận',
      new Date().toLocaleString('vi-VN'),
      data.price         || 0,
      data.duration      || ''
    ]);

    // Gửi email thông báo cho chủ tiệm
    sendEmailNotification(data);

    return ok({ saved: true });

  } catch(err) {
    return err_response(err.message);
  }
}

// ── GET: Lấy danh sách lịch ──────────────────────────────────
function doGet(e) {
  try {
    const sheet  = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const rows   = sheet.getDataRange().getValues();
    const shopId = (e && e.parameter && e.parameter.shopId) ? e.parameter.shopId : '';

    // Bỏ hàng header (row 0)
    const data = rows.slice(1)
      .filter(r => r[0] || r[1]) // bỏ hàng trống
      .filter(r => !shopId || r[0] === shopId)
      .map((r, i) => ({
        rowIndex:      i + 2, // 1-indexed + bỏ header
        shopId:        r[0]  || '',
        customerName:  r[1]  || '',
        customerPhone: r[2]  || '',
        service:       r[3]  || '',
        date:          r[4]  ? String(r[4]).slice(0, 10) : '',
        time:          r[5]  || '',
        note:          r[6]  || '',
        status:        r[7]  || 'Chờ xác nhận',
        createdAt:     r[8]  || '',
        price:         r[9]  || 0,
        duration:      r[10] || ''
      }));

    return ok({ data, total: data.length });

  } catch(err) {
    return err_response(err.message);
  }
}

// ── EMAIL THÔNG BÁO ──────────────────────────────────────────
function sendEmailNotification(data) {
  try {
    const subject = `📅 Lịch mới: [${data.shopId}] – ${data.customerName}`;
    const body = `
BookEasy – Lịch hẹn mới!

🏪 Tiệm     : ${data.shopId}
👤 Khách    : ${data.customerName}
📞 SĐT      : ${data.customerPhone}
💆 Dịch vụ : ${data.service}
📅 Ngày     : ${data.date}
🕐 Giờ      : ${data.time}
💰 Giá      : ${Number(data.price || 0).toLocaleString('vi-VN')}đ
📝 Ghi chú  : ${data.note || 'Không có'}

⏰ Đặt lúc  : ${new Date().toLocaleString('vi-VN')}

→ Truy cập Dashboard để xác nhận lịch hẹn.
    `.trim();

    GmailApp.sendEmail(NOTIFY_EMAIL, subject, body);
  } catch(err) {
    Logger.log('Email error: ' + err.message);
  }
}

// ── HELPERS ──────────────────────────────────────────────────
function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', ...data }))
    .setMimeType(ContentService.MimeType.JSON);
}
function err_response(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'error', message: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
