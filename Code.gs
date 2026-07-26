/**
 * ==============================================================
 *  BACKEND NHẬN BÀI DỰ THI — Google Apps Script
 *  - Lưu file bài thi vào Google Drive (1 thư mục chỉ định)
 *  - Ghi thông tin người nộp vào Google Sheet (tab dữ liệu chính)
 *  - Quản lý "Tên cuộc thi / Ngày mở / Ngày kết thúc" qua admin.html
 *    (lưu trong 1 tab riêng tên "CaiDat")
 *
 *  Thứ tự cột tab dữ liệu chính: A=Stt | B=Họ và tên | C=Link file | D=Thời gian nộp
 * ==============================================================
 *
 * CÁCH CÀI ĐẶT:
 * 1. Tạo 1 Google Sheet mới, copy ID trong URL (giữa /d/ và /edit)
 *    -> dán vào SHEET_ID bên dưới.
 * 2. Tạo 1 thư mục Google Drive chứa bài dự thi, copy ID thư mục
 *    -> dán vào FOLDER_ID bên dưới.
 * 3. Đặt 1 mật khẩu quản trị bất kỳ -> dán vào ADMIN_PASSWORD bên dưới.
 * 4. Vào https://script.google.com -> New project -> xoá code mẫu,
 *    dán toàn bộ nội dung file này vào.
 * 5. Deploy > New deployment > "Web app":
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Bấm Deploy, cấp quyền (Authorize) khi được hỏi.
 * 6. Copy URL Web app (dạng .../exec) -> dán vào biến SCRIPT_URL
 *    trong CẢ 2 file: nop-bai-du-thi.html VÀ admin.html.
 *
 * Tab "CaiDat" sẽ được TỰ ĐỘNG tạo trong lần chạy đầu tiên, không
 * cần tạo tay. Admin chỉnh sửa qua trang admin.html, không cần vào
 * thẳng Google Sheet.
 * ==============================================================
 */

const SHEET_ID = '1oZwEdTrl_NDaeH5vbrAhRop13_coZyNKbLhjXL9n0u4';
const SHEET_NAME = 'Sheet1';       // tab chứa danh sách bài nộp
const SETTINGS_SHEET_NAME = 'CaiDat'; // tab chứa cấu hình cuộc thi (tự tạo)
const FOLDER_ID = '1w4QAeUANi9_M0e8kjGbAjgSmYwPaOPNU';
const ADMIN_PASSWORD = '123456';

// ============================== doGet ==============================
// Dùng để lấy cấu hình cuộc thi (tên, ngày mở, ngày kết thúc) — công khai,
// vì thông tin này vốn đã hiển thị sẵn trên trang nộp bài.
function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getSettings') {
    return jsonOutput(getSettings());
  }
  return jsonOutput({ status: 'error', message: 'Action không hợp lệ.' });
}

// ============================== doPost ==============================
// Xử lý 2 loại yêu cầu: nộp bài (mặc định) và cập nhật cấu hình (admin).
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || 'submit';

    if (action === 'updateSettings') {
      return updateSettings(data);
    }
    return submitEntry(data);

  } catch (err) {
    return jsonOutput({ status: 'error', message: err.toString() });
  }
}

// ---------------------- Nộp bài dự thi ----------------------
function submitEntry(data) {
  const fullName = (data.fullName || '').toString().trim();
  const fileName = (data.fileName || 'bai-du-thi').toString();
  const fileType = data.fileType || 'application/octet-stream';
  const fileData = data.fileData;

  if (!fullName || !fileData) {
    return jsonOutput({ status: 'error', message: 'Thiếu dữ liệu.' });
  }

  // Kiểm tra thời gian mở/đóng nộp bài
  const settings = getSettings();
  const now = new Date();
  if (settings.openDate && now < new Date(settings.openDate)) {
    return jsonOutput({ status: 'error', message: 'Cuộc thi chưa đến thời gian nhận bài.' });
  }
  if (settings.endDate && now > new Date(settings.endDate)) {
    return jsonOutput({ status: 'error', message: 'Cuộc thi đã kết thúc nhận bài.' });
  }

  // 1) Lưu file vào Google Drive
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const blob = Utilities.newBlob(
    Utilities.base64Decode(fileData),
    fileType,
    `${fullName} - ${fileName}`
  );
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 2) Ghi vào Google Sheet — đúng thứ tự cột A, B, C, D
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const stt = sheet.getLastRow(); // dòng 1 là tiêu đề -> dòng dữ liệu đầu tiên Stt = 1

  sheet.appendRow([
    stt,             // Cột A - Stt
    fullName,        // Cột B - Họ và tên
    file.getUrl(),   // Cột C - Link file trên Drive
    new Date()       // Cột D - Thời gian nộp
  ]);

  return jsonOutput({ status: 'success' });
}

// ---------------------- Cấu hình cuộc thi (admin) ----------------------
function getSettings() {
  const sheet = getOrCreateSettingsSheet();
  const values = sheet.getRange(2, 1, 1, 3).getValues()[0];
  return {
    status: 'success',
    contestName: values[0] ? values[0].toString() : '',
    openDate: values[1] ? new Date(values[1]).toISOString() : '',
    endDate: values[2] ? new Date(values[2]).toISOString() : ''
  };
}

function updateSettings(data) {
  if (data.password !== ADMIN_PASSWORD) {
    return jsonOutput({ status: 'error', message: 'Sai mật khẩu quản trị.' });
  }
  const sheet = getOrCreateSettingsSheet();
  sheet.getRange(2, 1, 1, 3).setValues([[
    (data.contestName || '').toString(),
    data.openDate ? new Date(data.openDate) : '',
    data.endDate ? new Date(data.endDate) : ''
  ]]);
  return jsonOutput({ status: 'success' });
}

function getOrCreateSettingsSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    sheet.getRange(1, 1, 1, 3).setValues([['Tên cuộc thi', 'Ngày mở', 'Ngày kết thúc']]);
    const now = new Date();
    const in30days = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
    sheet.getRange(2, 1, 1, 3).setValues([['Cuộc thi dự thi', now, in30days]]);
  }
  return sheet;
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
