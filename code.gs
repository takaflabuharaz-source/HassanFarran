/**
 * نظام إدارة مؤسسة حسن خالد حسن فران للمقاولات العامة
 * Backend Controller - آمن ويحفظ البيانات من المسح
 */

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'getAllData') {
    return ContentService.createTextOutput(JSON.stringify(getAllAppData()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (e && e.parameter && e.parameter.action === 'checkLogin') {
    return ContentService.createTextOutput(JSON.stringify(checkLogin(e.parameter.email, e.parameter.password)))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('نظام إدارة مؤسسة حسن خالد حسن فران للمقاولات')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var contents = JSON.parse(e.postData.contents);

    if (contents.action === 'saveRecord') {
      var res = saveRecord(contents.sheetName, contents.record);
      return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
    }

    if (contents.action === 'deleteRecord') {
      var res = deleteRecord(contents.sheetName, contents.id);
      return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
    }

    if (contents.action === 'uploadFile') {
      var res = uploadFileToDrive(contents.fileData);
      return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ "result": "success", "data": contents }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "result": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// دالة جلب البيانات الآمنة (تقرأ كافة البيانات الموجودة بالشيت بدون مسح)
function getAllAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const result = {};

  sheets.forEach(sheet => {
    const name = sheet.getName();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      result[name] = [];
      return;
    }

    const headers = data[0].map(h => h ? h.toString().trim() : '');
    const rows = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i].every(cell => cell === "" || cell === null)) continue;
      let obj = {};
      headers.forEach((h, index) => {
        if (!h) return;
        let val = data[i][index];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        obj[h] = (val !== undefined && val !== null) ? val : "";
      });
      rows.push(obj);
    }
    result[name] = rows;
  });

  return result;
}

// تسجيل الدخول المرن
function checkLogin(email, password) {
  if (String(email).trim().toLowerCase() === 'admin@farran.com' && String(password).trim() === '123456') {
    return { success: true, user: { name: 'م. حسن فران', email: email, role: 'مدير النظام' } };
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المستخدمين");
  if (!sheet) return { success: false, message: "جدول المستخدمين غير موجود" };

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: false, message: "لا يوجد مستخدمين مسجلين" };

  var headers = data[0];
  var emailIdx = headers.indexOf("البريد الإلكتروني");
  var passIdx = headers.indexOf("كلمة المرور");
  var roleIdx = headers.indexOf("الدور");
  var nameIdx = headers.indexOf("الاسم");

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][emailIdx]).trim().toLowerCase() === String(email).trim().toLowerCase() &&
        String(data[i][passIdx]).trim() === String(password).trim()) {
      return {
        success: true,
        user: { name: data[i][nameIdx], email: data[i][emailIdx], role: data[i][roleIdx] || 'مدير النظام' }
      };
    }
  }
  return { success: false, message: "بيانات الدخول غير صحيحة" };
}

// إضافة أو تحديث سجل دون المساس بباقي الجداول
function saveRecord(sheetName, record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const data = sheet.getDataRange().getValues();
  let headers = data.length > 0 ? data[0] : [];

  if (headers.length === 0) {
    headers = Object.keys(record);
    sheet.appendRow(headers);
  }

  if (record['ID']) {
    const idIndex = headers.indexOf('ID');
    if (idIndex !== -1) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][idIndex]) === String(record['ID'])) {
          headers.forEach((h, colIndex) => {
            if (record[h] !== undefined) {
              sheet.getRange(i + 1, colIndex + 1).setValue(record[h]);
            }
          });
          return { success: true, message: 'تم تحديث البيانات بنجاح' };
        }
      }
    }
  }

  record['ID'] = 'ID-' + new Date().getTime();
  const rowToAppend = headers.map(h => record[h] !== undefined ? record[h] : '');
  sheet.appendRow(rowToAppend);
  return { success: true, message: 'تم الحفظ بنجاح' };
}

// حذف سجل
function deleteRecord(sheetName, id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, message: 'الجدول غير موجود' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf('ID');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIndex]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'تم الحذف بنجاح' };
    }
  }
  return { success: false, message: 'لم يتم العثور على السجل' };
}

// رفع الملفات لـ Google Drive
function uploadFileToDrive(fileData) {
  try {
    const folderName = "فواتير ومستندات المؤسسة";
    let folders = DriveApp.getFoldersByName(folderName);
    let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    const contentType = fileData.mimeType || 'application/octet-stream';
    var base64String = fileData.data;
    if (base64String.indexOf(',') !== -1) {
      base64String = base64String.split(',')[1];
    }

    const bytes = Utilities.base64Decode(base64String);
    const blob = Utilities.newBlob(bytes, contentType, fileData.fileName);

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var previewUrl = "https://drive.google.com/file/d/" + file.getId() + "/preview";

    return { success: true, url: previewUrl, rawUrl: file.getUrl() };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
