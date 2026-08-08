// ==========================================
// نظام إدارة المؤسسة - Google Apps Script
// ==========================================

// إعدادات الشيتات والمجلدات
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const FOLDER_ID = ''; // ضع ID مجلد Google Drive المخصص لرفع المستندات والفواتير هنا إذا رغبت، أو اتركه فارغاً للحفظ في المجلد الرئيسي

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('نظام إدارة مؤسسة حسن خالد حسن فران للمقاولات')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// دالة جلب البيانات العامة من أي شيت
function getData(sheetName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const headers = data[0];
    const rows = data.slice(1);
    
    return rows.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        let val = row[index];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        obj[header] = val;
      });
      return obj;
    });
  } catch (error) {
    Logger.log("Error in getData: " + error.toString());
    return [];
  }
}

// دالة حفظ أو تحديث السجلات
function saveRecord(sheetName, record) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    const data = sheet.getDataRange().getValues();
    let headers = data.length > 0 ? data[0] : [];
    
    // إنشاء الهيدر إذا كان الشيت جديداً
    if (headers.length === 0) {
      headers = Object.keys(record);
      sheet.appendRow(headers);
    }
    
    // البحث عن السجل للتعديل أو الإضافة
    let rowIndex = -1;
    if (record.ID) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(record.ID)) {
          rowIndex = i + 1;
          break;
        }
      }
    } else {
      record.ID = 'ID-' + new Date().getTime();
    }
    
    const rowData = headers.map(header => record[header] !== undefined ? record[header] : '');
    
    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    
    return { success: true, message: 'تم حفظ البيانات بنجاح', id: record.ID };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// دالة حذف سجل
function deleteRecord(sheetName, id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'الشيت غير موجود' };
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'تم الحذف بنجاح' };
      }
    }
    return { success: false, message: 'لم يتم العثور على السجل' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// دالة رفع الملفات إلى Google Drive
function uploadFileToDrive(fileData) {
  try {
    const folder = FOLDER_ID ? DriveApp.getFolderById(FOLDER_ID) : DriveApp.getRootFolder();
    
    // استخراج بيانات Base64
    const splitData = fileData.data.split(',');
    const contentType = splitData[0].split(';')[0].replace('data:', '');
    const bytes = Utilities.base64Decode(splitData[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileData.fileName);
    
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return {
      success: true,
      url: file.getUrl(),
      fileId: file.getId()
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// دالة تصفية المصروفات بالمدى الزمني
function getExpensesByDateRange(startDate, endDate) {
  const allExpenses = getData('المصروفات');
  if (!startDate && !endDate) return allExpenses;

  const start = startDate ? new Date(startDate) : new Date('1970-01-01');
  const end = endDate ? new Date(endDate) : new Date('2099-12-31');
  end.setHours(23, 59, 59, 999);

  return allExpenses.filter(item => {
    if (!item['التاريخ']) return false;
    const itemDate = new Date(item['التاريخ']);
    return itemDate >= start && itemDate <= end;
  });
}

// المعالج الرئيسي للطلبات المباشرة عبر POST
function doPost(e) {
  try {
    var contents = JSON.parse(e.postData.contents);
    var action = contents.action;
    var res = {};

    if (action === 'getData') {
      res = getData(contents.sheetName);
    } else if (action === 'saveRecord') {
      res = saveRecord(contents.sheetName, contents.record);
    } else if (action === 'deleteRecord') {
      res = deleteRecord(contents.sheetName, contents.id);
    } else if (action === 'uploadFile') {
      res = uploadFileToDrive(contents.fileData);
    } else if (action === 'getExpensesByDateRange') {
      res = getExpensesByDateRange(contents.startDate, contents.endDate);
    } else {
      res = { result: "error", message: "Action not recognized" };
    }

    return ContentService.createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
