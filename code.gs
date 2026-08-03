function doGet(e) {
  // للتعامل مع طلبات البيانات الخارجية عبر GitHub عبر HTTP GET
  if (e && e.parameter && e.parameter.action) {
    let result = {};
    const action = e.parameter.action;

    try {
      if (action === 'getDashboard') {
        result = getDashboardData();
      } else if (action === 'getData') {
        result = getData(e.parameter.sheetName);
      } else if (action === 'checkLogin') {
        result = checkLogin(e.parameter.email, e.parameter.password);
      } else {
        result = { success: false, message: 'Action not recognized' };
      }
    } catch (err) {
      result = { success: false, message: err.toString() };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // في حال فتح الرابط مباشرة داخل جوجل
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('نظام إدارة مؤسسة حسن خالد حسن فران للمقاولات')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let result = {};
  try {
    const contents = JSON.parse(e.postData.contents);
    const action = contents.action;

    if (action === 'saveRecord') {
      result = saveRecord(contents.sheetName, contents.record);
    } else if (action === 'deleteRecord') {
      result = deleteRecord(contents.sheetName, contents.id);
    } else if (action === 'uploadFile') {
      result = uploadFileToDrive(contents.fileData);
    } else if (action === 'checkLogin') {
      result = checkLogin(contents.email, contents.password);
    } else {
      result = { success: true, data: contents };
    }
  } catch (error) {
    result = { success: false, message: error.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// تهيئة الجداول وتحديث الأعمدة
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = [
    { name: 'المشاريع', headers: ['ID', 'كود المشروع', 'اسم المشروع', 'العميل', 'الموقع', 'القيمة الاجمالية', 'المستخلص المحصل', 'مدير المشروع', 'تاريخ البدء', 'تاريخ النهاية', 'حالة المشروع'] },
    { name: 'المصروفات', headers: ['ID', 'كود المشروع', 'بند المصروف', 'المبلغ', 'التاريخ', 'ملاحظات', 'رابط الفاتورة'] },
    { name: 'الموظفين', headers: ['ID', 'رقم الموظف', 'اسم الموظف', 'رقم الهوية', 'تاريخ الميلاد', 'فصيلة الدم', 'تاريخ انتهاء الإقامة', 'المهنة', 'الراتب الأساسي'] },
    { name: 'مسير الرواتب', headers: ['ID', 'رقم الموظف', 'اسم الموظف', 'الشهر/السنة', 'الراتب الأساسي', 'الخصومات', 'الإضافي', 'صافي الراتب', 'حالة الصرف', 'تاريخ الصرف'] },
    { name: 'العقود والخطابات', headers: ['ID', 'اسم العقد', 'الطرف الثاني', 'التاريخ', 'رابط المستند'] },
    { name: 'وثائق الشركة', headers: ['ID', 'اسم الوثيقة', 'رقم الوثيقة', 'تاريخ الانتهاء', 'رابط المستند'] },
    { name: 'المستخدمين', headers: ['ID', 'الاسم', 'البريد الإلكتروني', 'كلمة المرور', 'الدور'] }
  ];

  sheets.forEach(sheetInfo => {
    let sheet = ss.getSheetByName(sheetInfo.name);
    if (!sheet) {
      sheet = ss.insertSheet(sheetInfo.name);
      sheet.appendRow(sheetInfo.headers);
      sheet.getRange(1, 1, 1, sheetInfo.headers.length).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    }
  });

  return { status: 'success', message: 'تمت تهيئة الجداول بنجاح' };
}

// دالة رفع الملفات إلى Google Drive وإرجاع الرابط
function uploadFileToDrive(fileData) {
  try {
    const folderName = "فواتير ومستندات المؤسسة";
    let folders = DriveApp.getFoldersByName(folderName);
    let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    const contentType = fileData.mimeType;
    const bytes = Utilities.base64Decode(fileData.data.split(',')[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileData.fileName);
    
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return { success: true, url: file.getUrl() };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// تسجيل الدخول
function checkLogin(email, password) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المستخدمين");
  if (!sheet) {
    return { success: true, user: { email: email, role: 'مدير النظام' } }; 
  }
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { success: true, user: { email: email, role: 'مدير النظام' } };
  }
  
  var headers = data[0];
  var emailIdx = headers.indexOf("البريد الإلكتروني");
  var passIdx = headers.indexOf("كلمة المرور");
  var roleIdx = headers.indexOf("الدور");
  var nameIdx = headers.indexOf("الاسم");
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][emailIdx]).trim() === String(email).trim() && 
        String(data[i][passIdx]).trim() === String(password).trim()) {
      return {
        success: true,
        user: {
          name: data[i][nameIdx],
          email: data[i][emailIdx],
          role: data[i][roleIdx] || 'مدير النظام'
        }
      };
    }
  }
  
  return { success: false, message: "البريد الإلكتروني أو كلمة المرور غير صحيحة" };
}

// جلب البيانات من أي شيت
function getData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; 
  
  const headers = data[0].map(h => h ? h.toString().trim() : '');
  const rows = data.slice(1).filter(row => row.some(cell => cell !== "" && cell !== null));
  
  return rows.map(row => {
    let obj = {};
    headers.forEach((h, index) => {
      if (!h) return;
      
      let val = row[index];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      obj[h] = (val !== undefined && val !== null) ? val : "";
    });
    return obj;
  });
}

// حفظ أو تعديل سجل
function saveRecord(sheetName, record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, message: 'الشيت غير موجود' };
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  if (record['ID']) {
    const idIndex = headers.indexOf('ID');
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

  record['ID'] = 'ID-' + new Date().getTime();
  const rowToAppend = headers.map(h => record[h] !== undefined ? record[h] : '');
  sheet.appendRow(rowToAppend);
  return { success: true, message: 'تم الحفظ بنجاح' };
}

// حذف سجل
function deleteRecord(sheetName, id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
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

// بيانات لوحة التحكم
function getDashboardData() {
  const projects = getData('المشاريع');
  const expenses = getData('المصروفات');
  const docs = getData('وثائق الشركة');
  const employees = getData('الموظفين');

  let totalContractValue = 0;
  let totalCollected = 0;
  let totalExpenses = 0;

  projects.forEach(p => {
    totalContractValue += parseFloat(p['القيمة الاجمالية']) || 0;
    totalCollected += parseFloat(p['المستخلص المحصل']) || 0;
  });

  expenses.forEach(e => {
    totalExpenses += parseFloat(e['المبلغ']) || 0;
  });

  const today = new Date();
  const expiringDocs = docs.filter(d => {
    if (!d['تاريخ الانتهاء']) return false;
    const exp = new Date(d['تاريخ الانتهاء']);
    return (exp - today) / (1000 * 60 * 60 * 24) <= 30;
  });

  const expiringIqamas = employees.filter(emp => {
    if (!emp['تاريخ انتهاء الإقامة']) return false;
    const exp = new Date(emp['تاريخ انتهاء الإقامة']);
    return (exp - today) / (1000 * 60 * 60 * 24) <= 30;
  });

  return {
    stats: {
      totalContractValue,
      totalCollected,
      totalRemaining: totalContractValue - totalCollected,
      totalExpenses,
      projectsCount: projects.length
    },
    expiringDocs,
    expiringIqamas,
    projects
  };
}

function savePayrollRecord(record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('مسير الرواتب');
  if (!sheet) {
    setupDatabase();
  }
  return saveRecord('مسير الرواتب', record);
}
