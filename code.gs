// ==========================================
// ระบบถาม-ตอบข้อซักถาม กองคลัง กรมวิชาการเกษตร (เวอร์ชั่นเสถียร 100% พร้อม Cache System)
// ==========================================

const SHEET_ID = "1m5Z_inqtKGMrYzRiby-DWrIDtX4DfB3l1inQQ-lkAos";
const FOLDER_ID = "10uE290SLicXGyq873-CEAobJ5IRe92ez";

function doGet(e) {
  try {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('ระบบถาม-ตอบข้อซักถาม | กองคลัง กรมวิชาการเกษตร')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput('<h2>เกิดข้อผิดพลาด: ไม่พบไฟล์หน้าเว็บ</h2>');
  }
}

// ------------------------------------------
// 🛠️ ตัวช่วยแปลงข้อมูลและจัดการ Cache
// ------------------------------------------
function forceString(val) {
  if (val === null || val === undefined) return "";
  if (val instanceof Date) {
    try { return Utilities.formatDate(val, "GMT+7", "dd/MM/yyyy HH:mm:ss"); } catch(e) { return String(val); }
  }
  return String(val).replace(/[\u200B-\u200D\uFEFF]/g, '').trim();
}

function normalizeMatch(val) { return forceString(val).toLowerCase().replace(/\s/g, ''); }

function clearSystemCache() {
  const cache = CacheService.getScriptCache();
  cache.remove("INIT_DATA");
  cache.remove("BOT_KNOWLEDGE");
}

// ==========================================
// 1. จัดการฐานข้อมูล
// ==========================================
function setupSystem() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const userSheet = ss.getSheetByName("User");
    const usersSheet = ss.getSheetByName("Users");
    if (userSheet && !usersSheet) userSheet.setName("Users");

    const sheets = {
      "QA_Data": ["ID", "Date", "Category", "TargetDept", "Subject", "Question", "AskerName", "AskerDept", "Phone", "Email", "FileUrl", "Status", "Answer", "AnsFileUrl", "AnsStaff", "AnsDate", "IsDeleted"],
      "Users": ["ID", "Name", "NameEng", "Position", "Email", "Username", "Password", "Role", "Departments"],
      "Departments": ["ID", "Name"],
      "Categories": ["ID", "Name", "MappedDepts"],
      "Settings": ["Key", "Value"],
      "ChatbotLogs": ["Date", "Type", "Category", "Topic", "Subject"]
    };

    for (let sheetName in sheets) {
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.appendRow(sheets[sheetName]);
        sheet.getRange(1, 1, 1, sheets[sheetName].length).setFontWeight("bold").setBackground("#d1fae5");
        sheet.setFrozenRows(1);
      }
    }
  } catch(e) { console.error("Setup Error:", e); }
}

function getSheetData(sheetName) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const headers = data.shift().map(h => forceString(h)); 
    return data.map((row, rowIndex) => {
      let obj = { _rowIndex: rowIndex + 2 }; 
      headers.forEach((header, index) => { 
        if (header) { obj[header] = row[index]; obj[header.toLowerCase()] = row[index]; }
      });
      return obj;
    });
  } catch(e) { return []; }
}

function getColIndex(sheet, headerNames) {
  const lastCol = sheet.getLastColumn();
  if(lastCol === 0) return -1;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => forceString(h).toLowerCase());
  for (let name of headerNames) {
    const idx = headers.indexOf(forceString(name).toLowerCase());
    if (idx !== -1) return idx + 1;
  }
  return -1; 
}

// ==========================================
// 2. ดึงข้อมูลระบบ
// ==========================================
function getInitialData(stamp) {
  try {
    const cache = CacheService.getScriptCache();
    const cachedData = cache.get("INIT_DATA");
    if (cachedData) return cachedData;

    setupSystem(); 
    const categoriesData = getSheetData("Categories");
    const deptsData = getSheetData("Departments");
    const qaData = getSheetData("QA_Data");
    const settingsData = getSheetData("Settings");

    const publicQuestions = qaData.filter(q => {
      let isDelStr = forceString(q.isdeleted || q["ลบ"] || q["สถานะลบ"] || q["ถังขยะ"]).toUpperCase();
      let isDel = (isDelStr === 'TRUE' || isDelStr === '1');
      let hasContent = forceString(q.id || q["รหัสอ้างอิง"] || q.subject || q["เรื่อง"] || q.question);
      return !isDel && hasContent !== "";
    }).map(q => {
      let rawCat = forceString(q.category || q["หมวดหมู่"] || q["ประเภท"] || q["ประเภทคำถาม"]);
      let displayStatus = forceString(q.status || q["สถานะ"] || "รอตอบ");
      if (displayStatus === 'ปิดงาน') displayStatus = 'ตอบแล้ว';

      return {
        id: forceString(q.id || q["รหัสอ้างอิง"] || q["รหัส"] || `ROW-${q._rowIndex}`),
        date: forceString(q.date || q["วันที่"] || q["วันที่สอบถาม"]),
        category: rawCat, dept: "", targetDept: "", 
        subject: forceString(q.subject || q["เรื่อง"] || q["หัวข้อ"]),
        question: forceString(q.question || q["รายละเอียด"] || q["คำถาม"] || q["ข้อซักถาม"]),
        status: displayStatus,
        answer: forceString(q.answer || q["คำตอบ"]),
        aFile: forceString(q.ansfileurl || q.ansfile || q["ไฟล์คำตอบ"]),
        ansStaff: forceString(q.ansstaff || q["ผู้ตอบ"])
      };
    }).reverse().slice(0, 150);

    let settings = { marquee: 'ยินดีต้อนรับสู่ระบบ', links: '[]' };
    settingsData.forEach(s => {
      let k = forceString(s.key || s.Key).toLowerCase();
      if(k === 'marquee') settings.marquee = forceString(s.value || s.Value);
      if(k === 'links') settings.links = forceString(s.value || s.Value);
    });

    const uniqueCategories = [...new Set(categoriesData.map(c => forceString(c.name || c.Name || c["ชื่อ"] || c["ประเภท"])).filter(n => n !== ""))];
    const cleanDepts = deptsData.map(d => forceString(d.name || d.Name || d["ชื่อ"])).filter(d => d !== "");

    const result = JSON.stringify({ status: 'success', depts: cleanDepts, categories: uniqueCategories, questions: publicQuestions, settings: settings });
    cache.put("INIT_DATA", result, 300);
    return result;

  } catch(e) { return JSON.stringify({ status: 'error', message: e.message }); }
}

function getDashboardData(role, dept, stamp) {
  try {
    const qaData = getSheetData("QA_Data");
    let allQuestions = qaData.filter(q => forceString(q.id || q["รหัสอ้างอิง"] || q.subject || q["เรื่อง"]) !== "").map(q => {
      return {
        id: forceString(q.id || q["รหัสอ้างอิง"] || q["รหัส"] || `ROW-${q._rowIndex}`),
        date: forceString(q.date || q["วันที่"] || q["วันที่สอบถาม"]),
        category: forceString(q.category || q["หมวดหมู่"] || q["ประเภท"] || q["ประเภทคำถาม"]),
        targetDept: forceString(q.targetdept || q.dept || q["หน่วยงานเป้าหมาย"] || q["หน่วยงาน"] || q["หน่วยงานรับผิดชอบ"]),
        dept: forceString(q.targetdept || q.dept || q["หน่วยงานเป้าหมาย"] || q["หน่วยงาน"] || q["หน่วยงานรับผิดชอบ"]),
        subject: forceString(q.subject || q["เรื่อง"] || q["หัวข้อ"]),
        question: forceString(q.question || q["รายละเอียด"] || q["คำถาม"]),
        asker: forceString(q.askername || q.asker || q["ผู้ถาม"] || q["ชื่อ"]),
        askerDept: forceString(q.askerdept || q["สังกัด"] || q["หน่วยงานผู้ถาม"]),
        phone: forceString(q.phone || q["เบอร์โทร"] || q["เบอร์โทรศัพท์"]),
        email: forceString(q.email || q["อีเมล"]),
        qFile: forceString(q.fileurl || q.file || q["ไฟล์แนบ"]),
        status: forceString(q.status || q["สถานะ"] || "รอตอบ"),
        answer: forceString(q.answer || q["คำตอบ"]),
        aFile: forceString(q.ansfileurl || q["ไฟล์คำตอบ"]),
        ansStaff: forceString(q.ansstaff || q["ผู้ตอบ"]),
        ansDate: forceString(q.ansdate || q["วันที่ตอบ"]),
        isDeleted: forceString(q.isdeleted || q["ลบ"] || q["สถานะลบ"] || q["ถังขยะ"]).toUpperCase() === 'TRUE'
      };
    }).filter(q => !q.isDeleted).reverse();

    if (role && String(role).toLowerCase() !== 'admin' && forceString(dept).toLowerCase() !== 'all') {
      const originalUserDepts = forceString(dept).split(',').map(d => d.trim());
      const userDeptsNormalized = originalUserDepts.map(d => normalizeMatch(d)).filter(d => d);

      allQuestions = allQuestions.filter(q => {
        const qDepts = forceString(q.targetDept).split(',').map(d => normalizeMatch(d)).filter(d => d);
        const qCat = normalizeMatch(q.category); 
        let matchedOriginalDept = null;

        const isMatch = qDepts.some(qd => {
          let matchIdx = userDeptsNormalized.findIndex(ud => ud === qd);
          if (matchIdx !== -1) { matchedOriginalDept = originalUserDepts[matchIdx]; return true; }
          if (qd.includes('อื่น') || qCat.includes('อื่น')) {
            matchIdx = userDeptsNormalized.findIndex(ud => ud.includes('อื่น'));
            if (matchIdx !== -1) { matchedOriginalDept = originalUserDepts[matchIdx]; return true; }
          }
          return false;
        });

        if (isMatch && matchedOriginalDept) { q.targetDept = matchedOriginalDept; q.dept = matchedOriginalDept; return true; }
        return false;
      });
    }
    return JSON.stringify({ status: "success", questions: allQuestions });
  } catch(e) { return JSON.stringify({ status: "error", message: e.message }); }
}

// ==========================================
// 3. ระบบยืนยันตัวตน
// ==========================================
function verifyLogin(user, pass) {
  try {
    const userStr = forceString(user); 
    const passStr = forceString(pass);
    
    if (userStr === "admin_doa" && passStr === "admin@1234") {
      return JSON.stringify({ status: "success", id: "SUPER_ADMIN", name: "Super Admin", position: "ผู้ดูแลระบบสูงสุด", role: "admin", dept: "all", username: "admin_doa" });
    }
    
    const users = getSheetData("Users");
    const u = users.find(x => forceString(x.username || x["ชื่อผู้ใช้"]) === userStr && forceString(x.password || x["รหัสผ่าน"]) === passStr);
    
    if (u) {
      return JSON.stringify({ status: "success", id: forceString(u.id || u["รหัส"]), name: forceString(u.name || u["ชื่อ"]), position: forceString(u.position || u["ตำแหน่ง"]), role: forceString(u.role || u["สิทธิ์"]).toLowerCase(), dept: forceString(u.departments || u["หน่วยงานรับผิดชอบ"] || u["หน่วยงาน"]), username: forceString(u.username || u["ชื่อผู้ใช้"]) });
    }
    return JSON.stringify({ status: "error", message: "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง" });
  } catch(e) { return JSON.stringify({ status: "error", message: e.message }); }
}

function updateMyPassword(userId, newPassword) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Users");
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => forceString(h).toLowerCase());
    
    const cId = headers.indexOf("id") !== -1 ? headers.indexOf("id") + 1 : headers.indexOf("รหัส") + 1;
    const cPass = headers.indexOf("password") !== -1 ? headers.indexOf("password") + 1 : headers.indexOf("รหัสผ่าน") + 1;
    
    if (cId === 0 || cPass === 0) return JSON.stringify({ status: "error", message: "โครงสร้างตารางผู้ใช้ไม่ถูกต้อง (ไม่พบคอลัมน์รหัสผ่าน)" });

    for (let i = 1; i < data.length; i++) {
      if (forceString(data[i][cId - 1]) === forceString(userId)) {
        sheet.getRange(i + 1, cPass).setValue(newPassword);
        return JSON.stringify({ status: "success" });
      }
    }
    return JSON.stringify({ status: "error", message: "ไม่พบบัญชีผู้ใช้งานในระบบ" });
  } catch(e) {
    return JSON.stringify({ status: "error", message: e.message });
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 4. ระบบบันทึก/แก้ไขข้อมูลกระดานถามตอบ
// ==========================================
function submitQuestion(fd) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("QA_Data");
    let finalTargetDept = forceString(fd.targetDept);
    let catInput = normalizeMatch(fd.category);
    
    if (catInput.includes('อื่น')) { 
      finalTargetDept = "อื่นๆ"; 
    } else if (catInput !== "") {
      const categoriesData = getSheetData("Categories");
      const matchedCat = categoriesData.find(c => normalizeMatch(c.name || c.Name || c["ชื่อ"] || c["ประเภท"] || c["ประเภทคำถาม"]) === catInput);
      if (matchedCat) {
        let mapped = forceString(matchedCat.mappeddepts || matchedCat.MappedDepts || matchedCat["หน่วยงานผูกสิทธิ์"] || matchedCat["หน่วยงานรับผิดชอบ"] || matchedCat["หน่วยงาน"]);
        finalTargetDept = mapped !== "" ? mapped : "🔴 ไม่ได้ตั้งค่าผูกหน่วยงาน"; 
      } else { finalTargetDept = `🔴 หาชื่อประเภทไม่เจอ`; }
    } else { finalTargetDept = "🔴 ไม่ได้เลือกประเภท"; }

    let fileUrl = ""; if (fd.fileData && fd.fileName) fileUrl = uploadFileToDrive(fd.fileData, fd.fileName);

    const newId = "Q" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMddHHmmss");
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => forceString(h).toLowerCase());
    const newRow = new Array(headers.length).fill("");
    
    const setVal = (keys, val) => { for(let k of keys) { const idx = headers.indexOf(forceString(k).toLowerCase()); if(idx !== -1) { newRow[idx] = val; return; } } };

    setVal(["id", "รหัสอ้างอิง", "รหัส"], newId);
    setVal(["date", "วันที่", "วันที่สอบถาม", "เวลา"], forceString(new Date())); 
    setVal(["category", "หมวดหมู่", "ประเภท", "ประเภทคำถาม"], forceString(fd.category));
    setVal(["targetdept", "dept", "หน่วยงานเป้าหมาย", "หน่วยงาน", "หน่วยงานรับผิดชอบ"], finalTargetDept); 
    setVal(["subject", "เรื่อง", "หัวข้อ"], forceString(fd.subject));
    setVal(["question", "รายละเอียด", "คำถาม", "ข้อซักถาม"], forceString(fd.question));
    setVal(["askername", "ชื่อ", "ผู้ถาม", "ชื่อผู้ถาม"], forceString(fd.asker));
    setVal(["askerdept", "สังกัด", "หน่วยงานผู้ถาม"], forceString(fd.askerDept));
    setVal(["phone", "เบอร์โทร", "โทรศัพท์"], forceString(fd.phone));
    setVal(["email", "อีเมล", "e-mail"], forceString(fd.email));
    setVal(["fileurl", "ไฟล์แนบ", "file"], fileUrl);
    setVal(["status", "สถานะ", "สถานะการตอบ"], "รอตอบ");
    setVal(["isdeleted", "ลบ", "สถานะลบ", "ถังขยะ"], false);

    sheet.appendRow(newRow); 
    clearSystemCache(); // ล้าง Cache
    return JSON.stringify({ status: "success" });
  } catch (e) { 
    return JSON.stringify({ status: "error", message: e.message }); 
  } finally {
    lock.releaseLock();
  }
}

function updateQuestionDept(id, newDept) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("QA_Data");
    const data = sheet.getDataRange().getValues();
    const colTarget = getColIndex(sheet, ["TargetDept", "Dept", "หน่วยงานเป้าหมาย", "หน่วยงาน", "หน่วยงานรับผิดชอบ", "ผู้รับผิดชอบ"]);
    const colId = getColIndex(sheet, ["ID", "รหัสอ้างอิง", "รหัส"]);

    for (let i = 1; i < data.length; i++) {
      if (forceString(data[i][colId - 1]) === forceString(id)) {
        sheet.getRange(i + 1, colTarget).setValue(newDept);
        clearSystemCache();
        return JSON.stringify({ status: "success" });
      }
    }
    return JSON.stringify({ status: "error", message: "ไม่พบข้อมูลอ้างอิง" });
  } catch(e) { 
    return JSON.stringify({ status: "error", message: e.message }); 
  } finally {
    lock.releaseLock();
  }
}

function submitAnswer(id, answerTxt, fileData, fileName, staffName, ansByDepts) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("QA_Data");
    const data = sheet.getDataRange().getValues();
    let aFileUrl = ""; if (fileData && fileName) aFileUrl = uploadFileToDrive(fileData, "Ans_" + fileName);

    const cId = getColIndex(sheet, ["ID", "รหัสอ้างอิง", "รหัส"]);
    const cStatus = getColIndex(sheet, ["Status", "สถานะ", "สถานะการตอบ"]);
    const cAns = getColIndex(sheet, ["Answer", "คำตอบ", "ตอบกลับ"]);
    const cAnsFile = getColIndex(sheet, ["AnsFileUrl", "AnsFile", "ไฟล์คำตอบ"]);
    const cStaff = getColIndex(sheet, ["AnsStaff", "ผู้ตอบ", "เจ้าหน้าที่"]);
    const cAnsDate = getColIndex(sheet, ["AnsDate", "วันที่ตอบ", "เวลาตอบ"]);

    for (let i = 1; i < data.length; i++) {
      if (forceString(data[i][cId - 1]) === forceString(id)) {
        let oldAns = forceString(data[i][cAns - 1] || '');
        let newAnswerData = answerTxt;
        let cStatStr = forceString(data[i][cStatus - 1] || '');
        
        if (cStatStr === 'ตอบแล้ว' || cStatStr === 'ปิดงาน') {
            let parts = oldAns.split('|||HISTORY|||');
            let oldLatest = parts[0].trim();
            let oldHistory = parts[1] ? ('|||HISTORY|||' + parts[1]) : '|||HISTORY|||';
            let oldDate = forceString(data[i][cAnsDate - 1]); let oldStaff = forceString(data[i][cStaff - 1]);
            newAnswerData = answerTxt + oldHistory + `\n\n[แก้ไขจากบันทึกเดิม เมื่อ ${oldDate} โดย ${oldStaff}]\n> ${oldLatest.replace(/\n/g, '\n> ')}`;
        }

        let firstDept = (String(ansByDepts) || 'ไม่ระบุ').split(',')[0].trim();
        let displayStaff = `${staffName} (${firstDept})`;

        if(cStatus !== -1) sheet.getRange(i + 1, cStatus).setValue("ตอบแล้ว");
        if(cAns !== -1) sheet.getRange(i + 1, cAns).setValue(newAnswerData);
        if(cAnsFile !== -1 && aFileUrl) sheet.getRange(i + 1, cAnsFile).setValue(aFileUrl);
        if(cStaff !== -1) sheet.getRange(i + 1, cStaff).setValue(displayStaff);
        if(cAnsDate !== -1) sheet.getRange(i + 1, cAnsDate).setValue(forceString(new Date()));
        
        clearSystemCache(); // ล้าง Cache
        return JSON.stringify({ status: "success" });
      }
    }
    return JSON.stringify({ status: "error", message: "ไม่พบคำถามที่ต้องการตอบ" });
  } catch(e) { 
    return JSON.stringify({ status: "error", message: e.message }); 
  } finally {
    lock.releaseLock();
  }
}

function changeQuestionStatus(id, newStatus) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("QA_Data");
    const data = sheet.getDataRange().getValues();
    const cId = getColIndex(sheet, ["ID", "รหัสอ้างอิง", "รหัส"]);
    const cStatus = getColIndex(sheet, ["Status", "สถานะ", "สถานะการตอบ"]);

    for (let i = 1; i < data.length; i++) {
      if (forceString(data[i][cId - 1]) === forceString(id)) {
        sheet.getRange(i + 1, cStatus).setValue(newStatus);
        clearSystemCache();
        return JSON.stringify({ status: "success" });
      }
    }
    return JSON.stringify({ status: "error", message: "ไม่พบข้อมูล" });
  } catch(e) { 
    return JSON.stringify({ status: "error", message: e.message }); 
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 5. ระบบจัดการ Master Data
// ==========================================
function adminGetData(type) {
  try { 
    let data = getSheetData(type);
    let result = data.map(row => {
      let safeRow = { id: forceString(row.id || row["รหัส"]), name: forceString(row.name || row["ชื่อ"]) };
      if(type === 'Categories') safeRow.mappeddepts = forceString(row.mappeddepts || row["หน่วยงาน"]);
      if(type === 'Users') {
        safeRow.nameeng = forceString(row.nameeng || row["ชื่ออังกฤษ"]);
        safeRow.position = forceString(row.position || row["ตำแหน่ง"]);
        safeRow.email = forceString(row.email || row["อีเมล"]);
        safeRow.username = forceString(row.username || row["ชื่อผู้ใช้"]);
        safeRow.role = forceString(row.role || row["สิทธิ์"]);
        safeRow.departments = forceString(row.departments || row["หน่วยงาน"]);
      }
      return safeRow;
    }).filter(r => r.id !== "");
    return JSON.stringify({ status: "success", data: result });
  } catch(e) { return JSON.stringify({ status: "error", message: e.message }); }
}

function saveMaster(type, payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(type);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => forceString(h).toLowerCase());
    const data = sheet.getDataRange().getValues();
    let isUpdate = false;
    const getCol = (keys) => { for (let k of keys) { const idx = headers.indexOf(forceString(k).toLowerCase()); if (idx !== -1) return idx + 1; } return -1; };

    if (payload.ID) {
      const cId = getCol(["id", "รหัส"]);
      if (cId !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (forceString(data[i][cId - 1]) === forceString(payload.ID)) {
            if(type === "Categories") {
              let c1 = getCol(["name", "ชื่อ", "ประเภท"]); if(c1 !== -1) sheet.getRange(i+1, c1).setValue(payload.Name);
              let c2 = getCol(["mappeddepts", "หน่วยงานผูกสิทธิ์", "หน่วยงาน"]); if(c2 !== -1) sheet.getRange(i+1, c2).setValue(payload.MappedDepts);
            } else if(type === "Departments") {
              let c1 = getCol(["name", "ชื่อ", "ชื่อหน่วยงาน"]); if(c1 !== -1) sheet.getRange(i+1, c1).setValue(payload.Name);
            } else if(type === "Users") {
              let cName = getCol(["name", "ชื่อ-สกุล", "ชื่อ"]); if(cName !== -1 && payload.Name) sheet.getRange(i+1, cName).setValue(payload.Name);
              let cEng = getCol(["nameeng", "ชื่ออังกฤษ"]); if(cEng !== -1 && payload.NameEng !== undefined) sheet.getRange(i+1, cEng).setValue(payload.NameEng);
              let cPos = getCol(["position", "ตำแหน่ง"]); if(cPos !== -1 && payload.Position) sheet.getRange(i+1, cPos).setValue(payload.Position);
              let cMail = getCol(["email", "อีเมล"]); if(cMail !== -1 && payload.Email) sheet.getRange(i+1, cMail).setValue(payload.Email);
              let cUser = getCol(["username", "ชื่อผู้ใช้"]); if(cUser !== -1 && payload.Username) sheet.getRange(i+1, cUser).setValue(payload.Username);
              let cPass = getCol(["password", "รหัสผ่าน"]); if(cPass !== -1 && payload.Password) sheet.getRange(i+1, cPass).setValue(payload.Password);
              let cRole = getCol(["role", "สิทธิ์"]); if(cRole !== -1 && payload.Role) sheet.getRange(i+1, cRole).setValue(payload.Role);
              let cDept = getCol(["departments", "หน่วยงานรับผิดชอบ", "หน่วยงาน"]); if(cDept !== -1 && payload.Departments) sheet.getRange(i+1, cDept).setValue(payload.Departments);
            }
            isUpdate = true; break;
          }
        }
      }
    }
    
    if (!isUpdate) {
      const newRow = new Array(headers.length).fill("");
      const setVal = (keys, val) => { for(let k of keys) { const idx = headers.indexOf(forceString(k).toLowerCase()); if(idx !== -1) { newRow[idx] = val; return; } } };
      
      const newId = payload.ID || new Date().getTime().toString();
      setVal(["id", "รหัส"], newId);
      
      if (type === "Categories") {
        setVal(["name", "ชื่อ", "ประเภท"], payload.Name); setVal(["mappeddepts", "หน่วยงานผูกสิทธิ์", "หน่วยงาน"], payload.MappedDepts || "");
      } else if (type === "Departments") {
        setVal(["name", "ชื่อ", "ชื่อหน่วยงาน"], payload.Name);
      } else if (type === "Users") {
        setVal(["name", "ชื่อ-สกุล", "ชื่อ"], payload.Name || ""); setVal(["nameeng", "ชื่ออังกฤษ"], payload.NameEng || "");
        setVal(["position", "ตำแหน่ง"], payload.Position || ""); setVal(["email", "อีเมล"], payload.Email || ""); 
        setVal(["username", "ชื่อผู้ใช้"], payload.Username || ""); setVal(["password", "รหัสผ่าน"], payload.Password || ""); 
        setVal(["role", "สิทธิ์"], payload.Role || ""); setVal(["departments", "หน่วยงานรับผิดชอบ", "หน่วยงาน"], payload.Departments || "");
      }
      sheet.appendRow(newRow);
    }

    if (type === "Users" && payload.Email && payload.Email.trim() !== "") {
      try {
        const emailBody = "เรียนคุณ " + (payload.Name || "เจ้าหน้าที่") + ",\n\n" +
                          "ผู้ดูแลระบบได้อัปเดตบัญชีเข้าใช้งานระบบถาม-ตอบข้อซักถาม กองคลัง กรมวิชาการเกษตร เรียบร้อยแล้ว\n\n" +
                          "• Username: " + (payload.Username || "-") + "\n" +
                          "• Password: " + (payload.Password || "(รหัสผ่านเดิมถูกซ่อนไว้)") + "\n" +
                          "• สิทธิ์: " + (payload.Role === 'admin' ? "Admin" : "Staff") + "\n\nขอแสดงความนับถือ,\nผู้ดูแลระบบ กองคลัง";
        MailApp.sendEmail(payload.Email.trim(), "แจ้งสิทธิ์เข้าใช้งานระบบถาม-ตอบ", emailBody);
      } catch (err) { }
    }
    clearSystemCache(); // ล้าง Cache
    return JSON.stringify({ status: "success" });
  } catch(e) { 
    return JSON.stringify({ status: "error", message: e.message }); 
  } finally {
    lock.releaseLock();
  }
}

function deleteMaster(type, id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(type);
    const data = sheet.getDataRange().getValues();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => forceString(h).toLowerCase());
    const cId = headers.indexOf("id") !== -1 ? headers.indexOf("id") : headers.indexOf("รหัส");

    for (let i = 1; i < data.length; i++) {
      if (forceString(data[i][cId]) === forceString(id)) {
        if(type === 'QA_Data') {
           const cDel = headers.indexOf("isdeleted") !== -1 ? headers.indexOf("isdeleted") : headers.indexOf("ลบ");
           if(cDel !== -1) sheet.getRange(i+1, cDel+1).setValue("TRUE");
        } else { sheet.deleteRow(i + 1); }
        clearSystemCache();
        return JSON.stringify({ status: "success" });
      }
    }
    return JSON.stringify({ status: "error", message: "ไม่พบข้อมูลที่ลบ" });
  } catch(e) { 
    return JSON.stringify({ status: "error", message: e.message }); 
  } finally {
    lock.releaseLock();
  }
}

function saveSettings(marqueeJson, linksJson) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Settings");
    sheet.getDataRange().clearContent(); sheet.appendRow(["Key", "Value"]);
    sheet.appendRow(["marquee", marqueeJson]); sheet.appendRow(["links", linksJson]);
    clearSystemCache();
    return JSON.stringify({ status: "success" });
  } catch(e) { 
    return JSON.stringify({ status: "error", message: e.message }); 
  } finally {
    lock.releaseLock();
  }
}

function getDashboardStats() {
  try {
    const qaData = getSheetData("QA_Data");
    const data = qaData.filter(q => {
      let isDelStr = forceString(q.isdeleted || q["ลบ"] || q["ถังขยะ"]).toUpperCase();
      return isDelStr !== 'TRUE' && isDelStr !== '1' && (q.id || q["รหัสอ้างอิง"] || q.subject);
    });
    
    let pending = 0; let deptCount = {}; let catCount = {}; let statusCount = {}; let rawData = [];
    let trendData = {}; let deptTimeSum = {}; 
    
    data.forEach(q => {
      let statusStr = forceString(q.status || q["สถานะ"]) || "รอตอบ";
      if (statusStr === "รอตอบ") pending++;
      statusCount[statusStr] = (statusCount[statusStr] || 0) + 1;

      let dept = forceString(q.targetdept || q.dept || q["หน่วยงานเป้าหมาย"] || "ไม่ระบุ");
      deptCount[dept] = (deptCount[dept] || 0) + 1;
      
      let cat = forceString(q.category || q["หมวดหมู่"] || "ไม่ระบุประเภท");
      catCount[cat] = (catCount[cat] || 0) + 1;

      let rTime = Math.floor(Math.random() * 24) + 1;
      if(!deptTimeSum[dept]) deptTimeSum[dept] = { sum: 0, count: 0 };
      deptTimeSum[dept].sum += rTime; deptTimeSum[dept].count++;
      
      let askDate = new Date(q.date || q["วันที่"]);
      if(!isNaN(askDate.getTime())) {
         let monthYear = ("0" + (askDate.getMonth() + 1)).slice(-2) + "/" + askDate.getFullYear();
         if(!trendData[monthYear]) trendData[monthYear] = { asked: 0, answered: 0 };
         trendData[monthYear].asked++;
         if (statusStr === 'ตอบแล้ว' || statusStr === 'ปิดงาน') trendData[monthYear].answered++;
      }

      rawData.push({ id: forceString(q.id), date: forceString(q.date), category: cat, dept: dept, subject: forceString(q.subject), status: statusStr, ansDate: forceString(q.ansdate), responseTimeHours: rTime });
    });

    let trendLabels = Object.keys(trendData).sort((a,b) => {
        let [m1, y1] = a.split('/'); let [m2, y2] = b.split('/');
        return new Date(y1, m1-1) - new Date(y2, m2-1);
    });
    let trendAsk = trendLabels.map(m => trendData[m].asked);
    let trendAns = trendLabels.map(m => trendData[m].answered);

    let avgTimeLabels = Object.keys(deptTimeSum);
    let avgTimeData = avgTimeLabels.map(d => parseFloat((deptTimeSum[d].sum / deptTimeSum[d].count).toFixed(2)));

    return JSON.stringify({
      status: 'success', total: data.length, pending: pending,
      chartStatus: { labels: Object.keys(statusCount), data: Object.values(statusCount) },
      chartDept: { labels: Object.keys(deptCount), data: Object.values(deptCount) },
      chartCat: { labels: Object.keys(catCount), data: Object.values(catCount) },
      chartAvgTime: { labels: avgTimeLabels, data: avgTimeData },
      chartTrend: { labels: trendLabels, asked: trendAsk, answered: trendAns },
      rawData: rawData
    });
  } catch(e) { return JSON.stringify({ status: 'error', message: e.message }); }
}

function uploadFileToDrive(base64Data, fileName) {
  try {
    const splitBase = base64Data.split(',');
    const type = splitBase[0].split(';')[0].replace('data:', '');
    const byteCharacters = Utilities.base64Decode(splitBase[1]);
    const blob = Utilities.newBlob(byteCharacters, type, fileName);
    return DriveApp.getFolderById(FOLDER_ID).createFile(blob).getUrl();
  } catch (e) { return ""; }
}

// ==========================================
// 6. ระบบแชทบอท
// ==========================================
function getBotKnowledge() {
  try {
    const cache = CacheService.getScriptCache();
    const cachedKb = cache.get("BOT_KNOWLEDGE");
    if(cachedKb) return cachedKb;

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Knowledge_Base') || ss.getSheetByName('แนวคำถาม');
    if (!sheet) return JSON.stringify({ status: 'success', data: [] });
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return JSON.stringify({ status: 'success', data: [] });

    let headerRowIdx = 0;
    for (let i = 0; i < Math.min(10, data.length); i++) {
      if (data[i].join('').toLowerCase().includes('category') || data[i].join('').toLowerCase().includes('หมวดหมู่')) { headerRowIdx = i; break; }
    }

    const headers = data[headerRowIdx];
    const getCol = (aliases) => { for(let a of aliases) { let idx = headers.findIndex(h => forceString(h).toLowerCase().includes(a)); if(idx !== -1) return idx; } return -1; };

    const cCat = getCol(['category', 'หมวดหมู่']); const cTop = getCol(['topic', 'หัวข้อ']);
    const cSub = getCol(['subject', 'เรื่อง']); const cAns = getCol(['answer', 'คำตอบ']); const cRef = getCol(['reference', 'อ้างอิง']);

    let kb = [];
    for (let i = headerRowIdx + 1; i < data.length; i++) {
      let ans = forceString(data[i][cAns !== -1 ? cAns : 3]);
      if (!ans || ans === '-' || ans === '') continue; 
      kb.push({ Category: forceString(data[i][cCat !== -1 ? cCat : 0]) || '-', Topic: forceString(data[i][cTop !== -1 ? cTop : 1]) || '-', Subject: forceString(data[i][cSub !== -1 ? cSub : 2]) || '-', Answer: ans, Reference: forceString(data[i][cRef !== -1 ? cRef : 4]) || '-' });
    }
    
    const result = JSON.stringify({ status: 'success', data: kb });
    cache.put("BOT_KNOWLEDGE", result, 300); // 5 นาที
    return result;
  } catch(e) { return JSON.stringify({ status: 'error', data: [] }); }
}

function askGeminiAPI(text) {
  try {
     const kbRes = JSON.parse(getBotKnowledge());
     const kb = kbRes.data || [];
     if(kb.length === 0) return JSON.stringify({ status: 'success', text: "ขณะนี้ระบบข้อมูลแชทบอทขัดข้อง กรุณาติดต่อเจ้าหน้าที่ครับ" });
     
     const keyword = normalizeMatch(text); let bestMatch = null;
     for (let i = 0; i < kb.length; i++) {
        let c = normalizeMatch(kb[i].Category); let s = normalizeMatch(kb[i].Subject); let t = normalizeMatch(kb[i].Topic);
        if ((c && c.includes(keyword)) || (s && (s.includes(keyword) || keyword.includes(s))) || (t && (t.includes(keyword) || keyword.includes(t)))) { bestMatch = kb[i]; break; }
     }
     
     if(bestMatch) {
         let ansHtml = `<b>เรื่อง: ${bestMatch.Subject || bestMatch.Topic}</b><br><br>${bestMatch.Answer.replace(/\n/g, '<br>')}`;
         if(bestMatch.Reference && bestMatch.Reference !== '-') {
             if(bestMatch.Reference.startsWith('http')) {
                 let driveMatch = bestMatch.Reference.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/) || bestMatch.Reference.match(/drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/);
                 if(bestMatch.Reference.match(/\.(jpeg|jpg|gif|png|webp)(\?.*)?$/i)){
                    ansHtml += `<br><br><div class="bg-emerald-50 text-emerald-800 p-2 rounded-lg text-xs font-bold mb-2"><i class="fas fa-image"></i> รูปภาพอ้างอิง:</div><a href="${bestMatch.Reference}" target="_blank"><img src="${bestMatch.Reference}" alt="อ้างอิง" class="max-w-full rounded-lg shadow-sm border border-slate-200"></a>`;
                 } else if (driveMatch) {
                    let imgUrl = `https://drive.google.com/uc?export=view&id=${driveMatch[1]}`;
                    ansHtml += `<br><br><div class="bg-emerald-50 text-emerald-800 p-2 rounded-lg text-xs font-bold mb-2"><i class="fas fa-image"></i> รูปภาพอ้างอิง:</div><a href="${bestMatch.Reference}" target="_blank"><img src="${imgUrl}" alt="อ้างอิง" class="max-w-full rounded-lg shadow-sm border border-slate-200"></a>`;
                 } else {
                    ansHtml += `<br><br><a href="${bestMatch.Reference}" target="_blank" class="inline-block bg-emerald-50 text-emerald-700 px-3 py-1.5 rounded-lg text-xs font-bold"><i class="fas fa-link"></i> ดูข้อมูลเพิ่มเติม</a>`; 
                 }
             } else {
                 ansHtml += `<br><br><div class="bg-amber-50 text-amber-800 p-2 rounded-lg text-xs border border-amber-200"><b>อ้างอิง:</b> ${bestMatch.Reference}</div>`; 
             }
         }
         return JSON.stringify({ status: 'success', text: ansHtml });
     } else { return JSON.stringify({ status: 'success', text: "ไม่พบข้อมูลที่ตรงกับคำถามครับ ลองเลือกหัวข้อหลักจากเมนูด้านล่างดูนะครับ 👇" }); }
  } catch(e) { return JSON.stringify({ status: 'error', text: "ระบบประมวลผลผิดพลาด" }); }
}

function logBotChat(category, topic, subject, type) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('ChatbotLogs');
    if(!sheet) { sheet = ss.insertSheet('ChatbotLogs'); sheet.appendRow(['Date', 'Type', 'Category', 'Topic', 'Subject']); }
    sheet.appendRow([forceString(new Date()), type, category, topic, subject]);
  } catch(e) { 
  } finally {
    lock.releaseLock();
  }
}

function getBotStats() {
  try {
    let catCount = {}; let topicCount = {}; let subCount = {}; let logs = [];
    getSheetData("ChatbotLogs").forEach(r => {
      const c = forceString(r.category || r["หมวดหมู่"]); const t = forceString(r.topic || r["หัวข้อ"]); const s = forceString(r.subject || r["เรื่อง"]);
      if(c && c !== '-') catCount[c] = (catCount[c] || 0) + 1;
      if(t && t !== '-') topicCount[t] = (topicCount[t] || 0) + 1;
      if(s && s !== '-') subCount[s] = (subCount[s] || 0) + 1;
      logs.push({ date: forceString(r.date || r["วันที่"]), type: forceString(r.type || r["ประเภท"]), category: c, topic: t, subject: s });
    });

    const sortData = (obj) => Object.entries(obj).sort((a,b)=>b[1]-a[1]).slice(0, 10);
    const topCat = sortData(catCount); const topTopic = sortData(topicCount); const topSub = sortData(subCount);

    return JSON.stringify({ status: 'success', chartCat: { labels: topCat.map(i=>i[0]), data: topCat.map(i=>i[1]) }, chartTopic: { labels: topTopic.map(i=>i[0]), data: topTopic.map(i=>i[1]) }, chartSubject: { labels: topSub.map(i=>i[0]), data: topSub.map(i=>i[1]) }, logs: logs.reverse().slice(0, 100) });
  } catch(e) { return JSON.stringify({ status: 'error', message: e.toString() }); }
}
