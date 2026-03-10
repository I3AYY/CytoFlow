// =====================================================================
// CytoFlow v1.3.3 AI Edition - Multi-Database Architecture (SaaS Ready)
// =====================================================================

// --- DATABASE CONFIGURATION ---
// นำ ID ของ Google Sheets ทั้ง 4 ไฟล์มาใส่ที่นี่
const DB_FILES = {
  SYSTEM: 'XXXX',      // [Sheets]: Master_Config
  USER: 'XXXX',        // [Sheets]: User_Config
  REF: 'XXXX',         // [Sheets]: Reference_Data_Config
  LOG: 'XXXX'          // [Sheets]: Log_Config
};

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('CytoFlow 2026 (v1.3.3 AI Edition)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function formatDateVal(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(val);
}

function formatDateTimeVal(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  }
  return String(val);
}

// Helper: สำหรับเปรียบเทียบเพื่อทำ Audit Log
function safeString(val) {
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  if (val === null || val === undefined) return "";
  return String(val).trim();
}

function getDiffString(oldArr, newArr, headers) {
  let diffs = [];
  for (let i = 0; i < oldArr.length; i++) {
     let oldVal = safeString(oldArr[i]);
     let newVal = safeString(newArr[i]);
     
     // Remove leading tick for clean logging
     if(oldVal.startsWith("'")) oldVal = oldVal.substring(1);
     if(newVal.startsWith("'")) newVal = newVal.substring(1);

     if (oldVal !== newVal) {
        diffs.push(`[${headers[i]}]: '${oldVal}' -> '${newVal}'`);
     }
  }
  return diffs.length > 0 ? diffs.join(' | ') : 'No data changed';
}

// Helper: ดึง Sheet ข้อมูลคนไข้ตามปี
function getDbSheet(year) {
  const sysSS = SpreadsheetApp.openById(DB_FILES.SYSTEM);
  const configSheet = sysSS.getSheetByName('Year');
  if(!configSheet) throw new Error("ไม่พบแท็บ Year ในไฟล์ Master_Config");
  
  const data = configSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) == String(year)) {
      return SpreadsheetApp.openById(data[i][1]).getSheetByName('Data');
    }
  }
  throw new Error("ไม่พบ Config สำหรับปี: " + year);
}

// --- LOGGING (แยกไฟล์เพื่อความปลอดภัยระดับ PDPA) ---
function logSystem(action, detail, username) {
  try {
    const logSS = SpreadsheetApp.openById(DB_FILES.LOG);
    let sheet = logSS.getSheetByName('System_Logs');
    if (!sheet) {
      sheet = logSS.insertSheet('System_Logs');
      sheet.appendRow(['Timestamp', 'User', 'Action', 'Detail']);
    }
    sheet.appendRow([new Date(), username, action, detail]);
  } catch(e) { console.log("Log Sys Error: " + e); }
}

function logData(year, action, detail, username, cytoNo) {
  try {
    const logSS = SpreadsheetApp.openById(DB_FILES.LOG);
    const sheetName = 'Log_' + year;
    let sheet = logSS.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = logSS.insertSheet(sheetName);
      sheet.appendRow(['Timestamp', 'User', 'Action', 'CytoNo', 'Detail']);
      sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#e0e7ff");
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date(), username, action, cytoNo, detail]);
  } catch(e) { 
    console.log("Log Data Error: " + e); 
  }
}

function getCurrentYear() {
  const today = new Date();
  return today.getFullYear() + 543; // ปี พ.ศ. ปัจจุบัน
}

// ให้ Frontend เรียกใช้บันทึก Log ได้ (เช่น Logout, Cancel)
function apiFrontendLog(action, detail, username) {
  logSystem(action, detail, username);
  return { status: 'success' };
}

// --- API: GET MASTER DATA (ดึงจากหลายไฟล์) ---
function apiGetMasterData() {
  try {
    const refSS = SpreadsheetApp.openById(DB_FILES.REF);
    const userSS = SpreadsheetApp.openById(DB_FILES.USER);
    
    const unitsSheet = refSS.getSheetByName('Sampling_Unit');
    const districtSheet = refSS.getSheetByName('District');
    const adequacySheet = refSS.getSheetByName('SPECIMEN ADEQUACY');
    const sqSheet = refSS.getSheetByName('200 Squamous Cell');
    const glSheet = refSS.getSheetByName('200 Glandular Cell');
    const cat300Sheet = refSS.getSheetByName('300 OTHER');
    
    const cytoTechSheet = userSS.getSheetByName('Cytotechnologist');
    const pathoSheet = userSS.getSheetByName('Pathologist');
    
    let units = []; let districts = []; let adequacyMaster = [];
    let cytoTechs = []; let pathos = [];
    let masterSquamous = []; let masterGlandular = []; let masterCat300 = [];

    if (unitsSheet) {
      const uData = unitsSheet.getRange(2, 1, Math.max(1, unitsSheet.getLastRow() - 1)).getValues();
      units = uData.map(r => String(r[0]).trim()).filter(Boolean);
    }
    if (districtSheet) {
      const dData = districtSheet.getRange(2, 1, Math.max(1, districtSheet.getLastRow() - 1)).getValues();
      districts = dData.map(r => String(r[0]).trim()).filter(Boolean);
    }
    if (adequacySheet) {
      const adData = adequacySheet.getRange(2, 1, Math.max(1, adequacySheet.getLastRow() - 1), 2).getValues();
      adequacyMaster = adData.map(r => ({ group: String(r[0]).trim(), text: String(r[1]).trim() })).filter(x => x.text);
    }
    if (sqSheet) {
      const sqData = sqSheet.getRange(2, 1, Math.max(1, sqSheet.getLastRow() - 1), 3).getValues();
      masterSquamous = sqData.map(r => ({ main: String(r[0]).trim(), detail1: String(r[1]).trim(), detail2: String(r[2]).trim() })).filter(x => x.main || x.detail1);
    }
    if (glSheet) {
      const glData = glSheet.getRange(2, 1, Math.max(1, glSheet.getLastRow() - 1), 2).getValues();
      masterGlandular = glData.map(r => ({ main: String(r[0]).trim(), detail1: String(r[1]).trim() })).filter(x => x.main || x.detail1);
    }
    if (cat300Sheet) {
      const c300Data = cat300Sheet.getRange(2, 1, Math.max(1, cat300Sheet.getLastRow() - 1)).getValues();
      masterCat300 = c300Data.map(r => String(r[0]).trim()).filter(Boolean);
    }
    
    if (cytoTechSheet) {
      const ctData = cytoTechSheet.getRange(2, 1, Math.max(1, cytoTechSheet.getLastRow() - 1), 2).getValues();
      cytoTechs = ctData.map(r => (String(r[0]).trim() + " " + String(r[1]).trim()).trim()).filter(Boolean);
    }
    if (pathoSheet) {
      const ptData = pathoSheet.getRange(2, 1, Math.max(1, pathoSheet.getLastRow() - 1), 2).getValues();
      pathos = ptData.map(r => (String(r[0]).trim() + " " + String(r[1]).trim()).trim()).filter(Boolean);
    }
    
    return { 
      status: 'success', 
      units: units, 
      districts: districts, 
      adequacyMaster: adequacyMaster, 
      cytoTechs: cytoTechs, 
      pathos: pathos,
      masterSquamous: masterSquamous,
      masterGlandular: masterGlandular,
      masterCat300: masterCat300
    };
  } catch (e) { return { status: 'error', message: 'Master Data Error: ' + e.message }; }
}

function apiGetNextCytoNo(year) {
  try {
    const sheet = getDbSheet(year);
    const yearPrefix = year.toString().substring(2);
    const lastRow = sheet.getLastRow();
    let nextNum = 1;
    if (lastRow > 1) {
      const lastId = sheet.getRange(lastRow, 1).getValue().toString();
      if (lastId.startsWith(yearPrefix)) {
        const numPart = parseInt(lastId.substring(2)); 
        if (!isNaN(numPart)) nextNum = numPart + 1;
      }
    }
    return { status: 'success', cytoNoPreview: yearPrefix + nextNum.toString().padStart(4, '0') };
  } catch (e) { return { status: 'error', message: e.message }; }
}

// --- API: LOGIN ---
function apiLoginStep1(username, password) {
  try {
    const userSS = SpreadsheetApp.openById(DB_FILES.USER);
    const sheet = userSS.getSheetByName('Users');
    if(!sheet) return { status: 'error', message: 'ไม่พบฐานข้อมูล Users' };
    
    const data = sheet.getDataRange().getValues();
    let logoUrl = "https://drive.google.com/thumbnail?id=142CkRafzFxGXtCS5q5D0Iqct7rKr4HSA&sz=w200";
    
    try {
      const sysSS = SpreadsheetApp.openById(DB_FILES.SYSTEM);
      const logoSheet = sysSS.getSheetByName('App_Logo');
      if (logoSheet && logoSheet.getLastRow() > 1) logoUrl = logoSheet.getRange(2, 2).getValue();
    } catch(e) { console.log("Logo fetch error: " + e); }

    let userFound = false;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) == String(username) && String(data[i][1]) == String(password)) {
        userFound = true;
        const email = data[i][5];
        if (!email || email === "") {
          logSystem("Login Failed", "Account missing email", username);
          return { status: 'error', message: 'Account นี้ยังไม่ระบุ Email ในระบบ' };
        }
        
        const otp = Math.floor(100000 + Math.random() * 900000).toString();
        // ตั้งเวลา OTP หมดอายุที่ 300 วินาที (5 นาที)
        CacheService.getScriptCache().put("OTP_" + username, otp, 300);

        // --- NEW OTP EMAIL TEMPLATE (GitHub Style with PDPA) ---
        const htmlTemplate = `
        <div style="font-family: -apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif,'Apple Color Emoji','Segoe UI Emoji'; color: #24292f; max-width: 600px; margin: 0 auto; padding: 20px;">
            <div style="text-align: center; margin-bottom: 24px;">
                <img src="${logoUrl}" alt="CytoFlow Logo" style="width: 64px; height: 64px; border-radius: 50%; object-fit: contain; vertical-align: middle;">
                <span style="font-size: 32px; font-weight: 600; color: #24292f; vertical-align: middle; margin-left: 12px; letter-spacing: -0.5px; display: inline-block;">CytoFlow</span>
            </div>
            <h2 style="font-size: 24px; font-weight: 400; text-align: center; margin-bottom: 24px; color: #24292f;">
                กรุณายืนยันตัวตนของคุณ, <strong>${username}</strong>
            </h2>
            <div style="background-color: #ffffff; border: 1px solid #d0d7de; border-radius: 6px; padding: 24px;">
                <p style="margin-top: 0; margin-bottom: 16px; font-size: 14px;">นี่คือรหัส OTP สำหรับยืนยันการเข้าสู่ระบบของคุณ:</p>
                <div style="text-align: center; font-size: 32px; font-family: ui-monospace,SFMono-Regular,SF Mono,Menlo,Consolas,Liberation Mono,monospace; letter-spacing: 8px; color: #24292f; margin: 24px 0;">
                    ${otp}
                </div>
                <p style="font-size: 14px; margin-bottom: 16px;">
                    รหัสนี้มีอายุการใช้งาน <strong>5 นาที</strong> และสามารถใช้ได้เพียงครั้งเดียว
                </p>
                <p style="font-size: 14px; margin-bottom: 16px;">
                    <strong>โปรดอย่าแชร์รหัสนี้กับใคร:</strong> ทีมงานจะไม่ขอรหัสผ่านหรือ OTP ของคุณทางโทรศัพท์หรืออีเมล<br>
                    เด็ดขาด (มาตรการรักษาความปลอดภัยข้อมูลส่วนบุคคล - PDPA)
                </p>
                <p style="font-size: 14px; margin-bottom: 0;">
                    ขอบคุณ,<br>ทีมงาน CytoFlow
                </p>
            </div>
            <div style="margin-top: 32px; font-size: 12px; color: #6e7781; text-align: center; line-height: 1.5;">
                คุณได้รับอีเมลฉบับนี้เนื่องจากมีการร้องขอรหัสยืนยันสำหรับบัญชีผู้ใช้ CytoFlow ของคุณ หากคุณไม่ได้เป็นผู้ร้องขอ<br>
                โปรดเพิกเฉยต่ออีเมลฉบับนี้ หรือแจ้งผู้ดูแลระบบทันที
            </div>
        </div>
        `;

        try { 
          MailApp.sendEmail({ 
            to: email, 
            subject: "รหัส OTP สำหรับเข้าสู่ระบบ CytoFlow", 
            htmlBody: htmlTemplate,
            name: "CytoFlow" 
          }); 
          logSystem("OTP Requested", "OTP sent to user email", username);
        } 
        catch (mailErr) { 
          logSystem("OTP Error", "Failed to send OTP email: " + mailErr.message, username);
          return { status: 'error', message: 'ส่งอีเมล OTP ไม่สำเร็จ: ' + mailErr.message }; 
        }

        const maskedEmail = email.replace(/^(.)(.*)(.@.*)$/, "$1***$3");
        return { status: 'otp_required', message: 'กรุณากรอกรหัส OTP ที่ส่งไปยัง ' + maskedEmail, systemLogo: logoUrl };
      }
    }
    
    logSystem("Login Failed", "Invalid username or password", username);
    return { status: 'error', message: 'Username หรือ Password ไม่ถูกต้อง' };
  } catch (e) { 
    return { status: 'error', message: 'System Error: ' + e.message }; 
  }
}

function apiVerifyOtp(username, inputOtp) {
  try {
    const cache = CacheService.getScriptCache();
    const storedOtp = cache.get("OTP_" + username);
    if (storedOtp && storedOtp === inputOtp) {
      cache.remove("OTP_" + username);
      
      const userSS = SpreadsheetApp.openById(DB_FILES.USER);
      const sheet = userSS.getSheetByName('Users');
      const data = sheet.getDataRange().getValues();
      let userData = null;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) == String(username)) {
          userData = { name: data[i][2], position: data[i][3], role: data[i][4], image: data[i][6] || "", username: username };
          break;
        }
      }
      
      const sysSS = SpreadsheetApp.openById(DB_FILES.SYSTEM);
      const configSheet = sysSS.getSheetByName('Year');
      const years = configSheet.getDataRange().getValues().slice(1).map(r => String(r[0]));
      
      logSystem("Login Success", "Successfully verified OTP", username);
      return { status: 'success', user: userData, years: years, currentYear: getCurrentYear() };
    } else { 
      logSystem("Login Failed", "Invalid or Expired OTP", username);
      return { status: 'error', message: 'รหัส OTP ไม่ถูกต้อง หรือหมดอายุ' }; 
    }
  } catch (e) { return { status: 'error', message: 'Verify Error: ' + e.message }; }
}

function apiVerifyPassword(username, password) {
  try {
    const userSS = SpreadsheetApp.openById(DB_FILES.USER);
    const sheet = userSS.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) == String(username) && String(data[i][1]) == String(password)) {
        logSystem("Unlock Screen", "Successfully unlocked screen", username);
        return { status: 'success' };
      }
    }
    logSystem("Unlock Failed", "Invalid password during screen unlock", username);
    return { status: 'error', message: 'รหัสผ่านไม่ถูกต้อง' };
  } catch (e) { return { status: 'error', message: 'System Error: ' + e.message }; }
}

// --- API: DASHBOARD ---
function apiGetDashboardData(year) {
  try {
    const sheet = getDbSheet(year);
    const lastRow = sheet.getLastRow();
    let stats = { total: 0, reported: 0, abnormal: 0 };
    let patients = [];

    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 42).getValues();
      data.forEach((r, index) => {
        const status = r[41] ? String(r[41]) : "Pending";
        
        stats.total++;
        if (status === "Reported") stats.reported++;

        patients.push({
          rowId: index + 2, cytoNo: String(r[0]), hn: String(r[1]), cid: String(r[2]),
          prefix: r[3], fname: r[4], lname: r[5], age: r[6], sex: r[7],
          specimenDate: formatDateVal(r[8]), recDate: formatDateVal(r[9]),
          unit: r[10] ? String(r[10]) : "ไม่ระบุ", district: r[11], hcode: r[12], coordinator: r[13], phone: r[14],
          para: r[15], last: r[16], lmp: formatDateVal(r[17]), contraception: r[18], prevTx: r[19],
          clinFind: r[20], clinDx: r[21], lastPap: r[22], method: r[23], registerName: r[24],
          regTimestamp: formatDateTimeVal(r[25]),
          adequacy: r[26], adequacyDetail: r[27], additional: r[28], 
          organism: r[29] ? String(r[29]) : "",
          nonNeo: r[30] ? String(r[30]) : "",
          squamousMain: r[31] ? String(r[31]) : "",
          squamousSub: r[32] ? String(r[32]) : "",
          glandularMain: r[33] ? String(r[33]) : "",
          glandularSub: r[34] ? String(r[34]) : "",
          cat300: r[35], comment: r[36], 
          cytoName: r[37], cytoDateTime: String(r[38]), 
          pathoName: r[39], pathoDateTime: String(r[40]), 
          status: status 
        });
      });
    }
    return { status: 'success', stats: stats, patients: patients.reverse() };
  } catch (e) { return { status: 'error', message: "Data Error: " + e.message }; }
}

// --- API: REGISTER ---
function apiRegisterSample(form, year, username) {
  const lock = LockService.getScriptLock(); lock.tryLock(10000);
  try {
    const sheet = getDbSheet(year);
    const yearPrefix = year.toString().substring(2);
    const lastRow = sheet.getLastRow();
    let nextNum = 1;
    if (lastRow > 1) {
      const lastId = sheet.getRange(lastRow, 1).getValue().toString();
      if (lastId.startsWith(yearPrefix)) { const numPart = parseInt(lastId.substring(2)); if (!isNaN(numPart)) nextNum = numPart + 1; }
    }
    const cytoNo = yearPrefix + nextNum.toString().padStart(4, '0');
    const phoneStr = form.phone ? "'" + form.phone : ""; 

    let record = [
      cytoNo, String(form.hn), String(form.cid), form.prefix, form.fname, form.lname, form.age, form.sex,
      form.specimenDate, form.receivedDate, form.unit, form.district, form.hcode, form.coordinator, phoneStr,
      form.para, form.last, form.lmp, form.contraception, form.prevTx, form.clinFind, form.clinDx,
      form.lastPap, form.method, form.registerName
    ]; 
    record.push(new Date()); 
    for(let i = 0; i < 15; i++) { record.push(""); }
    record.push("Pending");

    sheet.appendRow(record);
    logData(year, "Register", "Created new sample", username, cytoNo);
    return { status: 'success', cytoNo: cytoNo };
  } catch (e) { return { status: 'error', message: "Save Failed: " + e.message }; }
  finally { lock.releaseLock(); }
}

// --- API: UPDATE (WITH FORENSIC DIFF AUDIT LOG) ---
function apiUpdateSample(form, year, rowId, username) {
  const lock = LockService.getScriptLock(); lock.tryLock(10000);
  try {
    const sheet = getDbSheet(year); 
    const rowIndex = parseInt(rowId);
    if (rowIndex > sheet.getLastRow()) return { status: 'error', message: 'Row not found' };
    
    // 1. Get Old Data for Diff
    const oldData = sheet.getRange(rowIndex, 2, 1, 24).getValues()[0];
    
    // 2. Prepare New Data
    const phoneStr = form.phone ? "'" + form.phone : ""; 
    const newRecord = [
      String(form.hn), String(form.cid), form.prefix, form.fname, form.lname, form.age, form.sex,
      form.specimenDate, form.receivedDate, form.unit, form.district, form.hcode, form.coordinator, phoneStr,
      form.para, form.last, form.lmp, form.contraception, form.prevTx, form.clinFind, form.clinDx,
      form.lastPap, form.method, form.registerName
    ];
    
    // 3. Generate Diff String
    const headers = ['HN', 'CID', 'Prefix', 'Fname', 'Lname', 'Age', 'Sex', 'SpecimenDate', 'RecDate', 'Unit', 'District', 'HCode', 'Coordinator', 'Phone', 'Para', 'Last', 'LMP', 'Contraception', 'PrevTx', 'ClinFind', 'ClinDx', 'LastPap', 'Method', 'RegName'];
    const diffLog = getDiffString(oldData, newRecord, headers);
    
    // 4. Update Database
    sheet.getRange(rowIndex, 2, 1, 24).setValues([newRecord]); 
    const cytoNo = sheet.getRange(rowIndex, 1).getValue();
    
    // 5. Save Log
    logData(year, "Edit Registration", diffLog, username, cytoNo);
    
    return { status: 'success', cytoNo: cytoNo };
  } catch (e) { return { status: 'error', message: "Update Failed: " + e.message }; }
  finally { lock.releaseLock(); }
}

// --- API: REPORT (WITH FORENSIC DIFF AUDIT LOG) ---
function apiSubmitReport(form, year, username) {
  try {
    const sheet = getDbSheet(year);
    const row = parseInt(form.rowId);
    
    // 1. Get Old Data for Diff
    const oldData = sheet.getRange(row, 27, 1, 16).getValues()[0]; // Col 27 to 42
    
    // 2. Prepare New Data
    const cytoDT = form.cytoDateTime ? "'" + form.cytoDateTime : "";
    const pathoDT = form.pathoDateTime ? "'" + form.pathoDateTime : "";
    
    const newRecord = [
      form.adequacy, form.adequacyDetail, form.additional, form.organism, form.nonNeo,
      form.squamousMain, form.squamousSub, form.glandularMain, form.glandularSub,
      form.cat300, form.comment, form.cytoName, cytoDT, form.pathoName, pathoDT, "Reported"
    ];
    
    // 3. Update Database
    sheet.getRange(row, 27, 1, 16).setValues([newRecord]); 

    // 4. Generate Diff & Log
    const cytoNo = sheet.getRange(row, 1).getValue();
    const headers = ['Adequacy', 'AdqDetail', 'Additional', 'Organism', 'NonNeo', 'SqMain', 'SqSub', 'GlMain', 'GlSub', 'Cat300', 'Comment', 'CytoName', 'CytoDT', 'PathoName', 'PathoDT', 'Status'];
    
    if (form.isEdit) {
      const diffLog = getDiffString(oldData, newRecord, headers);
      logData(year, "Edit Report", diffLog, username, cytoNo);
    } else {
      logData(year, "Submit Report", "Initial Report Submitted", username, cytoNo);
    }
    
    return { status: 'success' };
  } catch (e) { return { status: 'error', message: e.message }; }
}

// --- API: PROFILE IMAGE & LOGO ---
function apiSaveProfileImage(username, base64Data) {
  try {
    const userSS = SpreadsheetApp.openById(DB_FILES.USER); 
    const sheet = userSS.getSheetByName('Users'); 
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1; let oldFileUrl = "";
    for (let i = 1; i < data.length; i++) { if (String(data[i][0]) === String(username)) { rowIndex = i + 1; oldFileUrl = data[i][6]; break; } }
    if (rowIndex === -1) return { status: 'error', message: 'User not found' };
    if (oldFileUrl && oldFileUrl.includes("drive.google.com")) { try { const idMatch = oldFileUrl.match(/id=([^&]+)/); if (idMatch && idMatch[1]) DriveApp.getFileById(idMatch[1]).setTrashed(true); } catch (e) {} }
    
    const folderName = "CytoFlow_Profiles"; const folders = DriveApp.getFoldersByName(folderName);
    let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    const contentType = base64Data.substring(5, base64Data.indexOf(';')); const bytes = Utilities.base64Decode(base64Data.substr(base64Data.indexOf('base64,')+7));
    const blob = Utilities.newBlob(bytes, contentType, `profile_${username}_${Date.now()}.jpg`); const file = folder.createFile(blob); file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = `https://drive.google.com/thumbnail?id=${file.getId()}&sz=s400`; sheet.getRange(rowIndex, 7).setValue(fileUrl);
    
    logSystem("Change Profile Pic", "Updated profile image", username); return { status: 'success', url: fileUrl };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function apiChangePassword(username, newPassword) {
  try {
    const userSS = SpreadsheetApp.openById(DB_FILES.USER); 
    const sheet = userSS.getSheetByName('Users'); 
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (String(data[i][0]) === String(username)) { sheet.getRange(i + 1, 2).setValue(newPassword); logSystem("Change Password", "Updated password", username); return { status: 'success' }; } }
    return { status: 'error', message: 'User not found' };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function apiSaveSystemLogo(base64Data, username) {
  try {
    const sysSS = SpreadsheetApp.openById(DB_FILES.SYSTEM); 
    let sheet = sysSS.getSheetByName('App_Logo');
    if (!sheet) { sheet = sysSS.insertSheet('App_Logo'); sheet.appendRow(['Name', 'Url']); }
    
    const folderName = "CytoFlow_Logo"; const folders = DriveApp.getFoldersByName(folderName);
    let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    const contentType = base64Data.substring(5, base64Data.indexOf(';')); let ext = "png"; if (contentType.includes("gif")) ext = "gif"; else if (contentType.includes("jpeg")) ext = "jpg";
    const bytes = Utilities.base64Decode(base64Data.substr(base64Data.indexOf('base64,')+7)); const blob = Utilities.newBlob(bytes, contentType, `app_logo_${Date.now()}.${ext}`);
    const file = folder.createFile(blob); file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = `https://drive.google.com/thumbnail?id=${file.getId()}&sz=s1000`;
    
    if (sheet.getLastRow() < 2) sheet.appendRow(['MainLogo', fileUrl]); else sheet.getRange(2, 2).setValue(fileUrl);
    logSystem("Change Logo", "Updated system logo", username); return { status: 'success', url: fileUrl };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}
