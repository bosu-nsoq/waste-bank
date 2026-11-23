/**
 * ระบบธนาคารขยะ (Waste Bank) - Backend Script
 * พัฒนาสำหรับ Google Apps Script & รองรับเชื่อมต่อ GitHub Pages
 * Script ID: 1p7TBhgvIctwB854mT3PJLdmR_Q54a20mGxYa3o7jqVkP5stFYxAZHqD1
 * รหัส PIN ผู้ดูแลระบบเริ่มต้น: 8055
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('Page')
    .evaluate()
    .setTitle('Waste bank - ระบบธนาคารขยะ')
    .setFaviconUrl('https://i.postimg.cc/3wqjKXmy/Waste-Bank.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 🌐 Web API Endpoint สำหรับรองรับการเรียกจากภายนอก (GitHub Pages)
function doPost(e) {
  try {
    var contents = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    var data = JSON.parse(contents);
    var action = data.action;
    var args = data.args || [];
    
    var fn = (typeof globalThis !== 'undefined' && typeof globalThis[action] === 'function') 
             ? globalThis[action] 
             : (typeof this[action] === 'function' ? this[action] : null);

    if (!fn) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "error",
        error: "ไม่พบฟังก์ชัน: " + action 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var result = fn.apply(null, args);
    return ContentService.createTextOutput(JSON.stringify(result !== undefined ? result : { status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "error", 
      message: err.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// -------------------------------------------------------------
// Helper: ตรวจสอบและสร้างชีตเริ่มต้นหากยังไม่มี
// -------------------------------------------------------------
function getOrCreateSheet(sheetName, defaultHeaders, defaultRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (defaultHeaders && defaultHeaders.length > 0) {
      sheet.appendRow(defaultHeaders);
      sheet.getRange(1, 1, 1, defaultHeaders.length).setFontWeight('bold').setBackground('#E8F5E9');
    }
    if (defaultRows && defaultRows.length > 0) {
      defaultRows.forEach(r => sheet.appendRow(r));
    }
  }
  return sheet;
}

function initSheetsIfNeeded() {
  getOrCreateSheet('Settings', ['รหัส PIN', 'ชื่อเจ้าหน้าที่'], [
    ['8055', 'ผู้ดูแลระบบ (Admin)'],
    ['1234', 'ครูผู้ดูแลธนาคารขยะ']
  ]);

  getOrCreateSheet('Prices', ['ประเภทวัสดุ', 'คะแนนต่อชิ้น'], [
    ['ขวดพลาสติกใส (PET)', 1],
    ['ขวดพลาสติกขุ่น (HDPE)', 1],
    ['กระป๋องอลูมิเนียม', 2],
    ['ลังกระดาษ/กล่อง', 1],
    ['กระดาษขาวดำ/สมุด', 1],
    ['แก้วน้ำพลาสติก', 0.5],
    ['ขวดแก้ว', 1]
  ]);

  getOrCreateSheet('Users', ['รหัสนักเรียน', 'ชื่อ-สกุล', 'ชั้นเรียน', 'คะแนนสะสม'], [
    ['1001', 'ด.ช. กิตติศักดิ์ รักษ์โลก', 'ป.4/1', 25],
    ['1002', 'ด.ญ. สุพรรณษา รักสะอาด', 'ป.4/1', 40],
    ['1003', 'ด.ช. ธีรภัทร ใจดี', 'ป.5/2', 15],
    ['1004', 'ด.ญ. ปนัดดา รักธรรมชาติ', 'ป.6/1', 60]
  ]);

  getOrCreateSheet('Transactions', ['วัน-เวลา', 'ผู้ทำรายการ', 'รหัสนักเรียน', 'ชื่อ-สกุล', 'ชั้นเรียน', 'ประเภท', 'วัสดุ', 'จำนวน/น้ำหนัก', 'คะแนน'], []);
}

// -------------------------------------------------------------
// 1. ตรวจสอบ PIN Admin (รองรับ PIN 8055 เป็น Master PIN เสมอ)
// -------------------------------------------------------------
function checkAdminPin(inputPin) {
  if (!inputPin) return { valid: false, name: '' };
  const pinStr = inputPin.toString().trim();

  // ตรวจสอบ Master PIN 8055
  if (pinStr === '8055') {
    return { valid: true, name: 'ผู้ดูแลระบบ (Admin)' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Settings');
    if (!sheet) {
      initSheetsIfNeeded();
      sheet = ss.getSheetByName('Settings');
    }
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === pinStr) {
        return { valid: true, name: data[i][1] ? data[i][1].toString() : 'เจ้าหน้าที่' };
      }
    }
  } catch (e) {
    console.error('checkAdminPin error: ', e);
  }

  return { valid: false, name: '' };
}

// -------------------------------------------------------------
// 2. ดึงรายการวัสดุและอัตราคะแนน
// -------------------------------------------------------------
function getPrices() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Prices');
    if (!sheet) {
      initSheetsIfNeeded();
      sheet = ss.getSheetByName('Prices');
    }
    const data = sheet.getDataRange().getValues();
    let prices = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        prices.push({ 
          name: data[i][0].toString(), 
          price: parseFloat(data[i][1]) || 0 
        });
      }
    }
    if (prices.length === 0) {
      initSheetsIfNeeded();
      return getPrices();
    }
    return prices;
  } catch (e) {
    console.error('getPrices error:', e);
    return [
      { name: 'ขวดพลาสติกใส (PET)', price: 1 },
      { name: 'ขวดพลาสติกขุ่น (HDPE)', price: 1 },
      { name: 'กระป๋องอลูมิเนียม', price: 2 },
      { name: 'ลังกระดาษ/กล่อง', price: 1 },
      { name: 'แก้วน้ำพลาสติก', price: 0.5 },
      { name: 'ขวดแก้ว', price: 1 }
    ];
  }
}

// -------------------------------------------------------------
// 3. ค้นหาข้อมูลนักเรียนจากรหัส 4 หลัก
// -------------------------------------------------------------
function getStudentInfo(studentId) {
  if (!studentId) return { found: false, name: '', className: '', balance: 0 };
  const targetId = studentId.toString().trim();
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Users');
    if (!sheet) {
      initSheetsIfNeeded();
      sheet = ss.getSheetByName('Users');
    }
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === targetId) {
        return { 
          found: true, 
          studentId: targetId,
          name: data[i][1] ? data[i][1].toString() : '', 
          className: data[i][2] ? data[i][2].toString() : '', 
          balance: parseFloat(data[i][3]) || 0 
        };
      }
    }
  } catch (e) {
    console.error('getStudentInfo error:', e);
  }
  return { found: false, studentId: targetId, name: '', className: '', balance: 0 };
}

// -------------------------------------------------------------
// 3.1 ดึงประวัติธุรกรรมของนักเรียน (สำหรับโหมด Guest ตรวจสอบ)
// -------------------------------------------------------------
function getStudentHistory(studentId) {
  if (!studentId) return [];
  const targetId = studentId.toString().trim();
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let txSheet = ss.getSheetByName('Transactions');
    if (!txSheet) return [];
    
    const data = txSheet.getDataRange().getValues();
    let history = [];
    
    for (let i = data.length - 1; i >= 1; i--) {
      let row = data[i];
      if (row[2] && row[2].toString().trim() === targetId) {
        let d = new Date(row[0]);
        let y = d.getFullYear();
        let yBE = y > 2400 ? y : y + 543;
        let dateStr = !isNaN(d.getTime()) 
          ? (Utilities.formatDate(d, "GMT+7", "dd/MM/") + yBE + " " + Utilities.formatDate(d, "GMT+7", "HH:mm")) 
          : String(row[0]);
        
        history.push({
          date: dateStr,
          admin: row[1] || '-',
          type: row[5] || 'ฝาก',
          material: row[6] || '-',
          weight: row[7] !== '' && row[7] !== '-' ? parseFloat(row[7]) : '-',
          amount: parseFloat(row[8]) || 0
        });
        
        if (history.length >= 20) break; // ดึงย้อนหลัง 20 รายการล่าสุด
      }
    }
    return history;
  } catch (e) {
    console.error('getStudentHistory error:', e);
    return [];
  }
}

// -------------------------------------------------------------
// 4. บันทึกการฝากขยะ
// -------------------------------------------------------------
function saveDeposit(adminName, studentId, studentName, studentClass, items) {
  if (!items || !items.length) return 'ไม่มีรายการขยะที่ต้องบันทึก';
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let txSheet = ss.getSheetByName('Transactions');
  let userSheet = ss.getSheetByName('Users');
  if (!txSheet || !userSheet) {
    initSheetsIfNeeded();
    txSheet = ss.getSheetByName('Transactions');
    userSheet = ss.getSheetByName('Users');
  }
  
  const date = new Date();
  let totalAmount = 0;
  
  items.forEach(item => {
    let amt = parseFloat(item.amount) || 0;
    let wgt = parseFloat(item.weight) || 0;
    txSheet.appendRow([
      date, 
      adminName || 'Admin', 
      studentId.toString().trim(), 
      studentName, 
      studentClass, 
      'ฝาก', 
      item.material, 
      wgt, 
      amt
    ]);
    totalAmount += amt;
  });
  
  updateUserBalance(userSheet, studentId, studentName, studentClass, totalAmount);
  return 'บันทึกการฝากสำเร็จ ยอดรวม: ' + totalAmount.toLocaleString('th-TH') + ' คะแนน';
}

// -------------------------------------------------------------
// 5. ระบบแลกคะแนน
// -------------------------------------------------------------
function withdrawMoney(adminName, studentId, studentName, studentClass, amount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let txSheet = ss.getSheetByName('Transactions');
  let userSheet = ss.getSheetByName('Users');
  if (!txSheet || !userSheet) {
    initSheetsIfNeeded();
    txSheet = ss.getSheetByName('Transactions');
    userSheet = ss.getSheetByName('Users');
  }
  
  let numAmount = parseFloat(amount);
  if (isNaN(numAmount) || numAmount <= 0) return 'จำนวนคะแนนที่ต้องการแลกไม่ถูกต้อง';
  
  let studentData = getStudentInfo(studentId);
  if (studentData.balance < numAmount) {
    return 'คะแนนไม่พอให้แลก (คงเหลือ ' + studentData.balance + ' คะแนน)';
  }
  
  txSheet.appendRow([
    new Date(), 
    adminName || 'Admin', 
    studentId.toString().trim(), 
    studentName, 
    studentClass, 
    'แลก', 
    '-', 
    '-', 
    -numAmount
  ]);
  
  updateUserBalance(userSheet, studentId, studentName, studentClass, -numAmount);
  return 'แลกสำเร็จ จำนวน ' + numAmount.toLocaleString('th-TH') + ' คะแนน';
}

// อัปเดตยอดคงเหลือในชีต Users
function updateUserBalance(sheet, studentId, studentName, studentClass, amount) {
  const data = sheet.getDataRange().getValues();
  let found = false;
  const sid = studentId.toString().trim();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === sid) {
      let currentBal = parseFloat(data[i][3]) || 0;
      let newBalance = currentBal + amount;
      if (studentName) sheet.getRange(i + 1, 2).setValue(studentName);
      if (studentClass) sheet.getRange(i + 1, 3).setValue(studentClass);
      sheet.getRange(i + 1, 4).setValue(newBalance);
      found = true;
      break;
    }
  }
  if (!found) { 
    sheet.appendRow([sid, studentName, studentClass, amount]);
  }
}

// -------------------------------------------------------------
// 6. ดึงข้อมูลสถิติ
// -------------------------------------------------------------
function getStatistics() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Transactions');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    let stats = {};
    let totalItems = 0;
    let totalPoints = 0;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][5] === 'ฝาก') {
        let mat = data[i][6] ? data[i][6].toString() : 'ทั่วไป';
        let weight = parseFloat(data[i][7]) || 0;
        let points = parseFloat(data[i][8]) || 0;
        if (!stats[mat]) stats[mat] = { weight: 0, points: 0 };
        stats[mat].weight += weight;
        stats[mat].points += points;
        totalItems += weight;
        totalPoints += points;
      }
    }
    
    let result = Object.keys(stats).map(key => ({
      material: key,
      weight: Math.round(stats[key].weight * 100) / 100,
      points: Math.round(stats[key].points * 100) / 100
    }));
    
    result.sort((a, b) => b.weight - a.weight);
    return result;
  } catch (e) {
    console.error('getStatistics error:', e);
    return [];
  }
}

// -------------------------------------------------------------
// 6.1 ดึงอันดับนักเรียนรักษ์โลก (Leaderboard)
// -------------------------------------------------------------
function getLeaderboard() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Users');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    let list = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1]) {
        list.push({
          studentId: data[i][0].toString(),
          name: data[i][1].toString(),
          className: data[i][2] ? data[i][2].toString() : '-',
          balance: parseFloat(data[i][3]) || 0
        });
      }
    }
    list.sort((a, b) => b.balance - a.balance);
    return list.slice(0, 10); // 10 อันดับแรก
  } catch (e) {
    console.error('getLeaderboard error:', e);
    return [];
  }
}

// -------------------------------------------------------------
// 7. ดึงรายงานตามช่วงวันที่ (รองรับทั้ง YYYY-MM-DD และ DD/MM/YYYY ตามเกณฑ์ พ.ศ.)
// -------------------------------------------------------------
function parseDateInput(val) {
  if (!val) return null;
  val = String(val).trim();
  let y, m, d;
  if (val.includes('/')) {
    let parts = val.split('/');
    d = parseInt(parts[0], 10);
    m = parseInt(parts[1], 10) - 1;
    y = parseInt(parts[2], 10);
  } else if (val.includes('-')) {
    let parts = val.split('-');
    y = parseInt(parts[0], 10);
    m = parseInt(parts[1], 10) - 1;
    d = parseInt(parts[2], 10);
  } else {
    let dt = new Date(val);
    return isNaN(dt.getTime()) ? null : dt;
  }
  if (y > 2400) y -= 543; // แปลง พ.ศ. กลับเป็น ค.ศ. สำหรับเปรียบเทียบในระบบ
  return new Date(y, m, d);
}

function getReport(startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Transactions');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    let report = [];
    
    let startDt = parseDateInput(startDate);
    let endDt = parseDateInput(endDate);
    
    let start = startDt ? startDt.setHours(0,0,0,0) : 0;
    let end = endDt ? endDt.setHours(23,59,59,999) : Number.MAX_SAFE_INTEGER;
    
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      let d = new Date(data[i][0]);
      let txTime = d.getTime();
      
      if (!isNaN(txTime) && txTime >= start && txTime <= end) {
        let y = d.getFullYear();
        let yBE = y > 2400 ? y : y + 543;
        let formattedDate = Utilities.formatDate(d, "GMT+7", "dd/MM/") + yBE + " " + Utilities.formatDate(d, "GMT+7", "HH:mm");
        
        report.push({
          date: formattedDate,
          admin: data[i][1] || '-',
          studentId: data[i][2] || '-',
          studentName: data[i][3] || '-',
          studentClass: data[i][4] || '-',
          type: data[i][5] || '-',
          material: data[i][6] || '-',
          weight: data[i][7] !== '' && data[i][7] !== '-' ? data[i][7] : '-',
          amount: parseFloat(data[i][8]) || 0
        });
      }
    }
    return report.reverse(); // เรียงจากล่าสุดไปเก่าสุด
  } catch (e) {
    console.error('getReport error:', e);
    return [];
  }
}

// -------------------------------------------------------------
// 8. ดึงข้อมูลภาพรวมระบบในครั้งเดียว (เร็วขึ้น ไม่ต้องเรียกหลายรอบ)
// -------------------------------------------------------------
function getSystemOverview() {
  return {
    prices: getPrices(),
    stats: getStatistics(),
    leaderboard: getLeaderboard()
  };
}
