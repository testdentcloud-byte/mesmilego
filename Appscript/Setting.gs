
// Utility: เลือกชื่อชีตตามสาขา
function getSheetNameByBranch(baseName, branch) {
  // กรณีสาขาหลัก (1) ใช้ชื่อเดิม
  if (!branch || branch === 'main' || branch === '1') {
    return baseName;
  }
  
  // กรณีสาขาอื่นๆ (2, 3, ...) ให้เพิ่มหมายเลขสาขาต่อท้าย
  // รูปแบบ: ชื่อชีตเดิม_สาขา2, ชื่อชีตเดิม_สาขา3
  return `${baseName}_สาขา${branch}`;
}
const SPREADSHEET_ID = '1djH94Ht78Ig2sUoxdXaVzL37fcggxHSKeafHaWggcsY';
const SHEET_NAME_MEMBER = 'รายชื่อสมาชิก';
const SHEET_NAME_SERVICE = 'บริการ';
const SHEET_NAME_SETTING = 'ตั้งค่าทั่วไป';
const SHEET_ADMIN = 'แอดมิน';
const DRIVE_FOLDER_ID = '15T8WipLolB2Yb31OXS_TLHr5QwlbWNuk'; // ตรวจสอบให้แน่ใจว่าเป็น Folder ID ที่ถูกต้อง

// --- Web Entry Point ---
function doGet(e) {
  const action = e.parameter.action;
  const branch = e.parameter.branch || '1';

  // =========================
  // 1) กรณี adminLogin (เช็ก username/password)
  // =========================
  if (action === 'adminLogin') {
    // ดึง username/password จาก query params
    const username = e.parameter.username || '';
    const password = e.parameter.password || '';
    // เรียกฟังก์ชันตรวจสอบจากชีต "แอดมิน"
    const result = checkAdminCredentials({ username, password });

    // คืนค่า JSON (ไม่ต้องเช็ค branchId แต่อย่างใด)
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'fetchData') return fetchServiceData(branch);                // บริการ
  if (action === 'fetchTimeSlots') return fetchTimeSlots(branch);             // ตั้งค่า - ช่วงเวลา
  if (action === 'fetchHolidays') return fetchHolidays(branch);               // ตั้งค่า - วันหยุด
  if (action === 'fetchMemberData') return fetchMembers(branch);              // สมาชิก
  if (action === 'fetchGeneralSettings') return fetchGeneralSettings(branch);
  if (action === 'fetchDoctors') return fetchDoctors(branch);                 // รายชื่อหมอ

  // เพิ่มการจัดการข้อผิดพลาดสำหรับ action ที่ไม่รู้จัก
  return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let msg;
  try {
    const req = JSON.parse(e.postData.contents);

    // สมาชิก
    if (req.action === 'insertOrUpdateMember') msg = insertOrUpdateMember(req.data);
    else if (req.action === 'deleteMember') msg = deleteMember(req.id, req.branch);

    // บริการ
    else if (req.action === 'insertOrUpdateService') msg = insertOrUpdateService(req.data); // <-- ตรงนี้ที่แก้ไข
    else if (req.action === 'deleteService') msg = deleteService(req.id, req.branch);

    // ตั้งค่าทั่วไป
    else if (req.action === 'insertOrUpdateTimeSlot') msg = insertOrUpdateTimeSlot(req.data);
    else if (req.action === 'deleteTimeSlot') msg = deleteTimeSlot(req.id, req.branch);
    else if (req.action === 'addHoliday') msg = addHoliday(req.date, req.branch);
    else if (req.action === 'deleteHoliday') msg = deleteHoliday(req.date, req.branch);
    else if (req.action === 'updatePermanentHolidays') msg = updatePermanentHolidays(req.data, req.branch);

    // รายชื่อหมอ
    else if (req.action === 'insertOrUpdateDoctor') msg = insertOrUpdateDoctor(req.data);
    else if (req.action === 'deleteDoctor') msg = deleteDoctor(req.id, req.branch);

    else throw new Error('Unknown action');

    // แก้ไข: ให้ return เป็น object ที่มี success: true เสมอเมื่อสำเร็จ
    return ContentService.createTextOutput(JSON.stringify(msg)) // msg ควรเป็น {success: true, message: "...", id: "..."}
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // แก้ไข: Catch error และส่งกลับ response ที่ชัดเจน
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: `Error: ${error.message}` }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// === สมาชิก ===
function fetchMembers(branch) {
  branch = branch || '1';
  const sheetName = getSheetNameByBranch(SHEET_NAME_MEMBER, branch);
  // ปรับหัวตารางใหม่: ['รหัส', 'ไลน์ไอดี', 'ชื่อ', 'นามสกุล', 'เบอร์โทร', 'เลขบัตร/ประกัน', 'โรค/แพ้', 'หมายเหตุ', 'เวลาบันทึก']
  const sheet = getOrCreateSheet(sheetName, ['รหัส', 'ไลน์ไอดี', 'ชื่อ', 'นามสกุล', 'เบอร์โทร', 'เลขบัตร/ประกัน', 'โรค/แพ้', 'หมายเหตุ', 'เวลาบันทึก']);
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.shift();
  // Map to object with key ภาษาไทย
  return ContentService.createTextOutput(JSON.stringify(values.map(r => ({
    'รหัส': r[0],
    'ไลน์ไอดี': r[1],
    'ชื่อ': r[2],
    'นามสกุล': r[3],
    'เบอร์โทร': r[4],
    'เลขบัตร/ประกัน': r[5],
    'โรค/แพ้': r[6],
    'หมายเหตุ': r[7],
    'เวลาบันทึก': r[8]
  }))))
    .setMimeType(ContentService.MimeType.JSON);
}

function insertOrUpdateMember(data) {
  const branch = data.branch || '1';
  const sheetName = getSheetNameByBranch(SHEET_NAME_MEMBER, branch);
  // ปรับหัวตารางใหม่: ['รหัส', 'UserID', 'ชื่อ', 'นามสกุล', 'เบอร์โทร', 'เลขบัตร/ประกัน', 'โรค/แพ้', 'หมายเหตุ', 'เวลาบันทึก']
  const sheet = getOrCreateSheet(sheetName, ['รหัส', 'UserID', 'ชื่อ', 'นามสกุล', 'เบอร์โทร', 'เลขบัตร/ประกัน', 'โรค/แพ้', 'หมายเหตุ', 'เวลาบันทึก']);
  const values = sheet.getDataRange().getValues();
  const now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');

  // Map key อังกฤษ → ไทย ถ้ายังไม่มี key ภาษาไทย
  if (!data['ชื่อ'] && data['firstName']) data['ชื่อ'] = data['firstName'];
  if (!data['นามสกุล'] && data['lastName']) data['นามสกุล'] = data['lastName'];
  if (!data['เลขบัตร/ประกัน'] && data['idCardOrSocial']) data['เลขบัตร/ประกัน'] = data['idCardOrSocial'];
  if (!data['โรค/แพ้'] && data['diseaseAllergy']) data['โรค/แพ้'] = data['diseaseAllergy'];
  if (!data['หมายเหตุ'] && data['note']) data['หมายเหตุ'] = data['note'];

  if (!data['ชื่อ'] || !data['นามสกุล'] || !data['เบอร์โทร']) throw new Error('ต้องระบุชื่อ นามสกุล และเบอร์โทร');

  let found = false;
  for (let i = 1; i < values.length; i++) { // เริ่มจาก 1 เพื่อข้าม header
    if (String(values[i][0]) === String(data['รหัส'])) { // เปรียบเทียบ ID เป็น String
      sheet.getRange(i + 1, 2, 1, 8).setValues([[data['UserID'] || '', data['ชื่อ'], data['นามสกุล'], data['เบอร์โทร'], data['เลขบัตร/ประกัน'] || '', data['โรค/แพ้'] || '', data['หมายเหตุ'] || '', now]]);
      found = true;
      break;
    }
  }

  if (!found) {
    const newId = String(sheet.getLastRow() + 1);
    sheet.appendRow([newId, data['UserID'] || '', data['ชื่อ'], data['นามสกุล'], data['เบอร์โทร'], data['เลขบัตร/ประกัน'] || '', data['โรค/แพ้'] || '', data['หมายเหตุ'] || '', now]);
    return { success: true, message: 'เพิ่มข้อมูลสมาชิกใหม่เรียบร้อย', id: newId };
  }
  return { success: true, message: 'อัปเดตข้อมูลสมาชิกเรียบร้อย', id: data['รหัส'] };
}

function deleteMember(id, branch) {
  branch = branch || '1';
  const sheetName = getSheetNameByBranch(SHEET_NAME_MEMBER, branch);
  const sheet = getOrCreateSheet(sheetName);
  const values = sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) { // เริ่มจาก 1 เพื่อข้าม header
    if (String(values[i][0]) === String(id)) { // เปรียบเทียบ ID เป็น String
      sheet.deleteRow(i + 1);
      return { success: true, message: 'ลบข้อมูลสมาชิกแล้ว' };
    }
  }
  return { success: false, message: 'ไม่พบ ID สมาชิกที่จะลบ' };
}

// === บริการ ===
function fetchServiceData(branch) {
  branch = branch || '1';
  const sheetName = getSheetNameByBranch(SHEET_NAME_SERVICE, branch);
  const sheet = getOrCreateSheet(sheetName, ['ไอดี', 'รายการ', 'รายละเอียด', 'ราคา', 'รูปภาพ']);
  const values = sheet.getDataRange().getDisplayValues(); // ใช้ getDisplayValues() เพื่อให้ได้ค่าตามที่แสดงในชีท (รวมถึง URL รูปภาพ)
  const headers = values.length > 0 ? values.shift() : []; // ตรวจสอบว่ามีข้อมูลก่อน shift headers

  // ตรวจสอบว่า headers มีข้อมูลครบถ้วนหรือไม่ เพื่อป้องกัน index out of bounds
  if (headers.length === 0) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Map values to objects, ensuring all expected keys are present
  const data = values.map(r => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = r[i] !== undefined ? r[i] : ''; // กำหนดค่าว่างถ้าไม่มีข้อมูลใน cell
    });
    return obj;
  });

  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function insertOrUpdateService(data) {
  const branch = data.branch || '1';
  const sheetName = getSheetNameByBranch(SHEET_NAME_SERVICE, branch);
  const sheet = getOrCreateSheet(sheetName, ['ไอดี', 'รายการ', 'รายละเอียด', 'ราคา', 'รูปภาพ']);
  const values = sheet.getDataRange().getValues(); // ใช้ getValues() เพื่อให้ได้ค่าดิบ (เช่น ID เป็นตัวเลข)

  let id = data['ไอดี']; // นี่คือ ID ที่ส่งมาจากฟอร์ม (อาจเป็นค่าว่างสำหรับรายการใหม่ หรือ ID เดิมสำหรับการอัปเดต)
  let isUpdate = false;
  let targetRowIndex = -1; // Index 0-based ในอาร์เรย์ values (headers อยู่ที่ index 0)

  // 1. ตรวจสอบว่า ID ที่ส่งมา มีอยู่ในชีทอยู่แล้วหรือไม่
  if (id) {
    // เริ่มค้นหาจากแถวที่ 2 (index 1) เพื่อข้ามหัวข้อ
    for (let i = 1; i < values.length; i++) {
      // เปรียบเทียบ ID โดยแปลงเป็น String ทั้งคู่เพื่อความแน่นอน
      if (String(values[i][0]) === String(id)) {
        targetRowIndex = i; // พบแถวที่ต้องการอัปเดต
        isUpdate = true;
        break;
      }
    }
  }

  // 2. กำหนด ID หากเป็นรายการใหม่
  if (!isUpdate) {
    id = String(sheet.getLastRow() + 1); // ID ใหม่สำหรับรายการใหม่
  }

  // 3. จัดการรูปภาพ
  let imageUrl = data['รูปภาพOLD'] || ''; // เริ่มต้นด้วย URL รูปภาพเก่า (ถ้ามี) หรือสตริงว่าง
  if (data['รูปภาพBase64']) {
    // ถ้ามีการอัปโหลดรูปภาพใหม่ ให้สร้าง URL ใหม่
    imageUrl = generateFileUrl(data['รูปภาพBase64'], data['รูปภาพMimeType'], `${data['รายการ'] || 'service'}_${Date.now()}`, DRIVE_FOLDER_ID);
  } else if (isUpdate && !data['รูปภาพOLD']) {
    // ถ้าเป็นการอัปเดตแต่ไม่มีรูปภาพเก่าส่งมา (เช่น ลบรูปภาพทิ้ง) ให้ตั้งค่าเป็นว่าง
    imageUrl = '';
  }
  // ถ้าไม่มีการอัปโหลดใหม่ และมีรูปภาพเก่า ก็จะใช้ค่าจาก data['รูปภาพOLD'] ที่ตั้งไว้แต่แรก

  // 4. เตรียมข้อมูลสำหรับบันทึก (ราคาเป็น 0 ถ้าไม่มีข้อมูลหรือไม่ใช่ตัวเลข)
  let price = parseFloat(data['ราคา']);
  if (isNaN(price)) price = 0;
  const rowData = [
    String(id), // ID ต้องเป็น String เพื่อความสอดคล้อง
    data['รายการ'] || '',
    data['รายละเอียด'] || '',
    price,
    imageUrl
  ];

  // 5. ดำเนินการบันทึก
  try {
    if (isUpdate && targetRowIndex !== -1) {
      // อัปเดตแถวที่มีอยู่ (targetRowIndex + 1 เพราะ Range เป็น 1-based index)
      sheet.getRange(targetRowIndex + 1, 1, 1, rowData.length).setValues([rowData]);
      return { success: true, message: 'อัปเดตข้อมูลบริการเรียบร้อย', id: id };
    } else {
      // เพิ่มแถวใหม่
      sheet.appendRow(rowData);
      return { success: true, message: 'เพิ่มข้อมูลบริการใหม่เรียบร้อย', id: id };
    }
  } catch (e) {
    return { success: false, message: `เกิดข้อผิดพลาดในการบันทึกข้อมูล: ${e.message}` };
  }
}

function deleteService(id, branch) {
  branch = branch || '1';
  const sheetName = getSheetNameByBranch(SHEET_NAME_SERVICE, branch);
  const sheet = getOrCreateSheet(sheetName);
  const values = sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) { // เริ่มจาก 1 เพื่อข้าม header
    if (String(values[i][0]) === String(id)) { // เปรียบเทียบ ID เป็น String
      sheet.deleteRow(i + 1);
      return { success: true, message: 'ลบข้อมูลบริการแล้ว' };
    }
  }
  return { success: false, message: 'ไม่พบ ID บริการที่จะลบ' };
}

// === ตั้งค่าทั่วไป ===
function fetchTimeSlots(branch) {
  branch = branch || '1';
  const sheetName = getSheetNameByBranch(SHEET_NAME_SETTING, branch);
  const sheet = getOrCreateSheet(sheetName, ['ไอดี', 'ช่วงเวลา', 'สลอต', 'วันหยุด']);
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.shift();
  // Filter out rows that are not valid time slots (e.g., just holiday entries)
  const timeSlots = values.map(r => ({
    'ไอดี': r[0] || '',
    'ช่วงเวลา': r[1] || '',
    'สลอต': r[2] || ''
  })).filter(item => item['ช่วงเวลา'] !== ''); // Only return rows that actually have a time slot defined

  return ContentService.createTextOutput(JSON.stringify(timeSlots))
    .setMimeType(ContentService.MimeType.JSON);
}


function insertOrUpdateTimeSlot(data) {
  const branch = data.branch || '1';
  const sheetName = getSheetNameByBranch(SHEET_NAME_SETTING, branch);
  const sheet = getOrCreateSheet(sheetName, ['ไอดี', 'ช่วงเวลา', 'สลอต', 'วันหยุด']);
  const values = sheet.getDataRange().getValues();
  let id = data['ไอดี'];
  let isUpdate = false;
  let targetRowIndex = -1;

  if (id) {
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(id)) {
        targetRowIndex = i;
        isUpdate = true;
        break;
      }
    }
  }

  if (!isUpdate) {
    id = String(sheet.getLastRow() + 1); // Generate new ID for new entry
    // Try to find an empty row for a new slot without creating a new ID
    for (let i = 1; i < values.length; i++) {
      if (!values[i][1] && !values[i][2] && !values[i][3]) { // Check if row is mostly empty (excluding ID)
        targetRowIndex = i;
        isUpdate = true; // Treat as update to an empty row
        break;
      }
    }
    if (!isUpdate) { // Still no empty row found, truly append
      sheet.appendRow([id, data['ช่วงเวลา'] || '', data['สลอต'] || '', '']);
      return { success: true, message: 'เพิ่มช่วงเวลาใหม่แล้ว', id: id };
    }
  }

  // If we found a target row (either existing ID or empty row for new slot)
  const rowData = [
    String(id),
    data['ช่วงเวลา'] || '',
    data['สลอต'] || '',
    (targetRowIndex !== -1 && values[targetRowIndex][3]) ? values[targetRowIndex][3] : '' // Preserve existing holiday if any
  ];
  sheet.getRange(targetRowIndex + 1, 1, 1, rowData.length).setValues([rowData]);
  return { success: true, message: 'อัปเดตช่วงเวลาแล้ว', id: id };
}


function deleteTimeSlot(id, branch) {
  branch = branch || '1';
  const sheetName = getSheetNameByBranch(SHEET_NAME_SETTING, branch);
  const sheet = getOrCreateSheet(sheetName);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      // Clear values but don't delete row, to preserve potential holidays in column D
      sheet.getRange(i + 1, 1, 1, 3).setValues([['', '', '']]); // Clear ID, ช่วงเวลา, สลอต
      return { success: true, message: 'ลบช่วงเวลาแล้ว' };
    }
  }
  return { success: false, message: 'ไม่พบ ID ช่วงเวลา' };
}


function fetchHolidays(branch) {
  branch = branch || '1';
  const sheetName = getSheetNameByBranch(SHEET_NAME_SETTING, branch);
  const sheet = getOrCreateSheet(sheetName);
  const values = sheet.getDataRange().getDisplayValues();
  values.shift(); // Remove header row
  // Filter out empty values and return only non-empty holiday dates
  return ContentService.createTextOutput(JSON.stringify(values.map(r => r[3]).filter(d => d && d.trim() !== '')))
    .setMimeType(ContentService.MimeType.JSON);
}


function addHoliday(dateStr, branch) {
  branch = branch || '1';
  const sheetName = getSheetNameByBranch(SHEET_NAME_SETTING, branch);
  const [yyyy, mm, dd] = dateStr.split('-');
  const formatted = `${dd}-${mm}-${yyyy}`; // Format to DD-MM-YYYY
  const sheet = getOrCreateSheet(sheetName);
  const values = sheet.getDataRange().getValues();

  // Check if holiday already exists
  if (values.some(r => r[3] === formatted)) {
    return { success: false, message: 'มีวันหยุดนี้แล้ว' };
  }

  // Find an empty slot in column D for holiday
  for (let i = 1; i < values.length; i++) {
    if (!values[i][3]) { // If column D (index 3) is empty
      sheet.getRange(i + 1, 4).setValue(formatted);
      return { success: true, message: 'เพิ่มวันหยุดแล้ว' };
    }
  }

  // If no empty slot, append a new row (might create a new ID)
  const id = String(sheet.getLastRow() + 1);
  sheet.appendRow([id, '', '', formatted]);
  return { success: true, message: 'เพิ่มวันหยุดใหม่แล้ว' };
}


function deleteHoliday(dateStr, branch) {
  branch = branch || '1';
  const sheetName = getSheetNameByBranch(SHEET_NAME_SETTING, branch);
  const sheet = getOrCreateSheet(sheetName);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][3] === dateStr) {
      sheet.getRange(i + 1, 4).setValue(''); // Clear the holiday cell
      return { success: true, message: 'ลบวันหยุดแล้ว' };
    }
  }
  return { success: false, message: 'ไม่พบวันหยุดนี้' };
}


// === ฟังก์ชันสำหรับจัดการวันหยุดประจำ (เพิ่มใหม่) ===

// ฟังก์ชันสำหรับอ่านค่าวันหยุดประจำจากชีต
function fetchGeneralSettings(branch) {
  branch = branch || '1';
  const sheetName = getSheetNameByBranch(SHEET_NAME_SETTING, branch);
  const sheet = getOrCreateSheet(sheetName, ['ไอดี', 'ช่วงเวลา', 'สลอต', 'วันหยุด', 'วันหยุดประจำสัปดาห์ (0-6)', 'รายชื่อหมอ']);
  // อ่านค่าจากเซลล์ E2 (แถว 2, คอลัมน์ 5)
  const permanentHolidays = sheet.getRange(2, 5).getValue();
  return ContentService.createTextOutput(JSON.stringify({ permanentHolidays: permanentHolidays }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ฟังก์ชันสำหรับบันทึก/อัปเดตค่าวันหยุดประจำ
function updatePermanentHolidays(days, branch) {
  branch = branch || '1';
  const sheetName = getSheetNameByBranch(SHEET_NAME_SETTING, branch);
  if (!Array.isArray(days)) {
    return { success: false, message: 'ข้อมูลไม่ถูกต้อง' };
  }
  const sheet = getOrCreateSheet(sheetName);
  // แปลง Array ['6', '0'] เป็น String "6,0" แล้วบันทึกลงเซลล์ E2
  const valueToSave = days.join(',');
  sheet.getRange(2, 5).setValue(valueToSave);
  return { success: true, message: 'บันทึกการตั้งค่าวันหยุดประจำเรียบร้อย' };
}

// === รายชื่อหมอ ===
function fetchDoctors(branch) {
  branch = branch || '1';
  const sheetName = getSheetNameByBranch(SHEET_NAME_SETTING, branch);
  const sheet = getOrCreateSheet(sheetName, ['ไอดี', 'ช่วงเวลา', 'สลอต', 'วันหยุด', 'วันหยุดประจำสัปดาห์ (0-6)', 'รายชื่อหมอ']);
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.shift();
  // Filter out rows that have doctor names (column F - index 5)
  const doctors = values.map(r => ({
    'ไอดี': r[0] || '',
    'รายชื่อหมอ': r[5] || ''
  })).filter(item => item['รายชื่อหมอ'] !== ''); // Only return rows with doctor names

  return ContentService.createTextOutput(JSON.stringify(doctors))
    .setMimeType(ContentService.MimeType.JSON);
}

function insertOrUpdateDoctor(data) {
  const branch = data.branch || '1';
  const sheetName = getSheetNameByBranch(SHEET_NAME_SETTING, branch);
  const sheet = getOrCreateSheet(sheetName, ['ไอดี', 'ช่วงเวลา', 'สลอต', 'วันหยุด', 'วันหยุดประจำสัปดาห์ (0-6)', 'รายชื่อหมอ']);
  const values = sheet.getDataRange().getValues();
  let id = data['ไอดี'];
  let isUpdate = false;
  let targetRowIndex = -1;

  if (!data['รายชื่อหมอ'] || data['รายชื่อหมอ'].trim() === '') {
    return { success: false, message: 'กรุณาระบุชื่อหมอ' };
  }

  // ถ้ามี ID แล้ว (กรณีแก้ไข) ให้หาแถวที่มี ID นั้น
  if (id) {
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(id)) {
        targetRowIndex = i;
        isUpdate = true;
        break;
      }
    }
  }

  // กรณีเพิ่มใหม่ (ไม่มี ID หรือหาไม่เจอ)
  if (!isUpdate) {
    // ค้นหาแถวที่ยังไม่มีชื่อหมอในคอลัมน์ F โดยเริ่มจากแถว 2
    let emptyDoctorRowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (!values[i][5] || values[i][5].toString().trim() === '') { // คอลัมน์ F (index 5) ว่าง
        emptyDoctorRowIndex = i;
        break; // หยุดทันทีเมื่อเจอแถวว่างแถวแรก
      }
    }

    if (emptyDoctorRowIndex !== -1) {
      // เจอแถวที่คอลัมน์ F ว่าง ให้ใช้แถวนั้น
      id = String(values[emptyDoctorRowIndex][0]) || String(emptyDoctorRowIndex + 1);
      sheet.getRange(emptyDoctorRowIndex + 1, 6).setValue(data['รายชื่อหมอ']); // อัปเดตเฉพาะคอลัมน์ F
      return { success: true, message: 'เพิ่มหมอใหม่แล้ว', id: id };
    } else {
      // ไม่เจอแถวว่าง ให้เพิ่มแถวใหม่ท้ายสุด
      id = String(sheet.getLastRow() + 1);
      sheet.appendRow([id, '', '', '', '', data['รายชื่อหมอ']]);
      return { success: true, message: 'เพิ่มหมอใหม่แล้ว', id: id };
    }
  }

  // กรณีอัปเดต (มี ID และเจอแถวแล้ว) - อัปเดตเฉพาะคอลัมน์รายชื่อหมอ
  sheet.getRange(targetRowIndex + 1, 6).setValue(data['รายชื่อหมอ']);
  return { success: true, message: 'อัปเดตข้อมูลหมอแล้ว', id: id };
}

function deleteDoctor(id, branch) {
  branch = branch || '1';
  const sheetName = getSheetNameByBranch(SHEET_NAME_SETTING, branch);
  const sheet = getOrCreateSheet(sheetName);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      // Clear only doctor name column (F - index 5)
      sheet.getRange(i + 1, 6).setValue('');
      return { success: true, message: 'ลบข้อมูลหมอแล้ว' };
    }
  }
  return { success: false, message: 'ไม่พบ ID หมอที่จะลบ' };
}



// === Shared ===
function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers) {
      sheet.appendRow(headers);
    }
  }
  return sheet;
}

function generateFileUrl(base64, mime, name, folderId) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), mime, name);
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return `https://lh5.googleusercontent.com/d/${file.getId()}`;
  } catch (e) {
    Logger.log(`Error generating file URL: ${e.message}`);
    throw new Error(`ไม่สามารถสร้าง URL รูปภาพได้: ${e.message}`);
  }
}


// ================================
// ----------------------------------------------------------------
// เปลี่ยนฟังก์ชันตรวจสอบ username/password (เพิ่ม displayName)
function checkAdminCredentials(data) {
  try {
    const usernameInput = (data.username || '').trim();
    const passwordInput = data.password || '';

    if (!usernameInput || !passwordInput) {
      return { success: false, message: 'กรุณากรอก username และ password' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_ADMIN);
    if (!sheet) {
      Logger.log('ไม่พบชีต "แอดมิน" ในสเปรดชีต');
      return { success: false, message: 'ระบบไม่พบข้อมูล Admin' };
    }

    // อ่านข้อมูลทั้งหมด ยกเว้นแถวหัวคอลัมน์
    // คอลัมน์ในชีต "แอดมิน" คือ:
    // A: ID | B: UserID | C: ชื่อ | D: เบอร์โทร | E: username | F: password | G: สาขาที่ดูแล
    const rows = sheet.getDataRange().getValues().slice(1);
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const storedUsername = String(row[4] || '').trim(); // คอลัมน์ E: username
      const storedPassword = String(row[5] || '');       // คอลัมน์ F: password (plaintext)

      if (usernameInput === storedUsername && passwordInput === storedPassword) {
        const displayName = String(row[2] || '').trim(); // คอลัมน์ C: ชื่อ
        const branches = String(row[6] || '1').trim();   // คอลัมน์ G: สาขาที่ดูแล (เช่น "1,2" หรือ "all")
        
        // แปลงสาขาเป็น array
        let branchArray = [];
        if (branches.toLowerCase() === 'all' || branches === '*') {
          branchArray = ['all']; // แอดมินดูแลทุกสาขา
        } else {
          branchArray = branches.split(',').map(b => b.trim()).filter(b => b);
          if (branchArray.length === 0) branchArray = ['1']; // default ถ้าไม่มีข้อมูล
        }
        
        return {
          success: true,
          message: 'ล็อกอินสำเร็จ',
          displayName: displayName || storedUsername,
          branches: branchArray, // ส่งรายการสาขาที่ดูแล
          defaultBranch: branchArray[0] === 'all' ? '1' : branchArray[0] // สาขาเริ่มต้น
        };
      }
    }

    return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  } catch (err) {
    Logger.log('Error in checkAdminCredentials: ' + err);
    return { success: false, message: 'เกิดข้อผิดพลาดระบบ' };
  }
}
