// ================================
// ค่าคงที่สำหรับ Google Sheet, Calendar และ LINE (Flex message)
const SPREADSHEET_ID = '1DzQaGJRZNcv0I_ieFmdd9_-ltngtm9Fo2f7SwqQ3igQ';
const CALENDAR_ID = '851f56a6d18a7aa852c930872af57c3f7fa24734f7c948cf38ecae7cc09d3df6@group.calendar.google.com';
const LINE_TOKEN = '48lb16bl7sn8aGEiNkaIgFWmkgU4EsHauKAikCZXPaqg/t6bJto1pll2DdRJouTOPbJPRGZA5snAlwQtEGMUiNva1f1agAasgEf+QWK7xRakz+F64nqdwEsHdOZlZwDlP4mFFXLwXop18tv0dEZUhAdB04t89/1O/w1cDnyilFU=';

const LIFF_ID_CONFIRM = '2007432636-Rw2NO5DL';  

// ================================
// ค่าคงที่สำหรับ Telegram Bot
const TELEGRAM_BOT_TOKEN = 'XXXXXXX';  // Bot Token ของคุณ
const TELEGRAM_CHAT_ID = 'YYYYYYY';  // Chat ID

const SHEET_DATETIME = 'ตั้งค่าทั่วไป';   // ชีตเวลานัดหมาย
const SHEET_BOOKING = 'บันทึกนัดหมาย';  // ชีตการจอง
const SHEET_MEMBER = 'รายชื่อสมาชิก';   // ชีตสมาชิก
const SHEET_SERVICES = 'บริการ';         // ชีตข้อมูลบริการ

// ----------------------------------------------------------------
// ส่งแจ้งเตือนผ่าน Telegram
function sendTelegramNotify(message) {
  const url = 'https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage';
  const payload = { chat_id: TELEGRAM_CHAT_ID, text: message };
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload) };
  try {
    const response = UrlFetchApp.fetch(url, options);
    Logger.log('Telegram response: ' + response.getContentText());
  } catch (error) {
    Logger.log('Telegram error: ' + error.toString());
  }
}

// ----------------------------------------------------------------
// ดึงรายการบริการ
function fetchServices() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_SERVICES);
  const data = sheet.getDataRange().getValues().slice(1);
  const services = data.map(row => ({
    id: row[0],
    name: row[1],
    details: row[2],
    price: row[3]
  }));
  return ContentService.createTextOutput(JSON.stringify(services))
    .setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------
// ดึงรายละเอียดบริการตาม id
function getServiceById(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_SERVICES);
  const data = sheet.getDataRange().getValues().slice(1);
  for (let r of data) {
    if (r[0].toString() === id.toString()) {
      return { id: r[0], name: r[1], details: r[2], price: r[3] };
    }
  }
  return null;
}

// ----------------------------------------------------------------
// ดึง services + price จาก SHEET_DATETIME (เดิม)
function fetchServicesAndprice() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_DATETIME);
  const data = sheet.getDataRange().getValues().slice(1);
  const services = [...new Set(data.map(r => r[4]).filter(Boolean))];
  const price = [...new Set(data.map(r => r[5]).filter(Boolean))];
  const maxBookings = data.reduce((acc, row) => {
    if (row[0] && row[1]) acc[row[0]] = row[1];
    return acc;
  }, {});
  return ContentService.createTextOutput(JSON.stringify({ services, price, maxBookings }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------
// Routing: doGet
function doGet(e) {
  const action = e.parameter.action;
  switch (action) {
    case 'fetchServices':
      return fetchServices();
    case 'fetchServicesAndprice':
      return fetchServicesAndprice();
    case 'fetchBookings':
      return fetchBookings();
    default:
      return checkAvailability(e);
  }
}


// ----------------------------------------------------------------
// Routing: doPost
function doPost(e) {
  const action = e.parameter.action;
  if (action === 'makeBooking') {
    return makeBooking(e);
  } else if (action === 'cancelBooking') {
    return cancelBooking(e.parameter);
  } else if (action === 'confirmBooking') {
    return confirmBooking(e.parameter);
  } else if (action === 'completeBooking') {
    return completeBooking(e);
  } else if (action === 'sendReminder') {
    return sendReminder(e.parameter.id);
  } else if (action === 'sendReminderWithPayment') {
    return sendReminderWithPayment(e.parameter);
  } else if (action === 'sendReminderWithoutPayment') {
    return sendReminderWithoutPayment(e.parameter);
  }
  // รองรับ LINE postback JSON
  if (e.postData && e.postData.type === 'application/json') {
    const data = JSON.parse(e.postData.contents);
    if (data.events) {
      data.events.forEach(evt => {
        if (evt.type === 'postback') {
          const params = parseQueryString(evt.postback.data);
          if (params.action === 'confirmBooking') return confirmBooking(params);
          if (params.action === 'cancelBooking') return cancelBooking(params);
        }
      });
    }
  }
  return ContentService.createTextOutput('Invalid action');
}

// แปลง query string เป็น object
function parseQueryString(qs) {
  return qs.split('&').reduce((o, p) => {
    const [k, v] = p.split('=');
    o[k] = decodeURIComponent(v);
    return o;
  }, {});
}

// ----------------------------------------------------------------
// ตรวจสอบความพร้อมจอง & วันหยุด (เวอร์ชันปรับปรุง + เพิ่มความเสถียร)
function checkAvailability(e) {
  if (!e.parameter.date && !e.parameter.getHolidays) {
    return ContentService.createTextOutput(JSON.stringify({ error: "Missing params" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_DATETIME);
  const rows = sheet.getDataRange().getValues();

  // --- ส่วนที่ปรับปรุง: เพิ่มการตรวจสอบว่ามีข้อมูลในแถวที่ 2 หรือไม่ ---
  // ป้องกัน Error หากในชีตมีแค่ Header
  const permanentHolidaysSetting = (rows && rows.length > 1 && rows[1][4]) ? String(rows[1][4]) : '';
  const permanentHolidayDays = permanentHolidaysSetting.split(',').map(d => parseInt(d.trim())).filter(d => !isNaN(d));
  // --- สิ้นสุดส่วนที่ปรับปรุง ---

  const dataRows = rows.slice(1);

  if (e.parameter.getHolidays) {
    const holidays = dataRows.map(r => {
      const v = r[3];
      if (!v) return null;
      if (typeof v === 'string' && /^\d{2}-\d{2}-\d{4}$/.test(v)) {
        const [d, m, y] = v.split('-');
        return `${y}-${m}-${d}`;
      }
      if (v instanceof Date) {
        return Utilities.formatDate(v, "Asia/Bangkok", "yyyy-MM-dd");
      }
      return null;
    }).filter(Boolean);

    return ContentService.createTextOutput(JSON.stringify({
      holidays: Array.from(new Set(holidays)),
      permanentHolidays: permanentHolidayDays
    }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const inputDate = new Date(e.parameter.date);
  if (isNaN(inputDate)) {
    return ContentService.createTextOutput(JSON.stringify({ error: "Invalid date" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const times = dataRows.map(r => {
    const t = r[1];
    if (!t) return null;
    if (t instanceof Date) return Utilities.formatDate(t, "Asia/Bangkok", "HH:mm");
    if (typeof t === 'string' && /^\d{2}:\d{2}:\d{2}$/.test(t)) return t.substring(0, 5);
    return t.toString();
  }).filter(Boolean);
  const maxBookings = dataRows.map(r => Number(r[2]) || 0);

  const holidays = dataRows.map(r => {
    const v = r[3];
    if (!v) return null;
    if (typeof v === 'string' && /^\d{2}-\d{2}-\d{4}$/.test(v)) {
      const [d, m, y] = v.split('-');
      return new Date(`${y}-${m}-${d}`);
    }
    if (v instanceof Date) return v;
    return null;
  }).filter(Boolean);

  const formattedInputDate = Utilities.formatDate(inputDate, "Asia/Bangkok", "yyyy-MM-dd");

  const isSpecificHoliday = holidays.some(h => Utilities.formatDate(h, "Asia/Bangkok", "yyyy-MM-dd") === formattedInputDate);
  const isPermanentHoliday = permanentHolidayDays.includes(inputDate.getDay());
  const isHoliday = isSpecificHoliday || isPermanentHoliday;

  let availability = {};
  if (inputDate < today) {
    times.forEach(t => availability[t] = { status: 'เต็ม', remaining: 0, total: 0 });
  } else {
    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    times.forEach((t, i) => {
      let status = 'Available', remaining = maxBookings[i], total = maxBookings[i];
      if (isHoliday) {
        status = 'ไม่ว่าง';
        remaining = 0;
      } else {
        const [hh, mm] = t.split(':').map(Number);
        const start = new Date(inputDate); start.setHours(hh, mm);
        const end = new Date(start.getTime() + 30 * 60 * 1000);
        const events = calendar.getEvents(start, end).filter(ev => ev.getTitle().indexOf('นัดคิว:') > -1);
        remaining = total - events.length;
        if (remaining <= 0) { remaining = 0; status = 'เต็ม'; }
      }
      availability[t] = { status, remaining, total };
    });
  }

  return ContentService.createTextOutput(JSON.stringify({
    availability,
    isHoliday,
    color: isHoliday ? 'red' : 'default'
  })).setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------
// สร้างการจอง
function makeBooking(e) {
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_BOOKING);
  const dtSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_DATETIME);
  const memSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_MEMBER);

  const userlineid = e.parameter.userlineid;
  const name = e.parameter.name;
  let phonenumber = e.parameter.phonenumber;
   // เพิ่มการดึงค่าใหม่จาก parameter
  var nationalId = e.parameter.national_id;
  var socialSecurityNumber = e.parameter.social_security_number;

  const note = e.parameter.note;
  const timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd-MM-yyyy HH:mm:ss");
  const date = e.parameter.date;
  const time = e.parameter.time;
  const datetime = new Date(`${date}T${time}:00`);

  // ✅ เพิ่ม ' เพื่อป้องกันการตัด 0 ออก
  phonenumber = "'" + String(phonenumber).trim();

  // ตรวจ capacity
  const times = dtSheet.getRange('B2:B').getValues().flat();
  const idx = times.indexOf(time);
  if (idx >= 0) {
    const maxBookingsData = dtSheet.getDataRange().getValues().slice(1).map(r => Number(r[2]) || 0);
    const maxBk = maxBookingsData[idx];
    const events = cal.getEvents(datetime, new Date(datetime.getTime() + 30 * 60 * 1000));
    if (events.length >= maxBk) {
      return ContentService.createTextOutput('ช่วงเวลาที่เลือกเต็มแล้ว');
    }
  }

  const formattedDate = Utilities.formatDate(datetime, Session.getScriptTimeZone(), 'dd-MM-yyyy');

  // ดึงบริการ
  let serviceNames = "", totalPrice = 0;
  try {
    const arr = JSON.parse(e.parameter.selectedServices || '[]');
    const details = arr.map(id => getServiceById(id)).filter(s => s);
    serviceNames = details.map(s => `${s.name} (ราคา ${s.price})`).join(", ");
    totalPrice = details.reduce((sum, s) => sum + Number(s.price), 0);
  } catch {
    serviceNames = e.parameter.selectedServices;
  }

  // สร้าง Event
  const event = cal.createEvent(`นัดคิว: ${serviceNames}`, datetime, new Date(datetime.getTime() + 30 * 60 * 1000), {
    description: `ลูกค้า: ${name}\nเบอร์: ${phonenumber}\nหมายเหตุ: ${note}`
  });
  const calendarEventId = event.getId();

  const nextRow = sheet.getLastRow() + 1;
  const id = 'BK' + (nextRow - 1);
  const newRow = [
    id,               // A
    userlineid,       // B
    name,             // C
    phonenumber,      // D  
    nationalId,                // <-- ข้อมูลใหม่
    socialSecurityNumber,      // <-- ข้อมูลใหม่
    note,             // E
    serviceNames,     // F
    totalPrice,       // G
    formattedDate,    // H
    time,             // I
    'รอการยืนยัน',     // J
    timestamp,        // K
    calendarEventId   // L
  ];
  sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow]);

  // กำหนดฟอร์แมทเป็นข้อความ
  sheet.getRange(nextRow, 4).setNumberFormat("@"); // Phone
  sheet.getRange(nextRow, 8).setNumberFormat("@"); // Date
  sheet.getRange(nextRow, 9).setNumberFormat("@"); // Time
  sheet.getRange(nextRow, 11).setNumberFormat("@"); // Timestamp
  sheet.getRange(nextRow, 12).setNumberFormat("@"); // Calendar Event ID

  // อัปเดตหรือเพิ่มสมาชิก
  const dataMem = memSheet.getDataRange().getValues();
  const existIdx = dataMem.findIndex(r => r[1] === userlineid) + 1;
  if (existIdx > 1) {
    memSheet.getRange(existIdx, 3).setValue(name);
    memSheet.getRange(existIdx, 4).setNumberFormat("@").setValue(phonenumber);
    memSheet.getRange(existIdx, 5).setValue(nationalId);
    memSheet.getRange(existIdx, 6).setValue(socialSecurityNumber);
    memSheet.getRange(existIdx, 7).setValue(timestamp);
  } else {
    memSheet.appendRow([id, userlineid, name, phonenumber, timestamp]);
    memSheet.getRange(memSheet.getLastRow(), 4).setNumberFormat("@");
  }

  // แจ้ง Telegram
  sendTelegramNotify(
    `📅 นัดหมายใหม่\n👤: ${name}\n📋: ${serviceNames}\n💰: ${totalPrice}\n📅: ${formattedDate}\n⏰: ${time}`
  );

  return ContentService.createTextOutput('Booking successful!');
}


// ----------------------------------------------------------------
// แจ้งเตือนพร้อมจ่ายเงิน
function sendReminderWithPayment(params) {
  const id = params.id;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_BOOKING);
  const data = sheet.getDataRange().getValues();
  const idx = data.findIndex(r => r[0] == id);
  if (idx < 0) return ContentService.createTextOutput('Booking not found');
  const b = data[idx];
  const totalPrice = b[6];
  const qrCodeUrl = `https://promptpay.io/0623733306/${totalPrice}.png`;
  const liffUrl = `https://liff.line.me/2006029649-EbKnbZJ0?id=${id}&userlineid=${b[1]}&name=${encodeURIComponent(b[2])}&contact=${encodeURIComponent(b[3])}&serviceNames=${encodeURIComponent(b[5])}&totalPrice=${totalPrice}`;
  const reminderData = {
    idKey: b[0], service: b[5], price: totalPrice,
    date: b[7], time: b[8], name: b[2],
    qrCodeUrl, liffUrl
  };
  sendLineMessage(b[1], 'reminder_with_payment', reminderData);
  return ContentService.createTextOutput('Reminder with payment sent.');
}

// ----------------------------------------------------------------
// แจ้งเตือน without payment
function sendReminderWithoutPayment(params) {
  const id = params.id;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_BOOKING);
  const data = sheet.getDataRange().getValues();
  const idx = data.findIndex(r => r[0] == id);
  if (idx < 0) return ContentService.createTextOutput('Booking not found');
  const b = data[idx];
  const reminderData = {
    idKey: b[0], service: b[5], price: b[6],
    date: b[7], time: b[8], name: b[2]
  };
  sendLineMessage(b[1], 'reminder_no_payment', reminderData);
  return ContentService.createTextOutput('Reminder without payment sent.');
}

// ----------------------------------------------------------------
// ยืนยันการจอง
function confirmBooking(params) {
  const idKey = params.id;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_BOOKING);
  const data = sheet.getDataRange().getValues();
  const idx = data.findIndex(r => r[0] == idKey);
  if (idx < 0) return ContentService.createTextOutput('ไม่พบการนัดหมายนี้.');
  const booking = data[idx];
  const current = booking[9];
  if (current !== 'รอการยืนยัน') {
    const msg = `ไม่สามารถยืนยันได้ สถานะปัจจุบันคือ "${current}"`;
    sendLineMessage(booking[1], 'notification', { message: msg });
    return ContentService.createTextOutput(msg);
  }
  sheet.getRange(idx + 1, 10).setValue('ยืนยันแล้ว');
  sendLineMessage(booking[1], 'confirmBooking', {
    service: booking[5], price: booking[6],
    date: booking[7], time: booking[8],
    name: booking[2], status: 'ยืนยันแล้ว', idKey
  });
  return ContentService.createTextOutput('Booking confirmed successfully.');
}

// ----------------------------------------------------------------
// เสร็จสิ้นการจอง
function completeBooking(e) {
  const idKey = e.parameter.id;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_BOOKING);
  const data = sheet.getDataRange().getValues();
  const idx = data.findIndex(r => r[0] == idKey);
  if (idx < 0) return ContentService.createTextOutput('Booking not found.');
  sheet.getRange(idx + 1, 10).setValue('เสร็จสิ้น');
  const booking = data[idx];
  sendLineMessage(booking[1], 'completion', {
    service: booking[5], price: booking[6],
    date: booking[7], time: booking[8],
    name: booking[2], status: 'เสร็จสิ้น'
  });
  return ContentService.createTextOutput('Booking completed successfully.');
}

// ----------------------------------------------------------------
// ยกเลิกการจอง
function cancelBooking(params) {
  const idKey = params.id;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_BOOKING);
  const data = sheet.getDataRange().getValues();
  const idx = data.findIndex(r => r[0] == idKey);
  if (idx < 0) return ContentService.createTextOutput('ไม่พบการนัดหมายนี้.');
  const booking = data[idx];
  const current = booking[9]; // Column J: Status
  if (current !== 'รอการยืนยัน') {
    const msg = `ไม่สามารถยกเลิกได้ สถานะปัจจุบันคือ "${current}"`;
    sendLineMessage(booking[1], 'notification', { message: msg });
    return ContentService.createTextOutput(msg);
  }
  sheet.getRange(idx + 1, 10).setValue('ยกเลิก'); // Update status in sheet (Column J)

  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  try {
    const calendarEventId = booking[11]; // <<< ดึง Calendar Event ID จากคอลัมน์ใหม่ (Index 11 คือคอลัมน์ L)
    if (calendarEventId) { // ตรวจสอบว่ามี Event ID อยู่จริง
      const ev = cal.getEventById(calendarEventId);
      if (ev) {
        ev.deleteEvent();
        Logger.log(`Event ID ${calendarEventId} deleted successfully.`);
      } else {
        Logger.log(`Calendar event with ID ${calendarEventId} not found.`);
      }
    } else {
      Logger.log(`No Calendar Event ID found for booking ${idKey}.`);
    }
  } catch (e) {
    Logger.log(`Error deleting calendar event for booking ${idKey}: ${e.toString()}`); // <<< เพิ่ม Logger เพื่อดู error
  }

  sendLineMessage(booking[1], 'cancellation', {
    service: booking[5], price: booking[6],
    date: booking[7], time: booking[8],
    name: booking[2], status: 'ยกเลิก'
  });
  return ContentService.createTextOutput('Booking cancelled successfully.');
}
// ----------------------------------------------------------------
// ส่ง reminder ทันที
function sendReminder(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_BOOKING);
  const data = sheet.getDataRange().getValues();
  const idx = data.findIndex(r => r[0] == id);
  if (idx < 0) return ContentService.createTextOutput('Booking not found.');
  const b = data[idx];
  sendLineMessage(b[1], 'reminder', {
    idKey: b[0], service: b[5], price: b[6],
    date: b[7], time: b[8], name: b[2], status: b[9]
  });
  return ContentService.createTextOutput('Reminder sent successfully.');
}

// ----------------------------------------------------------------
// --- วางทับฟังก์ชันนี้ทั้งหมด ---
function sendReminders() {
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_BOOKING);
  const data = sheet.getDataRange().getValues().slice(1);

  data.forEach(row => {
    try {
      const dateStr = row[7]; // วันที่รูปแบบ "DD-MM-YYYY"
      const status = row[9];

      if (!dateStr || status !== 'รอการยืนยัน') {
        return; // ข้ามแถวนี้ไปถ้าไม่มีวันที่ หรือสถานะไม่ใช่ 'รอการยืนยัน'
      }

      // *** จุดที่แก้ไข: เปลี่ยนวิธีแปลงวันที่ให้ปลอดภัยและแน่นอน ***
      const dateParts = dateStr.split('-'); // ["DD", "MM", "YYYY"]
      if (dateParts.length !== 3) return; // ข้ามถ้า format ผิด

      // new Date(year, monthIndex, day) // เดือนใน JS เริ่มจาก 0 (ม.ค.=0)
      const bookingDate = new Date(parseInt(dateParts[2]), parseInt(dateParts[1]) - 1, parseInt(dateParts[0]));

      // เปรียบเทียบวันที่
      if (bookingDate.getDate() === tomorrow.getDate() &&
        bookingDate.getMonth() === tomorrow.getMonth() &&
        bookingDate.getFullYear() === tomorrow.getFullYear()) {

        sendLineMessage(row[1], 'reminder', {
          idKey: row[0],
          service: row[5],
          price: row[6],
          date: row[7],
          time: row[8],
          name: row[2],
          status: 'รอการยืนยัน'
        });
      }
    } catch (e) {
      Logger.log(`Error processing row for sendReminders: ${e.toString()}. Row data: ${row}`);
    }
  });
}

// สร้าง Trigger สำหรับ sendReminders
function createReminderTrigger() {
  ScriptApp.newTrigger('sendReminders')
    .timeBased().everyDays(1).atHour(12).create();
}

// ==========================================================
// ========== ฟังก์ชันสำหรับทดสอบ (DEMO) ==========
// ==========================================================
function testReminderFlex() {
  // 1. จำลองข้อมูลการจอง (เหมือนข้อมูลที่ดึงมาจากชีต)
  // คุณสามารถลองเปลี่ยนค่าพวกนี้เพื่อทดสอบได้
  const mockBookingData = {
    idKey: 'BK-DEMO-001',
    name: 'ลูกค้าทดสอบ',
    service: 'ตัดผม, ทำสีผม',
    price: '1200',
    date: '15-08-2025',
    time: '13:00',
    status: 'รอการยืนยัน'
  };

  try {
    // 2. เรียกใช้ฟังก์ชัน createFlexMessage โดยตรงด้วยข้อมูลจำลอง
    // และระบุ type เป็น 'reminder'
    const flexMessageJson = createFlexMessage('reminder', mockBookingData);

    // 3. แสดงผลลัพธ์ใน Log เพื่อให้เรานำไปตรวจสอบ
    Logger.log("========== DEMO START ==========");
    Logger.log("--- INPUT DATA ---");
    Logger.log(JSON.stringify(mockBookingData, null, 2)); // แสดงข้อมูลจำลองที่ใช้
    Logger.log("--- OUTPUT FLEX JSON ---");
    Logger.log(JSON.stringify(flexMessageJson, null, 2)); // แสดงผลลัพธ์ Flex Message ที่สร้างได้
    Logger.log("========== DEMO END ==========");

  } catch (e) {
    Logger.log("!!! AN ERROR OCCURRED DURING FLEX CREATION !!!");
    Logger.log(e.toString());
  }
}
// ----------------------------------------------------------------
// ส่งข้อความผ่าน LINE
function sendLineMessage(userId, type, data) {
  const url = 'https://api.line.me/v2/bot/message/push';
  const msg = createFlexMessage(type, data);
  const payload = { to: userId, messages: [msg] };
  const opts = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + LINE_TOKEN },
    payload: JSON.stringify(payload)
  };
  try {
    const resp = UrlFetchApp.fetch(url, opts);
    Logger.log('LINE push response: ' + resp.getResponseCode());
  } catch (err) {
    Logger.log('LINE error: ' + err);
  }
}
// --- วางทับฟังก์ชันนี้ทั้งหมดใน AP V2.js ---
function createFlexMessage(type, data) {
  if (type === 'notification') {
    return {
      type: 'text',
      text: data.message
    };
  }

  const serviceText = data.service || 'ไม่มีบริการระบุ';
  const priceText = (data.price !== null && data.price !== undefined) ? data.price.toString() : '0';
  const nameText = data.name || 'ไม่มีชื่อ';
  const dateText = data.date || 'ไม่มีวันที่';
  const timeText = data.time || 'ไม่มีเวลา';
  const idKeyText = data.idKey || '';

  const headerText = {
    confirmBooking: '✅ ยืนยันการนัดหมาย',
    cancellation: '⚠️ ยกเลิกการนัดหมาย',
    completion: '✅ นัดหมายเสร็จสิ้น',
    reminder: '🔔 กรุณาตรวจสอบนัดหมาย',
    confirmation: '🔔 การนัดหมาย'
  }[type] || '🔔 การนัดหมาย';

  const statusText = {
    confirmBooking: 'ยืนยันแล้ว',
    cancellation: 'ยกเลิก',
    completion: 'เสร็จสิ้น',
    reminder: 'รอการยืนยัน',
    confirmation: 'รอการยืนยัน'
  }[type] || 'รอการยืนยัน';

  const statusColor = {
    confirmBooking: '#06C755',
    cancellation: '#FF0000',
    completion: '#06C755',
    reminder: '#FFA500',
    confirmation: '#1f1f97'
  }[type] || '#2F1A87';

  const bodyContents = [{
    type: 'text',
    text: headerText,
    weight: 'bold',
    color: statusColor,
    size: 'sm'
  }, {
    type: 'separator',
    margin: 'md'
  }, {
    type: 'box',
    layout: 'vertical',
    margin: 'md',
    spacing: 'sm',
    contents: [{
      type: 'box',
      layout: 'horizontal',
      contents: [{
        type: 'text',
        text: 'ชื่อผู้จอง',
        size: 'sm',
        color: '#2F1A87',
        weight: 'bold'
      }, {
        type: 'text',
        text: nameText,
        size: 'sm',
        color: '#111111',
        align: 'end'
      }]
    }, {
      type: 'box',
      layout: 'horizontal',
      contents: [{
        type: 'text',
        text: 'บริการ',
        size: 'sm',
        color: '#2F1A87',
        weight: 'bold'
      }]
    }, {
      type: 'box',
      layout: 'vertical',
      margin: 'sm',
      contents: serviceText.split(',').map(item => ({
        type: 'text',
        text: item.trim(),
        size: 'sm',
        color: '#111111',
        wrap: true
      }))
    }, {
      type: 'box',
      layout: 'horizontal',
      contents: [{
        type: 'text',
        text: 'ราคา',
        size: 'sm',
        color: '#2F1A87',
        weight: 'bold'
      }, {
        type: 'text',
        text: priceText,
        size: 'sm',
        color: '#111111',
        align: 'end'
      }]
    }, {
      type: 'box',
      layout: 'horizontal',
      contents: [{
        type: 'text',
        text: 'วันที่นัดหมาย',
        size: 'sm',
        color: '#2F1A87',
        weight: 'bold'
      }, {
        type: 'text',
        text: dateText,
        size: 'sm',
        color: '#111111',
        align: 'end'
      }]
    }, {
      type: 'box',
      layout: 'horizontal',
      contents: [{
        type: 'text',
        text: 'เวลา',
        size: 'sm',
        color: '#2F1A87',
        weight: 'bold'
      }, {
        type: 'text',
        text: timeText,
        size: 'sm',
        color: '#111111',
        align: 'end'
      }]
    }, {
      type: 'box',
      layout: 'horizontal',
      contents: [{
        type: 'text',
        text: 'สถานะ',
        size: 'sm',
        color: '#2F1A87',
        weight: 'bold'
      }, {
        type: 'text',
        text: statusText,
        size: 'sm',
        color: statusColor,
        align: 'end'
      }]
    }]
  }];

  const footer = [];
  let footerBg = '#FFFFFF';

  if (type === 'reminder') {
    const params = `id=${encodeURIComponent(idKeyText)}&name=${encodeURIComponent(nameText)}&service=${encodeURIComponent(serviceText)}&price=${encodeURIComponent(priceText)}&date=${encodeURIComponent(dateText)}&time=${encodeURIComponent(timeText)}`;
    const liffUrl = `https://liff.line.me/${LIFF_ID_CONFIRM}?${params}`;

    footer.push({
      type: 'box',
      layout: 'vertical',
      spacing: 'sm',
      contents: [{
        type: 'button',
        action: {
          type: 'uri',
          label: 'จัดการนัดหมาย',
          uri: liffUrl
        },
        style: 'primary',
        height: 'sm',
        color: '#1f1f97'
      }]
    });
    footerBg = '#F3F2FA';

  } else if (type === 'cancellation') {
    footer.push({
      type: 'box',
      layout: 'vertical',
      contents: [{
        type: 'text',
        text: 'การนัดหมายถูกยกเลิกแล้ว',
        size: 'xs',
        color: '#ffffff',
        align: 'center'
      }]
    });
    footerBg = '#EC726E';

  } else if (type === 'completion') {
    footer.push({
      type: 'box',
      layout: 'vertical',
      contents: [{
        type: 'text',
        text: 'ขอบคุณที่ใช้บริการ',
        size: 'xs',
        color: '#ffffff',
        align: 'center'
      }]
    });
    footerBg = '#03B555';

  } else if (type === 'confirmation') {
    footer.push({
      type: 'box',
      layout: 'vertical',
      contents: [{
        type: 'text',
        text: 'บันทึกนัดหมายแล้ว',
        size: 'xs',
        color: '#ffffff',
        align: 'center'
      }]
    });
    footerBg = '#1f1f97';

  } else if (type === 'confirmBooking') {
    footer.push({
      type: 'box',
      layout: 'vertical',
      contents: [{
        type: 'text',
        text: 'กรุณามาก่อน 10-20 นาที',
        size: 'xs',
        color: '#ffffff',
        align: 'center'
      }]
    });
    footerBg = '#06c755';
  }

  if (data.qrCodeUrl && data.liffUrl) {
    bodyContents.push({
      type: 'box',
      layout: 'vertical',
      margin: 'md',
      spacing: 'sm',
      contents: [{
        type: 'image',
        url: data.qrCodeUrl,
        size: 'full',
        aspectMode: 'cover',
        aspectRatio: '1:1'
      }, {
        type: 'button',
        action: {
          type: 'uri',
          label: 'ชำระเงิน',
          uri: data.liffUrl
        },
        style: 'primary',
        color: '#2F1A87',
        margin: 'md'
      }]
    });
  }

  return {
    type: 'flex',
    altText: 'กำหนดการนัดหมาย',
    contents: {
      type: 'bubble',
      body: {
        type: 'box',
        layout: 'vertical',
        contents: bodyContents
      },
      footer: {
        type: 'box',
        layout: 'vertical',
        contents: footer,
        backgroundColor: footerBg
      }
    }
  };
}
// ----------------------------------------------------------------
// ดึงรายการจอง
function fetchBookings() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_BOOKING);
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    result.push({
      idKey: r[0], userlineid: r[1], name: r[2],
      phonenumber: r[3], national_id: r[4], social_security_number: r[5], note: r[6], service: r[7],
      price: r[8], date: r[9], time: r[10],
      status: r[11], timestamp: r[12]
    });
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------
// Test functions
function testConfirmBooking() {
  const resp = confirmBooking({ id: 2 });
  Logger.log(resp.getContent());
}
function testCancelBooking() {
  const resp = cancelBooking({ id: 1 });
  Logger.log(resp.getContent());
}
function testCompleteBooking() {
  const resp = completeBooking({ parameter: { id: 1 } });
  Logger.log(resp.getContent());
}
