/** * ระบบลงทะเบียนและเช็คอินอัตโนมัติ (Version 2.0 - Stable)
 * พัฒนาโดยใช้ Google Apps Script
 */

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Event Check-in System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // อนุญาตให้ Embed iFrame ได้
}

/**
 * ฟังก์ชันทำงานเมื่อมีการส่งฟอร์ม (Form Submit)
 * ต้องตั้งค่า Trigger: "On form submit" ในเมนู Triggers ของ Apps Script
 */
function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = e.range.getRow();
  
  // ดึงค่าจากฟอร์ม (อ้างอิงตามลำดับคอลัมน์ A: Timestamp, B: Name, C: Email)
  const name = e.values[1];
  const email = e.values[2];
  
  // 1. สร้าง Unique ID แบบสุ่มตัวอักษร 6 หลัก + เลขแถว
  const regId = "REG-" + Math.random().toString(36).substr(2, 6).toUpperCase() + "-" + row;
  sheet.getRange(row, 4).setValue(regId); // บันทึกในคอลัมน์ D
  
  // 2. สร้าง QR Code URL ผ่าน QuickChart API (ขนาด 300px)
  const qrUrl = "https://quickchart.io/qr?text=" + encodeURIComponent(regId) + "&size=300&margin=2";
  
  // 3. ออกแบบเนื้อหาอีเมลยืนยัน (HTML)
  const subject = "ยืนยันการลงทะเบียน: " + name;
  const htmlBody = `
    <div style="font-family: 'Kanit', sans-serif; background-color: #f8f9fa; padding: 40px 20px;">
      <div style="max-width: 500px; margin: auto; background: white; border-radius: 20px; overflow: hidden; box-shadow: 0 10px 30px rgba(0,0,0,0.1); border-top: 8px solid #ed008c;">
        <div style="padding: 30px; text-align: center;">
          <h2 style="color: #ed008c; margin: 0;">ลงทะเบียนสำเร็จ</h2>
          <p style="color: #666;">ขอบคุณที่ร่วมเป็นส่วนหนึ่งของงานเรา</p>
          <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">
          <p style="text-align: left;">สวัสดีคุณ <strong>${name}</strong>,</p>
          <p style="text-align: left; color: #444;">นี่คือบัตรเข้าร่วมงานและรหัสยืนยันของคุณ โปรดแสดง QR Code นี้แก่เจ้าหน้าที่ ณ จุดลงทะเบียนเช็คอิน</p>
          
          <div style="background: #fff5f8; border: 2px dashed #ed008c; padding: 20px; border-radius: 15px; margin: 25px 0;">
            <p style="margin: 0; font-size: 12px; color: #ed008c; text-transform: uppercase; letter-spacing: 1px;">Registration ID</p>
            <p style="margin: 5px 0; font-size: 24px; font-weight: bold; color: #333;">${regId}</p>
            <img src="${qrUrl}" width="200" style="margin-top: 15px; border-radius: 10px; border: 1px solid #eee;">
          </div>
          
          <p style="font-size: 12px; color: #999;">*รหัสนี้ใช้สำหรับเช็คอินเข้างานเท่านั้น</p>
        </div>
      </div>
    </div>
  `;
  
  try {
    MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody });
    sheet.getRange(row, 5).setValue("ส่งสำเร็จ"); // บันทึกในคอลัมน์ E
  } catch (err) {
    sheet.getRange(row, 5).setValue("ผิดพลาด: " + err.message);
  }
}

/**
 * ระบบประมวลผลการเช็คอินหน้างาน (เรียกผ่านการสื่อสารจาก Client-side)
 * @param {string} regId รหัสที่ได้จากการสแกน QR Code
 */
function processCheckIn(regId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  // ค้นหารหัสในคอลัมน์ D (index 3)
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] == regId) {
      const name = data[i][1];
      const attendance = data[i][5]; // คอลัมน์ F (index 5)
      
      if (!attendance || attendance === "") {
        // บันทึกเวลาที่เช็คอิน
        const time = Utilities.formatDate(new Date(), "GMT+7", "HH:mm:ss");
        sheet.getRange(i + 1, 6).setValue("Present (" + time + ")");
        return { 
          status: "success", 
          message: "เช็คอินสำเร็จ!", 
          name: name 
        };
      } else {
        // กรณีเคยสแกนไปแล้ว
        return { 
          status: "already", 
          message: "เช็คอินไปแล้ว", 
          name: name 
        };
      }
    }
  }
  
  // กรณีสแกนแล้วไม่เจอ ID ในระบบ
  return { 
    status: "not_found", 
    message: "ไม่พบรหัสในระบบ", 
    name: "Unknown" 
  };
}
