/**
 * ฟังก์ชันหลักที่ทำงานเมื่อมีการส่ง Google Form
 */
function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = e.range;
  const row = range.getRow();
  
  // ตรวจสอบและเพิ่มหัวข้อคอลัมน์ใหม่ถ้ายังไม่มี (เพิ่มคอลัมน์ QR Code Image)
  const headers = sheet.getRange(1, 1, 1, 7).getValues()[0];
  if (headers[3] !== "Registration ID") {
    sheet.getRange(1, 4).setValue("Registration ID");
    sheet.getRange(1, 5).setValue("Email Status");
    sheet.getRange(1, 6).setValue("Attendance");
    sheet.getRange(1, 7).setValue("QR Code Image");
  }

  const name = e.namedValues['ชื่อ-นามสกุล'] ? e.namedValues['ชื่อ-นามสกุล'][0] : "ผู้ร่วมงาน";
  const email = e.namedValues['อีเมล'] ? e.namedValues['อีเมล'][0] : "";
  const timestamp = new Date().getTime();
  
  // 1. สร้าง Registration ID
  const regId = `REG-${timestamp}-${row}`;
  sheet.getRange(row, 4).setValue(regId);
  
  // 2. สร้าง QR Code URL สำหรับใช้ใน Google Sheet และ Email
  const qrUrl = `https://quickchart.io/qr?text=${encodeURIComponent(regId)}&size=150`;
  
  // 3. ใส่รูป QR Code ลงในเซลล์โดยใช้สูตร =IMAGE()
  // คอลัมน์ที่ 7 คือ G
  sheet.getRange(row, 7).setFormula(`=IMAGE("${qrUrl}")`);
  
  // 4. ส่งอีเมลพร้อม QR Code
  if (email) {
    try {
      sendConfirmationEmail(email, name, regId);
      sheet.getRange(row, 5).setValue("ส่งแล้ว");
    } catch (err) {
      sheet.getRange(row, 5).setValue("ส่งล้มเหลว: " + err.toString());
    }
  }
}

/**
 * ฟังก์ชันส่งอีเมล HTML พร้อม QR Code
 */
function sendConfirmationEmail(email, name, regId) {
  const qrUrlForEmail = `https://quickchart.io/qr?text=${encodeURIComponent(regId)}&size=300`;
  
  const htmlBody = `
    <div style="font-family: 'Kanit', sans-serif; border: 2px solid #E00084; padding: 20px; border-radius: 15px; max-width: 500px; margin: auto; text-align: center;">
      <h2 style="color: #E00084;">ยืนยันการลงทะเบียนสำเร็จ</h2>
      <p>สวัสดีคุณ <strong>${name}</strong></p>
      <p>ขอบคุณที่ลงทะเบียนร่วมงานกับเรา นี่คือรหัสและ QR Code สำหรับการเช็คอินหน้างาน:</p>
      <div style="background-color: #fce4ec; padding: 15px; border-radius: 10px; margin: 20px 0;">
        <h3 style="margin: 0; color: #333;">ID: ${regId}</h3>
      </div>
      <img src="${qrUrlForEmail}" alt="QR Code" style="width: 200px; height: 200px; border: 5px solid #fff; box-shadow: 0 0 10px rgba(0,0,0,0.1);" />
      <p style="font-size: 0.9em; color: #666; margin-top: 20px;">กรุณาแสดง QR Code นี้แก่เจ้าหน้าที่ ณ จุดลงทะเบียน</p>
    </div>
  `;
  
  MailApp.sendEmail({
    to: email,
    subject: `ยืนยันการลงทะเบียนงานอีเวนต์ - ${name}`,
    htmlBody: htmlBody
  });
}

/**
 * Web App Entry Point สำหรับรับข้อมูลจาก Netlify
 */
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const registrationId = params.registrationId;
    const result = processCheckIn(registrationId);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Error: " + err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ฟังก์ชันประมวลผลการเช็คอิน
 */
function processCheckIn(registrationId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === registrationId) {
      const name = data[i][1];
      const attendance = data[i][5];
      
      if (!attendance) {
        const now = new Date();
        // ปรับรูปแบบเป็น วัน/เดือน/ปี ชั่วโมง:นาที:วินาที
        const dateTimeStr = Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy HH:mm:ss");
        // เปลี่ยนข้อความจาก Present เป็น เข้าร่วมสัมมนา
        sheet.getRange(i + 1, 6).setValue(`เข้าร่วมสัมมนา (${dateTimeStr})`);
        return { success: true, message: `เช็คอินสำเร็จ: คุณ ${name}` };
      } else {
        return { success: false, message: `คุณ ${name} ได้เช็คอินไปแล้ว` };
      }
    }
  }
  return { success: false, message: "ไม่พบรหัสลงทะเบียนนี้ในระบบ" };
}
