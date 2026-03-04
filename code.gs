/**
 * ฟังก์ชันหลักที่ทำงานเมื่อมีการส่ง Google Form
 */
function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = e.range;
  const row = range.getRow();
  
  // ตรวจสอบและเพิ่มหัวข้อคอลัมน์ใหม่ถ้ายังไม่มี
  const headers = sheet.getRange(1, 1, 1, 7).getValues()[0];
  if (headers[3] !== "Registration ID") {
    sheet.getRange(1, 4).setValue("Registration ID");
    sheet.getRange(1, 5).setValue("Email Status");
    sheet.getRange(1, 6).setValue("Attendance");
    sheet.getRange(1, 7).setValue("QR Code Image");
  }

  // ดึงค่าโดยรองรับทั้งชื่อคอลัมน์ไทยและอังกฤษ และตัดช่องว่าง (trim)
  const name = (e.namedValues['ชื่อ-นามสกุล'] || e.namedValues['Name'] || ["ผู้ร่วมงาน"])[0].trim();
  const email = (e.namedValues['อีเมล'] || e.namedValues['Email'] || e.namedValues['email'] || [""])[0].trim();
  
  const timestamp = new Date().getTime();
  
  // 1. สร้าง Registration ID
  const regId = `REG-${timestamp}-${row}`;
  sheet.getRange(row, 4).setValue(regId);
  
  // 2. สร้าง QR Code URL สำหรับใช้ใน Google Sheet และ Email
  const qrUrl = `https://quickchart.io/qr?text=${encodeURIComponent(regId)}&size=150`;
  
  // 3. ใส่รูป QR Code ลงในเซลล์
  sheet.getRange(row, 7).setFormula(`=IMAGE("${qrUrl}")`);
  
  // 4. ตรวจสอบและส่งอีเมล
  if (email !== "") {
    try {
      sendConfirmationEmail(email, name, regId);
      sheet.getRange(row, 5).setValue("ส่งแล้ว");
    } catch (err) {
      sheet.getRange(row, 5).setValue("ส่งล้มเหลว: " + err.toString());
    }
  } else {
    sheet.getRange(row, 5).setValue("ไม่พบข้อมูลอีเมลในฟอร์ม");
  }
}

/**
 * ฟังก์ชันส่งอีเมล HTML พร้อมรายละเอียดงานและ QR Code
 */
function sendConfirmationEmail(email, name, regId) {
  const qrUrlForEmail = `https://quickchart.io/qr?text=${encodeURIComponent(regId)}&size=300`;
  const eventName = "โครงการพัฒนาศักยภาพบุคลากรเพื่อรองรับการเปลี่ยนแปลงและยกระดับคุณภาพการให้บริการของมหาวิทยาลัย ประจำปี 2569";
  const organizers = "สำนักบริหารการเงิน การบัญชี และการพัสดุ และศูนย์วิเคราะห์รายได้และปฏิบัติการลงทุน";
  const eventDate = "5 - 7 มีนาคม พ.ศ. 2569";
  const eventLocation = "Cape Dara Resort Pattaya จังหวัดชลบุรี";

  const htmlBody = `
    <div style="font-family: 'Kanit', sans-serif; background-color: #f9f9f9; padding: 20px;">
      <div style="max-width: 600px; margin: auto; background-color: #ffffff; border-top: 10px solid #E00084; border-radius: 15px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); overflow: hidden;">
        <div style="padding: 30px; text-align: center;">
          <h2 style="color: #E00084; margin-top: 0;">ยืนยันการลงทะเบียนสำเร็จ</h2>
          <p style="font-size: 16px; color: #333;">สวัสดีคุณ <strong>${name}</strong></p>
          <div style="background-color: #fce4ec; border-radius: 10px; padding: 20px; text-align: left; margin-bottom: 25px; border-left: 5px solid #E00084;">
            <p style="margin: 5px 0; font-weight: bold; color: #E00084;">ชื่องาน:</p>
            <p style="margin: 0 0 10px 0; color: #333;">${eventName}</p>
            <p style="margin: 5px 0; font-weight: bold; color: #E00084;">จัดโดย:</p>
            <p style="margin: 0 0 10px 0; color: #333;">${organizers}</p>
            <p style="margin: 5px 0; font-weight: bold; color: #E00084;">วันที่:</p>
            <p style="margin: 0 0 10px 0; color: #333;">${eventDate}</p>
            <p style="margin: 5px 0; font-weight: bold; color: #E00084;">สถานที่:</p>
            <p style="margin: 0; color: #333;">${eventLocation}</p>
          </div>
          
          <p style="color: #666; font-size: 14px;">กรุณาแสดง QR Code ด้านล่างนี้ต่อเจ้าหน้าที่เพื่อเช็คอินเข้างาน</p>
          <div style="margin: 20px auto; display: inline-block; padding: 10px; background: white; border: 2px solid #E00084; border-radius: 10px;">
            <img src="${qrUrlForEmail}" alt="QR Code" style="width: 250px; height: 250px; display: block;" />
            <p style="margin-top: 10px; font-weight: bold; color: #333;">ID: ${regId}</p>
          </div>
        </div>
        <div style="background-color: #fce4ec; padding: 15px; text-align: center; color: #ad0066; font-size: 12px;">
          ขอบคุณที่ให้ความสนใจเข้าร่วมโครงการกับเรา
        </div>
      </div>
    </div>
  `;
  
  MailApp.sendEmail({
    to: email,
    subject: `[ยืนยันลงทะเบียน] ${eventName}`,
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
      
      if (!attendance || attendance === "") {
        const now = new Date();
        const dateTimeStr = Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy HH:mm:ss");
        sheet.getRange(i + 1, 6).setValue("'เข้าร่วมสัมมนา (" + dateTimeStr + ")");
        return { success: true, message: `เช็คอินสำเร็จ: คุณ ${name}` };
      } else {
        return { success: false, message: `คุณ ${name} ได้เช็คอินไปแล้ว` };
      }
    }
  }
  return { success: false, message: "ไม่พบรหัสลงทะเบียนนี้ในระบบ" };
}
