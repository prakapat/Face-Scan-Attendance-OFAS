/**
 * แก้ไขส่วนนี้: เพิ่มฟังก์ชัน doPost เพื่อรับค่าจากภายนอก (Netlify)
 */
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const registrationId = params.registrationId;
    
    // เรียกใช้ฟังก์ชันประมวลผลเดิม
    const result = processCheckIn(registrationId);
    
    // ส่งผลลัพธ์กลับเป็น JSON
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Error: " + err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ฟังก์ชันประมวลผลการเช็คอิน (เหมือนเดิม)
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
        const timeStr = Utilities.formatDate(now, "GMT+7", "HH:mm:ss");
        sheet.getRange(i + 1, 6).setValue(`Present (${timeStr})`);
        return { success: true, message: `เช็คอินสำเร็จ: คุณ ${name}` };
      } else {
        return { success: false, message: `คุณ ${name} ได้เช็คอินไปแล้ว` };
      }
    }
  }
  return { success: false, message: "ไม่พบรหัสลงทะเบียนนี้ในระบบ" };
}

// ฟังก์ชัน onFormSubmit และ sendConfirmationEmail ยังคงเดิมตามที่เคยเขียนไว้
