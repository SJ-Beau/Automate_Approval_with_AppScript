/**
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e
 */
function processFormSubmissionToDoc(e) {
  try {
    // 1. กำหนด ID ของ Google Doc Template ทั้งหมดที่คุณมี
    //    สำคัญ: 'Key' ต้องตรงกับ 'ชื่อตัวเลือก' ใน Google Form ของคุณ
    const templateIds = {
      'ขอโอนย้ายไปยังบริการอื่นๆ': "1LbINjjA25MOIGbapEKkkBwIzcRJVngWyPCJ1BEAhNNU",
      'ขอขยายวันบริการ 30 วัน': "1kyEha3cT_bFkkiGIQIo22tuzhKLXBZU_YIpuVFrLxDo",
      'ขอเปิดงานเร่งด่วน / นอกเวลาทำการ': "1AkmJzqq3Lgsg0x3gkM7sezTjpttVmuDmDNx02RVw0SU",
      'ขอขึ้นออนไลน์เว็บไซต์โดยไม่รอเสร็จสมบูรณ์': "1hAmCZLJKwZSHdHT7AAcW3cqyIVYPmT2wJngR412sxYY",
      'การเติมค่าคลิกแรกเข้าไม่ครบขั้นต่ำ': "1QR-oPSc_xhp7gjfbdjsFX1p3BOyOSuZYeBDpuamkIDs",
      'ขอเปิดงานก่อนจ่ายค่าคลิก': "1jo9WPJ3gHIb0OILcFOY4CcsCddoyIb1biXOYOxymamE",
      'อื่นๆ': "1ACh5y2sAMEAaezE7ckRDrAZn8rRjedPIDMYGpQLLE4U",
      'ขอโปรโมชันค่าคลิก 12,000': "1SG73od4HWnEcOxchISwl9do5YPOaUZ36IGdElT6c5aY",
      'ขอแบ่งจ่ายแพ็กเกจ Total Solution': "1c70YgxS-Vg0cEzn3pY38bReaDt2yDPQ-n0yo2vlfJ-0",
      'ขอ Advance ค่าคลิกทีม DMS': "1na2blImtXBpjTEkmSBnIN6eD_Be16BR9y0Zu10fQ6uo"
    };
    
    // 2. กำหนด ID ของโฟลเดอร์ชั่วคราวสำหรับเก็บสำเนา Doc ที่แก้ไขแล้ว
    const tempFolderId = "1a__KHv_eO4KwDPZbwxOySbfXjtAupz_X"; 

    // 3. กำหนด ID ของโฟลเดอร์ปลายทางสำหรับเก็บไฟล์ PDF
    const pdfOutputFolderId = "1OuNNmbvOMqoMN9vjYESdsfRHvcxUicD-"; 

    // ดึงข้อมูลที่ส่งมาจากฟอร์ม
    const info = e.namedValues;

    // 4. ดึงคำตอบจากคำถามเงื่อนไขที่ใช้เลือก Template
    //    ใช้ 'choose_temp' ตามที่คุณระบุ
    const templateChoice = info['choose_temp'] ? info['choose_temp'][0] : null;

    // ตรวจสอบว่ามีการเลือก Template หรือไม่ และ Template ID นั้นมีอยู่จริง
    if (!templateChoice || !templateIds[templateChoice]) {
      console.error('Error: Template choice is missing or invalid. Please check your form question name and templateIds mapping.');
      return; 
    }

    // เลือก templateDocId ตามคำตอบที่ได้จากฟอร์ม
    const selectedTemplateDocId = templateIds[templateChoice];

    // เข้าถึง Google Doc Template และโฟลเดอร์ต่างๆ
    const templateDoc = DriveApp.getFileById(selectedTemplateDocId); 
    const tempFolder = DriveApp.getFolderById(tempFolderId);
    const pdfOutputFolder = DriveApp.getFolderById(pdfOutputFolderId);

    // กำหนดชื่อไฟล์สำหรับเอกสารชั่วคราวและ PDF
    const requestorName = info['name_request'] ? info['name_request'][0] : 'ไม่ระบุชื่อ';
    const formattedDate = Utilities.formatDate(new Date(), "GMT+0700", "dd/MM/yyyy HH:mm:ss");

    const tempDocName = `ใบขออนุมัติ_ชั่วคราว_${templateChoice}_${requestorName}_${formattedDate}`;
    const newTempFile = templateDoc.makeCopy(tempFolder).setName(tempDocName);
    
    // เปิดเอกสารที่คัดลอกมาเพื่อแก้ไขเนื้อหา
    const openDoc = DocumentApp.openById(newTempFile.getId());
    const body = openDoc.getBody();

    // วนลูปเพื่อแทนที่ตัวยึด (placeholders) ในเอกสาร
    for (const key in info) {
      const placeholder = "{" + key + "}"; 
      const value = (info[key] && info[key].length > 0) ? info[key][0] : ""; 

      try {
        body.replaceText(placeholder, value || "-"); 
      } catch (replaceError) {
        // Placeholder not found in this template, no action needed
      }
    }
    
    // บันทึกการเปลี่ยนแปลงและปิดเอกสาร
    openDoc.saveAndClose();

    // --- ส่วน: แปลงเป็น PDF และบันทึก ---
    
    const blobPDF = newTempFile.getBlob();

    const pdfFileName = `ใบขออนุมัติ${templateChoice} โดย${requestorName}_${formattedDate}.pdf`;

    const pdfFile = pdfOutputFolder.createFile(blobPDF).setName(pdfFileName);
    
    // ลบไฟล์ Google Doc ชั่วคราว
    newTempFile.setTrashed(true);

    console.log(`เอกสาร PDF '${pdfFileName}' ถูกสร้างและบันทึกเรียบร้อยแล้ว.`);
    const pdfLink = pdfFile.getUrl(); // เก็บลิงก์ PDF ไว้ใช้ในอีเมล
    console.log(`ลิงก์ไฟล์ PDF: ${pdfLink}`);

    // --- บันทึก PDF Link ลงใน Google Sheet ---
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const row = e.range.getRow(); // ดึงหมายเลขแถวที่ข้อมูลถูกส่งเข้ามา
    const columnB = 2; // คอลัมน์ B คือคอลัมน์ที่ 2

    // เขียนลิงก์ PDF ลงในคอลัมน์ A ของแถวที่ส่งฟอร์มเข้ามา
    sheet.getRange(row, columnB).setValue(pdfLink);
    console.log(`ลิงก์ PDF ถูกบันทึกลงในเซลล์ B${row} ของ Google Sheet แล้ว.`);

    // --- ส่งอีเมล ---

    // 5. โหลดเนื้อหา Email Template จากไฟล์ emailTemplate.html
    const emailTemplateHtml = HtmlService.createTemplateFromFile('emailTemplate').evaluate().getContent();

    // 6. ดึงอีเมลผู้รับจากฟอร์ม
    const recipientEmail = info['email_request'] ? info['email_request'][0] : null;
    const cus_web = info['cus_web'] ? info['cus_web'][0] : '-'; 

    // ตรวจสอบว่ามีอีเมลผู้รับหรือไม่
    if (recipientEmail) {
      // แทนที่ตัวยึดใน Email Template ด้วยข้อมูลจริง
      let finalEmailHtml = emailTemplateHtml.replace('{pdfLink}', pdfLink);
      finalEmailHtml = finalEmailHtml.replace('{name_request}', requestorName);
      finalEmailHtml = finalEmailHtml.replace('{cus_web}', cus_web);
      // คุณสามารถเพิ่มการแทนที่ตัวยึดอื่นๆ ที่คุณมีใน emailTemplate.html ได้ที่นี่
      // ตัวอย่าง: finalEmailHtml = finalEmailHtml.replace('{website_name}', info['website_name'][0]);

      // กำหนดหัวข้ออีเมล
      const emailSubject = `ใบขออนุมัติ${templateChoice} ของคุณ${requestorName}`;

      // กำหนดชื่อผู้ส่ง
      const senderName = "ITOPPLUS";

      // กำหนดผู้รับ CC
      const ccRecipients = 'siriwan@theiconweb.com';

      // ส่งอีเมล
      MailApp.sendEmail({
      to: recipientEmail,
      subject: emailSubject,
      htmlBody: finalEmailHtml,
      name: senderName,
      cc: ccRecipients
      });

      console.log(`อีเมลถูกส่งไปยัง ${recipientEmail} และ CC ไปยัง ${ccRecipients} เรียบร้อยแล้ว.`);
    } else {
      console.warn('ไม่พบอีเมลผู้รับในข้อมูลฟอร์ม ไม่สามารถส่งอีเมลได้.');
    }

  } catch (error) {
    console.error('เกิดข้อผิดพลาดในการประมวลผลฟอร์ม, สร้าง PDF หรือส่งอีเมล:', error);
  }
}

// ฟังก์ชันสำหรับรับวันที่และเวลาในรูปแบบที่กำหนด (เหมือนเดิม)
function getFormattedDate() {
  const today = new Date();
  const timeZone = "GMT+0700"; 
  return Utilities.formatDate(today, timeZone, "dd/MM/yyyy HH:mm:ss");
}
