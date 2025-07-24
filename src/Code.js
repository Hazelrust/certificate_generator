/** ใส่ชื่อ Sheet จาก google spreadsheet ข้างล่างนี้ */
const SHEET_NAME = 'Sheet1'; // ชื่อ sheet ที่เก็บรายชื่อ

// Template Configuration
/** ใส่ Template Slide ID ข้างล่างนี้ (อย่าลืมเปิดแชร์)*/
const TEMPLATE_SLIDE_ID = '1RiLgvhzens2oCVfu2RQPpVevymPkdSXZU9TGQ2NnHuQ'; // ID ของ Google Slides template
/** ใส่ ID ของโฟลเดอร์ที่จะเอาไปเก็บไฟล์ (อย่าลืมเปิดแชร์)*/
const OUTPUT_FOLDER_ID = '1iZawh5LNgj-qNY51-p_qC1w0u4D62Sb4'; // ID ของโฟลเดอร์ที่จะเก็บไฟล์ (ถ้าไม่ระบุจะเก็บใน root)

/** ใน slide อยากใส่อะไรให้เขียนแบบนั้น เช่น <<ชื่อสกุล>> */
const PLACEHOLDERS = {
  NAME: '<<ชื่อสกุล>>',
};

/**
 * สร้าง Web App
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * รวมไฟล์ HTML/CSS/JS
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ตรวจสอบรายชื่อใน Google Sheets
 * คอลัมน์ A = Email, คอลัมน์ B = ชื่อ-นามสกุล (เช่น "สมชาย ใจดี")
 */
function checkName(fullName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    const cleanInputName = fullName.trim().replace(/\s+/g, ' ');
    
    // ข้าม header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const email = row[0] ? row[0].toString().trim() : '';
      const sheetFullName = row[1] ? row[1].toString().trim().replace(/\s+/g, ' ') : '';
      
      // เปรียบเทียบชื่อ-นามสกุล
      if (sheetFullName.toLowerCase() === cleanInputName.toLowerCase()) {
        // แยกชื่อ-นามสกุลจากข้อมูลใน sheet
        const nameParts = sheetFullName.split(' ');
        const firstName = nameParts[0] || '';
        const lastName = nameParts.slice(1).join(' ') || '';
        
        return {
          success: true,
          message: 'พบรายชื่อของคุณในระบบ',
          data: {
            fullName: sheetFullName,
            firstName: firstName,
            lastName: lastName,
            email: email,
          }
        };
      }
    }
    
    return {
      success: false,
      message: 'ไม่พบรายชื่อของคุณในระบบ กรุณาตรวจสอบชื่อ-นามสกุลให้ถูกต้อง (เช่น สมชาย ใจดี)'
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'เกิดข้อผิดพลาด: ' + error.toString()
    };
  }
}

/**
 * สร้างเกียรติบัตรจาก Google Slides template
 */
function generateCertificateFromSlides(userData) {
  try {
    if (!TEMPLATE_SLIDE_ID === TEMPLATE_SLIDE_ID) {
      return {
        success: false,
        message: 'กรุณากำหนด Template Slide ID ในโค้ด'
      };
    }

    const template = DriveApp.getFileById(TEMPLATE_SLIDE_ID);
    const certificateName = `เกียรติบัตร_${userData.fullName}_${new Date().toLocaleDateString('th-TH')}`;
    
    let targetFolder;
    if (OUTPUT_FOLDER_ID == OUTPUT_FOLDER_ID) {
      targetFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
    } else {
      targetFolder = DriveApp.getRootFolder();
    }
    
    const newFile = template.makeCopy(certificateName, targetFolder);
    const presentation = SlidesApp.openById(newFile.getId());
    
    // const currentDate = new Date().toLocaleDateString('th-TH');
    const replacements = [
      { placeholder: PLACEHOLDERS.NAME, value: userData.fullName },
      // { placeholder: PLACEHOLDERS.EVENT, value: userData.eventName },
      // { placeholder: PLACEHOLDERS.DATE, value: currentDate },
      // { placeholder: PLACEHOLDERS.ORGANIZATION, value: userData.organization }
    ];
    
    // แทนที่ข้อความในทุกหน้า
    const slides = presentation.getSlides();
    slides.forEach(slide => {
      replacements.forEach(replacement => {
        slide.replaceAllText(replacement.placeholder, replacement.value);
      });
    });
    
    // บันทึกการเปลี่ยนแปลง
    presentation.saveAndClose();
    
    // สร้าง PDF
    const pdfBlob = newFile.getAs(MimeType.PDF);
    const pdfFile = targetFolder.createFile(pdfBlob);
    pdfFile.setName(`${certificateName}.pdf`);
    
    // ลบไฟล์ slide ชั่วคราว (ถ้าต้องการเก็บแค่ PDF)
    // newFile.setTrashed(true);
    DriveApp.getFileById(newFile.getId()).setTrashed(true);
    
    // สร้าง download link
    const downloadUrl = `https://drive.google.com/uc?export=download&id=${pdfFile.getId()}`;
    
    return {
      success: true,
      message: 'สร้างเกียรติบัตรสำเร็จ',
      downloadUrl: downloadUrl,
      fileId: pdfFile.getId(),
      fileName: pdfFile.getName()
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการสร้างเกียรติบัตร: ' + error.toString()
    };
  }
}

/**
 * ดึง thumbnail ของ Google Slides เพื่อแสดงตัวอย่าง
 */
function getSlideThumbnail() {
  try {
    if (!TEMPLATE_SLIDE_ID  === TEMPLATE_SLIDE_ID) {
      return {
        success: false,
        message: 'ไม่พบ Template ID'
      };
    }

    // สร้าง thumbnail URL
    const thumbnailUrl = `https://docs.google.com/presentation/d/${TEMPLATE_SLIDE_ID}/export/png?pageid=p`;
    
    return {
      success: true,
      thumbnailUrl: thumbnailUrl
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'ไม่สามารถโหลดภาพตัวอย่างได้: ' + error.toString()
    };
  }
}


/**
 * สร้าง URL สำหรับดาวน์โหลดไฟล์จาก Google Drive
 */
function getDownloadUrl(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    
    // ตั้งค่าให้ทุกคนสามารถดูได้
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return {
      success: true,
      downloadUrl: `https://drive.google.com/uc?export=download&id=${fileId}`,
      viewUrl: `https://drive.google.com/file/d/${fileId}/view`
    };
  } catch (error) {
    return {
      success: false,
      message: 'ไม่สามารถสร้าง download link ได้: ' + error.toString()
    };
  }
}

/**
 * บันทึกการดาวน์โหลด (Log)
 */
function logDownload(fullName, email, fileId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('DownloadLog');
    
    if (!logSheet) {
      logSheet = ss.insertSheet('DownloadLog');
      logSheet.getRange(1, 1, 1, 5).setValues([['วันที่-เวลา', 'ชื่อ-นามสกุล', 'Email', 'File ID', 'สถานะ']]);
    }
    
    logSheet.appendRow([
      new Date(),
      fullName,
      email || 'ไม่ระบุ',
      fileId || 'ไม่ระบุ',
      'สร้างเกียรติบัตรสำเร็จ'
    ]);
    
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * ฟังก์ชันช่วยเหลือ: ดึงรายชื่อทั้งหมดเพื่อแสดงตัวอย่าง
 */
function getAllNames() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const names = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const email = row[0] ? row[0].toString().trim() : '';
      const fullName = row[1] ? row[1].toString().trim() : '';
      
      if (fullName) {
        names.push({ email: email, fullName: fullName });
      }
    }
    
    // Remove after production
    console.log(names)
    return { success: true, names: names };
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * ฟังก์ชันทดสอบการสร้างเกียรติบัตร
 */
function testCertificateGeneration() {
  const testData = {
    fullName: 'ทดสอบ ระบบ',
    email: 'test@example.com',
  };
  
  const result = generateCertificateFromSlides(testData);
  console.log('Test Result:', result);
  return result;
}