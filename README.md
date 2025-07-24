# 🏆 Certificate Generator with Google Apps Script / ระบบทำเกียรติบัตรออนไลน์ด้วย Google Apps Script 

[For English description](#for-en)

<img width="873" height="817" alt="image" src="https://github.com/user-attachments/assets/4aee19b7-f809-4e66-a681-ef650c2b6209" />

ระบบสร้างเกียรติบัตรออนไลน์ที่ใช้ Google Slides เป็น template และสร้างไฟล์ PDF ให้ผู้ใช้ดาวน์โหลดอัตโนมัติ โดยตรวจสอบรายชื่อจาก Google Sheets และสร้างเกียรติบัตรเฉพาะผู้ที่มีสิทธิ์เท่านั้น

## ✨ คุณสมบัติหลัก

- 🎨 **Template ที่ปรับแต่งได้**: ใช้ Google Slides ออกแบบเกียรติบัตรตามต้องการ
- 📊 **เชื่อมต่อ Google Sheets**: ตรวจสอบรายชื่อจากฐานข้อมูลอัตโนมัติ
- 📱 **Responsive Design**: ใช้งานได้ทุกอุปกรณ์ (มือถือ แท็บเล็ต คอมพิวเตอร์)
- 🔒 **ควบคุมการเข้าถึง**: เฉพาะผู้ที่ลงทะเบียนแล้วเท่านั้นที่สร้างเกียรติบัตรได้
- 📁 **จัดเก็บอย่างเป็นระบบ**: บันทึกไฟล์ใน Google Drive folder ที่กำหนด
- 📋 **ติดตามการดาวน์โหลด**: บันทึก log การสร้างเกียรติบัตรครบถ้วน
- 🌐 **รองรับภาษาไทย**: แสดงผลและประมวลผลภาษาไทยได้อย่างถูกต้อง

## 🚀 เริ่มต้นใช้งาน

### สิ่งที่ต้องมี

- บัญชี Google ที่สามารถเข้าถึง:
  - Google Sheets
  - Google Slides  
  - Google Drive
  - Google Apps Script

### 1. ตั้งค่า Google Sheets

สร้าง Google Sheets ที่มีโครงสร้างดังนี้:

<img width="352" height="137" alt="image" src="https://github.com/user-attachments/assets/6662a80f-eb0a-4209-bd21-3003d58e1414" />

| A (Email) | B (ชื่อ-นามสกุล) |
|-----------|----------------|
| john@email.com | John Doe |
| jane@email.com | Jane Smith |
| somchai@email.com | สมชาย ใจดี |


**สำคัญ**: ชื่อในคอลัมน์ B ต้องมีช่องว่างระหว่างชื่อและนามสกุล

### 2. สร้าง Template เกียรติบัตร

<img width="941" height="678" alt="image" src="https://github.com/user-attachments/assets/7d555816-07b6-4331-a3c9-654aa2df2a73" />

1. สร้าง Google Slides ใหม่
2. ออกแบบเกียรติบัตรตามต้องการ
3. ใส่ตัวแปรในตำแหน่งที่ต้องการให้เปลี่ยนแปลง:
   - `<<ชื่อสกุล>>` หรือ `{{NAME}}` - จะถูกแทนที่ด้วยชื่อผู้เข้าร่วม
   - `{{EVENT}}` - จะถูกแทนที่ด้วยชื่อกิจกรรม (ไม่บังคับ)
   - `{{DATE}}` - จะถูกแทนที่ด้วยวันที่ปัจจุบัน (ไม่บังคับ)
   - `{{ORGANIZATION}}` - จะถูกแทนที่ด้วยชื่อหน่วยงาน (ไม่บังคับ)

### 3. ตั้งค่า Google Apps Script

1. เปิด Google Sheets ที่สร้างไว้
2. ไปที่ **ส่วนขยาย** → **Apps Script**
3. ลบโค้ดเดิมในไฟล์ `Code.gs`
4. Copy โค้ดจากไฟล์ [`Code.gs`](./Code.gs) มาแทน
5. สร้างไฟล์ HTML ใหม่: **ไฟล์** → **ใหม่** → **ไฟล์ HTML**
6. ตั้งชื่อไฟล์ว่า `Index`
7. Copy โค้ดจากไฟล์ [`Index.html`](./Index.html) มาแทน

### 4. กำหนดค่าตัวแปร

แก้ไขค่าตัวแปรในไฟล์ `Code.gs`:

```javascript
// การตั้งค่าที่จำเป็น
const SHEET_NAME = 'Sheet1'; // ชื่อ sheet ที่เก็บข้อมูลผู้เข้าร่วม
const TEMPLATE_SLIDE_ID = 'ใส่ ID ของ Slide Template'; // ID ของ Google Slides template
const OUTPUT_FOLDER_ID = 'ใส่ ID ของ Folder'; // ID ของ Google Drive folder

// การตั้งค่าเพิ่มเติม (ไม่บังคับ)
const EVENT_NAME = 'ชื่อกิจกรรมของคุณ';
const ORGANIZATION = 'ชื่อหน่วยงานของคุณ';
const PLACEHOLDERS = {
  NAME: '<<ชื่อสกุล>>', // ปรับให้ตรงกับ template ของคุณ
  EVENT: '{{EVENT}}',
  DATE: '{{DATE}}',
  ORGANIZATION: '{{ORGANIZATION}}'
};
```

### 5. วิธีหา ID ที่ต้องการ

#### ID ของ Google Slides Template
จาก URL ของ Google Slides:
```
https://docs.google.com/presentation/d/[SLIDE_ID]/edit
```
Copy ส่วน `SLIDE_ID`

#### ID ของ Google Drive Folder  
จาก URL ของ Google Drive folder:
```
https://drive.google.com/drive/folders/[FOLDER_ID]
```
Copy ส่วน `FOLDER_ID`

### 6. เปิดใช้งาน API ที่จำเป็น

1. ใน Apps Script คลิก **บริการ** (Services) ที่แถบด้านซ้าย
2. เพิ่มบริการเหล่านี้:
   - **Google Slides API**
   - **Google Drive API**
3. คลิก **บันทึก**

### 7. Deploy Web App

1. คลิก **Deploy** → **การใช้งานใหม่**
2. ประเภท: **เว็บแอป**
3. คำอธิบาย: `ระบบสร้างเกียรติบัตร`
4. ดำเนินการเป็น: **ฉัน**
5. ผู้ที่มีสิทธิ์เข้าถึง: **ทุกคน** (สำหรับการเข้าถึงสาธารณะ)
6. คลิก **Deploy**
7. **Copy URL ของ web app**

### 8. ตั้งค่าสิทธิ์

1. ตั้งค่า Google Slides template เป็น **ทุกคนที่มีลิงก์สามารถดูได้**
2. ตรวจสอบให้แน่ใจว่า Google Drive folder อนุญาตให้ script เขียนไฟล์ได้

## 🔧 การปรับแต่ง

### เปลี่ยนดีไซน์ Template

1. แก้ไข Google Slides template ของคุณ
2. ตรวจสอบให้แน่ใจว่าตัวแปรตรงกับที่กำหนดใน object `PLACEHOLDERS`
3. การเปลี่ยนแปลงจะมีผลทันที (ไม่ต้อง deploy ใหม่)

### แก้ไขตัวแปร (Placeholders)

อัปเดต object `PLACEHOLDERS` ในไฟล์ `Code.gs`:

```javascript
const PLACEHOLDERS = {
  NAME: '{{ชื่อผู้เข้าร่วม}}',        // ตัวแปรที่คุณกำหนดเอง
  EVENT: '{{ชื่อกิจกรรม}}',         // ตัวแปรที่คุณกำหนดเอง
  DATE: '{{วันที่เกียรติบัตร}}',    // ตัวแปรที่คุณกำหนดเอง
  ORGANIZATION: '{{ชื่อองค์กร}}'    // ตัวแปรที่คุณกำหนดเอง
};
```

### ปรับแต่งหน้าเว็บ

หน้าเว็บรองรับทุกอุปกรณ์และปรับแต่งได้เต็มที่ แก้ไข CSS ในไฟล์ `Index.html`:

- **สี**: แก้ไขสีไล่ระดับใน class `.header` และ `.btn`
- **ฟอนต์**: เปลี่ยน `font-family` ใน selector `body`
- **เลย์เอาต์**: ปรับระยะห่าง padding และความกว้างของ container
- **มือถือ**: จุดแบ่งสำหรับ responsive ได้ตั้งค่าไว้แล้ว

### เพิ่มการตรวจสอบแบบกำหนดเอง

เพิ่ม logic การตรวจสอบในฟังก์ชัน `checkName()`:

```javascript
// ตัวอย่าง: จำกัดเฉพาะ email domain เฉพาะ
if (!email.endsWith('@yourcompany.com')) {
  return {
    success: false,
    message: 'อนุญาตเฉพาะอีเมลของบริษัทเท่านั้น'
  };
}
```

## 📊 การติดตามและ Log

ระบบจะสร้าง sheet ชื่อ `DownloadLog` อัตโนมัติ ที่มีข้อมูล:

<img width="791" height="470" alt="image" src="https://github.com/user-attachments/assets/b2bebf72-d629-403d-92d9-1c516d586437" />


- **วันที่-เวลา**: เวลาที่สร้างเกียรติบัตร
- **ชื่อ-นามสกุล**: ชื่อผู้เข้าร่วม
- **Email**: อีเมลผู้เข้าร่วมจากฐานข้อมูล
- **File ID**: ID ไฟล์ของเกียรติบัตรที่สร้างใน Google Drive
- **สถานะ**: สถานะการสร้าง


### ตัวอย่างการวิเคราะห์ Log

```javascript
// นับจำนวนเกียรติบัตรทั้งหมดที่สร้าง
=COUNTIF(E:E,"สร้างเกียรติบัตรสำเร็จ")

// หาการดาวน์โหลดซ้ำ
=COUNTIF(B:B,B2)>1

// แสดงรายชื่อผู้เข้าร่วมที่ไม่ซ้ำ
=UNIQUE(B:B)
```

## 🔒 ความปลอดภัย

### การควบคุมการเข้าถึง
- เฉพาะผู้เข้าร่วมที่อยู่ใน Google Sheets เท่านั้นที่สร้างเกียรติบัตรได้
- ชื่อต้องตรงกันทุกตัวอักษร (ไม่คำนึงตัวพิมพ์เล็ก-ใหญ่)
- การสร้างทุกครั้งจะถูกบันทึก log พร้อม timestamp

### ป้องกันการใช้งานผิดปกติ
เพิ่มการจำกัดอัตราใน `Code.gs`:

```javascript
// ตัวอย่าง: หนึ่งเกียรติบัตรต่อคนต่อวัน
const today = new Date().toDateString();
const recentLogs = logSheet.getDataRange().getValues()
  .filter(row => row[1] === fullName && new Date(row[0]).toDateString() === today);

if (recentLogs.length > 0) {
  return {
    success: false,
    message: 'คุณได้รับเกียรติบัตรในวันนี้แล้ว'
  };
}
```

### ความเป็นส่วนตัวของข้อมูล
- ข้อมูลผู้เข้าร่วมอยู่ในบัญชี Google ของคุณ
- ไม่มีการใช้บริการภายนอก
- การเข้ารหัส SSL ผ่านโครงสร้างพื้นฐานของ Google

## 🚨 การแก้ไขปัญหา

### ปัญหาที่พบบ่อย

**"กรุณากำหนด Template Slide ID ในโค้ด"**
- ตรวจสอบว่า `TEMPLATE_SLIDE_ID` ตั้งค่าถูกต้อง
- ตรวจสอบว่า ID ไม่มีข้อความตัวอย่างอยู่
- ยืนยันว่า Google Slides เข้าถึงได้

**"ไม่พบรายชื่อของคุณในระบบ"**
- ตรวจสอบการสะกดและช่องว่างของชื่อใน Google Sheets
- ตรวจสอบว่าคอลัมน์ B มีชื่อเต็มพร้อมช่องว่าง
- ยืนยันว่า `SHEET_NAME` ตรงกับชื่อ sheet ของคุณ

**"เกิดข้อผิดพลาดในการสร้างเกียรติบัตร"**
- ยืนยันว่า Google Slides API เปิดใช้งานแล้ว
- ตรวจสอบว่า output folder มีอยู่และเข้าถึงได้
- ยืนยันว่าข้อความตัวแปรตรงกับ template

**Web app โหลดไม่ได้**
- ตรวจสอบว่า deployment ตั้งค่าการเข้าถึงเป็น "ทุกคน"
- ตรวจสอบว่า URL ของ web app ถูกต้อง
- ลอง deploy เวอร์ชันใหม่

### โหมดการแก้ไขจุดบกพร่อง

เปิดใช้งาน debugging ใน `Code.gs`:

```javascript
function testCertificateGeneration() {
  const testData = {
    fullName: 'ผู้ใช้ทดสอบ',
    email: 'test@example.com',
    eventName: EVENT_NAME,
    organization: ORGANIZATION
  };
  
  const result = generateCertificateFromSlides(testData);
  console.log('ผลการทดสอบ:', result);
  return result;
}
```

เรียกใช้ฟังก์ชันนี้เพื่อทดสอบการสร้างเกียรติบัตร

## 🤝 การมีส่วนร่วม

1. Fork repository
2. สร้าง feature branch (`git checkout -b feature/คุณสมบัติใหม่`)
3. Commit การเปลี่ยนแปลง (`git commit -m 'เพิ่มคุณสมบัติใหม่'`)
4. Push ไปยัง branch (`git push origin feature/คุณสมบัติใหม่`)
5. เปิด Pull Request

## 📄 ลิขสิทธิ์

โปรเจกต์นี้อยู่ภายใต้ลิขสิทธิ์ MIT License - ดูรายละเอียดในไฟล์ [LICENSE](LICENSE)

## 🙏 การขอบคุณ

- สร้างด้วย Google Apps Script และ Google Workspace APIs
- Responsive design ใช้เทคนิค CSS สมัยใหม่
- การสร้างเกียรติบัตรใช้พลัง Google Slides API

---

## 📧 การสนับสนุน

หากพบปัญหาหรือมีคำถาม:

1. ตรวจสอบหัวข้อ [การแก้ไขปัญหา](#-การแก้ไขปัญหา) ก่อน
2. ตรวจสอบการตั้งค่าของคุณ
3. เปิด issue บน GitHub พร้อมระบุ:
   - รายละเอียดปัญหา
   - ข้อความแสดงข้อผิดพลาด (ถ้ามี)
   - ขั้นตอนการทำซ้ำ
   - การตั้งค่าของคุณ (โดยไม่ระบุ ID ที่เป็นความลับ)

---

## 💡 เคล็ดลับการใช้งาน

### การปรับปรุงประสิทธิภาพ
- ใช้โฟลเดอร์แยกสำหรับแต่ละกิจกรรม
- ลบไฟล์ Slides ชั่วคราวเพื่อประหยัดพื้นที่
- ตั้งค่า rate limiting เพื่อป้องกันการใช้งานมากเกินไป

### การจัดการหลายกิจกรรม
```javascript
// ตัวอย่างการตั้งค่าหลาย template
const TEMPLATES = {
  'workshop': '1abc123...',
  'seminar': '1def456...',
  'conference': '1ghi789...'
};

// เลือก template ตามประเภทกิจกรรม
const templateId = TEMPLATES[eventType] || TEMPLATE_SLIDE_ID;
```

### การ Backup ข้อมูล
- Export Google Sheets เป็น CSV เป็นประจำ
- Backup DownloadLog sheet
- เก็บ copy ของ Slides template

---

# for-en

# 🏆 Certificate Generator with Google Apps Script

<img width="352" height="137" alt="image" src="https://github.com/user-attachments/assets/6662a80f-eb0a-4209-bd21-3003d58e1414" />

A web-based certificate generation system that automatically creates personalized certificates from Google Slides templates and exports them as PDFs. Users can enter their name to verify against a Google Sheets database and download their certificate instantly.

## ✨ Features

- 🎨 **Custom Templates**: Use your own Google Slides as certificate templates
- 📊 **Google Sheets Integration**: Automatic name verification from spreadsheet data
- 📱 **Responsive Design**: Works seamlessly on all devices (mobile, tablet, desktop)
- 🔒 **Access Control**: Only registered participants can generate certificates
- 📁 **Organized Storage**: Certificates saved to specified Google Drive folder
- 📋 **Download Tracking**: Complete logs of certificate generation activity
- 🌐 **Multi-language Support**: Thai and English text support

## 🚀 Quick Start

### Prerequisites

- Google Account with access to:
  - Google Sheets
  - Google Slides  
  - Google Drive
  - Google Apps Script

### 1. Setup Google Sheets

Create a Google Sheets with the following structure:

<img width="352" height="137" alt="image" src="https://github.com/user-attachments/assets/6662a80f-eb0a-4209-bd21-3003d58e1414" />

| A (Email) | B (Full Name) |
|-----------|---------------|
| john@email.com | John Doe |
| jane@email.com | Jane Smith |
| somchai@email.com | สมชาย ใจดี |

**Important**: Names in column B must have spaces between first and last names.

### 2. Create Certificate Template

<img width="941" height="678" alt="image" src="https://github.com/user-attachments/assets/2517b2dc-d917-42d8-8c5b-c23e23d80c71" />

1. Create a new Google Slides presentation
2. Design your certificate template
3. Insert placeholders where dynamic content should appear:
   - `<<ชื่อสกุล>>` or `{{NAME}}` - Will be replaced with participant name
   - `{{EVENT}}` - Will be replaced with event name (optional)
   - `{{DATE}}` - Will be replaced with current date (optional)
   - `{{ORGANIZATION}}` - Will be replaced with organization name (optional)

### 3. Setup Google Apps Script

1. Open your Google Sheets
2. Go to **Extensions** → **Apps Script**
3. Delete existing code in `Code.gs`
4. Copy and paste the code from [`Code.gs`](./Code.gs)
5. Create a new HTML file: **File** → **New** → **HTML file**
6. Name it `Index`
7. Copy and paste the code from [`Index.html`](./Index.html)

### 4. Configure Settings

Edit the configuration variables in `Code.gs`:

```javascript
// Required Configuration
const SHEET_NAME = 'Sheet1'; // Name of your sheet with participant data
const TEMPLATE_SLIDE_ID = 'YOUR_SLIDE_ID_HERE'; // Google Slides template ID
const OUTPUT_FOLDER_ID = 'YOUR_FOLDER_ID_HERE'; // Google Drive folder ID

// Optional Configuration  
const EVENT_NAME = 'Your Event Name';
const ORGANIZATION = 'Your Organization';
const PLACEHOLDERS = {
  NAME: '<<ชื่อสกุล>>', // Adjust to match your template
  EVENT: '{{EVENT}}',
  DATE: '{{DATE}}',
  ORGANIZATION: '{{ORGANIZATION}}'
};
```

### 5. Get Required IDs

#### Google Slides Template ID
From your Google Slides URL:
```
https://docs.google.com/presentation/d/[SLIDE_ID]/edit
```
Copy the `SLIDE_ID` part.

#### Google Drive Folder ID  
From your Google Drive folder URL:
```
https://drive.google.com/drive/folders/[FOLDER_ID]
```
Copy the `FOLDER_ID` part.

### 6. Enable Required APIs

1. In Apps Script, click **Services** (left sidebar)
2. Add these services:
   - **Google Slides API**
   - **Google Drive API**
3. Click **Save**

### 7. Deploy Web App

1. Click **Deploy** → **New deployment**
2. Type: **Web app**
3. Description: `Certificate Generator`
4. Execute as: **Me**
5. Who has access: **Anyone** (for public access)
6. Click **Deploy**
7. **Copy the web app URL**

### 8. Set Permissions

1. Make sure your Google Slides template is set to **Anyone with the link can view**
2. Ensure the Google Drive folder allows the script to write files

## 🔧 Customization

### Changing the Template Design

1. Edit your Google Slides template
2. Ensure placeholders match those defined in `PLACEHOLDERS` object
3. The changes will be reflected immediately (no need to redeploy)

### Modifying Placeholders

Update the `PLACEHOLDERS` object in `Code.gs`:

```javascript
const PLACEHOLDERS = {
  NAME: '{{PARTICIPANT_NAME}}',     // Your custom placeholder
  EVENT: '{{EVENT_TITLE}}',        // Your custom placeholder  
  DATE: '{{CERTIFICATE_DATE}}',    // Your custom placeholder
  ORGANIZATION: '{{ORG_NAME}}'     // Your custom placeholder
};
```

### Styling the Web Interface

The interface is fully responsive and customizable. Edit the CSS in `Index.html`:

- **Colors**: Modify gradient colors in `.header` and `.btn` classes
- **Fonts**: Change `font-family` in the `body` selector
- **Layout**: Adjust spacing, padding, and container widths
- **Mobile**: Responsive breakpoints are already configured

### Adding Custom Validation

Add custom validation logic in the `checkName()` function:

```javascript
// Example: Restrict to specific email domains
if (!email.endsWith('@yourcompany.com')) {
  return {
    success: false,
    message: 'Only company email addresses are allowed'
  };
}
```

## 📊 Monitoring and Logs

The system automatically creates a `DownloadLog` sheet with:

<img width="791" height="470" alt="image" src="https://github.com/user-attachments/assets/b2bebf72-d629-403d-92d9-1c516d586437" />

- **Date/Time**: When certificate was generated
- **Full Name**: Participant name  
- **Email**: Participant email from database
- **File ID**: Google Drive file ID of generated certificate
- **Status**: Generation status

### Log Analysis Examples

```javascript
// Count total certificates generated
=COUNTIF(E:E,"สร้างเกียรติบัตรสำเร็จ")

// Find duplicate downloads
=COUNTIF(B:B,B2)>1

// List unique participants
=UNIQUE(B:B)
```

## 🔒 Security Considerations

### Access Control
- Only participants in the Google Sheets can generate certificates
- Names must match exactly (case-insensitive)
- Each generation is logged with timestamp

### Preventing Abuse
Add rate limiting in `Code.gs`:

```javascript
// Example: One certificate per person per day
const today = new Date().toDateString();
const recentLogs = logSheet.getDataRange().getValues()
  .filter(row => row[1] === fullName && new Date(row[0]).toDateString() === today);

if (recentLogs.length > 0) {
  return {
    success: false,
    message: 'คุณได้รับเกียรติบัตรในวันนี้แล้ว'
  };
}
```

### Data Privacy
- Participant data stays in your Google account
- No external services involved
- SSL encryption via Google infrastructure

## 🚨 Troubleshooting

### Common Issues

**"กรุณากำหนด Template Slide ID ในโค้ด"**
- Check that `TEMPLATE_SLIDE_ID` is set correctly
- Ensure the ID doesn't contain the example placeholder text
- Verify the Google Slides is accessible

**"ไม่พบรายชื่อของคุณในระบบ"**
- Check name spelling and spacing in Google Sheets
- Ensure column B contains full names with spaces
- Verify `SHEET_NAME` matches your sheet name

**"เกิดข้อผิดพลาดในการสร้างเกียรติบัตร"**
- Confirm Google Slides API is enabled
- Check that output folder exists and is accessible
- Verify placeholder text matches template

**Web app not loading**
- Ensure deployment is set to "Anyone" access
- Check that the web app URL is correct
- Try deploying a new version

### Debug Mode

Enable debugging in `Code.gs`:

```javascript
function testCertificateGeneration() {
  const testData = {
    fullName: 'Test User',
    email: 'test@example.com',
    eventName: EVENT_NAME,
    organization: ORGANIZATION
  };
  
  const result = generateCertificateFromSlides(testData);
  console.log('Test Result:', result);
  return result;
}
```

Run this function to test certificate generation.

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built with Google Apps Script and Google Workspace APIs
- Responsive design using modern CSS techniques
- Certificate generation powered by Google Slides API

---

## 📧 Support

If you encounter any issues or have questions:

1. Check the [Troubleshooting](#-troubleshooting) section
2. Review your configuration settings
3. Open an issue on GitHub with:
   - Description of the problem
   - Error messages (if any)
   - Steps to reproduce
   - Your configuration (without sensitive IDs)

---

**Happy Certificate Generating! 🎉**


