# 🏆 Certificate Generator with Google Apps Script

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

| A (Email) | B (Full Name) |
|-----------|---------------|
| john@email.com | John Doe |
| jane@email.com | Jane Smith |
| somchai@email.com | สมชาย ใจดี |

**Important**: Names in column B must have spaces between first and last names.

### 2. Create Certificate Template

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
