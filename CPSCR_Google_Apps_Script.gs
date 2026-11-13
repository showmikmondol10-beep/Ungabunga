/**
 * CPSCR Cultural Fest registration receiver
 *
 * Setup:
 * 1. Create a Google Spreadsheet.
 * 2. Open Extensions > Apps Script from that spreadsheet.
 * 3. Paste this entire file into Code.gs and save.
 * 4. Deploy > New deployment > Web app.
 *    Execute as: Me
 *    Who has access: Anyone
 * 5. Copy the Web App URL and paste it into the website's
 *    GOOGLE_APPS_SCRIPT_URL constant.
 *
 * The notification recipient is fixed to the organizer's address.
 * Do not ask students for banking PINs, passwords, or verification codes.
 */

const ADMIN_EMAIL = 'showmikmondol04@gmail.com';
const SHEET_NAME = 'Registrations';

const ALLOWED_CLASSES = ['Class 9', 'Class 10', 'Class 11', 'Class 12'];
const ALLOWED_SECTIONS = ['Shapla', 'Joba', 'Golap', 'Bely'];
const ALLOWED_PAYMENTS = ['Cash', 'bKash', 'Nagad', 'Rocket', 'Bank transfer', 'Other'];

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Empty submission.');
    }

    const data = JSON.parse(e.postData.contents);
    const name = cleanText(data.name, 100);
    const studentId = cleanText(data.studentId, 40);
    const studentClass = cleanText(data.class, 20);
    const section = cleanText(data.section, 30);
    const payment = cleanText(data.payment, 30);

    if (!name || !studentId) throw new Error('Name and Student ID are required.');
    if (!ALLOWED_CLASSES.includes(studentClass)) throw new Error('Invalid class.');
    if (!ALLOWED_SECTIONS.includes(section)) throw new Error('Invalid section.');
    if (!ALLOWED_PAYMENTS.includes(payment)) throw new Error('Invalid payment method.');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('This script must be attached to the registration spreadsheet.');

    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Name', 'Class', 'Section', 'Student ID', 'Payment Method']);
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      safeForSheet(name),
      safeForSheet(studentClass),
      safeForSheet(section),
      safeForSheet(studentId),
      safeForSheet(payment)
    ]);

    const submittedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy, hh:mm a');
    const subject = 'New CPSCR Cultural Fest Registration';
    const htmlBody =
      '<div style="font-family:Arial,sans-serif;max-width:560px;margin:auto;padding:24px;border:1px solid #eee;border-radius:12px">' +
      '<h2 style="color:#6757d9;margin-top:0">New registration received</h2>' +
      '<p>A student has submitted the CPSCR Cultural Fest form.</p>' +
      '<table cellpadding="8" cellspacing="0" style="border-collapse:collapse;width:100%">' +
      emailRow('Name', name) +
      emailRow('Class', studentClass) +
      emailRow('Section', section) +
      emailRow('Student ID', studentId) +
      emailRow('Payment method', payment) +
      emailRow('Submitted at', submittedAt) +
      '</table><p style="color:#777;font-size:12px">This is an automated notification from the CPSCR Cultural Fest registration form.</p></div>';

    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: subject,
      htmlBody: htmlBody,
      body: 'New CPSCR Cultural Fest registration\n' +
            'Name: ' + name + '\nClass: ' + studentClass +
            '\nSection: ' + section + '\nStudent ID: ' + studentId +
            '\nPayment method: ' + payment + '\nSubmitted at: ' + submittedAt
    });

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    console.error(err);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: String(err.message || err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function cleanText(value, maxLength) {
  return String(value == null ? '' : value).trim().replace(/[<>]/g, '').slice(0, maxLength);
}

function safeForSheet(value) {
  // Prevent a submitted value from being interpreted as a spreadsheet formula.
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}

function emailRow(label, value) {
  return '<tr><td style="border-bottom:1px solid #eee;color:#666;width:38%"><b>' +
    escapeHtml(label) + '</b></td><td style="border-bottom:1px solid #eee">' +
    escapeHtml(value) + '</td></tr>';
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, function(ch) {
    return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch];
  });
}
