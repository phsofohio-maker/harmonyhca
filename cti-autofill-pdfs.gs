/**
 * Parrish Health Systems - Patient Certification Notification Script
 * Automatically checks for patients needing certification and sends notifications
 */

// Configuration
const CONFIG = {
  SHEET_ID: '13-V3b2gckV1C2tyXZ0iq9cEfv9y4DyBRFxa3I-8U15w',
  EMAIL_LIST: [
    'kobet@parrishhealthsystems.org',
    'reneesha@parrishhealthsystems.org',
    'tajuanna@parrishhealthsystems.org'
  ],
  DOCTOR_NAME: 'Dr. Thomas Smallwood',
  DOC_TEMPLATES: {
    '60DAY': '1CkUx8NCYOwNJEDnVkIQGsl_gRZt_AjD_hlW0ZG0AzTw',
    '90DAY1': '1-OFQEG2c2B4v65Rpyr7gn6A3fAqIC4-Qr65BEJWuUTo',
    '90DAY2': '1IB9I_BOGwweBZJUtu7XDOKwlghQJvC9MF4XAx6Vkjmk',
    'ATTEND_CERT': '1H74TZgRCXL4hdoTBdjXwRIVwGi-QjdAX1wXvBGF9Ee8',
    'PROGRESS_NOTE': '1PObRDB6JVBLvlMgMOw_6owbvucBUJdak58lH2u61YCM',
    'PATIENT_HISTORY': '1MRYBd6soKZMhx8Autzegm78FpGF4mi1a9L2Eva7ORaM'
  },
  // Column indices (0-based for array access)
  COLUMNS: {
    ADMISSION_DATE: 1,  // B
    NOTIFY_DATE: 5,     // F
    CDATE_1: 6,         // G
    CDATE_2: 7,         // H
    MR_NUMBER: 10,      // K
    PATIENT_NAME: 11    // L
  }
};

/**
 * Main function - Run daily at 9am via trigger
 */
function checkCertificationNotifications() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const today = normalizeDate(new Date());
  
  // Skip header row
  const patients = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const notifyDate = parseDate(row[CONFIG.COLUMNS.NOTIFY_DATE]);
    
    if (!notifyDate) continue;
    
    const normalizedNotify = normalizeDate(notifyDate);
    
    // Check if notify date equals today
    if (normalizedNotify.getTime() === today.getTime()) {
      const patientData = extractPatientData(row, i + 1);
      if (patientData) {
        patients.push(patientData);
      }
    }
  }
  
  // Process each patient needing notification
  if (patients.length > 0) {
    processPatients(patients);
  }
  
  Logger.log(`Processed ${patients.length} patient(s) for certification notification.`);
}

/**
 * Extract patient data from a row
 */
function extractPatientData(row, rowNum) {
  const admissionDate = parseDate(row[CONFIG.COLUMNS.ADMISSION_DATE]);
  const patientName = row[CONFIG.COLUMNS.PATIENT_NAME];
  const mrNumber = row[CONFIG.COLUMNS.MR_NUMBER];
  
  if (!admissionDate || !patientName) {
    Logger.log(`Row ${rowNum}: Missing required data, skipping.`);
    return null;
  }
  
  const today = new Date();
  const daysSinceAdmission = calculateDaysBetween(admissionDate, today);
  
  return {
    rowNumber: rowNum,
    patientName: String(patientName).trim(),
    mrNumber: String(mrNumber).trim(),
    admissionDate: admissionDate,
    notifyDate: parseDate(row[CONFIG.COLUMNS.NOTIFY_DATE]),
    cDate1: parseDate(row[CONFIG.COLUMNS.CDATE_1]),
    cDate2: parseDate(row[CONFIG.COLUMNS.CDATE_2]),
    daysSinceAdmission: daysSinceAdmission,
    certPeriod: determineCertPeriod(daysSinceAdmission)
  };
}

/**
 * Determine certification period and required documents
 */
function determineCertPeriod(days) {
  if (days <= 90) {
    return {
      name: 'Initial (0-90 days)',
      documents: ['90DAY1', 'ATTEND_CERT', 'PATIENT_HISTORY']
    };
  } else if (days <= 180) {
    return {
      name: 'Second Period (91-180 days)',
      documents: ['90DAY2', 'PROGRESS_NOTE']
    };
  } else {
    return {
      name: 'Subsequent (180+ days)',
      documents: ['60DAY', 'PROGRESS_NOTE']
    };
  }
}

/**
 * Process all patients needing notification
 */
function processPatients(patients) {
  const allPreparedDocs = [];
  
  for (const patient of patients) {
    const preparedDocs = prepareDocuments(patient);
    allPreparedDocs.push({
      patient: patient,
      documents: preparedDocs
    });
  }
  
  // Send consolidated email
  sendNotificationEmail(allPreparedDocs);
}

/**
 * Prepare documents for a patient by copying templates and filling placeholders
 * Exports as PDF for email attachment
 */
function prepareDocuments(patient) {
  const preparedDocs = [];
  const placeholders = buildPlaceholders(patient);
  
  for (const docKey of patient.certPeriod.documents) {
    const templateId = CONFIG.DOC_TEMPLATES[docKey];
    
    try {
      // Create a copy of the template
      const templateFile = DriveApp.getFileById(templateId);
      const docName = `${docKey} - ${patient.patientName} - ${formatDate(new Date())}`;
      const copy = templateFile.makeCopy(docName);
      
      // Open the copy and replace placeholders
      const doc = DocumentApp.openById(copy.getId());
      replacePlaceholders(doc, placeholders);
      doc.saveAndClose();
      
      // Export as PDF
      const pdfBlob = convertDocToPdf(copy.getId(), docName);
      
      preparedDocs.push({
        name: docKey,
        docName: docName,
        url: copy.getUrl(),
        id: copy.getId(),
        pdfBlob: pdfBlob
      });
      
      Logger.log(`Prepared document: ${docName}`);
    } catch (e) {
      Logger.log(`Error preparing ${docKey} for ${patient.patientName}: ${e.message}`);
    }
  }
  
  return preparedDocs;
}

/**
 * Convert a Google Doc to PDF blob
 */
function convertDocToPdf(docId, fileName) {
  const url = `https://docs.google.com/document/d/${docId}/export?format=pdf`;
  const options = {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const pdfBlob = response.getBlob().setName(`${fileName}.pdf`);
  
  return pdfBlob;
}

/**
 * Build placeholder replacement map
 */
function buildPlaceholders(patient) {
  return {
    '{{Patient_Name}}': patient.patientName,
    '{{MR_Number}}': patient.mrNumber,
    '{{Doctor_Name}}': CONFIG.DOCTOR_NAME,
    '{{Admission_Date}}': formatDate(patient.admissionDate),
    '{{Notify_Date}}': formatDate(patient.notifyDate),
    '{{CDate_1}}': formatDate(patient.cDate1),
    '{{CDate_2}}': formatDate(patient.cDate2),
    '{{Today_Date}}': formatDate(new Date()),
    '{{Days_Since_Admission}}': patient.daysSinceAdmission.toString(),
    '{{Cert_Period}}': patient.certPeriod.name
  };
}

/**
 * Replace all placeholders in a document (body, headers, footers, and all nested elements)
 */
function replacePlaceholders(doc, placeholders) {
  const body = doc.getBody();
  const header = doc.getHeader();
  const footer = doc.getFooter();
  
  // Replace in body
  for (const [placeholder, value] of Object.entries(placeholders)) {
    body.replaceText(escapeRegex(placeholder), value || '');
  }
  
  // Replace in header if exists
  if (header) {
    replaceInElement(header, placeholders);
  }
  
  // Replace in footer if exists
  if (footer) {
    replaceInElement(footer, placeholders);
  }
  
  // Also search through all footnotes
  const footnotes = body.getFootnotes();
  if (footnotes) {
    for (const footnote of footnotes) {
      replaceInElement(footnote.getFootnoteContents(), placeholders);
    }
  }
}

/**
 * Recursively replace placeholders in any document element
 * Handles tables, text boxes, paragraphs, and nested elements
 */
function replaceInElement(element, placeholders) {
  if (!element) return;
  
  const type = element.getType();
  
  // First try direct replaceText if available
  if (typeof element.replaceText === 'function') {
    for (const [placeholder, value] of Object.entries(placeholders)) {
      try {
        element.replaceText(escapeRegex(placeholder), value || '');
      } catch (e) {
        // Some elements don't support replaceText, continue
      }
    }
  }
  
  // Handle different element types with children
  if (type === DocumentApp.ElementType.BODY_SECTION ||
      type === DocumentApp.ElementType.HEADER_SECTION ||
      type === DocumentApp.ElementType.FOOTER_SECTION ||
      type === DocumentApp.ElementType.DOCUMENT ||
      type === DocumentApp.ElementType.PARAGRAPH ||
      type === DocumentApp.ElementType.LIST_ITEM ||
      type === DocumentApp.ElementType.TABLE ||
      type === DocumentApp.ElementType.TABLE_ROW ||
      type === DocumentApp.ElementType.TABLE_CELL) {
    
    // Iterate through child elements
    if (typeof element.getNumChildren === 'function') {
      const numChildren = element.getNumChildren();
      for (let i = 0; i < numChildren; i++) {
        replaceInElement(element.getChild(i), placeholders);
      }
    }
  }
  
  // Handle text elements directly
  if (type === DocumentApp.ElementType.TEXT) {
    let text = element.getText();
    for (const [placeholder, value] of Object.entries(placeholders)) {
      if (text.includes(placeholder)) {
        text = text.replace(new RegExp(escapeRegex(placeholder), 'g'), value || '');
      }
    }
    try {
      element.setText(text);
    } catch (e) {
      // Read-only text, skip
    }
  }
}

/**
 * Send notification email to staff
 */
function sendNotificationEmail(allPreparedDocs) {
  const subject = `Patient Certification Alert - ${formatDate(new Date())} - Action Required`;
  const htmlBody = buildEmailHtml(allPreparedDocs);
  const plainBody = buildEmailPlain(allPreparedDocs);
  
  MailApp.sendEmail({
    to: CONFIG.EMAIL_LIST.join(','),
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody
  });
  
  Logger.log(`Notification email sent to: ${CONFIG.EMAIL_LIST.join(', ')}`);
}

/**
 * Build HTML email body
 */
function buildEmailHtml(allPreparedDocs) {
  let patientSections = '';
  
  for (const item of allPreparedDocs) {
    const patient = item.patient;
    const docs = item.documents;
    
    let docLinks = docs.map(d => 
      `<li><a href="${d.url}">${d.name}</a></li>`
    ).join('');
    
    patientSections += `
      <div style="background:#f9f9f9;border-left:4px solid #2c5282;padding:15px;margin:15px 0;border-radius:4px;">
        <h3 style="margin:0 0 10px 0;color:#2c5282;">${patient.patientName}</h3>
        <table style="width:100%;border-collapse:collapse;">
          <tr><td style="padding:4px 10px 4px 0;color:#666;">MR Number:</td><td style="padding:4px 0;"><strong>${patient.mrNumber}</strong></td></tr>
          <tr><td style="padding:4px 10px 4px 0;color:#666;">Admission Date:</td><td style="padding:4px 0;">${formatDate(patient.admissionDate)}</td></tr>
          <tr><td style="padding:4px 10px 4px 0;color:#666;">Days Since Admission:</td><td style="padding:4px 0;">${patient.daysSinceAdmission} days</td></tr>
          <tr><td style="padding:4px 10px 4px 0;color:#666;">Certification Period:</td><td style="padding:4px 0;"><strong>${patient.certPeriod.name}</strong></td></tr>
          <tr><td style="padding:4px 10px 4px 0;color:#666;">Attending Physician:</td><td style="padding:4px 0;">${CONFIG.DOCTOR_NAME}</td></tr>
        </table>
        <p style="margin:15px 0 5px 0;font-weight:bold;color:#333;">Prepared Documents:</p>
        <ul style="margin:5px 0;padding-left:20px;">${docLinks}</ul>
      </div>
    `;
  }
  
  return `
    <!DOCTYPE html>
    <html>
    <head><meta charset="UTF-8"></head>
    <body style="font-family:Arial,sans-serif;line-height:1.6;color:#333;max-width:700px;margin:0 auto;">
      <div style="background:#2c5282;color:white;padding:20px;text-align:center;">
        <h1 style="margin:0;font-size:24px;">Parrish Health Systems of Ohio</h1>
        <p style="margin:5px 0 0 0;opacity:0.9;">Patient Certification Notification</p>
      </div>
      
      <div style="padding:20px;">
        <p>Dear Team,</p>
        <p>This is an automated notification regarding patient certification(s) requiring attention. 
           The following patient(s) have upcoming certification deadlines and the required documents 
           have been prepared for your review.</p>
        
        <h2 style="color:#2c5282;border-bottom:2px solid #2c5282;padding-bottom:5px;">
          Patients Requiring Attention (${allPreparedDocs.length})
        </h2>
        
        ${patientSections}
        
        <div style="background:#fff3cd;border:1px solid #ffc107;padding:15px;margin:20px 0;border-radius:4px;">
          <strong>⚠️ Action Required:</strong> Please review the prepared documents and complete 
          the certification process in a timely manner to ensure continuity of care.
        </div>
        
        <p>If you have questions or need assistance, please contact your supervisor.</p>
        
        <p style="margin-top:30px;">
          Best regards,<br>
          <strong>Parrish Health Systems of Ohio</strong><br>
          <em>Automated Certification Management System</em>
        </p>
      </div>
      
      <div style="background:#f1f1f1;padding:15px;text-align:center;font-size:12px;color:#666;">
        This is an automated message from Parrish Health Systems. Please do not reply directly to this email.
      </div>
    </body>
    </html>
  `;
}

/**
 * Build plain text email body
 */
function buildEmailPlain(allPreparedDocs) {
  let content = `PARRISH HEALTH SYSTEMS OF OHIO
Patient Certification Notification
${formatDate(new Date())}

Dear Team,

This is an automated notification regarding patient certification(s) requiring attention.

PATIENTS REQUIRING ATTENTION (${allPreparedDocs.length})
${'='.repeat(50)}

`;

  for (const item of allPreparedDocs) {
    const patient = item.patient;
    const docs = item.documents;
    
    content += `
Patient: ${patient.patientName}
MR Number: ${patient.mrNumber}
Admission Date: ${formatDate(patient.admissionDate)}
Days Since Admission: ${patient.daysSinceAdmission} days
Certification Period: ${patient.certPeriod.name}
Attending Physician: ${CONFIG.DOCTOR_NAME}

Prepared Documents:
${docs.map(d => `  - ${d.name}: ${d.url}`).join('\n')}

${'-'.repeat(50)}
`;
  }

  content += `
ACTION REQUIRED: Please review the prepared documents and complete 
the certification process in a timely manner.

Best regards,
Parrish Health Systems of Ohio
Automated Certification Management System
`;

  return content;
}

// ============ UTILITY FUNCTIONS ============

/**
 * Parse various date formats from sheet cells
 */
function parseDate(value) {
  if (!value) return null;
  
  // Already a Date object
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }
  
  // Number (Excel serial date)
  if (typeof value === 'number') {
    // Google Sheets uses Dec 30, 1899 as epoch
    const date = new Date(Date.UTC(1899, 11, 30 + value));
    return isNaN(date.getTime()) ? null : date;
  }
  
  // String - try parsing
  if (typeof value === 'string') {
    const parsed = new Date(value);
    return isNaN(parsed.getTime()) ? null : parsed;
  }
  
  return null;
}

/**
 * Normalize date to midnight for comparison
 */
function normalizeDate(date) {
  if (!date) return null;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

/**
 * Calculate days between two dates
 */
function calculateDaysBetween(startDate, endDate) {
  const start = normalizeDate(startDate);
  const end = normalizeDate(endDate);
  const diffTime = Math.abs(end - start);
  return Math.floor(diffTime / (1000 * 60 * 60 * 24));
}

/**
 * Format date as MM/DD/YYYY
 */
function formatDate(date) {
  if (!date) return 'N/A';
  const d = new Date(date);
  if (isNaN(d.getTime())) return 'N/A';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd/yyyy');
}

/**
 * Escape special regex characters in placeholder strings
 */
function escapeRegex(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ============ TRIGGER MANAGEMENT ============

/**
 * Set up the daily 9am trigger
 * Run this function once to create the trigger
 */
function createDailyTrigger() {
  // Remove existing triggers first
  deleteTriggers();
  
  // Create new daily trigger at 9am
  ScriptApp.newTrigger('checkCertificationNotifications')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();
  
  Logger.log('Daily trigger created for 9:00 AM');
}

/**
 * Delete all existing triggers for this script
 */
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'checkCertificationNotifications') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  Logger.log('Existing triggers deleted');
}

/**
 * Set up the weekly Monday trigger for summary emails
 * Run this function once to create the trigger
 */
function createWeeklyTrigger() {
  // Remove existing weekly triggers first
  deleteWeeklyTriggers();
  
  // Create new weekly trigger for Monday at 8am
  ScriptApp.newTrigger('sendWeeklySummary')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
  
  Logger.log('Weekly trigger created for Monday at 8:00 AM');
}

/**
 * Delete existing weekly summary triggers
 */
function deleteWeeklyTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'sendWeeklySummary') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  Logger.log('Existing weekly triggers deleted');
}

/**
 * Weekly Summary - Find all patients with notify dates this month
 * and send a consolidated summary email with documents
 */
function sendWeeklySummary() {
  Logger.log('Starting weekly summary...');
  
  const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  const currentMonth = today.getMonth();
  const currentYear = today.getFullYear();
  
  // Get end of current month
  const endOfMonth = new Date(currentYear, currentMonth + 1, 0);
  
  const upcomingPatients = [];
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const notifyDate = parseDate(row[CONFIG.COLUMNS.NOTIFY_DATE]);
    
    if (!notifyDate) continue;
    
    // Check if notify date is within this month (from today to end of month)
    if (notifyDate >= normalizeDate(today) && notifyDate <= endOfMonth) {
      const patientData = extractPatientData(row, i + 1);
      if (patientData) {
        upcomingPatients.push(patientData);
      }
    }
  }
  
  if (upcomingPatients.length === 0) {
    Logger.log('No upcoming certifications this month.');
    sendNoUpcomingEmail();
    return;
  }
  
  // Sort by notify date
  upcomingPatients.sort((a, b) => a.notifyDate - b.notifyDate);
  
  // Prepare documents for all patients
  const allPreparedDocs = [];
  for (const patient of upcomingPatients) {
    const preparedDocs = prepareDocuments(patient);
    allPreparedDocs.push({
      patient: patient,
      documents: preparedDocs
    });
  }
  
  // Send weekly summary email
  sendWeeklySummaryEmail(allPreparedDocs, endOfMonth);
  
  Logger.log(`Weekly summary complete. Processed ${upcomingPatients.length} patient(s).`);
}

/**
 * Send email when no upcoming certifications
 */
function sendNoUpcomingEmail() {
  const today = new Date();
  const monthName = today.toLocaleString('default', { month: 'long' });
  
  const subject = `Weekly Certification Summary - ${monthName} ${today.getFullYear()} - No Upcoming`;
  
  const htmlBody = `
    <!DOCTYPE html>
    <html>
    <head><meta charset="UTF-8"></head>
    <body style="font-family:Arial,sans-serif;line-height:1.6;color:#333;max-width:700px;margin:0 auto;">
      <div style="background:#2c5282;color:white;padding:20px;text-align:center;">
        <h1 style="margin:0;font-size:24px;">Parrish Health Systems of Ohio</h1>
        <p style="margin:5px 0 0 0;opacity:0.9;">Weekly Certification Summary</p>
      </div>
      
      <div style="padding:20px;">
        <p>Dear Team,</p>
        <p>This is your weekly certification summary for <strong>${monthName} ${today.getFullYear()}</strong>.</p>
        
        <div style="background:#d4edda;border:1px solid #28a745;padding:15px;margin:20px 0;border-radius:4px;">
          <strong>✓ All Clear:</strong> There are no patient certifications due for the remainder of this month.
        </div>
        
        <p style="margin-top:30px;">
          Best regards,<br>
          <strong>Parrish Health Systems of Ohio</strong><br>
          <em>Automated Certification Management System</em>
        </p>
      </div>
      
      <div style="background:#f1f1f1;padding:15px;text-align:center;font-size:12px;color:#666;">
        Weekly summary generated on ${formatDate(today)}
      </div>
    </body>
    </html>
  `;
  
  MailApp.sendEmail({
    to: CONFIG.EMAIL_LIST.join(','),
    subject: subject,
    body: `Weekly Certification Summary - No upcoming certifications for ${monthName}.`,
    htmlBody: htmlBody
  });
  
  Logger.log('No upcoming certifications email sent.');
}

/**
 * Send weekly summary email with all upcoming patients and documents
 */
function sendWeeklySummaryEmail(allPreparedDocs, endOfMonth) {
  const today = new Date();
  const monthName = today.toLocaleString('default', { month: 'long' });
  
  const subject = `Weekly Certification Summary - ${monthName} ${today.getFullYear()} - ${allPreparedDocs.length} Patient(s) Upcoming`;
  
  let patientSections = '';
  const allAttachments = [];
  
  for (const item of allPreparedDocs) {
    const patient = item.patient;
    const docs = item.documents;
    
    // Calculate days until notify date
    const daysUntil = calculateDaysBetween(today, patient.notifyDate);
    const urgencyColor = daysUntil <= 7 ? '#dc3545' : (daysUntil <= 14 ? '#ffc107' : '#28a745');
    const urgencyText = daysUntil <= 7 ? 'This Week!' : (daysUntil <= 14 ? 'Soon' : 'Upcoming');
    
    let docLinks = docs.map(d => 
      `<li><strong>${d.name}</strong>: <a href="${d.url}">View in Google Docs</a> | <em>PDF attached</em></li>`
    ).join('');
    
    // Collect attachments
    for (const doc of docs) {
      if (doc.pdfBlob) {
        allAttachments.push(doc.pdfBlob);
      }
    }
    
    patientSections += `
      <div style="background:#f9f9f9;border-left:4px solid ${urgencyColor};padding:15px;margin:15px 0;border-radius:4px;">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
          <h3 style="margin:0;color:#2c5282;">${patient.patientName}</h3>
          <span style="background:${urgencyColor};color:white;padding:3px 10px;border-radius:12px;font-size:12px;font-weight:bold;">${urgencyText}</span>
        </div>
        <table style="width:100%;border-collapse:collapse;">
          <tr><td style="padding:4px 10px 4px 0;color:#666;">MR Number:</td><td style="padding:4px 0;"><strong>${patient.mrNumber}</strong></td></tr>
          <tr><td style="padding:4px 10px 4px 0;color:#666;">Notify Date:</td><td style="padding:4px 0;"><strong>${formatDate(patient.notifyDate)}</strong> (${daysUntil} days)</td></tr>
          <tr><td style="padding:4px 10px 4px 0;color:#666;">Admission Date:</td><td style="padding:4px 0;">${formatDate(patient.admissionDate)}</td></tr>
          <tr><td style="padding:4px 10px 4px 0;color:#666;">Days Since Admission:</td><td style="padding:4px 0;">${patient.daysSinceAdmission} days</td></tr>
          <tr><td style="padding:4px 10px 4px 0;color:#666;">Certification Period:</td><td style="padding:4px 0;"><strong>${patient.certPeriod.name}</strong></td></tr>
          <tr><td style="padding:4px 10px 4px 0;color:#666;">Attending Physician:</td><td style="padding:4px 0;">${CONFIG.DOCTOR_NAME}</td></tr>
        </table>
        <p style="margin:15px 0 5px 0;font-weight:bold;color:#333;">Prepared Documents:</p>
        <ul style="margin:5px 0;padding-left:20px;">${docLinks}</ul>
      </div>
    `;
  }
  
  const htmlBody = `
    <!DOCTYPE html>
    <html>
    <head><meta charset="UTF-8"></head>
    <body style="font-family:Arial,sans-serif;line-height:1.6;color:#333;max-width:700px;margin:0 auto;">
      <div style="background:#2c5282;color:white;padding:20px;text-align:center;">
        <h1 style="margin:0;font-size:24px;">Parrish Health Systems of Ohio</h1>
        <p style="margin:5px 0 0 0;opacity:0.9;">Weekly Certification Summary</p>
      </div>
      
      <div style="padding:20px;">
        <p>Dear Team,</p>
        <p>This is your weekly certification summary. The following <strong>${allPreparedDocs.length} patient(s)</strong> 
           have certifications due between now and <strong>${formatDate(endOfMonth)}</strong>.</p>
        
        <div style="background:#e7f3ff;border:1px solid #2c5282;padding:15px;margin:15px 0;border-radius:4px;">
          <strong>📅 Summary:</strong>
          <ul style="margin:10px 0 0 0;">
            <li><span style="color:#dc3545;font-weight:bold;">Red</span> = Due this week</li>
            <li><span style="color:#ffc107;font-weight:bold;">Yellow</span> = Due within 2 weeks</li>
            <li><span style="color:#28a745;font-weight:bold;">Green</span> = Due later this month</li>
          </ul>
        </div>
        
        <h2 style="color:#2c5282;border-bottom:2px solid #2c5282;padding-bottom:5px;">
          Upcoming Certifications (${allPreparedDocs.length})
        </h2>
        
        ${patientSections}
        
        <div style="background:#fff3cd;border:1px solid #ffc107;padding:15px;margin:20px 0;border-radius:4px;">
          <strong>📎 Attachments:</strong> All prepared documents are attached to this email as PDFs (${allAttachments.length} total).
          You can also click the Google Docs links above to edit if needed.
        </div>
        
        <p style="margin-top:30px;">
          Best regards,<br>
          <strong>Parrish Health Systems of Ohio</strong><br>
          <em>Automated Certification Management System</em>
        </p>
      </div>
      
      <div style="background:#f1f1f1;padding:15px;text-align:center;font-size:12px;color:#666;">
        Weekly summary generated on ${formatDate(today)} | Next summary: ${getNextMonday()}
      </div>
    </body>
    </html>
  `;
  
  const plainBody = buildWeeklySummaryPlain(allPreparedDocs, endOfMonth);
  
  MailApp.sendEmail({
    to: CONFIG.EMAIL_LIST.join(','),
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody,
    attachments: allAttachments
  });
  
  Logger.log(`Weekly summary email sent with ${allAttachments.length} attachments.`);
}

/**
 * Build plain text version of weekly summary
 */
function buildWeeklySummaryPlain(allPreparedDocs, endOfMonth) {
  const today = new Date();
  const monthName = today.toLocaleString('default', { month: 'long' });
  
  let content = `PARRISH HEALTH SYSTEMS OF OHIO
Weekly Certification Summary
${formatDate(today)}

Dear Team,

This is your weekly certification summary. The following ${allPreparedDocs.length} patient(s) 
have certifications due between now and ${formatDate(endOfMonth)}.

UPCOMING CERTIFICATIONS
${'='.repeat(50)}

`;

  for (const item of allPreparedDocs) {
    const patient = item.patient;
    const docs = item.documents;
    const daysUntil = calculateDaysBetween(today, patient.notifyDate);
    
    content += `
Patient: ${patient.patientName}
MR Number: ${patient.mrNumber}
Notify Date: ${formatDate(patient.notifyDate)} (${daysUntil} days)
Admission Date: ${formatDate(patient.admissionDate)}
Days Since Admission: ${patient.daysSinceAdmission} days
Certification Period: ${patient.certPeriod.name}
Attending Physician: ${CONFIG.DOCTOR_NAME}

Prepared Documents:
${docs.map(d => `  - ${d.name}: ${d.url} (PDF attached)`).join('\n')}

${'-'.repeat(50)}
`;
  }

  content += `
All prepared documents are attached as PDFs.

Best regards,
Parrish Health Systems of Ohio
Automated Certification Management System

Next summary: ${getNextMonday()}
`;

  return content;
}

/**
 * Get the date of next Monday
 */
function getNextMonday() {
  const today = new Date();
  const dayOfWeek = today.getDay();
  const daysUntilMonday = dayOfWeek === 0 ? 1 : (8 - dayOfWeek);
  const nextMonday = new Date(today);
  nextMonday.setDate(today.getDate() + daysUntilMonday);
  return formatDate(nextMonday);
}

/**
 * Test function for weekly summary - run manually to test
 */
function testWeeklySummary() {
  Logger.log('Starting weekly summary test...');
  sendWeeklySummary();
  Logger.log('Weekly summary test complete. Check your email.');
}

/**
 * Manual test function - run to test without waiting for trigger
 */
function testRun() {
  Logger.log('Starting manual test run...');
  checkCertificationNotifications();
  Logger.log('Test run complete. Check logs for details.');
}

/**
 * Test function to generate ALL 6 documents with sample patient data
 * and send them via email as PDF attachments for review
 */
function testAllDocuments() {
  Logger.log('Starting full document test...');
  
  // Create test patient data
  const testPatient = {
    rowNumber: 0,
    patientName: 'John Test Smith',
    mrNumber: 'MR-TEST-12345',
    admissionDate: new Date(2025, 8, 15), // Sept 15, 2025
    notifyDate: new Date(),
    cDate1: new Date(2025, 11, 14), // Dec 14, 2025
    cDate2: new Date(2026, 2, 14),  // Mar 14, 2026
    daysSinceAdmission: 86,
    certPeriod: {
      name: 'TEST - All Documents',
      documents: ['60DAY', '90DAY1', '90DAY2', 'ATTEND_CERT', 'PROGRESS_NOTE', 'PATIENT_HISTORY']
    }
  };
  
  const placeholders = buildPlaceholders(testPatient);
  const preparedDocs = [];
  
  // Generate all 6 documents
  for (const docKey of testPatient.certPeriod.documents) {
    const templateId = CONFIG.DOC_TEMPLATES[docKey];
    
    try {
      const templateFile = DriveApp.getFileById(templateId);
      const docName = `TEST - ${docKey} - ${testPatient.patientName} - ${formatDate(new Date())}`;
      const copy = templateFile.makeCopy(docName);
      
      const doc = DocumentApp.openById(copy.getId());
      replacePlaceholders(doc, placeholders);
      doc.saveAndClose();
      
      // Export as PDF
      const pdfBlob = convertDocToPdf(copy.getId(), docName);
      
      preparedDocs.push({
        name: docKey,
        docName: docName,
        url: copy.getUrl(),
        id: copy.getId(),
        pdfBlob: pdfBlob
      });
      
      Logger.log(`✓ Prepared: ${docName}`);
    } catch (e) {
      Logger.log(`✗ Error with ${docKey}: ${e.message}`);
    }
  }
  
  // Send test email with PDF attachments
  sendTestEmail(testPatient, preparedDocs);
  
  Logger.log(`\nTest complete! Generated ${preparedDocs.length} documents.`);
  Logger.log('Check your email for the test notification with PDF attachments.');
}

/**
 * Send test email with all documents as PDF attachments
 */
function sendTestEmail(patient, docs) {
  const subject = `[TEST] All Documents Preview - ${formatDate(new Date())}`;
  
  let docList = docs.map(d => 
    `<li><strong>${d.name}</strong>: <a href="${d.url}">View in Google Docs</a> | <em>PDF attached</em></li>`
  ).join('');
  
  const htmlBody = `
    <!DOCTYPE html>
    <html>
    <head><meta charset="UTF-8"></head>
    <body style="font-family:Arial,sans-serif;line-height:1.6;color:#333;max-width:700px;margin:0 auto;">
      <div style="background:#d97706;color:white;padding:20px;text-align:center;">
        <h1 style="margin:0;font-size:24px;">⚠️ TEST EMAIL - Parrish Health Systems</h1>
        <p style="margin:5px 0 0 0;">Document Template Preview</p>
      </div>
      
      <div style="padding:20px;">
        <div style="background:#fef3c7;border:1px solid #d97706;padding:15px;margin-bottom:20px;border-radius:4px;">
          <strong>This is a TEST email.</strong> All 6 document templates have been generated 
          with sample patient data. View them in Google Docs or open the <strong>attached PDFs</strong>.
        </div>
        
        <h2 style="color:#2c5282;">Test Patient Data Used</h2>
        <table style="width:100%;border-collapse:collapse;margin-bottom:20px;">
          <tr><td style="padding:8px;border-bottom:1px solid #eee;color:#666;width:180px;">Patient Name:</td>
              <td style="padding:8px;border-bottom:1px solid #eee;"><strong>${patient.patientName}</strong></td></tr>
          <tr><td style="padding:8px;border-bottom:1px solid #eee;color:#666;">MR Number:</td>
              <td style="padding:8px;border-bottom:1px solid #eee;"><strong>${patient.mrNumber}</strong></td></tr>
          <tr><td style="padding:8px;border-bottom:1px solid #eee;color:#666;">Doctor Name:</td>
              <td style="padding:8px;border-bottom:1px solid #eee;">${CONFIG.DOCTOR_NAME}</td></tr>
          <tr><td style="padding:8px;border-bottom:1px solid #eee;color:#666;">Admission Date:</td>
              <td style="padding:8px;border-bottom:1px solid #eee;">${formatDate(patient.admissionDate)}</td></tr>
          <tr><td style="padding:8px;border-bottom:1px solid #eee;color:#666;">Notify Date:</td>
              <td style="padding:8px;border-bottom:1px solid #eee;">${formatDate(patient.notifyDate)}</td></tr>
          <tr><td style="padding:8px;border-bottom:1px solid #eee;color:#666;">CDate 1:</td>
              <td style="padding:8px;border-bottom:1px solid #eee;">${formatDate(patient.cDate1)}</td></tr>
          <tr><td style="padding:8px;border-bottom:1px solid #eee;color:#666;">CDate 2:</td>
              <td style="padding:8px;border-bottom:1px solid #eee;">${formatDate(patient.cDate2)}</td></tr>
          <tr><td style="padding:8px;border-bottom:1px solid #eee;color:#666;">Days Since Admission:</td>
              <td style="padding:8px;border-bottom:1px solid #eee;">${patient.daysSinceAdmission} days</td></tr>
        </table>
        
        <h2 style="color:#2c5282;">📎 Documents (${docs.length})</h2>
        <p>Click to view in Google Docs, or open the attached PDFs:</p>
        <ul style="line-height:2;">${docList}</ul>
        
        <div style="background:#e0f2fe;border:1px solid #0284c7;padding:15px;margin:20px 0;border-radius:4px;">
          <strong>📋 Review Checklist:</strong>
          <ul style="margin:10px 0 0 0;">
            <li>Open via Google Docs link to edit, or view attached PDFs</li>
            <li>Verify all {{placeholders}} were replaced correctly</li>
            <li>Check formatting in headers, footers, and body</li>
            <li>Delete test documents from Drive when done</li>
          </ul>
        </div>
      </div>
      
      <div style="background:#f1f1f1;padding:15px;text-align:center;font-size:12px;color:#666;">
        Test generated on ${formatDate(new Date())} | Parrish Health Systems of Ohio
      </div>
    </body>
    </html>
  `;
  
  const plainBody = `TEST EMAIL - Parrish Health Systems
Document Template Preview

This is a TEST email with all 6 documents generated using sample patient data.
View them in Google Docs or open the attached PDFs.

TEST PATIENT DATA:
- Patient Name: ${patient.patientName}
- MR Number: ${patient.mrNumber}
- Doctor Name: ${CONFIG.DOCTOR_NAME}
- Admission Date: ${formatDate(patient.admissionDate)}
- Days Since Admission: ${patient.daysSinceAdmission}

DOCUMENTS (View in Google Docs or open attached PDFs):
${docs.map(d => `- ${d.name}: ${d.url} (PDF attached)`).join('\n')}

Please review each PDF to verify placeholders were replaced correctly.
Delete test documents from Drive when done.
`;

  // Collect PDF attachments
  const attachments = docs
    .filter(d => d.pdfBlob)
    .map(d => d.pdfBlob);

  MailApp.sendEmail({
    to: CONFIG.EMAIL_LIST.join(','),
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody,
    attachments: attachments
  });
  
  Logger.log(`Test email sent to: ${CONFIG.EMAIL_LIST.join(', ')} with ${attachments.length} PDF attachments`);
}
