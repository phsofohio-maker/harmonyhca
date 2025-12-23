// =================================================================
// SCRIPT CONFIGURATION
// =================================================================

// --- ⬇️ EDIT THESE VALUES ⬇️ ---

/**
 * The email address to send the daily report to.
 * For multiple recipients, separate addresses with a comma.
 * e.g., 'email1@example.com, email2@example.com'
 */
const RECIPIENT_EMAIL = 'kobet@parrishhealthsystems.org,reneesha@parrishhealthsystems.org,Joyceboateng370@yahoo.com,Kaylapudvan@gmail.com,nassumpta@hotmail.com,olumideo@parrishhealthsystems.org,miarac@parrishhealthsystems.org,tajuanna@parrishhealthsystems.org,lewisgena291@gmail.com,quinisha.glaspie@gmail.com,kevo3415@yahoo.com';

/**
 * The name of the sheet in your Google Spreadsheet that contains patient data.
 */
const SHEET_NAME = 'HOPE/HUV';

// --- ⬆️ NO MORE EDITS NEEDED BELOW THIS LINE ⬆️ ---


// =================================================================
// Main Functions
// =================================================================

/**
 * Calculates and returns an object containing the HUV1 and HUV2 date windows.
 * @param {Date | string} startDate The start of care date.
 * @returns {object} An object with four Date objects: huv1Start, huv1End, huv2Start, huv2End.
 */
function getHUVWindows(startDate) {
  const socDate = new Date(startDate);

  // Create new date objects for each calculation
  const huv1Start = new Date(socDate);
  const huv1End = new Date(socDate);
  const huv2Start = new Date(socDate);
  const huv2End = new Date(socDate);

  // Calculate date windows
  huv1Start.setDate(socDate.getDate() + 5);
  huv1End.setDate(socDate.getDate() + 14);
  huv2Start.setDate(socDate.getDate() + 15);
  huv2End.setDate(socDate.getDate() + 28);

  return {
    huv1Start,
    huv1End,
    huv2Start,
    huv2End
  };
}

/**
 * Reads patient data from a Google Sheet, determines HUV status, and sends a summary email.
 * This function is intended to be run daily via a time-based trigger.
 */
function sendDailyHUVReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    console.error(`Error: Sheet named "${SHEET_NAME}" not found.`);
    return; // Exit if the sheet doesn't exist
  }

  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // Remove header row and store it

  const today = new Date();
  today.setHours(0, 0, 0, 0); // Normalize date to midnight for accurate comparisons

  let patientRows = [];

  for (const row of data) {
    const patientName = row[0];
    const socDate = new Date(row[1]);
    const isHuv1Complete = row[2]; // Assumes TRUE/FALSE from a checkbox or cell value
    const isHuv2Complete = row[3];

    // Skip empty rows or rows with an invalid date
    if (!patientName || socDate.toString() === 'Invalid Date') {
      continue;
    }

    const windows = getHUVWindows(socDate);
    const huv1Status = getStatus(today, windows.huv1Start, windows.huv1End, isHuv1Complete);
    const huv2Status = getStatus(today, windows.huv2Start, windows.huv2End, isHuv2Complete);

    patientRows.push({
      name: patientName,
      huv1Window: `(${formatDate(windows.huv1Start)} - ${formatDate(windows.huv1End)})`,
      huv1Status: huv1Status,
      huv2Window: `(${formatDate(windows.huv2Start)} - ${formatDate(windows.huv2End)})`,
      huv2Status: huv2Status
    });
  }

  // If no patients, don't send an email
  if (patientRows.length === 0) {
    console.log("No patient data found. Email not sent.");
    return;
  }
  
  // Sort patients so those needing action are at the top
  patientRows.sort((a, b) => {
    const priority = { "❌": 0, "❗": 1, "✅": 2, "🗓️": 3 };
    const statusA = a.huv1Status.charAt(0);
    const statusB = b.huv1Status.charAt(0);
    return priority[statusA] - priority[statusB];
  });


  const emailBody = createHtmlEmailBody(patientRows);
  const subject = `Daily HOPE Update Visit (HUV) Report - ${new Date().toLocaleDateString()}`;

  MailApp.sendEmail({
    to: RECIPIENT_EMAIL,
    subject: subject,
    htmlBody: emailBody
  });
}

// =================================================================
// Helper Functions
// =================================================================

/**
 * Determines the status of a visit based on today's date and its window.
 * @private
 */
function getStatus(today, startDate, endDate, isComplete) {
  if (isComplete === true) return '✅ Complete';
  if (today > endDate) return '❌ OVERDUE';
  if (today >= startDate && today <= endDate) return '❗ ACTION NEEDED';
  return '🗓️ Upcoming';
}

/**
 * Formats a Date object into a simple "MM/DD" string.
 * @private
 */
function formatDate(date) {
  return `${date.getMonth() + 1}/${date.getDate()}`;
}

/**
 * Creates an HTML-formatted table for the email body.
 * @private
 */
function createHtmlEmailBody(patientRows) {
  let rowsHtml = '';
  for (const patient of patientRows) {
    rowsHtml += `
      <tr>
        <td style="padding: 8px; border-bottom: 1px solid #ddd;">${patient.name}</td>
        <td style="padding: 8px; border-bottom: 1px solid #ddd;">${patient.huv1Status}</td>
        <td style="padding: 8px; border-bottom: 1px solid #ddd;">${patient.huv1Window}</td>
        <td style="padding: 8px; border-bottom: 1px solid #ddd;">${patient.huv2Status}</td>
        <td style="padding: 8px; border-bottom: 1px solid #ddd;">${patient.huv2Window}</td>
      </tr>
    `;
  }

  return `
    <body style="font-family: Arial, sans-serif; color: #333;">
      <h2>HOPE Update Visit (HUV) Daily Status</h2>
      <p>This report shows all active patients and the status of their HUV1 and HUV2 visits. Patients requiring attention are listed first.</p>
      <table style="width: 100%; border-collapse: collapse; text-align: left;">
        <thead>
          <tr style="background-color: #f2f2f2;">
            <th style="padding: 8px;">Patient Name</th>
            <th style="padding: 8px;">HUV1 Status</th>
            <th style="padding: 8px;">HUV1 Window</th>
            <th style="padding: 8px;">HUV2 Status</th>
            <th style="padding: 8px;">HUV2 Window</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml}
        </tbody>
      </table>
      <p style="font-size: 12px; color: #888; margin-top: 20px;">
        Status Legend: <br>
        ❗ ACTION NEEDED: Visit window is currently active. <br>
        ❌ OVERDUE: Visit window has passed and is not marked complete. <br>
        ✅ Complete: Visit is marked as complete. <br>
        🗓️ Upcoming: Visit window is in the future.
      </p>
    </body>
  `;
}

/**
 * Test function to log the return value of getHUVWindows.
 */
function testHUVCalculation() {
  const testDate = new Date('2/6/2024'); [cite_start]// Using the date from the PDF [cite: 1]
  const windows = getHUVWindows(testDate);

  // The function now returns a usable object
  console.log(windows);

  // You can access each date individually
  console.log(`HUV1 starts on: ${windows.huv1Start.toLocaleDateString()}`);
}