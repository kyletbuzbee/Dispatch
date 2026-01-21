/**
 * K&L DISPATCH MANAGER - DISPATCH & EMAILS
 * FEATURES:
 * 1. Add Jobs from Sidebar (inserts at top).
 * 2. Auto-Stamping: Marks "Date Completed" when you select "Complete".
 * 3. Custom Emails: Grouped by Yard, with Pick/Drop summaries.
 * 4. Archiving: Moves old completed jobs to a separate tab to keep the board clean.
 */

const CONFIG = {
  SHEET_NAME: 'Dispatch Sheet',
  ARCHIVE_PREFIX: 'Archive_',
  EMAIL_RECIPIENTS: [
    'mineola@kl-recycling.com',
    'ccrow@kl-recycling.com',
    'awells@kl-recycling.com',
    'jacksonville@kl-recycling.com',
    'houstoncounty@kl-recycling.com',
    'dfletcher@kl-recycling.com',
    'rhood@kl-recycling.com',
    'andersoncounty@kl-recycling.com',
    'kbritton@kl-recycling.com',
    'MARKETING@kl-recycling.com',
    'madfos@kl-recycling.com',
    'nacogdoches@kl-recycling.com',
    'premier@kl-recycling.com',
    'mwells@kl-recycling.com',
    'rsawler@kl-recycling.com'
  ],

  // Column Mapping (0-based Index: A=0, B=1, etc.)
  COLS: {
    DATE: 0,           // A
    COMPANY: 1,        // B
    CITY: 2,           // C
    ACTION: 3,         // D
    YARD: 4,           // E
    BOX: 5,            // F
    QTY: 6,            // G
    STATUS: 7,         // H
    DATE_COMPLETED: 8, // I
    NOTES: 9,          // J
    HELPER_DATE: 10    // K <--- NEW HELPER COLUMN
  }
};

// ==========================================
// 1. MENU & TRIGGERS
// ==========================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚛 Dispatch Tools')
    .addItem('📱 Open Sidebar', 'showSidebar')
    .addSeparator()
    .addItem('📅 Send Next Day Schedule', 'sendNextDaySchedule')
    .addItem('✅ Send Previous Day Report', 'sendPreviousDayReport')
    .addSeparator()
    .addItem('📦 Archive Completed Jobs', 'archiveCompletedRows')
    .addToUi();
}

function onEdit(e) {
  if (!e) return;
  const range = e.range;
  const sheet = range.getSheet();

  // Auto-Stamp Date when Status changes to "Complete"
  if (sheet.getName() === CONFIG.SHEET_NAME &&
      range.getColumn() === (CONFIG.COLS.STATUS + 1)) {

    const status = String(range.getValue()).trim();
    const row = range.getRow();

    if (status === 'Complete') {
      const today = new Date();
      // 1. Stamp Date Object (Column I / Index 8)
      sheet.getRange(row, CONFIG.COLS.DATE_COMPLETED + 1)
           .setValue(today)
           .setNumberFormat('MM/dd/yyyy');

      // 2. Stamp Helper Text String (Column K / Index 10) - THE KEY PART
      // We force a text string like "01/20/2026" that will never change timezone
      const dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy');
      sheet.getRange(row, CONFIG.COLS.HELPER_DATE + 1)
           .setValue(dateString)
           .setNumberFormat('@'); // Force Text Format

    } else {
      // Clear BOTH if status changed away from Complete
      sheet.getRange(row, CONFIG.COLS.DATE_COMPLETED + 1).clearContent();
      sheet.getRange(row, CONFIG.COLS.HELPER_DATE + 1).clearContent();
    }
  }
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('K&L Mobile')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function showSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('🚛 Dispatch')
      .setWidth(400)
  );
}

// ==========================================
// 2. CORE ACTIONS (ADD & ARCHIVE)
// ==========================================

function addNewJob(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.SHEET_NAME);

  // If no date selected, default to Today
  const dateVal = form.date ? new Date(form.date + 'T12:00:00') : new Date();

  const row = [
    dateVal,                  
    form.company,            
    form.city,               
    form.action,
    form.yard,
    form.boxSize,
    form.numBoxes || 1,
    'Pending',
    '',                       // Date Completed (empty)
    form.notes || '',        
    ''                        // Helper Date (empty, wait for onEdit to fill it)
  ];

  sh.insertRowBefore(2);
  sh.getRange(2, 1, 1, row.length).setValues([row]);
  sh.getRange(2, 1).setNumberFormat('MM/dd/yyyy'); // Format Date Received
  
  // OPTIONAL: Ensure Helper Col is Text format
  sh.getRange(2, CONFIG.COLS.HELPER_DATE + 1).setNumberFormat('@'); 

  return getDashboardStats();
}

function archiveCompletedRows(startDateStr, endDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const data = srcSheet.getDataRange().getValues();
  const headers = data.shift(); // Remove headers

  // 1. Setup Date Filter Boundaries
  let start = null;
  let end = null;
  
  if (startDateStr) {
    start = new Date(startDateStr);
    start.setHours(0, 0, 0, 0); // Start of that day
  }
  
  if (endDateStr) {
    end = new Date(endDateStr);
    end.setHours(23, 59, 59, 999); // End of that day
  }

  // 2. Find or Create Archive Sheet
  const year = new Date().getFullYear();
  const archiveName = CONFIG.ARCHIVE_PREFIX + year;
  let destSheet = ss.getSheetByName(archiveName);

  if (!destSheet) {
    destSheet = ss.insertSheet(archiveName);
    destSheet.appendRow(headers);
    destSheet.setFrozenRows(1);
  }

  const rowsToArchive = [];
  const rowsToKeep = [headers];

  // 3. Filter Loop
  data.forEach(row => {
    const status = String(row[CONFIG.COLS.STATUS]).trim();
    let shouldArchive = false;

    if (status === 'Complete') {
      // If no dates provided, archive ALL completed
      if (!start && !end) {
        shouldArchive = true;
      } else {
        // If dates provided, check the row date
        const rowDateVal = row[CONFIG.COLS.DATE_COMPLETED];
        
        if (rowDateVal instanceof Date) {
          // Check Start Boundary
          const afterStart = !start || (rowDateVal >= start);
          // Check End Boundary
          const beforeEnd = !end || (rowDateVal <= end);
          
          if (afterStart && beforeEnd) {
            shouldArchive = true;
          }
        }
      }
    }

    if (shouldArchive) {
      rowsToArchive.push(row);
    } else {
      rowsToKeep.push(row);
    }
  });

  if (rowsToArchive.length === 0) return 'No matching jobs found to archive.';

  // 4. Move to Archive
  const lastRow = destSheet.getLastRow();
  destSheet
    .getRange(lastRow + 1, 1, rowsToArchive.length, rowsToArchive[0].length)
    .setValues(rowsToArchive);

  // Formatting Archive
  destSheet.getRange(lastRow + 1, 1, rowsToArchive.length, 1).setNumberFormat('MM/dd/yyyy');
  destSheet.getRange(lastRow + 1, CONFIG.COLS.DATE_COMPLETED + 1, rowsToArchive.length, 1).setNumberFormat('MM/dd/yyyy');

  // 5. Update Main Sheet
  srcSheet.clearContents();
  srcSheet
    .getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length)
    .setValues(rowsToKeep);

  return `Archived ${rowsToArchive.length} jobs to ${archiveName}.`;
}

// ==========================================
// 3. EMAIL REPORTS
// ==========================================

// --- COMPLETED REPORT (PREVIOUS BUSINESS DAY) ---
function sendPreviousDayReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  // 1. Calculate the Target Date String (e.g. "01/19/2026")
  const targetDate = getTargetDate('prev');
  const targetStr = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');
  
  Logger.log(`Target Date: ${targetStr}`); 

  const data = sh.getDataRange().getValues();
  data.shift(); // remove headers

  const jobs = [];

  data.forEach((row, index) => {
    // 2. Check Status
    const status = String(row[CONFIG.COLS.STATUS]).trim();
    if (status !== 'Complete') return;

    // 3. READ THE HELPER COLUMN (Col K / Index 10)
    // The Sheet formula converts this to strict text "MM/dd/yyyy"
    const dateStr = String(row[CONFIG.COLS.HELPER_DATE] || '').trim();

    // 4. Simple Text Comparison (Reliable)
    if (dateStr === targetStr) {
      Logger.log(`Row ${index + 2}: MATCH! (${dateStr})`);
      jobs.push(mapRowToObject(row));
    }
  });

  const pendingCount = data.filter(r =>
    String(r[CONFIG.COLS.STATUS]).trim() === 'Pending'
  ).length;

  if (jobs.length === 0) {
    const msg = `No completed jobs found for ${targetStr}. Check the 'Executions' log for details.`;
    Logger.log(msg);
    return msg;
  }

  const grouped = groupBy(jobs, 'yard');
  let html = '<div style="font-family: Arial, sans-serif; color: #000;">';

  // Yard groups
  for (const yard in grouped) {
    html += `<h3 style="margin-bottom: 5px; color: #222; text-transform: uppercase;">${yard}</h3>`;
    grouped[yard].forEach(j => {
      const qtyStr = j.qty > 1 ? ` <strong>X${j.qty}</strong>` : '';
      html += `<div style="margin-bottom: 2px;">${j.company}${qtyStr}</div>`;
    });
    html += '<br>';
  }

  // Footer
  html += `
    <hr style="border: 1px solid #ccc;">
    <div style="font-weight: bold;">CUSTOMERS WAITING FOR AVAILABLE ASSETS: ${pendingCount}</div>
    <div style="font-weight: bold;">AVAILABLE ASSETS: 0</div>
  </div>`;

  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENTS.join(','),
    subject: `✅ Completed Report - ${targetStr}`,
    htmlBody: html
  });

  return `Report sent for ${jobs.length} jobs on ${targetStr}.`;
}

// --- SCHEDULED REPORT (ALL SCHEDULED) ---
function sendNextDaySchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  const data = sh.getDataRange().getValues();
  data.shift(); // headers

  // Filter: Status is 'Scheduled'
  const jobs = data
    .filter(row => String(row[CONFIG.COLS.STATUS]).trim() === 'Scheduled')
    .map(mapRowToObject);

  if (jobs.length === 0) return 'No jobs scheduled.';

  const grouped = groupBy(jobs, 'yard');

  let html = '<div style="font-family: Arial, sans-serif; color: #000;">';

  // 1. Yard Lists
  for (const yard in grouped) {
    html += `<h3 style="margin-bottom: 5px; color: #222; text-transform: uppercase;">${yard}</h3>`;
    grouped[yard].forEach(j => {
      const qtyStr = j.qty > 1 ? ` <strong>X${j.qty}</strong>` : '';
      html += `<div style="margin-bottom: 2px;">${j.company}${qtyStr}</div>`;
    });
    html += '<br>';
  }

  // 2. Picks Summary
  const picks = jobs.filter(j =>
    String(j.action || '').toLowerCase().includes('pick')
  );
  if (picks.length > 0) {
    html += `<h3 style="margin-bottom: 5px; color: #222; text-transform: uppercase;">PICK</h3>`;
    picks.forEach(j => {
      html += `<div style="margin-bottom: 2px;">${j.company} - ${j.city}</div>`;
    });
    html += '<br>';
  }

  // 3. Drops Summary
  const drops = jobs.filter(j =>
    String(j.action || '').toLowerCase().includes('drop')
  );
  if (drops.length > 0) {
    html += `<h3 style="margin-bottom: 5px; color: #222; text-transform: uppercase;">DROP</h3>`;
    drops.forEach(j => {
      html += `<div style="margin-bottom: 2px;">${j.company} - ${j.city}</div>`;
    });
  }

  html += '</div>';

  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENTS.join(','),
    subject: `📅 Scheduled - ${formatDate(getTargetDate('next'))}`,
    htmlBody: html
  });

  return `Schedule sent for ${jobs.length} jobs.`;
}

// ==========================================
// 4. HELPERS
// ==========================================

function getDashboardStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) return { scheduled: 0, complete: 0 };

  const data = sh.getDataRange().getValues();
  data.shift(); // headers

  let scheduled = 0;
  let complete = 0;

  data.forEach(row => {
    const status = String(row[CONFIG.COLS.STATUS]).trim();
    if (status === 'Scheduled') scheduled++;
    if (status === 'Complete') complete++;
  });

  return { scheduled, complete };
}

function getCompanyList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  // Get all companies from Column B
  const data = sh.getRange(2, 2, lastRow - 1, 1).getValues();
  const unique = [...new Set(data.flat().filter(String))];
  return unique.sort();
}

function groupBy(arr, key) {
  return arr.reduce((acc, obj) => {
    const k = (obj[key] || 'Unassigned').toUpperCase();
    (acc[k] = acc[k] || []).push(obj);
    return acc;
  }, {});
}

function getTargetDate(direction) {
  const d = new Date();
  const day = d.getDay(); // 0=Sun ... 6=Sat

  if (direction === 'next') {
    // If Friday (5), next business day is Monday (+3 days), else +1 day
    d.setDate(d.getDate() + (day === 5 ? 3 : 1));
  } else if (direction === 'prev') {
    // If Monday (1), previous business day is Friday (-3 days), else -1 day
    d.setDate(d.getDate() - (day === 1 ? 3 : 1));
  }

  d.setHours(0, 0, 0, 0);
  return d;
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy');
}

function mapRowToObject(row) {
  return {
    company: row[CONFIG.COLS.COMPANY],
    city: row[CONFIG.COLS.CITY],
    action: row[CONFIG.COLS.ACTION],
    yard: row[CONFIG.COLS.YARD],
    box: row[CONFIG.COLS.BOX],
    qty: row[CONFIG.COLS.QTY] || 1,
    notes: row[CONFIG.COLS.NOTES]
  };
}

function fixExistingHelperDates() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dispatch Sheet');
  const lastRow = sh.getLastRow();
  // Get all Completed Dates (Column I)
  const ranges = sh.getRange(2, 9, lastRow - 1, 1).getValues(); 
  const updates = [];
  
  ranges.forEach(r => {
    const val = r[0];
    if (val instanceof Date) {
      // Create the string
      updates.push([Utilities.formatDate(val, Session.getScriptTimeZone(), 'MM/dd/yyyy')]);
    } else {
      updates.push(['']);
    }
  });
  
  // Write to Column K
  sh.getRange(2, 11, updates.length, 1).setValues(updates);
}