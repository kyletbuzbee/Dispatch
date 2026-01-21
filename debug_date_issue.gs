/**
 * Debug script to test date comparison logic
 * This will help identify why sendPreviousDayReport can't find previous days completed
 */

function debugDateComparison() {
  Logger.log('🔍 Starting date comparison debug');

  // Get the target date (previous business day)
  const targetDate = getTargetDate('prev');
  const targetStr = formatDate(targetDate);

  Logger.log('Target date (Date object): ' + targetDate);
  Logger.log('Target date (formatted string): ' + targetStr);

  // Get spreadsheet data
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  const data = sh.getDataRange().getValues();
  data.shift(); // Remove headers

  Logger.log('Total rows in spreadsheet: ' + data.length);

  // Find all completed jobs and log their date completed values
  const completedJobs = [];
  data.forEach((row, index) => {
    const status = String(row[CONFIG.COLS.STATUS]).trim();
    if (status === 'Complete') {
      const rawDate = row[CONFIG.COLS.DATE_COMPLETED];
      const rowNum = index + 2; // +2 because we removed headers and rows start at 2

      Logger.log(`Row ${rowNum}: Company: ${row[CONFIG.COLS.COMPANY]}, Date Completed: ${rawDate} (type: ${typeof rawDate})`);

      // Try to convert the date
      let dateStr = '';
      if (rawDate instanceof Date) {
        dateStr = formatDate(rawDate);
      } else if (rawDate) {
        const parsed = new Date(rawDate);
        if (!isNaN(parsed.getTime())) {
          dateStr = formatDate(parsed);
        } else {
          dateStr = String(rawDate).trim();
        }
      }

      Logger.log(`  -> Converted to: "${dateStr}"`);
      Logger.log(`  -> Matches target "${targetStr}": ${dateStr === targetStr}`);

      completedJobs.push({
        row: rowNum,
        company: row[CONFIG.COLS.COMPANY],
        rawDate: rawDate,
        dateStr: dateStr,
        matches: dateStr === targetStr
      });
    }
  });

  Logger.log(`Found ${completedJobs.length} completed jobs total`);
  const matchingJobs = completedJobs.filter(job => job.matches);
  Logger.log(`Found ${matchingJobs.length} jobs matching target date ${targetStr}`);

  // Log timezone info
  Logger.log('Script timezone: ' + Session.getScriptTimeZone());
  Logger.log('Current date: ' + new Date());

  return {
    targetDate: targetDate,
    targetStr: targetStr,
    completedJobs: completedJobs,
    matchingJobs: matchingJobs
  };
}

/**
 * Test the getTargetDate function specifically
 */
function debugGetTargetDate() {
  Logger.log('📅 Testing getTargetDate function');

  const today = new Date();
  Logger.log('Today: ' + today);
  Logger.log('Today day of week: ' + today.getDay()); // 0=Sun, 1=Mon, ..., 6=Sat

  const prevDate = getTargetDate('prev');
  Logger.log('Previous business day: ' + prevDate);
  Logger.log('Previous business day of week: ' + prevDate.getDay());

  const nextDate = getTargetDate('next');
  Logger.log('Next business day: ' + nextDate);
  Logger.log('Next business day of week: ' + nextDate.getDay());

  // Test formatting
  Logger.log('Formatted previous: ' + formatDate(prevDate));
  Logger.log('Formatted next: ' + formatDate(nextDate));

  return {
    today: today,
    prevDate: prevDate,
    nextDate: nextDate
  };
}

/**
 * Test date parsing from spreadsheet
 */
function debugDateParsing() {
  Logger.log('📊 Testing date parsing from spreadsheet');

  // Test various date formats that might be in the spreadsheet
  const testDates = [
    new Date('2026-01-19'), // Date object
    '01/19/2026',           // String format
    '1/19/2026',            // Short string format
    '01-19-2026',           // Dash format
    'January 19, 2026',     // Long format
    ''                      // Empty
  ];

  const targetStr = '01/19/2026'; // Example target

  testDates.forEach((testDate, index) => {
    Logger.log(`Test ${index + 1}: Input = "${testDate}" (type: ${typeof testDate})`);

    let dateStr = '';
    if (testDate instanceof Date) {
      dateStr = formatDate(testDate);
    } else if (testDate) {
      const parsed = new Date(testDate);
      if (!isNaN(parsed.getTime())) {
        dateStr = formatDate(parsed);
      } else {
        dateStr = String(testDate).trim();
      }
    }

    Logger.log(`  -> Converted to: "${dateStr}"`);
    Logger.log(`  -> Matches target "${targetStr}": ${dateStr === targetStr}`);
  });
}
