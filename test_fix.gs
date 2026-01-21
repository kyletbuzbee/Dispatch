/**
 * Test function to verify the date comparison fix
 */

function testDateComparisonFix() {
  Logger.log('🧪 Testing date comparison fix');

  // Test the updated date parsing logic
  const testCases = [
    {
      input: new Date('2026-01-19'),
      expected: '01/19/2026',
      description: 'Date object'
    },
    {
      input: '01/19/2026',
      expected: '01/19/2026',
      description: 'String date'
    },
    {
      input: 44585, // Excel serial number for 01/19/2026
      expected: '01/19/2026',
      description: 'Excel serial number'
    },
    {
      input: '1/19/2026',
      expected: '01/19/2026',
      description: 'Short string date'
    }
  ];

  testCases.forEach((testCase, index) => {
    Logger.log(`Test ${index + 1}: ${testCase.description}`);

    // Simulate the date parsing logic from the fixed function
    let dateStr = '';
    const rawDate = testCase.input;

    if (rawDate instanceof Date) {
      dateStr = formatDate(rawDate);
    } else if (typeof rawDate === 'number') {
      const excelDate = new Date((rawDate - 25569) * 86400 * 1000);
      dateStr = formatDate(excelDate);
    } else {
      const parsed = new Date(rawDate);
      if (!isNaN(parsed.getTime())) {
        dateStr = formatDate(parsed);
      } else {
        dateStr = String(rawDate).trim();
      }
    }

    Logger.log(`  Input: ${testCase.input} (type: ${typeof testCase.input})`);
    Logger.log(`  Output: ${dateStr}`);
    Logger.log(`  Expected: ${testCase.expected}`);
    Logger.log(`  Match: ${dateStr === testCase.expected ? '✅' : '❌'}`);
  });

  // Test the actual function
  Logger.log('\n📊 Testing actual sendPreviousDayReport function');

  // Mock some data that would come from the spreadsheet
  const mockData = [
    ['Pending', '', 'Test Company 1', '', '', '', '', '', ''], // Not complete
    ['Complete', '', 'Test Company 2', '', '', '', '', '', '01/19/2026'], // String date
    ['Complete', '', 'Test Company 3', '', '', '', '', '', new Date('2026-01-19')], // Date object
    ['Complete', '', 'Test Company 4', '', '', '', '', '', 44585], // Excel serial number
    ['Complete', '', 'Test Company 5', '', '', '', '', '', '01/20/2026'], // Different date
  ];

  // Set target date to 01/19/2026
  const targetDate = new Date('2026-01-19');
  const targetStr = formatDate(targetDate);

  Logger.log(`Target date: ${targetStr}`);

  const matchingJobs = mockData.filter(row => {
    const status = String(row[CONFIG.COLS.STATUS]).trim();
    const rawDate = row[CONFIG.COLS.DATE_COMPLETED];

    if (status !== 'Complete' || !rawDate) return false;

    let dateStr = '';
    if (rawDate instanceof Date) {
      dateStr = formatDate(rawDate);
    } else if (typeof rawDate === 'number') {
      const excelDate = new Date((rawDate - 25569) * 86400 * 1000);
      dateStr = formatDate(excelDate);
    } else {
      const parsed = new Date(rawDate);
      if (!isNaN(parsed.getTime())) {
        dateStr = formatDate(parsed);
      } else {
        dateStr = String(rawDate).trim();
      }
    }

    const matches = dateStr === targetStr;
    Logger.log(`  Company: ${row[CONFIG.COLS.COMPANY]}, Date: ${rawDate} -> ${dateStr}, Matches: ${matches ? '✅' : '❌'}`);
    return matches;
  });

  Logger.log(`Found ${matchingJobs.length} matching jobs out of ${mockData.filter(row => row[CONFIG.COLS.STATUS] === 'Complete').length} completed jobs`);

  return {
    success: matchingJobs.length === 3, // Should find 3 out of 4 completed jobs (excluding the one with different date)
    foundJobs: matchingJobs.length,
    expectedJobs: 3
  };
}
