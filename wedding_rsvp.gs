// ═══════════════════════════════════════════════════════════════
//  Andrew & Cindy Wedding RSVP — Google Apps Script
//  Paste this entire file into your Google Apps Script editor
// ═══════════════════════════════════════════════════════════════

const SHEET_NAME = 'RSVPs';

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // Create sheet + headers if it doesn't exist yet
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      setupRSVPSheet(sheet);
    }

    // Parse incoming data
    const data = JSON.parse(e.postData.contents);

    // Figure out which events were selected
    const eventsRaw  = (data.events || '').toLowerCase();
    const ceremony   = eventsRaw.includes('ceremony')    ? '✓' : '✗';
    const lunch      = eventsRaw.includes('lunch')        ? '✓' : '✗';
    const afterparty = eventsRaw.includes('afterparty')   ? '✓' : '✗';

    // Shuttle
    const shuttle = data.shuttle || '不需要';

    // Append row — new column order: Name, Guests, Ceremony, Lunch, Afterparty, Shuttle, Notes
    sheet.appendRow([
      data.name      || '',
      Number(data.guests) || 1,
      ceremony,
      lunch,
      afterparty,
      shuttle,
      data.notes     || ''
    ]);

    // Rewrite summary every time
    updateSummary(ss, sheet);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Set up the RSVPs sheet with headers and styling ──
function setupRSVPSheet(sheet) {
  const headers = [
    '姓名 Name',
    '人數 Guests',
    '證婚 Ceremony',
    '午宴 Luncheon',
    'After Party',
    '接駁車 Shuttle',
    '備註 Notes'
  ];
  sheet.appendRow(headers);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1e3248');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // Column widths
  sheet.setColumnWidth(1, 140); // Name
  sheet.setColumnWidth(2, 80);  // Guests
  sheet.setColumnWidth(3, 100); // Ceremony
  sheet.setColumnWidth(4, 100); // Luncheon
  sheet.setColumnWidth(5, 100); // After Party
  sheet.setColumnWidth(6, 160); // Shuttle
  sheet.setColumnWidth(7, 220); // Notes
}

// ── Rebuild the Summary sheet from scratch each time ──
function updateSummary(ss, rsvpSheet) {
  let summary = ss.getSheetByName('Summary');
  if (!summary) {
    summary = ss.insertSheet('Summary');
  }
  summary.clearContents();
  summary.clearFormats();

  const lastRow = rsvpSheet.getLastRow();
  const totalGuests = lastRow > 1
    ? rsvpSheet.getRange(2, 2, lastRow - 1, 1).getValues()
        .reduce((s, r) => s + (Number(r[0]) || 0), 0)
    : 0;
  const totalRSVPs = lastRow > 1 ? lastRow - 1 : 0;

  // Count events — now cols 3,4,5 = ceremony, lunch, afterparty; col 6 = shuttle
  let cerCount = 0, lunCount = 0, aftCount = 0;
  const shuttleCounts = { '去程 高鐵→涵碧': 0, '回程 涵碧→高鐵': 0, '來回 Both': 0, '不需要': 0 };

  if (lastRow > 1) {
    const rows = rsvpSheet.getRange(2, 3, lastRow - 1, 4).getValues();
    rows.forEach(r => {
      if (r[0] === '✓') cerCount++;
      if (r[1] === '✓') lunCount++;
      if (r[2] === '✓') aftCount++;
      const s = r[3];
      if (shuttleCounts.hasOwnProperty(s)) shuttleCounts[s]++;
    });
  }

  // ── Write Summary ──
  const data = [
    ['Andrew & Cindy · RSVP 統計', ''],
    ['', ''],
    ['📋 總覽 Overview', ''],
    ['總出席人數 Total Guests', totalGuests],
    ['回覆人數 No. of RSVPs', totalRSVPs],
    ['', ''],
    ['🎊 活動出席 Event Attendance', ''],
    ['證婚 Ceremony', cerCount],
    ['午宴 Luncheon', lunCount],
    ['After Party', aftCount],
    ['', ''],
    ['🚌 接駁車 Shuttle', ''],
    ['去程 高鐵→涵碧', shuttleCounts['去程 高鐵→涵碧']],
    ['回程 涵碧→高鐵', shuttleCounts['回程 涵碧→高鐵']],
    ['來回 Both', shuttleCounts['來回 Both']],
    ['不需要 None', shuttleCounts['不需要']],
  ];

  summary.getRange(1, 1, data.length, 2).setValues(data);

  // Styling
  summary.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#1e3248');
  summary.getRange('A3').setFontWeight('bold').setFontColor('#1e3248');
  summary.getRange('A7').setFontWeight('bold').setFontColor('#1e3248');
  summary.getRange('A12').setFontWeight('bold').setFontColor('#1e3248');
  summary.getRange('B4').setFontSize(20).setFontWeight('bold').setFontColor('#7ab8d8');
  summary.getRange('B5').setFontSize(14);
  summary.getRange('A3:B3').setBackground('#e8eff7');
  summary.getRange('A7:B7').setBackground('#e8eff7');
  summary.getRange('A12:B12').setBackground('#e8eff7');
  summary.setColumnWidth(1, 240);
  summary.setColumnWidth(2, 100);
}

// Test the script is alive
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'alive', message: 'RSVP script is running!' }))
    .setMimeType(ContentService.MimeType.JSON);
}
