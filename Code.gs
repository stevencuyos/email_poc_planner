// ============================================================
//  POC Planner — Google Apps Script Backend
//  Schedule Source : 1J6X6Bxqh3vXFHP4iz5eqbrOpvYHHJC2w_4iBJI2haco
//  Schedule Tab    : "NEW Schedule File"
//  Publish Target  : 1DNQXuONhHRZUozekjM6NlbtN0e9nGmtEBlLgGo4D-TA
// ============================================================

const SCHEDULE_SS_ID  = '1J6X6Bxqh3vXFHP4iz5eqbrOpvYHHJC2w_4iBJI2haco';
const SCHEDULE_TAB    = 'NEW Schedule File';

const PUBLISH_SS_ID   = '1SwO6Wet3OWPQDkXC2jyQ3rbPTjCDMZbAc5fLbDK6HHU';
const SHIFT_HOURS     = 9;
const RESTRICT_START  = 2; // no POC in first 2 hrs of shift
const RESTRICT_END    = 0; // no POC in last 0 hr of shift

const LOG_TAB         = 'Change Log';

// ── POC Schedule master grid layout ──────────────────────────
const SCHEDULE_TAB_NAME = 'POC Schedule';
const FIRST_DATA_COL    = 2;   // column B — first week group starts here
const GROUP_HEADER_ROW  = 6;   // merged rows 6:7 — week band (e.g. "WS0712")
const DATE_HEADER_ROW   = 8;   // day+date sub-header, one column per day
const DATA_START_ROW    = 9;   // hour 00:00 row
const HOURS_PER_DAY     = 24;

const GROUP_PALETTE = [
  { band: '#1a1a3e', bandFg: '#ffffff', sub: '#5856D6', subFg: '#ffffff' },
  { band: '#0b3d2e', bandFg: '#ffffff', sub: '#248A3D', subFg: '#ffffff' },
  { band: '#7a3300', bandFg: '#ffffff', sub: '#C93400', subFg: '#ffffff' },
  { band: '#1f3a5f', bandFg: '#ffffff', sub: '#0F6FC5', subFg: '#ffffff' },
  { band: '#4a154b', bandFg: '#ffffff', sub: '#8B2E8F', subFg: '#ffffff' }
];

// ── Entry point ──────────────────────────────────────────────
function doGet() {
  try { appendLogRow('VIEW', '', 'Web app doGet() loaded'); } catch (e) { Logger.log('log failed: ' + e.message); }
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Email POC Planner')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Read schedule from "NEW Schedule File" tab ───────────────
function getScheduleData() {
  try {
    const ss    = SpreadsheetApp.openById(SCHEDULE_SS_ID);
    const sheet = ss.getSheetByName(SCHEDULE_TAB);
    if (!sheet) throw new Error('Tab "' + SCHEDULE_TAB + '" not found.');

    // 1. Get DISPLAY values for rows (keeps shift times as exact text strings)
    const displayData = sheet.getDataRange().getDisplayValues();

    // 2. Get RAW values specifically for Row 2 (keeps true Date objects for day matching)
    const rawDateRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Find header row (contains "Nickname"), then also locate LDAP and Site columns
    let headerRowIdx = -1, nameCol = -1, siteCol = -1, ldapCol = -1;
    for (let r = 0; r < Math.min(displayData.length, 10); r++) {
      const ni = displayData[r].findIndex(c => /nickname/i.test(String(c)));
      if (ni !== -1) {
        headerRowIdx = r;
        nameCol = ni;
        siteCol = displayData[r].findIndex(c => /site/i.test(String(c)));
        ldapCol = displayData[r].findIndex(c => /^ldap$/i.test(String(c).trim()));
        break;
      }
    }
    if (headerRowIdx === -1) throw new Error('Could not find "Nickname" header in first 10 rows.');
    if (siteCol === -1) throw new Error('Could not find "Site" column after Nickname.');
    if (ldapCol === -1) throw new Error('Could not find "LDAP" column in header row.');

    const DAY_NAMES = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    const MON_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

    // Find XG column
    let scheduleStartCol = -1;
    const headerRow = displayData[headerRowIdx];
    for (let c = siteCol + 1; c < headerRow.length; c++) {
      const cellRef = sheet.getRange(headerRowIdx + 1, c + 1).getA1Notation();
      if (cellRef.startsWith('XG')) {
        scheduleStartCol = c;
        break;
      }
    }

    if (scheduleStartCol === -1) {
      Logger.log('getScheduleData: "XG" column marker not found by A1 letter — falling back to column 215.');
      scheduleStartCol = 215; // known-good fallback for this sheet's layout
    }

    // Collect ALL date columns from scheduleStartCol onwards using the true raw Date objects
    const dateCols = [];
    for (let c = scheduleStartCol; c < headerRow.length; c++) {
      const dateRaw = rawDateRow[c];

      if (!dateRaw || !(dateRaw instanceof Date) || isNaN(dateRaw.getTime())) continue;

      // Format label as "May-24" manually from the true date object
      const dateLabel = MON_NAMES[dateRaw.getMonth()] + '-' + String(dateRaw.getDate()).padStart(2,'0');
      const dayLabel = DAY_NAMES[dateRaw.getDay()];

      const label = dayLabel + ' ' + dateLabel;
      dateCols.push({ col: c, label });
    }

    if (dateCols.length === 0) throw new Error('No date columns found starting from XG.');

    // Read support rows using display values for text comparison.
    // Use LDAP (always populated) as both the display name AND the row
    // terminator — Nickname is blank for many agents, which used to cut
    // the list short at the first blank one.
    const supports = [];
    for (let r = headerRowIdx + 1; r < displayData.length; r++) {
      const row  = displayData[r];
      const ldap = String(row[ldapCol] || '').trim();
      if (!ldap || /interval/i.test(ldap)) break;

      const site  = String(row[siteCol] || '').trim().toUpperCase();
      const sched = dateCols.map(dc => {
        const cellText = row[dc.col];
        return parseTimeToHour(cellText);
      });
      supports.push({ name: ldap, site, sched });
    }

    try {
      appendLogRow('VIEW', dateCols[0].label,
        'getScheduleData loaded — ' + supports.length + ' agent(s), ' + dateCols.length + ' day(s)');
    } catch (logErr) { Logger.log('log failed: ' + logErr.message); }

    return {
      ok:        true,
      supports,
      days:      dateCols.map(dc => dc.label),
      weekStart: dateCols[0].label
    };
  } catch (e) {
    try { appendLogRow('ERROR', '', 'getScheduleData failed: ' + e.message); } catch (logErr) { Logger.log('log failed: ' + logErr.message); }
    return { ok: false, error: e.message };
  }
}

// ── Parse time values to hour integer ───────────────────────
function parseTimeToHour(val) {
  if (!val) return 'OFF';
  const str = String(val).trim().toUpperCase();

  // 1. Handle OFF, VL, LOA
  if (/^(OFF|VL|LOA)$/.test(str)) return str;

  // 2. Extract the hour from text (e.g., "4:00 AM" -> 4, "15:00" -> 15)
  const timeMatch = str.match(/^(\d{1,2})(?::\d{2})?\s*(AM|PM)?/);
  if (timeMatch) {
    let h = parseInt(timeMatch[1], 10);
    const ampm = timeMatch[2];

    if (ampm === 'PM' && h < 12) h += 12;
    if (ampm === 'AM' && h === 12) h = 0;
    return h;
  }

  return 'OFF';
}

// ── Publish POC assignments → POC Schedule master grid ───────
// payload.dayIndices (optional): array of 0-based day-indices (within the
// week's `days` array) to publish. Omitted/empty = publish the whole week
// (unchanged behavior). Any day NOT in dayIndices is left exactly as it
// currently is on the sheet — this is what powers "publish selected day(s)".
function publishPOC(payloadJson) {
  let payload;
  try {
    payload = JSON.parse(payloadJson);
    const result = writeToPOCSchedule(payload.weekLabel, payload.days, payload.assignments, payload.dayIndices);

    const scopeNote = (payload.dayIndices && payload.dayIndices.length)
      ? 'partial publish — ' + payload.dayIndices.length + ' of ' + payload.days.length + ' day(s)'
      : 'full week';

    try {
      appendLogRow('PUBLISH', payload.weekLabel,
        'Assignments: ' + Object.keys(payload.assignments).length + ' — ' + scopeNote + ' — ' + result.note);
    } catch (logErr) { Logger.log('log failed: ' + logErr.message); }

    return { ok: true, weekLabel: payload.weekLabel, note: result.note, scope: scopeNote };
  } catch (e) {
    try {
      appendLogRow('ERROR', (payload && payload.weekLabel) || '', 'publishPOC failed: ' + e.message);
    } catch (logErr) { Logger.log('log failed: ' + logErr.message); }
    return { ok: false, error: e.message };
  }
}

// ── Locate existing week-group bands in row GROUP_HEADER_ROW ─
// Returns [{ label, startCol, endCol, width }, ...] in left-to-right order.
// Width is inferred from the gap to the next group's start column, so this
// works even though not every group is the same width (partial weeks).
function scanScheduleGroups(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < FIRST_DATA_COL) return [];

  const headerVals = sheet
    .getRange(GROUP_HEADER_ROW, FIRST_DATA_COL, 1, lastCol - FIRST_DATA_COL + 1)
    .getDisplayValues()[0];

  const groups = [];
  headerVals.forEach((val, i) => {
    const trimmed = String(val || '').trim();
    if (trimmed) groups.push({ label: trimmed, startCol: FIRST_DATA_COL + i });
  });
  groups.forEach((g, i) => {
    g.endCol = (i + 1 < groups.length) ? groups[i + 1].startCol - 1 : lastCol;
    g.width  = g.endCol - g.startCol + 1;
  });
  return groups;
}

// ── Write one week's assignments into the POC Schedule grid ──
// Creates a new week band + date header if this week isn't there yet;
// otherwise overwrites just the hourly data for the existing band.
// dayIndices (optional): 0-based day-indices within `days` to actually
// overwrite. Columns outside this set keep their current sheet values.
function writeToPOCSchedule(weekLabel, days, weekAssignments, dayIndices) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return writeToPOCScheduleLocked(weekLabel, days, weekAssignments, dayIndices);
  } finally {
    lock.releaseLock();
  }
}

function writeToPOCScheduleLocked(weekLabel, days, weekAssignments, dayIndices) {
  const ss    = SpreadsheetApp.openById(PUBLISH_SS_ID);
  const sheet = ss.getSheetByName(SCHEDULE_TAB_NAME);
  if (!sheet) throw new Error('Tab "' + SCHEDULE_TAB_NAME + '" not found.');

  const groups = scanScheduleGroups(sheet);
  const match  = groups.find(g => g.label === weekLabel);

  let startCol, writeWidth, note;
  let isNewGroup = false;

  if (match) {
    startCol   = match.startCol;
    writeWidth = Math.min(match.width, days.length);
    note = (match.width === days.length)
      ? 'existing group updated'
      : 'existing group is ' + match.width + ' col(s) wide, publishing ' + days.length +
        ' day(s) — wrote ' + writeWidth + ' col(s), rest left untouched (manual column insert needed to grow it)';
  } else {
    isNewGroup = true;
    startCol   = groups.length ? groups[groups.length - 1].endCol + 1 : FIRST_DATA_COL;
    writeWidth = days.length;
    const palette = GROUP_PALETTE[groups.length % GROUP_PALETTE.length];

    const bandRange = sheet.getRange(GROUP_HEADER_ROW, startCol, 2, writeWidth);
    bandRange.merge();
    bandRange.setValue(weekLabel);
    bandRange.setBackground(palette.band);
    bandRange.setFontColor(palette.bandFg);
    bandRange.setFontWeight('bold');
    bandRange.setHorizontalAlignment('center');
    bandRange.setVerticalAlignment('middle');

    for (let i = 0; i < writeWidth; i++) {
      const cell  = sheet.getRange(DATE_HEADER_ROW, startCol + i);
      const parts = String(days[i] || '').match(/^([A-Za-z]+)\s+(.+)$/);
      cell.setValue(parts ? parts[1] + '\n' + parts[2] : (days[i] || ''));
      cell.setBackground(palette.sub);
      cell.setFontColor(palette.subFg);
      cell.setFontWeight('bold');
      cell.setHorizontalAlignment('center');
      cell.setWrap(true);
    }
    sheet.setColumnWidths(startCol, writeWidth, 90);
    note = 'new group created';
  }

  // Which day-columns (0-based, relative to this group) are we actually
  // allowed to overwrite this call? Empty/omitted dayIndices = all of them.
  const publishAll = !dayIndices || dayIndices.length === 0;
  const publishSet = publishAll ? null : new Set(dayIndices);

  // Snapshot the previous values so the log records what changed, not just that it changed
  // — and so untouched columns can be preserved verbatim on a partial publish.
  const beforeRange  = sheet.getRange(DATA_START_ROW, startCol, HOURS_PER_DAY, writeWidth);
  const beforeValues = isNewGroup ? null : beforeRange.getValues();

  const rows = [];
  for (let h = 0; h < HOURS_PER_DAY; h++) {
    const rowVals = [];
    for (let di = 0; di < writeWidth; di++) {
      const shouldWrite = publishAll || publishSet.has(di);
      if (shouldWrite) {
        rowVals.push(weekAssignments[di + ':' + h] || '');
      } else {
        // Keep whatever is already on the sheet for a column we're not publishing.
        const existing = beforeValues && beforeValues[h] ? beforeValues[h][di] : '';
        rowVals.push(existing || '');
      }
    }
    rows.push(rowVals);
  }
  const dataRange = sheet.getRange(DATA_START_ROW, startCol, HOURS_PER_DAY, writeWidth);
  dataRange.setValues(rows);
  dataRange.setHorizontalAlignment('center');

  if (!publishAll) {
    note += ' (partial: col-index ' + Array.from(publishSet).sort((a,b)=>a-b).join(',') + ' written, others preserved)';
  }

  try {
    const changedCells = countChangedCells(beforeValues, rows);
    appendLogRow('WRITE', weekLabel,
      (isNewGroup ? 'New group written' : changedCells + ' cell(s) changed') +
      ' at col ' + startCol + ', width ' + writeWidth);
  } catch (logErr) { Logger.log('log failed: ' + logErr.message); }

  return { startCol, width: writeWidth, note };
}

// ── Count how many cells actually changed between two grids ──
function countChangedCells(beforeValues, afterValues) {
  if (!beforeValues) return afterValues.length * (afterValues[0] ? afterValues[0].length : 0);
  let count = 0;
  for (let r = 0; r < afterValues.length; r++) {
    for (let c = 0; c < afterValues[r].length; c++) {
      const b = beforeValues[r] ? beforeValues[r][c] : '';
      if (String(b) !== String(afterValues[r][c])) count++;
    }
  }
  return count;
}

// ── Read back published assignments for ONE week (hydrate on load) ───
// Mirrors the "di:h" -> name shape that publishPOC() writes, so the
// frontend can drop the result straight into its assignments map.
function getPublishedAssignments(weekLabel, days) {
  try {
    const ss    = SpreadsheetApp.openById(PUBLISH_SS_ID);
    const sheet = ss.getSheetByName(SCHEDULE_TAB_NAME);
    if (!sheet) throw new Error('Tab "' + SCHEDULE_TAB_NAME + '" not found.');

    const groups = scanScheduleGroups(sheet);
    const match  = groups.find(g => g.label === weekLabel);
    if (!match) return { ok: true, found: false, assignments: {} };

    const width = days && days.length ? Math.min(match.width, days.length) : match.width;
    const range = sheet.getRange(DATA_START_ROW, match.startCol, HOURS_PER_DAY, width).getDisplayValues();

    const assignments = {};
    for (let h = 0; h < HOURS_PER_DAY; h++) {
      for (let di = 0; di < width; di++) {
        const val = String(range[h][di] || '').trim();
        if (val) assignments[di + ':' + h] = val;
      }
    }
    return { ok: true, found: true, assignments, width };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ── Read back published assignments for EVERY week band currently on
//    the POC Schedule grid. Used to (a) hydrate the planner across all
//    weeks on open, and (b) power the "Published Weeks" overview panel. ──
function getAllPublishedWeeks() {
  try {
    const t0    = new Date().getTime();
    const ss    = SpreadsheetApp.openById(PUBLISH_SS_ID);
    const sheet = ss.getSheetByName(SCHEDULE_TAB_NAME);
    if (!sheet) throw new Error('Tab "' + SCHEDULE_TAB_NAME + '" not found.');

    const groups = scanScheduleGroups(sheet);
    if (!groups.length) return { ok: true, weeks: [] };

    // Bound the batched read to the actual last published week band, NOT
    // sheet.getLastColumn(). getLastColumn() reports the sheet's last
    // formatted/touched column, which can drift far past real data (a
    // stray border, a cell someone clicked once, leftover formatting from
    // deleting a week) — and grows every time the tab is edited. Reading
    // out to that phantom edge instead of the real last week's end column
    // can turn one fast batched read into one that scans thousands of
    // empty columns and hangs with no error ever reaching the browser.
    const lastCol   = groups[groups.length - 1].endCol;
    const totalCols = lastCol - FIRST_DATA_COL + 1;

    // Single batched read covering every group's header + data range, instead
    // of one getRange()/getDisplayValues() round-trip PER published week.
    // Each extra week used to add another full network call here — with
    // several weeks now published, that was adding up to real, growing delay.
    const dateHeaders = sheet
      .getRange(DATE_HEADER_ROW, FIRST_DATA_COL, 1, totalCols)
      .getDisplayValues()[0];
    const allData = sheet
      .getRange(DATA_START_ROW, FIRST_DATA_COL, HOURS_PER_DAY, totalCols)
      .getDisplayValues();

    const weeksOut = groups.map(g => {
      const colOffset = g.startCol - FIRST_DATA_COL;

      const days = [];
      for (let i = 0; i < g.width; i++) {
        const raw = dateHeaders[colOffset + i] || '';
        days.push(String(raw).replace(/\n/g, ' ').trim());
      }

      const assignments = {};
      for (let h = 0; h < HOURS_PER_DAY; h++) {
        for (let i = 0; i < g.width; i++) {
          const val = String(allData[h][colOffset + i] || '').trim();
          if (val) assignments[i + ':' + h] = val;
        }
      }

      return { label: g.label, days, assignments };
    });

    Logger.log('getAllPublishedWeeks: ' + groups.length + ' week(s), ' + (new Date().getTime() - t0) + 'ms');
    return { ok: true, weeks: weeksOut };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ── Delete an entire week group from the POC Schedule grid ───
// Clears the band header, date header, and all hourly data for that week.
function deleteWeekGroup(weekLabel) {
  try {
    const ss    = SpreadsheetApp.openById(PUBLISH_SS_ID);
    const sheet = ss.getSheetByName(SCHEDULE_TAB_NAME);
    if (!sheet) throw new Error('Tab "' + SCHEDULE_TAB_NAME + '" not found.');

    const groups = scanScheduleGroups(sheet);
    const match  = groups.find(g => g.label === weekLabel);
    if (!match) throw new Error('Week "' + weekLabel + '" not found.');

    sheet.getRange(GROUP_HEADER_ROW, match.startCol, 2, match.width).breakApart();
    sheet.getRange(GROUP_HEADER_ROW, match.startCol, (DATA_START_ROW - GROUP_HEADER_ROW) + HOURS_PER_DAY, match.width).clearContent();
    sheet.getRange(GROUP_HEADER_ROW, match.startCol, 2, match.width).setBackground(null).setFontColor(null);
    sheet.getRange(DATE_HEADER_ROW, match.startCol, 1, match.width).setBackground(null).setFontColor(null);

    appendLogRow('DELETE', weekLabel, 'Week group cleared — ' + match.width + ' col(s) at col ' + match.startCol);
    return { ok: true, weekLabel };
  } catch (e) {
    try { appendLogRow('ERROR', weekLabel || '', 'deleteWeekGroup failed: ' + e.message); } catch (logErr) { Logger.log('log failed: ' + logErr.message); }
    return { ok: false, error: e.message };
  }
}

// ── Change Log ────────────────────────────────────────────────
function getLogSheet(ss) {
  ss = ss || SpreadsheetApp.openById(PUBLISH_SS_ID);
  let sheet = ss.getSheetByName(LOG_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_TAB, 0);
    const header = ['Timestamp', 'User', 'Action', 'Week', 'Details'];
    const hRange = sheet.getRange(1, 1, 1, header.length);
    hRange.setValues([header]);
    hRange.setBackground('#1a1a3e');
    hRange.setFontColor('#ffffff');
    hRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 110);
    sheet.setColumnWidth(4, 90);
    sheet.setColumnWidth(5, 400);
  }
  return sheet;
}

const ACTION_COLORS = {
  OPEN:        { bg: '#E3F2FD', fg: '#0D47A1' },
  VIEW:        { bg: '#E3F2FD', fg: '#0D47A1' },
  EDIT_BATCH:  { bg: '#FFF8E1', fg: '#8D6E00' },
  EDIT_MANUAL: { bg: '#FFF3E0', fg: '#B15400' },
  WRITE:       { bg: '#E8F5E9', fg: '#1B5E20' },
  PUBLISH:     { bg: '#E0F7FA', fg: '#00695C' },
  DELETE:      { bg: '#FCE4EC', fg: '#B71C1C' },
  ERROR:       { bg: '#FFEBEE', fg: '#C62828' },
  CLIENT_ERROR:{ bg: '#FFEBEE', fg: '#C62828' }
};

function appendLogRow(action, week, details, ss) {
  const sheet = getLogSheet(ss);
  const user = Session.getActiveUser().getEmail() || 'unknown';
  sheet.appendRow([new Date(), user, action, week || '', details || '']);

  const lastRow = sheet.getLastRow();
  const colors  = ACTION_COLORS[action];
  if (colors) {
    const actionCell = sheet.getRange(lastRow, 3);
    actionCell.setBackground(colors.bg);
    actionCell.setFontColor(colors.fg);
    actionCell.setFontWeight('bold');
  }
}

function logOpen() {
  try { appendLogRow('OPEN', '', ''); } catch (e) { Logger.log('logOpen failed: ' + e.message); }
  return { ok: true };
}

function logEditBatch(payloadJson) {
  try {
    const payload = JSON.parse(payloadJson);
    appendLogRow('EDIT_BATCH', payload.week || '', payload.summary || '');
  } catch (e) { Logger.log('logEditBatch failed: ' + e.message); }
  return { ok: true };
}

// ── Generic client-callable logger for any other action ──────
// Call from the frontend as: google.script.run.logAction('CLOSE', weekLabel, 'details text')
function logAction(action, week, details) {
  try { appendLogRow(action, week || '', details || ''); } catch (e) { Logger.log('logAction failed: ' + e.message); }
  return { ok: true };
}

// ── Client-side failure logger ────────────────────────────────
// Every withFailureHandler in the frontend calls this so that errors
// happening on the browser side (network hiccups, a rejected promise,
// a thrown exception in a success handler, etc.) land in the same
// Change Log as server-side ERROR rows — one unified error trail.
// context: short string identifying which action failed (e.g. "publish",
// "refresh", "deleteWeek"). message: the error text shown to the user.
function logClientError(context, message, week) {
  try {
    appendLogRow('CLIENT_ERROR', week || '', '[' + (context || 'unknown') + '] ' + (message || 'no message'));
  } catch (e) {
    Logger.log('logClientError failed: ' + e.message);
  }
  return { ok: true };
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 POC Planner')
    .addItem('Open Planner (in Sheets)', 'openPlanner')
    .addItem('Open Full Web App (New Tab)', 'openWebAppLink')
    .addToUi();
  try { appendLogRow('VIEW', '', 'Spreadsheet opened — POC Planner menu added'); } catch (e) { Logger.log('log failed: ' + e.message); }
}

// ── Open the standalone deployed web app in a new browser tab ─
// Server-side code can't force-open a client tab directly, so this shows
// a tiny dialog with a real <a target="_blank"> link — clicking it is a
// genuine user gesture, so it opens cleanly without being popup-blocked.
function openWebAppLink() {
  let url = null;
  try { url = ScriptApp.getService().getUrl(); } catch (e) { url = null; }

  const body = url
    ? '<a href="' + url + '" target="_blank" ' +
      'style="display:inline-block;padding:10px 22px;background:#1a73e8;color:#fff;' +
      'border-radius:20px;text-decoration:none;font-weight:500;font-size:13px" ' +
      'onclick="setTimeout(function(){google.script.host.close();},150)">Open Web App ↗</a>'
    : '<p style="font-size:13px;color:#c5221f;line-height:1.5">' +
      'No live deployment found. In the Apps Script editor, use ' +
      '<b>Deploy → New deployment → Web app</b> to publish one, then try again.</p>';

  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Roboto,Arial,sans-serif;padding:24px;text-align:center">' +
    '<p style="margin-bottom:16px;font-size:14px;color:#3c4043">Opens the full planner in its own browser tab.</p>' +
    body +
    '</div>'
  ).setWidth(360).setHeight(160);

  SpreadsheetApp.getUi().showModalDialog(html, 'Open Full Web App');
  try { appendLogRow('VIEW', '', 'Web app new-tab link dialog opened'); } catch (e) { Logger.log('log failed: ' + e.message); }
}

function handleEditTrigger(e) {
  try {
    const sheet     = e.range.getSheet();
    const sheetName = sheet.getName();
    if (sheetName === LOG_TAB) return; // don't log edits to the log itself

    const cellRef   = e.range.getA1Notation();
    const oldValue  = (e.oldValue !== undefined) ? e.oldValue : '(multi-cell/paste)';
    const newValue  = e.range.getDisplayValue();

    appendLogRow('EDIT_MANUAL', '', 'Tab: "' + sheetName + '" | ' + cellRef + ': "' + oldValue + '" → "' + newValue + '"');
  } catch (err) {
    Logger.log('handleEditTrigger log failed: ' + err.message);
  }
}

// ── Run this ONCE manually from the Apps Script editor to install the trigger ──
function createEditTrigger() {
  const ss = SpreadsheetApp.openById(PUBLISH_SS_ID);
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'handleEditTrigger') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('handleEditTrigger')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}

function openPlanner() {
  const html = HtmlService
    .createHtmlOutputFromFile('index')
    .setWidth(1300)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Email POC Planner');
  try { appendLogRow('VIEW', '', 'Planner opened from spreadsheet menu (dialog)'); } catch (e) { Logger.log('log failed: ' + e.message); }
}

function diagnoseSheetBounds() {
  const ss    = SpreadsheetApp.openById(PUBLISH_SS_ID);
  const sheet = ss.getSheetByName(SCHEDULE_TAB_NAME);
  const groups = scanScheduleGroups(sheet);
  const realLastCol = groups.length ? groups[groups.length - 1].endCol : FIRST_DATA_COL;
  Logger.log('sheet.getLastColumn(): ' + sheet.getLastColumn());
  Logger.log('Actual last published week ends at column: ' + realLastCol);
  Logger.log('Difference (phantom columns): ' + (sheet.getLastColumn() - realLastCol));
}

// ============================================================
//  POC 30-Minute Reminder Emails
// ============================================================
const NOTIF_LOG_TAB = 'Notification Log';
const REMINDER_LEAD_MIN = 30;
const TZ = 'Asia/Manila';

function getNotifLogSheet() {
  const ss = SpreadsheetApp.openById(PUBLISH_SS_ID);
  let sheet = ss.getSheetByName(NOTIF_LOG_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(NOTIF_LOG_TAB);
    const header = ['Timestamp', 'LDAP', 'Date Label', 'Hour', 'Week', 'Status'];
    const hRange = sheet.getRange(1, 1, 1, header.length);
    hRange.setValues([header]).setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, header.length, 140);
  }
  return sheet;
}

function alreadyNotified_(sheet, ldap, dateLabel, hour) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const data = sheet.getRange(2, 2, lastRow - 1, 3).getValues();
  return data.some(r => String(r[0]) === ldap && String(r[1]) === dateLabel && Number(r[2]) === hour);
}

function markNotified_(sheet, ldap, dateLabel, hour, week) {
  sheet.appendRow([new Date(), ldap, dateLabel, hour, week, 'SENT']);
}

function checkPOCReminders() {
  try {
    const ss = SpreadsheetApp.openById(PUBLISH_SS_ID);
    const sheet = ss.getSheetByName(SCHEDULE_TAB_NAME);
    if (!sheet) return;

    const target = new Date(Date.now() + REMINDER_LEAD_MIN * 60000);
    const targetParts = Utilities.formatDate(target, TZ, 'EEE|MMM-dd|H').split('|');
    const targetDayName = targetParts[0];
    const targetDateLabel = targetParts[1];
    const targetHour = parseInt(targetParts[2], 10);
    const headerLabel = targetDayName + '\n' + targetDateLabel;

    const groups = scanScheduleGroups(sheet);
    if (!groups.length) return;

    const logSheet = getNotifLogSheet();

    for (const g of groups) {
      const dateHeaders = sheet.getRange(DATE_HEADER_ROW, g.startCol, 1, g.width).getDisplayValues()[0];
      const colOffset = dateHeaders.findIndex(h => String(h).trim() === headerLabel);
      if (colOffset === -1) continue;

      const col = g.startCol + colOffset;
      const ldap = String(sheet.getRange(DATA_START_ROW + targetHour, col).getDisplayValue() || '').trim();
      if (!ldap) continue;

      if (alreadyNotified_(logSheet, ldap, targetDateLabel, targetHour)) continue;

      sendPOCReminderEmail_(ldap, targetDateLabel, targetHour, g.label);
      markNotified_(logSheet, ldap, targetDateLabel, targetHour, g.label);
    }
  } catch (e) {
    try { appendLogRow('ERROR', '', 'checkPOCReminders failed: ' + e.message); } catch (logErr) {}
  }
}

function sendPOCReminderEmail_(ldap, dateLabel, hour, weekLabel) {
  const email = ldap + '@google.com';
  const shiftStart = formatHour_(hour);
  const shiftEnd = formatHour_((hour + 1) % 24);
  const initials = ldap.slice(0, 2).toUpperCase();
  const avatarUrl = 'https://moma-teams-photos.corp.google.com/photos/' + encodeURIComponent(ldap) + '?sz=72';

  const subject = 'You\u2019re POC in 30 minutes \u2014 ' + dateLabel;

  const body = `
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"></head>
<body style="margin:0;padding:0;background:#f4f6fb;font-family:'Google Sans',Roboto,Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6fb;padding:32px 0;">
<tr><td align="center">
<table width="520" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:8px;border:1px solid #e0e0e0;max-width:520px;">

<tr><td style="padding:24px 32px 0">
<table cellpadding="0" cellspacing="0"><tr>
<td style="width:24px;height:24px;background:#1a73e8;border-radius:5px;text-align:center;vertical-align:middle;color:#fff;font-size:12px;font-weight:700">P</td>
<td style="padding-left:10px;font-size:13px;color:#5f6368;font-weight:500">Email POC Planner</td>
</tr></table>
</td></tr>

<tr><td style="padding:20px 32px 4px">
<table cellpadding="0" cellspacing="0"><tr>
<td style="width:36px">
<img src="${avatarUrl}" width="36" height="36" style="border-radius:50%;display:block;background:#e8eaed" alt="${initials}">
</td>
<td style="padding-left:12px;font-size:20px;color:#202124;font-weight:400">You're POC in <span style="color:#1a73e8;font-weight:500">30 minutes</span></td>
</tr></table>
</td></tr>

<tr><td style="padding:4px 32px 20px">
<div style="font-size:13px;color:#5f6368;line-height:1.5">Your Email point-of-contact schedule starts soon. Here are the details.</div>
</td></tr>

<tr><td style="padding:0 32px">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f8f9fa;border-radius:8px;">
<tr><td style="padding:16px 20px">
${reminderRow_('Shift', shiftStart + ' \u2013 ' + shiftEnd)}
${reminderRow_('Date', dateLabel)}
${reminderRow_('Week', weekLabel)}
${reminderRow_('LDAP', ldap)}
</td></tr>
</table>
</td></tr>

<tr><td style="padding:24px 32px 8px">
<a href="${ScriptApp.getService().getUrl()}" style="display:inline-block;background:#1a73e8;color:#fff;font-size:14px;font-weight:500;padding:10px 24px;border-radius:20px;text-decoration:none">Open POC Schedule</a>
</td></tr>

<tr><td style="padding:16px 32px 24px;border-top:1px solid #e8eaed;margin-top:12px">
<div style="font-size:11px;color:#80868b;line-height:1.6">This is an automated reminder from the Email POC Planner. If you're not available for this slot, swap it in the planner before your shift starts.</div>
</td></tr>

</table>
</td></tr>
</table>
</body>
</html>`;

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: body,
    name: 'Email POC Planner',
    noReply: true
  });
}

function reminderRow_(label, value) {
  return `<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:6px"><tr>
<td style="font-size:13px;color:#5f6368;width:80px">${label}</td>
<td style="font-size:13px;color:#202124;font-weight:500">${value}</td>
</tr></table>`;
}

function formatHour_(h) {
  const ampm = h >= 12 ? 'PM' : 'AM';
  let hr = h % 12;
  if (hr === 0) hr = 12;
  return hr + ':00 ' + ampm;
}

function setupPOCReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkPOCReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkPOCReminders')
    .timeBased()
    .everyMinutes(10)
    .create();
}

function testPOCReminderEmail() {
  sendPOCReminderEmail_('groyonjr', 'Aug-07', 16, 'WS0802');
}
