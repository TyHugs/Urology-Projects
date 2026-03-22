// ============================================================
// UROLOGY KEY PROJECTS TRACKER — Google Apps Script Backend
// Deploy this as a Web App to connect your GitHub Pages site
// to your Google Sheet.
// ============================================================

const SHEET_NAME = 'Projects';
const RECURRING_SHEET = 'Recurring';
const HISTORY_SHEET = 'Recurring History';
const CALENDAR_SHEET = 'Calendar Events';
const SCHEDULE_SHEET = 'Provider Schedules';
const GOALS_SHEET = 'Goals';

/**
 * GET handler — returns all project data as JSON
 */
function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return jsonResponse({ error: 'Sheet not found. Create a sheet named "Projects".' });
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return jsonResponse({ projects: [], headers: data[0] || [] });
    }

    const headers = data[0];
    const projects = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0] && !row[1]) continue; // skip empty rows
      const project = {};
      headers.forEach((header, idx) => {
        project[header] = row[idx] !== undefined ? row[idx] : '';
      });
      project._row = i + 1; // 1-indexed row number for updates
      projects.push(project);
    }

    // --- Also fetch Recurring obligations ---
    let recurring = [];
    let recurringHeaders = [];
    const rSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECURRING_SHEET);
    if (rSheet && rSheet.getLastRow() >= 2) {
      const rData = rSheet.getDataRange().getValues();
      recurringHeaders = rData[0];
      for (let i = 1; i < rData.length; i++) {
        const row = rData[i];
        if (!row[0] && !row[1]) continue;
        const item = {};
        recurringHeaders.forEach((h, idx) => { item[h] = row[idx] !== undefined ? row[idx] : ''; });
        item._row = i + 1;
        recurring.push(item);
      }
    }

    // --- Also fetch History ---
    let history = [];
    const hSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HISTORY_SHEET);
    if (hSheet && hSheet.getLastRow() >= 2) {
      const hData = hSheet.getDataRange().getValues();
      const hHeaders = hData[0];
      for (let i = 1; i < hData.length; i++) {
        const row = hData[i];
        if (!row[0]) continue;
        const item = {};
        hHeaders.forEach((h, idx) => { item[h] = row[idx] !== undefined ? row[idx] : ''; });
        history.push(item);
      }
    }

    // --- Also fetch Calendar Events ---
    let calendarEvents = [];
    const cSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CALENDAR_SHEET);
    if (cSheet && cSheet.getLastRow() >= 2) {
      const cData = cSheet.getDataRange().getValues();
      const cHeaders = cData[0];
      for (let i = 1; i < cData.length; i++) {
        const row = cData[i];
        if (!row[0] && !row[1]) continue;
        const item = {};
        cHeaders.forEach((h, idx) => { item[h] = row[idx] !== undefined ? row[idx] : ''; });
        item._row = i + 1;
        calendarEvents.push(item);
      }
    }

    // --- Also fetch Provider Schedules ---
    let schedules = [];
    const sSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCHEDULE_SHEET);
    if (sSheet && sSheet.getLastRow() >= 2) {
      const sData = sSheet.getDataRange().getValues();
      const sHeaders = sData[0];
      for (let i = 1; i < sData.length; i++) {
        const row = sData[i];
        if (!row[0] && !row[1]) continue;
        const item = {};
        sHeaders.forEach((h, idx) => { item[h] = row[idx] !== undefined ? row[idx] : ''; });
        item._row = i + 1;
        schedules.push(item);
      }
    }

    // --- Also fetch Goals ---
    let goals = [];
    const gSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GOALS_SHEET);
    if (gSheet && gSheet.getLastRow() >= 2) {
      const gData = gSheet.getDataRange().getValues();
      const gHeaders = gData[0];
      for (let i = 1; i < gData.length; i++) {
        const row = gData[i];
        if (!row[0] && !row[1]) continue;
        const item = {};
        gHeaders.forEach((h, idx) => { item[h] = row[idx] !== undefined ? row[idx] : ''; });
        item._row = i + 1;
        goals.push(item);
      }
    }

    return jsonResponse({ projects, headers, recurring, recurringHeaders, history, calendarEvents, schedules, goals });

  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
}

/**
 * POST handler — updates a specific cell or adds/deletes a row
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    if (!sheet) {
      return jsonResponse({ error: 'Sheet not found.' });
    }

    // --- UPDATE a cell ---
    if (payload.action === 'update') {
      const { row, column, value } = payload;
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const colIndex = headers.indexOf(column);

      if (colIndex === -1) {
        return jsonResponse({ error: `Column "${column}" not found.` });
      }

      sheet.getRange(row, colIndex + 1).setValue(value);

      // Update "Last Updated" timestamp
      const tsIndex = headers.indexOf('Last Updated');
      if (tsIndex !== -1) {
        sheet.getRange(row, tsIndex + 1).setValue(new Date().toISOString());
      }

      return jsonResponse({ success: true, row, column, value });
    }

    // --- ADD a new project row ---
    if (payload.action === 'add') {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const newRow = headers.map(h => {
        if (h === 'ID') return generateId(sheet);
        if (h === 'Status') return payload['Status'] || 'Not Started';
        if (h === 'Priority') return payload['Priority'] || 'Medium';
        if (h === 'Leadership Status') return payload['Leadership Status'] || 'Idea';
        if (h === 'Last Updated') return new Date().toISOString();
        return payload[h] || '';
      });
      sheet.appendRow(newRow);
      const newRowNum = sheet.getLastRow();
      return jsonResponse({ success: true, row: newRowNum, id: newRow[0] });
    }

    // --- DELETE a project row ---
    if (payload.action === 'delete') {
      const { row } = payload;
      if (row > 1 && row <= sheet.getLastRow()) {
        sheet.deleteRow(row);
        return jsonResponse({ success: true, deletedRow: row });
      }
      return jsonResponse({ error: 'Invalid row number.' });
    }

    // --- UPDATE a recurring obligation cell ---
    if (payload.action === 'recurring_update') {
      const rSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECURRING_SHEET);
      if (!rSheet) return jsonResponse({ error: 'Recurring sheet not found.' });
      const { row, column, value } = payload;
      const headers = rSheet.getRange(1, 1, 1, rSheet.getLastColumn()).getValues()[0];
      const colIndex = headers.indexOf(column);
      if (colIndex === -1) return jsonResponse({ error: `Column "${column}" not found in Recurring.` });
      rSheet.getRange(row, colIndex + 1).setValue(value);
      return jsonResponse({ success: true, row, column, value });
    }

    // --- ADD a recurring obligation ---
    if (payload.action === 'recurring_add') {
      const rSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECURRING_SHEET);
      if (!rSheet) return jsonResponse({ error: 'Recurring sheet not found.' });
      const headers = rSheet.getRange(1, 1, 1, rSheet.getLastColumn()).getValues()[0];
      const newRow = headers.map(h => {
        if (h === 'ID') return generateRecurringId(rSheet);
        if (h === 'Created') return new Date().toISOString();
        return payload[h] || '';
      });
      rSheet.appendRow(newRow);
      return jsonResponse({ success: true, row: rSheet.getLastRow(), id: newRow[0] });
    }

    // --- DELETE a recurring obligation ---
    if (payload.action === 'recurring_delete') {
      const rSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECURRING_SHEET);
      if (!rSheet) return jsonResponse({ error: 'Recurring sheet not found.' });
      const { row } = payload;
      if (row > 1 && row <= rSheet.getLastRow()) {
        rSheet.deleteRow(row);
        return jsonResponse({ success: true, deletedRow: row });
      }
      return jsonResponse({ error: 'Invalid row number.' });
    }

    // --- COMPLETE a recurring obligation (log history + advance date) ---
    if (payload.action === 'recurring_complete') {
      const rSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECURRING_SHEET);
      const hSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HISTORY_SHEET);
      if (!rSheet || !hSheet) return jsonResponse({ error: 'Recurring or History sheet not found.' });

      const { row, completedBy, notes } = payload;
      const headers = rSheet.getRange(1, 1, 1, rSheet.getLastColumn()).getValues()[0];
      const rowData = rSheet.getRange(row, 1, 1, headers.length).getValues()[0];
      const item = {};
      headers.forEach((h, idx) => { item[h] = rowData[idx]; });

      // Log to history
      hSheet.appendRow([
        item['ID'], item['Title'], completedBy || item['Owner'],
        new Date().toISOString(), item['Next Due'], item['Cadence'], notes || ''
      ]);

      // Calculate next due date
      const cadenceMonths = {
        'Monthly': 1, 'Quarterly': 3, 'Semi-Annual': 6,
        'Annual': 12, '2-Year': 24, '3-Year': 36, '5-Year': 60
      };
      const months = cadenceMonths[item['Cadence']] || 12;
      const currentDue = new Date(item['Next Due']);
      currentDue.setMonth(currentDue.getMonth() + months);
      const nextDue = currentDue.toISOString().slice(0, 10);

      // Update Next Due and Last Completed
      const dueIdx = headers.indexOf('Next Due');
      const lastCompIdx = headers.indexOf('Last Completed');
      if (dueIdx !== -1) rSheet.getRange(row, dueIdx + 1).setValue(nextDue);
      if (lastCompIdx !== -1) rSheet.getRange(row, lastCompIdx + 1).setValue(new Date().toISOString());

      return jsonResponse({ success: true, nextDue, row });
    }

    // --- ADD a calendar event ---
    if (payload.action === 'cal_add') {
      const cSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CALENDAR_SHEET);
      if (!cSheet) return jsonResponse({ error: 'Calendar sheet not found.' });
      const headers = cSheet.getRange(1, 1, 1, cSheet.getLastColumn()).getValues()[0];
      const newRow = headers.map(h => {
        if (h === 'ID') return generateCalId(cSheet);
        if (h === 'Created') return new Date().toISOString();
        return payload[h] || '';
      });
      cSheet.appendRow(newRow);
      return jsonResponse({ success: true, row: cSheet.getLastRow(), id: newRow[0] });
    }

    // --- UPDATE a calendar event cell ---
    if (payload.action === 'cal_update') {
      const cSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CALENDAR_SHEET);
      if (!cSheet) return jsonResponse({ error: 'Calendar sheet not found.' });
      const { row, column, value } = payload;
      const headers = cSheet.getRange(1, 1, 1, cSheet.getLastColumn()).getValues()[0];
      const colIndex = headers.indexOf(column);
      if (colIndex === -1) return jsonResponse({ error: `Column "${column}" not found.` });
      cSheet.getRange(row, colIndex + 1).setValue(value);
      return jsonResponse({ success: true, row, column, value });
    }

    // --- DELETE a calendar event ---
    if (payload.action === 'cal_delete') {
      const cSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CALENDAR_SHEET);
      if (!cSheet) return jsonResponse({ error: 'Calendar sheet not found.' });
      const { row } = payload;
      if (row > 1 && row <= cSheet.getLastRow()) {
        cSheet.deleteRow(row);
        return jsonResponse({ success: true, deletedRow: row });
      }
      return jsonResponse({ error: 'Invalid row number.' });
    }

    // --- ADD a schedule entry ---
    if (payload.action === 'sched_add') {
      const sSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCHEDULE_SHEET);
      if (!sSheet) return jsonResponse({ error: 'Schedule sheet not found.' });
      const headers = sSheet.getRange(1, 1, 1, sSheet.getLastColumn()).getValues()[0];
      const newRow = headers.map(h => {
        if (h === 'ID') return generateSchedId(sSheet);
        return payload[h] || '';
      });
      sSheet.appendRow(newRow);
      return jsonResponse({ success: true, row: sSheet.getLastRow(), id: newRow[0] });
    }

    // --- UPDATE a schedule entry ---
    if (payload.action === 'sched_update') {
      const sSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCHEDULE_SHEET);
      if (!sSheet) return jsonResponse({ error: 'Schedule sheet not found.' });
      const { row, column, value } = payload;
      const headers = sSheet.getRange(1, 1, 1, sSheet.getLastColumn()).getValues()[0];
      const colIndex = headers.indexOf(column);
      if (colIndex === -1) return jsonResponse({ error: `Column "${column}" not found.` });
      sSheet.getRange(row, colIndex + 1).setValue(value);
      return jsonResponse({ success: true, row, column, value });
    }

    // --- DELETE a schedule entry ---
    if (payload.action === 'sched_delete') {
      const sSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCHEDULE_SHEET);
      if (!sSheet) return jsonResponse({ error: 'Schedule sheet not found.' });
      const { row } = payload;
      if (row > 1 && row <= sSheet.getLastRow()) {
        sSheet.deleteRow(row);
        return jsonResponse({ success: true, deletedRow: row });
      }
      return jsonResponse({ error: 'Invalid row number.' });
    }

    // --- BULK ADD schedule entries ---
    if (payload.action === 'sched_bulk_add') {
      const sSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCHEDULE_SHEET);
      if (!sSheet) return jsonResponse({ error: 'Schedule sheet not found.' });
      const headers = sSheet.getRange(1, 1, 1, sSheet.getLastColumn()).getValues()[0];
      const entries = payload.entries || [];
      let added = 0;
      entries.forEach(entry => {
        const newRow = headers.map(h => {
          if (h === 'ID') return generateSchedId(sSheet);
          return entry[h] || '';
        });
        sSheet.appendRow(newRow);
        added++;
      });
      return jsonResponse({ success: true, added });
    }

    // --- ADD a goal ---
    if (payload.action === 'goal_add') {
      const gSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GOALS_SHEET);
      if (!gSheet) return jsonResponse({ error: 'Goals sheet not found. Run setupGoalsSheet() first.' });
      const headers = gSheet.getRange(1, 1, 1, gSheet.getLastColumn()).getValues()[0];
      const newRow = headers.map(h => {
        if (h === 'ID') return generateGoalId(gSheet);
        return payload[h] || '';
      });
      gSheet.appendRow(newRow);
      return jsonResponse({ success: true, row: gSheet.getLastRow(), id: newRow[0] });
    }

    // --- UPDATE a goal ---
    if (payload.action === 'goal_update') {
      const gSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GOALS_SHEET);
      if (!gSheet) return jsonResponse({ error: 'Goals sheet not found.' });
      const { row, column, value } = payload;
      const headers = gSheet.getRange(1, 1, 1, gSheet.getLastColumn()).getValues()[0];
      const colIndex = headers.indexOf(column);
      if (colIndex === -1) return jsonResponse({ error: `Column "${column}" not found.` });
      gSheet.getRange(row, colIndex + 1).setValue(value);
      return jsonResponse({ success: true, row, column, value });
    }

    // --- DELETE a goal ---
    if (payload.action === 'goal_delete') {
      const gSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GOALS_SHEET);
      if (!gSheet) return jsonResponse({ error: 'Goals sheet not found.' });
      const { row } = payload;
      if (row > 1 && row <= gSheet.getLastRow()) {
        gSheet.deleteRow(row);
        return jsonResponse({ success: true, deletedRow: row });
      }
      return jsonResponse({ error: 'Invalid row number.' });
    }

    return jsonResponse({ error: 'Unknown action.' });

  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
}

/**
 * Generate a sequential project ID like PRJ-001
 */
function generateId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'PRJ-001';
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(Boolean);
  const nums = ids.map(id => {
    const match = String(id).match(/PRJ-(\d+)/);
    return match ? parseInt(match[1]) : 0;
  });
  const next = Math.max(...nums, 0) + 1;
  return `PRJ-${String(next).padStart(3, '0')}`;
}

/**
 * Generate a sequential recurring ID like REC-001
 */
function generateRecurringId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'REC-001';
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(Boolean);
  const nums = ids.map(id => {
    const match = String(id).match(/REC-(\d+)/);
    return match ? parseInt(match[1]) : 0;
  });
  const next = Math.max(...nums, 0) + 1;
  return `REC-${String(next).padStart(3, '0')}`;
}

/**
 * Generate a sequential calendar event ID like CAL-001
 */
function generateCalId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'CAL-001';
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(Boolean);
  const nums = ids.map(id => {
    const match = String(id).match(/CAL-(\d+)/);
    return match ? parseInt(match[1]) : 0;
  });
  const next = Math.max(...nums, 0) + 1;
  return `CAL-${String(next).padStart(3, '0')}`;
}

/**
 * Generate a sequential schedule ID like SCH-001
 */
function generateSchedId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'SCH-001';
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(Boolean);
  const nums = ids.map(id => {
    const match = String(id).match(/SCH-(\d+)/);
    return match ? parseInt(match[1]) : 0;
  });
  const next = Math.max(...nums, 0) + 1;
  return `SCH-${String(next).padStart(3, '0')}`;
}

/**
 * Return a CORS-friendly JSON response
 */
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Run this once to create the initial sheet with headers
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  const headers = [
    'ID', 'Project', 'Description', 'Category', 'Owner',
    'Key Stakeholder', 'Priority', 'Status', 'Leadership Status',
    'Est. Time', 'Start Date', 'Due Date', '% Complete',
    'AI Potential?', 'Last Updated', 'Notes'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#00274C')
    .setFontColor('#FFFFFF');

  // Set column widths
  sheet.setColumnWidth(1, 80);   // ID
  sheet.setColumnWidth(2, 250);  // Project
  sheet.setColumnWidth(3, 350);  // Description
  sheet.setColumnWidth(4, 140);  // Category
  sheet.setColumnWidth(5, 140);  // Owner
  sheet.setColumnWidth(6, 160);  // Key Stakeholder
  sheet.setColumnWidth(7, 100);  // Priority
  sheet.setColumnWidth(8, 130);  // Status
  sheet.setColumnWidth(9, 160);  // Leadership Status
  sheet.setColumnWidth(10, 100); // Est. Time
  sheet.setColumnWidth(11, 120); // Start Date
  sheet.setColumnWidth(12, 120); // Due Date
  sheet.setColumnWidth(13, 100); // % Complete
  sheet.setColumnWidth(14, 120); // AI Potential?
  sheet.setColumnWidth(15, 180); // Last Updated
  sheet.setColumnWidth(16, 300); // Notes

  // Data validation for Priority
  const prioRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Critical', 'Urgent', 'Important', 'Medium', 'Low', 'Non-Urgent'])
    .build();
  sheet.getRange(2, 7, 500, 1).setDataValidation(prioRule);

  // Data validation for Status
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Not Started', 'Planning', 'In Progress', 'In Review', 'Complete', 'Delayed'])
    .build();
  sheet.getRange(2, 8, 500, 1).setDataValidation(statusRule);

  // Data validation for Leadership Status
  const lsRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Priority', 'Actively Managed', 'Idea'])
    .build();
  sheet.getRange(2, 9, 500, 1).setDataValidation(lsRule);

  // Data validation for AI Potential?
  const aiRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No', 'Maybe'])
    .build();
  sheet.getRange(2, 14, 500, 1).setDataValidation(aiRule);

  sheet.setFrozenRows(1);

  // Add sample data — columns: ID, Project, Description, Category, Owner, Key Stakeholder, Priority, Status, Leadership Status, Est. Time, Start Date, Due Date, % Complete, AI Potential?, Last Updated, Notes
  const sampleData = [
    ['PRJ-001', 'FY27 Budget Narrative', 'Final review and Dean\'s Office submission', 'Finance', 'Tyler Hughes', 'Dr. Palapattu', 'Critical', 'In Progress', 'Priority', '8 hours', '2026-03-12', '2026-04-01', 75, 'No', new Date().toISOString(), 'Final stretch — all sections complete'],
    ['PRJ-002', 'BizIQ Platform Expansion', 'Scale to surgical departments with new dashboards', 'Analytics', 'Tyler Hughes', 'Dr. Palapattu', 'Urgent', 'In Progress', 'Priority', 'Multi-day', '2026-03-12', '2026-06-30', 40, 'Yes', new Date().toISOString(), '75+ dashboards, 280+ providers served'],
    ['PRJ-003', 'New Physician Onboarding Analytics', 'wRVU projections and P&L modeling for new hires', 'Faculty', 'Tyler Hughes', '', 'Urgent', 'In Progress', 'Actively Managed', '4 hours', '2026-03-12', '2026-05-15', 55, 'No', new Date().toISOString(), 'Dr. Koh visiting faculty arrangement'],
    ['PRJ-004', 'OR Block Reallocation', 'Analyze and optimize OR block allocation across sites', 'Operations', 'Tyler Hughes', '', 'Medium', 'Planning', 'Actively Managed', 'Multi-day', '2026-04-13', '2026-07-01', 10, 'Yes', new Date().toISOString(), 'UH, BCSC, Brighton, Chelsea'],
    ['PRJ-005', 'Fellowship Onboarding Comms', 'Coordinate fellowship onboarding/offboarding communications', 'Education', '', '', 'Medium', 'Not Started', 'Idea', '2 hours', '2026-04-20', '2026-06-01', 0, 'No', new Date().toISOString(), ''],
  ];

  if (sheet.getLastRow() < 2) {
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  }

  SpreadsheetApp.flush();
  Logger.log('Sheet setup complete with sample data.');
}

/**
 * Run this once to create the Recurring and History sheets
 */
function setupRecurringSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Recurring sheet ---
  let rSheet = ss.getSheetByName(RECURRING_SHEET);
  if (!rSheet) {
    rSheet = ss.insertSheet(RECURRING_SHEET);
  }

  const rHeaders = [
    'ID', 'Title', 'Owner', 'Category', 'Cadence',
    'Next Due', 'Alert Lead (Months)', 'Last Completed', 'Notes', 'Created'
  ];

  rSheet.getRange(1, 1, 1, rHeaders.length).setValues([rHeaders]);
  rSheet.getRange(1, 1, 1, rHeaders.length)
    .setFontWeight('bold')
    .setBackground('#00274C')
    .setFontColor('#FFFFFF');

  rSheet.setColumnWidth(1, 80);    // ID
  rSheet.setColumnWidth(2, 280);   // Title
  rSheet.setColumnWidth(3, 150);   // Owner
  rSheet.setColumnWidth(4, 140);   // Category
  rSheet.setColumnWidth(5, 120);   // Cadence
  rSheet.setColumnWidth(6, 120);   // Next Due
  rSheet.setColumnWidth(7, 140);   // Alert Lead
  rSheet.setColumnWidth(8, 160);   // Last Completed
  rSheet.setColumnWidth(9, 300);   // Notes
  rSheet.setColumnWidth(10, 160);  // Created
  rSheet.setFrozenRows(1);

  // --- History sheet ---
  let hSheet = ss.getSheetByName(HISTORY_SHEET);
  if (!hSheet) {
    hSheet = ss.insertSheet(HISTORY_SHEET);
  }

  const hHeaders = [
    'Obligation ID', 'Title', 'Completed By', 'Completed Date',
    'Due Date', 'Cadence', 'Notes'
  ];

  hSheet.getRange(1, 1, 1, hHeaders.length).setValues([hHeaders]);
  hSheet.getRange(1, 1, 1, hHeaders.length)
    .setFontWeight('bold')
    .setBackground('#00274C')
    .setFontColor('#FFFFFF');

  hSheet.setColumnWidth(1, 100);
  hSheet.setColumnWidth(2, 280);
  hSheet.setColumnWidth(3, 150);
  hSheet.setColumnWidth(4, 180);
  hSheet.setColumnWidth(5, 120);
  hSheet.setColumnWidth(6, 120);
  hSheet.setColumnWidth(7, 300);
  hSheet.setFrozenRows(1);

  SpreadsheetApp.flush();
  Logger.log('Recurring and History sheets created.');
}

/**
 * Run this once to create the Calendar Events sheet
 */
function setupCalendarSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cSheet = ss.getSheetByName(CALENDAR_SHEET);
  if (!cSheet) {
    cSheet = ss.insertSheet(CALENDAR_SHEET);
  }

  const headers = [
    'ID', 'Title', 'Date', 'Type', 'Owner', 'Notes', 'Created'
  ];

  cSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  cSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#00274C')
    .setFontColor('#FFFFFF');

  cSheet.setColumnWidth(1, 80);
  cSheet.setColumnWidth(2, 300);
  cSheet.setColumnWidth(3, 120);
  cSheet.setColumnWidth(4, 160);
  cSheet.setColumnWidth(5, 150);
  cSheet.setColumnWidth(6, 350);
  cSheet.setColumnWidth(7, 160);
  cSheet.setFrozenRows(1);

  SpreadsheetApp.flush();
  Logger.log('Calendar Events sheet created.');
}

/**
 * Run this once to create the Provider Schedules sheet
 */
function setupScheduleSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sSheet = ss.getSheetByName(SCHEDULE_SHEET);
  if (!sSheet) {
    sSheet = ss.insertSheet(SCHEDULE_SHEET);
  }

  const headers = [
    'ID', 'Provider', 'Date', 'Assignment', 'Location', 'Time', 'Notes'
  ];

  sSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#00274C')
    .setFontColor('#FFFFFF');

  sSheet.setColumnWidth(1, 80);
  sSheet.setColumnWidth(2, 180);
  sSheet.setColumnWidth(3, 120);
  sSheet.setColumnWidth(4, 160);
  sSheet.setColumnWidth(5, 160);
  sSheet.setColumnWidth(6, 120);
  sSheet.setColumnWidth(7, 300);
  sSheet.setFrozenRows(1);

  SpreadsheetApp.flush();
  Logger.log('Provider Schedules sheet created.');
}

/**
 * Generate a sequential goal ID like GOL-001
 */
function generateGoalId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'GOL-001';
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(Boolean);
  const nums = ids.map(id => {
    const match = String(id).match(/GOL-(\d+)/);
    return match ? parseInt(match[1]) : 0;
  });
  const next = Math.max(...nums, 0) + 1;
  return `GOL-${String(next).padStart(3, '0')}`;
}

/**
 * ONE-TIME: Create the Goals sheet with proper headers and formatting
 * Columns: ID, Title, Domain, Status, Notes, Created
 */
function setupGoalsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let gSheet = ss.getSheetByName(GOALS_SHEET);
  if (gSheet) {
    Logger.log('Goals sheet already exists.');
    return;
  }
  gSheet = ss.insertSheet(GOALS_SHEET);
  const headers = ['ID', 'Title', 'Domain', 'Status', 'Notes', 'Created'];
  gSheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#00274C')
    .setFontColor('#FFFFFF');

  gSheet.setColumnWidth(1, 80);
  gSheet.setColumnWidth(2, 300);
  gSheet.setColumnWidth(3, 140);
  gSheet.setColumnWidth(4, 120);
  gSheet.setColumnWidth(5, 400);
  gSheet.setColumnWidth(6, 120);
  gSheet.setFrozenRows(1);

  // Add data validation for Status
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Not Started', 'In Progress', 'Complete'])
    .build();
  gSheet.getRange(2, 4, 500, 1).setDataValidation(statusRule);

  SpreadsheetApp.flush();
  Logger.log('Goals sheet created.');
}
