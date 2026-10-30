function getOverviewData() {
  try {
    var overviewCache = CacheService.getScriptCache();
    var cachedOverview = overviewCache.get('iso_overview_v3');
    if (cachedOverview) {
      try { return JSON.parse(cachedOverview); } catch (cacheParseError) { /* rebuild below */ }
    }
    var ss = SpreadsheetApp.openByUrl(DATABASE_URL);
    
    var sheet = null;
    var candidateNames = ['Overview P/F or data', 'Overview P/F', 'Overview'];
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var sName = sheets[i].getName().trim().toLowerCase();
      for (var j = 0; j < candidateNames.length; j++) {
        if (sName === candidateNames[j].toLowerCase() || sName.indexOf(candidateNames[j].toLowerCase()) !== -1) {
          sheet = sheets[i]; break;
        }
      }
      if (sheet) break;
    }

    if (!sheet) return { success: false, error: "Overview P/F sheet not found in database." };

    var data = sheet.getDataRange().getDisplayValues();
    if (data.length < 5) return { success: false, error: "Overview sheet is empty." };

    var robots = [];
    var tests = [];
    var robotCols = {}; 
    
    var cols = {
      testItem: -1, purpose: -1, sample: -1, passRate: -1, failRate: -1,
      max: -1, min: -1, avg: -1, std: -1, avgPlus: -1, avgMinus: -1,
      summary: -1, sugSW: -1, sugHW: -1, actionPP: -1, note: -1
    };

    // 1. Dynamic header mapping
    for (var r = 0; r < 5; r++) {
      for (var c = 0; c < data[r].length; c++) {
        var val = data[r][c].toString().trim().toLowerCase();
        if (!val) continue;
        
        if (val === 'test item') cols.testItem = c;
        if (val === 'purpose') cols.purpose = c;
        if (val.indexOf('sample size') !== -1 || val.indexOf('robots have been tested') !== -1) cols.sample = c;
        if (val === 'pass rate(%)') cols.passRate = c;
        if (val === 'fail rate(%)') cols.failRate = c;
        if (val === 'max') cols.max = c;
        if (val === 'min') cols.min = c;
        if (val === 'avg') cols.avg = c;
        if (val === 'std') cols.std = c;
        if (val === 'avg+3std') cols.avgPlus = c;
        if (val === 'avg-3std') cols.avgMinus = c;
        if (val === 'summary') cols.summary = c;
        if (val === 'suggestion to sw') cols.sugSW = c;
        if (val === 'suggestion to hw') cols.sugHW = c;
        if (val.indexOf('following action of pp') !== -1) cols.actionPP = c;
        if (val.indexOf('note from meeting') !== -1) cols.note = c;

        if (val.toUpperCase().indexOf('DVT') !== -1 && val.length < 15) {
          if (!robotCols[c]) {
            robotCols[c] = val;
            robots.push({
              name: val, col: c, sw: 'N/A', task: 'N/A', temp: 'N/A', hum: 'N/A',
              owner: 'N/A', rbtTicket: 'N/A', robotIp: 'N/A', issueNote: 'N/A', hwFix: 'N/A'
            });
          }
        }
      }
    }

    if (cols.testItem === -1) cols.testItem = 4;
    if (cols.purpose === -1) cols.purpose = 7;

    // 2. Robot Metadata
    for (var r = 0; r < 25; r++) {
      if (!data[r]) continue;
      var rowLabel = data[r].join(' ').toLowerCase();
      var isSW = rowLabel.indexOf('sw update') !== -1 || rowLabel.indexOf('sw release') !== -1;
      var isTask = rowLabel.indexOf('current main task') !== -1;
      var isOwner = rowLabel.indexOf('owner in lab') !== -1 || rowLabel === 'owner';
      var isRbtTicket = rowLabel.indexOf('rbt ticket') !== -1 || rowLabel.indexOf('rb ticket') !== -1;
      var isRobotIp = rowLabel.indexOf('robot ip') !== -1;
      var isIssueNote = rowLabel.indexOf('issue note') !== -1;
      var isHwFix = rowLabel.indexOf('hw component') !== -1;
      var isTemp = rowLabel.indexOf('temperature') !== -1;
      var isHum = rowLabel.indexOf('humidity') !== -1;

      if (isSW || isTask || isOwner || isRbtTicket || isRobotIp || isIssueNote || isHwFix || isTemp || isHum) {
        robots.forEach(function(rob) {
          var cellVal = data[r][rob.col] ? data[r][rob.col].toString().trim() : '';
          if (isSW && cellVal) rob.sw = cellVal;
          if (isTask && cellVal) rob.task = cellVal;
          if (isOwner && cellVal) rob.owner = cellVal;
          if (isRbtTicket && cellVal) rob.rbtTicket = cellVal;
          if (isRobotIp && cellVal) rob.robotIp = cellVal;
          if (isIssueNote && cellVal) rob.issueNote = cellVal;
          if (isHwFix && cellVal) rob.hwFix = cellVal;
          if (isTemp && cellVal) rob.temp = cellVal;
          if (isHum && cellVal) rob.hum = cellVal;
        });
      }
    }

    function getVal(arr, index) { return (index !== -1 && arr[index]) ? arr[index].toString().trim() : ''; }

    // 3. Test Rows with row index preservation
    var currentTestCategory = "General Validation";
    for (var r = 10; r < data.length; r++) {
      var tItem = getVal(data[r], cols.testItem);
      var tPurp = getVal(data[r], cols.purpose);
      if (tItem !== '') currentTestCategory = tItem;

      var hasResult = false;
      var results = {};
      robots.forEach(function(rob) {
        var res = getVal(data[r], rob.col);
        results[rob.name] = res;
        if (res !== '') hasResult = true;
      });

      // Keep continuation rows as long as they contain robot results. Several
      // A&R and dynamic ISO blocks use merged labels, leaving later label cells
      // blank even though the rows contain valid test data.
      if (hasResult && currentTestCategory) {
        tests.push({
          sheetRow: r + 1, // 1-based row index in Google Sheet
          item: currentTestCategory,
          purpose: tPurp !== '' ? tPurp : currentTestCategory,
          purposeCol: cols.purpose,
          sample: getVal(data[r], cols.sample),
          passRate: getVal(data[r], cols.passRate),
          failRate: getVal(data[r], cols.failRate),
          max: getVal(data[r], cols.max),
          min: getVal(data[r], cols.min),
          avg: getVal(data[r], cols.avg),
          std: getVal(data[r], cols.std),
          avgPlus: getVal(data[r], cols.avgPlus),
          avgMinus: getVal(data[r], cols.avgMinus),
          summary: getVal(data[r], cols.summary),
          sugSW: getVal(data[r], cols.sugSW),
          sugHW: getVal(data[r], cols.sugHW),
          actionPP: getVal(data[r], cols.actionPP),
          note: getVal(data[r], cols.note),
          results: results
        });
      }
    }

    var overviewResult = { success: true, robots: robots, tests: tests };
    try { overviewCache.put('iso_overview_v3', JSON.stringify(overviewResult), 60); } catch (cacheWriteError) { /* cache is optional */ }
    return overviewResult;
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// BATCH UPDATER: Direct in-cell edits written simultaneously to Google Sheets
function batchUpdateOverview(updates) {
  try {
    var ss = SpreadsheetApp.openByUrl(DATABASE_URL);
    var candidateNames = ['Overview P/F or data', 'Overview P/F', 'Overview'];
    var sheet = null;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var sName = sheets[i].getName().trim().toLowerCase();
      for (var j = 0; j < candidateNames.length; j++) {
        if (sName === candidateNames[j].toLowerCase() || sName.indexOf(candidateNames[j].toLowerCase()) !== -1) {
          sheet = sheets[i]; break;
        }
      }
      if (sheet) break;
    }

    if (!sheet) return { success: false, error: "Overview sheet not found." };
    if (!updates || updates.length === 0) return { success: true, count: 0 };

    updates.forEach(function(u) {
      if (u.row && u.col !== undefined) {
        sheet.getRange(u.row, u.col + 1).setValue(u.val);
      }
    });

    try { CacheService.getScriptCache().remove('iso_overview_v2'); } catch (cacheError) { /* optional */ }

    return { success: true, count: updates.length };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// Add a new robot column to the Overview sheet from the dashboard UI.
function addOverviewRobot(robot) {
  try {
    robot = robot || {};
    var name = String(robot.name || '').trim().replace(/\s+/g, ' ');
    if (!name) return { success: false, error: 'Robot name is required.' };
    if (!/dvt\s*[-_]?\s*\d+/i.test(name)) return { success: false, error: 'Robot name must include a DVT cohort.' };
    if (/^dvt\s*[-_]?\s*\d+/i.test(name)) name = 'AC3110-' + name;

    var ss = SpreadsheetApp.openByUrl(DATABASE_URL);
    var candidateNames = ['Overview P/F or data', 'Overview P/F', 'Overview'];
    var sheet = null;
    ss.getSheets().some(function(s) {
      var lower = s.getName().trim().toLowerCase();
      var matched = candidateNames.some(function(candidate) { return lower === candidate.toLowerCase() || lower.indexOf(candidate.toLowerCase()) !== -1; });
      if (matched) { sheet = s; return true; }
      return false;
    });
    if (!sheet) return { success: false, error: 'Overview sheet not found.' };

    var values = sheet.getDataRange().getDisplayValues();
    var headerRow = -1;
    var existing = {};
    var robotColumns = [];
    for (var r = 0; r < Math.min(5, values.length); r++) {
      for (var c = 0; c < values[r].length; c++) {
        var cell = String(values[r][c] || '').trim();
        var cohortMatch = cell.match(/dvt\s*[-_]?\s*(\d+)/i);
        if (cohortMatch) {
          if (headerRow === -1) headerRow = r + 1;
          existing[cell.toLowerCase()] = true;
          robotColumns.push({ column: c + 1, cohort: 'dvt' + cohortMatch[1], name: cell });
        }
      }
    }
    if (headerRow === -1) return { success: false, error: 'Robot header row could not be detected.' };

    var newRobotMatch = name.match(/dvt\s*[-_]?\s*(\d+)/i);
    var newCohort = 'dvt' + newRobotMatch[1];
    var hasSequence = /[-_\s]\d+\s*$/.test(name);
    if (!hasSequence) {
      var nextSequence = 1;
      robotColumns.forEach(function(item) {
        if (item.cohort !== newCohort) return;
        var existingMatch = String(item.name || '').match(/[-_\s](\d+)\s*$/);
        if (existingMatch) nextSequence = Math.max(nextSequence, Number(existingMatch[1]) + 1);
      });
      name += '-' + nextSequence;
    }
    if (existing[name.toLowerCase()]) return { success: false, error: 'This robot already exists.' };

    var insertAfter = 0;
    robotColumns.forEach(function(item) {
      if (item.cohort === newCohort && item.column > insertAfter) insertAfter = item.column;
    });
    // For a new cohort (for example DVT3), place it after the last robot
    // column instead of after unrelated columns elsewhere in the sheet.
    if (!insertAfter) {
      robotColumns.forEach(function(item) {
        if (item.column > insertAfter) insertAfter = item.column;
      });
    }
    if (!insertAfter) return { success: false, error: 'Last robot column could not be detected.' };
    var templateColumn = insertAfter;
    sheet.insertColumnAfter(insertAfter);
    var newCol = insertAfter + 1;
    var maxRows = sheet.getMaxRows();
    if (templateColumn > 0) {
      var templateRange = sheet.getRange(1, templateColumn, maxRows, 1);
      var newRange = sheet.getRange(1, newCol, maxRows, 1);
      templateRange.copyTo(newRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      templateRange.copyTo(newRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
      sheet.setColumnWidth(newCol, sheet.getColumnWidth(templateColumn));
    }

    // Keep the robot header visually identical to the previous robot. The
    // second header row contains the result type (normally "pass/fail") and
    // must be copied explicitly because format/data-validation copy does not
    // copy cell content.
    var headerLabel = name.replace(/^AC3110-/i, 'AC3110\n');
    sheet.getRange(headerRow, newCol).setValue(headerLabel).setWrap(true);
    if (headerRow < maxRows) {
      var resultHeader = sheet.getRange(headerRow + 1, templateColumn).getDisplayValue().trim();
      if (!resultHeader) resultHeader = 'pass/fail';
      sheet.getRange(headerRow + 1, newCol).setValue(resultHeader).setWrap(true);
    }

    // Re-apply the template backgrounds/alignment to the full new column so
    // the inserted DVT column keeps the same visual blocks as DVT1/DVT2.
    sheet.getRange(1, templateColumn, maxRows, 1)
      .copyTo(sheet.getRange(1, newCol, maxRows, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    for (var row = 0; row < values.length; row++) {
      var rowText = values[row].join(' ').toLowerCase();
      if (rowText.indexOf('sw update') !== -1 || rowText.indexOf('sw release') !== -1) sheet.getRange(row + 1, newCol).setValue(String(robot.sw || '').trim());
      else if (rowText.indexOf('temperature') !== -1) sheet.getRange(row + 1, newCol).setValue(String(robot.temp || '').trim());
      else if (rowText.indexOf('humidity') !== -1) sheet.getRange(row + 1, newCol).setValue(String(robot.humidity || '').trim());
    }
    try { CacheService.getScriptCache().remove('iso_overview_v2'); } catch (cacheError) { /* optional */ }
    return { success: true, name: name, column: newCol };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function updateOverviewRobotMetadata(payload) {
  try {
    payload = payload || {};
    var robotName = String(payload.robotName || '').trim();
    if (!robotName) return { success: false, error: 'Robot name is required.' };
    var ss = SpreadsheetApp.openByUrl(DATABASE_URL);
    var candidateNames = ['Overview P/F or data', 'Overview P/F', 'Overview'];
    var sheet = null;
    ss.getSheets().some(function(s) {
      var lower = s.getName().trim().toLowerCase();
      var matched = candidateNames.some(function(candidate) { return lower === candidate.toLowerCase() || lower.indexOf(candidate.toLowerCase()) !== -1; });
      if (matched) { sheet = s; return true; }
      return false;
    });
    if (!sheet) return { success: false, error: 'Overview sheet not found.' };

    var values = sheet.getDataRange().getDisplayValues();
    var robotColumn = -1;
    for (var r = 0; r < Math.min(5, values.length); r++) {
      for (var c = 0; c < values[r].length; c++) {
        if (String(values[r][c] || '').trim().toLowerCase() === robotName.toLowerCase()) { robotColumn = c + 1; break; }
      }
      if (robotColumn !== -1) break;
    }
    if (robotColumn === -1) return { success: false, error: 'Robot column not found.' };

    for (var row = 0; row < values.length; row++) {
      var rowText = values[row].join(' ').toLowerCase();
      if (rowText.indexOf('sw update') !== -1 || rowText.indexOf('sw release') !== -1) sheet.getRange(row + 1, robotColumn).setValue(String(payload.sw || '').trim());
      else if (rowText.indexOf('temperature') !== -1) sheet.getRange(row + 1, robotColumn).setValue(String(payload.temp || '').trim());
      else if (rowText.indexOf('humidity') !== -1) sheet.getRange(row + 1, robotColumn).setValue(String(payload.humidity || '').trim());
    }
    try { CacheService.getScriptCache().remove('iso_overview_v2'); } catch (cacheError) { /* optional */ }
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function deleteOverviewRobot(robotName) {
  try {
    robotName = String(robotName || '').trim();
    if (!robotName) return { success: false, error: 'Robot name is required.' };
    var ss = SpreadsheetApp.openByUrl(DATABASE_URL);
    var candidateNames = ['Overview P/F or data', 'Overview P/F', 'Overview'];
    var sheet = null;
    ss.getSheets().some(function(s) {
      var lower = s.getName().trim().toLowerCase();
      var matched = candidateNames.some(function(candidate) { return lower === candidate.toLowerCase() || lower.indexOf(candidate.toLowerCase()) !== -1; });
      if (matched) { sheet = s; return true; }
      return false;
    });
    if (!sheet) return { success: false, error: 'Overview sheet not found.' };

    var values = sheet.getDataRange().getDisplayValues();
    var robotColumn = -1;
    for (var r = 0; r < Math.min(5, values.length); r++) {
      for (var c = 0; c < values[r].length; c++) {
        if (String(values[r][c] || '').trim().toLowerCase() === robotName.toLowerCase()) { robotColumn = c + 1; break; }
      }
      if (robotColumn !== -1) break;
    }
    if (robotColumn === -1) return { success: false, error: 'Robot column not found.' };
    sheet.deleteColumn(robotColumn);
    try { CacheService.getScriptCache().remove('iso_overview_v2'); } catch (cacheError) { /* optional */ }
    return { success: true, name: robotName };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}
