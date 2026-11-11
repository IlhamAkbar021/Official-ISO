function getLidarData(forceRefresh) {
  var cache = CacheService.getScriptCache();
  var cached = forceRefresh ? null : cache.get('iso_lidar_data');
  if (cached) {
    try { return JSON.parse(cached); } catch (cacheError) { /* rebuild */ }
  }
  var ss = SpreadsheetApp.openByUrl(DATABASE_URL);
  var sheet = getSheetFlexible(ss, ['lidar', 'LiDAR']);
  if (!sheet) return { error: "Sheet 'lidar' not found in database." };
  
  var range = sheet.getDataRange();
  var displayValues = range.getDisplayValues(); 

  if (displayValues.length <= 2) return { error: "No records found in LiDAR sheet." };
  
  var headerRow = displayValues[1]; 
  var dataRows = displayValues.slice(2); 
  
  var accuracyHeaders = ["Robot ID", "Date", "SW Ver", "Tester", "Option & Range", "Distance (m)", "STD (m)", "Valid Points", "Actual Dist (m)", "Pass/Fail", "Note / Link"];
  
  var accuracyRows = dataRows.map(function(r, rIdx) {
    var accData = r.slice(0, 11);
    return accData.map(function(cellVal, cIdx) {
      var val = cellVal ? cellVal.toString().trim() : '';
      return val;
    });
  });
  
  var colBricks = -1;
  for (var i = 0; i < headerRow.length; i++) {
    var h = (headerRow[i] || '').toString().toLowerCase().trim();
    if (h.indexOf('brick') !== -1 && colBricks === -1) { colBricks = i; }
  }
  if (colBricks === -1) colBricks = 11;

  var tiltHeaders = ["Robot ID", "Date", "SW Ver", "Tester", "Number of Bricks", "Detected Distance (m)", "Height From (cm)", "Height To (cm)", "LiDAR Tilt Status", "Angle From", "Angle To", "Note / Link"];
  
  var tiltRows = [];
  
  function getCellValue(rIdx, cIdx) {
    if (cIdx < 0 || !dataRows[rIdx] || cIdx >= dataRows[rIdx].length) return '';
    var val = dataRows[rIdx][cIdx] ? dataRows[rIdx][cIdx].toString().trim() : '';
    return val;
  }

  var activeRobotId = "";
  var activeDate = "";
  var activeSwVer = "";
  var activeTester = "";

  dataRows.forEach(function(r, rIdx) {
    if (!r || r.join('').trim() === '') return;

    if (r[0] && r[0].toString().trim() !== '') activeRobotId = r[0].toString().trim();
    if (r[1] && r[1].toString().trim() !== '') activeDate = r[1].toString().trim();
    if (r[2] && r[2].toString().trim() !== '') activeSwVer = r[2].toString().trim();
    if (r[3] && r[3].toString().trim() !== '') activeTester = r[3].toString().trim();

    var bricks        = getCellValue(rIdx, colBricks);     
    var detectedDist  = getCellValue(rIdx, colBricks + 1); 
    var heightFrom    = getCellValue(rIdx, colBricks + 2); 
    var heightTo      = getCellValue(rIdx, colBricks + 3); 
    var tiltStatus    = getCellValue(rIdx, colBricks + 4); 
    var angleFrom     = getCellValue(rIdx, colBricks + 5); 
    var angleTo       = getCellValue(rIdx, colBricks + 6); 
    var noteLink      = getCellValue(rIdx, colBricks + 7); 

    if (bricks !== '' || detectedDist !== '' || tiltStatus !== '') {
      tiltRows.push([ activeRobotId, activeDate, activeSwVer, activeTester, bricks, detectedDist, heightFrom, heightTo, tiltStatus, angleFrom, angleTo, noteLink ]);
    }
  });

  var result = { isSplit: true, accuracy: { headers: accuracyHeaders, rows: accuracyRows }, tilt: { headers: tiltHeaders, rows: tiltRows } };
  try { cache.put('iso_lidar_data', JSON.stringify(result), 60); } catch (cacheWriteError) { /* cache is optional */ }
  return result;
}

function normalizeLidarImportPayload_(payload) {
  payload = payload || {};
  var normalized = {};
  Object.keys(payload).forEach(function(key) { normalized[key] = payload[key]; });
  normalized.lidarType = normalized.lidarType === 'tilt' ? 'tilt' : 'accuracy';
  normalized.robotId = String(normalized.robotId || '').trim().toUpperCase().replace(/\s+/g, '-');
  normalized.date = String(normalized.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd')).trim();
  normalized.tester = String(normalized.tester || 'Unassigned').trim();
  normalized.swVer = String(normalized.swVer || '').trim();
  if (normalized.lidarType === 'tilt') {
    var brickCount = Number(normalized.bricks);
    var tiltProfile = null;
    if (isFinite(brickCount) && brickCount >= 0 && brickCount <= 8 && Math.floor(brickCount) === brickCount) {
      tiltProfile = {
        heightFrom: (brickCount - 1) * 10,
        heightTo: brickCount * 10,
        angleFrom: ((brickCount - 4) * 0.38).toFixed(2),
        angleTo: ((brickCount - 3) * 0.38).toFixed(2),
        tiltStatus: brickCount < 4 ? 'Tilt Down, Acceptable' : (brickCount === 4 ? 'Normal' : (brickCount === 5 ? 'Tilt Up, but Acceptable' : 'Tilt Up, Not Acceptable'))
      };
    }
    if (tiltProfile) {
      normalized.heightFrom = String(tiltProfile.heightFrom);
      normalized.heightTo = String(tiltProfile.heightTo);
      normalized.angleFrom = String(tiltProfile.angleFrom);
      normalized.angleTo = String(tiltProfile.angleTo);
      normalized.tiltStatus = tiltProfile.tiltStatus;
    }
  }
  ['runA', 'runB', 'runC'].forEach(function(key) {
    var run = normalized[key] || {};
    normalized[key] = {
      distance: String(run.distance || '').trim().replace(',', '.'),
      std: String(run.std || '0.00').trim().replace(',', '.'),
      validPoints: String(run.validPoints || '9').trim()
    };
  });
  return normalized;
}

function addLidarRecord(payload) {
  try {
    payload = normalizeLidarImportPayload_(payload);
    var ss = SpreadsheetApp.openByUrl(DATABASE_URL);
    var sheet = getSheetFlexible(ss, ['lidar', 'LiDAR']);
    if (!sheet) return { success: false, error: "LiDAR Sheet not found." };

    var lastRow = sheet.getLastRow();
    var numCols = sheet.getLastColumn();
    var writtenFirstRow = lastRow + 1;

    if (payload.lidarType === 'accuracy') {
      var templateStartRow = Math.max(3, lastRow - 2);
      if (lastRow >= 5) {
        var optionValues = sheet.getRange(3, 5, lastRow - 2, 1).getDisplayValues();
        for (var optionIndex = optionValues.length - 3; optionIndex >= 0; optionIndex--) {
          var optionA = String(optionValues[optionIndex][0] || '').toUpperCase();
          var optionB = String(optionValues[optionIndex + 1][0] || '').toUpperCase();
          var optionC = String(optionValues[optionIndex + 2][0] || '').toUpperCase();
          if (/A\s*-?\s*5/.test(optionA) && /B\s*-?\s*10/.test(optionB) && /C\s*-?\s*15/.test(optionC)) {
            templateStartRow = optionIndex + 3;
            break;
          }
        }
      }
      sheet.insertRowsAfter(lastRow, 3);
      var firstRow = lastRow + 1;
      writtenFirstRow = firstRow;
      sheet.getRange(templateStartRow, 1, 3, numCols).copyTo(sheet.getRange(firstRow, 1, 3, numCols));
      var options = ['A - 5', 'B - 10', 'C - 15'];
      var runs = [payload.runA || {}, payload.runB || {}, payload.runC || {}];
      var targets = [5, 10, 15];
      var sequenceValues = [];
      for (var i = 0; i < 3; i++) {
        sequenceValues.push([options[i], runs[i].distance || '', runs[i].std || '', runs[i].validPoints || '']);
      }
      sheet.getRange(firstRow, 5, 3, 4).setValues(sequenceValues);
      if (numCols >= 11) sheet.getRange(firstRow, 11, 3, numCols - 10).clearContent();

      var sharedValues = [payload.robotId || '', payload.date || '', payload.swVer || '', payload.tester || ''];
      sharedValues.forEach(function(value, columnIndex) {
        var sharedRange = sheet.getRange(firstRow, columnIndex + 1, 3, 1);
        sharedRange.breakApart();
        sharedRange.clearContent();
        try {
          sharedRange.mergeVertically();
          sheet.getRange(firstRow, columnIndex + 1).setValue(value);
        } catch (mergeError) {
          sharedRange.setValues([[value], [''], ['']]);
        }
      });
      sheet.getRange(firstRow, 1, 3, Math.min(numCols, 10)).setVerticalAlignment('middle');
    } else if (payload.lidarType === 'tilt') {
      var headerRow = sheet.getRange(2, 1, 1, numCols).getValues()[0];
      var colBricksIdx = -1;
      for (var k = 0; k < headerRow.length; k++) {
        var h = (headerRow[k] || '').toString().toLowerCase();
        if (h.indexOf('brick') !== -1) { colBricksIdx = k; break; }
      }
      if (colBricksIdx === -1) colBricksIdx = 11;
      var values = sheet.getRange(3, 1, Math.max(1, lastRow - 2), Math.min(numCols, 4)).getDisplayValues();
      var targetRow = -1;
      for (var rowOffset = 0; rowOffset < values.length; rowOffset++) {
        var storedDate = String(values[rowOffset][1] || '').trim().replace(/\//g, '-');
        var requestedDate = String(payload.date || '').trim().replace(/\//g, '-');
        if (String(values[rowOffset][0] || '').trim().toUpperCase() === String(payload.robotId || '').trim().toUpperCase() &&
            (!requestedDate || storedDate === requestedDate)) {
          targetRow = rowOffset + 3;
          break;
        }
      }
      if (targetRow === -1) {
        sheet.insertRowsAfter(lastRow, 3);
        targetRow = lastRow + 1;
        sheet.getRange(targetRow, 1, 3, numCols).copyTo(sheet.getRange(targetRow, 1, 3, numCols));
        sheet.getRange(targetRow, 1, 3, 4).setValues([[payload.robotId || '', payload.date || '', payload.swVer || '', payload.tester || ''], ['', '', '', ''], ['', '', '', '']]);
      }
      writtenFirstRow = targetRow;
      sheet.getRange(targetRow, colBricksIdx + 1, 1, 8).setValues([[
        payload.bricks || '', payload.detectedDist || '', payload.heightFrom || '', payload.heightTo || '',
        payload.tiltStatus || '', payload.angleFrom || '', payload.angleTo || '', payload.note || ''
      ]]);
    }

    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove('iso_lidar_data');
    return { success: true, firstRow: writtenFirstRow };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function smartImportLidarRecord(payload) {
  try {
    var normalized = normalizeLidarImportPayload_(payload);
    if (!/^AC3110-DVT\d+(?:-\d+)?$/i.test(normalized.robotId)) {
      return { success: false, error: 'Robot ID must match AC3110-DVTn or AC3110-DVTn-n.' };
    }
    if (normalized.lidarType === 'tilt') {
      if (String(normalized.detectedDist || '').trim() === '' && String(normalized.bricks || '').trim() === '') {
        return { success: false, error: 'Tilt import requires a detected distance or bricks value.' };
      }
      if (!Number.isInteger(Number(normalized.bricks)) || Number(normalized.bricks) < 0 || Number(normalized.bricks) > 8) {
        return { success: false, error: 'Tilt import requires Bricks to be a whole number from 0 to 8.' };
      }
    }
    var result = addLidarRecord(normalized);
    if (!result || !result.success) return result || { success: false, error: 'LiDAR import failed.' };
    return { success: true, imported: { lidarType: normalized.lidarType, robotId: normalized.robotId, date: normalized.date, firstRow: result.firstRow }, data: getLidarData() };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function deleteLidarRecord(payload) {
  try {
    var ss = SpreadsheetApp.openByUrl(DATABASE_URL);
    var sheet = getSheetFlexible(ss, ['lidar', 'LiDAR']);
    if (!sheet) return { success: false, error: "LiDAR Sheet not found." };
    var targetRow = Number(payload.rowIndex) + 3;
    if (payload.robotId) {
      var records = sheet.getRange(3, 1, Math.max(1, sheet.getLastRow() - 2), 2).getDisplayValues();
      var wantedRobot = String(payload.robotId).trim().toUpperCase();
      var wantedDate = String(payload.date || '').trim().replace(/\//g, '-');
      var found = -1;
      for (var i = 0; i < records.length; i++) {
        var actualDate = String(records[i][1] || '').trim().replace(/\//g, '-');
        if (String(records[i][0] || '').trim().toUpperCase() === wantedRobot && (!wantedDate || actualDate === wantedDate)) {
          found = i + 3;
          break;
        }
      }
      if (found < 0) return { success:false, error:'The selected LiDAR record could not be located.' };
      targetRow = found;
    }
    if (payload.lidarType === 'accuracy') {
      sheet.deleteRows(targetRow, 3);
    } else {
      var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      var brickCol = headers.findIndex(function(h){ return String(h || '').toLowerCase().indexOf('brick') !== -1; });
      if (brickCol < 0) brickCol = 11;
      // Tilt belongs to the Accuracy block. Clear only Tilt columns.
      sheet.getRange(targetRow, brickCol + 1, 1, 8).clearContent();
    }
    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove('iso_lidar_data');
    return { success: true };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function updateLidarRecord(payload) {
  try {
    payload = payload || {};
    var ss = SpreadsheetApp.openByUrl(DATABASE_URL);
    var sheet = getSheetFlexible(ss, ['lidar', 'LiDAR']);
    if (!sheet) return { success:false, error:'LiDAR sheet not found.' };
    var row = Number(payload.rowIndex) + 3;
    if (!isFinite(row) || row < 3) return { success:false, error:'Invalid LiDAR row.' };
    var type = payload.lidarType === 'tilt' ? 'tilt' : 'accuracy';
    if (type === 'accuracy') {
      var shared = [[payload.date || '', payload.swVer || '', payload.tester || '']];
      sheet.getRange(row, 2, 1, 3).setValues(shared);
      var accuracyRuns = [payload.runA || {}, payload.runB || {}, payload.runC || {}];
      for (var runIndex = 0; runIndex < 3; runIndex++) {
        sheet.getRange(row + runIndex, 6, 1, 3).setValues([[
          accuracyRuns[runIndex].distance || '',
          accuracyRuns[runIndex].std || '',
          accuracyRuns[runIndex].validPoints || ''
        ]]);
      }
    } else {
      var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      var brickCol = headers.findIndex(function(h){ return String(h || '').toLowerCase().indexOf('brick') !== -1; });
      if (brickCol < 0) brickCol = 11;
      sheet.getRange(row, 2, 1, 3).setValues([[payload.date || '', payload.swVer || '', payload.tester || '']]);
      sheet.getRange(row, brickCol + 1, 1, 8).setValues([[payload.bricks || '', payload.detectedDist || '', payload.heightFrom || '', payload.heightTo || '', payload.tiltStatus || '', payload.angleFrom || '', payload.angleTo || '', payload.note || '']]);
    }
    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove('iso_lidar_data');
    return {success:true, data:getLidarData(true)};
  } catch (err) { return {success:false, error:String(err)}; }
}
