function getLidarData() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('iso_lidar_data');
  if (cached) {
    try { return JSON.parse(cached); } catch (cacheError) { /* rebuild */ }
  }
  var ss = SpreadsheetApp.openByUrl(DATABASE_URL);
  var sheet = getSheetFlexible(ss, ['lidar', 'LiDAR']);
  if (!sheet) return { error: "Sheet 'lidar' not found in database." };
  
  var range = sheet.getDataRange();
  var displayValues = range.getDisplayValues(); 
  var richTextValues = range.getRichTextValues();
  var formulas = range.getFormulas();

  if (displayValues.length <= 2) return { error: "No records found in LiDAR sheet." };
  
  var headerRow = displayValues[1]; 
  var dataRows = displayValues.slice(2); 
  var richTextRows = richTextValues.slice(2);
  var formulaRows = formulas.slice(2);
  
  var accuracyHeaders = ["Robot ID", "Date", "SW Ver", "Tester", "Option & Range", "Distance (m)", "STD (m)", "Valid Points", "Actual Dist (m)", "Pass/Fail", "Note / Link"];
  
  var accuracyRows = dataRows.map(function(r, rIdx) {
    var accData = r.slice(0, 11);
    return accData.map(function(cellVal, cIdx) {
      var val = cellVal ? cellVal.toString().trim() : '';
      var richText = richTextRows[rIdx] ? richTextRows[rIdx][cIdx] : null;
      var url = richText ? richText.getLinkUrl() : null;
      if (!url && formulaRows[rIdx] && formulaRows[rIdx][cIdx]) {
        var match = formulaRows[rIdx][cIdx].match(/HYPERLINK\(\s*["']([^"']+)["']/i);
        if (match && match[1]) url = match[1];
      }
      return url ? url : val;
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
    var richText = richTextRows[rIdx] ? richTextRows[rIdx][cIdx] : null;
    var url = richText ? richText.getLinkUrl() : null;
    if (!url && formulaRows[rIdx] && formulaRows[rIdx][cIdx]) {
      var match = formulaRows[rIdx][cIdx].match(/HYPERLINK\(\s*["']([^"']+)["']/i);
      if (match && match[1]) url = match[1];
    }
    return url ? url : val;
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
      sheet.getRange(templateStartRow, 1, 3, numCols).copyTo(sheet.getRange(firstRow, 1, 3, numCols));
      var options = ['A - 5', 'B - 10', 'C - 15'];
      var runs = [payload.runA || {}, payload.runB || {}, payload.runC || {}];
      var targets = [5, 10, 15];
      var sequenceValues = [];
      for (var i = 0; i < 3; i++) {
        var measured = Number(runs[i].distance);
        var status = isFinite(measured) && Math.abs(measured - targets[i]) <= 0.050000001 ? 'Pass' : 'Fail';
        sequenceValues.push([options[i], runs[i].distance, runs[i].std, runs[i].validPoints, targets[i], status]);
      }
      sheet.getRange(firstRow, 5, 3, 6).setValues(sequenceValues);
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
      sheet.getRange(lastRow, 1, 1, numCols).copyTo(sheet.getRange(lastRow + 1, 1));

      var headerRow = sheet.getRange(2, 1, 1, numCols).getValues()[0];
      var colBricksIdx = -1;
      for (var k = 0; k < headerRow.length; k++) {
        var h = (headerRow[k] || '').toString().toLowerCase();
        if (h.indexOf('brick') !== -1) { colBricksIdx = k; break; }
      }
      if (colBricksIdx === -1) colBricksIdx = 11;

      sheet.getRange(lastRow + 1, 1).setValue(payload.robotId || '');                  
      sheet.getRange(lastRow + 1, 2).setValue(payload.date || '');
      sheet.getRange(lastRow + 1, 3).setValue(payload.swVer || '');
      sheet.getRange(lastRow + 1, 4).setValue(payload.tester || '');                   
      sheet.getRange(lastRow + 1, colBricksIdx + 1).setValue(payload.bricks || '');      
      sheet.getRange(lastRow + 1, colBricksIdx + 2).setValue(payload.detectedDist || '');
      sheet.getRange(lastRow + 1, colBricksIdx + 3).setValue(payload.heightFrom || '');
      sheet.getRange(lastRow + 1, colBricksIdx + 4).setValue(payload.heightTo || '');
      sheet.getRange(lastRow + 1, colBricksIdx + 5).setValue(payload.tiltStatus || '');
      sheet.getRange(lastRow + 1, colBricksIdx + 6).setValue(payload.angleFrom || '');
      sheet.getRange(lastRow + 1, colBricksIdx + 7).setValue(payload.angleTo || '');
      sheet.getRange(lastRow + 1, colBricksIdx + 8).setValue(payload.note || '');
    }

    CacheService.getScriptCache().remove('iso_lidar_data');
    return { success: true, firstRow: lastRow + 1 };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function smartImportLidarRecord(payload) {
  try {
    var normalized = normalizeLidarImportPayload_(payload);
    if (!/^AC3110-DVT\d+(?:-\d+)?$/i.test(normalized.robotId)) {
      return { success: false, error: 'Robot ID must match AC3110-DVTn or AC3110-DVTn-n.' };
    }
    if (normalized.lidarType === 'accuracy') {
      var runs = [normalized.runA || {}, normalized.runB || {}, normalized.runC || {}];
      if (!runs.every(function(run) { return String(run.distance || '').trim() !== '' && isFinite(Number(run.distance)); })) {
        return { success: false, error: 'All three LiDAR measurements (5 m, 10 m, 15 m) must be valid numbers.' };
      }
      if (!runs.every(function(run) { return /^\d+$/.test(String(run.validPoints || '')) && Number(run.validPoints) > 0; })) {
        return { success: false, error: 'Samples must be a positive whole number for all three LiDAR rows.' };
      }
    } else if (String(normalized.detectedDist || '').trim() === '' && String(normalized.angleFrom || '').trim() === '' && String(normalized.angleTo || '').trim() === '') {
      return { success: false, error: 'Tilt import requires a detected distance or angle value.' };
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
    var targetRow = payload.rowIndex + 3;
    if (payload.lidarType === 'accuracy') { sheet.deleteRows(targetRow, 3); } 
    else { sheet.deleteRow(targetRow); }
    CacheService.getScriptCache().remove('iso_lidar_data');
    return { success: true };
  } catch (err) { return { success: false, error: err.toString() }; }
}
