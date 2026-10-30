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

function addLidarRecord(payload) {
  try {
    var ss = SpreadsheetApp.openByUrl(DATABASE_URL);
    var sheet = getSheetFlexible(ss, ['lidar', 'LiDAR']);
    if (!sheet) return { success: false, error: "LiDAR Sheet not found." };

    var lastRow = sheet.getLastRow();
    var numCols = sheet.getLastColumn();

    if (payload.lidarType === 'accuracy') {
      var templateStartRow = Math.max(3, lastRow - 2);
      sheet.getRange(templateStartRow, 1, 3, numCols).copyTo(sheet.getRange(lastRow + 1, 1));

      var options = ['A - 5', 'B - 10', 'C - 15'];
      var runs = [payload.runA || {}, payload.runB || {}, payload.runC || {}];

      for (var i = 0; i < 3; i++) {
        var rowIdx = lastRow + 1 + i;
        sheet.getRange(rowIdx, 1).setValue(payload.robotId || '');       
        sheet.getRange(rowIdx, 2).setValue(payload.date || '');          
        sheet.getRange(rowIdx, 3).setValue(payload.swVer || '');         
        sheet.getRange(rowIdx, 4).setValue(payload.tester || '');        
        sheet.getRange(rowIdx, 5).setValue(options[i]);                  
        sheet.getRange(rowIdx, 6).setValue(runs[i].distance || '');      
        sheet.getRange(rowIdx, 7).setValue(runs[i].std || '');           
        sheet.getRange(rowIdx, 8).setValue(runs[i].validPoints || '');   
      }
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
      sheet.getRange(lastRow + 1, 4).setValue(payload.tester || '');                   
      sheet.getRange(lastRow + 1, colBricksIdx + 1).setValue(payload.bricks || '');      
      sheet.getRange(lastRow + 1, colBricksIdx + 2).setValue(payload.detectedDist || '');
    }

    CacheService.getScriptCache().remove('iso_lidar_data');
    return { success: true };
  } catch (err) { return { success: false, error: err.toString() }; }
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
