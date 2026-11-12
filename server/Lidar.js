function getLidarData_(forceRefresh) {
  var cached = forceRefresh ? null : ISOApp.Server.Cache.readJson('iso_lidar_data');
  if (cached) return cached;
  var ss = ISOApp.Server.Database.open();
  var sheet = ISOApp.Server.Database.findSheet(ss, ['lidar', 'LiDAR']);
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
  return ISOApp.Server.Cache.writeJson('iso_lidar_data', result, 60);
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
    var brickValue = String(normalized.bricks === undefined || normalized.bricks === null ? '' : normalized.bricks).trim();
    var brickCount = brickValue === '' ? NaN : Number(brickValue);
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

// The LiDAR sheet also contains summaries and formatting far to the right or
// below the records. getLastRow() counts those cells, so use the first empty
// row in the actual record area (A:V) and insert there. This keeps new records
// directly after the existing data while shifting any later auxiliary content
// down intact.
function findLidarInsertRow_(sheet) {
  var firstDataRow = 3;
  var scanLastRow = Math.max(sheet.getLastRow(), firstDataRow);
  var width = Math.min(Math.max(sheet.getLastColumn(), 22), 22);
  var values = sheet.getRange(firstDataRow, 1, scanLastRow - firstDataRow + 1, width).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    var hasRecordData = values[i].some(function(value) { return String(value || '').trim() !== ''; });
    if (!hasRecordData) return firstDataRow + i;
  }
  return scanLastRow + 1;
}

function addLidarRecord_(payload) {
  try {
    payload = normalizeLidarImportPayload_(payload);
    var ss = ISOApp.Server.Database.open();
    var sheet = ISOApp.Server.Database.findSheet(ss, ['lidar', 'LiDAR']);
    if (!sheet) return { success: false, error: "LiDAR Sheet not found." };

    var lastRow = sheet.getLastRow();
    var insertRow = findLidarInsertRow_(sheet);
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
      sheet.insertRowsBefore(insertRow, 3);
      var firstRow = insertRow;
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
      var targetRow = Number(payload.targetRow) >= 3 ? Number(payload.targetRow) : -1;
      for (var rowOffset = 0; rowOffset < values.length; rowOffset++) {
        if (targetRow >= 3) break;
        var storedDate = String(values[rowOffset][1] || '').trim().replace(/\//g, '-');
        var requestedDate = String(payload.date || '').trim().replace(/\//g, '-');
        if (String(values[rowOffset][0] || '').trim().toUpperCase() === String(payload.robotId || '').trim().toUpperCase() &&
            (!requestedDate || storedDate === requestedDate)) {
          targetRow = rowOffset + 3;
          break;
        }
      }
      if (targetRow === -1) {
        sheet.insertRowsBefore(insertRow, 3);
        targetRow = insertRow;
        sheet.getRange(targetRow, 1, 3, numCols).copyTo(sheet.getRange(targetRow, 1, 3, numCols));
        // The copied template may contain vertically merged A:D cells. Break
        // those merges before writing, then recreate the intended shared
        // metadata cells. setValues() across a merged 3-row range throws and
        // makes the Tilt & Angle save appear to do nothing.
        [payload.robotId || '', payload.date || '', payload.swVer || '', payload.tester || ''].forEach(function(value, columnIndex) {
          var sharedRange = sheet.getRange(targetRow, columnIndex + 1, 3, 1);
          sharedRange.breakApart();
          sharedRange.clearContent();
          try {
            sharedRange.mergeVertically();
            sheet.getRange(targetRow, columnIndex + 1).setValue(value);
          } catch (mergeError) {
            sharedRange.setValues([[value], [''], ['']]);
          }
        });
      }
      writtenFirstRow = targetRow;
      sheet.getRange(targetRow, colBricksIdx + 1, 1, 8).setValues([[
        payload.bricks || '', payload.detectedDist || '', payload.heightFrom || '', payload.heightTo || '',
        payload.tiltStatus || '', payload.angleFrom || '', payload.angleTo || '', payload.note || ''
      ]]);
    }

    SpreadsheetApp.flush();
    ISOApp.Server.Cache.remove('iso_lidar_data');
    ISOApp.Server.Cache.remove('iso_static_input_choices');
    return { success: true, firstRow: writtenFirstRow };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function smartImportLidarRecord_(payload) {
  try {
    var normalized = normalizeLidarImportPayload_(payload);
    if (!/^AC3110-DVT\d+(?:-\d+)?$/i.test(normalized.robotId)) {
      return { success: false, error: 'Robot ID must match AC3110-DVTn or AC3110-DVTn-n.' };
    }
    var includeAccuracy = payload && payload.includeAccuracy !== undefined
      ? payload.includeAccuracy === true
      : normalized.lidarType !== 'tilt';
    var includeTilt = payload && payload.includeTilt !== undefined
      ? payload.includeTilt === true
      : normalized.lidarType === 'tilt';
    if (!includeAccuracy && !includeTilt) {
      return { success: false, error: 'Enter at least one Accuracy or Tilt & Angle value before saving.' };
    }
    if (includeTilt) {
      if (String(normalized.detectedDist || '').trim() === '' && String(normalized.bricks || '').trim() === '') {
        return { success: false, error: 'Tilt import requires a detected distance or bricks value.' };
      }
      if (String(normalized.bricks || '').trim() !== '' && (!Number.isInteger(Number(normalized.bricks)) || Number(normalized.bricks) < 0 || Number(normalized.bricks) > 8)) {
        return { success: false, error: 'Tilt import requires Bricks to be a whole number from 0 to 8.' };
      }
    }
    var accuracyResult = null;
    if (includeAccuracy) {
      var accuracyPayload = Object.assign({}, normalized, { lidarType: 'accuracy' });
      accuracyResult = addLidarRecord_(accuracyPayload);
      if (!accuracyResult || !accuracyResult.success) return accuracyResult || { success: false, error: 'Accuracy record save failed.' };
    }
    var tiltResult = null;
    if (includeTilt) {
      var tiltPayload = normalizeLidarImportPayload_(Object.assign({}, payload, {
        lidarType: 'tilt',
        targetRow: accuracyResult && accuracyResult.firstRow ? accuracyResult.firstRow : payload.targetRow
      }));
      tiltResult = addLidarRecord_(tiltPayload);
      if (!tiltResult || !tiltResult.success) return tiltResult || { success: false, error: 'Tilt & Angle record save failed.' };
    }
    var finalResult = accuracyResult || tiltResult;
    return { success: true, imported: { lidarType: includeAccuracy && includeTilt ? 'combined' : (includeTilt ? 'tilt' : 'accuracy'), robotId: normalized.robotId, date: normalized.date, firstRow: finalResult.firstRow }, data: getLidarData_(true) };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function deleteLidarRecord_(payload) {
  try {
    var ss = ISOApp.Server.Database.open();
    var sheet = ISOApp.Server.Database.findSheet(ss, ['lidar', 'LiDAR']);
    if (!sheet) return { success: false, error: "LiDAR Sheet not found." };
    var targetRow = Number(payload.rowIndex) + 3;
    if (payload.robotId) {
      var records = sheet.getRange(3, 1, Math.max(1, sheet.getLastRow() - 2), 2).getDisplayValues();
      var wantedRobot = String(payload.robotId).trim().toUpperCase();
      var wantedDate = String(payload.date || '').trim().replace(/\//g, '-');
      var indexedRobot = targetRow >= 3 && targetRow <= sheet.getLastRow() ? String(records[targetRow - 3][0] || '').trim().toUpperCase() : '';
      var indexedDate = targetRow >= 3 && targetRow <= sheet.getLastRow() ? String(records[targetRow - 3][1] || '').trim().replace(/\//g, '-') : '';
      // Accuracy rows carry the raw data index from the UI. Prefer that exact
      // physical row when its identity matches, avoiding deletion of the first
      // duplicate Robot ID/date elsewhere in the sheet.
      if (payload.lidarType === 'accuracy' && indexedRobot === wantedRobot && (!wantedDate || indexedDate === wantedDate)) {
        // Keep targetRow from the UI index.
      } else {
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
    ISOApp.Server.Cache.remove('iso_lidar_data');
    ISOApp.Server.Cache.remove('iso_static_input_choices');
    return { success: true };
  } catch (err) { return { success: false, error: err.toString() }; }
}

/** LiDAR orchestration service used by public RPC adapters in Code.js. */
ISOApp.Services = ISOApp.Services || {};
ISOApp.Services.Lidar = (function() {
  function owns_(object, key) {
    return Object.prototype.hasOwnProperty.call(object || {}, key);
  }

  function valueOrExisting_(object, key, existingValue) {
    return owns_(object, key) ? object[key] : existingValue;
  }

  function normalizedDate_(value) {
    return String(value || '').trim().replace(/\//g, '-');
  }

  function accuracyRow_(sheet, payload) {
    var row = Number(payload.rowIndex) + 3;
    return isFinite(row) && row >= 3 && row <= sheet.getLastRow() ? row : -1;
  }

  function tiltRow_(sheet, payload, brickColumn) {
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return -1;
    var values = sheet.getRange(3, 1, lastRow - 2, sheet.getLastColumn()).getDisplayValues();
    var tiltIndex = Number(payload.rowIndex);
    var visibleTiltIndex = -1;

    // rowIndex on the Tilt table is an index in the filtered Tilt dataset,
    // not a physical Sheet row. Resolve that mapping before using identity.
    if (isFinite(tiltIndex) && tiltIndex >= 0) {
      for (var i = 0; i < values.length; i++) {
        var hasTilt = String(values[i][brickColumn] || '').trim() !== '' ||
          String(values[i][brickColumn + 1] || '').trim() !== '' ||
          String(values[i][brickColumn + 4] || '').trim() !== '';
        if (hasTilt) {
          visibleTiltIndex++;
          if (visibleTiltIndex === tiltIndex) return i + 3;
        }
      }
    }

    var wantedRobot = String(payload.robotId || '').trim().toUpperCase();
    var wantedDate = normalizedDate_(payload.date);
    for (var rowIndex = 0; rowIndex < values.length; rowIndex++) {
      var actualRobot = String(values[rowIndex][0] || '').trim().toUpperCase();
      var actualDate = normalizedDate_(values[rowIndex][1]);
      if (actualRobot === wantedRobot && (!wantedDate || actualDate === wantedDate)) return rowIndex + 3;
    }
    return -1;
  }

  function brickColumn_(sheet) {
    var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i] || '').toLowerCase().indexOf('brick') !== -1) return i;
    }
    return 11;
  }

  function updateAccuracy_(sheet, payload) {
    var row = accuracyRow_(sheet, payload);
    if (row < 3) throw new Error('The selected LiDAR accuracy record could not be located.');

    var sharedExisting = sheet.getRange(row, 1, 1, 4).getValues()[0];
    sheet.getRange(row, 1, 1, 4).setValues([[
      valueOrExisting_(payload, 'robotId', sharedExisting[0]),
      valueOrExisting_(payload, 'date', sharedExisting[1]),
      valueOrExisting_(payload, 'swVer', sharedExisting[2]),
      valueOrExisting_(payload, 'tester', sharedExisting[3])
    ]]);

    var runs = [payload.runA || {}, payload.runB || {}, payload.runC || {}];
    var currentRuns = sheet.getRange(row, 6, 3, 3).getValues();
    var nextRuns = runs.map(function(run, index) {
      return [
        valueOrExisting_(run, 'distance', currentRuns[index][0]),
        valueOrExisting_(run, 'std', currentRuns[index][1]),
        valueOrExisting_(run, 'validPoints', currentRuns[index][2])
      ];
    });
    sheet.getRange(row, 6, 3, 3).setValues(nextRuns);
    return row;
  }

  function updateTilt_(sheet, payload, preferredRow) {
    payload = normalizeLidarImportPayload_(payload);
    var brickColumn = brickColumn_(sheet);
    var row = preferredRow >= 3 ? preferredRow : tiltRow_(sheet, payload, brickColumn);
    if (row < 3) throw new Error('The selected LiDAR tilt record could not be located.');

    var sharedExisting = sheet.getRange(row, 1, 1, 4).getValues()[0];
    sheet.getRange(row, 1, 1, 4).setValues([[
      valueOrExisting_(payload, 'robotId', sharedExisting[0]),
      valueOrExisting_(payload, 'date', sharedExisting[1]),
      valueOrExisting_(payload, 'swVer', sharedExisting[2]),
      valueOrExisting_(payload, 'tester', sharedExisting[3])
    ]]);

    var current = sheet.getRange(row, brickColumn + 1, 1, 8).getValues()[0];
    var keys = ['bricks', 'detectedDist', 'heightFrom', 'heightTo', 'tiltStatus', 'angleFrom', 'angleTo', 'note'];
    var next = keys.map(function(key, index) { return valueOrExisting_(payload, key, current[index]); });
    sheet.getRange(row, brickColumn + 1, 1, 8).setValues([next]);
    return row;
  }

  function save(rowData) {
    try {
      rowData = rowData || {};
      if (rowData.action !== 'updateLidarRow') return ISOApp.Server.Result.fail('Unsupported LiDAR save action.');
      if (!rowData.accuracyPayload && !rowData.tiltPayload) {
        return ISOApp.Server.Result.fail('No LiDAR row data was provided.');
      }

      var spreadsheet = ISOApp.Server.Database.open();
      var sheet = ISOApp.Server.Database.requireSheet(spreadsheet, ['lidar', 'LiDAR'], 'LiDAR');
      var accuracyRow = -1;
      var updated = [];

      if (rowData.accuracyPayload) {
        var accuracyPayload = rowData.accuracyPayload;
        var requestedAccuracyRow = Number(accuracyPayload.rowIndex) + 3;
        var shouldCreateAccuracy = accuracyPayload.isNew === true || !isFinite(requestedAccuracyRow) || requestedAccuracyRow < 3 || requestedAccuracyRow > sheet.getLastRow();
        if (shouldCreateAccuracy) {
          var createdAccuracy = addLidarRecord_(accuracyPayload);
          if (!createdAccuracy || !createdAccuracy.success) throw new Error(createdAccuracy && createdAccuracy.error || 'LiDAR accuracy record could not be appended.');
          accuracyRow = Number(createdAccuracy.firstRow || -1);
          updated.push({ type: 'accuracy', row: accuracyRow, created: true });
        } else {
          accuracyRow = updateAccuracy_(sheet, accuracyPayload);
          updated.push({ type: 'accuracy', row: accuracyRow, created: false });
        }
      }
      if (rowData.tiltPayload) {
        var tiltPayload = rowData.tiltPayload;
        var shouldCreateTilt = tiltPayload.isNew === true;
        var tiltRow = -1;
        if (!shouldCreateTilt) {
          var tiltBrickColumn = brickColumn_(sheet);
          var normalizedTiltPayload = normalizeLidarImportPayload_(tiltPayload);
          var existingTiltRow = tiltRow_(sheet, normalizedTiltPayload, tiltBrickColumn);
          if (existingTiltRow >= 3) tiltRow = updateTilt_(sheet, tiltPayload, accuracyRow);
        }
        if (tiltRow < 3) {
          var createdTilt = addLidarRecord_(tiltPayload);
          if (!createdTilt || !createdTilt.success) throw new Error(createdTilt && createdTilt.error || 'LiDAR tilt record could not be appended.');
          tiltRow = Number(createdTilt.firstRow || -1);
          updated.push({ type: 'tilt', row: tiltRow, created: true });
        } else {
          updated.push({ type: 'tilt', row: tiltRow, created: false });
        }
      }

      SpreadsheetApp.flush();
      ISOApp.Server.Cache.remove('iso_lidar_data');
      ISOApp.Server.Cache.remove('iso_static_input_choices');
      return ISOApp.Server.Result.ok({ updated: updated, data: getLidarData_(true) });
    } catch (error) {
      return ISOApp.Server.Result.fail(error, 'LiDAR save failed.');
    }
  }

  function update(payload) {
    payload = payload || {};
    var request = { action: 'updateLidarRow' };
    if (payload.lidarType === 'tilt') request.tiltPayload = payload;
    else request.accuracyPayload = payload;
    return save(request);
  }

  return Object.freeze({
    list:getLidarData_,
    create:addLidarRecord_,
    smartImport:smartImportLidarRecord_,
    remove:deleteLidarRecord_,
    update:update,
    save:save
  });
})();
