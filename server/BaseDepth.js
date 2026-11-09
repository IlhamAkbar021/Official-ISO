/** Base Depth Camera module entry point. Shared parsing remains in Modules.js. */
function getBaseDepthData() { return getModuleData('basedepth'); }

var BASE_DEPTH_TOLERANCE_M_ = 0.05;

function findBaseDepthColumn_(headers, patterns, fallback) {
  for (var column = 0; column < headers.length; column++) {
    var label = String(headers[column] || '').toLowerCase().replace(/[_-]+/g, ' ');
    for (var patternIndex = 0; patternIndex < patterns.length; patternIndex++) {
      if (patterns[patternIndex].test(label)) return column;
    }
  }
  return fallback !== undefined && fallback < headers.length ? fallback : -1;
}

function getBaseDepthColumnMap_(headers) {
  return {
    robot: findBaseDepthColumn_(headers, [/robot/, /unit/, /serial/], 0),
    date: findBaseDepthColumn_(headers, [/^date/, /test date/], 1),
    sw: findBaseDepthColumn_(headers, [/^sw\b/, /software.*version/, /firmware/], 2),
    tester: findBaseDepthColumn_(headers, [/tester/, /operator/, /engineer/], 3),
    obstacle: findBaseDepthColumn_(headers, [/obstacle.*size/, /object.*size/, /brick.*size/, /obstacle/, /setup/], 4),
    point: findBaseDepthColumn_(headers, [/test.*point/, /^point/, /^run\b/, /^trial\b/, /sequence/, /^#$/], 5),
    scriptDist: findBaseDepthColumn_(headers, [/script.*dist/, /command.*dist/, /program.*dist/, /set.*dist/], 6),
    realDist: findBaseDepthColumn_(headers, [/real.*dist/, /actual.*dist/, /measured.*dist/], 7),
    coordX: findBaseDepthColumn_(headers, [/coord.*\bx\b/, /\bx\b.*coord/, /position.*\bx\b/], 8),
    coordY: findBaseDepthColumn_(headers, [/coord.*\by\b/, /\by\b.*coord/, /position.*\by\b/], 9),
    coord: findBaseDepthColumn_(headers, [/coordinate/, /^coord$/], -1),
    accuracy: findBaseDepthColumn_(headers, [/accuracy/, /percent/, /rate/], 10),
    delta: findBaseDepthColumn_(headers, [/delta/, /difference/, /\bdiff\b/, /error/], 11),
    status: findBaseDepthColumn_(headers, [/status/, /result/, /pass.*fail/], 12)
  };
}

function calculateBaseDepthResult_(scriptDistance, realDistanceCm) {
  var script = Number(String(scriptDistance || '').replace(',', '.'));
  var realCm = Number(String(realDistanceCm || '').replace(',', '.'));
  if (!isFinite(script) || !isFinite(realCm)) return null;
  var delta = realCm / 100 - script;
  var accuracy = script === 0 ? null : Math.max(0, (1 - Math.abs(delta) / Math.abs(script)) * 100);
  return { delta: delta, accuracy: accuracy, status: Math.abs(delta) <= BASE_DEPTH_TOLERANCE_M_ + 1e-9 ? 'Pass' : 'Fail' };
}

function appendBaseDepthRecord(payload) {
  try {
    payload = payload || {};
    var sheetName = String(payload.sheetName || '').trim();
    var robotId = String(payload.robotId || '').trim();
    var tester = String(payload.tester || '').trim();
    var swVersion = String(payload.swVersion || '').trim();
    var date = String(payload.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd')).trim();
    if (!sheetName || !robotId || !tester || !swVersion) return { success:false, error:'Robot ID, Tester, SW Version and sheet are required.' };

    var sections = payload.obstacles || {};
    var sizes = ['3cm', '2cm'];
    var normalizedRows = [];
    sizes.forEach(function(size) {
      var rows = Array.isArray(sections[size]) ? sections[size] : [];
      if (rows.length !== 6) throw new Error(size + ' must contain exactly 6 test points.');
      rows.forEach(function(row, index) {
        var entry = {
          size:size,
          point:index + 1,
          scriptDist:String(row.scriptDist || '').trim().replace(',', '.'),
          realDist:String(row.realDist || '').trim().replace(',', '.'),
          coordX:String(row.coordX || '').trim().replace(',', '.'),
          coordY:String(row.coordY || '').trim().replace(',', '.')
        };
        if (![entry.scriptDist, entry.realDist, entry.coordX, entry.coordY].every(function(value) { return value !== '' && isFinite(Number(value)); })) {
          throw new Error(size + ' point ' + (index + 1) + ' contains incomplete or invalid raw values.');
        }
        normalizedRows.push(entry);
      });
    });

    var sheet = SpreadsheetApp.openByUrl(DATABASE_URL).getSheetByName(sheetName);
    if (!sheet) return { success:false, error:'Base Depth sheet not found.' };
    var displayValues = sheet.getDataRange().getDisplayValues();
    var headerIndex = detectIsoHeaderRow_(displayValues);
    var width = sheet.getLastColumn();
    var headers = makeUniqueIsoHeaders_(displayValues[headerIndex] || [], width);
    var columns = getBaseDepthColumnMap_(headers);
    if (columns.scriptDist < 0 || columns.realDist < 0 || (columns.coord < 0 && (columns.coordX < 0 || columns.coordY < 0))) {
      return { success:false, error:'Base Depth raw input columns could not be matched in Google Sheets.' };
    }

    var lastRow = sheet.getLastRow();
    sheet.insertRowsAfter(lastRow, normalizedRows.length);
    var firstRow = lastRow + 1;
    var output = normalizedRows.map(function(entry) {
      var row = new Array(width).fill('');
      var calculated = calculateBaseDepthResult_(entry.scriptDist, entry.realDist);
      if (columns.obstacle >= 0) row[columns.obstacle] = entry.size;
      if (columns.point >= 0) row[columns.point] = entry.point;
      row[columns.scriptDist] = entry.scriptDist;
      row[columns.realDist] = entry.realDist;
      if (columns.coord >= 0) row[columns.coord] = entry.coordX + ', ' + entry.coordY;
      if (columns.coordX >= 0) row[columns.coordX] = entry.coordX;
      if (columns.coordY >= 0) row[columns.coordY] = entry.coordY;
      if (columns.accuracy >= 0 && calculated) row[columns.accuracy] = calculated.accuracy === null ? '' : calculated.accuracy.toFixed(2) + '%';
      if (columns.delta >= 0 && calculated) row[columns.delta] = calculated.delta.toFixed(3);
      if (columns.status >= 0 && calculated) row[columns.status] = calculated.status;
      return row;
    });
    sheet.getRange(firstRow, 1, output.length, width).setValues(output);

    var shared = [
      { column:columns.robot, value:robotId },
      { column:columns.date, value:date },
      { column:columns.sw, value:swVersion },
      { column:columns.tester, value:tester }
    ];
    shared.forEach(function(item) {
      if (item.column < 0) return;
      var range = sheet.getRange(firstRow, item.column + 1, output.length, 1);
      try {
        range.mergeVertically();
        sheet.getRange(firstRow, item.column + 1).setValue(item.value);
      } catch (mergeError) {
        var values = new Array(output.length).fill(null).map(function(_, index) { return [index === 0 ? item.value : '']; });
        range.setValues(values);
      }
    });
    sheet.getRange(firstRow, 1, output.length, width).setVerticalAlignment('middle');
    CacheService.getScriptCache().remove('iso_module_basedepth');
    return { success:true, firstRow:firstRow, rowsAdded:output.length, data:getBaseDepthData() };
  } catch (err) {
    return { success:false, error:err.toString() };
  }
}
