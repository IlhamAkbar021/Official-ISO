function getWifiData() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('iso_wifi_data');
  if (cached) {
    try { return JSON.parse(cached); } catch (cacheError) { /* rebuild */ }
  }
  var ss = SpreadsheetApp.openByUrl(DATABASE_URL);
  var sheet = getSheetFlexible(ss, ['Wi-Fi Distance/Antenna Diversity', 'WiFi', 'Wi-Fi']);
  if (!sheet) return { error: "Sheet 'Wi-Fi Distance/Antenna Diversity' not found in database." };
  
  var apVal = sheet.getRange('D2').getDisplayValue().trim();
  var distVal = sheet.getRange('D3').getDisplayValue().trim();
  var sigVal = sheet.getRange('D4').getDisplayValue().trim();

  if (distVal === '' || distVal === '30.00' || distVal === '30') { distVal = '10'; }

  var setupData = {
    apName: apVal !== '' ? apVal : 'meshtest',
    distance: distVal,
    signal: sigVal !== '' ? sigVal : '-55'
  };

  var range = sheet.getDataRange();
  var displayValues = range.getDisplayValues().slice(5); 
  var richTextValues = range.getRichTextValues().slice(5);
  var formulas = range.getFormulas().slice(5);

  if (displayValues.length < 1) return { setup: setupData, headers: [], rows: [] };

  var headers = displayValues[0];
  var dataRows = displayValues.slice(1);
  var richTextRows = richTextValues.slice(1);
  var formulaRows = formulas.slice(1);

  var processedRows = dataRows.map(function(row, rIdx) {
    return row.map(function(cellVal, cIdx) {
      var val = cellVal ? cellVal.toString().trim() : '';
      var richText = richTextRows[rIdx] ? richTextRows[rIdx][cIdx] : null;
      var url = richText ? richText.getLinkUrl() : null;

      if (!url && formulaRows[rIdx] && formulaRows[rIdx][cIdx]) {
        var match = formulaRows[rIdx][cIdx].match(/HYPERLINK\(\s*["']([^"']+)["']/i);
        if (match && match[1]) { url = match[1]; }
      }
      return url ? url : val;
    });
  });

  var result = { setup: setupData, headers: headers, rows: processedRows };
  try { cache.put('iso_wifi_data', JSON.stringify(result), 60); } catch (cacheWriteError) { /* cache is optional */ }
  return result;
}

function addWifiRecord(payload) {
  try {
    var ss = SpreadsheetApp.openByUrl(DATABASE_URL);
    var sheet = getSheetFlexible(ss, ['Wi-Fi Distance/Antenna Diversity', 'WiFi', 'Wi-Fi']);
    if (!sheet) return { success: false, error: "Wi-Fi Sheet not found." };
    
    var lastRow = sheet.getLastRow();
    var numCols = sheet.getLastColumn();
    
    sheet.getRange(lastRow, 1, 1, numCols).copyTo(sheet.getRange(lastRow + 1, 1));
    
    var lastId = sheet.getRange(lastRow, 1).getValue();
    var newId = (typeof lastId === 'number') ? lastId + 1 : "";
    sheet.getRange(lastRow + 1, 1).setValue(newId);
    
    sheet.getRange(lastRow + 1, 2).setValue(payload.robotId);        
    sheet.getRange(lastRow + 1, 3).setValue(payload.tester);         
    sheet.getRange(lastRow + 1, 4).setValue(payload.rightAntenna);   
    sheet.getRange(lastRow + 1, 8).setValue(payload.rightLink || "");
    sheet.getRange(lastRow + 1, 10).setValue(payload.leftAntenna);  
    sheet.getRange(lastRow + 1, 14).setValue(payload.leftLink || "");
    
    CacheService.getScriptCache().remove('iso_wifi_data');
    return { success: true };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function updateWifiRecord(payload) {
  try {
    var ss = SpreadsheetApp.openByUrl(DATABASE_URL);
    var sheet = getSheetFlexible(ss, ['Wi-Fi Distance/Antenna Diversity', 'WiFi', 'Wi-Fi']);
    if (!sheet) return { success: false, error: "Wi-Fi Sheet not found." };
    var targetRow = payload.rowIndex + 7;
    sheet.getRange(targetRow, 3).setValue(payload.tester);          
    sheet.getRange(targetRow, 4).setValue(payload.rightAntenna);    
    sheet.getRange(targetRow, 8).setValue(payload.rightLink || ""); 
    sheet.getRange(targetRow, 10).setValue(payload.leftAntenna);    
    sheet.getRange(targetRow, 14).setValue(payload.leftLink || ""); 
    CacheService.getScriptCache().remove('iso_wifi_data');
    return { success: true };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function deleteWifiRecord(payload) {
  try {
    var ss = SpreadsheetApp.openByUrl(DATABASE_URL);
    var sheet = getSheetFlexible(ss, ['Wi-Fi Distance/Antenna Diversity', 'WiFi', 'Wi-Fi']);
    if (!sheet) return { success: false, error: "Wi-Fi Sheet not found." };
    var targetRow = payload.rowIndex + 7;
    sheet.deleteRow(targetRow);
    CacheService.getScriptCache().remove('iso_wifi_data');
    return { success: true };
  } catch (err) { return { success: false, error: err.toString() }; }
}
