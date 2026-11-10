/** Head Depth Camera module entry point. Shared parsing remains in Modules.js. */
function urlToBase64(url) {
  if (!url || typeof url !== 'string' || url.indexOf('http') !== 0) return null;
  try {
    var response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true
    });
    if (response.getResponseCode() === 200) {
      var blob = response.getBlob();
      return 'data:' + (blob.getContentType() || 'image/png') + ';base64,' + Utilities.base64Encode(blob.getBytes());
    }
  } catch (e) {
    console.error('Base64 convert error:', e);
  }
  return null;
}

function getHeadDepthData() {
  var result = getModuleData('headdepth');
  if (!result || !Array.isArray(result.rows)) return result;

  // Sheets internal image URLs are not directly loadable from the sandboxed UI.
  // Fetch the two image columns in one batch without Auth headers and convert to base64.
  var urlMap = {};
  var requests = [];
  result.rows.forEach(function(row) {
    [8, 9].forEach(function(index) {
      var value = String(row[index] || '').trim();
      if (value.indexOf('http') === 0 && !Object.prototype.hasOwnProperty.call(urlMap, value)) {
        urlMap[value] = null;
        requests.push({
          url: value,
          muteHttpExceptions: true
        });
      }
    });
  });

  if (requests.length) {
    var responses = UrlFetchApp.fetchAll(requests);
    requests.forEach(function(request, index) {
      var response = responses[index];
      if (response && response.getResponseCode() === 200) {
        var blob = response.getBlob();
        urlMap[request.url] = 'data:' + (blob.getContentType() || 'image/png') + ';base64,' + Utilities.base64Encode(blob.getBytes());
      }
    });
  }

  result.rows.forEach(function(row) {
    [8, 9].forEach(function(index) {
      var value = String(row[index] || '').trim();
      if (urlMap[value]) row[index] = urlMap[value];
    });
  });

  return result;
}

/** Append one Head Depth Camera record in the sheet's fixed 14-column order. */
function appendHeadDepthRecord(payload) {
  try {
    payload = payload || {};
    var sheetName = String(payload.sheetName || '').trim();
    var values = Array.isArray(payload.values) ? payload.values.slice(0, 14) : [];
    if (!sheetName || values.length !== 14) return { success:false, error:'Head Depth requires exactly 14 columns.' };

    var sheet = SpreadsheetApp.openByUrl(DATABASE_URL).getSheetByName(sheetName);
    if (!sheet) return { success:false, error:'Head Depth sheet not found.' };
    var width = Math.max(14, sheet.getLastColumn());
    var row = new Array(width).fill('');
    values.forEach(function(value, index) { row[index] = value === undefined ? '' : String(value); });
    var firstRow = sheet.getLastRow() + 1;
    sheet.getRange(firstRow, 1, 1, width).setValues([row]);
    try { CacheService.getScriptCache().remove('iso_module_headdepth'); } catch (cacheError) { /* optional */ }

    return { success:true, imported:{ firstRow:firstRow }, data:getHeadDepthData() };
  } catch (err) {
    return { success:false, error:err.toString() };
  }
}