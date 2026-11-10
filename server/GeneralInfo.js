/**
 * General Info data service.
 * Keeps sheet-specific image extraction separate from generic module services.
 */
function extractImageAsBase64(sheet, cellA1) {
  try {
    var range = sheet.getRange(cellA1);
    var value = range.getValue();
    var formula = range.getFormula();
    var targetUrl = null;

    if (formula && formula.toUpperCase().indexOf('=IMAGE') !== -1) {
      var match = formula.match(/"([^"]+)"/);
      if (match && match[1]) targetUrl = match[1];
    } else if (typeof value === 'string' && value.indexOf('http') === 0) {
      targetUrl = value;
    } else if (value && typeof value === 'object' && typeof value.getContentUrl === 'function') {
      targetUrl = value.getContentUrl();
    }

    if (targetUrl) {
      var attempts = [{ muteHttpExceptions: true, followRedirects: true }];
      try { attempts.unshift({ muteHttpExceptions: true, followRedirects: true, headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }); } catch (tokenError) {}
      for (var i = 0; i < attempts.length; i++) {
        var response = UrlFetchApp.fetch(targetUrl, attempts[i]);
        if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) continue;
        var blob = response.getBlob();
        var type = String(blob.getContentType() || '').toLowerCase();
        if (type.indexOf('image/') === 0) return 'data:' + type + ';base64,' + Utilities.base64Encode(blob.getBytes());
      }
    }

    if (typeof sheet.getImages === 'function') {
      var row = range.getRow(), col = range.getColumn();
      var images = sheet.getImages();
      for (var j = 0; j < images.length; j++) {
        var anchor = images[j].getAnchorCell();
        if ((anchor.getRow() === row || (anchor.getRow() >= row - 1 && anchor.getRow() <= row + 1)) && anchor.getColumn() === col) {
          var imageBlob = images[j].getBlob();
          return 'data:' + (imageBlob.getContentType() || 'image/png') + ';base64:' + Utilities.base64Encode(imageBlob.getBytes());
        }
      }
    }
  } catch (error) {
    console.error('Error extracting General Info image from ' + cellA1 + ':', error);
  }
  return null;
}

function getGeneralInfoDataV2() {
  return getCachedModuleData('general-info-v4', function() {
    return fetchGeneralInfoDataV2_();
  }, 21600);
}

function getCachedModuleData(cacheKey, fetchFunction, expirationSec) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  var freshData = fetchFunction();
  try {
    var payload = JSON.stringify(freshData);
    if (payload.length < 100000) cache.put(cacheKey, payload, expirationSec || 21600);
  } catch (e) {}
  return freshData;
}

function fetchGeneralInfoDataV2_() {
  try {
    var ss = typeof DATABASE_URL !== 'undefined' && DATABASE_URL
      ? SpreadsheetApp.openByUrl(DATABASE_URL)
      : SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('General info');
    if (!sheet) return { success: false, error: 'Sheet General info not found' };
    var floatingMap = {};
    try {
      sheet.getImages().forEach(function(image) {
        var row = image.getAnchorCell().getRow();
        var blob = image.getBlob();
        floatingMap[row] = 'data:' + (blob.getContentType() || 'image/png') + ';base64:' + Utilities.base64Encode(blob.getBytes());
      });
    } catch (imageError) {}
    function cellImage(cell) { return extractImageAsBase64(sheet, cell); }
    function note(cell, fallback) { return String(sheet.getRange(cell).getValue() || fallback).trim(); }
    function link(cell) {
      try {
        var rich = sheet.getRange(cell).getRichTextValue();
        if (rich && rich.getLinkUrl()) return rich.getLinkUrl();
        if (rich && rich.getRuns) {
          var runs = rich.getRuns();
          for (var i = 0; i < runs.length; i++) if (runs[i].getLinkUrl()) return runs[i].getLinkUrl();
        }
      } catch (e) {}
      return null;
    }
  return {
    success: true,
    items: [
      { id: 'mindmap', title: 'Mind Map of ISO Test', img: cellImage('B3') || floatingMap[3] || floatingMap[2] || null, note: note('C3', 'Test structure & dependencies'), link: link('C3') },
      { id: 'process', title: 'EVT2 Testing Process / Structure', img: cellImage('B4') || floatingMap[4] || null, note: note('C4', 'Workflow and factory validation'), link: link('C4') },
      { id: 'station', title: 'Test Station Layout', img: cellImage('B5') || floatingMap[5] || null, note: note('C5', 'Physical room arrangement'), link: link('C5') },
      { id: 'robot', title: 'Robot LU Definition & Pose', img: cellImage('B6') || floatingMap[6] || null, note: String((sheet.getRange('B7').getValue() || 'LU = ⌈w/500⌉ × 500 mm') + ' | ' + (sheet.getRange('C7').getValue() || '1.17 m x 0.56 m x 0.56 m')).trim(), link: link('C7') }
    ]
  };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getGeneralInfoImagesV2() {
  return getGeneralInfoDataV2();
}
