/** Head Depth Camera module entry point. Shared parsing remains in Modules.js. */
function urlToBase64_(url) {
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

function getHeadDepthData_() {
  // Return metadata and image URLs immediately. The browser loads thumbnails
  // lazily after the table has painted; no UrlFetchApp or Base64 work blocks
  // the main Head Depth response.
  return getModuleData('headdepth');
}

function saveHeadDepthImage_(payload) {
  try {
    payload = payload || {};
    var dataUrl = String(payload.dataUrl || '');
    var match = dataUrl.match(/^data:(image\/[\w.+-]+);base64,(.+)$/);
    if (!match) return { success:false, error:'Invalid image upload.' };
    if (match[2].length > 12000000) return { success:false, error:'Image is too large. Please choose a smaller image.' };
    var bytes = Utilities.base64Decode(match[2]);
    var blob = Utilities.newBlob(bytes, match[1], String(payload.filename || 'head-depth-camera-image'));
    var file = DriveApp.createFile(blob);
    return { success:true, url:file.getUrl(), name:file.getName() };
  } catch (error) {
    return { success:false, error:String(error) };
  }
}

/** Append one Head Depth Camera record in the sheet's fixed 14-column order. */
function appendHeadDepthRecord_(payload) {
  try {
    payload = payload || {};
    var sheetName = String(payload.sheetName || '').trim();
    var values = Array.isArray(payload.values) ? payload.values.slice(0, 14) : [];
    if (!sheetName || values.length !== 14) return { success:false, error:'Head Depth requires exactly 14 columns.' };

    var sheet = ISOApp.Server.Database.open().getSheetByName(sheetName);
    if (!sheet) return { success:false, error:'Head Depth sheet not found.' };
    var width = Math.max(14, sheet.getLastColumn());
    var row = ISOApp.Server.Rows.strings(values, width);
    var firstRow = sheet.getLastRow() + 1;
    sheet.getRange(firstRow, 1, 1, width).setValues([row]);
    ISOApp.Server.Cache.remove('iso_module_headdepth');
    ISOApp.Server.Cache.remove('iso_static_input_choices');

    return { success:true, imported:{ firstRow:firstRow }, data:getHeadDepthData_() };
  } catch (err) {
    return { success:false, error:err.toString() };
  }
}

ISOApp.Services = ISOApp.Services || {};
ISOApp.Services.HeadDepth = Object.freeze({
  list:getHeadDepthData_,
  saveImage:saveHeadDepthImage_,
  append:appendHeadDepthRecord_
});

// Public RPC adapters.
function getHeadDepthData() { return ISOApp.Services.HeadDepth.list(); }
function saveHeadDepthImage(payload) { return ISOApp.Services.HeadDepth.saveImage(payload); }
function appendHeadDepthRecord(payload) { return ISOApp.Services.HeadDepth.append(payload); }
