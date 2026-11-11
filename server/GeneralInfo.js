/** General Info metadata service. Images are always returned as direct URLs. */
function generalInfoDriveFileId_(value) {
  var url = String(value || '').trim();
  if (!url) return '';
  var patterns = [
    /\/d\/([a-zA-Z0-9_-]{20,})/,
    /[?&]id=([a-zA-Z0-9_-]{20,})/,
    /\/thumbnail\?id=([a-zA-Z0-9_-]{20,})/,
    /^([a-zA-Z0-9_-]{20,})$/
  ];
  for (var i = 0; i < patterns.length; i++) {
    var match = url.match(patterns[i]);
    if (match) return match[1];
  }
  return '';
}

function normalizeGeneralInfoImageUrl_(value) {
  var url = String(value || '').trim();
  var fileId = generalInfoDriveFileId_(url);
  if (fileId) return 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(fileId) + '&sz=w1200';
  return /^https?:\/\//i.test(url) ? url : '';
}

function extractGeneralInfoImageUrl_(formula, value, rich, note, raw) {
  try {
    var candidates = [];
    if (raw && typeof raw === 'object' && typeof raw.getContentUrl === 'function') {
      try { candidates.push(raw.getContentUrl()); } catch (imageError) {}
    }
    candidates = candidates.concat(String(formula || '').match(/https?:\/\/[^"')\s,;]+/gi) || []);
    if (/^https?:\/\//i.test(String(value || '').trim())) candidates.push(value);
    if (rich) {
      if (rich.getLinkUrl && rich.getLinkUrl()) candidates.push(rich.getLinkUrl());
      if (rich.getRuns) rich.getRuns().forEach(function(run) { if (run.getLinkUrl()) candidates.push(run.getLinkUrl()); });
    }
    candidates = candidates.concat(String(note || '').match(/https?:\/\/\S+/gi) || []);
    for (var i = 0; i < candidates.length; i++) {
      var normalized = normalizeGeneralInfoImageUrl_(candidates[i]);
      if (normalized) return normalized;
    }
  } catch (error) {
    console.error('General Info image URL error: ' + error);
  }
  return '';
}

function generalInfoRichLink_(rich) {
  try {
    if (rich && rich.getLinkUrl && rich.getLinkUrl()) return rich.getLinkUrl();
    if (rich && rich.getRuns) {
      var runs = rich.getRuns();
      for (var i = 0; i < runs.length; i++) if (runs[i].getLinkUrl()) return runs[i].getLinkUrl();
    }
  } catch (error) {}
  return '';
}

function getGeneralInfoDataV2_() {
  var cacheKey = 'general-info-cdn-v3';
  var cached = ISOApp.Server.Cache.readJson(cacheKey);
  if (cached) return cached;
  var result = fetchGeneralInfoDataV2_();
  if (result && result.success) ISOApp.Server.Cache.writeJson(cacheKey, result, 21600);
  return result;
}

function fetchGeneralInfoDataV2_() {
  try {
    var spreadsheet = ISOApp.Server.Database.open();
    var sheet = ISOApp.Server.Database.requireSheet(spreadsheet, ['General info', 'General Info'], 'General Info');
    // Read the whole reference block once. Individual getRange calls are
    // noticeably slower in Apps Script, especially on a cold page load.
    var referenceRange = sheet.getRange('B3:C7');
    var values = referenceRange.getDisplayValues();
    var formulas = referenceRange.getFormulas();
    var richValues = referenceRange.getRichTextValues();
    var notes = referenceRange.getNotes();
    var rawValues = referenceRange.getValues();
    function image(row) { return extractGeneralInfoImageUrl_(formulas[row][0], values[row][0], richValues[row][0], notes[row][0], rawValues[row][0]); }
    function imageRef(row) {
      var raw = rawValues[row] && rawValues[row][0];
      return raw && typeof raw === 'object' && typeof raw.getContentUrl === 'function'
        ? { sheetName:sheet.getName(), row:row + 3, column:2 }
        : null;
    }
    function description(row, fallback) { return String(values[row][1] || fallback).trim(); }
    function link(row) { return generalInfoRichLink_(richValues[row][1]); }
    return ISOApp.Server.Result.ok({
      items: [
        { id:'mindmap', title:'Mind Map of ISO Test', img:image(0) || normalizeGeneralInfoImageUrl_(link(0)), imageRef:imageRef(0), note:description(0, 'Test structure & dependencies'), link:link(0) },
        { id:'process', title:'EVT2 Testing Process / Structure', img:image(1) || normalizeGeneralInfoImageUrl_(link(1)), imageRef:imageRef(1), note:description(1, 'Workflow and factory validation'), link:link(1) },
        { id:'station', title:'Test Station Layout', img:image(2) || normalizeGeneralInfoImageUrl_(link(2)), imageRef:imageRef(2), note:description(2, 'Physical room arrangement'), link:link(2) },
        { id:'robot', title:'Robot LU Definition & Pose', img:image(3) || normalizeGeneralInfoImageUrl_(link(4)), imageRef:imageRef(3), note:String((values[4][0] || 'LU = ⌈w/500⌉ × 500 mm') + ' | ' + (values[4][1] || '1.17 m x 0.56 m x 0.56 m')).trim(), link:link(4) }
      ]
    });
  } catch (error) {
    return ISOApp.Server.Result.fail(error, 'General Info could not be loaded.');
  }
}

function getGeneralInfoImagesV2_() { return getGeneralInfoDataV2_(); }

ISOApp.Services = ISOApp.Services || {};
ISOApp.Services.GeneralInfo = Object.freeze({ list:getGeneralInfoDataV2_, images:getGeneralInfoImagesV2_ });

function getGeneralInfoDataV2() { return ISOApp.Services.GeneralInfo.list(); }
function getGeneralInfoImagesV2() { return ISOApp.Services.GeneralInfo.images(); }
