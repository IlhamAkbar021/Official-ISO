/**
 * Generic ISO module reader.
 *
 * The workbook is maintained by several teams, so sheet names and header row
 * positions are not always identical. This reader deliberately discovers the
 * best sheet and header row instead of hard-coding column coordinates.
 *
 * Sections: discovery -> cell/image helpers -> column inference -> row
 * summarization -> public CRUD endpoints.
 */

// -----------------------------------------------------------------------------
// Sheet and header discovery
// -----------------------------------------------------------------------------
function normalizeIsoName_(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

function findIsoModuleSheet_(ss, config) {
  var sheets = ss.getSheets();
  var best = null;
  var bestScore = 0;

  sheets.forEach(function(sheet) {
    var sheetName = normalizeIsoName_(sheet.getName());
    config.aliases.forEach(function(alias) {
      var candidate = normalizeIsoName_(alias);
      var score = 0;
      if (sheetName === candidate) score = 100 + candidate.length;
      else if (sheetName.indexOf(candidate) !== -1) score = 60 + candidate.length;
      else if (candidate.indexOf(sheetName) !== -1 && sheetName.length > 4) score = 30 + sheetName.length;
      if (score > bestScore) {
        bestScore = score;
        best = sheet;
      }
    });
  });

  // Some workbooks keep both detection methods in one A&R sheet. Use that
  // combined sheet only when a detector-specific tab does not exist.
  if (!best && config.fallbackAliases) {
    return findIsoModuleSheet_(ss, { aliases: config.fallbackAliases });
  }

  return best;
}

function detectIsoHeaderRow_(values) {
  var keywords = /robot|unit|serial|date|tester|owner|version|result|status|pass|fail|speed|distance|depth|height|angle|trial|run|note|remark|link|value|target|actual|sample/i;
  var maxRows = Math.min(values.length, 18);
  var bestIndex = 0;
  var bestScore = -1;

  for (var r = 0; r < maxRows; r++) {
    var nonEmpty = 0;
    var keywordHits = 0;
    var unique = {};
    for (var c = 0; c < values[r].length; c++) {
      var text = String(values[r][c] || '').trim();
      if (!text) continue;
      nonEmpty++;
      unique[text.toLowerCase()] = true;
      if (keywords.test(text)) keywordHits++;
    }
    var score = nonEmpty + keywordHits * 3 + Object.keys(unique).length * 0.2;
    if (nonEmpty >= 2 && score > bestScore) {
      bestScore = score;
      bestIndex = r;
    }
  }
  return bestIndex;
}

function makeUniqueIsoHeaders_(row, width) {
  var used = {};
  var headers = [];
  for (var c = 0; c < width; c++) {
    var base = String(row[c] || '').replace(/[\r\n]+/g, ' ').trim() || ('Column ' + (c + 1));
    var name = base;
    var suffix = 2;
    while (used[name.toLowerCase()]) name = base + ' ' + suffix++;
    used[name.toLowerCase()] = true;
    headers.push(name);
  }
  return headers;
}

function getIsoCell_(displayValues, richValues, formulas, rawValues, r, c) {
  var display = displayValues[r] && displayValues[r][c] !== undefined ? String(displayValues[r][c]).trim() : '';
  var rich = richValues[r] ? richValues[r][c] : null;
  var url = rich && rich.getLinkUrl ? rich.getLinkUrl() : null;
  var formula = formulas[r] ? formulas[r][c] : '';
  var raw = rawValues && rawValues[r] ? rawValues[r][c] : null;

  // Images inserted directly into a Google Sheets cell are returned as a
  // CellImage object rather than display text.
  if (raw && typeof raw === 'object' && typeof raw.getContentUrl === 'function') {
    try {
      // Prefer the actual CellImage content URL over its rich-text link. The
      // rich link commonly points to a Drive viewer HTML page, which cannot be
      // rendered by an <img> element.
      url = raw.getContentUrl();
    } catch (imageError) { /* fall through to the original value */ }
  }
  if (!url && formula) {
    var match = String(formula).match(/HYPERLINK\(\s*["']([^"']+)["']/i);
    if (match) url = match[1];
    if (!url) {
      var imageMatch = String(formula).match(/IMAGE\(\s*["']([^"']+)["']/i);
      if (imageMatch) url = imageMatch[1];
    }
  }
  return url || display;
}

/** Fetch one protected Sheets image only when its thumbnail is needed. */
function getIsoImageData(payload) {
  try {
    payload = payload || {};
    var sheetName = String(payload.sheetName || '').trim();
    var row = Number(payload.row);
    var column = Number(payload.column);
    if (!sheetName || !row || !column) return { success: false, error: 'Invalid image cell.' };

    var sheet = SpreadsheetApp.openByUrl(DATABASE_URL).getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Image sheet not found.' };
    var cell = sheet.getRange(row, column);
    var raw = cell.getValue();
    var formula = cell.getFormula();
    var rich = cell.getRichTextValue();
    var url = rich && rich.getLinkUrl ? rich.getLinkUrl() : '';
    if (raw && typeof raw === 'object' && typeof raw.getContentUrl === 'function') url = raw.getContentUrl();
    if (!url && typeof raw === 'string' && /^https?:\/\//i.test(raw.trim())) url = raw.trim();
    if (!url && formula) {
      var imageMatch = String(formula).match(/IMAGE\(\s*["']([^"']+)["']/i);
      if (imageMatch) url = imageMatch[1];
    }
    if (!url) return { success: false, error: 'No image found in this cell.' };

    // Drive share links normally return an HTML viewer page through UrlFetch.
    // Read the file blob directly so the browser receives actual image bytes.
    var driveMatch = String(url).match(/(?:\/d\/|[?&]id=)([-\w]{20,})/i);
    if (driveMatch && typeof DriveApp !== 'undefined') {
      try {
        var driveFile = DriveApp.getFileById(driveMatch[1]);
        // Drive's thumbnail is much smaller and faster than downloading the
        // original camera image. Fall back to the original only when no
        // thumbnail is available.
        var driveBlob = driveFile.getThumbnail() || driveFile.getBlob();
        var driveType = String(driveBlob.getContentType() || '').toLowerCase();
        if (driveType.indexOf('image/') === 0) {
          return { success: true, dataUrl: 'data:' + driveType + ';base64,' + Utilities.base64Encode(driveBlob.getBytes()) };
        }
      } catch (driveError) { /* fall through to URL fetch variants */ }
    }

    var attempts = [];
    // CellImage content URLs are commonly protected. Try the Apps Script
    // authorization token first so the image is read from the Sheet directly.
    try { attempts.push({ muteHttpExceptions: true, followRedirects: true, headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }); } catch (tokenError) { /* public image only */ }
    attempts.push({ muteHttpExceptions: true, followRedirects: true });
    for (var attemptIndex = 0; attemptIndex < attempts.length; attemptIndex++) {
      var response = UrlFetchApp.fetch(url, attempts[attemptIndex]);
      if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) continue;
      var blob = response.getBlob();
      var contentType = String(blob.getContentType() || '').toLowerCase();
      if (contentType.indexOf('image/') !== 0) continue;
      return { success: true, dataUrl: 'data:' + contentType + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
    }
    return { success: false, error: 'Image access denied or returned a non-image response.' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/** Read image cells from the General/Specification information sheet. */
function getGeneralInfoImages() {
  try {
    var ss = SpreadsheetApp.openByUrl(DATABASE_URL);
    var preferred = ['General Info', 'General Information', 'Specification', 'Factory Info', 'Environment'];
    var sheet = null;
    ss.getSheets().some(function(candidate) {
      var name = normalizeIsoName_(candidate.getName());
      if (preferred.some(function(alias) { return name === normalizeIsoName_(alias) || name.indexOf(normalizeIsoName_(alias)) !== -1; })) {
        sheet = candidate;
        return true;
      }
      return false;
    });
    if (!sheet) return { success: true, images: [] };

    var range = sheet.getDataRange();
    var display = range.getDisplayValues();
    var rich = range.getRichTextValues();
    var formulas = range.getFormulas();
    var raw = range.getValues();
    var images = [];

    // Floating images inserted with Sheets' Insert Image feature do not appear
    // in getDataRange(). Read them from the sheet image layer as well.
    if (typeof sheet.getImages === 'function') {
      sheet.getImages().forEach(function(floatingImage, imageIndex) {
        try {
          var anchor = floatingImage.getAnchorCell();
          var sourceUrl = typeof floatingImage.getUrl === 'function' ? floatingImage.getUrl() : '';
          if (!sourceUrl) return;
          images.push({
            value: sourceUrl,
            header: 'Reference image ' + (imageIndex + 1),
            row: anchor.getRow(),
            column: anchor.getColumn(),
            sheetName: sheet.getName()
          });
        } catch (floatingImageError) { /* continue with cell images */ }
      });
    }

    for (var r = 0; r < display.length; r++) {
      var rowLabel = '';
      for (var labelCol = 0; labelCol < Math.min(3, display[r].length); labelCol++) {
        if (String(display[r][labelCol] || '').trim()) { rowLabel = String(display[r][labelCol]).trim(); break; }
      }
      for (var c = 0; c < display[r].length; c++) {
        var value = getIsoCell_(display, rich, formulas, raw, r, c);
        if (!/^https?:\/\//i.test(String(value || '')) && !/^data:image\//i.test(String(value || ''))) continue;
        var imageResult = getIsoImageData({ sheetName: sheet.getName(), row: r + 1, column: c + 1 });
        images.push({
          value: imageResult && imageResult.success && imageResult.dataUrl ? imageResult.dataUrl : value,
          header: rowLabel || ('Information image ' + (images.length + 1)),
          row: r + 1,
          column: c + 1,
          sheetName: sheet.getName()
        });
      }
    }
    return { success: true, sheetName: sheet.getName(), images: images };
  } catch (err) {
    return { success: false, error: err.toString(), images: [] };
  }
}

// -----------------------------------------------------------------------------
// Column inference and row normalization
// -----------------------------------------------------------------------------

function classifyIsoColumns_(headers) {
  var roles = { robot: -1, tester: -1, date: -1, status: [] };
  headers.forEach(function(header, index) {
    var h = normalizeIsoName_(header);
    if (roles.robot === -1 && /robot|unit id|serial|machine id/.test(h)) roles.robot = index;
    if (roles.tester === -1 && /tester|owner|operator|engineer/.test(h)) roles.tester = index;
    if (roles.date === -1 && /date|time/.test(h)) roles.date = index;
    if (/pass|fail|status|result|judgement|judgment|conclusion/.test(h)) roles.status.push(index);
  });
  return roles;
}

function inferIsoRolesFromRows_(headers, rows, roles) {
  if (!rows.length) return roles;
  var width = headers.length;
  var robotScores = [];
  var dateScores = [];
  var statusScores = [];

  for (var c = 0; c < width; c++) {
    var robotScore = 0, dateScore = 0, statusScore = 0, populated = 0;
    rows.forEach(function(row) {
      var value = String(row[c] || '').trim();
      if (!value) return;
      populated++;
      if (/\b(?:AC\d{3,}|DVT\d|EVT\d|robot)\b/i.test(value) || /^re[-\s]?run\s+.*(?:AC|DVT)/i.test(value)) robotScore += 4;
      if (/^\d{4}[\/-]\d{1,2}[\/-]\d{1,2}$/.test(value)) dateScore += 2;
      if (/^(?:pass|fail|done|pending|normal|acceptable|ok|ng)$/i.test(value)) statusScore += 2;
    });
    robotScores[c] = robotScore;
    dateScores[c] = dateScore;
    statusScores[c] = populated ? statusScore / populated : 0;
  }

  function bestIndex(scores) {
    var index = -1, score = 0;
    scores.forEach(function(value, i) { if (value > score) { score = value; index = i; } });
    return { index: index, score: score };
  }

  var robotBest = bestIndex(robotScores);
  var dateBest = bestIndex(dateScores);
  if (roles.robot < 0 && robotBest.score >= 4) roles.robot = robotBest.index;
  if (roles.date < 0 && dateBest.score >= 2) roles.date = dateBest.index;
  if (!roles.status.length) {
    statusScores.forEach(function(score, index) { if (score >= 0.35) roles.status.push(index); });
  }
  return roles;
}

function findIsoSectionBounds_(values, config) {
  if (!config.sectionKeywords || !config.sectionKeywords.length) return { start: 0, end: values.length };
  var start = -1;
  var end = values.length;
  for (var r = 0; r < values.length; r++) {
    var rowText = normalizeIsoName_(values[r].join(' '));
    var isStart = config.sectionKeywords.some(function(keyword) {
      return rowText.indexOf(normalizeIsoName_(keyword)) !== -1;
    });
    if (start === -1 && isStart) {
      start = r;
      continue;
    }
    if (start !== -1 && r > start && config.stopKeywords && config.stopKeywords.some(function(keyword) {
      return rowText.indexOf(normalizeIsoName_(keyword)) !== -1;
    })) {
      end = r;
      break;
    }
  }
  return start === -1 ? { start: 0, end: values.length } : { start: start, end: end };
}

function summarizeIsoRows_(headers, rows, roles) {
  var robots = {};
  var pass = 0;
  var fail = 0;
  var warning = 0;
  var statusCount = 0;

  rows.forEach(function(row) {
    if (roles.robot >= 0 && row[roles.robot]) robots[String(row[roles.robot]).trim()] = true;
    roles.status.forEach(function(col) {
      var value = String(row[col] || '').toLowerCase().trim();
      if (!value) return;
      statusCount++;
      if (/\bfail\b|\bng\b|reject|bad/.test(value)) fail++;
      else if (/\bpass\b|\bok\b|acceptable|normal|done/.test(value)) pass++;
      else if (/warn|pending|wait|progress|partial|retest/.test(value)) warning++;
    });
  });

  var numeric = [];
  headers.forEach(function(header, col) {
    if (/id|date|time|version|link|note|remark|status|result/i.test(header)) return;
    var values = [];
    rows.forEach(function(row) {
      var raw = String(row[col] || '').replace(/,/g, '').match(/-?\d+(?:\.\d+)?/);
      if (raw) values.push(parseFloat(raw[0]));
    });
    if (values.length >= Math.max(2, Math.ceil(rows.length * 0.35))) {
      var sum = values.reduce(function(total, value) { return total + value; }, 0);
      numeric.push({
        label: header,
        average: Math.round((sum / values.length) * 100) / 100,
        minimum: Math.min.apply(null, values),
        maximum: Math.max.apply(null, values),
        samples: values.length
      });
    }
  });

  return {
    records: rows.length,
    robots: Object.keys(robots).length,
    pass: pass,
    fail: fail,
    warning: warning,
    statusCount: statusCount,
    numeric: numeric.slice(0, 3)
  };
}

// -----------------------------------------------------------------------------
// Public module endpoints
// -----------------------------------------------------------------------------

function getModuleData(moduleKey) {
  try {
    var config = getIsoModuleConfig_(moduleKey);
    if (!config) return { success: false, error: 'Unknown ISO module.' };

    // Module tabs are read-only between refreshes most of the time. A short
    // cache prevents repeated navigation/search actions from re-reading a
    // large Google Sheet on every click while keeping the UI effectively live.
    var cache = CacheService.getScriptCache();
    var cacheKey = 'iso_module_' + String(moduleKey).replace(/[^a-z0-9_]/gi, '_');
    var cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (cacheError) { /* rebuild below */ }
    }

    var ss = SpreadsheetApp.openByUrl(DATABASE_URL);
    var sheet = findIsoModuleSheet_(ss, config);
    if (!sheet) {
      return {
        success: false,
        missingSheet: true,
        title: config.title,
        aliases: config.aliases,
        availableSheets: ss.getSheets().map(function(item) { return item.getName(); }),
        error: 'No matching Google Sheet tab was found for ' + config.title + '.'
      };
    }

    var range = sheet.getDataRange();
    var displayValues = range.getDisplayValues();
    // Rich text, formulas, and raw CellImage values are only needed by Head
    // Depth's image/link column. Avoiding these calls cuts server time for the
    // other ISO tabs substantially.
    var needsRichCells = moduleKey === 'headdepth';
    // Rich-text extraction is very expensive on image-heavy sheets. CellImage
    // objects are available from getValues(), while formulas are handled below,
    // so do not read the full rich-text matrix for Head Depth.
    var richValues = [];
    var formulas = needsRichCells ? range.getFormulas() : [];
    var rawValues = [];
    if (needsRichCells) {
      // Image cells can occasionally make getValues() fail on a workbook
      // while the normal display matrix is still readable. Keep the table
      // usable and let image preview fall back to the source link.
      try { rawValues = range.getValues(); } catch (imageMatrixError) { rawValues = []; }
    }
    if (!displayValues.length) return { success: true, title: config.title, sheetName: sheet.getName(), headers: [], rows: [], summary: { records: 0, robots: 0, pass: 0, fail: 0, warning: 0, statusCount: 0, numeric: [] } };

    var sectionBounds = findIsoSectionBounds_(displayValues, config);
    var sectionOffset = sectionBounds.start;
    displayValues = displayValues.slice(sectionBounds.start, sectionBounds.end);
    richValues = richValues.slice(sectionBounds.start, sectionBounds.end);
    formulas = formulas.slice(sectionBounds.start, sectionBounds.end);
    rawValues = rawValues.slice(sectionBounds.start, sectionBounds.end);

    var headerIndex = detectIsoHeaderRow_(displayValues);
    var lastColumn = 0;
    for (var r = headerIndex; r < displayValues.length; r++) {
      for (var c = displayValues[r].length - 1; c >= 0; c--) {
        if (String(displayValues[r][c] || '').trim() !== '') {
          lastColumn = Math.max(lastColumn, c + 1);
          break;
        }
      }
    }
    if (!lastColumn) lastColumn = displayValues[headerIndex].length;

    var headers = makeUniqueIsoHeaders_(displayValues[headerIndex], lastColumn);
    var distanceLabels = [];
    if (moduleKey === 'headdepth') {
      var distancePattern = /^\d+(?:\.\d+)?\s*(?:m|cm)$/i;
      var distanceValues = [];
      displayValues.slice(0, headerIndex).forEach(function(preHeaderRow) {
        preHeaderRow.forEach(function(value) {
          var label = String(value || '').replace(/\s+/g, ' ').trim();
          if (distancePattern.test(label) && distanceValues.indexOf(label) === -1) distanceValues.push(label);
        });
      });
      for (var distanceIndex = 7; distanceIndex < 10; distanceIndex++) {
        var headerDistance = String(displayValues[headerIndex][distanceIndex] || '').trim();
        distanceLabels.push(distancePattern.test(headerDistance) ? headerDistance : (distanceValues[distanceIndex - 7] || ''));
      }
    }
    // Head Depth sheets commonly place a short test-purpose explanation above
    // the column headers. Preserve it as context instead of dropping it when
    // the data table is normalized.
    var purpose = '';
    if (moduleKey === 'headdepth') {
      for (var purposeRow = 0; purposeRow < headerIndex; purposeRow++) {
        var purposeCells = (displayValues[purposeRow] || []).map(function(value) {
          return String(value || '').replace(/[\r\n]+/g, ' ').trim();
        }).filter(function(value) { return value; });
        if (purposeCells.some(function(value) { return /purpose/i.test(value); })) {
          purpose = purposeCells.filter(function(value) { return !/^purpose$/i.test(value); }).join(' ');
          break;
        }
      }
    }
    // The Head Depth workbook reserves its final column for the source/evidence
    // URL.  Keep that contract even when the sheet header is blank or generic.
    if (moduleKey === 'headdepth' && headers.length) {
      headers[headers.length - 1] = 'Link';
    }
    var rows = [];
    for (var rowIndex = headerIndex + 1; rowIndex < displayValues.length; rowIndex++) {
      var row = [];
      var hasData = false;
      for (var col = 0; col < lastColumn; col++) {
        var cell = getIsoCell_(displayValues, richValues, formulas, rawValues, rowIndex, col);
        if (cell !== '') hasData = true;
        row.push(cell);
      }
      if (hasData) rows.push(row);
    }

    var roles = inferIsoRolesFromRows_(headers, rows, classifyIsoColumns_(headers));
    var result = {
      success: true,
      key: moduleKey,
      title: config.title,
      phase: config.phase,
      sheetName: sheet.getName(),
      purpose: purpose,
      distanceLabels: distanceLabels,
      headerRow: sectionOffset + headerIndex + 1,
      headers: headers,
      rows: rows,
      roles: roles,
      summary: summarizeIsoRows_(headers, rows, roles),
      updatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss')
    };
    // Applies to every generic module: Base Depth, IR Cliff, Head Depth,
    // Dynamic ISO modules, and Accuracy & Repeatability modules.
    try { cache.put(cacheKey, JSON.stringify(result), 60); } catch (cacheWriteError) { /* cache is optional */ }
    return result;
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/** Update one cell from a module table. Used by the compact Base Depth editor. */
function updateIsoModuleCell(payload) {
  try {
    payload = payload || {};
    var sheetName = String(payload.sheetName || '').trim();
    var row = Number(payload.row);
    var column = Number(payload.column);
    if (!sheetName || !row || !column || row < 1 || column < 1) return { success: false, error: 'Invalid cell address.' };
    var ss = SpreadsheetApp.openByUrl(DATABASE_URL);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sheet tab not found.' };
    sheet.getRange(row, column).setValue(payload.value === undefined ? '' : String(payload.value));
    try { CacheService.getScriptCache().remove('iso_module_basedepth'); } catch (cacheError) { /* optional */ }
    return { success: true, row: row, column: column };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/** Update only fields changed in the professional row editor. */
function updateIsoModuleCells(payload) {
  try {
    payload = payload || {};
    var sheetName = String(payload.sheetName || '').trim();
    var row = Number(payload.row);
    var changes = Array.isArray(payload.changes) ? payload.changes : [];
    if (!sheetName || !row || row < 1 || !changes.length) return { success: false, error: 'No valid changes supplied.' };
    var sheet = SpreadsheetApp.openByUrl(DATABASE_URL).getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sheet tab not found.' };
    changes.forEach(function(change) {
      var column = Number(change.column);
      if (column >= 1 && column <= sheet.getMaxColumns()) sheet.getRange(row, column).setValue(change.value === undefined ? '' : String(change.value));
    });
    try {
      var cacheKey = 'iso_module_' + String(payload.moduleKey || '').replace(/[^a-z0-9_]/gi, '_');
      if (payload.moduleKey) CacheService.getScriptCache().remove(cacheKey);
    } catch (cacheError) { /* optional */ }
    return { success: true, count: changes.length };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/** Append a new record using the discovered module header order. */
function appendIsoModuleRecord(payload) {
  try {
    payload = payload || {};
    var sheetName = String(payload.sheetName || '').trim();
    var values = Array.isArray(payload.values) ? payload.values : [];
    if (!sheetName || !values.length) return { success: false, error: 'No module values supplied.' };
    var ss = SpreadsheetApp.openByUrl(DATABASE_URL);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sheet tab not found.' };
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, values.length).setValues([values.map(function(value) { return value === undefined ? '' : String(value); })]);
    try {
      var cacheKey = 'iso_module_' + String(payload.moduleKey || '').replace(/[^a-z0-9_]/gi, '_');
      if (payload.moduleKey) CacheService.getScriptCache().remove(cacheKey);
    } catch (cacheError) { /* optional */ }
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/** Delete one normalized module record by its original Google Sheets row. */
function deleteIsoModuleRecord(payload) {
  try {
    payload = payload || {};
    var sheetRow = Number(payload.sheetRow);
    var moduleKey = String(payload.moduleKey || '').trim();
    var config = getIsoModuleConfig_(moduleKey);
    if (!config || !sheetRow || sheetRow < 1) return { success: false, error: 'Invalid module record.' };
    var sheet = findIsoModuleSheet_(SpreadsheetApp.openByUrl(DATABASE_URL), config);
    if (!sheet) return { success: false, error: 'Module sheet not found.' };
    sheet.deleteRow(sheetRow);
    CacheService.getScriptCache().remove('iso_module_' + moduleKey.replace(/[^a-z0-9_]/gi, '_'));
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}
