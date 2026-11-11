/**
 * Shared server library.
 *
 * Only top-level adapter functions (for example getWifiData) are public RPC
 * endpoints. Everything in this IIFE is server-internal; implementation
 * details stay private inside the closure.
 */
var ISOApp = ISOApp || {};

ISOApp.Server = (function() {
  function errorText_(error) {
    if (error === null || error === undefined) return 'Unknown server error.';
    return error && error.message ? String(error.message) : String(error);
  }

  function normalizeNames_(names) {
    return (Array.isArray(names) ? names : [])
      .map(function(name) { return String(name || '').trim(); })
      .filter(function(name, index, values) { return name && values.indexOf(name) === index; });
  }

  function findSheet_(spreadsheet, candidateNames) {
    if (!spreadsheet) return null;
    var candidates = normalizeNames_(candidateNames).map(function(name) { return name.toLowerCase(); });
    var sheets = spreadsheet.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (candidates.indexOf(String(sheets[i].getName()).trim().toLowerCase()) !== -1) return sheets[i];
    }
    for (var sheetIndex = 0; sheetIndex < sheets.length; sheetIndex++) {
      var actualName = String(sheets[sheetIndex].getName()).trim().toLowerCase();
      if (candidates.some(function(candidate) { return actualName.indexOf(candidate) !== -1; })) return sheets[sheetIndex];
    }
    return null;
  }

  function scriptCache_() {
    return CacheService.getScriptCache();
  }

  return {
    Result: {
      ok: function(data) {
        var result = { success: true };
        Object.keys(data || {}).forEach(function(key) { result[key] = data[key]; });
        return result;
      },
      fail: function(error, fallback) {
        return { success: false, error: errorText_(error || fallback) };
      },
      errorText: errorText_
    },

    Cache: {
      read: function(key) { return scriptCache_().get(key); },
      write: function(key, value, seconds) {
        try { scriptCache_().put(key, String(value), seconds || 60); } catch (error) { /* cache is optional */ }
        return value;
      },
      readJson: function(key) {
        var raw = scriptCache_().get(key);
        if (!raw) return null;
        try { return JSON.parse(raw); } catch (error) { return null; }
      },
      writeJson: function(key, value, seconds) {
        try { scriptCache_().put(key, JSON.stringify(value), seconds || 60); } catch (error) { /* cache is optional */ }
        return value;
      },
      remove: function(key) {
        try { scriptCache_().remove(key); } catch (error) { /* cache is optional */ }
      },
      removeAll: function(keys) {
        try { scriptCache_().removeAll(keys || []); } catch (error) { /* cache is optional */ }
      }
    },

    Database: {
      open: function() {
        return SpreadsheetApp.openByUrl(ISOApp.Config.databaseUrl);
      },
      activeOrConfigured: function() {
        return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openByUrl(ISOApp.Config.databaseUrl);
      },
      findSheet: findSheet_,
      requireSheet: function(spreadsheet, candidateNames, label) {
        var sheet = findSheet_(spreadsheet, candidateNames);
        if (!sheet) throw new Error((label || candidateNames[0] || 'Requested') + ' sheet was not found.');
        return sheet;
      },
      quoteSheetName: function(name) {
        return "'" + String(name || '').replace(/'/g, "''") + "'";
      }
    },

    Rows: {
      blank: function(width) { return new Array(Math.max(0, Number(width) || 0)).fill(''); },
      strings: function(values, width) {
        var output = new Array(Math.max(0, Number(width) || (values || []).length)).fill('');
        (values || []).slice(0, output.length).forEach(function(value, index) {
          output[index] = value === undefined || value === null ? '' : String(value);
        });
        return output;
      }
    },

    Http: {
      parseJson: function(event) {
        return JSON.parse(event && event.postData && event.postData.contents || '{}');
      },
      json: function(value) {
        return ContentService.createTextOutput(JSON.stringify(value))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
  };
})();

ISOApp.Services = ISOApp.Services || {};
