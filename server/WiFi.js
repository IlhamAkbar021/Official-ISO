/** Wi-Fi domain service. Private helpers are enclosed; global functions are RPC adapters. */
var ISOApp = ISOApp || {};
ISOApp.Services = ISOApp.Services || {};

ISOApp.Services.Wifi = (function() {
  var CACHE_KEY = 'iso_wifi_data';
  var SHEET_ALIASES = ['Wi-Fi Distance/Antenna Diversity', 'WiFi', 'Wi-Fi'];
  var FIRST_DATA_ROW = 7;

  function sheet_() {
    return ISOApp.Server.Database.requireSheet(ISOApp.Server.Database.open(), SHEET_ALIASES, 'Wi-Fi');
  }

  function invalidate_() {
    ISOApp.Server.Cache.remove(CACHE_KEY);
    ISOApp.Server.Cache.remove('iso_static_input_choices');
  }

  function value_(value) {
    return value === undefined || value === null ? '' : value;
  }

  function list() {
    var cached = ISOApp.Server.Cache.readJson(CACHE_KEY);
    if (cached) return cached;
    try {
      var sheet = sheet_();
      var apValue = sheet.getRange('D2').getDisplayValue().trim();
      var distanceValue = sheet.getRange('D3').getDisplayValue().trim();
      var signalValue = sheet.getRange('D4').getDisplayValue().trim();
      if (!distanceValue || distanceValue === '30.00' || distanceValue === '30') distanceValue = '10';
      var range = sheet.getDataRange();
      var displayValues = range.getDisplayValues().slice(5);
      var richTextValues = range.getRichTextValues().slice(5);
      var formulas = range.getFormulas().slice(5);
      var result = {
        setup: { apName: apValue || 'meshtest', distance: distanceValue, signal: signalValue || '-55' },
        headers: displayValues.length ? displayValues[0] : [],
        rows: []
      };
      if (displayValues.length > 1) {
        result.rows = displayValues.slice(1).map(function(row, rowIndex) {
          return row.map(function(cellValue, columnIndex) {
            var richText = richTextValues[rowIndex + 1] ? richTextValues[rowIndex + 1][columnIndex] : null;
            var url = richText ? richText.getLinkUrl() : null;
            var formula = formulas[rowIndex + 1] ? formulas[rowIndex + 1][columnIndex] : '';
            if (!url && formula) {
              var match = formula.match(/HYPERLINK\(\s*["']([^"']+)["']/i);
              if (match) url = match[1];
            }
            return url || String(cellValue || '').trim();
          });
        });
      }
      return ISOApp.Server.Cache.writeJson(CACHE_KEY, result, 60);
    } catch (error) {
      return { error: ISOApp.Server.Result.errorText(error), setup: {}, headers: [], rows: [] };
    }
  }

  function create(payload) {
    try {
      payload = payload || {};
      var sheet = sheet_();
      var lastRow = sheet.getLastRow();
      var width = sheet.getLastColumn();
      sheet.getRange(lastRow, 1, 1, width).copyTo(sheet.getRange(lastRow + 1, 1));
      var lastId = sheet.getRange(lastRow, 1).getValue();
      var newId = typeof lastId === 'number' ? lastId + 1 : '';
      [[1, newId], [2, payload.robotId], [3, payload.tester], [4, payload.rightAntenna],
       [8, payload.rightLink], [10, payload.leftAntenna], [14, payload.leftLink]]
        .forEach(function(update) { sheet.getRange(lastRow + 1, update[0]).setValue(value_(update[1])); });
      invalidate_();
      return ISOApp.Server.Result.ok({ row: lastRow + 1 });
    } catch (error) {
      return ISOApp.Server.Result.fail(error, 'Wi-Fi record could not be created.');
    }
  }

  function update(payload) {
    try {
      payload = payload || {};
      var sheet = sheet_();
      var targetRow = Number(payload.rowIndex) + FIRST_DATA_ROW;
      if (!isFinite(targetRow) || targetRow < FIRST_DATA_ROW) throw new Error('Invalid Wi-Fi row index.');
      [[3, payload.tester], [4, payload.rightAntenna], [8, payload.rightLink],
       [10, payload.leftAntenna], [14, payload.leftLink]]
        .forEach(function(updateValue) { sheet.getRange(targetRow, updateValue[0]).setValue(value_(updateValue[1])); });
      invalidate_();
      return ISOApp.Server.Result.ok({ row: targetRow });
    } catch (error) {
      return ISOApp.Server.Result.fail(error, 'Wi-Fi record could not be updated.');
    }
  }

  function remove(payload) {
    try {
      var targetRow = Number(payload && payload.rowIndex) + FIRST_DATA_ROW;
      if (!isFinite(targetRow) || targetRow < FIRST_DATA_ROW) throw new Error('Invalid Wi-Fi row index.');
      sheet_().deleteRow(targetRow);
      invalidate_();
      return ISOApp.Server.Result.ok({ row: targetRow });
    } catch (error) {
      return ISOApp.Server.Result.fail(error, 'Wi-Fi record could not be deleted.');
    }
  }

  return Object.freeze({ list: list, create: create, update: update, remove: remove });
})();

// Public RPC adapters. Keep these global for google.script.run compatibility.
function getWifiData() { return ISOApp.Services.Wifi.list(); }
function addWifiRecord(payload) { return ISOApp.Services.Wifi.create(payload); }
function updateWifiRecord(payload) { return ISOApp.Services.Wifi.update(payload); }
function deleteWifiRecord(payload) { return ISOApp.Services.Wifi.remove(payload); }
