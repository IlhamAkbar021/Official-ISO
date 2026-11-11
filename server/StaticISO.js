/** Static ISO domain metadata. Shared behavior lives in ServerLibrary.js. */
var ISOApp = ISOApp || {};
ISOApp.StaticISO = ISOApp.StaticISO || {
  cacheKeys: {
    lidar: 'iso_lidar_data',
    wifi: 'iso_wifi_data',
    headDepth: 'iso_module_headdepth',
    baseDepth: 'iso_module_basedepth',
    cliff: 'iso_module_cliff'
  }
};

/** Build shared Robot ID and SW suggestions from every Static ISO data tab. */
function getStaticIsoInputChoices_() {
  var cacheKey = 'iso_static_input_choices';
  var cached = ISOApp.Server.Cache.readJson(cacheKey);
  if (cached) return cached;

  try {
    var robotNames = [];
    var swVersions = [];
    var robotSeen = {};
    var swSeen = {};
    function collect_(target, seen, value) {
      value = String(value || '').trim();
      if (!value || value === '—' || value === '-') return;
      var key = value.toLowerCase();
      if (!seen[key]) { seen[key] = true; target.push(value); }
    }
    function collectSheet_(sheet, moduleKey) {
      if (!sheet) return;
      var values = sheet.getDataRange().getDisplayValues();
      if (!values.length) return;
      var config = getStaticIsoModuleConfig_(moduleKey) || {};
      var bounds = findIsoSectionBounds_(values, config);
      var section = values.slice(bounds.start, bounds.end);
      if (!section.length) return;
      var headerIndex = detectIsoHeaderRow_(section);
      var headers = section[headerIndex] || [];
      var robotColumn = headers.findIndex(function(header) { return /robot\s*id|unit\s*id|serial/i.test(String(header || '')); });
      var swColumn = headers.findIndex(function(header) { return /\bsw\b|software|version|release/i.test(String(header || '')); });
      if (robotColumn < 0 && moduleKey === 'headdepth') robotColumn = 0;
      if (swColumn < 0 && moduleKey === 'headdepth') swColumn = 2;
      for (var rowIndex = headerIndex + 1; rowIndex < section.length; rowIndex++) {
        var row = section[rowIndex] || [];
        if (robotColumn >= 0) collect_(robotNames, robotSeen, row[robotColumn]);
        if (swColumn >= 0) collect_(swVersions, swSeen, row[swColumn]);
      }
    }

    var spreadsheet = ISOApp.Server.Database.open();
    ['headdepth', 'basedepth', 'cliff'].forEach(function(moduleKey) {
      try {
        var config = getStaticIsoModuleConfig_(moduleKey);
        collectSheet_(findIsoModuleSheet_(spreadsheet, config), moduleKey);
      } catch (moduleError) { console.warn('Static ISO choices skipped for ' + moduleKey, moduleError); }
    });

    // LiDAR has a dedicated reader because its sheet contains merged rows and
    // two record types. Use its normalized accuracy records for the choices.
    try {
      var lidar = ISOApp.Services.Lidar.list();
      ((lidar && lidar.accuracy && lidar.accuracy.rows) || []).forEach(function(row) {
        collect_(robotNames, robotSeen, row[0]);
        collect_(swVersions, swSeen, row[2]);
      });
      ((lidar && lidar.tilt && lidar.tilt.rows) || []).forEach(function(row) {
        collect_(robotNames, robotSeen, row[0]);
        collect_(swVersions, swSeen, row[2]);
      });
    } catch (lidarError) { console.warn('LiDAR choices unavailable', lidarError); }

    // Include the existing Overview registry so newly registered robots are
    // selectable before they have a measurement row in a Static ISO tab.
    try {
      var overview = getOverviewData_();
      ((overview && overview.robots) || []).forEach(function(robot) {
        collect_(robotNames, robotSeen, robot.name);
        collect_(swVersions, swSeen, robot.sw);
      });
    } catch (overviewError) { /* sheet-based options remain available */ }

    // Wi-Fi has no SW field, but its robot IDs are part of the shared Static
    // ISO list. Its data starts at row 7; robot IDs are in column B.
    try {
      var wifiSheet = ISOApp.Server.Database.findSheet(spreadsheet, ['Wi-Fi Distance/Antenna Diversity', 'WiFi', 'Wi-Fi']);
      if (wifiSheet && wifiSheet.getLastRow() >= 7) {
        wifiSheet.getRange(7, 2, wifiSheet.getLastRow() - 6, 1).getDisplayValues().forEach(function(row) {
          collect_(robotNames, robotSeen, row[0]);
        });
      }
    } catch (wifiError) { /* optional module */ }

    robotNames.sort(function(a, b) { return a.localeCompare(b, undefined, { numeric:true, sensitivity:'base' }); });
    swVersions.sort(function(a, b) { return a.localeCompare(b, undefined, { numeric:true, sensitivity:'base' }); });
    return ISOApp.Server.Cache.writeJson(cacheKey, { success:true, names:robotNames, swVersions:swVersions }, 60);
  } catch (error) {
    return { success:false, error:ISOApp.Server.Result.errorText(error), names:[], swVersions:[] };
  }
}

ISOApp.StaticISO.getInputChoices = getStaticIsoInputChoices_;
