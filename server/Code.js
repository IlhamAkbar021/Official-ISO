/**
 * Application entry points.
 *
 * This file contains only deployment-wide configuration and the web-app
 * bootstrap. Feature-specific server code lives in the other JS files.
 */

// Force the next Static ISO read to come directly from Google Sheets.
function clearStaticIsoCache() {
  var keys = [
    'iso_lidar_data', 'iso_wifi_data', 'iso_overview_v3',
    'iso_static_input_choices',
    'iso_module_headdepth', 'iso_module_basedepth', 'iso_module_cliff',
    'iso_module_apriltag', 'iso_module_estop', 'iso_module_ratedspeed',
    'iso_module_turningwidth', 'iso_module_clifflowerspeed',
    'iso_module_traversegap', 'iso_module_narrow'
  ];
  ISOApp.Server.Cache.removeAll(keys);
  return ISOApp.Server.Result.ok();
}

function getStaticIsoInputChoices() {
  return ISOApp.StaticISO.getInputChoices();
}

function getRegisteredTesterNames() {
  return new ISOApp.TesterService().list();
}

function saveRegisteredTesterNames(names) {
  return new ISOApp.TesterService().save(names);
}

function doGet() {
  try {
    return new ISOApp.WebAppController().get();
  } catch (err) {
    return new ISOApp.WebAppController().error('App Template Error', err);
  }
}

/** Optional JSON web endpoint; the in-app UI normally uses google.script.run. */
function doPost(event) {
  try {
    return ISOApp.Server.Http.json(saveLidarData(ISOApp.Server.Http.parseJson(event)));
  } catch (error) {
    return ISOApp.Server.Http.json(ISOApp.Server.Result.fail(error));
  }
}

/** Public RPC adapter for LiDAR saves. */
function saveLidarData(data) {
  try {
    if (!data || typeof data !== 'object' || Array.isArray(data)) {
      return ISOApp.Server.Result.fail('A valid LiDAR save payload is required.');
    }
    return ISOApp.Services.Lidar.save(data);
  } catch (error) {
    return ISOApp.Server.Result.fail(error, 'LiDAR save failed.');
  }
}

function getLidarData(forceRefresh) {
  return ISOApp.Services.Lidar.list(forceRefresh);
}

function addLidarRecord(payload) {
  return ISOApp.Services.Lidar.create(payload);
}

function smartImportLidarRecord(payload) {
  return ISOApp.Services.Lidar.smartImport(payload);
}

function updateLidarRecord(payload) {
  return ISOApp.Services.Lidar.update(payload);
}

function deleteLidarRecord(payload) {
  return ISOApp.Services.Lidar.remove(payload);
}

// Temporary compatibility adapters for older deployed clients. Both use the
// same service path, so an edit can never accidentally append a duplicate row.
function saveRecord(data) { return saveLidarData(data); }
function fastSaveRecord(data) { return saveLidarData(data); }

/** Includes an HTML partial in index.html. */
function include(filename) {
  return new ISOApp.WebAppController().include(filename);
}
