/**
 * ISO module registry.
 *
 * Keep all workbook-tab aliases and UI metadata in one place.  The reader in
 * Modules.js only asks this registry for configuration; it does not know
 * anything about individual ISO modules.
 */
var ISO_MODULE_DEFINITIONS = {
  static: {
    headdepth: { title: 'Head Depth Camera', phase: 'Static ISO', aliases: ['head depth camera', 'head depth', 'head rgbd', 'head camera', 'rgbd head'] },
    basedepth: { title: 'Base Depth Camera', phase: 'Static ISO', aliases: ['base depth camera', 'base depth', 'base rgbd', 'bottom depth', 'rgbd base'] },
    cliff: { title: 'IR Cliff Sensor', phase: 'Static ISO', aliases: ['ir cliff', 'cliff sensor', 'cliff'] },
    wifi: { title: 'Wi-Fi', phase: 'Static ISO', customReader: true, aliases: ['wi-fi distance/antenna diversity', 'wi-fi', 'wifi'] }
  },
  dynamic: {
    ratedspeed: { title: 'Rated Speed', phase: 'Dynamic ISO', aliases: ['rated speed', 'rate speed', 'speed test', 'robot speed'] },
    estop: { title: 'E-Stop', phase: 'Dynamic ISO', aliases: ['e-stop', 'estop', 'emergency stop', 'stop characteristic', 'iso 18646-1 clause 6'] },
    turningwidth: { title: 'Turning Width', phase: 'Dynamic ISO', aliases: ['turning width', 'turn width', 'turning diameter', 'turning test'] },
    clifflowerspeed: { title: 'Cliff at Lower Speed', phase: 'Dynamic ISO', aliases: ['cliff with lower speed', 'cliff lower speed', 'cliff at lower speed', 'low speed cliff'] },
    traversegap: { title: 'Traverse Gap', phase: 'Dynamic ISO', aliases: ['traverse gap', 'gap traverse', 'gap test', 'traverse'] },
    narrow: { title: 'Narrow Corridor', phase: 'Dynamic ISO', aliases: ['narrow corridor', 'narrow passage', 'corridor'] },
    apriltag: { title: 'AprilTag', phase: 'Dynamic ISO', aliases: ['apriltag', 'april tag', 'tag detection'] }
  },
  accuracy: {
    accuracylaser: { title: 'Accuracy & Repeatability — Laser Detection', phase: 'Accuracy & Repeatability', aliases: ['accuracy and repeatability laser detection', 'accuracy repeatability laser', 'accuracy and repeatabilty laser', 'a&r laser detection', 'laser detection a&r'], fallbackAliases: ['accuracy and repeatability', 'accuracy and repeatabilty', 'accuracy repeatability'], sectionKeywords: ['laser detection', 'lidar detection', 'laser'], stopKeywords: ['apriltag detection', 'april tag detection'] },
    accuracyapriltag: { title: 'Accuracy & Repeatability — AprilTag Detection', phase: 'Accuracy & Repeatability', aliases: ['accuracy and repeatability apriltag detection', 'accuracy repeatability apriltag', 'accuracy and repeatabilty april tag', 'a&r apriltag detection', 'april tag detection a&r'], fallbackAliases: ['accuracy and repeatability', 'accuracy and repeatabilty', 'accuracy repeatability'], sectionKeywords: ['apriltag detection', 'april tag detection', 'apriltag'], stopKeywords: ['laser detection', 'lidar detection'] }
  }
};

/** Small registry wrapper keeps lookup logic out of the data reader. */
function IsoModuleRegistry(definitions) {
  this.definitions = definitions || {};
}

IsoModuleRegistry.prototype.get = function(moduleKey) {
  var groups = ['static', 'dynamic', 'accuracy'];
  for (var i = 0; i < groups.length; i++) {
    var config = this.definitions[groups[i]][moduleKey];
    if (config) return config;
  }
  return null;
};

var ISO_MODULE_REGISTRY = new IsoModuleRegistry(ISO_MODULE_DEFINITIONS);

function getIsoModuleConfig_(moduleKey) { return ISO_MODULE_REGISTRY.get(moduleKey); }

function getStaticIsoModuleConfig_(moduleKey) { return ISO_MODULE_DEFINITIONS.static[moduleKey] || null; }
function getDynamicIsoModuleConfig_(moduleKey) { return ISO_MODULE_DEFINITIONS.dynamic[moduleKey] || null; }
function getAccuracyIsoModuleConfig_(moduleKey) { return ISO_MODULE_DEFINITIONS.accuracy[moduleKey] || null; }
