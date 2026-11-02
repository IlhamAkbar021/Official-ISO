/**
 * Application entry points.
 *
 * This file contains only deployment-wide configuration and the web-app
 * bootstrap. Feature-specific server code lives in the other JS files.
 */

// Google Sheets data source used by all server-side readers.
// Compatibility alias. New server code should use ISOApp.Config.databaseUrl.
var DATABASE_URL = ISOApp.Config.databaseUrl;

var APP_CONFIG = ISOApp.Config;

// Shared tester directory for every user/device of this Apps Script deployment.
// The first run seeds the two currently registered testers.
var REGISTERED_TESTERS_PROPERTY = ISOApp.Config.testerProperty;
var DEFAULT_REGISTERED_TESTERS = ISOApp.Config.defaultTesters;

function normalizeRegisteredTesters_(names) {
  return ISOApp.Text.normalizeNames(names);
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

/** Includes an HTML partial in index.html. */
function include(filename) {
  return new ISOApp.WebAppController().include(filename);
}

function renderServerError_(title, error) {
  return new ISOApp.WebAppController().error(title, error);
}
