/**
 * Application entry points.
 *
 * This file contains only deployment-wide configuration and the web-app
 * bootstrap. Feature-specific server code lives in the other JS files.
 */

// Google Sheets data source used by all server-side readers.
var DATABASE_URL = 'https://docs.google.com/spreadsheets/d/1pV5pDJ3Q_eDbIhRUpIUY5-0JjAr9Yb_JUYegFBmMewc/edit?gid=1990766678#gid=1990766678';

var APP_CONFIG = {
  title: 'AC3110 DVT ISO test',
  template: 'index',
  viewport: 'width=device-width, initial-scale=1'
};

// Shared tester directory for every user/device of this Apps Script deployment.
// The first run seeds the two currently registered testers.
var REGISTERED_TESTERS_PROPERTY = 'registered_testers';
var DEFAULT_REGISTERED_TESTERS = ['Kalman', 'Yuvenika'];

function normalizeRegisteredTesters_(names) {
  var unique = [];
  (Array.isArray(names) ? names : []).forEach(function(name) {
    name = String(name || '').trim();
    if (name && unique.indexOf(name) === -1) unique.push(name);
  });
  return unique;
}

function getRegisteredTesterNames() {
  var properties = PropertiesService.getScriptProperties();
  var raw = properties.getProperty(REGISTERED_TESTERS_PROPERTY);
  var names;
  try {
    names = raw ? JSON.parse(raw) : DEFAULT_REGISTERED_TESTERS;
  } catch (e) {
    names = DEFAULT_REGISTERED_TESTERS;
  }
  names = normalizeRegisteredTesters_(names);
  if (!raw) properties.setProperty(REGISTERED_TESTERS_PROPERTY, JSON.stringify(names));
  return names;
}

function saveRegisteredTesterNames(names) {
  var normalized = normalizeRegisteredTesters_(names);
  PropertiesService.getScriptProperties().setProperty(
    REGISTERED_TESTERS_PROPERTY,
    JSON.stringify(normalized)
  );
  return { success: true, names: normalized };
}

function doGet() {
  try {
    return HtmlService.createTemplateFromFile(APP_CONFIG.template)
        .evaluate()
        .setTitle(APP_CONFIG.title)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', APP_CONFIG.viewport);
  } catch (err) {
    return renderServerError_('App Template Error', err);
  }
}

/** Includes an HTML partial in index.html. */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    return '<!-- Error loading ' + filename + ': ' + e.toString() + ' -->';
  }
}

function renderServerError_(title, error) {
  return HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif; padding:2rem; color:#dc2626; background:#fee2e2; border-radius:8px; margin:2rem;">' +
    '<h3>⚠️ ' + title + '</h3>' +
    '<p>' + error.toString() + '</p>' +
    '</div>'
  );
}
