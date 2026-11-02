/**
 * Application core layer.
 *
 * Apps Script project files are loaded into one global runtime, so this file
 * provides explicit namespaces/classes instead of adding more loose globals.
 * Public Apps Script entry points remain in Code.js as compatibility adapters.
 */
var ISOApp = ISOApp || {};

ISOApp.Config = {
  databaseUrl: 'https://docs.google.com/spreadsheets/d/1pV5pDJ3Q_eDbIhRUpIUY5-0JjAr9Yb_JUYegFBmMewc/edit?gid=1990766678#gid=1990766678',
  title: 'AC3110 DVT ISO test',
  template: 'index',
  viewport: 'width=device-width, initial-scale=1',
  testerProperty: 'registered_testers',
  defaultTesters: ['Kalman', 'Yuvenika']
};

ISOApp.Text = {
  normalizeNames: function(names) {
    var unique = [];
    (Array.isArray(names) ? names : []).forEach(function(name) {
      name = String(name || '').trim();
      if (name && unique.indexOf(name) === -1) unique.push(name);
    });
    return unique;
  }
};

ISOApp.TesterService = class {
  constructor(properties) {
    this.properties = properties || PropertiesService.getScriptProperties();
    this.key = ISOApp.Config.testerProperty;
  }

  list() {
    var raw = this.properties.getProperty(this.key);
    var names;
    try {
      names = raw ? JSON.parse(raw) : ISOApp.Config.defaultTesters;
    } catch (error) {
      names = ISOApp.Config.defaultTesters;
    }
    names = ISOApp.Text.normalizeNames(names);
    if (!raw) this.save(names);
    return names;
  }

  save(names) {
    var normalized = ISOApp.Text.normalizeNames(names);
    this.properties.setProperty(this.key, JSON.stringify(normalized));
    return { success: true, names: normalized };
  }
};

ISOApp.WebAppController = class {
  get() {
    return HtmlService.createTemplateFromFile(ISOApp.Config.template)
      .evaluate()
      .setTitle(ISOApp.Config.title)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', ISOApp.Config.viewport);
  }

  include(filename) {
    try {
      return HtmlService.createHtmlOutputFromFile(filename).getContent();
    } catch (error) {
      return '<!-- Error loading ' + filename + ': ' + error.toString() + ' -->';
    }
  }

  error(title, error) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif; padding:2rem; color:#dc2626; background:#fee2e2; border-radius:8px; margin:2rem;">' +
      '<h3>⚠️ ' + title + '</h3>' +
      '<p>' + error.toString() + '</p>' +
      '</div>'
    );
  }
};
