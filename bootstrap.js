var PP2Zotero;
var chromeHandle;

// Polyfill setImmediate for JSZip (not available in Gecko)
if (typeof setImmediate === "undefined") {
  var setImmediate = fn => setTimeout(fn, 0);
}

function install(data, reason) {}

function startup({ id, version, resourceURI, rootURI }, reason) {
  // Register chrome://pp2zotero/content/ mapping
  var aomStartup = Cc["@mozilla.org/addons/addon-manager-startup;1"]
    .getService(Ci.amIAddonManagerStartup);
  var manifestURI = Services.io.newURI(rootURI + "manifest.json");
  chromeHandle = aomStartup.registerChrome(manifestURI, [
    ["content", "pp2zotero", rootURI + "content/"]
  ]);

  var chromeBase = "chrome://pp2zotero/content/";
  Services.scriptloader.loadSubScript(chromeBase + "lib/jszip.min.js");
  Services.scriptloader.loadSubScript(chromeBase + "pp2zotero.js");
  Services.scriptloader.loadSubScript(chromeBase + "scanner.js");
  Services.scriptloader.loadSubScript(chromeBase + "matcher.js");
  Services.scriptloader.loadSubScript(chromeBase + "converter.js");
  Services.scriptloader.loadSubScript(chromeBase + "odf-writer.js");

  PP2Zotero.init({ id, version, rootURI });
}

function shutdown({ id, version, resourceURI, rootURI }, reason) {
  if (reason === APP_SHUTDOWN) return;
  PP2Zotero.shutdown();
  PP2Zotero = undefined;
  if (chromeHandle) {
    chromeHandle.destruct();
    chromeHandle = null;
  }
}

function uninstall(data, reason) {}
