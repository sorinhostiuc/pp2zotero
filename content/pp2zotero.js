var PP2Zotero = {
  id: null,
  version: null,
  rootURI: null,
  _initialized: false,

  init({ id, version, rootURI }) {
    if (this._initialized) return;
    this.id = id;
    this.version = version;
    this.rootURI = rootURI;
    this._initialized = true;

    this._addMenuWhenReady();
    Zotero.debug("PP2Zotero: initialized v" + version);
  },

  async _addMenuWhenReady() {
    // Wait for Zotero UI to be fully ready before touching the DOM
    if (Zotero.uiReadyPromise) {
      await Zotero.uiReadyPromise;
    }

    const win = Zotero.getMainWindow();
    if (!win) {
      Zotero.debug("PP2Zotero: no main window after uiReadyPromise, giving up");
      return;
    }

    this._addMenuItem();

    // Re-add menu item when a new main window opens (e.g. after close+reopen)
    Zotero.getMainWindows().forEach(w => {
      w.addEventListener("unload", () => {
        // Window closing — nothing to do
      });
    });
  },

  shutdown() {
    this._removeMenuItem();
    this._initialized = false;
    Zotero.debug("PP2Zotero: shutdown");
  },

  _addMenuItem() {
    const menuId = "pp2zotero-menu-convert";
    const doc = Zotero.getMainWindow().document;
    const menuTools = doc.getElementById("menu_ToolsPopup");
    if (!menuTools) return;

    const separator = doc.createXULElement("menuseparator");
    separator.id = "pp2zotero-separator";
    menuTools.appendChild(separator);

    const menuItem = doc.createXULElement("menuitem");
    menuItem.id = menuId;
    menuItem.setAttribute("label", "Convert Paperpile Citations...");
    menuItem.addEventListener("command", () => this.openDialog());
    menuTools.appendChild(menuItem);

    this._menuItemID = menuId;
  },

  _removeMenuItem() {
    const doc = Zotero.getMainWindow()?.document;
    if (!doc) return;
    const item = doc.getElementById(this._menuItemID);
    const sep = doc.getElementById("pp2zotero-separator");
    if (item) item.remove();
    if (sep) sep.remove();
  },

  async readFile(filePath) {
    return IOUtils.read(filePath);
  },

  openDialog() {
    const io = {
      plugin: this,
      rootURI: this.rootURI,
      Zotero: Zotero
    };

    // Use chrome:// URI registered in bootstrap.js
    const dialogURL = "chrome://pp2zotero/content/ui/dialog.xhtml";
    Zotero.debug("PP2Zotero: opening dialog at " + dialogURL);

    Zotero.getMainWindow().openDialog(
      dialogURL,
      "pp2zotero-dialog",
      "chrome,dialog=no,centerscreen,resizable=yes,width=620,height=500",
      io
    );
  }
};
