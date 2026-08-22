/* Dialog controller — manages 4 screens and communicates with plugin modules */

(function () {
  "use strict";

  const io = window.arguments ? window.arguments[0] : null;
  const Zotero = io ? io.Zotero : null;
  const rootURI = io ? io.rootURI : "";
  const plugin = io ? io.plugin : null;

  let selectedFilePath = null;
  let selectedFilePaths = [];
  let scanResults = null;
  let matchResults = null;
  let uniqueRefs = [];
  let convertStats = {};
  let batchResults = [];

  function showScreen(id) {
    document.querySelectorAll(".screen").forEach(s => s.classList.remove("active"));
    document.getElementById(id).classList.add("active");

    if (id === "screen-results") {
      window.resizeTo(820, 600);
    } else if (id === "screen-progress") {
      window.resizeTo(600, 300);
    } else if (id === "screen-complete") {
      window.resizeTo(600, 500);
    } else {
      window.resizeTo(620, 820);
    }
  }

  function initScreen1() {
    const dropZone = document.getElementById("file-drop-zone");
    const btnScan = document.getElementById("btn-scan");

    dropZone.addEventListener("click", async () => {
      const fp = Components.classes["@mozilla.org/filepicker;1"]
        .createInstance(Components.interfaces.nsIFilePicker);
      fp.init(window, "Select Word document(s)", fp.modeOpenMultiple);
      fp.appendFilter("Word Documents", "*.docx");

      const result = await new Promise(resolve => fp.open(resolve));
      if (result === fp.returnOK) {
        selectedFilePaths = [];
        const files = fp.files;
        while (files.hasMoreElements()) {
          const f = files.getNext().QueryInterface(Components.interfaces.nsIFile);
          selectedFilePaths.push(f.path);
        }
        if (selectedFilePaths.length > 0) {
          selectedFilePath = selectedFilePaths[0];
          if (selectedFilePaths.length === 1) {
            showFileInfo(fp.file);
          } else {
            showBatchFileInfo(selectedFilePaths);
          }
        }
      }
    });

    dropZone.addEventListener("dragover", e => { e.preventDefault(); dropZone.classList.add("dragover"); });
    dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragover"));
    dropZone.addEventListener("drop", e => {
      e.preventDefault();
      dropZone.classList.remove("dragover");
      if (e.dataTransfer.files.length > 0) {
        selectedFilePaths = [];
        for (let i = 0; i < e.dataTransfer.files.length; i++) {
          const file = e.dataTransfer.files[i];
          if (file.name.endsWith(".docx")) {
            selectedFilePaths.push(file.mozFullPath || file.path);
          }
        }
        if (selectedFilePaths.length === 1) {
          selectedFilePath = selectedFilePaths[0];
          showFileInfo(e.dataTransfer.files[0]);
        } else if (selectedFilePaths.length > 1) {
          selectedFilePath = selectedFilePaths[0];
          showBatchFileInfo(selectedFilePaths);
        }
      }
    });

    btnScan.addEventListener("click", () => startScan());
    document.getElementById("btn-cancel-1").addEventListener("click", () => window.close());
    document.getElementById("btn-new-collection").addEventListener("click", () => createNewCollection());

    initSegControl();
    initCollectionPicker();
  }

  /* ---- Segmented controls ---- */

  let _matchStrategy = "doi_title_year";
  let _fieldMode = "fieldcodes";

  function initSegControl() {
    const btns = document.querySelectorAll("#opt-match-strategy .seg-btn");
    btns.forEach(btn => {
      btn.addEventListener("click", () => {
        btns.forEach(b => b.classList.remove("active"));
        btn.classList.add("active");
        _matchStrategy = btn.dataset.value;
      });
    });

    const fmBtns = document.querySelectorAll("#opt-field-mode .seg-btn");
    fmBtns.forEach(btn => {
      btn.addEventListener("click", () => {
        fmBtns.forEach(b => b.classList.remove("active"));
        btn.classList.add("active");
        _fieldMode = btn.dataset.value;
      });
    });
  }

  function getMatchStrategy() {
    return _matchStrategy;
  }

  /* ---- Collection picker (button + Services.prompt.select) ---- */

  let _selectedCollectionID = null;
  let _collectionsList = []; // { id, name }

  function initCollectionPicker() {
    _buildCollectionsList();

    const btn = document.getElementById("opt-collection-btn");
    btn.addEventListener("click", () => {
      if (_collectionsList.length === 0) {
        _buildCollectionsList();
      }
      const labels = _collectionsList.map(c => c.name);
      const selected = { value: 0 };

      // Find current selection index
      if (_selectedCollectionID) {
        const idx = _collectionsList.findIndex(c => c.id === _selectedCollectionID);
        if (idx >= 0) selected.value = idx;
      }

      const ok = Services.prompt.select(
        window, "PP2Zotero", "Select collection:", labels, selected
      );
      if (ok) {
        const col = _collectionsList[selected.value];
        _selectedCollectionID = col.id;
        btn.textContent = col.name.replace(/^\u00A0+/, "").trim() || "My Library";
      }
    });

    // Pre-select current collection if one is active
    try {
      const activeCollection = Zotero.getActiveZoteroPane()?.getSelectedCollection();
      if (activeCollection) {
        _selectedCollectionID = activeCollection.id;
        btn.textContent = activeCollection.name;
      }
    } catch (e) { /* ignore */ }
  }

  function _buildCollectionsList() {
    _collectionsList = [{ id: null, name: "My Library" }];
    const libraryID = Zotero.Libraries.userLibraryID;
    const collections = Zotero.Collections.getByLibrary(libraryID, true);

    const buildTree = (parentID, depth) => {
      const children = collections.filter(c => (c.parentID || 0) === parentID);
      children.sort((a, b) => a.name.localeCompare(b.name));
      for (const col of children) {
        _collectionsList.push({ id: col.id, name: "\u00A0\u00A0".repeat(depth) + col.name });
        buildTree(col.id, depth + 1);
      }
    };

    const topLevel = collections.filter(c => !c.parentID);
    topLevel.sort((a, b) => a.name.localeCompare(b.name));
    for (const col of topLevel) {
      _collectionsList.push({ id: col.id, name: col.name });
      buildTree(col.id, 1);
    }
  }


  function showFileInfo(file) {
    document.getElementById("file-drop-zone").classList.add("has-file");
    const fileSelected = document.getElementById("file-selected");
    fileSelected.classList.add("visible");
    document.getElementById("file-name").textContent = file.leafName || file.name;

    const sizeKB = Math.round((file.fileSize || file.size) / 1024);
    const sizeMB = (sizeKB / 1024).toFixed(1);
    const sizeStr = sizeKB > 1024 ? sizeMB + " MB" : sizeKB + " KB";
    document.getElementById("file-meta").textContent = sizeStr;

    if (selectedFilePaths.length === 0) {
      selectedFilePaths = [selectedFilePath];
    }
    document.getElementById("btn-scan").disabled = false;
  }

  function showBatchFileInfo(paths) {
    document.getElementById("file-drop-zone").classList.add("has-file");
    const fileSelected = document.getElementById("file-selected");
    fileSelected.classList.add("visible");
    document.getElementById("file-name").textContent = paths.length + " files selected";
    const names = paths.map(p => p.split("/").pop()).join(", ");
    document.getElementById("file-meta").textContent = names.length > 60 ? names.substring(0, 57) + "..." : names;
    document.getElementById("btn-scan").disabled = false;
  }

  async function startScan() {
    if (selectedFilePaths.length > 1) {
      startBatch();
      return;
    }

    const btnScan = document.getElementById("btn-scan");
    btnScan.disabled = true;
    btnScan.textContent = "Scanning...";

    try {
      const fileData = await plugin.readFile(selectedFilePath);
      scanResults = await plugin.Scanner.scan(fileData);

      if (scanResults.citations.length === 0) {
        Zotero.alert(window, "PP2Zotero", "No Paperpile citations found in this document.");
        btnScan.disabled = false;
        btnScan.textContent = "\uD83D\uDD0D Scan document";
        return;
      }

      const libraryID = Zotero.Libraries.userLibraryID;
      const strategy = getMatchStrategy();
      await plugin.Matcher.buildIndexes(libraryID);
      matchResults = plugin.Matcher.matchAll(scanResults.citations, strategy);

      if (document.getElementById("opt-auto-import").checked) {
        await autoImportMissing(libraryID);
      }

      // Add already-matched items to selected collection
      await addMatchedToCollection(libraryID);

      deduplicateRefs();
      showScreen("screen-results");
      renderResults();

    } catch (e) {
      Zotero.alert(window, "PP2Zotero Error", "Scan failed: " + e.message);
      Zotero.debug("PP2Zotero scan error: " + e.stack);
      btnScan.disabled = false;
      btnScan.textContent = "\uD83D\uDD0D Scan document";
    }
  }

  async function startBatch() {
    showScreen("screen-progress");
    const startTime = Date.now();
    const overwrite = document.getElementById("opt-overwrite").checked;
    const autoImport = document.getElementById("opt-auto-import").checked;
    const strategy = getMatchStrategy();
    const libraryID = Zotero.Libraries.userLibraryID;

    batchResults = [];
    let totalConverted = 0, totalCitations = 0, totalSkipped = 0;

    try {
      await plugin.Matcher.buildIndexes(libraryID);

      for (let fi = 0; fi < selectedFilePaths.length; fi++) {
        const filePath = selectedFilePaths[fi];
        const fileName = filePath.split("/").pop();
        const pct = Math.round((fi / selectedFilePaths.length) * 90);
        updateProgress(pct, "Processing " + (fi + 1) + "/" + selectedFilePaths.length + ": " + fileName);

        try {
          const fileData = await plugin.readFile(filePath);
          const scan = await plugin.Scanner.scan(fileData);

          if (scan.citations.length === 0) {
            batchResults.push({ file: fileName, status: "skipped", reason: "No Paperpile citations" });
            continue;
          }

          const matches = plugin.Matcher.matchAll(scan.citations, strategy);

          if (autoImport) {
            // Reuse autoImportMissing logic inline
            matchResults = matches;
            await autoImportMissing(libraryID);
          }

          // Add already-matched items to selected collection
          matchResults = matches;
          await addMatchedToCollection(libraryID);

          // Convert
          let outputPath;
          const batchIsODT = _fieldMode === "referencemarks";
          if (overwrite && !batchIsODT) {
            outputPath = filePath;
          } else if (batchIsODT) {
            outputPath = filePath.replace(/\.docx$/i, "_zotero.odt");
          } else {
            outputPath = filePath.replace(/\.docx$/i, "_zotero.docx");
          }

          if (document.getElementById("opt-backup").checked && overwrite) {
            await plugin.Converter.createBackup(filePath);
          }

          const hasSkipped = matches.some(mr =>
            mr.itemMatches.every(m => m.matchType === "none" || m.matchType === "skipped" || m.action === "skip")
          );

          for (const mr of matches) {
            for (const im of mr.itemMatches) {
              if (im.matchType === "none" || im.matchType === "skipped") {
                im.action = "skip";
              }
            }
          }

          const newData = await plugin.Converter.convert(fileData, matches, { addComments: hasSkipped, fieldMode: _fieldMode });
          await plugin.Converter.saveFile(outputPath, newData);

          let converted = 0, skipped = 0;
          for (const mr of matches) {
            const allSkip = mr.itemMatches.every(m => !m.match);
            if (!allSkip) converted++;
            else skipped++;
          }

          totalConverted += converted;
          totalCitations += scan.citations.length;
          totalSkipped += skipped;

          batchResults.push({
            file: fileName,
            output: outputPath.split("/").pop(),
            status: "ok",
            citations: scan.citations.length,
            converted, skipped
          });
        } catch (e) {
          batchResults.push({ file: fileName, status: "error", reason: e.message });
          Zotero.debug("PP2Zotero batch error on " + fileName + ": " + e.stack);
        }
      }

      const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
      convertStats = {
        citationsConverted: totalConverted,
        citationsTotal: totalCitations,
        uniqueConverted: totalConverted,
        uniqueTotal: totalConverted + totalSkipped,
        matchedDOI: 0, matchedTitle: 0, duplicateFound: 0, importedDOI: 0,
        skipped: totalSkipped,
        backupPath: null,
        elapsed,
        outputPath: selectedFilePaths.length + " files processed",
        isBatch: true,
        batchResults
      };

      updateProgress(100, "Complete!");
      setTimeout(() => {
        showScreen("screen-complete");
        renderComplete();
      }, 500);

    } catch (e) {
      Zotero.alert(window, "PP2Zotero Error", "Batch conversion failed: " + e.message);
      Zotero.debug("PP2Zotero batch error: " + e.stack);
      showScreen("screen-file");
    }
  }

  async function createNewCollection() {
    const name = { value: "" };
    const ok = Services.prompt.prompt(window, "PP2Zotero", "New collection name:", name, null, {});
    if (!ok || !name.value.trim()) return;

    try {
      const libraryID = Zotero.Libraries.userLibraryID;
      const collection = new Zotero.Collection();
      collection.libraryID = libraryID;
      collection.name = name.value.trim();

      // If a parent collection is selected, create as subcollection
      const parentID = getSelectedCollectionID();
      if (parentID) {
        collection.parentID = parentID;
      }

      await collection.saveTx();

      // Refresh list and select the new collection
      _buildCollectionsList();
      _selectedCollectionID = collection.id;
      document.getElementById("opt-collection-btn").textContent = name.value.trim();

      Zotero.debug("PP2Zotero: Created new collection: " + name.value.trim());
    } catch (e) {
      Zotero.debug("PP2Zotero: Failed to create collection: " + e.message);
    }
  }

  function getSelectedCollectionID() {
    return _selectedCollectionID || null;
  }

  function _refKey(cslData) {
    return cslData.DOI
      ? "doi:" + cslData.DOI.toLowerCase()
      : "title:" + (cslData.title || "").toLowerCase() + "|" + plugin.Matcher._extractYearFromCSL(cslData);
  }

  async function autoImportMissing(libraryID) {
    const collectionID = getSelectedCollectionID();

    // Track already-imported refs to propagate to duplicates
    const importedMap = new Map();

    for (const mr of matchResults) {
      for (const im of mr.itemMatches) {
        if (im.matchType !== "none") continue;

        const key = _refKey(im.cslData);

        // Check if we already imported this ref in a previous iteration
        if (importedMap.has(key)) {
          const prev = importedMap.get(key);
          im.match = prev.match;
          im.matchType = prev.matchType;
          im.confidence = prev.confidence;
          continue;
        }

        // Broader duplicate check before importing (PMID, ISBN, title+author)
        const existing = await plugin.Matcher.findExisting(im.cslData, libraryID);
        if (existing) {
          im.match = existing;
          im.matchType = "duplicate_found";
          im.confidence = 90;
          importedMap.set(key, im);
          continue;
        }

        let imported = null;

        // Try DOI import first (richer metadata from CrossRef)
        if (im.cslData.DOI) {
          imported = await plugin.Matcher.importByDOI(im.cslData.DOI, libraryID, collectionID);
        }

        // Fallback: create directly from Paperpile data
        if (!imported && im.cslData.title) {
          imported = await plugin.Matcher.createFromCSLData(im.cslData, libraryID, collectionID);
        }

        if (imported) {
          im.match = imported;
          im.matchType = "imported";
          im.confidence = 100;
          importedMap.set(key, im);
        } else {
          Zotero.debug("PP2Zotero: Failed to import/create: title=" + (im.cslData.title || "NONE") + ", DOI=" + (im.cslData.DOI || "NONE") + ", type=" + (im.cslData.type || "NONE"));
        }
      }
    }
  }

  async function addMatchedToCollection(libraryID) {
    const collectionID = getSelectedCollectionID();
    if (!collectionID) return; // "My Library" selected, nothing to do

    const added = new Set();
    for (const mr of matchResults) {
      for (const im of mr.itemMatches) {
        if (!im.match || !im.match.id) continue;
        if (added.has(im.match.id)) continue;
        added.add(im.match.id);

        // Check if item is already in this collection
        const collections = im.match.getCollections ? im.match.getCollections() : [];
        if (collections.includes(collectionID)) continue;

        try {
          im.match.addToCollection(collectionID);
          await im.match.saveTx();
          Zotero.debug("PP2Zotero: Added existing item " + im.match.id + " to collection " + collectionID);
        } catch (e) {
          Zotero.debug("PP2Zotero: Failed to add item to collection: " + e.message);
        }
      }
    }
  }

  function deduplicateRefs() {
    const seen = new Map();
    uniqueRefs = [];

    for (const mr of matchResults) {
      for (const im of mr.itemMatches) {
        const key = im.cslData.DOI
          ? "doi:" + im.cslData.DOI.toLowerCase()
          : "title:" + (im.cslData.title || "").toLowerCase() + "|" + plugin.Matcher._extractYearFromCSL(im.cslData);

        if (!seen.has(key)) {
          seen.set(key, im);
          uniqueRefs.push(im);
        }
      }
    }
  }

  let currentSort = { col: null, asc: true };

  function sortRefs(col) {
    if (currentSort.col === col) {
      currentSort.asc = !currentSort.asc;
    } else {
      currentSort.col = col;
      currentSort.asc = true;
    }

    const statusOrder = { "none": 0, "imported": 1, "duplicate_found": 2, "title_only": 3, "title_year": 4, "doi": 5, "skipped": 6 };

    uniqueRefs.sort((a, b) => {
      let va, vb;
      if (col === "title") {
        va = (a.cslData.title || "").toLowerCase();
        vb = (b.cslData.title || "").toLowerCase();
      } else if (col === "authors") {
        va = (a.cslData.author && a.cslData.author[0] ? a.cslData.author[0].family : "").toLowerCase();
        vb = (b.cslData.author && b.cslData.author[0] ? b.cslData.author[0].family : "").toLowerCase();
      } else if (col === "year") {
        va = parseInt(plugin.Matcher._extractYearFromCSL(a.cslData)) || 0;
        vb = parseInt(plugin.Matcher._extractYearFromCSL(b.cslData)) || 0;
      } else if (col === "doi") {
        va = (a.cslData.DOI || "").toLowerCase();
        vb = (b.cslData.DOI || "").toLowerCase();
      } else if (col === "match") {
        va = statusOrder[a.matchType] !== undefined ? statusOrder[a.matchType] : 5;
        vb = statusOrder[b.matchType] !== undefined ? statusOrder[b.matchType] : 5;
      } else {
        return 0;
      }

      if (va < vb) return currentSort.asc ? -1 : 1;
      if (va > vb) return currentSort.asc ? 1 : -1;
      return 0;
    });

    renderResults();
  }

  function renderResults() {
    const tbody = document.getElementById("results-tbody");
    tbody.innerHTML = "";

    let totalCitations = matchResults.length;
    let matched = 0, missing = 0, imported = 0;

    uniqueRefs.forEach(ref => {
      if (ref.matchType === "doi" || ref.matchType === "title_year" || ref.matchType === "title_only" || ref.matchType === "duplicate_found") matched++;
      else if (ref.matchType === "imported") imported++;
      else missing++;
    });

    document.getElementById("stat-unique").textContent = uniqueRefs.length;
    document.getElementById("stat-citations").textContent = totalCitations;
    document.getElementById("stat-matched").textContent = matched;
    document.getElementById("stat-missing").textContent = missing;
    document.getElementById("stat-imported").textContent = imported;

    // Make headers sortable
    document.querySelectorAll("th[data-sort]").forEach(th => {
      th.style.cursor = "pointer";
      const col = th.dataset.sort;
      const arrow = currentSort.col === col ? (currentSort.asc ? " \u25B2" : " \u25BC") : "";
      // Remove old arrow
      th.textContent = th.textContent.replace(/\s*[\u25B2\u25BC]$/, "") + arrow;
      th.onclick = () => sortRefs(col);
    });

    let idx = 0;

    uniqueRefs.forEach((ref, i) => {
      idx++;
      const tr = document.createElement("tr");
      if (ref.matchType === "none") tr.className = "row-missing";
      if (ref.matchType === "imported") tr.className = "row-imported";

      const dotClass = ref.matchType === "none" ? "red"
        : ref.matchType === "imported" ? "blue" : "green";

      const authors = formatAuthors(ref.cslData.author);
      const year = plugin.Matcher._extractYearFromCSL(ref.cslData);
      const doi = ref.cslData.DOI || "";

      // Build cells using DOM methods to avoid XHTML innerHTML parsing failures
      const addCell = (cls, content) => {
        const td = document.createElement("td");
        td.className = cls;
        if (typeof content === "string") td.textContent = content;
        else if (content) td.appendChild(content);
        return tr.appendChild(td);
      };

      // Status dot
      const dot = document.createElement("span");
      dot.className = "status-dot " + dotClass;
      addCell("col-status", dot);

      // Index
      addCell("col-idx", String(idx));

      // Title
      addCell("col-title", ref.cslData.title || "Untitled");

      // Authors
      addCell("col-authors", authors);

      // Year
      addCell("col-year", year);

      // DOI
      const doiTd = addCell("col-doi", null);
      if (doi) {
        doiTd.textContent = doi.substring(0, 14) + "\u2026";
      } else {
        const doiSpan = document.createElement("span");
        doiSpan.style.cssText = "color:var(--accent-red);font-size:9px";
        doiSpan.textContent = "missing";
        doiTd.appendChild(doiSpan);
      }

      // Match badge
      const matchTd = addCell("col-match", null);
      const badge = document.createElement("span");
      if (ref.matchType === "doi") { badge.className = "match-badge exact"; badge.textContent = "\u2713 DOI exact"; }
      else if (ref.matchType === "title_year") { badge.className = "match-badge exact"; badge.textContent = "\u2713 Title+Year"; }
      else if (ref.matchType === "title_only") { badge.className = "match-badge exact"; badge.textContent = "\u2713 Title match"; }
      else if (ref.matchType === "duplicate_found") { badge.className = "match-badge exact"; badge.textContent = "\u2713 Found existing"; }
      else if (ref.matchType === "imported") { badge.className = "match-badge imported"; badge.textContent = "\u2B07 Imported DOI"; }
      else { badge.className = "match-badge none"; badge.textContent = "\u2717 Missing"; }
      matchTd.appendChild(badge);

      // Action
      const actionTd = addCell("col-action", null);
      if (ref.matchType === "none") {
        const wrap = document.createElement("div");
        wrap.style.cssText = "display:flex;gap:3px;flex-wrap:wrap;";
        const makeBtn = (label, action, cls) => {
          const btn = document.createElement("button");
          btn.className = "action-btn" + (cls ? " " + cls : "");
          btn.dataset.idx = i;
          btn.dataset.action = action;
          btn.textContent = label;
          return wrap.appendChild(btn);
        };
        makeBtn("\u270E Edit", "edit");
        makeBtn(doi ? "\u2B07 Import" : "\u2B07 Create", "import", "import-doi");
        makeBtn("Skip", "skip");
        actionTd.appendChild(wrap);
      } else {
        const dash = document.createElement("span");
        dash.style.cssText = "color:var(--text-muted);font-size:11px";
        dash.textContent = "\u2014";
        actionTd.appendChild(dash);
      }

      tbody.appendChild(tr);
    });

    updateConvertCount();

    tbody.querySelectorAll(".action-btn").forEach(btn => {
      btn.addEventListener("click", handleAction);
    });

    document.getElementById("btn-cancel-2").addEventListener("click", () => window.close());
    document.getElementById("btn-convert").addEventListener("click", () => startConversion());
  }

  function handleAction(e) {
    const idx = parseInt(e.target.dataset.idx);
    const action = e.target.dataset.action;
    const ref = uniqueRefs[idx];

    if (action === "edit") {
      toggleEditRow(idx, e.target.closest("tr"));
    } else if (action === "skip") {
      ref.action = "skip";
      ref.matchType = "skipped";
      const row = e.target.closest("tr");
      const oldBadge = row.querySelector(".match-badge");
      const newBadge = document.createElement("span");
      newBadge.className = "match-badge none";
      newBadge.style.opacity = "0.5";
      newBadge.textContent = "Skipped";
      oldBadge.parentNode.replaceChild(newBadge, oldBadge);
      const actionTd = e.target.closest("td");
      actionTd.textContent = "";
      const skippedSpan = document.createElement("span");
      skippedSpan.style.cssText = "color:var(--text-muted);font-size:11px";
      skippedSpan.textContent = "skipped";
      actionTd.appendChild(skippedSpan);
    } else if (action === "import") {
      importSingleDOI(idx, e.target);
    }

    updateConvertCount();
  }

  function toggleEditRow(idx, row) {
    const ref = uniqueRefs[idx];
    const existingEdit = row.nextElementSibling;
    if (existingEdit && existingEdit.classList.contains("edit-row")) {
      existingEdit.remove();
      return;
    }

    const typeOptions = [
      ["journalArticle", "Journal Article"],
      ["book", "Book"],
      ["bookSection", "Book Section"],
      ["conferencePaper", "Conference Paper"],
      ["thesis", "Thesis"],
      ["report", "Report"],
      ["webpage", "Web Page"],
      ["document", "Document"],
      ["patent", "Patent"]
    ];
    const csl = ref.cslData;
    const currentType = {
      "article-journal": "journalArticle", "article": "journalArticle",
      "book": "book", "chapter": "bookSection", "paper-conference": "conferencePaper",
      "thesis": "thesis", "report": "report", "webpage": "webpage", "patent": "patent"
    }[csl.type] || "document";

    const authorsStr = (csl.author || []).map(a => a.isInstitution ? (a.family || "") : ((a.given || "") + " " + (a.family || "")).trim()).join("; ");
    const year = plugin.Matcher._extractYearFromCSL(csl);

    const editTr = document.createElement("tr");
    editTr.className = "edit-row";
    const editTd = document.createElement("td");
    editTd.setAttribute("colspan", "8");
    editTd.className = "edit-cell";

    const form = document.createElement("div");
    form.className = "edit-form";
    const fields = document.createElement("div");
    fields.className = "edit-row-fields";

    const makeField = (labelText, cls, type, value, attrs) => {
      const label = document.createElement("label");
      label.textContent = labelText + " ";
      const input = document.createElement(type === "select" ? "select" : "input");
      input.className = cls;
      if (type === "select") {
        typeOptions.forEach(([v, l]) => {
          const opt = document.createElement("option");
          opt.value = v;
          opt.textContent = l;
          if (v === currentType) opt.selected = true;
          input.appendChild(opt);
        });
      } else {
        input.value = value || "";
        if (attrs) Object.entries(attrs).forEach(([k, v]) => { input.style[k] !== undefined ? input.setAttribute(k, v) : input.setAttribute(k, v); });
      }
      label.appendChild(input);
      fields.appendChild(label);
    };

    makeField("Type", "edit-type", "select");
    makeField("Title", "edit-title", "text", csl.title || "");
    makeField("Authors", "edit-authors", "text", authorsStr, { placeholder: "First Last; First Last" });
    makeField("Year", "edit-year", "text", year, { style: "width:60px" });
    makeField("DOI", "edit-doi", "text", csl.DOI || "");
    makeField("Publisher", "edit-publisher", "text", csl.publisher || "");

    form.appendChild(fields);

    const actions = document.createElement("div");
    actions.className = "edit-actions";
    const saveBtn = document.createElement("button");
    saveBtn.className = "action-btn import-doi";
    saveBtn.dataset.idx = idx;
    saveBtn.dataset.action = "save-edit";
    saveBtn.textContent = "\u2714 Save & Create";
    actions.appendChild(saveBtn);
    const cancelBtn = document.createElement("button");
    cancelBtn.className = "action-btn";
    cancelBtn.dataset.idx = idx;
    cancelBtn.dataset.action = "cancel-edit";
    cancelBtn.textContent = "\u2716 Cancel";
    actions.appendChild(cancelBtn);
    form.appendChild(actions);

    editTd.appendChild(form);
    editTr.appendChild(editTd);
    row.after(editTr);

    editTr.querySelector('[data-action="save-edit"]').addEventListener("click", async (e) => {
      const cell = editTr.querySelector(".edit-cell");
      const newTitle = cell.querySelector(".edit-title").value.trim();
      const newAuthorsStr = cell.querySelector(".edit-authors").value.trim();
      const newYear = cell.querySelector(".edit-year").value.trim();
      const newDOI = cell.querySelector(".edit-doi").value.trim();
      const newPublisher = cell.querySelector(".edit-publisher").value.trim();
      const newType = cell.querySelector(".edit-type").value;

      // Update cslData
      csl.title = newTitle;
      csl.DOI = newDOI;
      csl.publisher = newPublisher;

      // Parse authors
      if (newAuthorsStr) {
        csl.author = newAuthorsStr.split(";").map(a => {
          a = a.trim();
          const parts = a.split(/\s+/);
          const family = parts.pop() || "";
          const given = parts.join(" ");
          return { family, given };
        });
      } else {
        csl.author = [];
      }

      // Update year
      if (newYear) {
        csl.issued = { "date-parts": [[parseInt(newYear) || newYear]] };
      }

      // Update type
      const reverseTypeMap = {
        "journalArticle": "article-journal", "book": "book", "bookSection": "chapter",
        "conferencePaper": "paper-conference", "thesis": "thesis", "report": "report",
        "webpage": "webpage", "document": "article", "patent": "patent"
      };
      csl.type = reverseTypeMap[newType] || "article";

      editTr.remove();
      // Now import
      importSingleDOI(idx, e.target);
    });

    editTr.querySelector('[data-action="cancel-edit"]').addEventListener("click", () => {
      editTr.remove();
    });
  }

  function escapeAttr(str) {
    return str.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  async function importSingleDOI(idx, btn) {
    const ref = uniqueRefs[idx];
    btn.textContent = "Importing...";
    btn.disabled = true;

    const libraryID = Zotero.Libraries.userLibraryID;
    const collectionID = getSelectedCollectionID();

    let imported = null;
    if (ref.cslData.DOI) {
      imported = await plugin.Matcher.importByDOI(ref.cslData.DOI, libraryID, collectionID);
    }
    if (!imported && ref.cslData.title) {
      imported = await plugin.Matcher.createFromCSLData(ref.cslData, libraryID, collectionID);
    }

    if (imported) {
      ref.match = imported;
      ref.matchType = "imported";
      ref.confidence = 100;

      const refKey = ref.cslData.DOI ? ref.cslData.DOI.toLowerCase() : (ref.cslData.title || "").toLowerCase();
      for (const mr of matchResults) {
        for (const im of mr.itemMatches) {
          const imKey = im.cslData.DOI ? im.cslData.DOI.toLowerCase() : (im.cslData.title || "").toLowerCase();
          if (imKey === refKey) {
            im.match = imported;
            im.matchType = "imported";
            im.confidence = 100;
          }
        }
      }

      renderResults();
    } else {
      btn.textContent = "Failed";
      btn.disabled = true;
    }
  }

  function updateConvertCount() {
    const resolved = uniqueRefs.filter(r => r.matchType !== "none").length;
    document.getElementById("convert-count").textContent = resolved + "/" + uniqueRefs.length;
  }

  async function startConversion() {
    showScreen("screen-progress");
    const startTime = Date.now();

    try {
      const overwrite = document.getElementById("opt-overwrite").checked;

      // Determine output path
      let outputPath;
      const isODT = _fieldMode === "referencemarks";
      if (overwrite && !isODT) {
        outputPath = selectedFilePath;
      } else if (isODT) {
        outputPath = selectedFilePath.replace(/\.docx$/i, "_zotero.odt");
      } else {
        // filename.docx -> filename_zotero.docx
        outputPath = selectedFilePath.replace(/\.docx$/i, "_zotero.docx");
      }

      let backupPath = null;
      if (document.getElementById("opt-backup").checked && overwrite) {
        updateProgress(0, "Creating backup...");
        backupPath = await plugin.Converter.createBackup(selectedFilePath);
      }

      updateProgress(5, "Reading document...");
      const fileData = await plugin.readFile(selectedFilePath);

      for (const mr of matchResults) {
        for (const im of mr.itemMatches) {
          if (im.matchType === "none" || im.matchType === "skipped") {
            im.action = "skip";
          }
        }
      }

      const progressMsg = _fieldMode === "referencemarks" ? "Converting to ODF reference marks..."
        : _fieldMode === "bookmarks" ? "Converting to bookmarks..." : "Converting field codes...";
      updateProgress(10, progressMsg);
      const hasSkipped = matchResults.some(mr =>
        mr.itemMatches.every(m => m.matchType === "none" || m.matchType === "skipped" || m.action === "skip")
      );
      const newData = await plugin.Converter.convert(fileData, matchResults, { addComments: hasSkipped, fieldMode: _fieldMode });

      updateProgress(95, "Saving document...");
      await plugin.Converter.saveFile(outputPath, newData);

      const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
      convertStats = buildStats(backupPath, elapsed);
      convertStats.outputPath = outputPath;
      convertStats.overwrite = overwrite;

      updateProgress(100, "Complete!");
      setTimeout(() => {
        showScreen("screen-complete");
        renderComplete();
      }, 500);

    } catch (e) {
      Zotero.alert(window, "PP2Zotero Error", "Conversion failed: " + e.message);
      Zotero.debug("PP2Zotero conversion error: " + e.stack);
      showScreen("screen-results");
    }
  }

  function updateProgress(pct, detail) {
    document.getElementById("progress-bar").style.width = pct + "%";
    if (detail) document.getElementById("progress-detail").textContent = detail;
  }

  function buildStats(backupPath, elapsed) {
    let citationsConverted = 0, citationsTotal = matchResults.length;
    let matchedDOI = 0, matchedTitle = 0, importedDOI = 0, duplicateFound = 0, skipped = 0;

    const counted = new Set();
    for (const mr of matchResults) {
      let allSkip = true;
      for (const im of mr.itemMatches) {
        const key = im.cslData.DOI || im.cslData.title;
        if (!counted.has(key)) {
          counted.add(key);
          if (im.matchType === "doi") matchedDOI++;
          else if (im.matchType === "title_year" || im.matchType === "title_only") matchedTitle++;
          else if (im.matchType === "duplicate_found") duplicateFound++;
          else if (im.matchType === "imported") importedDOI++;
          else skipped++;
        }
        if (im.match) allSkip = false;
      }
      if (!allSkip) citationsConverted++;
    }

    return {
      citationsConverted, citationsTotal,
      uniqueConverted: uniqueRefs.filter(r => r.match).length,
      uniqueTotal: uniqueRefs.length,
      matchedDOI, matchedTitle, duplicateFound, importedDOI, skipped,
      backupPath, elapsed
    };
  }

  function renderComplete() {
    const s = convertStats;
    const summary = document.getElementById("complete-summary");
    summary.textContent = "";

    const addRow = (label, value, cls) => {
      const row = document.createElement("div");
      row.className = "summary-row";
      const lbl = document.createElement("span");
      lbl.className = "label";
      lbl.textContent = label;
      const val = document.createElement("span");
      val.className = "value" + (cls ? " " + cls : "");
      val.textContent = value;
      row.appendChild(lbl);
      row.appendChild(val);
      summary.appendChild(row);
    };
    const addDivider = () => {
      const d = document.createElement("div");
      d.className = "summary-divider";
      summary.appendChild(d);
    };

    if (s.isBatch) {
      addRow("Files processed", s.batchResults.length);
      const okCount = s.batchResults.filter(r => r.status === "ok").length;
      const errCount = s.batchResults.filter(r => r.status === "error").length;
      const skipCount = s.batchResults.filter(r => r.status === "skipped").length;
      addRow("Successful", okCount, "green");
      if (errCount > 0) addRow("Errors", errCount, "red");
      if (skipCount > 0) addRow("Skipped (no citations)", skipCount);
      addDivider();
      addRow("Total citations", s.citationsTotal);
      addRow("Citations converted", s.citationsConverted, "green");
      if (s.skipped > 0) addRow("Citations skipped", s.skipped, "red");
      addDivider();
      // Per-file details
      for (const br of s.batchResults) {
        if (br.status === "ok") {
          addRow(br.file, br.converted + "/" + br.citations + " converted", br.skipped > 0 ? "yellow" : "green");
        } else if (br.status === "error") {
          addRow(br.file, "Error: " + br.reason, "red");
        } else {
          addRow(br.file, br.reason);
        }
      }
      addDivider();
      addRow("Total time", s.elapsed + "s");
      // Hide Open button in batch mode
      document.getElementById("btn-open-file").style.display = "none";
    } else {
      addRow("Citations converted", s.citationsConverted + " / " + s.citationsTotal, "green");
      addRow("Unique references converted", s.uniqueConverted + " / " + s.uniqueTotal, "green");
      addDivider();
      addRow("Exact match (DOI)", s.matchedDOI);
      addRow("Exact match (title)", s.matchedTitle);
      if (s.duplicateFound > 0) addRow("Found existing (broad match)", s.duplicateFound, "green");
      addRow("Imported from DOI (CrossRef)", s.importedDOI, "blue");
      addDivider();
      addRow("Citations skipped", s.skipped, "red");
      addDivider();
      const savedRow = document.createElement("div");
      savedRow.className = "summary-row";
      const savedLbl = document.createElement("span");
      savedLbl.className = "label";
      savedLbl.textContent = "Saved to";
      const savedVal = document.createElement("span");
      savedVal.className = "value";
      savedVal.style.fontSize = "10px";
      savedVal.textContent = s.outputPath.split("/").pop();
      savedRow.appendChild(savedLbl);
      savedRow.appendChild(savedVal);
      summary.appendChild(savedRow);
      if (s.backupPath) {
        const bkRow = document.createElement("div");
        bkRow.className = "summary-row";
        const bkLbl = document.createElement("span");
        bkLbl.className = "label";
        bkLbl.textContent = "Backup saved";
        const bkVal = document.createElement("span");
        bkVal.className = "value";
        bkVal.style.fontSize = "10px";
        bkVal.textContent = s.backupPath.split("/").pop();
        bkRow.appendChild(bkLbl);
        bkRow.appendChild(bkVal);
        summary.appendChild(bkRow);
      }
      addRow("Total time", s.elapsed + "s");
    }

    if (s.skipped > 0) {
      document.getElementById("warning-box").style.display = "flex";
      const warnText = document.getElementById("warning-text");
      warnText.textContent = "";
      const strong = document.createElement("strong");
      strong.textContent = s.skipped + " unconverted citation(s)";
      warnText.appendChild(strong);
      warnText.appendChild(document.createTextNode(" were left as Paperpile fields in the document. Word comments have been added at each unconverted citation to help you find them."));
    }

    document.getElementById("btn-close").addEventListener("click", () => window.close());
    document.getElementById("btn-copy-report").addEventListener("click", () => {
      const text = summary.innerText;
      Components.classes["@mozilla.org/widget/clipboardhelper;1"]
        .getService(Components.interfaces.nsIClipboardHelper)
        .copyString(text);
    });

    // Update button label based on output format
    const openBtn = document.getElementById("btn-open-file");
    if (_fieldMode === "referencemarks") {
      openBtn.textContent = "\uD83D\uDCC2 Open in LibreOffice";
    } else {
      openBtn.textContent = "\uD83D\uDCC2 Open in Word";
    }

    // Open file in default app
    openBtn.addEventListener("click", () => {
      try {
        const file = Components.classes["@mozilla.org/file/local;1"]
          .createInstance(Components.interfaces.nsIFile);
        file.initWithPath(s.outputPath);
        file.launch();
      } catch (e) {
        Zotero.debug("PP2Zotero: Failed to open file: " + e.message);
      }
    });

    // Export detailed report
    document.getElementById("btn-export-report").addEventListener("click", async () => {
      const fp = Components.classes["@mozilla.org/filepicker;1"]
        .createInstance(Components.interfaces.nsIFilePicker);
      fp.init(window, "Save conversion report", fp.modeSave);
      fp.appendFilter("CSV files", "*.csv");
      fp.appendFilter("Text files", "*.txt");
      fp.defaultString = "pp2zotero_report.csv";

      const result = await new Promise(resolve => fp.open(resolve));
      if (result === fp.returnOK || result === fp.returnReplace) {
        const report = buildDetailedReport(fp.file.path.endsWith(".txt") ? "txt" : "csv");
        await IOUtils.writeUTF8(fp.file.path, report);
      }
    });
  }

  function buildDetailedReport(format) {
    const s = convertStats;
    if (format === "txt") {
      let report = "PP2Zotero Conversion Report\n";
      report += "==========================\n\n";
      report += "Date: " + new Date().toLocaleString() + "\n";
      report += "Time: " + s.elapsed + "s\n\n";

      if (s.isBatch) {
        report += "Batch mode: " + s.batchResults.length + " files\n\n";
        for (const br of s.batchResults) {
          report += "  " + br.file + ": ";
          if (br.status === "ok") report += br.converted + "/" + br.citations + " converted\n";
          else if (br.status === "error") report += "ERROR - " + br.reason + "\n";
          else report += br.reason + "\n";
        }
        report += "\nTotal citations: " + s.citationsConverted + "/" + s.citationsTotal + " converted\n";
        report += "Skipped: " + s.skipped + "\n";
      } else {
        report += "Source: " + selectedFilePath.split("/").pop() + "\n";
        report += "Output: " + s.outputPath.split("/").pop() + "\n\n";
        report += "Citations: " + s.citationsConverted + "/" + s.citationsTotal + " converted\n";
        report += "Unique refs: " + s.uniqueConverted + "/" + s.uniqueTotal + " resolved\n";
        report += "  DOI match: " + s.matchedDOI + "\n";
        report += "  Title match: " + s.matchedTitle + "\n";
        if (s.duplicateFound > 0) report += "  Found existing: " + s.duplicateFound + "\n";
        report += "  Imported: " + s.importedDOI + "\n";
        report += "  Skipped: " + s.skipped + "\n\n";
        report += "Detailed reference list:\n";
        report += "-".repeat(60) + "\n";
        uniqueRefs.forEach((ref, i) => {
          const authors = formatAuthors(ref.cslData.author);
          const year = plugin.Matcher._extractYearFromCSL(ref.cslData);
          report += (i + 1) + ". " + (ref.cslData.title || "Untitled") + "\n";
          report += "   Authors: " + (authors || "N/A") + "\n";
          report += "   Year: " + (year || "N/A") + "\n";
          report += "   DOI: " + (ref.cslData.DOI || "N/A") + "\n";
          report += "   Status: " + ref.matchType + "\n\n";
        });
      }
      return report;
    }

    // CSV format
    let csv = "\uFEFF"; // BOM for Excel UTF-8
    if (s.isBatch) {
      csv += "File,Status,Citations,Converted,Skipped,Error\n";
      for (const br of s.batchResults) {
        const file = (br.file || "").replace(/"/g, '""');
        csv += '"' + file + '",' + br.status + ',' + (br.citations || 0) + ',' + (br.converted || 0) + ',' + (br.skipped || 0) + ',"' + (br.reason || "").replace(/"/g, '""') + '"\n';
      }
    } else {
      csv += "Index,Title,Authors,Year,DOI,Type,Match Status\n";
      uniqueRefs.forEach((ref, i) => {
        const authors = formatAuthors(ref.cslData.author);
        const year = plugin.Matcher._extractYearFromCSL(ref.cslData);
        const title = (ref.cslData.title || "").replace(/"/g, '""');
        const authStr = (authors || "").replace(/"/g, '""');
        const doi = ref.cslData.DOI || "";
        const type = ref.cslData.type || "";
        csv += (i + 1) + ',"' + title + '","' + authStr + '",' + year + ',' + doi + ',' + type + ',' + ref.matchType + "\n";
      });
    }
    return csv;
  }

  function formatAuthors(authors) {
    if (!authors || !authors.length) return "";
    const name = a => a.family || a.given || "";
    if (authors.length === 1) return name(authors[0]);
    if (authors.length === 2) return name(authors[0]) + ", " + name(authors[1]);
    return name(authors[0]) + " et al.";
  }

  function escapeHtml(str) {
    const div = document.createElement("div");
    div.textContent = str;
    return div.innerHTML;
  }

  window.addEventListener("load", () => {
    if (!io) {
      document.body.style.background = "#fff";
      showScreen("screen-file");
      return;
    }
    initScreen1();
  });

})();
