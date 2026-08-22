/* ODF Writer Module — converts OOXML (.docx) content to ODF (.odt) with Zotero reference marks */

if (typeof PP2Zotero === "undefined") var PP2Zotero = {};

PP2Zotero.ODFWriter = {
  W_NS: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",

  _autoStyles: {},
  _autoStyleCounter: 0,
  _paraAutoStyles: {},
  _paraStyleCounter: 0,
  _usedFonts: new Set(),
  _footnoteCounter: 0,
  _endnoteCounter: 0,

  async convertToODT(docxData, matchResults, options) {
    // === 2-STEP APPROACH ===
    // Step 1: Create .docx with placeholders instead of Paperpile field codes
    // Step 2: Use LibreOffice CLI to convert .docx → .odt (perfect formatting)
    // Step 3: Post-process .odt to replace placeholders with Zotero reference marks

    if (docxData && docxData.byteLength !== undefined && !(docxData instanceof Uint8Array)) {
      docxData = new Uint8Array(docxData);
    }
    const JSZipRef = typeof JSZip !== "undefined" ? JSZip : (await import("./lib/jszip.min.js")).default;

    // --- Step 1: Replace Paperpile field codes with placeholders in .docx ---
    const citationMap = {}; // placeholderID → citation reference mark name
    const placeholderDocx = await this._createPlaceholderDocx(
      JSZipRef, docxData, matchResults, citationMap, options
    );

    // Save temp .docx — use /tmp directly (PathUtils.tempDir may differ)
    const tempDir = "/tmp";
    const tempName = "pp2z_temp_" + Date.now();
    const tempDocx = tempDir + "/" + tempName + ".docx";
    const tempOdt = tempDir + "/" + tempName + ".odt";
    await IOUtils.write(tempDocx, placeholderDocx);
    Zotero.log("PP2Zotero: Wrote temp docx: " + tempDocx);

    // --- Step 2: Convert .docx → .odt using LibreOffice ---
    Zotero.log("PP2Zotero: Converting via LibreOffice CLI...");
    await this._convertWithLibreOffice(tempDocx, tempDir);

    // Find the .odt — LibreOffice may use slightly different naming
    let odtPath = tempOdt;
    if (!await IOUtils.exists(odtPath)) {
      // List /tmp for any pp2z*.odt files as fallback
      // Read stderr log for diagnostics
      let errLog = "";
      try { errLog = new TextDecoder().decode(await IOUtils.read("/tmp/pp2z_lo_err.log")); } catch (e) {}
      Zotero.log("PP2Zotero: LO stderr: " + errLog);
      throw new Error("LibreOffice .odt not found at " + odtPath + (errLog ? " | stderr: " + errLog : ""));
    }
    Zotero.log("PP2Zotero: Reading .odt from " + odtPath);
    const odtData = await IOUtils.read(odtPath);

    // --- Step 3: Replace placeholders with Zotero reference marks ---
    const finalOdt = await this._insertReferenceMarks(JSZipRef, odtData, citationMap);

    // Cleanup temp files
    try { await IOUtils.remove(tempDocx); } catch (e) {}
    try { await IOUtils.remove(tempOdt); } catch (e) {}

    return finalOdt;
  },

  async _createPlaceholderDocx(JSZipRef, docxData, matchResults, citationMap, options) {
    const zip = await JSZipRef.loadAsync(docxData);
    const matchMap = new Map();
    for (const mr of matchResults) {
      matchMap.set(mr.citation.paperpileId, mr);
    }

    const xmlFiles = ["word/document.xml", "word/footnotes.xml", "word/endnotes.xml"];
    let placeholderIdx = 0;

    for (const xmlPath of xmlFiles) {
      const file = zip.file(xmlPath);
      if (!file) continue;
      let xmlText = await file.async("string");

      // Find all Paperpile field code blocks and replace with placeholders
      const blocks = PP2Zotero.Converter._findPaperpileFieldBlocks(xmlText);

      // Process in reverse to preserve positions
      for (let i = blocks.length - 1; i >= 0; i--) {
        const block = blocks[i];

        if (block.isBibliography) {
          const pid = "PP2Z_BIBL";
          const bibl = { uncited: [], omitted: [], custom: [] };
          const rnd = this._makeRnd();
          citationMap[pid] = "ZOTERO_BIBL " + JSON.stringify(bibl) + " CSL_BIBLIOGRAPHY" + rnd;

          // Replace field code with placeholder text + display content
          const placeholder = '<w:r><w:rPr><w:vanish/></w:rPr><w:t>' + pid + '</w:t></w:r>';
          xmlText = xmlText.substring(0, block.startPos) + placeholder + block.displayContent + xmlText.substring(block.endPos);
          continue;
        }

        if (!block.isCitation) continue;

        const clusterMatch = block.instrText.match(/&lt;clusterId&gt;(.*?)&lt;\/clusterId&gt;/)
          || block.instrText.match(/<clusterId>(.*?)<\/clusterId>/);
        if (!clusterMatch) continue;

        const ppId = clusterMatch[1];
        const matchResult = matchMap.get(ppId);
        if (!matchResult) continue;

        const allSkipped = matchResult.itemMatches.every(m =>
          m.matchType === "none" || m.matchType === "skipped" || m.action === "skip"
        );
        if (allSkipped) continue;

        const zc = PP2Zotero.Converter._buildZoteroCitation(matchResult);
        if (!zc) continue;

        // Strip itemData for shorter reference mark names
        const zcLight = JSON.parse(JSON.stringify(zc));
        delete zcLight.schema;
        for (const ci of (zcLight.citationItems || [])) {
          delete ci.itemData;
        }

        placeholderIdx++;
        const pid = "PP2Z_" + String(placeholderIdx).padStart(4, "0");
        const rnd = this._makeRnd();
        citationMap[pid] = "ZOTERO_ITEM CSL_CITATION " + JSON.stringify(zcLight) + rnd;

        // Replace field code with: hidden placeholder + display text
        const placeholder = '<w:r><w:rPr><w:vanish/></w:rPr><w:t>' + pid + '</w:t></w:r>';
        const endPlaceholder = '<w:r><w:rPr><w:vanish/></w:rPr><w:t>/' + pid + '</w:t></w:r>';
        xmlText = xmlText.substring(0, block.startPos)
          + placeholder + block.displayContent + endPlaceholder
          + xmlText.substring(block.endPos);
      }

      zip.file(xmlPath, xmlText);
    }

    return await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  },

  async _convertWithLibreOffice(inputPath, outDir) {
    // nsIProcess can't run shell scripts directly — use /bin/bash
    const bash = Components.classes["@mozilla.org/file/local;1"]
      .createInstance(Components.interfaces.nsIFile);
    bash.initWithPath("/bin/bash");

    const process = Components.classes["@mozilla.org/process/util;1"]
      .createInstance(Components.interfaces.nsIProcess);
    process.init(bash);

    // Find soffice absolute path
    let sofficePath = "/usr/bin/soffice";
    const candidates = ["/usr/bin/soffice", "/usr/bin/libreoffice", "/usr/local/bin/soffice"];
    for (const c of candidates) {
      try { if (await IOUtils.exists(c)) { sofficePath = c; break; } } catch (e) {}
    }

    const profileDir = "/tmp/pp2z_lo_profile_" + Date.now();
    // No single quotes around -env: — bash -c handles the whole string
    const cmd = sofficePath + " --headless"
      + " -env:UserInstallation=file://" + profileDir
      + " --convert-to odt --outdir " + outDir + " " + inputPath
      + " 2>/tmp/pp2z_lo_err.log ; rm -rf " + profileDir;

    Zotero.log("PP2Zotero: Running: " + cmd);
    const args = ["-c", cmd];

    try {
      process.run(true, args, args.length);
      Zotero.log("PP2Zotero: bash exit code: " + process.exitValue);
    } catch (e) {
      throw new Error("LibreOffice conversion failed: " + e.message);
    }
  },

  async _insertReferenceMarks(JSZipRef, odtData, citationMap) {
    const zip = await JSZipRef.loadAsync(odtData);

    // Process content.xml
    const contentFile = zip.file("content.xml");
    if (!contentFile) throw new Error("No content.xml in .odt");
    let content = await contentFile.async("string");

    // Replace placeholders with reference marks
    for (const [pid, refMarkName] of Object.entries(citationMap)) {
      const escapedName = this._escapeXml(refMarkName);

      if (pid === "PP2Z_BIBL") {
        // Bibliography: find placeholder, wrap with reference mark
        const startRe = new RegExp(this._escapeRegex(pid), "g");
        const endRe = new RegExp(this._escapeRegex("/" + pid), "g");
        content = content.replace(startRe, '<text:reference-mark-start text:name="' + escapedName + '"/>');
        content = content.replace(endRe, '<text:reference-mark-end text:name="' + escapedName + '"/>');
        continue;
      }

      // Citation: find start/end placeholders
      // The placeholders might be inside <text:span> elements after conversion
      // Find PP2Z_NNNN and /PP2Z_NNNN in the XML text content
      const startMarker = pid;
      const endMarker = "/" + pid;

      // Replace the placeholder text (might be wrapped in spans)
      content = this._replacePlaceholderInXml(content, startMarker,
        '<text:reference-mark-start text:name="' + escapedName + '"/>');
      content = this._replacePlaceholderInXml(content, endMarker,
        '<text:reference-mark-end text:name="' + escapedName + '"/>');
    }

    zip.file("content.xml", content);

    // Add ZOTERO_PREF as custom property in meta.xml
    const metaFile = zip.file("meta.xml");
    if (metaFile) {
      let meta = await metaFile.async("string");
      const sessionId = PP2Zotero.Converter._generateCitationID();
      const prefXml = '<data data-version="3" zotero-version="7.0.0">'
        + '<session id="' + sessionId + '"/>'
        + '<prefs><pref name="fieldType" value="ReferenceMark"/></prefs>'
        + '</data>';
      const prefProp = '<meta:user-defined meta:name="ZOTERO_PREF" meta:value-type="string">'
        + this._escapeXml(prefXml) + '</meta:user-defined>';

      if (meta.indexOf("ZOTERO_PREF") === -1) {
        meta = meta.replace("</office:meta>", prefProp + "\n</office:meta>");
        zip.file("meta.xml", meta);
      }
    }

    return await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  },

  _replacePlaceholderInXml(xml, placeholder, replacement) {
    // The placeholder text might appear directly or inside XML text nodes
    // It was set as hidden text (w:vanish) in docx, which LibreOffice may preserve or remove
    // Try multiple patterns:

    // Pattern 1: placeholder as plain text content (most common after LO conversion)
    const idx = xml.indexOf(placeholder);
    if (idx !== -1) {
      // Check if it's inside a text element - find the surrounding context
      // Replace just the text, keeping any surrounding XML tags
      return xml.substring(0, idx) + replacement + xml.substring(idx + placeholder.length);
    }
    return xml;
  },

  _makeRnd() {
    return " RND" + PP2Zotero.Converter._generateCitationID()
      + PP2Zotero.Converter._generateCitationID().substring(0, 2);
  },

  /* ---- OOXML parsing helpers ---- */

  async _parseNotes(zip, xmlPath, tagName) {
    const file = zip.file(xmlPath);
    if (!file) return {};
    const text = await file.async("string");
    const doc = new DOMParser().parseFromString(text, "application/xml");
    const notes = {};
    const noteEls = doc.getElementsByTagNameNS(this.W_NS, tagName);
    for (let i = 0; i < noteEls.length; i++) {
      const note = noteEls[i];
      const type = note.getAttribute("w:type");
      if (type === "separator" || type === "continuationSeparator") continue;
      const id = note.getAttribute("w:id");
      if (id) notes[id] = note;
    }
    return notes;
  },

  async _parseDocStyles(zip) {
    const result = { defaultPara: {}, defaultRun: {}, named: {} };
    const file = zip.file("word/styles.xml");
    if (!file) return result;
    const text = await file.async("string");
    const doc = new DOMParser().parseFromString(text, "application/xml");

    // Document defaults (docDefaults → rPrDefault/pPrDefault)
    const docDefaults = doc.getElementsByTagNameNS(this.W_NS, "docDefaults")[0];
    if (docDefaults) {
      const rPrDefault = this._deepChildNS(docDefaults, "rPrDefault");
      if (rPrDefault) {
        const rPr = this._deepChildNS(rPrDefault, "rPr");
        if (rPr) result.defaultRun = this._extractRunProps(rPr);
      }
      const pPrDefault = this._deepChildNS(docDefaults, "pPrDefault");
      if (pPrDefault) {
        const pPr = this._deepChildNS(pPrDefault, "pPr");
        if (pPr) result.defaultPara = this._extractParaProps(pPr);
      }
    }

    // Named styles (Normal, Heading1, etc.)
    const styleEls = doc.getElementsByTagNameNS(this.W_NS, "style");
    for (let i = 0; i < styleEls.length; i++) {
      const styleEl = styleEls[i];
      const styleId = styleEl.getAttribute("w:styleId");
      if (!styleId) continue;
      const entry = {};
      const pPr = this._deepChildNS(styleEl, "pPr");
      if (pPr) entry.para = this._extractParaProps(pPr);
      const rPr = this._deepChildNS(styleEl, "rPr");
      if (rPr) entry.run = this._extractRunProps(rPr);
      result.named[styleId] = entry;
    }

    return result;
  },

  _extractRunProps(rPr) {
    const props = {};
    if (!rPr) return props;
    if (this._deepChildNS(rPr, "b")) props.bold = true;
    if (this._deepChildNS(rPr, "i")) props.italic = true;
    if (this._deepChildNS(rPr, "u")) props.underline = true;
    if (this._deepChildNS(rPr, "strike")) props.strikethrough = true;
    const va = this._deepChildNS(rPr, "vertAlign");
    if (va) {
      const val = va.getAttribute("w:val");
      if (val === "superscript") props.superscript = true;
      if (val === "subscript") props.subscript = true;
    }
    const sz = this._deepChildNS(rPr, "sz");
    if (sz) {
      const halfPt = parseInt(sz.getAttribute("w:val"));
      if (halfPt) props.fontSize = (halfPt / 2) + "pt";
    }
    const rFonts = this._deepChildNS(rPr, "rFonts");
    if (rFonts) {
      const font = rFonts.getAttribute("w:ascii") || rFonts.getAttribute("w:hAnsi") || rFonts.getAttribute("w:cs");
      if (font) {
        props.fontFamily = font;
        this._usedFonts.add(font);
      }
    }
    const color = this._deepChildNS(rPr, "color");
    if (color) {
      const val = color.getAttribute("w:val");
      if (val && val !== "auto") props.color = "#" + val;
    }
    return props;
  },

  _extractPageLayout(bodyEl) {
    const layout = {};
    const sectPr = bodyEl.getElementsByTagNameNS(this.W_NS, "sectPr")[0];
    if (!sectPr) return layout;

    const pgSz = this._deepChildNS(sectPr, "pgSz");
    if (pgSz) {
      const w = pgSz.getAttribute("w:w");
      const h = pgSz.getAttribute("w:h");
      if (w) layout.pageWidth = (parseInt(w) / 567).toFixed(2) + "cm";
      if (h) layout.pageHeight = (parseInt(h) / 567).toFixed(2) + "cm";
    }

    const pgMar = this._deepChildNS(sectPr, "pgMar");
    if (pgMar) {
      const t = pgMar.getAttribute("w:top");
      const r = pgMar.getAttribute("w:right");
      const b = pgMar.getAttribute("w:bottom");
      const l = pgMar.getAttribute("w:left");
      if (t) layout.marginTop = (parseInt(t) / 567).toFixed(2) + "cm";
      if (r) layout.marginRight = (parseInt(r) / 567).toFixed(2) + "cm";
      if (b) layout.marginBottom = (parseInt(b) / 567).toFixed(2) + "cm";
      if (l) layout.marginLeft = (parseInt(l) / 567).toFixed(2) + "cm";
    }

    return layout;
  },

  // Like _firstChildNS but searches all descendants (getElementsByTagNameNS)
  _deepChildNS(parent, localName) {
    const els = parent.getElementsByTagNameNS(this.W_NS, localName);
    return els.length > 0 ? els[0] : null;
  },

  _firstChildNS(parent, localName) {
    for (let i = 0; i < parent.childNodes.length; i++) {
      const n = parent.childNodes[i];
      if (n.nodeType === 1 && n.localName === localName && n.namespaceURI === this.W_NS) return n;
    }
    return null;
  },

  /* ---- Body conversion ---- */

  _convertBody(bodyEl, matchMap, footnoteMap, endnoteMap) {
    let content = "";
    for (let i = 0; i < bodyEl.childNodes.length; i++) {
      const child = bodyEl.childNodes[i];
      if (child.nodeType !== 1) continue;
      if (child.localName === "p") {
        content += this._convertParagraph(child, matchMap, footnoteMap, endnoteMap);
      } else if (child.localName === "tbl") {
        content += this._convertTable(child, matchMap, footnoteMap, endnoteMap);
      }
      // Skip sectPr, sdt, and other elements
    }

    return content;
  },

  _convertParagraph(wP, matchMap, footnoteMap, endnoteMap) {
    let styleName = "Standard";
    let isHeading = false;
    let headingLevel = 0;

    const pPr = this._firstChildNS(wP, "pPr");
    if (pPr) {
      // Check outlineLvl first (most reliable for headings regardless of language)
      const outlineLvl = this._firstChildNS(pPr, "outlineLvl");
      if (outlineLvl) {
        const lvl = parseInt(outlineLvl.getAttribute("w:val"));
        if (!isNaN(lvl) && lvl >= 0 && lvl <= 8) {
          isHeading = true;
          headingLevel = lvl + 1;
          styleName = "Heading_20_" + headingLevel;
        }
      }

      // Fallback: check style name for heading patterns (any language)
      if (!isHeading) {
        const pStyle = this._firstChildNS(pPr, "pStyle");
        if (pStyle) {
          const val = pStyle.getAttribute("w:val") || "";
          const headingMatch = val.match(/^(?:[Hh]eading|Titre|Titlu|[Üü]berschrift|Encabezado|Intestazione)\s*(\d+)$/);
          if (headingMatch) {
            isHeading = true;
            headingLevel = parseInt(headingMatch[1]);
            styleName = "Heading_20_" + headingLevel;
          }
        }
      }

      // Resolve paragraph formatting from style + inline overrides
      if (!isHeading) {
        const pStyleEl = this._firstChildNS(pPr, "pStyle");
        const styleId = pStyleEl ? pStyleEl.getAttribute("w:val") : null;
        const styleDef = styleId ? this._docStyles.named[styleId] : null;

        // Merge paragraph props: doc defaults < named style < inline
        const mergedPara = {};
        Object.assign(mergedPara, this._docStyles.defaultPara);
        if (styleDef && styleDef.para) Object.assign(mergedPara, styleDef.para);
        Object.assign(mergedPara, this._extractParaProps(pPr));

        // Merge text props for this paragraph: doc defaults < named style < pPr/rPr
        const mergedRun = {};
        Object.assign(mergedRun, this._docStyles.defaultRun);
        if (styleDef && styleDef.run) Object.assign(mergedRun, styleDef.run);
        const pRPr = this._firstChildNS(pPr, "rPr");
        if (pRPr) Object.assign(mergedRun, this._extractRunProps(pRPr));

        if (Object.keys(mergedPara).length > 0 || Object.keys(mergedRun).length > 0) {
          styleName = this._getParaAutoStyleName({ para: mergedPara, run: mergedRun });
        }
      }
    }

    const innerContent = this._processChildren(wP, matchMap, footnoteMap, endnoteMap);

    if (isHeading) {
      return '<text:h text:style-name="' + styleName + '" text:outline-level="' + headingLevel + '">'
        + innerContent + '</text:h>\n';
    }
    return '<text:p text:style-name="' + styleName + '">' + innerContent + '</text:p>\n';
  },

  _extractParaProps(pPr) {
    const props = {};

    // Text alignment
    const jc = this._firstChildNS(pPr, "jc");
    if (jc) {
      const val = jc.getAttribute("w:val");
      const map = { left: "start", center: "center", right: "end", both: "justify" };
      if (map[val]) props.textAlign = map[val];
    }

    // Spacing (w:before/w:after in twips = 1/20 pt; w:line in 240ths of a line)
    const spacing = this._firstChildNS(pPr, "spacing");
    if (spacing) {
      const before = spacing.getAttribute("w:before");
      const after = spacing.getAttribute("w:after");
      const line = spacing.getAttribute("w:line");
      const lineRule = spacing.getAttribute("w:lineRule");

      if (before) props.marginTop = (parseInt(before) / 20) + "pt";
      if (after) props.marginBottom = (parseInt(after) / 20) + "pt";

      if (line) {
        if (lineRule === "auto" || !lineRule) {
          props.lineHeight = Math.round(parseInt(line) / 240 * 100) + "%";
        } else {
          props.lineHeight = (parseInt(line) / 20) + "pt";
        }
      }
    }

    // Indentation (values in twips → cm; 567 twips = 1 cm)
    const ind = this._firstChildNS(pPr, "ind");
    if (ind) {
      const left = ind.getAttribute("w:left") || ind.getAttribute("w:start");
      const right = ind.getAttribute("w:right") || ind.getAttribute("w:end");
      const firstLine = ind.getAttribute("w:firstLine");
      const hanging = ind.getAttribute("w:hanging");

      if (left) props.marginLeft = (parseInt(left) / 567).toFixed(2) + "cm";
      if (right) props.marginRight = (parseInt(right) / 567).toFixed(2) + "cm";
      if (firstLine) props.textIndent = (parseInt(firstLine) / 567).toFixed(2) + "cm";
      if (hanging) props.textIndent = "-" + (parseInt(hanging) / 567).toFixed(2) + "cm";
    }

    return props;
  },

  _getParaAutoStyleName(props) {
    const key = JSON.stringify(props);
    if (this._paraAutoStyles[key]) return this._paraAutoStyles[key].name;
    this._paraStyleCounter++;
    const name = "P" + this._paraStyleCounter;
    this._paraAutoStyles[key] = { name, props };
    return name;
  },

  _processChildren(parent, matchMap, footnoteMap, endnoteMap) {
    // Collect element children
    const children = [];
    for (let i = 0; i < parent.childNodes.length; i++) {
      if (parent.childNodes[i].nodeType === 1) children.push(parent.childNodes[i]);
    }

    let result = "";
    let i = 0;
    while (i < children.length) {
      const child = children[i];

      if (child.localName === "r" && child.namespaceURI === this.W_NS) {
        // Check for field code begin
        const fldChar = this._firstChildNS(child, "fldChar");
        if (fldChar && fldChar.getAttribute("w:fldCharType") === "begin") {
          const fc = this._processFieldCode(children, i, matchMap);
          result += fc.odf;
          i = fc.nextIndex;
          continue;
        }

        // Check for footnote reference
        const fnRef = this._firstChildNS(child, "footnoteReference");
        if (fnRef) {
          const fnId = fnRef.getAttribute("w:id");
          if (fnId && footnoteMap[fnId]) {
            result += this._buildNote(footnoteMap[fnId], "footnote", matchMap, footnoteMap, endnoteMap);
          }
          i++;
          continue;
        }

        // Check for endnote reference
        const enRef = this._firstChildNS(child, "endnoteReference");
        if (enRef) {
          const enId = enRef.getAttribute("w:id");
          if (enId && endnoteMap[enId]) {
            result += this._buildNote(endnoteMap[enId], "endnote", matchMap, footnoteMap, endnoteMap);
          }
          i++;
          continue;
        }

        // Regular run
        result += this._convertRun(child);
      } else if (child.localName === "hyperlink") {
        // Process inner runs of hyperlinks
        result += this._processChildren(child, matchMap, footnoteMap, endnoteMap);
      }
      // Skip pPr, bookmarkStart/End, proofErr, etc.

      i++;
    }

    return result;
  },

  /* ---- Field code → Reference mark conversion ---- */

  _processFieldCode(children, beginIdx, matchMap) {
    let instrText = "";
    let displayOdf = "";
    let phase = "instr";
    let nestLevel = 1;
    let j = beginIdx + 1;

    while (j < children.length && nestLevel > 0) {
      const child = children[j];
      if (child.localName !== "r" || child.namespaceURI !== this.W_NS) { j++; continue; }

      const fldChar = this._firstChildNS(child, "fldChar");
      if (fldChar) {
        const type = fldChar.getAttribute("w:fldCharType");
        if (type === "begin") {
          nestLevel++;
        } else if (type === "separate" && nestLevel === 1) {
          phase = "display";
        } else if (type === "end") {
          nestLevel--;
          if (nestLevel === 0) { j++; break; }
        }
        j++;
        continue;
      }

      if (phase === "instr") {
        const instrEl = this._firstChildNS(child, "instrText");
        if (instrEl) instrText += instrEl.textContent;
      } else if (phase === "display") {
        displayOdf += this._convertRun(child);
      }
      j++;
    }

    instrText = instrText.trim();

    // Paperpile citation → Zotero reference mark
    if (instrText.indexOf("paperpile_citation") !== -1) {
      const clusterMatch = instrText.match(/<clusterId>(.*?)<\/clusterId>/);
      if (clusterMatch) {
        const ppId = clusterMatch[1];
        const mr = matchMap.get(ppId);
        if (mr) {
          const allSkipped = mr.itemMatches.every(m =>
            m.matchType === "none" || m.matchType === "skipped" || m.action === "skip"
          );
          if (!allSkipped) {
            const zc = PP2Zotero.Converter._buildZoteroCitation(mr);
            if (zc) {
              // Strip itemData and schema to keep reference mark names short
              const zcLight = JSON.parse(JSON.stringify(zc));
              delete zcLight.schema;
              for (const ci of (zcLight.citationItems || [])) {
                delete ci.itemData;
              }
              // Zotero requires " RND" + 10 random chars suffix on reference mark names
              const rndSuffix = " RND" + PP2Zotero.Converter._generateCitationID() + PP2Zotero.Converter._generateCitationID().substring(0, 2);
              const name = "ZOTERO_ITEM CSL_CITATION " + JSON.stringify(zcLight) + rndSuffix;
              if (typeof Zotero !== "undefined") {
                Zotero.debug("PP2Zotero ODFWriter: refmark URI=" + (zc.citationItems[0] && zc.citationItems[0].uris ? zc.citationItems[0].uris[0] : "none"));
                Zotero.debug("PP2Zotero ODFWriter: refmark name length=" + name.length);
              }
              const esc = this._escapeXml(name);
              return {
                odf: '<text:reference-mark-start text:name="' + esc + '"/>'
                  + displayOdf
                  + '<text:reference-mark-end text:name="' + esc + '"/>',
                nextIndex: j
              };
            }
          }
        }
      }
    }

    // Paperpile bibliography → Zotero bibliography reference mark
    if (instrText.indexOf("paperpile_bibliography") !== -1) {
      const bibl = { uncited: [], omitted: [], custom: [] };
      const biblRnd = " RND" + PP2Zotero.Converter._generateCitationID() + PP2Zotero.Converter._generateCitationID().substring(0, 2);
      const name = "ZOTERO_BIBL " + JSON.stringify(bibl) + " CSL_BIBLIOGRAPHY" + biblRnd;
      const esc = this._escapeXml(name);
      return {
        odf: '<text:reference-mark-start text:name="' + esc + '"/>'
          + displayOdf
          + '<text:reference-mark-end text:name="' + esc + '"/>',
        nextIndex: j
      };
    }

    // Non-Paperpile field — just output display text
    return { odf: displayOdf, nextIndex: j };
  },

  /* ---- Run conversion ---- */

  _convertRun(wR) {
    let text = "";
    for (let i = 0; i < wR.childNodes.length; i++) {
      const child = wR.childNodes[i];
      if (child.nodeType !== 1) continue;
      if (child.localName === "t") {
        text += this._escapeXml(child.textContent);
      } else if (child.localName === "br") {
        text += "<text:line-break/>";
      } else if (child.localName === "tab") {
        text += "<text:tab/>";
      }
    }
    if (!text) return "";

    // Extract formatting properties
    const rPr = this._firstChildNS(wR, "rPr");
    const props = {};
    if (rPr) {
      // First, resolve character style (rStyle) — e.g., "Strong" → bold, "Emphasis" → italic
      const rStyleEl = this._firstChildNS(rPr, "rStyle");
      if (rStyleEl) {
        const rStyleId = rStyleEl.getAttribute("w:val");
        const styleDef = this._docStyles && this._docStyles.named[rStyleId];
        if (styleDef && styleDef.run) {
          Object.assign(props, styleDef.run);
        }
      }
      // Then, inline properties override the character style
      if (this._firstChildNS(rPr, "b")) props.bold = true;
      if (this._firstChildNS(rPr, "i")) props.italic = true;
      if (this._firstChildNS(rPr, "u")) props.underline = true;
      if (this._firstChildNS(rPr, "strike")) props.strikethrough = true;
      const va = this._firstChildNS(rPr, "vertAlign");
      if (va) {
        const val = va.getAttribute("w:val");
        if (val === "superscript") props.superscript = true;
        if (val === "subscript") props.subscript = true;
      }
      // Font size (w:sz is in half-points)
      const sz = this._firstChildNS(rPr, "sz");
      if (sz) {
        const halfPt = parseInt(sz.getAttribute("w:val"));
        if (halfPt) props.fontSize = (halfPt / 2) + "pt";
      }
      // Font family
      const rFonts = this._firstChildNS(rPr, "rFonts");
      if (rFonts) {
        const font = rFonts.getAttribute("w:ascii") || rFonts.getAttribute("w:hAnsi") || rFonts.getAttribute("w:cs");
        if (font) {
          props.fontFamily = font;
          this._usedFonts.add(font);
        }
      }
      // Text color
      const color = this._firstChildNS(rPr, "color");
      if (color) {
        const val = color.getAttribute("w:val");
        if (val && val !== "auto") props.color = "#" + val;
      }
    }

    if (Object.keys(props).length === 0) return text;

    const styleName = this._getAutoStyleName(props);
    return '<text:span text:style-name="' + styleName + '">' + text + '</text:span>';
  },

  /* ---- Footnote / Endnote conversion ---- */

  _buildNote(noteEl, noteClass, matchMap, footnoteMap, endnoteMap) {
    const counter = noteClass === "footnote" ? ++this._footnoteCounter : ++this._endnoteCounter;
    const noteId = (noteClass === "footnote" ? "ftn" : "edn") + counter;

    let bodyContent = "";
    for (let i = 0; i < noteEl.childNodes.length; i++) {
      const child = noteEl.childNodes[i];
      if (child.nodeType === 1 && child.localName === "p") {
        bodyContent += this._convertParagraph(child, matchMap, footnoteMap, endnoteMap);
      }
    }

    return '<text:note text:id="' + noteId + '" text:note-class="' + noteClass + '">'
      + '<text:note-citation>' + counter + '</text:note-citation>'
      + '<text:note-body>' + bodyContent + '</text:note-body>'
      + '</text:note>';
  },

  /* ---- Table conversion ---- */

  _convertTable(tblEl, matchMap, footnoteMap, endnoteMap) {
    this._tableCounter = (this._tableCounter || 0) + 1;
    let result = '<table:table table:style-name="Table' + this._tableCounter + '">';

    // Determine column count from first row
    const firstRow = tblEl.getElementsByTagNameNS(this.W_NS, "tr")[0];
    if (firstRow) {
      const cellCount = firstRow.getElementsByTagNameNS(this.W_NS, "tc").length;
      for (let c = 0; c < cellCount; c++) {
        result += '<table:table-column/>';
      }
    }

    const rows = tblEl.getElementsByTagNameNS(this.W_NS, "tr");
    for (let r = 0; r < rows.length; r++) {
      result += '<table:table-row>';
      const cells = rows[r].getElementsByTagNameNS(this.W_NS, "tc");
      for (let c = 0; c < cells.length; c++) {
        result += '<table:table-cell>';
        for (let p = 0; p < cells[c].childNodes.length; p++) {
          const child = cells[c].childNodes[p];
          if (child.nodeType === 1 && child.localName === "p") {
            result += this._convertParagraph(child, matchMap, footnoteMap, endnoteMap);
          }
        }
        result += '</table:table-cell>';
      }
      result += '</table:table-row>';
    }

    result += '</table:table>\n';
    return result;
  },

  /* ---- Automatic style management ---- */

  _getAutoStyleName(props) {
    const key = JSON.stringify(props);
    if (this._autoStyles[key]) return this._autoStyles[key].name;
    this._autoStyleCounter++;
    const name = "T" + this._autoStyleCounter;
    this._autoStyles[key] = { name, props };
    return name;
  },

  _buildAutoStylesXml() {
    let xml = "";
    // Text (run) auto styles
    for (const key of Object.keys(this._autoStyles)) {
      const style = this._autoStyles[key];
      const p = style.props;
      let tp = "";
      if (p.bold) tp += ' fo:font-weight="bold"';
      if (p.italic) tp += ' fo:font-style="italic"';
      if (p.underline) tp += ' style:text-underline-style="solid" style:text-underline-width="auto"';
      if (p.strikethrough) tp += ' style:text-line-through-style="solid"';
      if (p.superscript) tp += ' style:text-position="super 58%"';
      if (p.subscript) tp += ' style:text-position="sub 58%"';
      if (p.fontSize) tp += ' fo:font-size="' + p.fontSize + '"';
      if (p.fontFamily) tp += ' style:font-name="' + this._escapeXml(p.fontFamily) + '"';
      if (p.color) tp += ' fo:color="' + p.color + '"';
      xml += '<style:style style:name="' + style.name + '" style:family="text">'
        + '<style:text-properties' + tp + '/>'
        + '</style:style>\n';
    }
    // Paragraph auto styles (include both paragraph-properties and text-properties)
    for (const key of Object.keys(this._paraAutoStyles)) {
      const style = this._paraAutoStyles[key];
      const sp = style.props;
      const p = sp.para || sp; // support new {para, run} format and legacy
      const r = sp.run || {};
      let pp = "";
      if (p.textAlign) pp += ' fo:text-align="' + p.textAlign + '"';
      if (p.marginTop) pp += ' fo:margin-top="' + p.marginTop + '"';
      if (p.marginBottom) pp += ' fo:margin-bottom="' + p.marginBottom + '"';
      if (p.marginLeft) pp += ' fo:margin-left="' + p.marginLeft + '"';
      if (p.marginRight) pp += ' fo:margin-right="' + p.marginRight + '"';
      if (p.textIndent) pp += ' fo:text-indent="' + p.textIndent + '"';
      if (p.lineHeight) pp += ' fo:line-height="' + p.lineHeight + '"';
      let tp = "";
      if (r.fontSize) tp += ' fo:font-size="' + r.fontSize + '"';
      if (r.fontFamily) tp += ' style:font-name="' + this._escapeXml(r.fontFamily) + '"';
      if (r.bold) tp += ' fo:font-weight="bold"';
      if (r.italic) tp += ' fo:font-style="italic"';
      if (r.color) tp += ' fo:color="' + r.color + '"';
      xml += '<style:style style:name="' + style.name + '" style:family="paragraph" style:parent-style-name="Standard">'
        + (pp ? '<style:paragraph-properties' + pp + '/>' : '')
        + (tp ? '<style:text-properties' + tp + '/>' : '')
        + '</style:style>\n';
    }
    return xml;
  },

  /* ---- ODF XML builders ---- */

  _buildContentXml(bodyContent) {
    return '<?xml version="1.0" encoding="UTF-8"?>\n'
      + '<office:document-content'
      + ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"'
      + ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"'
      + ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"'
      + ' xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"'
      + ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"'
      + ' xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"'
      + ' office:version="1.2">\n'
      + this._buildFontFaceDecls()
      + '<office:automatic-styles>\n'
      + this._buildAutoStylesXml()
      + '</office:automatic-styles>\n'
      + '<office:body>\n'
      + '<office:text>\n'
      + bodyContent
      + '</office:text>\n'
      + '</office:body>\n'
      + '</office:document-content>';
  },

  _buildFontFaceDecls() {
    if (this._usedFonts.size === 0) return "";
    let xml = '<office:font-face-decls>\n';
    for (const font of this._usedFonts) {
      const escaped = this._escapeXml(font);
      const generic = /serif/i.test(font) || /times|garamond|georgia/i.test(font) ? "roman"
        : /mono|courier|consolas/i.test(font) ? "modern" : "swiss";
      xml += '<style:font-face style:name="' + escaped + '" svg:font-family="' + "'" + escaped + "'" + '" style:font-family-generic="' + generic + '"/>\n';
    }
    xml += '</office:font-face-decls>\n';
    return xml;
  },

  _buildStylesXml() {
    // Build "Standard" style from document defaults + Normal style
    const ds = this._docStyles || {};
    const defRun = { ...(ds.defaultRun || {}) };
    const defPara = { ...(ds.defaultPara || {}) };
    const normalStyle = (ds.named || {})["Normal"] || {};
    if (normalStyle.run) Object.assign(defRun, normalStyle.run);
    if (normalStyle.para) Object.assign(defPara, normalStyle.para);

    let standardParaProps = "";
    if (defPara.textAlign) standardParaProps += ' fo:text-align="' + defPara.textAlign + '"';
    if (defPara.lineHeight) standardParaProps += ' fo:line-height="' + defPara.lineHeight + '"';
    if (defPara.marginTop) standardParaProps += ' fo:margin-top="' + defPara.marginTop + '"';
    if (defPara.marginBottom) standardParaProps += ' fo:margin-bottom="' + defPara.marginBottom + '"';
    if (defPara.textIndent) standardParaProps += ' fo:text-indent="' + defPara.textIndent + '"';

    let standardTextProps = "";
    if (defRun.fontSize) standardTextProps += ' fo:font-size="' + defRun.fontSize + '"';
    if (defRun.fontFamily) standardTextProps += ' style:font-name="' + this._escapeXml(defRun.fontFamily) + '"';

    let standardStyle = '<style:style style:name="Standard" style:family="paragraph" style:class="text">\n';
    if (standardParaProps) standardStyle += '  <style:paragraph-properties' + standardParaProps + '/>\n';
    if (standardTextProps) standardStyle += '  <style:text-properties' + standardTextProps + '/>\n';
    standardStyle += '</style:style>\n';

    // Page layout from w:sectPr
    const pl = this._pageLayout || {};
    let pageLayoutProps = "";
    if (pl.pageWidth) pageLayoutProps += ' fo:page-width="' + pl.pageWidth + '"';
    if (pl.pageHeight) pageLayoutProps += ' fo:page-height="' + pl.pageHeight + '"';
    if (pl.marginTop) pageLayoutProps += ' fo:margin-top="' + pl.marginTop + '"';
    if (pl.marginRight) pageLayoutProps += ' fo:margin-right="' + pl.marginRight + '"';
    if (pl.marginBottom) pageLayoutProps += ' fo:margin-bottom="' + pl.marginBottom + '"';
    if (pl.marginLeft) pageLayoutProps += ' fo:margin-left="' + pl.marginLeft + '"';

    // Font face declarations for styles.xml
    let fontDecls = "";
    if (this._usedFonts.size > 0) {
      fontDecls = '<office:font-face-decls>\n';
      for (const font of this._usedFonts) {
        const esc = this._escapeXml(font);
        const generic = /serif|times|garamond|georgia/i.test(font) ? "roman"
          : /mono|courier|consolas/i.test(font) ? "modern" : "swiss";
        fontDecls += '<style:font-face style:name="' + esc + '" svg:font-family="' + "'" + esc + "'" + '" style:font-family-generic="' + generic + '"/>\n';
      }
      fontDecls += '</office:font-face-decls>\n';
    }

    return '<?xml version="1.0" encoding="UTF-8"?>\n'
      + '<office:document-styles'
      + ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"'
      + ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"'
      + ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"'
      + ' xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"'
      + ' xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"'
      + ' office:version="1.2">\n'
      + fontDecls
      + '<office:styles>\n'
      + standardStyle
      + '<style:style style:name="Heading_20_1" style:family="paragraph" style:parent-style-name="Standard" style:default-outline-level="1">\n'
      + '  <style:text-properties fo:font-size="18pt" fo:font-weight="bold"/>\n'
      + '</style:style>\n'
      + '<style:style style:name="Heading_20_2" style:family="paragraph" style:parent-style-name="Standard" style:default-outline-level="2">\n'
      + '  <style:text-properties fo:font-size="16pt" fo:font-weight="bold"/>\n'
      + '</style:style>\n'
      + '<style:style style:name="Heading_20_3" style:family="paragraph" style:parent-style-name="Standard" style:default-outline-level="3">\n'
      + '  <style:text-properties fo:font-size="14pt" fo:font-weight="bold"/>\n'
      + '</style:style>\n'
      + '<style:style style:name="Footnote" style:family="paragraph" style:parent-style-name="Standard">\n'
      + '  <style:text-properties fo:font-size="10pt"/>\n'
      + '</style:style>\n'
      + '</office:styles>\n'
      + (pageLayoutProps
        ? '<office:automatic-styles>\n'
          + '<style:page-layout style:name="pm1">\n'
          + '<style:page-layout-properties' + pageLayoutProps + '/>\n'
          + '</style:page-layout>\n'
          + '</office:automatic-styles>\n'
          + '<office:master-styles>\n'
          + '<style:master-page style:name="Standard" style:page-layout-name="pm1"/>\n'
          + '</office:master-styles>\n'
        : '')
      + '</office:document-styles>';
  },

  _buildMetaXml() {
    // Store ZOTERO_PREF as document custom property
    // (Zotero reads it via XDocumentPropertiesSupplier → custom properties)
    const sessionId = PP2Zotero.Converter._generateCitationID();
    const prefXml = '<data data-version="3" zotero-version="7.0.0">'
      + '<session id="' + sessionId + '"/>'
      + '<prefs>'
      + '<pref name="fieldType" value="ReferenceMark"/>'
      + '</prefs>'
      + '</data>';

    return '<?xml version="1.0" encoding="UTF-8"?>\n'
      + '<office:document-meta'
      + ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"'
      + ' xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"'
      + ' office:version="1.2">\n'
      + '<office:meta>\n'
      + '<meta:generator>PP2Zotero</meta:generator>\n'
      + '<meta:user-defined meta:name="ZOTERO_PREF" meta:value-type="string">' + this._escapeXml(prefXml) + '</meta:user-defined>\n'
      + '</office:meta>\n'
      + '</office:document-meta>';
  },

  _buildManifestXml() {
    return '<?xml version="1.0" encoding="UTF-8"?>\n'
      + '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">\n'
      + '<manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.text" manifest:full-path="/"/>\n'
      + '<manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>\n'
      + '<manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml"/>\n'
      + '<manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml"/>\n'
      + '</manifest:manifest>';
  },

  /* ---- Utilities ---- */

  _escapeXml(str) {
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  },

  _escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }
};
