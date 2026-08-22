/* Scanner Module — extracts Paperpile citations from .docx OOXML */

if (typeof PP2Zotero === "undefined") var PP2Zotero = {};

PP2Zotero.Scanner = {
  FIELD_PREFIX_CITATION: "ADDIN paperpile_citation",
  FIELD_PREFIX_BIBL: "ADDIN paperpile_bibliography",
  W_NS: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",

  async scan(docxData) {
    // Re-create Uint8Array in this scope to avoid cross-compartment instanceof failures with JSZip
    if (docxData && docxData.byteLength !== undefined && !(docxData instanceof Uint8Array)) {
      docxData = new Uint8Array(docxData);
    }
    const JSZipRef = typeof JSZip !== "undefined" ? JSZip : (await import("./lib/jszip.min.js")).default;
    const zip = await JSZipRef.loadAsync(docxData);

    const results = { citations: [], bibliography: null, errors: [] };

    const xmlFiles = ["word/document.xml", "word/footnotes.xml", "word/endnotes.xml"];
    for (const xmlPath of xmlFiles) {
      const file = zip.file(xmlPath);
      if (!file) continue;
      const text = await file.async("string");
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(text, "application/xml");
      const context = {
        isFootnote: xmlPath === "word/footnotes.xml",
        isEndnote: xmlPath === "word/endnotes.xml"
      };
      if (typeof Zotero !== "undefined") {
        const runs = xmlDoc.getElementsByTagNameNS(this.W_NS, "r");
        Zotero.debug("PP2Zotero Scanner: " + xmlPath + " has " + runs.length + " runs");
      }
      await this._extractFields(xmlDoc, results, context);
    }

    if (typeof Zotero !== "undefined") {
      Zotero.debug("PP2Zotero Scanner: found " + results.citations.length + " citations, " + results.errors.length + " errors");
      if (results.errors.length > 0) {
        Zotero.debug("PP2Zotero Scanner: first error: " + JSON.stringify(results.errors[0]));
      }
    }
    return results;
  },

  async _extractFields(xmlDoc, results, context) {
    const allRuns = xmlDoc.getElementsByTagNameNS(this.W_NS, "r");
    let i = 0;

    while (i < allRuns.length) {
      const run = allRuns[i];
      const fldChar = run.getElementsByTagNameNS(this.W_NS, "fldChar")[0];

      if (fldChar && fldChar.getAttribute("w:fldCharType") === "begin") {
        const field = this._collectField(allRuns, i);
        if (field) {
          await this._processField(field, results, context);
          i = field.endIndex + 1;
          continue;
        }
      }
      i++;
    }
  },

  _collectField(allRuns, beginIndex) {
    let instrTextParts = [];
    let displayTextParts = [];
    let phase = "instr";
    let endIndex = beginIndex;
    let nestLevel = 1;
    let footnoteId = null;

    for (let i = beginIndex + 1; i < allRuns.length; i++) {
      const run = allRuns[i];
      const fldChar = run.getElementsByTagNameNS(this.W_NS, "fldChar")[0];

      if (fldChar) {
        const charType = fldChar.getAttribute("w:fldCharType");
        if (charType === "begin") {
          nestLevel++;
        } else if (charType === "separate") {
          if (nestLevel === 1) phase = "display";
        } else if (charType === "end") {
          nestLevel--;
          if (nestLevel === 0) {
            endIndex = i;
            break;
          }
        }
        continue;
      }

      if (phase === "instr") {
        const instrEl = run.getElementsByTagNameNS(this.W_NS, "instrText")[0];
        if (instrEl) instrTextParts.push(instrEl.textContent);
      } else if (phase === "display") {
        const tEl = run.getElementsByTagNameNS(this.W_NS, "t")[0];
        if (tEl) displayTextParts.push(tEl.textContent);
      }
    }

    let fnParent = allRuns[beginIndex];
    while (fnParent && fnParent.parentNode) {
      fnParent = fnParent.parentNode;
      if (fnParent.localName === "footnote") {
        footnoteId = fnParent.getAttribute("w:id");
        break;
      }
      if (fnParent.localName === "endnote") {
        footnoteId = fnParent.getAttribute("w:id");
        break;
      }
    }

    const instrText = instrTextParts.join("").trim();
    if (!instrText) return null;

    return {
      instrText,
      displayText: displayTextParts.join(""),
      beginIndex,
      endIndex,
      footnoteId
    };
  },

  async _processField(field, results, context) {
    const { instrText, displayText } = field;

    if (instrText.indexOf("paperpile_bibliography") !== -1 || instrText.indexOf("PAPERPILE_BIBL") !== -1) {
      results.bibliography = {
        rawInstrText: instrText,
        fieldBeginIndex: field.beginIndex,
        fieldEndIndex: field.endIndex
      };
      return;
    }

    if (instrText.indexOf("paperpile_citation") === -1 && instrText.indexOf("PAPERPILE_CITATION") === -1) {
      if (typeof Zotero !== "undefined") Zotero.debug("PP2Zotero Scanner: skipping non-PP field: " + instrText.substring(0, 80));
      return;
    }
    if (typeof Zotero !== "undefined") Zotero.debug("PP2Zotero Scanner: found PP citation field, length=" + instrText.length);

    // Parse XML metadata from instrText
    // instrText looks like: ADDIN paperpile_citation <clusterId>XX</clusterId><version>...</version><metadata>...</metadata><data>BASE64</data> \* MERGEFORMAT
    // XML entities are NOT escaped here (instrText is already text content from the OOXML)
    const clusterId = this._extractTag(instrText, "clusterId");
    const dataB64 = this._extractTag(instrText, "data");

    if (!dataB64) {
      results.errors.push({ type: "citation_no_data", message: "No data blob", displayText });
      return;
    }

    // Decode base64 + zlib compressed JSON
    let items;
    try {
      items = await this._decodeDataBlob(dataB64);
    } catch (e) {
      results.errors.push({ type: "citation_decode_error", message: e.message, displayText });
      return;
    }

    // Parse metadata XML for locators, prefixes, suppress-author etc.
    const metadataXml = this._extractTag(instrText, "metadata");
    const metaCitations = metadataXml ? this._parseMetadataCitations(metadataXml) : [];

    // Build citation items by combining data items with metadata
    const citationItems = items.map((ppItem, idx) => {
      const meta = metaCitations.find(m => m.id === ppItem._id) || metaCitations[idx] || {};
      const cslData = this._paperpileToCSL(ppItem);
      return {
        paperpileItemId: ppItem._id || "",
        cslData: cslData,
        locator: meta.locator || null,
        locatorType: meta.locator_label || null,
        prefix: meta.prefix || null,
        suffix: meta.suffix || null,
        suppressAuthor: meta.no_author === true
      };
    });

    const citation = {
      fieldIndex: results.citations.length,
      rawInstrText: instrText,
      paperpileId: clusterId || "",
      citationItems: citationItems,
      formattedText: displayText,
      fieldBeginIndex: field.beginIndex,
      fieldEndIndex: field.endIndex,
      isInFootnote: context.isFootnote,
      isInEndnote: context.isEndnote,
      footnoteId: field.footnoteId || null
    };

    results.citations.push(citation);
  },

  _extractTag(text, tagName) {
    const open = "<" + tagName + ">";
    const close = "</" + tagName + ">";
    const start = text.indexOf(open);
    if (start === -1) return null;
    const end = text.indexOf(close, start);
    if (end === -1) return null;
    return text.substring(start + open.length, end);
  },

  async _decodeDataBlob(b64) {
    // Decode base64 to binary
    const binaryStr = atob(b64);
    const bytes = new Uint8Array(binaryStr.length);
    for (let i = 0; i < binaryStr.length; i++) {
      bytes[i] = binaryStr.charCodeAt(i);
    }

    // DecompressionStream is a Web API on Window, not available in bootstrap scope
    const mainWindow = typeof Zotero !== "undefined" ? Zotero.getMainWindow() : window;
    const ds = new mainWindow.DecompressionStream("deflate");
    const writer = ds.writable.getWriter();
    writer.write(bytes);
    writer.close();

    const reader = ds.readable.getReader();
    const chunks = [];
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
    }

    // Concatenate chunks and decode as UTF-8
    const totalLen = chunks.reduce((sum, c) => sum + c.length, 0);
    const merged = new Uint8Array(totalLen);
    let offset = 0;
    for (const chunk of chunks) {
      merged.set(chunk, offset);
      offset += chunk.length;
    }

    const jsonStr = new TextDecoder().decode(merged);
    return JSON.parse(jsonStr);
  },

  _parseMetadataCitations(metadataStr) {
    // Parse: <citation><id>XX</id><no_author/><prefix/><suffix/><locator/><locator_label>page</locator_label></citation>
    const citations = [];
    const citationPattern = /<citation>([\s\S]*?)<\/citation>/g;
    let match;
    while ((match = citationPattern.exec(metadataStr)) !== null) {
      const block = match[1];
      const id = this._extractTag(block, "id");
      const locator = this._extractTag(block, "locator");
      const locatorLabel = this._extractTag(block, "locator_label");
      const prefix = this._extractTag(block, "prefix");
      const suffix = this._extractTag(block, "suffix");
      const noAuthorTag = this._extractTag(block, "no_author");
      // <no_author/> is an empty placeholder (like <prefix/>, <suffix/>) — NOT a boolean flag.
      // Only treat as suppress-author when tag has explicit truthy content: <no_author>true</no_author> or <no_author>1</no_author>
      const noAuthor = noAuthorTag !== null && noAuthorTag !== "" && noAuthorTag !== "false" && noAuthorTag !== "0";
      if (typeof Zotero !== "undefined") {
        Zotero.debug("PP2Zotero: citation id=" + (id || "?") + " no_author tag=" + JSON.stringify(noAuthorTag) + " => suppressAuthor=" + noAuthor);
        Zotero.debug("PP2Zotero: metadata block: " + block.substring(0, 200));
      }
      citations.push({
        id: id || "",
        locator: locator || null,
        locator_label: locatorLabel || null,
        prefix: prefix || null,
        suffix: suffix || null,
        no_author: noAuthor
      });
    }
    return citations;
  },

  _paperpileToCSL(ppItem) {
    // Convert Paperpile item format to CSL-JSON-like structure for the matcher
    const csl = {
      title: ppItem.title || "",
      DOI: ppItem.doi || ppItem.DOI || "",
      PMID: ppItem.pmid || "",
      ISBN: ppItem.isbn || ""
    };

    // Authors
    if (ppItem.author && ppItem.author.length) {
      csl.author = ppItem.author.map(a => {
        if (a.collective) {
          return { family: a.collective, given: "", isInstitution: true };
        }
        return { family: a.last || "", given: a.first || "" };
      });
    }

    // Date: Paperpile uses published.year
    if (ppItem.published && ppItem.published.year) {
      csl.issued = {
        "date-parts": [[parseInt(ppItem.published.year) || ppItem.published.year]]
      };
    }

    // Journal
    if (ppItem.journal) csl["container-title"] = ppItem.journal;
    if (ppItem.journalfull) csl["container-title"] = ppItem.journalfull;
    if (ppItem.volume) csl.volume = ppItem.volume;
    if (ppItem.issue) csl.issue = ppItem.issue;
    if (ppItem.pages) csl.page = ppItem.pages;
    if (ppItem.publisher) csl.publisher = ppItem.publisher;
    if (ppItem.url) csl.URL = ppItem.url;
    if (ppItem.abstract) csl.abstract = ppItem.abstract;

    // Map pubtype
    const typeMap = {
      "PP_ARTICLE": "article-journal",
      "PP_BOOK": "book",
      "PP_BOOK_SECTION": "chapter",
      "PP_CONFERENCE": "paper-conference",
      "PP_THESIS": "thesis",
      "PP_REPORT": "report",
      "PP_WEBPAGE": "webpage",
      "PP_PATENT": "patent"
    };
    csl.type = typeMap[ppItem.pubtype] || "article";

    return csl;
  }
};
