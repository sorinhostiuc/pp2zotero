/* Matcher Module — matches Paperpile citations to Zotero library items */

if (typeof PP2Zotero === "undefined") var PP2Zotero = {};

PP2Zotero.Matcher = {
  _doiIndex: null,
  _titleYearIndex: null,
  _itemCache: null,

  async buildIndexes(libraryID) {
    this._doiIndex = new Map();
    this._titleYearIndex = new Map();
    this._titleOnlyIndex = new Map();
    this._itemCache = new Map();

    const s = new Zotero.Search();
    s.libraryID = libraryID;
    s.addCondition("itemType", "isNot", "attachment");
    s.addCondition("itemType", "isNot", "note");
    s.addCondition("includeDeleted", "false");
    const itemIDs = await s.search();
    const items = await Zotero.Items.getAsync(itemIDs);

    let indexedDOI = 0;
    let indexedTitle = 0;

    for (const item of items) {
      this._itemCache.set(item.id, item);

      const doi = item.getField("DOI");
      if (doi) {
        const normDOI = this._normalizeDOI(doi);
        if (normDOI) {
          this._doiIndex.set(normDOI, item.id);
          indexedDOI++;
        }
      }

      const title = item.getField("title");
      const date = item.getField("date");
      if (title) {
        const year = this._extractYear(date);
        const key = this._normalizeTitleYear(title, year);
        this._titleYearIndex.set(key, item.id);
        const titleKey = this._normalizeTitleYear(title, "").split("|")[0];
        if (!this._titleOnlyIndex.has(titleKey)) {
          this._titleOnlyIndex.set(titleKey, item.id);
        }
        indexedTitle++;
      }
    }

    Zotero.debug("PP2Zotero Matcher: indexed " + items.length + " items (" + indexedDOI + " DOI, " + indexedTitle + " title+year)");
    return { totalItems: items.length, indexedDOI, indexedTitle };
  },

  matchItem(citationItem, strategy) {
    const csl = citationItem.cslData;
    strategy = strategy || "doi_title_year";

    // DOI match (unless title_only strategy)
    if (strategy !== "title_only" && csl.DOI) {
      const normDOI = this._normalizeDOI(csl.DOI);
      if (normDOI && this._doiIndex.has(normDOI)) {
        const itemID = this._doiIndex.get(normDOI);
        return { match: this._itemCache.get(itemID), matchType: "doi", confidence: 100 };
      }
    }

    // Title+Year match (default strategy)
    if (strategy === "doi_title_year" && csl.title) {
      const year = this._extractYearFromCSL(csl);
      const key = this._normalizeTitleYear(csl.title, year);
      if (this._titleYearIndex.has(key)) {
        const itemID = this._titleYearIndex.get(key);
        return { match: this._itemCache.get(itemID), matchType: "title_year", confidence: 95 };
      }
    }

    // Title-only match (title_only strategy)
    if (strategy === "title_only" && csl.title) {
      const titleKey = this._normalizeTitleYear(csl.title, "").split("|")[0];
      if (this._titleOnlyIndex.has(titleKey)) {
        const itemID = this._titleOnlyIndex.get(titleKey);
        return { match: this._itemCache.get(itemID), matchType: "title_only", confidence: 85 };
      }
    }

    return { match: null, matchType: "none", confidence: 0 };
  },

  matchAll(citations, strategy) {
    const results = [];
    for (const citation of citations) {
      const itemMatches = citation.citationItems.map(citItem => {
        const result = this.matchItem(citItem, strategy);
        return {
          paperpileItemId: citItem.paperpileItemId,
          cslData: citItem.cslData,
          locator: citItem.locator,
          locatorType: citItem.locatorType,
          prefix: citItem.prefix,
          suffix: citItem.suffix,
          suppressAuthor: citItem.suppressAuthor,
          ...result
        };
      });
      results.push({ citation, itemMatches });
    }
    return results;
  },

  async findExisting(cslData, libraryID) {
    // Broader search beyond exact DOI / title+year — catches near-duplicates

    // Try PMID
    if (cslData.PMID) {
      try {
        const s = new Zotero.Search();
        s.libraryID = libraryID;
        s.addCondition("extra", "contains", "PMID: " + cslData.PMID);
        s.addCondition("includeDeleted", "false");
        const ids = await s.search();
        if (ids.length > 0) {
          const item = await Zotero.Items.getAsync(ids[0]);
          Zotero.debug("PP2Zotero: Found existing item by PMID " + cslData.PMID);
          return item;
        }
      } catch (e) { /* ignore search errors */ }
    }

    // Try ISBN
    if (cslData.ISBN) {
      try {
        const s = new Zotero.Search();
        s.libraryID = libraryID;
        s.addCondition("ISBN", "is", cslData.ISBN);
        s.addCondition("includeDeleted", "false");
        const ids = await s.search();
        if (ids.length > 0) {
          const item = await Zotero.Items.getAsync(ids[0]);
          Zotero.debug("PP2Zotero: Found existing item by ISBN " + cslData.ISBN);
          return item;
        }
      } catch (e) { /* ignore search errors */ }
    }

    // Try title-only (without requiring year match) + verify first author
    if (cslData.title && cslData.title.length > 10) {
      try {
        const s = new Zotero.Search();
        s.libraryID = libraryID;
        s.addCondition("title", "contains", cslData.title.substring(0, 60));
        s.addCondition("itemType", "isNot", "attachment");
        s.addCondition("itemType", "isNot", "note");
        s.addCondition("includeDeleted", "false");
        const ids = await s.search();
        for (const id of ids) {
          const item = await Zotero.Items.getAsync(id);
          const existTitle = this._normalizeTitleYear(item.getField("title"), "").split("|")[0];
          const searchTitle = this._normalizeTitleYear(cslData.title, "").split("|")[0];
          if (existTitle === searchTitle) {
            // Verify at least first author matches
            const creators = item.getCreators();
            const firstAuthor = cslData.author && cslData.author[0];
            if (!firstAuthor || !creators.length) {
              Zotero.debug("PP2Zotero: Found existing item by title (no author check): " + cslData.title.substring(0, 50));
              return item;
            }
            const existFamily = (creators[0].lastName || "").toLowerCase();
            const searchFamily = (firstAuthor.family || "").toLowerCase();
            if (existFamily === searchFamily) {
              Zotero.debug("PP2Zotero: Found existing item by title+author: " + cslData.title.substring(0, 50));
              return item;
            }
          }
        }
      } catch (e) { /* ignore search errors */ }
    }

    return null;
  },

  async importByDOI(doi, libraryID, collectionID) {
    try {
      const translate = new Zotero.Translate.Search();
      translate.setIdentifier({ DOI: doi });

      const translators = await translate.getTranslators();
      if (!translators.length) {
        Zotero.debug("PP2Zotero: No translator found for DOI " + doi);
        return null;
      }

      translate.setTranslator(translators[0]);
      const newItems = await translate.translate({ libraryID });

      if (newItems && newItems.length > 0) {
        const item = newItems[0];

        if (collectionID) {
          item.addToCollection(collectionID);
          await item.saveTx();
        }

        this._itemCache.set(item.id, item);
        const itemDOI = item.getField("DOI");
        if (itemDOI) {
          this._doiIndex.set(this._normalizeDOI(itemDOI), item.id);
        }
        const title = item.getField("title");
        const date = item.getField("date");
        if (title) {
          const year = this._extractYear(date);
          this._titleYearIndex.set(this._normalizeTitleYear(title, year), item.id);
        }

        Zotero.debug("PP2Zotero: Imported item via DOI: " + doi);
        return item;
      }
    } catch (e) {
      Zotero.debug("PP2Zotero: DOI import failed for " + doi + ": " + e.message);
    }
    return null;
  },

  async createFromCSLData(cslData, libraryID, collectionID) {
    try {
      // Map CSL type to Zotero item type
      const typeMap = {
        "article-journal": "journalArticle",
        "article": "journalArticle",
        "book": "book",
        "chapter": "bookSection",
        "paper-conference": "conferencePaper",
        "thesis": "thesis",
        "report": "report",
        "webpage": "webpage",
        "patent": "patent"
      };
      const itemType = typeMap[cslData.type] || "document";

      const item = new Zotero.Item(itemType);
      item.libraryID = libraryID;

      // Safe setField - some fields aren't valid for all item types
      const safeSet = (field, value) => {
        if (!value) return;
        try { item.setField(field, String(value)); } catch (e) {
          Zotero.debug("PP2Zotero: skipped field " + field + " on " + itemType + ": " + e.message);
        }
      };

      safeSet("title", cslData.title);
      safeSet("DOI", cslData.DOI);
      safeSet("url", cslData.URL);
      safeSet("volume", cslData.volume);
      safeSet("issue", cslData.issue);
      safeSet("pages", cslData.page);
      safeSet("publisher", cslData.publisher);
      safeSet("ISBN", cslData.ISBN);
      safeSet("abstractNote", cslData.abstract);
      if (cslData.PMID) safeSet("extra", "PMID: " + cslData.PMID);

      // Journal/book title
      if (cslData["container-title"]) {
        const journalField = (itemType === "bookSection") ? "bookTitle"
          : (itemType === "conferencePaper") ? "proceedingsTitle"
          : "publicationTitle";
        safeSet(journalField, cslData["container-title"]);
      }

      // Date
      if (cslData.issued && cslData.issued["date-parts"] && cslData.issued["date-parts"][0]) {
        const year = cslData.issued["date-parts"][0][0];
        if (year && year !== 0 && year !== "0") {
          safeSet("date", String(year));
        }
      }

      // Authors
      if (cslData.author && cslData.author.length) {
        const creators = cslData.author.map(a => {
          if (a.isInstitution) {
            return { lastName: a.family || "", creatorType: "author", fieldMode: 1 };
          }
          return { firstName: a.given || "", lastName: a.family || "", creatorType: "author" };
        });
        item.setCreators(creators);
      }

      if (collectionID) {
        item.addToCollection(collectionID);
      }

      await item.saveTx();

      // Add to indexes
      this._itemCache.set(item.id, item);
      if (cslData.DOI) {
        this._doiIndex.set(this._normalizeDOI(cslData.DOI), item.id);
      }
      if (cslData.title) {
        const year = this._extractYearFromCSL(cslData);
        this._titleYearIndex.set(this._normalizeTitleYear(cslData.title, year), item.id);
      }

      Zotero.debug("PP2Zotero: Created " + itemType + " from Paperpile data: " + (cslData.title || "").substring(0, 50));
      return item;
    } catch (e) {
      Zotero.debug("PP2Zotero: Failed to create item from data: " + e.message);
      return null;
    }
  },

  _normalizeDOI(doi) {
    if (!doi) return null;
    return doi.trim().toLowerCase()
      .replace(/^https?:\/\/(dx\.)?doi\.org\//, "")
      .replace(/^doi:\s*/i, "");
  },

  _normalizeTitleYear(title, year) {
    const normTitle = title.toLowerCase()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .replace(/[^\w\s]/g, "")
      .replace(/\s+/g, " ")
      .trim();
    return normTitle + "|" + (year || "");
  },

  _extractYear(dateStr) {
    if (!dateStr) return "";
    const match = dateStr.match(/(\d{4})/);
    return match ? match[1] : "";
  },

  _extractYearFromCSL(cslData) {
    if (cslData.issued && cslData.issued["date-parts"] && cslData.issued["date-parts"][0]) {
      return String(cslData.issued["date-parts"][0][0] || "");
    }
    return "";
  },

  getCSLJSON(zoteroItem) {
    return Zotero.Utilities.itemToCSLJSON(zoteroItem);
  },

  buildURI(zoteroItem) {
    let uri = null;

    // Try Zotero's built-in API first
    try {
      uri = Zotero.URI.getItemURI(zoteroItem);
    } catch (e) {
      Zotero.debug("PP2Zotero: Zotero.URI.getItemURI() failed: " + e.message);
    }

    // Manual fallback if API failed
    if (!uri) {
      const key = zoteroItem.key;
      const libraryType = zoteroItem.library ? zoteroItem.library.libraryType : "user";
      if (libraryType === "group") {
        const groupID = zoteroItem.library.id || zoteroItem.libraryID;
        uri = "http://zotero.org/groups/" + groupID + "/items/" + key;
      } else {
        const userID = Zotero.Users.getCurrentUserID();
        if (userID) {
          uri = "http://zotero.org/users/" + userID + "/items/" + key;
        } else {
          try {
            uri = "http://zotero.org/users/local/" + Zotero.Users.getLocalUserKey() + "/items/" + key;
          } catch (e2) {
            uri = "http://zotero.org/users/0/items/" + key;
          }
        }
      }
    }

    // Log to error console so user can see it (first 5 items only)
    if (!this._uriLogCount) this._uriLogCount = 0;
    if (this._uriLogCount < 5) {
      Zotero.log("PP2Zotero URI #" + this._uriLogCount + ": " + uri);
      this._uriLogCount++;
    }

    return uri;
  }
};
