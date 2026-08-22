/* Node.js test for scanner XML parsing logic */
const fs = require("fs");
const { DOMParser } = require("@xmldom/xmldom");
const assert = require("assert");

global.DOMParser = DOMParser;

const xmlContent = fs.readFileSync(__dirname + "/fixtures/sample-document.xml", "utf-8");
const parser = new DOMParser();
const xmlDoc = parser.parseFromString(xmlContent, "application/xml");

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// Test: find all fldChar begin elements
const allRuns = xmlDoc.getElementsByTagNameNS(W_NS, "r");
console.log("Total <w:r> elements:", allRuns.length);

// Count PAPERPILE fields
let ppCitationCount = 0;
let ppBiblCount = 0;
for (let i = 0; i < allRuns.length; i++) {
  const instrText = allRuns[i].getElementsByTagNameNS(W_NS, "instrText")[0];
  if (instrText) {
    const text = instrText.textContent.trim();
    if (text.startsWith("ADDIN PAPERPILE_CITATION")) ppCitationCount++;
    if (text.startsWith("ADDIN PAPERPILE_BIBL")) ppBiblCount++;
  }
}

assert.strictEqual(ppCitationCount, 2, "Should find 2 Paperpile citations");
assert.strictEqual(ppBiblCount, 1, "Should find 1 Paperpile bibliography");

// Test JSON extraction from first citation
for (let i = 0; i < allRuns.length; i++) {
  const instrText = allRuns[i].getElementsByTagNameNS(W_NS, "instrText")[0];
  if (instrText && instrText.textContent.trim().startsWith("ADDIN PAPERPILE_CITATION")) {
    const jsonStr = instrText.textContent.trim().substring("ADDIN PAPERPILE_CITATION".length).trim();
    const parsed = JSON.parse(jsonStr);

    assert.ok(parsed.id, "Should have Paperpile ID");
    assert.ok(Array.isArray(parsed.citationItems), "Should have citationItems array");
    assert.ok(parsed.citationItems[0].itemData, "First item should have itemData");
    assert.ok(parsed.citationItems[0].itemData.title, "First item should have title");

    console.log("First citation OK:", parsed.id, "-", parsed.citationItems[0].itemData.title);
    break;
  }
}

// Test cluster citation (multiple items in one field)
let clusterFound = false;
for (let i = 0; i < allRuns.length; i++) {
  const instrText = allRuns[i].getElementsByTagNameNS(W_NS, "instrText")[0];
  if (instrText && instrText.textContent.includes("XkRm2")) {
    const jsonStr = instrText.textContent.trim().substring("ADDIN PAPERPILE_CITATION".length).trim();
    const parsed = JSON.parse(jsonStr);

    assert.strictEqual(parsed.citationItems.length, 2, "Cluster should have 2 items");
    assert.strictEqual(parsed.citationItems[1].locator, "45", "Second item should have locator");
    assert.strictEqual(parsed.citationItems[1].prefix, "vezi ", "Second item should have prefix");

    console.log("Cluster citation OK: 2 items with locator and prefix");
    clusterFound = true;
    break;
  }
}
assert.ok(clusterFound, "Should find cluster citation");

console.log("\nAll scanner tests passed!");
