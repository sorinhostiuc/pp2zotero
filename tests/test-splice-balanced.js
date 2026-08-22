// Standalone test for _spliceBalanced and _findPaperpileFieldBlocks.
// Loads the real document.xml from /tmp/orig_unzipped, simulates field rewriting
// using a dummy replacement. Validation is done in a separate shell step.

const fs = require("fs");
const vm = require("vm");

const sandbox = { PP2Zotero: {}, Zotero: { debug: () => {} }, console };
const ctx = vm.createContext(sandbox);

const converterSrc = fs.readFileSync(
  "Y:/Zotero/Paperpile to zotero/pp2zotero/content/converter.js",
  "utf8"
);
vm.runInContext(converterSrc, ctx);

const C = sandbox.PP2Zotero.Converter;
if (!C) {
  console.error("ERROR: Converter not loaded");
  process.exit(1);
}

const TMP = "C:/Users/sorin/AppData/Local/Temp";
const xml = fs.readFileSync(TMP + "/orig_unzipped/word/document.xml", "utf8");
console.log("Loaded document.xml:", xml.length, "chars");

const blocks = C._findPaperpileFieldBlocks(xml);
console.log("Found", blocks.length, "paperpile blocks");

let lebaresIdx = -1;
for (let i = 0; i < blocks.length; i++) {
  if (blocks[i].displayContent && blocks[i].displayContent.indexOf("Lebares") !== -1) {
    lebaresIdx = i;
    break;
  }
}
console.log("Lebares block index:", lebaresIdx);
if (lebaresIdx >= 0) {
  const b = blocks[lebaresIdx];
  console.log("  startPos:", b.startPos, "endPos:", b.endPos);
  console.log("  cut length:", b.endPos - b.startPos);
  console.log();

  // Show all w:ins/w:del tag positions inside the cut
  const cut = xml.substring(b.startPos, b.endPos);
  const tagRegex = /<(\/?)w:(ins|del)\b([^>]*)>/g;
  const stacks = { ins: [], del: [] };
  const unmatchedCloses = { ins: 0, del: 0 };
  let m;
  console.log("  Track-change tags inside cut:");
  while ((m = tagRegex.exec(cut)) !== null) {
    const isClose = m[1] === '/';
    const elem = m[2];
    const attrs = m[3];
    const isSelfClose = attrs.trimEnd().endsWith('/');
    let action = "";
    if (isClose) {
      if (stacks[elem].length > 0) { stacks[elem].pop(); action = "POP"; }
      else { unmatchedCloses[elem]++; action = "ORPHAN_CLOSE"; }
    } else {
      if (isSelfClose) { action = "SELF_CLOSE"; }
      else { stacks[elem].push(attrs); action = "PUSH"; }
    }
    const idMatch = attrs.match(/w:id="(\d+)"/);
    const id = idMatch ? idMatch[1] : "?";
    console.log("    pos=" + m.index + " " + (isClose ? "</w:" : "<w:") + elem + " id=" + id + " => " + action);
  }
  console.log("  Final unmatchedCloses:", unmatchedCloses);
  console.log("  Final stacks.ins remaining opens:", stacks.ins.length);
  console.log("  Final stacks.del remaining opens:", stacks.del.length);

  console.log("\n  SUFFIX (first 200 chars after endPos):");
  console.log("  " + xml.substring(b.endPos, b.endPos + 200));
  console.log("  PREFIX (last 200 chars before startPos):");
  console.log("  " + xml.substring(b.startPos - 200, b.startPos));
}

const sortedBlocks = blocks.slice().sort((a, b) => b.startPos - a.startPos);
let result = xml;
const realisticJson = JSON.stringify({
  citationID: "ABC12345",
  properties: { formattedCitation: "(Author, 2024)", plainCitation: "(Author, 2024)", noteIndex: 0 },
  citationItems: [
    { id: 12345, uris: ["http://zotero.org/users/local/abcdefgh/items/AAAA1111"], uri: ["http://zotero.org/users/local/abcdefgh/items/AAAA1111"],
      itemData: { id: 12345, type: "article-journal", title: "A test article with a long title to simulate real CSL JSON content",
        author: [{ family: "Smith", given: "John A." }, { family: "Doe", given: "Jane B." }],
        issued: { "date-parts": [[2024, 6, 15]] }, "container-title": "Journal of Testing", volume: "10", issue: "3", page: "100-120",
        DOI: "10.1234/test.2024.001", abstract: "Lorem ipsum ".repeat(40) }
    }
  ],
  schema: "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
});
const escapedJson = realisticJson.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
for (const b of sortedBlocks) {
  // Simulate _stripTrackChangeWrappers on display content
  const cleanDisplay = (b.displayContent || '')
    .replace(/<w:(ins|del)\b([^>]*)>/g, function(match, elem, attrs) {
      return attrs.trimEnd().endsWith('/') ? match : '';
    })
    .replace(/<\/w:(ins|del)>/g, '');
  const replacement =
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM CSL_CITATION ' + escapedJson + ' </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    cleanDisplay +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  result = C._spliceBalanced(result, b.startPos, b.endPos, replacement);
}

fs.writeFileSync(TMP + "/result.xml", result);
console.log("\nWrote " + TMP + "/result.xml:", result.length, "chars");
