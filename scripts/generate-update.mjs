import { createHash } from "node:crypto";
import { readFileSync, writeFileSync } from "node:fs";
import { basename } from "node:path";

const xpi = process.argv[2];
if (!xpi) throw new Error("Usage: node scripts/generate-update.mjs <xpi>");

const manifest = JSON.parse(readFileSync("manifest.json", "utf8"));
const zotero = manifest.applications.zotero;
const repository = process.env.GITHUB_REPOSITORY || "mapispev/pp2zotero";
const tag = process.env.GITHUB_REF_NAME || `v${manifest.version}`;
const hash = createHash("sha256").update(readFileSync(xpi)).digest("hex");

const update = {
  addons: {
    [zotero.id]: {
      updates: [{
        version: manifest.version,
        update_link: `https://github.com/${repository}/releases/download/${tag}/${basename(xpi)}`,
        update_hash: `sha256:${hash}`,
        applications: {
          zotero: {
            strict_min_version: zotero.strict_min_version,
            strict_max_version: zotero.strict_max_version,
          },
        },
      }],
    },
  },
};

writeFileSync("update.json", `${JSON.stringify(update, null, 2)}\n`);
