#!/bin/bash
set -e
cd "$(dirname "$0")"

VERSION=$(grep '"version"' manifest.json | head -1 | sed 's/.*: *"\(.*\)".*/\1/')
XPI_NAME="pp2zotero-${VERSION}.xpi"

echo "Building ${XPI_NAME}..."

rm -f "${XPI_NAME}"

zip -r "${XPI_NAME}" \
  manifest.json \
  bootstrap.js \
  prefs.js \
  content/ \
  locale/ \
  -x "*.DS_Store" \
  -x "tests/*" \
  -x "node_modules/*"

echo "Built: ${XPI_NAME} ($(du -h "${XPI_NAME}" | cut -f1))"
echo "Install: Open Zotero > Tools > Add-ons > gear icon > Install Add-on From File"
