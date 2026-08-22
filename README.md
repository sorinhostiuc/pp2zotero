# PP2Zotero

PP2Zotero converts Paperpile citation fields in Microsoft Word `.docx` files into Zotero citation fields. It works on a copy by default, supports batch conversion, and matches references by identifiers and bibliographic metadata.

## Requirements

- Zotero versions 7-9
- A `.docx` document created with the Paperpile Word add-in
- Microsoft Word with the Zotero plugin for checking the converted document

## Installation

Download `pp2zotero-3.0.3.xpi` from the latest GitHub release. In Zotero, open `Tools > Plugins`, select `Install Add-on From File`, and choose the downloaded file.

## Converting a document

Open `Tools > Convert Paperpile Citations...` in Zotero. Select the `.docx` files, review the matching results, then start the conversion. Unless you enable overwriting, PP2Zotero writes `filename_zotero.docx` beside the original file.

The converter can search the Zotero library by DOI, PMID, ISBN, title, and year. When enabled, it imports missing DOI records; failing that, it creates items from the citation metadata. Unresolved citations receive Word comments so that they can be checked manually.

See [HELP.md](HELP.md) for every option and the troubleshooting notes.

## Building from source

Install Node.js and run `npm install`. Tests run with `npm test`. On Windows, build the XPI with `npm run build:windows`; on Unix-like systems, use `npm run build`.

## License

PP2Zotero is released under the MIT License. See [LICENSE](LICENSE).
