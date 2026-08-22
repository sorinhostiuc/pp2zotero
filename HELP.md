# PP2Zotero Help

## What is PP2Zotero?

PP2Zotero converts Paperpile citations in Word documents (.docx) to Zotero citations. After conversion, you can open the document in Word and use the Zotero Word plugin to manage your citations normally.

## Requirements

- Zotero 7 (version 7.0.x)
- A `.docx` file with Paperpile citations
- Your Zotero library should already contain most of the cited references (or have DOIs so PP2Zotero can import them)

## Installation

1. Download `pp2zotero.xpi`
2. In Zotero: **Tools > Add-ons > gear icon > Install Add-on From File**
3. Select the `.xpi` file
4. Restart Zotero

## How to Use

### Step 1: Open the Converter

Go to **Tools > Convert Paperpile Citations...** in Zotero's menu bar.

### Step 2: Select Your Document

- **Drag and drop** a `.docx` file onto the drop zone, or **click** to browse
- You can select **multiple files** for batch processing
- Only `.docx` files are supported (not `.doc`, `.odt`, etc.)

### Step 3: Configure Options

#### Overwrite original file
- **Unchecked** (default): Saves the converted file as `filename_zotero.docx` alongside the original
- **Checked**: Overwrites the original file in place

#### Create automatic backup
- **Checked** (default): Creates a `.docx.bak` copy before conversion (only when overwriting)

#### Auto-import missing references from DOI
- **Checked** (default): If a Paperpile citation is not found in your Zotero library, PP2Zotero will try to import it automatically:
  1. First checks if the item already exists under a different format (PMID, ISBN, or similar title)
  2. If not found, tries to import from CrossRef using the DOI
  3. If no DOI, creates the item directly from the Paperpile metadata

#### Match strategy
Controls how PP2Zotero matches Paperpile citations to your Zotero library:

- **DOI + Title** (default, recommended): First tries to match by DOI, then falls back to title + year matching. Best for most users.
- **DOI only**: Matches exclusively by DOI. Use this if you have many items with similar titles across different editions.
- **Title only**: Matches by normalized title only (ignores DOI). Use this if your Paperpile DOIs differ from Zotero DOIs.

#### Save references to
Choose which Zotero collection to save imported references to. Click the button to open a collection picker. Use the **+** button to create a new collection.

All references cited in the document (both newly imported and already existing in your library) will be added to the selected collection. This is useful for organizing references per document or project.

### Step 4: Scan

Click **Scan document**. PP2Zotero will:
1. Extract all Paperpile citation fields from the document (including footnotes and endnotes)
2. Match each citation against your Zotero library
3. Auto-import missing references (if enabled)

### Step 5: Review Results

The results screen shows a table of all unique references with their match status:

| Badge | Meaning |
|-------|---------|
| **DOI exact** (green) | Matched by DOI - highest confidence |
| **Title+Year** (green) | Matched by normalized title and year |
| **Title match** (green) | Matched by title only |
| **Found existing** (green) | Found via broader search (PMID, ISBN, or title+author) |
| **Imported DOI** (blue) | Successfully imported from CrossRef or created from metadata |
| **Missing** (red) | No match found - needs manual action |

#### Actions for Missing References

- **Edit**: Opens an inline editor to correct the title, authors, year, DOI, or item type before importing
- **Import / Create**: Imports from CrossRef (if DOI available) or creates directly from metadata
- **Skip**: Leaves this citation unconverted in the document

You can **sort** the table by clicking any column header.

### Step 6: Convert

Click **Convert N/N** (shows how many references are resolved). PP2Zotero will:
1. Rewrite all matched citation field codes from Paperpile to Zotero format
2. Rewrite the bibliography field (if present)
3. Add Word comments at any unconverted citations so you can find them easily
4. Save the converted document

### Step 7: Completion

The completion screen shows conversion statistics and offers:

- **Copy report**: Copies a text summary to clipboard
- **Export report**: Saves a detailed CSV or TXT report
- **Open in Word**: Opens the converted document
- **Close**: Closes the dialog

## Batch Processing

Select multiple `.docx` files (via file picker or drag-and-drop) to convert them all in one go. Each file is processed independently with the same settings.

## After Conversion

1. Open the converted `.docx` in Microsoft Word
2. The Zotero Word plugin should recognize all converted citations
3. You can use Zotero's "Refresh" button to update citation formatting
4. Add new citations or modify existing ones using Zotero as usual

## Troubleshooting

### "No Paperpile citations found"
- Make sure the document was created with Paperpile's Word plugin
- Documents exported as plain text (without field codes) cannot be converted
- PDF files are not supported

### Some citations show as "Missing"
- The reference may not exist in your Zotero library
- Try clicking **Edit** to correct the metadata, then **Import**
- Check that the DOI is correct
- If auto-import is enabled and still fails, the CrossRef database may not have the DOI

### Converted citations don't appear in Zotero Word plugin
- Make sure you're using the Zotero 7 Word plugin
- Try clicking "Refresh" in the Zotero toolbar in Word
- If the document was converted from a very old Paperpile format, some fields may not convert correctly

### Word comments appear in the document
- These mark citations that could not be converted
- Search for comments by "PP2Zotero" in Word to find them
- You can manually insert Zotero citations at these locations, then delete the comments

### Plugin doesn't appear in Tools menu
- Restart Zotero after installation
- Check that the plugin is enabled in **Tools > Add-ons**
- Verify Zotero version is 7.0.x

## Supported Citation Features

PP2Zotero preserves the following Paperpile citation properties during conversion:

- Page locators (e.g., "p. 45")
- Locator types (page, chapter, section, etc.)
- Citation prefixes (e.g., "see ")
- Citation suffixes
- Author suppression (e.g., "Author (2020)" vs "(Author 2020)")
- Multi-source citations (e.g., "(Smith 2020; Jones 2021)")
- Footnote and endnote citations
- Institutional/collective authors

## Supported Item Types

| Paperpile Type | Zotero Type |
|---------------|-------------|
| Journal Article | journalArticle |
| Book | book |
| Book Section | bookSection |
| Conference Paper | conferencePaper |
| Thesis | thesis |
| Report | report |
| Web Page | webpage |
| Patent | patent |
| Other | document |

## Contact & Issues

Report bugs or request features at: https://github.com/sorinhostiuc/pp2zotero
