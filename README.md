# url_to_book_tools

A pipeline that converts Vatican document URLs into print-ready Word (`.docx`) booklets, formatted for 5.5 × 8.5″ half-letter pages. Print double-sided on Letter paper, fold in half, and you have a bound booklet.

---

## Requirements

```bash
pip install requests beautifulsoup4 python-docx lxml
```

Python 3.10+

---

## Scripts

### `compile_booklet.py` — One or more documents

Fetch one or more Vatican URLs and build a single booklet. Multiple URLs are merged into one volume in the order given.

```bash
python compile_booklet.py <url> [options]
python compile_booklet.py <url1> <url2> ... [options]
```

**Examples:**
```bash
# Single encyclical
python compile_booklet.py \
  "https://www.vatican.va/content/john-paul-ii/en/encyclicals/documents/hf_jp-ii_enc_20030417_eccl-de-euch.html"

# Two documents combined into one booklet
python compile_booklet.py \
  "https://www.vatican.va/archive/.../lumen-gentium_en.html" \
  "https://www.vatican.va/archive/.../sacrosanctum-concilium_en.html" \
  --output LG_SC_booklet.docx
```

| Flag | Description |
|------|-------------|
| `--output`, `-o` | Output `.docx` path (default: named from URL, saved in current directory) |
| `--size`, `-s` | Page size: `half-letter` \| `a5` \| `letter` (default: `half-letter`) |
| `--toc` | Insert a Table of Contents page after the title page |
| `--keep-json` | Keep the intermediate JSON after building |
| `--json-only` | Fetch and extract only — skip the docx build |
| `--docx-only` | Skip Step 1 — build docx directly from existing JSON file(s) passed as arguments |

---

### `compile_volume.py` — Multi-document volume

Build a single booklet from a JSON manifest listing multiple Vatican URLs. Useful for assembling Vatican II volumes, collections of encyclicals, etc.

```bash
python compile_volume.py <manifest.json> [options]
```

**Example:**
```bash
python compile_volume.py manifests/vat2_constitutions.json
python compile_volume.py manifests/vat2_decrees.json --output MyVolume.docx
```

| Flag | Description |
|------|-------------|
| `--output`, `-o` | Output `.docx` path (overrides manifest `output` field) |
| `--size`, `-s` | Page size: `half-letter` \| `a5` \| `letter` (default: `half-letter`) |
| `--keep-json` | Keep intermediate per-document JSON files |
| `--json-only` | Fetch and extract all docs only — skip docx build |

---

## Manifest Format

Manifests live in the `manifests/` folder. Each is a JSON file describing a volume:

```json
{
  "title": "Vatican II — Constitutions",
  "author": "Second Vatican Council",
  "doc_type": "Dogmatic & Pastoral Constitutions",
  "output": "vat2_constitutions_booklet.docx",
  "toc": true,
  "documents": [
    {
      "url": "https://www.vatican.va/archive/.../sacrosanctum-concilium_en.html",
      "title": "Sacrosanctum Concilium"
    },
    {
      "url": "https://www.vatican.va/archive/.../lumen-gentium_en.html",
      "title": "Lumen Gentium"
    }
  ]
}
```

| Field | Required | Description |
|-------|----------|-------------|
| `title` | No | Volume title shown on the cover page |
| `author` | No | Author or body (e.g. "Second Vatican Council") |
| `doc_type` | No | Document type label shown above the title |
| `output` | No | Default output filename (overridden by `--output`) |
| `toc` | No | `true` to insert a two-level Table of Contents |
| `documents` | **Yes** | Ordered list of documents to include |
| `documents[].url` | **Yes** | Vatican.va URL for the document |
| `documents[].title` | No | Override the extracted title (recommended) |
| `documents[].author` | No | Override the extracted author |

### Included manifests

| File | Contents |
|------|----------|
| `manifests/vat2_constitutions.json` | The 4 Vatican II constitutions (SC, LG, DV, GS) |
| `manifests/vat2_constitutions_1.json` | Same as above, with `"toc": true` |
| `manifests/vat2_decrees.json` | The 9 Vatican II decrees |
| `manifests/vat2_declarations.json` | The 3 Vatican II declarations |

---

## Table of Contents

When `"toc": true` is set in a manifest (or `--toc` is passed to `compile_booklet.py`), a TOC page is inserted after the title page. It uses two levels:

- **Level 1** — Document titles (e.g. *Sacrosanctum Concilium*)
- **Level 2** — Chapter titles within each document

The TOC is a native Word field. Page numbers are populated when you open the file — Word will prompt you to update fields, or you can right-click the TOC and choose **Update Field**.

---

## Printing in Word (Book Fold)

1. Open the `.docx` in Microsoft Word
2. Go to **Layout → Margins → Custom Margins**
3. Set **Multiple pages** → `Book fold`
4. Print **double-sided**, flip on **short edge**

The document is pre-formatted for 5.5 × 8.5″ pages with mirror margins (inner 0.875″ binding gutter, outer 0.625″), so it prints correctly on standard Letter paper folded in half.

---

## Document Formatting

| Element | Style |
|---------|-------|
| Body text | Times New Roman 11pt, justified |
| Footnotes | Native Word footnotes, 9pt, single-spaced |
| Chapter headings | Times New Roman 13pt bold, centered |
| Document titles | Times New Roman 16pt bold italic, centered |
| Verse / poetry | Italic, left-aligned, stanza spacing preserved |
| Page numbers | Centered footer, 8pt gray |
| Header | Document title (left) on verso/recto pages |

---

## How It Works

### Step 1 — `vatican_fetch_content.py`

Fetches the Vatican HTML and extracts a structured JSON: title, author, chapters, paragraphs, and footnotes. Automatically detects which of three Vatican page templates the URL uses:

| Template | URL pattern | Used for |
|----------|-------------|----------|
| `archive` | `/archive/hist_councils/...` | Vatican II constitutions, decrees, declarations |
| `modern` | `/content/...` | Encyclicals, apostolic letters, homilies |
| `toc` | `/roman_curia/...` | Index pages — returns an error with chapter URLs |

Footnote reference formats handled automatically:

| Format | Example | Documents |
|--------|---------|-----------|
| Parenthetical | `(1)` | Most archive docs |
| Square bracket | `[1]` | Sacrosanctum Concilium |
| Superscript | `<sup>1</sup>` | Some archive docs |
| Caret marker | `^1` | Modern template docs |

Per-chapter footnote numbering restarts (common in Vatican II documents) are resolved by mapping all body references to sequential global keys in document order.

### Step 2 — `build_booklet_docx.py`

Converts the JSON into a formatted `.docx`. Source-agnostic — works with any JSON following the schema below.

Two-pass process:
- **Pass 1:** Build the document using python-docx, with superscript placeholders for footnote references
- **Pass 2:** Patch the zip directly to inject `word/footnotes.xml` and wire native Word `<w:footnoteReference>` elements

### Step 3 — `compile_booklet.py` / `compile_volume.py`

Orchestrates Steps 1 and 2. For volumes, merges multiple documents with correctly offset footnote numbering so all references are globally unique.

---

## JSON Schema

To use `build_booklet_docx.py` with a non-Vatican source, produce a JSON file in this format:

```json
{
  "title": "Document Title",
  "author": "Author Name",
  "doc_type": "Encyclical",
  "toc": false,
  "chapters": [
    {
      "number": 1,
      "title": "Chapter Title",
      "subtitle": "Optional subtitle",
      "paragraphs": [
        {
          "number": 1,
          "text": "Paragraph text. Supports *italic* and **bold** inline.\n\nMultiple blocks separated by blank lines."
        }
      ]
    }
  ],
  "footnotes": {
    "1": "Footnote text for reference 1.",
    "2": "Footnote text for reference 2."
  }
}
```

Inline formatting in paragraph text:

| Syntax | Result |
|--------|--------|
| `*text*` | Italic |
| `**text**` | Bold |
| `***text***` | Bold italic |
| `^N` | Footnote reference (linked to footnotes dict) |

---

## Known Limitations

- `vatican_fetch_content.py` is specific to Vatican.va. Other sources need a custom extractor that outputs the JSON schema above.
- Footnotes referenced within footnote text appear inline rather than as nested Word footnotes.
- Multi-page documents at `/roman_curia/` URLs are not supported as a single URL — pass each chapter URL individually.
- TOC page numbers require a manual update field step in Word after opening.
