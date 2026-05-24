# url_to_book_tools

A 3-script pipeline that converts Vatican document URLs into print-ready Word (`.docx`) booklets, formatted for Word's **Book Fold** feature.

Print double-sided on Letter paper, fold in half, and you have a bound booklet.

---

## Requirements

```bash
pip install requests beautifulsoup4 python-docx lxml reportlab
```

Python 3.10+

---

## Usage

```bash
python compile_booklet.py <url> [options]
```

**Example:**
```bash
python compile_booklet.py \
  "https://www.vatican.va/content/john-paul-ii/en/encyclicals/documents/hf_jp-ii_enc_20030417_eccl-de-euch.html"
```

Output is saved in the current directory, named after the document (e.g. `hf_jp-ii_enc_20030417_eccl-de-euch_booklet.docx`). Use `--output` to specify a path:

```bash
python compile_booklet.py <url> --output ~/Desktop/MyBooklet.docx
```

### Options

| Flag | Description |
|------|-------------|
| `--output`, `-o` | Output file path (default: current directory, named from URL) |
| `--keep-json` | Keep the intermediate JSON after building |
| `--json-only` | Run Step 1 only — fetch and extract, skip the build |
| `--pdf-only` | Run Step 2 only — pass an existing JSON file instead of a URL |

---

## The 3 Scripts

### 1. `vatican_fetch_content.py`
Fetches the Vatican HTML and extracts a structured JSON file: chapters, paragraph numbers, body text, and footnotes. Vatican.va-specific.

### 2. `build_booklet_docx.py`
Converts the JSON into a formatted `.docx`. Source-agnostic — works with any JSON following the schema (see below).

### 3. `compile_booklet.py`
The entry point. Takes a URL, runs Step 1 into a temp JSON, pipes it into Step 2, deletes the temp file. You never touch the intermediate JSON.

---

## Printing in Word (Book Fold)

1. Open the `.docx` in Microsoft Word
2. Go to **Layout → Margins → Custom Margins**
3. Set **Multiple pages** → `Book fold`
4. Print **double-sided**, flip on **short edge**

The document is pre-formatted for 5.5 × 8.5″ pages with mirror margins (inner 0.875″ binding gutter, outer 0.625″), so it prints correctly on standard Letter paper folded in half.

---

## Document Formatting

- **Body text:** Times New Roman 11pt, justified
- **Footnotes:** Native Word footnotes, 9pt bold numbers, single-spaced
- **Headings:** Chapter titles and section headings with paragraph numbers
- **Verse/poetry:** Italic, left-aligned, stanza spacing preserved

---

## JSON Schema

To use `build_booklet_docx.py` with a non-Vatican source, produce a JSON file in this format:

```json
{
  "title": "Document Title",
  "author": "Author Name",
  "doc_type": "Encyclical",
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

---

## Known Limitations

- `vatican_fetch_content.py` is specific to Vatican.va. Other sources need a custom extractor that outputs the JSON schema above.
- Footnotes referenced within footnote text are not wired as native `<w:footnoteReference>` and appear inline instead.
