#!/usr/bin/env python3
"""
compile_booklet.py — Master Pipeline Runner
============================================
Pass one or more Vatican document URLs and get a print-ready booklet .docx.

SINGLE DOCUMENT:
    python compile_booklet.py <url> [options]

MULTIPLE DOCUMENTS (combined into one booklet):
    python compile_booklet.py <url1> <url2> <url3> [options]

OPTIONS:
    --output, -o   Output .docx filename  (default: auto-named)
    --size,   -s   Page size: half-letter | a5 | letter  (default: half-letter)
    --keep-json    Keep intermediate JSON files after building
    --json-only    Only run Step 1 (fetch+extract), skip docx build
    --pdf-only     Only run Step 2 (build docx), requires existing JSON file(s)

EXAMPLES:
    # Single encyclical
    python compile_booklet.py \\
        https://www.vatican.va/content/john-paul-ii/en/encyclicals/documents/hf_jp-ii_enc_20030417_eccl-de-euch.html

    # Two documents combined into one booklet
    python compile_booklet.py \\
        https://www.vatican.va/.../lumen-gentium.html \\
        https://www.vatican.va/.../sacrosanctum-concilium.html \\
        --output Vatican_II_Combo.docx

    # A5 format
    python compile_booklet.py <url> --size a5 --output my_booklet.docx

REQUIREMENTS:
    pip install requests html2text beautifulsoup4 reportlab
"""

import sys
import os
import re
import json
import argparse
import tempfile
from pathlib import Path

# Ensure the pipeline directory is in path
sys.path.insert(0, str(Path(__file__).parent))
from vatican_fetch_content import fetch_and_extract
from build_booklet_docx import build_booklet_docx


# ─── Helpers ──────────────────────────────────────────────────────────────────

def url_to_slug(url: str) -> str:
    """Convert a URL to a safe filename slug."""
    slug = url.split('/')[-1].replace('.html', '').replace('.htm', '')
    slug = re.sub(r'[^\w-]', '_', slug)
    return slug[:60]


def merge_documents(docs: list[dict]) -> dict:
    """
    Merge multiple extracted documents into one combined document.
    Each document becomes its own chapter group, separated with a title page.
    """
    if len(docs) == 1:
        return docs[0]

    # Use first doc's metadata as the combined doc's header
    combined = {
        'doc_type': 'Vatican Documents',
        'title': 'Selected Vatican Documents',
        'author': '',
        'addressees': '',
        'subject': '',
        'latin_title': '',
        'url': '',
        'chapters': [],
        'footnotes': {},
    }

    footnote_offset = 0

    for doc in docs:
        # Add a divider "chapter" with the document title
        doc_title = doc.get('title') or doc.get('latin_title') or doc.get('doc_type', 'Document')
        doc_author = doc.get('author', '')

        divider = {
            'number': 0,
            'title': doc_title,
            'subtitle': doc_author,
            'paragraphs': [],
            '_is_doc_divider': True,
        }
        combined['chapters'].append(divider)

        # Offset footnote numbers to avoid collisions
        doc_footnotes = doc.get('footnotes', {})
        if footnote_offset > 0 and doc_footnotes:
            # Renumber footnotes and references in text
            offset_footnotes = {}
            for num, text in doc_footnotes.items():
                offset_footnotes[num + footnote_offset] = text
            # Update paragraph text references (superscript numbers)
            for ch in doc.get('chapters', []):
                for para in ch.get('paragraphs', []):
                    if para.get('text'):
                        para['text'] = offset_footnote_refs(para['text'], footnote_offset)
            combined['footnotes'].update(offset_footnotes)
        else:
            combined['footnotes'].update(doc_footnotes)

        combined['chapters'].extend(doc.get('chapters', []))
        footnote_offset += len(doc_footnotes)

    return combined


def offset_footnote_refs(text: str, offset: int) -> str:
    """Shift all footnote reference numbers in a text block by `offset`."""
    # This is a simple heuristic — bare numbers adjacent to punctuation
    def replacer(m):
        num = int(m.group(1))
        return m.group(0).replace(m.group(1), str(num + offset))

    return re.sub(r'(?<=[.,:;!?\'"])\s*(\d{1,3})(?=\s|$)', replacer, text)


# ─── Main Pipeline ────────────────────────────────────────────────────────────

def run_pipeline(
    urls: list[str],
    output_path: str,
    page_size: str = 'half-letter',
    keep_json: bool = False,
    json_only: bool = False,
    pdf_only: bool = False,
    existing_jsons: list[str] = None,
):
    work_dir = Path(output_path).parent
    work_dir.mkdir(parents=True, exist_ok=True)

    json_paths = []

    # ── Step 1: Fetch & Extract ───────────────────────────────────────────────
    if not pdf_only:
        for url in urls:
            slug = url_to_slug(url)
            json_path = work_dir / f"{slug}_extracted.json"

            print(f"\n{'='*60}")
            print(f"[Step 1] Extracting: {url}")
            print(f"{'='*60}")

            try:
                doc = fetch_and_extract(url)
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(doc, f, indent=2, ensure_ascii=False)
                print(f"[Step 1] ✓ Extracted → {json_path.name}")
                json_paths.append(str(json_path))
            except Exception as e:
                print(f"[Step 1] ✗ ERROR fetching {url}: {e}")
                raise

        if json_only:
            print(f"\n✓ JSON extraction complete. Files saved to: {work_dir}")
            return json_paths

    else:
        # pdf_only mode: use provided JSON files
        json_paths = existing_jsons or []

    # ── Step 2: Build PDF ─────────────────────────────────────────────────────
    print(f"\n{'='*60}")
    print(f"[Step 2] Building booklet docx")
    print(f"{'='*60}")

    # Load all extracted docs
    docs = []
    for jp in json_paths:
        with open(jp, 'r', encoding='utf-8') as f:
            docs.append(json.load(f))

    # Merge if multiple
    combined_doc = merge_documents(docs)

    # Auto-title output if multiple docs
    if len(docs) > 1 and 'Vatican Documents' in combined_doc.get('title', ''):
        titles = [d.get('title') or d.get('latin_title', '')[:20] for d in docs]
        combined_doc['title'] = ' & '.join(t for t in titles if t)

    # Write combined JSON for reference
    combined_json = work_dir / (Path(output_path).stem.replace('_booklet', '') + '_combined.json')
    if len(docs) > 1:
        with open(combined_json, 'w', encoding='utf-8') as f:
            json.dump(combined_doc, f, indent=2, ensure_ascii=False)

    # Build PDF from in-memory combined doc
    _build_from_doc(combined_doc, output_path, page_size)
    print(f"\n[Step 2] ✓ Booklet saved to: {output_path}")

    # Cleanup JSON files unless --keep-json
    if not keep_json and not pdf_only:
        for jp in json_paths:
            try:
                Path(jp).unlink()
            except Exception:
                pass

    print(f"\n{'='*60}")
    print(f"✓ DONE")
    print(f"  Output:    {output_path}")
    print(f"  Page size: {page_size}")
    chap_count = len(combined_doc.get('chapters', []))
    note_count = len(combined_doc.get('footnotes', {}))
    print(f"  Chapters:  {chap_count}   Endnotes: {note_count}")
    print(f"{'='*60}\n")


def _build_from_doc(doc: dict, output_path: str, page_size: str):
    """Write doc dict to a temp JSON and call the docx builder."""
    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as f:
        json.dump(doc, f, ensure_ascii=False)
        tmp_path = f.name

    try:
        build_booklet_docx(tmp_path, output_path, page_size)
    finally:
        os.unlink(tmp_path)


# ─── CLI ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description='Compile Vatican document URL(s) into a print-ready booklet .docx.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )

    parser.add_argument(
        'urls',
        nargs='+',
        help='Vatican document URL(s). For --pdf-only, pass JSON file path(s) instead.'
    )
    parser.add_argument(
        '--output', '-o',
        default=None,
        help='Output .docx filename (default: auto-named from document title/URL)'
    )
    parser.add_argument(
        '--size', '-s',
        default='half-letter',
        choices=['half-letter', 'a5', 'letter'],
        help='Page size: half-letter (5.5x8.5"), a5, or letter (default: half-letter)'
    )
    parser.add_argument(
        '--format', '-f',
        default='docx',
        choices=['docx', 'pdf'],
        help='Output format: docx (default, recommended for printing) or pdf'
    )
    parser.add_argument(
        '--keep-json',
        action='store_true',
        help='Keep intermediate JSON file(s) after building'
    )
    parser.add_argument(
        '--json-only',
        action='store_true',
        help='Only run Step 1 (fetch+extract), do not build PDF'
    )
    parser.add_argument(
        '--pdf-only',
        action='store_true',
        help='Only run Step 2 (build PDF from existing JSON files)'
    )

    args = parser.parse_args()

    # Determine output path
    if args.output is None:
        if len(args.urls) == 1:
            slug = url_to_slug(args.urls[0])
            args.output = f"{slug}_booklet.docx"
        else:
            args.output = "vatican_booklet.docx"

    # Validate pdf_only mode
    if args.pdf_only:
        for path in args.urls:
            if not Path(path).exists():
                print(f"ERROR: File not found for --pdf-only mode: {path}")
                sys.exit(1)
        run_pipeline(
            urls=[],
            output_path=args.output,
            page_size=args.size,
            pdf_only=True,
            existing_jsons=args.urls,
        )
    else:
        run_pipeline(
            urls=args.urls,
            output_path=args.output,
            page_size=args.size,
            keep_json=args.keep_json,
            json_only=args.json_only,
        )


if __name__ == '__main__':
    main()
