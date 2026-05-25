#!/usr/bin/env python3
"""
compile_volume.py — Manifest-Driven Volume Builder
====================================================
Build a single print-ready booklet .docx from a JSON manifest that lists
multiple Vatican document URLs. Useful for assembling Vatican II volumes,
collections of encyclicals, etc.

USAGE:
    python compile_volume.py <manifest.json> [options]

OPTIONS:
    --output, -o   Output .docx filename  (default: taken from manifest)
    --size,   -s   Page size: half-letter | a5 | letter  (default: half-letter)
    --keep-json    Keep intermediate per-document JSON files after building
    --json-only    Only run Step 1 (fetch+extract all docs), skip docx build

MANIFEST FORMAT:
    {
      "title": "Vatican II — Constitutions",
      "author": "Second Vatican Council",
      "doc_type": "Conciliar Documents",
      "output": "vatican2_constitutions_booklet.docx",
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

    Fields:
      title       Volume title shown on the cover page
      author      Volume author/body (optional)
      doc_type    Document type label (optional)
      output      Default output filename (overridden by --output)
      documents   List of documents to include, in order
        url       Vatican.va URL for the document
        title     Override the extracted title (optional but recommended)
        author    Override the extracted author (optional)

EXAMPLES:
    python compile_volume.py manifests/vat2_constitutions.json
    python compile_volume.py manifests/vat2_decrees.json --output MyVolume.docx
    python compile_volume.py manifests/vat2_constitutions.json --json-only

REQUIREMENTS:
    pip install requests beautifulsoup4 python-docx lxml
"""

import sys
import os
import re
import json
import argparse
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from vatican_fetch_content import fetch_and_extract
from build_booklet_docx import build_booklet_docx
from compile_booklet import merge_documents, offset_footnote_refs


# ── Helpers ────────────────────────────────────────────────────────────────────

def load_manifest(path: str) -> dict:
    with open(path, 'r', encoding='utf-8') as f:
        manifest = json.load(f)
    required = ['documents']
    for key in required:
        if key not in manifest:
            raise ValueError(f"Manifest missing required field: '{key}'")
    if not manifest['documents']:
        raise ValueError("Manifest 'documents' list is empty")
    return manifest


def slug_from_title(title: str) -> str:
    slug = re.sub(r'[^\w\s-]', '', title.lower())
    slug = re.sub(r'[\s-]+', '_', slug).strip('_')
    return slug[:60] or 'volume'


def url_to_slug(url: str) -> str:
    slug = url.split('/')[-1].replace('.html', '').replace('.htm', '')
    return re.sub(r'[^\w-]', '_', slug)[:60]


# ── Core Pipeline ──────────────────────────────────────────────────────────────

def run_volume(
    manifest: dict,
    output_path: str,
    page_size: str = 'half-letter',
    keep_json: bool = False,
    json_only: bool = False,
    work_dir: Path = None,
):
    if work_dir is None:
        work_dir = Path(output_path).parent
    work_dir.mkdir(parents=True, exist_ok=True)

    documents = manifest['documents']
    total = len(documents)

    print(f"\n{'='*60}")
    print(f"Volume: {manifest.get('title', 'Untitled')}")
    print(f"Documents: {total}")
    print(f"Output: {output_path}")
    print(f"{'='*60}")

    # ── Step 1: Fetch & extract each document ──────────────────────────────
    json_paths = []
    for i, doc_entry in enumerate(documents, 1):
        url = doc_entry['url']
        slug = url_to_slug(url)
        json_path = work_dir / f"{slug}_extracted.json"

        print(f"\n[{i}/{total}] Extracting: {doc_entry.get('title') or url.split('/')[-1]}")

        try:
            doc = fetch_and_extract(url)
        except Exception as e:
            print(f"  ✗ ERROR: {e}")
            raise

        # Apply manifest overrides
        if doc_entry.get('title'):
            doc['title'] = doc_entry['title']
        if doc_entry.get('author'):
            doc['author'] = doc_entry['author']

        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(doc, f, indent=2, ensure_ascii=False)

        fn_count = len(doc.get('footnotes', {}))
        para_count = sum(len(ch.get('paragraphs', [])) for ch in doc.get('chapters', []))
        print(f"  ✓ {len(doc.get('chapters', []))} chapters, {para_count} paragraphs, {fn_count} footnotes")
        json_paths.append(str(json_path))

    if json_only:
        print(f"\n✓ Extraction complete. JSON files in: {work_dir}")
        return json_paths

    # ── Step 2: Merge and build ────────────────────────────────────────────
    print(f"\n{'='*60}")
    print(f"[Step 2] Merging {total} documents and building booklet")
    print(f"{'='*60}")

    docs = []
    for jp in json_paths:
        with open(jp, 'r', encoding='utf-8') as f:
            docs.append(json.load(f))

    combined = merge_documents(docs)

    # Apply volume-level metadata from manifest
    if manifest.get('title'):
        combined['title'] = manifest['title']
    if manifest.get('author'):
        combined['author'] = manifest['author']
    if manifest.get('doc_type'):
        combined['doc_type'] = manifest['doc_type']

    # Save combined JSON for reference
    combined_slug = slug_from_title(manifest.get('title', 'volume'))
    combined_json = work_dir / f"{combined_slug}_combined.json"
    with open(combined_json, 'w', encoding='utf-8') as f:
        json.dump(combined, f, indent=2, ensure_ascii=False)
    print(f"  Combined JSON → {combined_json.name}")

    # Build docx
    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as f:
        json.dump(combined, f, ensure_ascii=False)
        tmp_path = f.name

    try:
        build_booklet_docx(tmp_path, output_path, page_size)
    finally:
        os.unlink(tmp_path)

    # Cleanup per-doc JSONs unless --keep-json
    if not keep_json:
        for jp in json_paths:
            try:
                Path(jp).unlink()
            except Exception:
                pass

    chap_count = len(combined.get('chapters', []))
    note_count = len(combined.get('footnotes', {}))

    print(f"\n{'='*60}")
    print(f"✓ DONE")
    print(f"  Output:     {output_path}")
    print(f"  Page size:  {page_size}")
    print(f"  Documents:  {total}")
    print(f"  Chapters:   {chap_count}   Footnotes: {note_count}")
    print(f"{'='*60}\n")

    return output_path


# ── CLI ────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description='Build a Vatican document volume from a JSON manifest.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument('manifest', help='Path to the volume manifest JSON file')
    parser.add_argument('--output', '-o', default=None,
                        help='Output .docx filename (overrides manifest "output" field)')
    parser.add_argument('--size', '-s', default='half-letter',
                        choices=['half-letter', 'a5', 'letter'],
                        help='Page size (default: half-letter)')
    parser.add_argument('--keep-json', action='store_true',
                        help='Keep intermediate per-document JSON files')
    parser.add_argument('--json-only', action='store_true',
                        help='Only fetch and extract; skip docx build')

    args = parser.parse_args()

    manifest = load_manifest(args.manifest)

    # Determine output path: CLI > manifest > auto-generated
    if args.output:
        output_path = args.output
    elif manifest.get('output'):
        output_path = manifest['output']
    else:
        slug = slug_from_title(manifest.get('title', 'volume'))
        output_path = f"{slug}_booklet.docx"

    work_dir = Path(output_path).parent

    run_volume(
        manifest=manifest,
        output_path=output_path,
        page_size=args.size,
        keep_json=args.keep_json,
        json_only=args.json_only,
        work_dir=work_dir,
    )


if __name__ == '__main__':
    main()
