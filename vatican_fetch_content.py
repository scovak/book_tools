#!/usr/bin/env python3
"""
STEP 1 — Fetch & Extract
========================
Fetches a Vatican document URL and extracts structured content to JSON.

Handles three Vatican site templates automatically:
  - archive   /archive/hist_councils/... (Vatican II docs, old simplepage)
  - modern    /content/...               (encyclicals, letters, homilies, angelus)
  - toc       /roman_curia/...           (TOC index pages — raises helpful error)

Usage:
    python vatican_fetch_content.py <url> [--output extracted.json]

Requirements:
    pip install requests beautifulsoup4
"""

import sys, re, json, argparse, requests
from pathlib import Path
from bs4 import BeautifulSoup, NavigableString, Tag

DOCUMENT_TYPES = [
    'ENCYCLICAL LETTER', 'APOSTOLIC CONSTITUTION', 'APOSTOLIC EXHORTATION',
    'APOSTOLIC LETTER', 'PASTORAL CONSTITUTION', 'DOGMATIC CONSTITUTION',
    'DECLARATION', 'DECREE', 'INSTRUCTION', 'MOTU PROPRIO',
]
MODERN_DOC_TYPES = [
    'LETTER', 'HOMILY', 'ANGELUS', 'REGINA CAELI', 'SPEECH', 'MESSAGE',
    'GENERAL AUDIENCE', 'ENCYCLICAL', 'EXHORTATION',
]
CHAPTER_WORDS = {
    'ONE':1,'TWO':2,'THREE':3,'FOUR':4,'FIVE':5,'SIX':6,'SEVEN':7,'EIGHT':8,'NINE':9,'TEN':10,
    'I':1,'II':2,'III':3,'IV':4,'V':5,'VI':6,'VII':7,'VIII':8,'IX':9,'X':10,
}
STRIP_CLASS_FRAGMENTS = [
    'nav','menu','header','footer','breadcrumb','sidebar','banner','cookie',
    'social','share','search','login','language','lang-','toolbar','zoom','pdf-icon','print',
]

# ── Template Detection ─────────────────────────────────────────────────────────

def detect_template(url, text=''):
    if '/archive/' in url:
        return 'archive'
    if '/content/' in url:
        return 'modern'
    if '/roman_curia/' in url or '/pontifical' in url:
        upper = text[:4000].upper()
        if 'TABLE OF CONTENTS' in upper or upper.count('\n- ') > 15:
            # Only raise a toc error for pure index pages with no body content.
            # Single-page documents (like the Compendium) have a TOC at the top
            # followed by the full text — anchor links, not separate pages.
            if not re.search(r'(?m)^\s*\d+\.\s+\S', text):
                return 'toc'
        return 'archive'
    return 'modern'

def infer_doc_type_from_url(url):
    parts = url.lower().rstrip('/').split('/')
    type_map = {
        'encyclicals':        'Encyclical Letter',
        'apost_letters':      'Apostolic Letter',
        'apost_exhortations': 'Apostolic Exhortation',
        'apost_constitutions':'Apostolic Constitution',
        'letters':            'Letter',
        'homilies':           'Homily',
        'angelus':            'Angelus',
        'speeches':           'Speech',
        'messages':           'Message',
        'audiences':          'General Audience',
        'bull':               'Papal Bull',
        'motu_proprio':       'Motu Proprio',
    }
    for part in parts:
        if part in type_map:
            return type_map[part]
    return 'Document'

# ── Fetch ──────────────────────────────────────────────────────────────────────

def fetch_html(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "text/html,application/xhtml+xml",
    }
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    # Vatican servers sometimes report wrong charset in headers; force UTF-8
    # so smart quotes and em-dashes aren't mis-decoded as latin-1 garbage.
    r.encoding = r.apparent_encoding or 'utf-8'
    return r.text

def node_to_text(node, depth=0):
    if isinstance(node, NavigableString):
        return str(node)
    if not isinstance(node, Tag):
        return ''
    tag = node.name.lower() if node.name else ''
    def ct():
        return ''.join(node_to_text(c, depth+1) for c in node.children)
    if tag in ('h1','h2','h3','h4','h5','h6'):
        return f'\n\n{ct().strip().upper()}\n\n'
    if tag == 'p':
        inner = ct().strip()
        return f'\n{inner}\n' if inner else ''
    if tag in ('div','section','article','blockquote','li','dd','dt'):
        return f'\n{ct()}\n'
    if tag == 'br':
        return '\n'
    if tag in ('b','strong'):
        inner = ct().strip()
        return f'**{inner}**' if inner else ''
    if tag in ('i','em','cite'):
        inner = ct().strip()
        return f'*{inner}*' if inner else ''
    if tag == 'sup':
        inner = ct().strip()
        # Archive pages use <sup>(2)</sup> — strip surrounding parens so we
        # get '^2' not '^(2)' (which would double-process with convert_archive_fn_refs)
        inner = re.sub(r'^\((\d{1,3})\)$', r'\1', inner)
        return f'^{inner}' if inner else ''
    if tag == 'a':
        return ct()
    if tag in ('ul','ol'):
        return '\n' + ct() + '\n'
    if tag == 'hr':
        return '\n---\n'
    if tag in ('nav','header','footer','script','style','noscript','iframe','form','button'):
        return ''
    return ct()

def extract_text_from_html(html):
    soup = BeautifulSoup(html, 'html.parser')
    for tag_name in ['nav','header','footer','script','style','noscript','iframe','form','button']:
        for el in soup.find_all(tag_name):
            el.decompose()
    for el in soup.find_all(True):
        if el.attrs is None:  # already decomposed (parent was stripped)
            continue
        cls = ' '.join(el.get('class', [])).lower()
        if any(f in cls for f in STRIP_CLASS_FRAGMENTS):
            el.decompose()
    best = soup.body or soup
    max_p = 0
    for div in soup.find_all(['div','article','main','section']):
        p_count = len(div.find_all('p'))
        if p_count > max_p:
            max_p = p_count
            best = div
    return node_to_text(best).strip()

# ── Archive-specific: (N) → ^N conversion ────────────────────────────────────

def convert_archive_fn_refs(text):
    """Convert archive-style (N) footnote refs to ^N superscript markers.
    Only matches purely numeric parens like (1), not scripture refs like (1 John 1:2)."""
    return re.sub(r'\((\d{1,3})\)', r'^\1', text)

# ── Boundary Detection ────────────────────────────────────────────────────────

def find_content_boundaries(lines):
    start, end = 0, len(lines)
    for i, line in enumerate(lines):
        if re.search(r'Copyright\s*©', line, re.IGNORECASE):
            end = i; break
    for i, line in enumerate(lines[:300]):
        clean = re.sub(r'[*_#\[\]]', '', line).strip().upper()
        if any(dt in clean for dt in DOCUMENT_TYPES):
            start = i; break
    if start == 0:
        for i, line in enumerate(lines[:300]):
            if re.match(r'^\s*1\.\s+', line) or re.match(r'^CHAPTER\s+ONE', line.strip(), re.IGNORECASE):
                start = max(0, i - 10); break
    return start, end

def find_content_boundaries_modern(lines):
    start, end = 0, len(lines)
    for i, line in enumerate(lines):
        if re.search(r'Copyright\s*©', line, re.IGNORECASE):
            end = i; break
    for i, line in enumerate(lines[:200]):
        s = line.strip()
        if re.match(r'^\*\*\*', s) or re.match(r'^\s*1\.\s+\S', s):
            start = max(0, i - 3)
            break
    return start, end

def find_footnote_boundary(lines):
    for i, line in enumerate(lines):
        s = line.strip()
        # **NOTES** header always marks the boundary — no digit check needed
        if re.match(r'^\*\*NOTES?\*\*$', s, re.IGNORECASE):
            return i
        if s == '---':
            for j in range(i+1, min(i+5, len(lines))):
                if lines[j].strip():
                    if re.match(r'^\d{1,3}', lines[j].strip()):
                        return i
                    break
    last_para = 0
    for i, line in enumerate(lines):
        if re.match(r'^\s*\d+\.\s+\S', line):
            last_para = i
    fn_re = re.compile(r'^\s*(\d{1,3})\s*([A-Za-z\*"\(]|Cf\.|Ibid\.)')
    for i in range(last_para, len(lines)):
        s = lines[i].strip()
        if fn_re.match(s) and not re.match(r'^\d+\.\s', s):
            return i
    return len(lines)

# ── Header Parser ─────────────────────────────────────────────────────────────

def split_compound_line(line):
    parts = re.split(r'(\*\*[^*]+\*\*|\*\*\*[^*]+\*\*\*)', line)
    return [p.strip() for p in parts if p.strip()]

def parse_header(lines):
    meta = {'doc_type':'','title':'','author':'','addressees':[],'subject':'','latin_title':''}
    for line in lines:
        for sub in split_compound_line(line):
            _classify_header_line(sub, meta)
    meta['title'] = meta['latin_title'] or meta['title']
    meta['addressees'] = '\n'.join(meta['addressees'])
    return meta

def _classify_header_line(line, meta):
    clean = re.sub(r'[*_\[\]#^]', '', line).strip()
    if not clean or len(clean) < 3:
        return
    cup = clean.upper()
    if not meta['doc_type']:
        for dt in DOCUMENT_TYPES + MODERN_DOC_TYPES:
            if dt in cup:
                meta['doc_type'] = clean.strip(); return
    if re.search(r'\b(POPE|JOHN PAUL|BENEDICT XVI|FRANCIS|PIUS XII|LEO XIII|GREGORY|CLEMENT|INNOCENT|URBAN)\b', cup):
        if not meta['author']:
            meta['author'] = clean.strip(); return
    if re.match(r'^TO\s+(THE|ALL|BISHOPS|PRIESTS|DEACONS|FAITHFUL)', cup):
        meta['addressees'].append(clean.strip()); return
    if meta['addressees'] and re.match(r'^(BISHOPS|PRIESTS|DEACONS|MEN AND WOMEN|ALL THE|IN THE|AND ALL|CONSECRATED)', cup):
        meta['addressees'].append(clean.strip()); return
    if re.match(r'^ON\s+', cup) or re.match(r'^CONCERNING\s+', cup):
        meta['subject'] = clean.strip(); return
    bi_match = re.search(r'\*\*\*([^*]+)\*\*\*', line)
    if bi_match:
        candidate = bi_match.group(1).strip()
        if not meta['latin_title'] and len(candidate) > 3:
            meta['latin_title'] = candidate; return
    if not meta['title'] and len(clean) > 5 and clean != meta['doc_type']:
        meta['title'] = clean

# ── Chapter & Paragraph Parsers ───────────────────────────────────────────────

def parse_chapters(body_lines):
    chapters, current = [], None
    subtitle_buf = None

    def flush():
        if current:
            raw = current.pop('_raw', '')
            current['paragraphs'] = parse_paragraphs(raw)
            chapters.append(current)

    def finalize_subtitle():
        nonlocal subtitle_buf
        if subtitle_buf is not None and current:
            sub = re.sub(r'\*', '', subtitle_buf).strip()
            if sub:
                current['subtitle'] = sub
                current['title'] = f"Chapter {current['number']}: {sub}"
        subtitle_buf = None

    for line in body_lines:
        s = line.strip()
        # Strip bold/italic markers before testing chapter patterns
        # (Vatican II archive pages wrap headings in **bold**)
        s_clean = re.sub(r'\*+', '', s).strip()
        m = re.match(
            r'^CHAPTER\s+(ONE|TWO|THREE|FOUR|FIVE|SIX|SEVEN|EIGHT|NINE|TEN|'
            r'I{1,3}|IV|VI{0,3}|IX|X{1,3}|\d+)$', s_clean, re.IGNORECASE
        )
        if m:
            finalize_subtitle()
            flush()
            word = m.group(1).upper()
            try: num = int(word)
            except: num = CHAPTER_WORDS.get(word, len(chapters)+1)
            current = {'number':num,'title':f'Chapter {num}','subtitle':'','_raw':''}
            subtitle_buf = ''
            continue

        if subtitle_buf is not None:
            if not s:
                continue
            if re.match(r'^\d+\.\s', s) or re.match(r'^CHAPTER\s+', s, re.IGNORECASE):
                finalize_subtitle()
            elif s.startswith('**') or not subtitle_buf or (subtitle_buf and not subtitle_buf.endswith('**')):
                subtitle_buf = (subtitle_buf + ' ' + s).strip()
                if subtitle_buf.endswith('**') and subtitle_buf.count('**') >= 2:
                    finalize_subtitle()
                continue
            else:
                finalize_subtitle()

        sect_m = re.match(r'^\*\*\s*(INTRODUCTION|CONCLUSION|EPILOGUE|PROLOGUE|PREFACE)\s*\*\*$', s, re.IGNORECASE)
        if sect_m:
            finalize_subtitle()
            flush()
            name = sect_m.group(1).title()
            current = {'number':0,'title':name,'subtitle':'','_raw':''}
            subtitle_buf = None
            continue

        if current is None:
            current = {'number':0,'title':'Preamble','subtitle':'','_raw':''}

        current['_raw'] = current.get('_raw','') + line + '\n'

    finalize_subtitle()
    flush()
    return chapters

def parse_paragraphs(text):
    matches = list(re.finditer(r'(?m)^(\d+)\.\s+', text))
    if not matches:
        c = clean_text(text)
        return [{'number':None,'text':c}] if c.strip() else []
    result = []
    for i, m in enumerate(matches):
        num = int(m.group(1))
        start = m.end()
        end = matches[i+1].start() if i+1 < len(matches) else len(text)
        result.append({'number':num,'text':clean_text(text[start:end].strip())})
    return result

def parse_footnotes(lines):
    """Parse footnotes into a global dict, handling per-chapter restarts.

    Returns (footnotes, chapter_offsets) where:
      - footnotes: dict mapping global footnote number → text
      - chapter_offsets: list of per-chapter offsets; chapter_offsets[i] is
        the value to ADD to local footnote numbers in body chapter i to get
        the global number stored in `footnotes`.

    Many Vatican II archive docs restart footnote numbering at 1 for each
    chapter. This is detected when we see local number 1 after already having
    collected footnotes (i.e. the local number resets).
    """
    footnotes = {}
    cur_num, cur_text = None, []
    global_offset = 0   # cumulative offset across chapter restarts
    local_max = 0       # highest local number seen in the current section
    chapter_offsets = [0]  # preamble section starts at offset 0

    fn_start = re.compile(r'^\s*(\d{1,3})\.?\s*(.*)')
    fn_start_caret = re.compile(r'^\^(\d{1,3})\s*(.*)')
    # Header lines within the notes block: "Chapter I:", "Preface Article 1:", etc.
    header_re = re.compile(
        r'^(?:Chapter|Preface|Article|Preamble|Prologue|Introduction|Part)\b',
        re.IGNORECASE
    )

    def flush():
        nonlocal cur_num, cur_text
        if cur_num is not None:
            footnotes[cur_num] = clean_text(' '.join(cur_text))
        cur_num, cur_text = None, []

    for line in lines:
        s = line.strip()
        if not s:
            continue
        if s in ('---', '***') or re.match(r'^\*\*NOTES?\*\*$', s, re.IGNORECASE):
            continue
        # Skip section header lines like "Chapter I" or "Chapter I:" or "Preface Article 1:"
        s_clean = re.sub(r'\*+', '', s).strip()
        if header_re.match(s_clean):
            continue

        m = fn_start_caret.match(s) or fn_start.match(s)
        if m:
            local_n = int(m.group(1))
            rest = m.group(2).strip()
            # Require some content after the number (reject bare standalone digits)
            if not rest:
                if cur_num is not None:
                    cur_text.append(s)
                continue
            # Detect per-chapter restart: local number resets to 1
            if local_max > 0 and local_n == 1:
                flush()
                global_offset += local_max
                chapter_offsets.append(global_offset)
                local_max = 0

            global_n = local_n + global_offset
            flush()
            cur_num = global_n
            cur_text = [rest]
            if local_n > local_max:
                local_max = local_n
            continue

        if cur_num is not None:
            cur_text.append(s)

    flush()
    return footnotes, chapter_offsets


def _offset_fn_refs(text, offset):
    """Shift all ^N footnote-ref markers in text by `offset`."""
    if offset == 0:
        return text
    def replacer(m):
        return f'^{int(m.group(1)) + offset}'
    return re.sub(r'\^(\d{1,3})', replacer, text)


def clean_text(text):
    text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', text)
    text = re.sub(r'!\[[^\]]*\]\([^\)]+\)', '', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()

# ── Template-specific parsers ─────────────────────────────────────────────────

def parse_archive(text, url=''):
    """Parse old Vatican archive pages. Converts (N) footnote refs to ^N."""
    lines = text.split('\n')
    start, end = find_content_boundaries(lines)
    content_lines = lines[start:end]
    fn_boundary = find_footnote_boundary(content_lines)
    body_lines  = content_lines[:fn_boundary]
    fn_lines    = content_lines[fn_boundary:]

    # Convert (N) refs in body only
    body_text  = convert_archive_fn_refs('\n'.join(body_lines))
    # TEMP DEBUG
    with open('debug_body.txt', 'w', encoding='utf-8') as _f:
        _f.write(body_text)
    body_lines = body_text.split('\n')

    _CHAPTER_RE = re.compile(
        r'^CHAPTER\s+(ONE|TWO|THREE|FOUR|FIVE|SIX|SEVEN|EIGHT|NINE|TEN|'
        r'I{1,3}|IV|VI{0,3}|IX|X{1,3}|\d+)$', re.IGNORECASE
    )
    header_end = 0
    for i, line in enumerate(body_lines):
        s = line.strip()
        s_clean = re.sub(r'\*+', '', s).strip()
        # Use full chapter regex (with $) to avoid matching TOC lines like
        # "CHAPTER ONE GOD'S PLAN OF LOVE FOR HUMANITY"
        if re.match(r'^\s*\d+\.\s+\S', s) or _CHAPTER_RE.match(s_clean):
            header_end = i; break

    print(f"  Parsing archive structure...")
    meta     = parse_header(body_lines[:header_end])
    chapters = parse_chapters(body_lines[header_end:])
    footnotes, chapter_offsets = parse_footnotes(fn_lines)

    # If footnotes use per-chapter restart numbering, patch ^N refs in body text
    # so they match the globally-numbered footnotes dict.
    #
    # We can't rely on index-based alignment between `chapters` and
    # `chapter_offsets` because the Preamble and Chapter I may share one
    # footnote block while having separate chapter entries. Instead, detect
    # the restart dynamically: when a chapter's first ref drops back to or
    # below the highest ref seen so far, the source restarted numbering and
    # we should advance to the next offset block.
    if len(chapter_offsets) > 1:
        block_idx = 0
        highest_seen = 0

        for chapter in chapters:
            # Find the first and highest local footnote ref in this chapter
            first_ref = None
            max_ref = 0
            for para in chapter.get('paragraphs', []):
                refs = [int(x) for x in re.findall(r'\^(\d{1,3})', para.get('text', ''))]
                if refs:
                    if first_ref is None:
                        first_ref = min(refs)
                    max_ref = max(max_ref, max(refs))

            # If the first ref drops back to or below what we've seen, it's a restart
            if first_ref is not None:
                if highest_seen > 0 and first_ref <= highest_seen:
                    block_idx += 1
                highest_seen = max(highest_seen, max_ref)

            offset = chapter_offsets[block_idx] if block_idx < len(chapter_offsets) else 0
            if offset == 0:
                continue
            for para in chapter.get('paragraphs', []):
                if para.get('text'):
                    para['text'] = _offset_fn_refs(para['text'], offset)

    total_p = sum(len(ch.get('paragraphs',[])) for ch in chapters)
    print(f"  → {len(chapters)} chapters, {total_p} paragraphs, {len(footnotes)} footnotes")
    return {**meta, 'url': url, 'chapters': chapters, 'footnotes': footnotes}


def parse_modern(text, url=''):
    """Parse modern Vatican CMS pages (/content/...). Handles docs with no chapters."""
    lines = text.split('\n')
    start, end = find_content_boundaries_modern(lines)
    content_lines = lines[start:end]
    fn_boundary = find_footnote_boundary(content_lines)
    body_lines  = content_lines[:fn_boundary]
    fn_lines    = content_lines[fn_boundary:]

    header_end = 0
    for i, line in enumerate(body_lines):
        if re.match(r'^\s*\d+\.\s+\S', line):
            header_end = i; break

    print(f"  Parsing modern structure...")
    meta = parse_header(body_lines[:header_end])

    if not meta['doc_type']:
        meta['doc_type'] = infer_doc_type_from_url(url)
    if not meta['author']:
        u = url.lower()
        if   'john-paul-ii' in u:  meta['author'] = 'Pope John Paul II'
        elif 'benedict-xvi' in u:  meta['author'] = 'Pope Benedict XVI'
        elif '/francis/' in u:     meta['author'] = 'Pope Francis'
        elif 'paul-vi' in u:       meta['author'] = 'Pope Paul VI'

    chapters = parse_chapters(body_lines[header_end:])
    if not chapters:
        para_text = '\n'.join(body_lines[header_end:])
        paras = parse_paragraphs(para_text)
        if paras:
            label = meta['doc_type'] or 'Content'
            chapters = [{'number': 0, 'title': label, 'subtitle': '', 'paragraphs': paras}]
    else:
        # Rename auto-generated 'Preamble' chapter to doc_type for short modern docs
        for ch in chapters:
            if ch['title'] == 'Preamble' and meta['doc_type']:
                ch['title'] = meta['doc_type']

    footnotes, _ = parse_footnotes(fn_lines)
    total_p = sum(len(ch.get('paragraphs',[])) for ch in chapters)
    print(f"  → {len(chapters)} chapters, {total_p} paragraphs, {len(footnotes)} footnotes")
    return {**meta, 'url': url, 'chapters': chapters, 'footnotes': footnotes}


def parse_toc_error(url):
    raise ValueError(
        f"This URL appears to be a table-of-contents index page, not a single document:\n"
        f"  {url}\n\n"
        f"Each chapter lives at its own URL. Either:\n"
        f"  (a) Pass the URL of a specific chapter, or\n"
        f"  (b) Pass multiple chapter URLs to compile_booklet.py to combine them."
    )

# ── Main Extract ──────────────────────────────────────────────────────────────

def fetch_and_extract(url):
    print(f"  Fetching HTML...")
    html = fetch_html(url)
    print(f"  Extracting text from HTML...")
    text = extract_text_from_html(html)

    template = detect_template(url, text)
    print(f"  Detected template: {template}")

    if template == 'toc':
        parse_toc_error(url)
    elif template == 'archive':
        return parse_archive(text, url)
    else:
        return parse_modern(text, url)

# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Fetch and extract a Vatican document to JSON.')
    parser.add_argument('url')
    parser.add_argument('--output','-o', default=None)
    args = parser.parse_args()
    if args.output is None:
        slug = re.sub(r'[^\w-]','_', args.url.split('/')[-1].replace('.html',''))
        args.output = f"{slug}_extracted.json"
    print(f"\n[Step 1] Fetching: {args.url}")
    doc = fetch_and_extract(args.url)
    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    with open(out,'w',encoding='utf-8') as f:
        json.dump(doc, f, indent=2, ensure_ascii=False)
    print(f"[Step 1] ✓ Saved to: {out}\n")
    return str(out)

if __name__ == '__main__':
    main()