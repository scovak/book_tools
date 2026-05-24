#!/usr/bin/env python3
"""
STEP 1 — Fetch & Extract
========================
Fetches a Vatican document URL and extracts structured content to JSON.

Usage:
    python fetch_content.py <url> [--output extracted.json]

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
CHAPTER_WORDS = {
    'ONE':1,'TWO':2,'THREE':3,'FOUR':4,'FIVE':5,'SIX':6,'SEVEN':7,'EIGHT':8,'NINE':9,'TEN':10,
    'I':1,'II':2,'III':3,'IV':4,'V':5,'VI':6,'VII':7,'VIII':8,'IX':9,'X':10,
}
STRIP_CLASS_FRAGMENTS = [
    'nav','menu','header','footer','breadcrumb','sidebar','banner','cookie',
    'social','share','search','login','language','lang-','toolbar','zoom','pdf-icon','print',
]

# ── Fetch ──────────────────────────────────────────────────────────────────────

def fetch_html(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "text/html,application/xhtml+xml",
    }
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
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

def find_footnote_boundary(lines):
    # Explicit NOTES section marker
    for i, line in enumerate(lines):
        s = line.strip()
        if re.match(r'^\*\*NOTES?\*\*$', s, re.IGNORECASE) or s == '---':
            for j in range(i+1, min(i+5, len(lines))):
                if lines[j].strip():
                    if re.match(r'^\d{1,3}', lines[j].strip()):
                        return i
                    break

    # Fallback: numbered footnote lines after last paragraph
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
    """Split '**POPE JOHN PAUL II**TO THE BISHOPS' into separate parts."""
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
        for dt in DOCUMENT_TYPES:
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
    # State: None = normal, 'subtitle' = accumulating chapter subtitle lines
    subtitle_buf = None

    def flush():
        if current:
            raw = current.pop('_raw', '')
            current['paragraphs'] = parse_paragraphs(raw)
            chapters.append(current)

    def finalize_subtitle():
        """Apply accumulated subtitle buffer to current chapter."""
        nonlocal subtitle_buf
        if subtitle_buf is not None and current:
            sub = re.sub(r'\*', '', subtitle_buf).strip()
            if sub:
                current['subtitle'] = sub
                current['title'] = f"Chapter {current['number']}: {sub}"
        subtitle_buf = None

    for line in body_lines:
        s = line.strip()

        # ── Chapter number line ──────────────────────────────────────────────
        m = re.match(
            r'^CHAPTER\s+(ONE|TWO|THREE|FOUR|FIVE|SIX|SEVEN|EIGHT|NINE|TEN|'
            r'I{1,3}|IV|VI{0,3}|IX|X{1,3}|\d+)$', s, re.IGNORECASE
        )
        if m:
            finalize_subtitle()
            flush()
            word = m.group(1).upper()
            try: num = int(word)
            except: num = CHAPTER_WORDS.get(word, len(chapters)+1)
            current = {'number':num,'title':f'Chapter {num}','subtitle':'','_raw':''}
            subtitle_buf = ''   # start accumulating subtitle
            continue

        # ── Accumulate subtitle lines (blank lines OK, stop at paragraph) ───
        if subtitle_buf is not None:
            if not s:
                continue   # skip blank lines between CHAPTER N and title
            # If this looks like a paragraph (numbered) or a new chapter, stop
            if re.match(r'^\d+\.\s', s) or re.match(r'^CHAPTER\s+', s, re.IGNORECASE):
                finalize_subtitle()
                # Fall through to process this line normally below
            elif s.startswith('**') or not subtitle_buf or (subtitle_buf and not subtitle_buf.endswith('**')):
                # Accumulate: bold opening line OR continuation of multi-line bold title
                subtitle_buf = (subtitle_buf + ' ' + s).strip()
                # If the accumulated buffer now has matching ** pairs, subtitle is complete
                # Count ** occurrences: need at least opening and closing
                if subtitle_buf.endswith('**') and subtitle_buf.count('**') >= 2:
                    finalize_subtitle()
                continue
            else:
                # Non-bold line after a complete subtitle → subtitle is done
                finalize_subtitle()
                # Fall through

        # ── Standalone section: INTRODUCTION, CONCLUSION ─────────────────────
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
    footnotes, cur_num, cur_text = {}, None, []
    fn_start = re.compile(r'^\s*(\d{1,3})\s*(.*)')
    for line in lines:
        s = line.strip()
        if not s: continue
        if s in ('---','***') or re.match(r'^\*\*NOTES?\*\*$', s, re.IGNORECASE):
            continue
        m = fn_start.match(s)
        if m:
            rest = m.group(2).strip()
            if rest and not rest.startswith('.'):
                if cur_num is not None:
                    footnotes[cur_num] = clean_text(' '.join(cur_text))
                cur_num, cur_text = int(m.group(1)), [rest]
                continue
        if cur_num is not None:
            cur_text.append(s)
    if cur_num is not None:
        footnotes[cur_num] = clean_text(' '.join(cur_text))
    return footnotes

def clean_text(text):
    text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', text)
    text = re.sub(r'!\[[^\]]*\]\([^\)]+\)', '', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()

# ── Main Extract ──────────────────────────────────────────────────────────────

def fetch_and_extract(url):
    print(f"  Fetching HTML...")
    html = fetch_html(url)
    print(f"  Extracting text from HTML...")
    text = extract_text_from_html(html)
    return parse_text(text, url)

def parse_text(text, url=''):
    lines = text.split('\n')
    print(f"  Identifying document boundaries ({len(lines)} lines)...")
    start, end = find_content_boundaries(lines)
    content_lines = lines[start:end]
    fn_boundary = find_footnote_boundary(content_lines)
    body_lines = content_lines[:fn_boundary]
    footnote_lines = content_lines[fn_boundary:]
    header_end = 0
    for i, line in enumerate(body_lines):
        if re.match(r'^\s*\d+\.\s+\S', line) or re.match(r'^CHAPTER\s+', line.strip(), re.IGNORECASE):
            header_end = i; break
    print(f"  Parsing structure...")
    meta = parse_header(body_lines[:header_end])
    chapters = parse_chapters(body_lines[header_end:])
    footnotes = parse_footnotes(footnote_lines)
    total_p = sum(len(ch.get('paragraphs',[])) for ch in chapters)
    print(f"  → {len(chapters)} chapters, {total_p} paragraphs, {len(footnotes)} footnotes")
    return {**meta, 'url':url, 'chapters':chapters, 'footnotes':footnotes}

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
