#!/usr/bin/env python3
"""
BRUHsailer guide build script.

Downloads the latest .docx export of each chapter from Google Docs, parses
them, and splices the resulting GUIDE array into base.html to produce
index.html.

Reads:
  base.html                                            (site template)
  source/Chapter{1,2,3}.docx                           (downloaded automatically)

Writes:
  source/Chapter{1,2,3}.docx                           (refreshed from Google Docs)
  index.html                                           (rebuilt site)

Usage:
  python build.py                                      (fetch + build, default)
  python build.py --no-fetch                           (use existing local docx)
  python build.py --source source/ --base base.html --output index.html
"""

import argparse
import html as html_lib
import json
import re
import sys
import urllib.request
import urllib.error
from pathlib import Path

import docx
from docx.oxml.ns import qn


# ═══════════════════════════════════════════════════════════════════════════
# CONFIG — Google Doc IDs for each chapter
# ═══════════════════════════════════════════════════════════════════════════
# These are the document IDs from the public Google Docs URLs. To find an ID,
# open the doc in a browser; the URL looks like:
#   https://docs.google.com/document/d/<THIS_PART_IS_THE_ID>/edit
# The docs must be shared as "Anyone with the link can view" for fetching
# to work without authentication.

GOOGLE_DOC_IDS = {
    1: "1gCez5XG5FA1kmmBYydur3RaI_cr-dYNJlnigRrByEX8",
    2: "1YQiZ6curEYPpgm3DtjZcWHPoEEkGpYdXZ-I0gCM5p10",
    3: "1O1VeAkwS6VAzGVy0GT205GqiNaOAbw17H5uyuMwz39o",
}


# ═══════════════════════════════════════════════════════════════════════════
# GOOGLE DOCS FETCHER
# ═══════════════════════════════════════════════════════════════════════════

def fetch_docx(doc_id, dest_path):
    """Download a Google Doc as .docx using the public export URL.

    The doc must be set to 'Anyone with the link can view' (or more open).
    Raises RuntimeError if the download fails or the result isn't a .docx.
    """
    url = f"https://docs.google.com/document/d/{doc_id}/export?format=docx"
    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0 (compatible; BRUHsailer-build/1.0)",
    })
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = resp.read()
    except urllib.error.HTTPError as e:
        raise RuntimeError(
            f"HTTP {e.code} fetching doc {doc_id}. "
            f"Check the doc is shared as 'Anyone with the link can view'."
        )
    except urllib.error.URLError as e:
        raise RuntimeError(f"Network error fetching doc {doc_id}: {e.reason}")

    # A real .docx is a ZIP archive, which always starts with the bytes 'PK'.
    # If Google returns an HTML page instead (e.g. login redirect), we'd see
    # '<!DO' or similar.
    if not data.startswith(b'PK'):
        snippet = data[:120].decode('utf-8', errors='replace')
        raise RuntimeError(
            f"Downloaded data for doc {doc_id} is not a .docx file. "
            f"Most likely the doc is not publicly viewable.\n"
            f"First bytes received: {snippet!r}"
        )

    dest_path = Path(dest_path)
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    with open(dest_path, 'wb') as f:
        f.write(data)
    return len(data)


def fetch_all_chapters(source_dir):
    """Fetch all configured chapters into source_dir. Returns list of paths."""
    source_dir = Path(source_dir)
    paths = []
    for chapter_num, doc_id in GOOGLE_DOC_IDS.items():
        dest = source_dir / f"Chapter{chapter_num}.docx"
        print(f"Fetching Chapter {chapter_num} from Google Docs (id {doc_id[:12]}...)")
        size = fetch_docx(doc_id, dest)
        print(f"  -> {dest} ({size:,} bytes)")
        paths.append((chapter_num, dest))
    return paths


# ═══════════════════════════════════════════════════════════════════════════
# DOCX PARSER
# ═══════════════════════════════════════════════════════════════════════════

def get_ilvl(p):
    """Bullet indentation level, or None if not a bullet."""
    pPr = p._element.find(qn('w:pPr'))
    if pPr is None:
        return None
    numpr = pPr.find(qn('w:numPr'))
    if numpr is None:
        return None
    ilvl_el = numpr.find(qn('w:ilvl'))
    return int(ilvl_el.get(qn('w:val'))) if ilvl_el is not None else 0


def get_links_map(doc):
    """Map relationship id -> URL for hyperlinks."""
    out = {}
    for rel_id, rel in doc.part.rels.items():
        if 'hyperlink' in rel.reltype:
            out[rel_id] = rel.target_ref
    return out


def _run_to_html(r):
    """Convert a w:r run to HTML, preserving bold and italic."""
    text_parts = []
    for t in r.findall(qn('w:t')):
        if t.text:
            text_parts.append(t.text)
    text = ''.join(text_parts)
    if not text:
        return ''
    rPr = r.find(qn('w:rPr'))
    is_bold = False
    is_italic = False
    if rPr is not None:
        b = rPr.find(qn('w:b'))
        if b is not None and b.get(qn('w:val')) not in ('0', 'false'):
            is_bold = True
        i = rPr.find(qn('w:i'))
        if i is not None and i.get(qn('w:val')) not in ('0', 'false'):
            is_italic = True
    escaped = html_lib.escape(text)
    if is_bold:
        escaped = f'<strong>{escaped}</strong>'
    if is_italic:
        escaped = f'<em>{escaped}</em>'
    return escaped


def runs_to_html(p, links_map):
    """Walk paragraph children in order, preserving bold/italic and hyperlinks."""
    parts = []
    for child in p._element:
        tag = child.tag.split('}')[-1]
        if tag == 'hyperlink':
            rel_id = child.get(qn('r:id'))
            url = links_map.get(rel_id, '')
            inner = ''.join(_run_to_html(r) for r in child.findall(qn('w:r')))
            if inner:
                parts.append(
                    f'<a href="{html_lib.escape(url)}" target="_blank" rel="noopener">{inner}</a>'
                )
        elif tag == 'r':
            parts.append(_run_to_html(child))
    return ''.join(parts)


def paragraph_is_italic(p):
    runs = [r for r in p.runs if r.text and r.text.strip()]
    if not runs:
        return False
    return all(bool(r.italic) for r in runs)


# ─── Sentence splitting (paren- and abbreviation-aware) ────────────────────

ABBREVS = ('e.g', 'i.e', 'vs', 'Mr', 'Mrs', 'Dr', 'St', 'cf', 'etc', 'Inc', 'Ltd')


def split_sentences(text):
    """Split text into sentences, respecting parens and common abbreviations.
    HTML-aware: doesn't split inside tags."""
    if not text:
        return []
    sentences = []
    buf = []
    paren_depth = 0
    in_tag = False
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if ch == '<':
            in_tag = True
            buf.append(ch); i += 1; continue
        if ch == '>':
            in_tag = False
            buf.append(ch); i += 1; continue
        if in_tag:
            buf.append(ch); i += 1; continue
        if ch == '(':
            paren_depth += 1
            buf.append(ch); i += 1; continue
        if ch == ')':
            if paren_depth > 0:
                paren_depth -= 1
            buf.append(ch); i += 1; continue
        if ch in '.!?':
            if paren_depth > 0:
                buf.append(ch); i += 1; continue
            current = ''.join(buf)
            if ch == '.':
                if i > 0 and text[i-1].isdigit() and i + 1 < n and text[i+1].isdigit():
                    buf.append(ch); i += 1; continue
                m = re.search(r'(\b\w+)$', current)
                if m and m.group(1) in ABBREVS:
                    buf.append(ch); i += 1; continue
            buf.append(ch)
            j = i + 1
            while j < n and text[j] in ' \t':
                j += 1
            if j >= n:
                sentences.append(''.join(buf).strip())
                buf = []
                i = j
                continue
            next_ch = text[j]
            if j > i + 1 and (next_ch.isupper() or next_ch.isdigit() or next_ch in '"\'<('):
                sentences.append(''.join(buf).strip())
                buf = []
                i = j
                continue
            i += 1
            continue
        buf.append(ch); i += 1
    if buf:
        last = ''.join(buf).strip()
        if last:
            sentences.append(last)
    return [s for s in sentences if s]


# ─── Metadata classification ───────────────────────────────────────────────

META_PATTERNS = {
    'gp':     re.compile(r'^\s*GP\s*(stack|after)[^:]*:\s*(.*)$', re.IGNORECASE),
    'items':  re.compile(r'^\s*Items\s*needed[^:]*:\s*(.*)$', re.IGNORECASE),
    'skills': re.compile(r'^\s*Skills\s*/\s*quests[^:]*:\s*(.*)$', re.IGNORECASE),
    'time':   re.compile(r'^\s*(Total\s+time|total\s+time)[^:]*:\s*(.*)$', re.IGNORECASE),
}
SECTION_REJECT = re.compile(r'(start of section|for section|of section)', re.IGNORECASE)


def classify_meta_line(text):
    if not text:
        return None
    colon_idx = text.find(':')
    if colon_idx > 0:
        if SECTION_REJECT.search(text[:colon_idx]):
            return None
    for key, pat in META_PATTERNS.items():
        m = pat.match(text)
        if m:
            return (key, m.groups()[-1].strip())
    return None


def classify_meta(text):
    """Classify a paragraph (possibly multi-line) as step metadata."""
    if not text:
        return None
    out = {}
    for line in text.split('\n'):
        line = line.strip()
        if not line:
            continue
        result = classify_meta_line(line)
        if result:
            key, value = result
            out[key] = value
    return out if out else None


def is_section_header(text):
    return bool(re.match(r'^\s*\d+\.\d+\s*:\s*\S', text))


# ─── Per-chapter parse ─────────────────────────────────────────────────────

def parse_chapter(path, chapter_num):
    doc = docx.Document(str(path))
    links_map = get_links_map(doc)
    paras = doc.paragraphs

    chapter_title = None
    intro_paras = []
    section_data = []
    current_section = None
    current_steps = []
    step_bullets = []
    step_meta = {}

    def close_step():
        nonlocal step_bullets, step_meta
        # Only emit if we have BOTH bullets AND metadata. This drops trailing
        # decorative content (end-of-chapter stats summaries, resource lists).
        if step_bullets and step_meta:
            current_steps.append({'bullets': step_bullets, 'meta': step_meta})
        step_bullets = []
        step_meta = {}

    def open_section(name):
        nonlocal current_section, current_steps
        if current_section is not None:
            section_data.append((current_section, current_steps))
        current_section = name
        current_steps = []

    for p in paras:
        text = p.text
        if not text or not text.strip():
            continue
        if chapter_title is None:
            chapter_title = text.strip()
            continue
        if is_section_header(text):
            close_step()
            open_section(text.strip())
            continue
        meta_dict = classify_meta(text)
        if meta_dict:
            step_meta.update(meta_dict)
            continue
        ilvl = get_ilvl(p)
        if ilvl is not None:
            if step_meta and step_bullets:
                close_step()
            html_body = runs_to_html(p, links_map)
            italic = paragraph_is_italic(p)
            step_bullets.append((ilvl, html_body, italic))
            continue
        # Non-bullet, non-meta paragraph: treat as barrier.
        if step_bullets and not step_meta:
            step_bullets = []
        if current_section is None:
            intro_paras.append(text.strip())

    close_step()
    if current_section is not None:
        section_data.append((current_section, current_steps))

    return chapter_title, intro_paras, section_data


# ─── Title and content rendering ───────────────────────────────────────────

def first_sentence(html_text, max_len=160):
    plain = re.sub(r'<[^>]+>', '', html_text)
    plain = html_lib.unescape(plain).strip()
    sents = split_sentences(plain)
    if not sents:
        return plain[:max_len]
    s = sents[0].rstrip()
    if len(s) > max_len:
        s = s[:max_len].rstrip()
    return s


def make_title(step_num, bullets):
    if not bullets:
        return f"Step {step_num}"
    first_title = first_sentence(bullets[0][1])
    # If first bullet is a short label ending in ':', append the next bullet
    if len(first_title) < 30 and first_title.rstrip().endswith(':') and len(bullets) > 1:
        second_title = first_sentence(bullets[1][1])
        if second_title:
            first_title = f"{first_title} {second_title}"
            if len(first_title) > 160:
                first_title = first_title[:160].rstrip()
    return f"Step {step_num}: {first_title}"


def render_bullet_html(html_body, do_split=True):
    if not do_split:
        return [html_body]
    sents = split_sentences(html_body)
    if not sents:
        return [html_body]
    return [s.strip() for s in sents if s.strip()]


def render_step_content(bullets, meta):
    lis_top = []
    current_top = None
    current_children = []

    for ilvl, body, italic in bullets:
        if ilvl == 0:
            if current_top is not None:
                lis_top.append((current_top, current_children))
            pieces = render_bullet_html(body, do_split=True)
            if italic:
                pieces = [f'<em>{p}</em>' if not p.startswith('<em>') else p for p in pieces]
            for piece in pieces[:-1]:
                lis_top.append((piece, []))
            current_top = pieces[-1] if pieces else body
            current_children = []
        else:
            child_html = body
            if italic and not child_html.startswith('<em>'):
                child_html = f'<em>{child_html}</em>'
            current_children.append((ilvl, child_html))
    if current_top is not None:
        lis_top.append((current_top, current_children))

    out = ['<ul class="main-steps">']
    for top_html, children in lis_top:
        if children:
            child_parts = ['<ul class="sub-steps">']
            i = 0
            while i < len(children):
                clvl, ctext = children[i]
                if clvl == 1:
                    nested = []
                    j = i + 1
                    while j < len(children) and children[j][0] >= 2:
                        nested.append(children[j])
                        j += 1
                    if nested:
                        child_parts.append(f'<li>{ctext}<ul class="sub-steps">')
                        for _, ntext in nested:
                            child_parts.append(f'<li>{ntext}</li>')
                        child_parts.append('</ul></li>')
                    else:
                        child_parts.append(f'<li>{ctext}</li>')
                    i = j
                else:
                    child_parts.append(f'<li>{ctext}</li>')
                    i += 1
            child_parts.append('</ul>')
            out.append(f'<li>{top_html}{"".join(child_parts)}</li>')
        else:
            out.append(f'<li>{top_html}</li>')
    out.append('</ul>')

    meta_parts = []
    if meta.get('gp'):
        meta_parts.append(f'💰 <strong>GP after step:</strong> {html_lib.escape(meta["gp"])}')
    if meta.get('items'):
        meta_parts.append(f'🎒 <strong>Items needed:</strong> {html_lib.escape(meta["items"])}')
    if meta.get('time'):
        meta_parts.append(f'⏱ <strong>Total time:</strong> {html_lib.escape(meta["time"])}')
    if meta_parts:
        out.append(f'<div class="step-meta">{"<br>".join(meta_parts)}</div>')
    return '\n'.join(out)


# ─── Final data structure assembly ─────────────────────────────────────────

def build_js_data(chapters_info):
    js_chapters = []
    for ch_num, ch_title, sections, _intro in chapters_info:
        section_objs = []
        global_step_idx = 0
        for sec_idx, (sec_name, steps) in enumerate(sections, start=1):
            step_objs = []
            for step in steps:
                global_step_idx += 1
                step_id = f"{ch_num}-{global_step_idx}"
                title = make_title(global_step_idx, step['bullets'])
                content_html = render_step_content(step['bullets'], step['meta'])
                step_objs.append({
                    'id': step_id,
                    'title': title,
                    'tags': [],
                    'content': content_html,
                })
            section_objs.append({
                'id': f"s{ch_num}-{sec_idx}",
                'name': sec_name,
                'steps': step_objs,
            })
        js_chapters.append({
            'chapter': f"Chapter {ch_num}",
            'chapterDesc': ch_title,
            'sections': section_objs,
        })
    return js_chapters


# ═══════════════════════════════════════════════════════════════════════════
# SPLICER (inject GUIDE array into base.html)
# ═══════════════════════════════════════════════════════════════════════════

def js_template_escape(s):
    return s.replace('\\', '\\\\').replace('`', '\\`').replace('${', '\\${')


def build_guide_js(data):
    js_parts = ['const GUIDE = [']
    for ch in data:
        js_parts.append('  {')
        js_parts.append(f'    chapter: "{ch["chapter"]}",')
        js_parts.append(f'    chapterDesc: {json.dumps(ch["chapterDesc"], ensure_ascii=False)},')
        js_parts.append('    sections: [')
        for sec in ch['sections']:
            js_parts.append('      {')
            js_parts.append(f'        id: "{sec["id"]}",')
            js_parts.append(f'        name: {json.dumps(sec["name"], ensure_ascii=False)},')
            js_parts.append('        steps: [')
            for step in sec['steps']:
                title_esc = js_template_escape(step['title'])
                content_esc = js_template_escape(step['content'])
                js_parts.append('          {')
                js_parts.append(f'            id: "{step["id"]}",')
                js_parts.append(f'            title: `{title_esc}`,')
                js_parts.append('            tags: [],')
                js_parts.append(f'            content: `{content_esc}`,')
                js_parts.append('          },')
            js_parts.append('        ]')
            js_parts.append('      },')
        js_parts.append('    ]')
        js_parts.append('  },')
    js_parts.append('];')
    return '\n'.join(js_parts)


def splice_into_base(data, base_path, output_path):
    with open(base_path, 'r', encoding='utf-8') as f:
        html = f.read()

    m = re.search(r'const GUIDE\s*=\s*\[', html)
    if not m:
        raise SystemExit("ERROR: could not find 'const GUIDE = [' in base.html")
    start_idx = m.start()

    # Find matching closing ];
    depth = 0
    in_str = None
    escape = False
    i = m.end() - 1
    n = len(html)
    end_idx = -1
    while i < n:
        c = html[i]
        if escape:
            escape = False; i += 1; continue
        if in_str:
            if c == '\\':
                escape = True
            elif c == in_str:
                in_str = None
            i += 1; continue
        if c in '"\'`':
            in_str = c; i += 1; continue
        if c == '[':
            depth += 1
        elif c == ']':
            depth -= 1
            if depth == 0:
                end_idx = i + 1
                j = end_idx
                while j < n and html[j] in ' \t':
                    j += 1
                if j < n and html[j] == ';':
                    end_idx = j + 1
                break
        i += 1

    if end_idx < 0:
        raise SystemExit("ERROR: could not find end of GUIDE array in base.html")

    new_guide_js = build_guide_js(data)
    new_html = html[:start_idx] + new_guide_js + html[end_idx:]

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(new_html)


# ═══════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════

def find_chapter_docx(source_dir, chapter_num):
    """Locate the docx for a given chapter. Tries exact 'ChapterN.docx' first,
    then any file with 'ChapterN' in the name (case-insensitive)."""
    source_dir = Path(source_dir)
    exact = source_dir / f'Chapter{chapter_num}.docx'
    if exact.exists():
        return exact
    matches = sorted(source_dir.glob(f'*[Cc]hapter{chapter_num}*.docx'))
    if matches:
        # Prefer the most recent by name (works well for ISO-date prefixes)
        return matches[-1]
    return None


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--source', default='source',
                    help='Directory containing the Chapter1/2/3 .docx files (default: source)')
    ap.add_argument('--base', default='base.html',
                    help='Base HTML template (default: base.html)')
    ap.add_argument('--output', default='index.html',
                    help='Output HTML file (default: index.html)')
    ap.add_argument('--no-fetch', action='store_true',
                    help='Skip downloading from Google Docs; use existing local docx files')
    args = ap.parse_args()

    # Download latest .docx from Google Docs (unless --no-fetch)
    if args.no_fetch:
        print("Skipping fetch; using existing docx files.")
        chapter_paths = []
        missing = []
        for n in (1, 2, 3):
            path = find_chapter_docx(args.source, n)
            if path is None:
                missing.append(n)
            else:
                chapter_paths.append((n, path))
        if missing:
            print(f"ERROR: Could not find docx for chapter(s): {missing}", file=sys.stderr)
            print(f"  Expected '{args.source}/ChapterN.docx' or '{args.source}/*ChapterN*.docx'",
                  file=sys.stderr)
            sys.exit(1)
    else:
        try:
            chapter_paths = fetch_all_chapters(args.source)
        except RuntimeError as e:
            print(f"ERROR: {e}", file=sys.stderr)
            sys.exit(1)

    # Parse
    chapters_info = []
    for n, path in chapter_paths:
        print(f"Parsing {path}...")
        title, intro, sections = parse_chapter(path, n)
        n_steps = sum(len(steps) for _, steps in sections)
        print(f"  Chapter {n}: {len(sections)} sections, {n_steps} steps")
        chapters_info.append((n, title, sections, intro))

    # Build JS data and splice into base
    data = build_js_data(chapters_info)
    total = sum(len(sec['steps']) for ch in data for sec in ch['sections'])
    print(f"Total: {total} steps")

    splice_into_base(data, args.base, args.output)
    out_size = Path(args.output).stat().st_size
    print(f"Wrote {args.output} ({out_size:,} bytes)")


if __name__ == '__main__':
    main()
