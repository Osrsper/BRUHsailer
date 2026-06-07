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
import urllib.parse
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
    # Cache-busting: Google's export endpoint sits behind a CDN/edge cache that
    # can serve a stale document body even after edits (while metadata like the
    # title updates sooner). A unique query param forces a cache miss, and the
    # no-cache headers ask any intermediary to revalidate, so we get the latest.
    import time as _time
    cache_bust = int(_time.time() * 1000)
    url = (f"https://docs.google.com/document/d/{doc_id}/export"
           f"?format=docx&_cb={cache_bust}")
    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0 (compatible; BRUHsailer-build/1.0)",
        "Cache-Control": "no-cache, no-store, max-age=0",
        "Pragma": "no-cache",
    })
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = resp.read()
            # The Content-Disposition filename is the Google Doc's title, e.g.
            # 'attachment; filename="20260525Chapter1.docx"'. We use the leading
            # YYYYMMDD in that title as the "last updated" date for the guide.
            disposition = resp.headers.get('Content-Disposition', '') or ''
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

    # Extract the doc title from the Content-Disposition header (handles both
    # plain filename="..." and RFC 5987 filename*=UTF-8''... forms).
    doc_title = ''
    m = re.search(r"filename\*=(?:UTF-8'')?([^;\r\n]+)", disposition, re.IGNORECASE)
    if not m:
        m = re.search(r'filename="?([^";\r\n]+)"?', disposition, re.IGNORECASE)
    if m:
        doc_title = urllib.parse.unquote(m.group(1)).strip()

    dest_path = Path(dest_path)
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    with open(dest_path, 'wb') as f:
        f.write(data)
    return len(data), doc_title


def parse_title_date(title):
    """Return a datetime.date parsed from a leading YYYYMMDD in a doc title,
    or None if not present/parseable."""
    import datetime
    m = re.match(r'\s*(\d{4})(\d{2})(\d{2})', title or '')
    if not m:
        return None
    try:
        return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def fetch_all_chapters(source_dir):
    """Fetch all configured chapters into source_dir.

    Returns (paths, latest_date) where paths is a list of (chapter_num, path)
    and latest_date is the most recent date parsed from the doc titles, or None.
    """
    source_dir = Path(source_dir)
    paths = []
    dates = []
    for chapter_num, doc_id in GOOGLE_DOC_IDS.items():
        dest = source_dir / f"Chapter{chapter_num}.docx"
        print(f"Fetching Chapter {chapter_num} from Google Docs (id {doc_id[:12]}...)")
        size, title = fetch_docx(doc_id, dest)
        d = parse_title_date(title)
        date_note = f", title '{title}' -> {d}" if title else ""
        print(f"  -> {dest} ({size:,} bytes){date_note}")
        if d:
            dates.append(d)
        paths.append((chapter_num, dest))
    latest_date = max(dates) if dates else None
    return paths, latest_date


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


# Colors that should NOT be reproduced as body text:
#   - black / auto / default        -> normal text
#   - 8d1d75 (purple)               -> BRUH's metadata-label colour, handled
#                                      structurally elsewhere; would turn text
#                                      purple if reproduced
#   - common hyperlink blues        -> links are styled by the <a> rule
COLOR_SKIP = {
    None, 'auto', '000000', '8d1d75',
    '0000ee', '1155cc', '000080', '0563c1',
}


def _brighten_for_dark_bg(hex6):
    """Ensure a colour is light enough to read on the dark parchment theme by
    enforcing a minimum HSL lightness while preserving hue and saturation."""
    try:
        r = int(hex6[0:2], 16) / 255.0
        g = int(hex6[2:4], 16) / 255.0
        b = int(hex6[4:6], 16) / 255.0
    except (ValueError, IndexError):
        return '#' + hex6
    mx, mn = max(r, g, b), min(r, g, b)
    l = (mx + mn) / 2.0
    MIN_L = 0.58
    if l >= MIN_L:
        return '#' + hex6
    # Scale toward white to lift lightness to the floor
    if l == 0:
        return '#' + hex6
    factor = MIN_L / l
    nr = min(1.0, r * factor)
    ng = min(1.0, g * factor)
    nb = min(1.0, b * factor)
    return '#%02x%02x%02x' % (round(nr * 255), round(ng * 255), round(nb * 255))


def _run_to_html(r):
    """Convert a w:r run to HTML, preserving bold, italic, underline and colour."""
    text_parts = []
    for t in r.findall(qn('w:t')):
        if t.text:
            text_parts.append(t.text)
    text = ''.join(text_parts)
    if not text:
        return ''
    rPr = r.find(qn('w:rPr'))
    is_bold = is_italic = is_underline = False
    color = None
    if rPr is not None:
        b = rPr.find(qn('w:b'))
        if b is not None and b.get(qn('w:val')) not in ('0', 'false'):
            is_bold = True
        i = rPr.find(qn('w:i'))
        if i is not None and i.get(qn('w:val')) not in ('0', 'false'):
            is_italic = True
        u = rPr.find(qn('w:u'))
        if u is not None and u.get(qn('w:val')) not in (None, 'none', '0', 'false'):
            is_underline = True
        c = rPr.find(qn('w:color'))
        if c is not None:
            val = (c.get(qn('w:val')) or '').lower()
            if val and val not in COLOR_SKIP:
                color = val
    escaped = html_lib.escape(text)
    if color:
        escaped = f'<span style="color:{_brighten_for_dark_bg(color)}">{escaped}</span>'
    if is_underline:
        escaped = f'<u>{escaped}</u>'
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


EOC_MARKER = re.compile(r'^=*\s*end of chapter', re.I)


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
    eoc_entries = []   # raw end-of-chapter content (after the ==END OF CHAPTER== marker)
    in_eoc = False

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
        # Detect the end-of-chapter marker. Everything after it is EOC content,
        # not steps.
        if not in_eoc and EOC_MARKER.match(text.strip()):
            close_step()
            if current_section is not None:
                section_data.append((current_section, current_steps))
                current_section = None
            in_eoc = True
            continue
        if in_eoc:
            eoc_entries.append({
                'ilvl': get_ilvl(p),
                'html': runs_to_html(p, links_map),
                'text': text,   # keep unstripped so leading tabs are preserved
            })
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

    return chapter_title, intro_paras, section_data, eoc_entries


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


_TAG_RE = re.compile(r'<(/?)([a-zA-Z][a-zA-Z0-9]*)([^>]*)>')
_VOID_TAGS = {'br', 'img', 'hr', 'wbr'}


def _balance_fragments(pieces):
    """Given HTML fragments produced by splitting a larger HTML string at
    sentence boundaries, return fragments that are each independently
    tag-balanced. Inline tags left open at the end of a fragment are closed,
    and reopened (with their original attributes) at the start of the next."""
    out = []
    carry = []  # list of (tagname, full_open_tag) still open from prior fragment
    for piece in pieces:
        prefix = ''.join(tag for _, tag in carry)
        stack = list(carry)
        for m in _TAG_RE.finditer(piece):
            closing, name, attrs = m.group(1), m.group(2).lower(), m.group(3)
            if name in _VOID_TAGS or attrs.rstrip().endswith('/'):
                continue
            if closing:
                for k in range(len(stack) - 1, -1, -1):
                    if stack[k][0] == name:
                        stack.pop(k)
                        break
            else:
                stack.append((name, m.group(0)))
        suffix = ''.join(f'</{name}>' for name, _ in reversed(stack))
        out.append(prefix + piece + suffix)
        carry = stack
    return out


def _build_bullet_tree(bullets):
    """Turn a flat list of (ilvl, html, italic) into a nested tree.

    Each node: {'html', 'italic', 'children': [...]}. Handles arbitrary depth
    by tracking a stack of open nodes keyed on indentation level.
    """
    root = []
    stack = []  # list of (ilvl, node)
    for ilvl, body, italic in bullets:
        node = {'html': body, 'italic': italic, 'children': []}
        while stack and stack[-1][0] >= ilvl:
            stack.pop()
        if stack:
            stack[-1][1]['children'].append(node)
        else:
            root.append(node)
        stack.append((ilvl, node))
    return root


def _render_nodes(nodes, css_class):
    """Render a list of tree nodes to a <ul>. Each node's text is split into
    sentences (so multi-sentence bullets become separate points, at any depth),
    fragments are tag-balanced, and children nest under the node's last sentence."""
    out = [f'<ul class="{css_class}">']
    for node in nodes:
        pieces = render_bullet_html(node['html'], do_split=True)
        pieces = _balance_fragments(pieces)
        if node['italic']:
            pieces = [p if '<em>' in p else f'<em>{p}</em>' for p in pieces]
        if not pieces:
            pieces = [node['html']]
        child_html = ''
        if node['children']:
            child_html = _render_nodes(node['children'], 'sub-steps')
        for idx, piece in enumerate(pieces):
            is_last = (idx == len(pieces) - 1)
            if is_last and child_html:
                out.append(f'<li>{piece}{child_html}</li>')
            else:
                out.append(f'<li>{piece}</li>')
    out.append('</ul>')
    return ''.join(out)


def render_step_content(bullets, meta):
    tree = _build_bullet_tree(bullets)
    out = [_render_nodes(tree, 'main-steps')]

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


# ═══════════════════════════════════════════════════════════════════════════
# END-OF-CHAPTER (EOC) PARSING + RENDERING
# ═══════════════════════════════════════════════════════════════════════════

# Map the many skill-name spellings BRUH uses to the short grid labels.
SKILL_ABBREV = {
    'attack': 'Atk', 'atk': 'Atk',
    'strength': 'Str', 'str': 'Str',
    'defence': 'Def', 'def': 'Def',
    'hitpoints': 'HP', 'hp': 'HP',
    'ranged': 'Range', 'range': 'Range',
    'prayer': 'Prayer', 'pray': 'Prayer',
    'magic': 'Magic',
    'runecraft': 'RC', 'runecrafting': 'RC', 'rc': 'RC',
    'construction': 'Cons', 'cons': 'Cons',
    'agility': 'Agil', 'agil': 'Agil',
    'herblore': 'Herb', 'herb': 'Herb',
    'thieving': 'Thiev', 'thiev': 'Thiev',
    'crafting': 'Craft', 'craft': 'Craft',
    'fletching': 'Fletch', 'fletch': 'Fletch',
    'slayer': 'Slay', 'slay': 'Slay',
    'hunter': 'Hunt', 'hunt': 'Hunt',
    'mining': 'Mining',
    'smithing': 'Smith', 'smith': 'Smith',
    'fishing': 'Fish', 'fish': 'Fish',
    'cooking': 'Cook', 'cook': 'Cook',
    'firemaking': 'FM', 'fm': 'FM',
    'woodcutting': 'WC', 'wc': 'WC',
    'farming': 'Farm', 'farm': 'Farm',
    'sailing': 'Sail', 'sail': 'Sail',
}

_STAT_LINE = re.compile(r'^\s*([A-Za-z][A-Za-z ]*?)\s*:\s*(\d+)\s*(.*)$')

_EOC_SECTION_HEADERS = [
    (re.compile(r'^quests missing', re.I), 'list', 'Quests missing for quest cape:'),
    (re.compile(r'^miniquests missing', re.I), 'list', 'Miniquests missing:'),
    (re.compile(r'^landlubber version', re.I), 'links', 'Landlubber version (pre sailing):'),
    (re.compile(r'^old version', re.I), 'links', 'OLD VERSION:'),
]


def _is_changelog_line(text):
    t = text.strip()
    if re.match(r'^beware compatibility', t, re.I):
        return True
    if re.search(r'\.txt\s*$', t, re.I):
        return True
    if re.match(r'^\d{8}\s*:', t):
        return True
    if re.match(r'^\d{8}\S*change\s*log', t, re.I):
        return True
    return False


def parse_eoc(entries):
    """Parse the raw end-of-chapter entries into structured data."""
    data = {'intro': '', 'stats': [], 'notes': [], 'total_level': None,
            'sections': [], 'routes': []}
    n = len(entries)
    i = 0

    # Optional intro / "Stats:" header
    if i < n:
        first = entries[i]['text'].strip()
        if first.lower().startswith('stats'):
            if first.rstrip().lower() in ('stats:', 'stats'):
                data['intro'] = 'Stats at completion of this chapter:'
            else:
                data['intro'] = first if first.endswith(':') else first + ':'
            i += 1

    # Stats block: consecutive "Skill: value (note)" lines
    while i < n:
        text = entries[i]['text'].strip()
        m = _STAT_LINE.match(text)
        if not m:
            break
        abbrev = SKILL_ABBREV.get(m.group(1).strip().lower())
        if not abbrev:
            break
        value = m.group(2)
        note = m.group(3).strip()
        data['stats'].append((abbrev, value))
        if note:
            disp = note
            if disp.startswith('(') and disp.endswith(')') and disp.count('(') == 1:
                disp = disp[1:-1].strip()
            data['notes'].append((abbrev, disp))
        i += 1

    # Remaining content: sections, references, changelog (stripped), routes
    section = None
    beyond_section = None
    seen_changelog = False
    routes = []
    while i < n:
        entry = entries[i]
        raw = entry['text']
        stripped = raw.strip()
        i += 1
        if not stripped:
            continue

        mtl = re.match(r'^total level\s*:\s*(.*)$', stripped, re.I)
        if mtl:
            data['total_level'] = mtl.group(1).strip()
            continue

        if _is_changelog_line(stripped):
            seen_changelog = True
            continue

        matched = False
        for pat, kind, heading in _EOC_SECTION_HEADERS:
            if pat.match(stripped):
                section = {'kind': kind, 'heading': heading, 'items': []}
                data['sections'].append(section)
                matched = True
                break
        if matched:
            continue

        if re.match(r'^overview document\s*:', stripped, re.I):
            section = {'kind': 'links', 'heading': '📎 References &amp; Links', 'items': []}
            data['sections'].append(section)
            section['items'].append(entry['html'])
            continue

        # After the changelog block, any remaining content is the Ch3
        # per-skill "routes to 99" material.
        if seen_changelog:
            leading_ws = raw[:1] in ('\t', ' ')
            # A skill header is a short, non-indented line (e.g. "Ranged",
            # "Magic+Defence"). Longer non-indented lines are route continuations.
            is_skill_header = (not leading_ws) and len(stripped) < 35
            routes.append({'route': not is_skill_header, 'text': stripped})
            continue

        # Orphan content before any known section = the Ch3 "Beyond BRUHsailer"
        # closing block (an intro line plus resource links with descriptions).
        if section is None:
            if beyond_section is None:
                beyond_section = {'kind': 'beyond',
                                  'heading': '🧭 Beyond BRUHsailer',
                                  'items': []}
                data['sections'].append(beyond_section)
            beyond_section['items'].append({
                'ilvl': entry.get('ilvl'),
                'html': entry['html'],
            })
            continue

        # Otherwise it's a content line for the current section
        section['items'].append(entry['html'])

    data['routes'] = routes
    return data


def render_eoc_html(data, chapter_num):
    emoji = '🏆' if chapter_num == 3 else '📊'
    p = []
    p.append(f'<div class="chapter-summary" id="eoc-chapter-{chapter_num}" '
             f'style="scroll-margin-top:var(--header-h,120px)">')
    p.append('  <div class="eoc-header" onclick="toggleEOCCollapsed(this)">')
    p.append(f'    <h3>{emoji} End of Chapter {chapter_num}</h3>'
             f'<span class="eoc-chevron">▼</span>')
    p.append('  </div>')
    p.append('  <div class="eoc-body">')
    # "Mark chapter complete" helper (a site feature, not from the doc)
    p.append('  <div class="eoc-complete-block">')
    p.append('    <div class="eoc-complete-label">⚡ Hopping over from an older '
             'version of BRUHsailer?</div>')
    p.append(f'    <button class="eoc-complete-btn" onclick="completeChapter({chapter_num})">'
             f'Mark all of Chapter {chapter_num} as complete</button>')
    p.append('    <div class="eoc-complete-hint">Useful if you already completed this '
             'chapter on the old site. You can untick individual steps later.</div>')
    p.append('  </div>')

    if data['intro']:
        p.append(f'  <p class="eoc-intro">{html_lib.escape(data["intro"])}</p>')

    if data['stats']:
        cells = ''.join(
            f'<div><span>{html_lib.escape(lbl)}</span>{html_lib.escape(val)}</div>'
            for lbl, val in data['stats'])
        p.append(f'  <div class="stats-grid">{cells}</div>')

    if data['total_level']:
        p.append(f'  <div class="eoc-line"><strong>Total level:</strong> '
                 f'{html_lib.escape(data["total_level"])}</div>')

    if data['notes']:
        notes = ''.join(
            f'<li><strong>{html_lib.escape(lbl)}:</strong> {html_lib.escape(note)}</li>'
            for lbl, note in data['notes'])
        p.append('  <h4 class="eoc-subheading">📝 Skill Notes</h4>')
        p.append(f'  <ul class="eoc-skill-notes">{notes}</ul>')

    for sec in data['sections']:
        if sec['kind'] == 'beyond':
            p.append(f'  <h4 class="eoc-subheading">{sec["heading"]}</h4>')
            p.append('  <div class="eoc-beyond">')
            first = True
            for it in sec['items']:
                ilvl = it.get('ilvl')
                ihtml = it['html']
                is_link_header = ihtml.lstrip().startswith('<a ')
                if first:
                    # Opening prose for the section
                    p.append(f'    <p class="eoc-beyond-intro">{ihtml}</p>')
                    first = False
                elif is_link_header and ilvl is None:
                    p.append(f'    <div class="eoc-beyond-item">{ihtml}</div>')
                else:
                    p.append(f'    <div class="eoc-beyond-desc">{ihtml}</div>')
            p.append('  </div>')
            continue
        p.append(f'  <h4 class="eoc-subheading">{sec["heading"]}</h4>')
        if sec['kind'] == 'list':
            items = ''.join(f'<li>{it}</li>' for it in sec['items'])
            p.append(f'  <ul class="eoc-list">{items}</ul>')
        else:
            for it in sec['items']:
                p.append(f'  <div class="eoc-line">{it}</div>')

    if data['routes']:
        p.append('  <h4 class="eoc-subheading">🎯 Skilling routes to 99</h4>')
        p.append('  <div class="eoc-routes">')
        for r in data['routes']:
            if r['route']:
                segs = [s for s in r['text'].split('\t') if s.strip()]
                if len(segs) >= 2:
                    rng = html_lib.escape(segs[0].strip())
                    method = html_lib.escape(' '.join(s.strip() for s in segs[1:]))
                    p.append(f'    <div class="eoc-route-line">'
                             f'<span class="route-range">{rng}</span> {method}</div>')
                else:
                    p.append(f'    <div class="eoc-route-line">'
                             f'{html_lib.escape(r["text"])}</div>')
            else:
                p.append(f'    <div class="eoc-route-skill">'
                         f'<strong>{html_lib.escape(r["text"])}</strong></div>')
        p.append('  </div>')

    p.append('  </div>')
    p.append('</div>')
    return '\n'.join(p)


def build_chapter_summaries_js(eoc_html_by_chapter):
    """Build the `const CHAPTER_SUMMARIES = {...};` JS block."""
    lines = ['const CHAPTER_SUMMARIES = {']
    for ch_num in sorted(eoc_html_by_chapter):
        html = eoc_html_by_chapter[ch_num]
        escaped = js_template_escape(html)
        lines.append(f'  "Chapter {ch_num}": `{escaped}`,')
    lines.append('};')
    return '\n'.join(lines)


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


def splice_into_base(data, base_path, output_path, last_updated_override=None,
                     chapter_summaries_js=None):
    with open(base_path, 'r', encoding='utf-8') as f:
        html = f.read()

    # Replace the GUIDE array
    g_start, g_end = _find_block_bounds(html, r'const GUIDE\s*=\s*\[', '[', ']')
    if g_start is None:
        raise SystemExit("ERROR: could not find 'const GUIDE = [' in base.html")
    new_guide_js = build_guide_js(data)
    new_html = html[:g_start] + new_guide_js + html[g_end:]

    # Replace the CHAPTER_SUMMARIES object (auto-generated end-of-chapter notes)
    if chapter_summaries_js:
        s_start, s_end = _find_block_bounds(
            new_html, r'const CHAPTER_SUMMARIES\s*=\s*\{', '{', '}')
        if s_start is None:
            raise SystemExit("ERROR: could not find 'const CHAPTER_SUMMARIES = {' in base.html")
        new_html = new_html[:s_start] + chapter_summaries_js + new_html[s_end:]

    # ── Determine "last updated" date ──────────────────────────────────────
    # Preferred: the date from the Google Doc titles (e.g. 20260525Chapter1),
    # which BRUH maintains precisely. Falls back to a content-change date only
    # if no title date was available (e.g. local --no-fetch builds).
    import datetime
    if last_updated_override is not None:
        last_updated = last_updated_override.strftime('%d %b %Y')
    else:
        today = datetime.datetime.now(datetime.timezone.utc).strftime('%d %b %Y')
        last_updated = today
        old_path = Path(output_path)
        if old_path.exists():
            try:
                old_html = old_path.read_text(encoding='utf-8')
                old_guide = _extract_guide_array(old_html)
                if old_guide is not None and old_guide.strip() == new_guide_js.strip():
                    m_date = re.search(r'<time>([^<]*)</time>', old_html)
                    if m_date and m_date.group(1).strip() and m_date.group(1) != '__LAST_UPDATED__':
                        last_updated = m_date.group(1).strip()
            except Exception:
                pass  # any problem -> just use today

    new_html = new_html.replace('__LAST_UPDATED__', last_updated)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(new_html)


def _find_block_bounds(html, decl_regex, open_ch, close_ch):
    """Find the [start, end) byte range of a JS declaration block, where the
    block body is delimited by matching open_ch/close_ch (e.g. [ ] or { }),
    string-aware, and includes a trailing semicolon. Returns (None, None) if
    the declaration isn't found."""
    m = re.search(decl_regex, html)
    if not m:
        return None, None
    start_idx = m.start()
    depth = 0
    in_str = None
    escape = False
    i = m.end() - 1
    n = len(html)
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
        if c == open_ch:
            depth += 1
        elif c == close_ch:
            depth -= 1
            if depth == 0:
                end_idx = i + 1
                j = end_idx
                while j < n and html[j] in ' \t':
                    j += 1
                if j < n and html[j] == ';':
                    end_idx = j + 1
                return start_idx, end_idx
        i += 1
    return start_idx, -1


def _extract_guide_array(html):
    """Extract the 'const GUIDE = [...]' block from an HTML string, or None."""
    m = re.search(r'const GUIDE\s*=\s*\[', html)
    if not m:
        return None
    start_idx = m.start()
    depth = 0
    in_str = None
    escape = False
    i = m.end() - 1
    n = len(html)
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
                return html[start_idx:end_idx]
        i += 1
    return None


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
    doc_date = None
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
            chapter_paths, doc_date = fetch_all_chapters(args.source)
        except RuntimeError as e:
            print(f"ERROR: {e}", file=sys.stderr)
            sys.exit(1)

    # Parse
    chapters_info = []
    eoc_html_by_chapter = {}
    for n, path in chapter_paths:
        print(f"Parsing {path}...")
        title, intro, sections, eoc_entries = parse_chapter(path, n)
        n_steps = sum(len(steps) for _, steps in sections)
        print(f"  Chapter {n}: {len(sections)} sections, {n_steps} steps, "
              f"{len(eoc_entries)} end-of-chapter lines")
        chapters_info.append((n, title, sections, intro))
        if eoc_entries:
            eoc_data = parse_eoc(eoc_entries)
            eoc_html_by_chapter[n] = render_eoc_html(eoc_data, n)
            print(f"    EOC: {len(eoc_data['stats'])} stats, "
                  f"{len(eoc_data['sections'])} sections, "
                  f"{len(eoc_data['routes'])} route lines")

    # Build JS data and splice into base
    data = build_js_data(chapters_info)
    total = sum(len(sec['steps']) for ch in data for sec in ch['sections'])
    print(f"Total: {total} steps")
    if doc_date:
        print(f"Last updated (from doc titles): {doc_date}")

    summaries_js = (build_chapter_summaries_js(eoc_html_by_chapter)
                    if eoc_html_by_chapter else None)

    splice_into_base(data, args.base, args.output,
                     last_updated_override=doc_date,
                     chapter_summaries_js=summaries_js)
    out_size = Path(args.output).stat().st_size
    print(f"Wrote {args.output} ({out_size:,} bytes)")


if __name__ == '__main__':
    main()
