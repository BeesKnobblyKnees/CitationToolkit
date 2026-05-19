"""
Citation Toolkit — Streamlit GUI v2
Adds App 2: Broken Citation Fixer (guided workflow for EndNote db-id issues)

Run with: streamlit run app_v2.py
"""

import base64
import html as html_module
import io
import re
import sqlite3
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from copy import deepcopy
from datetime import datetime
from pathlib import Path

import requests
import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Citation Toolkit",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# STYLES
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }

section[data-testid="stSidebar"] {
    background: #0f1117;
    border-right: 1px solid #1e2530;
}
section[data-testid="stSidebar"] * { color: #c8d0db !important; }

.main .block-container { padding-top: 2rem; max-width: 1000px; }

.section-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.7rem;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    color: #607080;
    border-bottom: 1px solid #1e2530;
    padding-bottom: 0.4rem;
    margin-bottom: 1rem;
}

/* Guided step cards */
.step-card {
    background: #141820;
    border: 1px solid #1e2a38;
    border-radius: 6px;
    padding: 1.2rem 1.4rem;
    margin-bottom: 1rem;
}
.step-card.active  { border-left: 4px solid #4fc3f7; }
.step-card.done    { border-left: 4px solid #43a047; background: #101a10; }
.step-card.waiting { border-left: 4px solid #2a3a4a; opacity: 0.6; }
.step-card.warning { border-left: 4px solid #ffb300; background: #1a1800; }

.step-header {
    display: flex;
    align-items: center;
    gap: 0.8rem;
    margin-bottom: 0.6rem;
}
.step-num {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.7rem;
    background: #1e2a38;
    color: #4fc3f7;
    padding: 0.15rem 0.6rem;
    border-radius: 3px;
    flex-shrink: 0;
}
.step-num.done    { background: #1b3a2a; color: #66bb6a; }
.step-num.warning { background: #2a2200; color: #ffb300; }
.step-title {
    font-size: 0.95rem;
    font-weight: 500;
    color: #c8d0db;
}
.step-desc {
    font-size: 0.83rem;
    color: #607080;
    line-height: 1.6;
    margin-bottom: 0.6rem;
}
.step-action {
    font-size: 0.8rem;
    color: #90a0b0;
    background: #0d1420;
    border: 1px solid #1e2a38;
    border-radius: 4px;
    padding: 0.5rem 0.8rem;
    margin-top: 0.5rem;
    font-family: 'IBM Plex Mono', monospace;
}

/* Ref list items */
.ref-item {
    padding: 0.5rem 0.8rem;
    border-radius: 4px;
    margin-bottom: 4px;
    font-size: 0.82rem;
    border-left: 3px solid;
}
.ref-item.ok      { border-color: #43a047; background: #0d1a0d; color: #90b090; }
.ref-item.missing { border-color: #e53935; background: #1a0d0d; color: #c09090; }
.ref-item.warn    { border-color: #ffb300; background: #1a1500; color: #c0a060; }

/* Metrics */
.metrics-row { display: flex; gap: 10px; margin-bottom: 1.2rem; }
.metric-box {
    flex: 1;
    background: #141820;
    border: 1px solid #1e2530;
    border-radius: 6px;
    padding: 0.8rem 1rem;
    text-align: center;
}
.metric-val { font-size: 26px; font-weight: 300; margin-bottom: 2px; }
.metric-lbl { font-size: 11px; color: #607080; letter-spacing: 0.06em; }

/* Match cards */
.match-card {
    background: #141820;
    border: 1px solid #1e2a38;
    border-left: 3px solid #4fc3f7;
    border-radius: 4px;
    padding: 1rem 1.2rem;
    margin-bottom: 0.8rem;
}
.match-card.accepted { border-left-color: #43a047; background: #101a10; }
.match-card.skipped  { border-left-color: #e53935; background: #1a1010; }
.match-sentence { font-size: 0.85rem; color: #90a0b0; font-style: italic; margin-bottom: 0.5rem; line-height: 1.6; }
.match-marker {
    font-family: 'IBM Plex Mono', monospace; font-size: 0.75rem;
    background: #1e2a38; color: #4fc3f7;
    padding: 0.1rem 0.5rem; border-radius: 3px;
    display: inline-block; margin-bottom: 0.5rem;
}
.score-pill {
    font-family: 'IBM Plex Mono', monospace; font-size: 0.7rem;
    padding: 0.1rem 0.5rem; border-radius: 2px;
    min-width: 3.5rem; text-align: center;
}
.score-high { background: #1b3a2a; color: #66bb6a; }
.score-mid  { background: #2a2a1a; color: #ffca28; }
.score-low  { background: #2a1a1a; color: #ef5350; }

/* Progress */
.progress-outer { height: 4px; background: #1e2530; border-radius: 2px; margin-bottom: 1rem; overflow: hidden; }
.progress-inner { height: 100%; background: #4fc3f7; transition: width 0.3s; }

.stButton > button {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.8rem;
    letter-spacing: 0.06em;
    border-radius: 3px;
    border: 1px solid #2a3a4a;
    background: #141820;
    color: #c8d0db;
    transition: all 0.15s ease;
}
.stButton > button:hover { background: #1e2a38; border-color: #4fc3f7; color: #4fc3f7; }

.badge {
    display: inline-block;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.65rem;
    padding: 0.15rem 0.6rem;
    border-radius: 2px;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    margin-right: 0.4rem;
}
.badge-green  { background: #1b3a2a; color: #66bb6a; }
.badge-red    { background: #3a1a1a; color: #ef5350; }
.badge-yellow { background: #2a2a1a; color: #ffca28; }
.badge-blue   { background: #1a2a3a; color: #4fc3f7; }

.instruction-box {
    background: #0d1525;
    border: 1px solid #1e3050;
    border-radius: 6px;
    padding: 1rem 1.2rem;
    margin: 0.8rem 0;
    font-size: 0.85rem;
    color: #90a8c8;
    line-height: 1.8;
}
.instruction-box b { color: #c8d0db; }
.instruction-box code {
    background: #1a2535;
    padding: 0.1rem 0.4rem;
    border-radius: 3px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.8rem;
    color: #4fc3f7;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

TOP_N           = 5
TFIDF_THRESHOLD = 0.12
PUBMED_MAX      = 5

MISSING_PATTERNS = [
    r'\[CITATION\]', r'\[REF\]', r'\[ref\]', r'\[\?\]',
    r'\[citation needed\]', r'\bXXX\b', r'\[#\]',
    r'<citation>', r'\[ *\]',
]
CITATION_MARKERS = re.compile('|'.join(MISSING_PATTERNS), re.IGNORECASE)

MATCH_THRESHOLD = 0.28
FUZZY_THRESHOLD = 0.08

PUBMED_ESEARCH = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
PUBMED_ESUM    = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi"

# ─────────────────────────────────────────────────────────────────────────────
# SHARED HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def xml_text(elem, path):
    node = elem.find(path)
    if node is None: return ""
    parts = [node.text or ""] + [c.text or "" for c in node]
    return " ".join(p.strip() for p in parts if p.strip())

def fmt_ref(ref, short=False):
    aa = ref.get("authors", [])
    if aa:
        a = aa[0].split(",")[0] if len(aa)==1 else \
            f"{aa[0].split(',')[0]} & {aa[1].split(',')[0]}" if len(aa)==2 else \
            f"{aa[0].split(',')[0]} et al."
    else: a = "Unknown"
    y = ref.get("year","n.d.")
    t = ref.get("title","")[:90] + ("..." if len(ref.get("title",""))>90 else "")
    return f"{a} ({y}) — {t}" if short else f"{a} ({y}). {t}. {ref.get('journal','')}"

def score_class(s):
    if s >= 0.20: return "score-high"
    if s >= 0.10: return "score-mid"
    return "score-low"

def doc_to_bytes(doc):
    buf = io.BytesIO(); doc.save(buf); buf.seek(0); return buf.read()

# ─────────────────────────────────────────────────────────────────────────────
# APP 2 — BROKEN CITATION FIXER (core logic)
# ─────────────────────────────────────────────────────────────────────────────

def analyze_docx_citations(docx_bytes):
    """
    Reads the raw Word XML and categorizes all EN.CITE fields:
    - working: instrText has full EndNote XML
    - broken:  instrText says only 'ADDIN EN.CITE' but fldData has base64 XML
    - empty:   neither has data
    Returns a dict of analysis results.
    """
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        raw = z.read('word/document.xml').decode('utf-8')

    total_fields = raw.count('ADDIN EN.CITE')
    working      = len(re.findall(r'&lt;EndNote&gt;', raw))
    broken_empty = len(re.findall(r'<w:instrText[^>]*> ADDIN EN\.CITE </w:instrText>', raw))
    flddata      = len(re.findall(r'<w:fldData', raw))

    # Decode all fldData to check they have content
    fld_pattern = re.compile(r'<w:fldData[^>]*>([\s\S+?]+?)</w:fldData>')
    fld_matches = fld_pattern.findall(raw)
    recoverable = 0
    for m in fld_matches:
        b64 = m.replace('\r','').replace('\n','').replace(' ','')
        pad = (4 - len(b64) % 4) % 4
        try:
            decoded = base64.b64decode(b64 + '='*pad).decode('utf-8', errors='replace')
            if '<EndNote>' in decoded:
                recoverable += 1
        except: pass

    # Get db-ids present
    db_id_pattern = re.compile(r'&lt;key[^&]*db-id=&quot;([^&]+)&quot;')
    db_ids = list(set(db_id_pattern.findall(raw)))

    return {
        'raw': raw,
        'total_fields': total_fields,
        'working': working,
        'broken_empty': broken_empty,
        'flddata_count': flddata,
        'recoverable': recoverable,
        'db_ids': db_ids,
    }

def fix_broken_fields(raw_xml):
    """Stage 1: Restore broken instrText from fldData base64."""
    pattern = re.compile(
        r'(<w:instrText[^>]*>) ADDIN EN\.CITE (</w:instrText>)'
        r'([\s\S]{0,2000}?)'
        r'<w:fldData[^>]*>([\s\S+?]+?)</w:fldData>',
        re.DOTALL
    )
    fixes = 0
    def replace_fn(m):
        nonlocal fixes
        instr_open, instr_close, between, b64_raw = m.group(1), m.group(2), m.group(3), m.group(4)
        b64 = b64_raw.replace('\r','').replace('\n','').replace(' ','')
        pad = (4 - len(b64) % 4) % 4
        b64 += '=' * pad
        try:
            decoded = base64.b64decode(b64).decode('utf-8', errors='replace')
            decoded = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', decoded)
            escaped = decoded.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')
            fixes += 1
            return (f'{instr_open} ADDIN EN.CITE {escaped}{instr_close}'
                    + between
                    + f'<w:fldData xml:space="preserve">{b64_raw}</w:fldData>')
        except: return m.group(0)
    result = pattern.sub(replace_fn, raw_xml)
    result = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', result)
    return result, fixes

def extract_karol_db_id(xml_export_bytes):
    """Extract the KAROL library db-id from an EndNote XML export."""
    try:
        content = xml_export_bytes.decode('utf-8', errors='replace')
        m = re.search(r'db-id="([a-z0-9]{20,45})"', content, re.IGNORECASE)
        if m: return m.group(1)
    except: pass
    return None

def get_karol_rec_nums(enl_bytes):
    """Read all record IDs from the KAROL .enl SQLite database."""
    with tempfile.NamedTemporaryFile(suffix='.enl', delete=False) as f:
        f.write(enl_bytes); tmp_path = f.name
    try:
        conn = sqlite3.connect(tmp_path)
        cursor = conn.cursor()
        cursor.execute('SELECT id, author, year, title FROM enl_refs WHERE trash_state=0 OR trash_state IS NULL')
        rows = cursor.fetchall()
        conn.close()
        return {str(row[0]): {'id': str(row[0]), 'author': row[1] or '', 'year': str(row[2] or ''), 'title': row[3] or ''} for row in rows}
    except Exception as e:
        return {}

def patch_db_ids(raw_xml, old_db_ids, new_db_id):
    """Stage 2: Replace all old db-id values with the KAROL db-id."""
    result = raw_xml
    replaced = 0
    for old_id in old_db_ids:
        if old_id != new_db_id:
            count = result.count(old_id)
            result = result.replace(old_id, new_db_id)
            replaced += count
    return result, replaced

def check_missing_from_karol(raw_xml, karol_rec_nums):
    """Find any RecNums in the doc that don't exist in KAROL."""
    all_rec_nums = set(re.findall(r'&lt;RecNum&gt;(\d+)&lt;/RecNum&gt;', raw_xml))
    missing = []
    for rn in all_rec_nums:
        if rn not in karol_rec_nums:
            missing.append(rn)
    return missing

def build_fixed_docx(original_bytes, fixed_xml):
    """Write the fixed XML back into the docx ZIP."""
    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(original_bytes)) as zin:
        all_files = {name: zin.read(name) for name in zin.namelist()}
    all_files['word/document.xml'] = fixed_xml.encode('utf-8')
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_files.items():
            zout.writestr(name, data)
    buf.seek(0)
    return buf.read()

def extract_traveling_library_xml(docx_bytes):
    """
    Extracts all references from the traveling library embedded in the Word doc
    and returns them as an EndNote XML string ready to import into KAROL.
    """
    import base64, html as html_module
    from lxml import etree

    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        raw = z.read('word/document.xml').decode('utf-8')

    all_cite_xml = []
    cite_pattern = re.compile(r'ADDIN EN\.CITE &lt;EndNote&gt;([\s\S]+?)&lt;/EndNote&gt;')
    for m in cite_pattern.findall(raw):
        all_cite_xml.append(html_module.unescape(f'<EndNote>{m}</EndNote>'))

    fld_pattern = re.compile(r'<w:fldData[^>]*>([\s\S+?]+?)</w:fldData>')
    for b64_raw in fld_pattern.findall(raw):
        b64 = b64_raw.replace('\r','').replace('\n','').replace(' ','')
        pad = (4 - len(b64) % 4) % 4
        try:
            decoded = base64.b64decode(b64+'='*pad).decode('utf-8',errors='replace').replace('\x00','')
            if '<EndNote>' in decoded:
                all_cite_xml.append(decoded)
        except: pass

    traveling_refs = {}
    for cite_xml in all_cite_xml:
        try:
            if not cite_xml.startswith('<EndNote>'): cite_xml = f'<EndNote>{cite_xml}</EndNote>'
            root = etree.fromstring(cite_xml.encode('utf-8'))
            for cite in root.findall('.//Cite'):
                rec_num = cite.findtext('RecNum') or ''
                record  = cite.find('record')
                if rec_num and record is not None and rec_num not in traveling_refs:
                    traveling_refs[rec_num] = record
        except: pass

    def get_t(elem, tag):
        node = elem.find(f'.//{tag}')
        return ''.join(node.itertext()).strip() if node is not None else ''

    import html as html_mod
    output = '<?xml version="1.0" encoding="UTF-8"?>\n<xml>\n  <records>\n'

    for rec_num, record in sorted(traveling_refs.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 9999):
        ref_type_el   = record.find('.//ref-type')
        ref_type_name = ref_type_el.get('name','Journal Article') if ref_type_el is not None else 'Journal Article'
        ref_type_num  = ref_type_el.text if ref_type_el is not None else '17'

        authors    = [''.join(a.itertext()).strip() for a in record.findall('.//contributors/authors/author')]
        sec_authors= [''.join(a.itertext()).strip() for a in record.findall('.//contributors/secondary-authors/author')]
        title      = get_t(record,'title')
        sec_title  = get_t(record,'secondary-title')
        tert_title = get_t(record,'tertiary-title')
        year       = get_t(record,'year')
        volume     = get_t(record,'volume')
        issue      = get_t(record,'number')
        pages      = get_t(record,'pages')
        edition    = get_t(record,'edition')
        publisher  = get_t(record,'publisher')
        place      = get_t(record,'pub-location')
        abstract   = get_t(record,'abstract')
        keywords   = [''.join(k.itertext()).strip() for k in record.findall('.//keyword')]

        r = f'    <record>\n'
        r += f'      <rec-number>{rec_num}</rec-number>\n'
        r += f'      <ref-type name="{html_mod.escape(ref_type_name)}">{ref_type_num}</ref-type>\n'
        if authors:
            r += '      <contributors>\n        <authors>\n'
            for a in authors:
                if a: r += f'          <author>{html_mod.escape(a)}</author>\n'
            r += '        </authors>\n'
            if sec_authors:
                r += '        <secondary-authors>\n'
                for a in sec_authors:
                    if a: r += f'          <author>{html_mod.escape(a)}</author>\n'
                r += '        </secondary-authors>\n'
            r += '      </contributors>\n'
        r += '      <titles>\n'
        if title:      r += f'        <title>{html_mod.escape(title)}</title>\n'
        if sec_title:  r += f'        <secondary-title>{html_mod.escape(sec_title)}</secondary-title>\n'
        if tert_title: r += f'        <tertiary-title>{html_mod.escape(tert_title)}</tertiary-title>\n'
        r += '      </titles>\n'
        if year:      r += f'      <dates><year>{html_mod.escape(year)}</year></dates>\n'
        if volume:    r += f'      <volume>{html_mod.escape(volume)}</volume>\n'
        if issue:     r += f'      <number>{html_mod.escape(issue)}</number>\n'
        if pages:     r += f'      <pages>{html_mod.escape(pages)}</pages>\n'
        if edition:   r += f'      <edition>{html_mod.escape(edition)}</edition>\n'
        if publisher: r += f'      <publisher>{html_mod.escape(publisher)}</publisher>\n'
        if place:     r += f'      <pub-location>{html_mod.escape(place)}</pub-location>\n'
        if abstract:  r += f'      <abstract>{html_mod.escape(abstract)}</abstract>\n'
        if [k for k in keywords if k]:
            r += '      <keywords>\n'
            for kw in keywords:
                if kw: r += f'        <keyword>{html_mod.escape(kw)}</keyword>\n'
            r += '      </keywords>\n'
        r += '    </record>\n'
        output += r

    output += '  </records>\n</xml>\n'
    return output, len(traveling_refs)

def remap_traveling_citations(docx_bytes, enl_bytes):
    """
    Finds citations still pointing to traveling library RecNums not in KAROL,
    matches them by author+year+title to correct KAROL records, and remaps them.
    Returns (fixed_bytes, report).
    """
    import base64 as b64mod, html as html_mod, sqlite3, tempfile, os
    from lxml import etree

    # Load KAROL library
    with tempfile.NamedTemporaryFile(suffix='.enl', delete=False) as f:
        f.write(enl_bytes); tmp_path = f.name
    try:
        conn = sqlite3.connect(tmp_path)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = [t[0] for t in cursor.fetchall()]
        tbl = 'refs' if 'refs' in tables else 'enl_refs'
        cursor.execute(f'SELECT id, author, year, title FROM {tbl} WHERE trash_state=0 OR trash_state IS NULL')
        karol_rows = cursor.fetchall()
        conn.close()
    finally:
        os.unlink(tmp_path)

    def norm(s):
        if not s: return ''
        return re.sub(r'[^a-z0-9]', '', s.lower())

    def auth_last(s):
        if not s: return ''
        return norm(s.split('\r')[0].split('\n')[0].strip().split(',')[0])

    karol_ids = set(str(r[0]) for r in karol_rows)
    karol_by_authyear = {}
    for row in karol_rows:
        kid, kauth, kyear, ktitle = row
        key = (auth_last(kauth or ''), norm(str(kyear or '')))
        if key not in karol_by_authyear:
            karol_by_authyear[key] = []
        karol_by_authyear[key].append(row)

    # Read docx
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        raw = z.read('word/document.xml').decode('utf-8')
        all_files = {name: z.read(name) for name in z.namelist()}

    cite_pat = re.compile(r'ADDIN EN\.CITE &lt;EndNote&gt;([\s\S]+?)&lt;/EndNote&gt;')
    report = []
    remap_dict = {}

    for cx_esc in cite_pat.findall(raw):
        cx = html_mod.unescape(f'<EndNote>{cx_esc}</EndNote>')
        try:
            root = etree.fromstring(cx.encode('utf-8'))
            for cite in root.findall('.//Cite'):
                rn  = cite.findtext('RecNum') or ''
                au  = cite.findtext('Author') or ''
                yr  = cite.findtext('Year') or ''
                if rn in karol_ids or rn in remap_dict:
                    continue
                # Get title from traveling library record
                title = ''
                rec = cite.find('record')
                if rec is not None:
                    tel = rec.find('.//title')
                    if tel is not None: title = ''.join(tel.itertext())
                # Match by author+year
                key = (auth_last(au), norm(yr))
                cands = karol_by_authyear.get(key, [])
                matched = None
                if len(cands) == 1:
                    matched = str(cands[0][0])
                elif len(cands) > 1:
                    for c in cands:
                        if norm(title)[:25] and norm(str(c[3] or ''))[:25] == norm(title)[:25]:
                            matched = str(c[0]); break
                    if not matched: matched = str(cands[0][0])
                remap_dict[rn] = matched
                report.append({'status': 'remapped' if matched else 'not_found',
                               'old_rec_num': rn, 'new_rec_num': matched,
                               'author': au, 'year': yr, 'title': title[:80]})
        except: pass

    if not remap_dict:
        return docx_bytes, report

    # Apply RecNum remapping to instrText
    fixed = raw
    for old_rn, new_rn in remap_dict.items():
        if new_rn:
            fixed = fixed.replace(
                f'&lt;RecNum&gt;{old_rn}&lt;/RecNum&gt;',
                f'&lt;RecNum&gt;{new_rn}&lt;/RecNum&gt;'
            )

    # Apply to fldData base64 blobs
    def fix_fld(m):
        b64r = m.group(1)
        b64 = b64r.replace('\r','').replace('\n','').replace(' ','')
        pad = (4 - len(b64) % 4) % 4
        try:
            dec = b64mod.b64decode(b64+'='*pad).decode('utf-8',errors='replace').replace('\x00','')
            mod = dec
            for old_rn, new_rn in remap_dict.items():
                if new_rn:
                    mod = mod.replace(f'<RecNum>{old_rn}</RecNum>', f'<RecNum>{new_rn}</RecNum>')
            if mod != dec:
                nb64 = b64mod.b64encode(mod.encode('utf-8')).decode('ascii')
                wrapped = '\r\n'.join(nb64[i:i+76] for i in range(0, len(nb64), 76))
                return f'<w:fldData xml:space="preserve">{wrapped}</w:fldData>'
        except: pass
        return m.group(0)

    fixed = re.compile(r'<w:fldData[^>]*>([\s\S+?]+?)</w:fldData>').sub(fix_fld, fixed)

    all_files['word/document.xml'] = fixed.encode('utf-8')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in all_files.items():
            zout.writestr(name, data)
    buf.seek(0)
    return buf.read(), report



def generate_vba_macro(doc_name):
    """Return the VBA macro text, customised with the document name."""
    return f'''
' ══════════════════════════════════════════════════════════════
'  EndNote Citation Relinker — Word VBA Macro
'  Generated by Citation Toolkit for: {doc_name}
'  Generated: {datetime.now():%Y-%m-%d %H:%M}
' ══════════════════════════════════════════════════════════════
'
' HOW TO USE:
'  1. Open "{doc_name}" in Word
'  2. Press Alt+F11 to open the VBA editor
'  3. Insert → Module
'  4. Paste this entire macro
'  5. Close the VBA editor
'  6. Press Alt+F8 → select RelinkAllCitations → Run
' ══════════════════════════════════════════════════════════════

Sub RelinkAllCitations()
    Dim oDoc As Document
    Set oDoc = ActiveDocument

    If oDoc Is Nothing Then
        MsgBox "No document is open.", vbCritical
        Exit Sub
    End If

    ' Step 1: Confirm EndNote is open with KAROL
    Dim msg As String
    msg = "Before running this macro, please confirm:" & vbCrLf & vbCrLf & _
          "  1. EndNote is open" & vbCrLf & _
          "  2. Your KAROL library is loaded in EndNote" & vbCrLf & _
          "  3. You have added Friedman 1995 to KAROL (if not already there)" & vbCrLf & vbCrLf & _
          "A backup will be saved automatically. Continue?"

    If MsgBox(msg, vbYesNo + vbQuestion, "EndNote Relinker") = vbNo Then Exit Sub

    ' Step 2: Save backup
    Dim sBackup As String
    sBackup = oDoc.Path & "\\{Path(doc_name).stem}_BACKUP_" & Format(Now,"YYYYMMDD_HHMMSS") & ".docx"
    oDoc.SaveAs2 sBackup, wdFormatXMLDocument
    MsgBox "Backup saved to:" & vbCrLf & sBackup, vbInformation

    ' Step 3: Convert to unformatted citations
    Application.StatusBar = "Converting citations to temporary format..."
    On Error Resume Next
    Application.Run "EndNote.UnformatAll"
    If Err.Number <> 0 Then
        Err.Clear
        ' Manual fallback
        Dim oField As Field
        For Each oField In oDoc.Fields
            If InStr(1, oField.Code.Text, "ADDIN EN.CITE", vbTextCompare) > 0 Then
                Dim sCode As String
                sCode = oField.Code.Text
                If InStr(sCode, "<Author>") > 0 Or InStr(sCode, "&lt;Author&gt;") > 0 Then
                    oField.Unlink
                End If
            End If
        Next oField
    End If
    On Error GoTo 0

    ' Step 4: Re-format / re-link to KAROL
    Application.StatusBar = "Re-linking citations to KAROL..."
    On Error Resume Next
    Application.Run "EndNote.FormatAll"
    If Err.Number <> 0 Then
        Err.Clear
        Application.Run "EndNote.UpdateAll"
    End If
    On Error GoTo 0

    ' Step 5: Save
    oDoc.Save
    Application.StatusBar = ""

    MsgBox "Done! Citations have been re-linked to your KAROL library." & vbCrLf & vbCrLf & _
           "If any citations appear highlighted in yellow, right-click each one " & vbCrLf & _
           "and choose Edit Citation > Find to fix manually.", _
           vbInformation, "EndNote Relinker Complete"
End Sub
'''

# ─────────────────────────────────────────────────────────────────────────────
# APP 1 LOGIC (citation repair)
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def parse_endnote_xml_bytes(xml_bytes):
    root = ET.fromstring(xml_bytes)
    refs = []
    for rec in root.iter("record"):
        authors = []
        for a in rec.findall(".//contributors/authors/author"):
            name = " ".join(p.strip() for p in ([a.text or ""]+[c.text or "" for c in a]) if p.strip())
            if name: authors.append(name)
        title   = xml_text(rec, ".//titles/title")
        journal = xml_text(rec, ".//periodical/full-title") or xml_text(rec, ".//periodical/abbr-1")
        year    = xml_text(rec, ".//dates/year")
        abstract= xml_text(rec, ".//abstract")
        rec_num = xml_text(rec, "rec-number")
        if not title: continue
        corpus = " ".join(filter(None,[title,abstract,journal,year]))
        refs.append(dict(rec_number=rec_num,authors=authors,title=title,journal=journal,year=year,corpus=corpus))
    return refs

@st.cache_resource(show_spinner=False)
def build_tfidf(corpora_tuple):
    vec = TfidfVectorizer(ngram_range=(1,2), sublinear_tf=True, max_features=50000)
    mat = vec.fit_transform(list(corpora_tuple))
    return vec, mat

def match_sentence(sentence, vec, mat, refs, top_n=TOP_N):
    sv   = vec.transform([sentence])
    sims = cosine_similarity(sv, mat)[0]
    idx  = sims.argsort()[::-1][:top_n]
    return [dict(ref=refs[i], score=float(sims[i])) for i in idx]

def extract_flagged(docx_bytes):
    doc = Document(io.BytesIO(docx_bytes))
    flagged = []
    for pi, para in enumerate(doc.paragraphs):
        text = para.text
        if not text.strip(): continue
        for m in CITATION_MARKERS.finditer(text):
            sents = re.split(r'(?<=[.!?])\s+', text)
            cum, target = 0, text
            for s in sents:
                if cum + len(s) >= m.start(): target = s; break
                cum += len(s) + 1
            flagged.append(dict(para_idx=pi, sentence=target, marker=m.group(), para_text=text))
    return flagged, doc

def author_label(ref):
    aa = ref.get("authors",[])
    last = aa[0].split(",")[0].strip().split()[-1] if aa else "Ref"
    return (last + " " + ref.get("year", "")).strip()

def insert_superscript(para, marker, label):
    if marker not in para.text: return False
    combined, run_map = "", []
    for run in para.runs:
        s = len(combined); combined += run.text
        run_map.append((s, s+len(run.text), run))
    pos = combined.find(marker)
    if pos == -1: return False
    for (s, e, run) in run_map:
        if s <= pos < e:
            before = run.text[:pos-s]; after = run.text[pos-s+len(marker):]
            run.text = before
            rPr = OxmlElement("w:rPr")
            va = OxmlElement("w:vertAlign"); va.set(qn("w:val"),"superscript"); rPr.append(va)
            nr = OxmlElement("w:r"); nr.append(deepcopy(rPr))
            t = OxmlElement("w:t"); t.text = f"[{label}]"
            t.set("{http://www.w3.org/XML/1998/namespace}space","preserve"); nr.append(t)
            run._r.addnext(nr)
            if after:
                tr = OxmlElement("w:r"); tt = OxmlElement("w:t")
                tt.text = after; tt.set("{http://www.w3.org/XML/1998/namespace}space","preserve")
                tr.append(tt); nr.addnext(tr)
            return True
    return False

def write_repair_report(decisions):
    doc = Document()
    doc.add_heading("Citation Repair Report", 0)
    doc.add_paragraph(f"Generated: {datetime.now():%Y-%m-%d %H:%M}")
    accepted = [d for d in decisions if d["action"]=="accepted"]
    skipped  = [d for d in decisions if d["action"]=="skipped"]
    pm_list  = [d for d in decisions if d["action"]=="pubmed"]
    doc.add_paragraph(f"Total: {len(decisions)} | Accepted: {len(accepted)} | Skipped: {len(skipped)} | PubMed: {len(pm_list)}")
    if accepted:
        doc.add_heading("Accepted", 1)
        for d in accepted:
            p = doc.add_paragraph(style="List Bullet")
            p.add_run(f"Marker: {d['marker']}\n").bold = True
            p.add_run(f"Context: {d['sentence'][:200]}\n")
            p.add_run(f"Inserted: {fmt_ref(d['ref'])}\n")
            p.add_run(f"Score: {d['score']:.3f}")
    if skipped:
        doc.add_heading("Skipped", 1)
        for d in skipped:
            p = doc.add_paragraph(style="List Bullet")
            r = p.add_run("NEEDS REVIEW\n"); r.font.color.rgb = RGBColor(0xC0,0,0)
            p.add_run(f"Marker: {d['marker']}\nContext: {d['sentence'][:200]}\n")
    return doc_to_bytes(doc)

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

defaults = dict(
    # App 1
    flagged=[], current_idx=0, decisions=[],
    doc_obj=None, repair_done=False,
    refs=[], vec=None, mat=None,
    # App 2
    fix_stage=1,
    fix_analysis=None,
    fix_raw_xml=None,
    fix_docx_bytes=None,
    fix_after_stage1=None,
    fix_after_stage2=None,
    fix_karol_db_id=None,
    fix_karol_rec_nums={},
    fix_missing_refs=[],
    fix_doc_name="document.docx",
    # App 3
    comp_result=None, comp_usage={}, comp_labels=("",""), comp_refs=([],[]),
)
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## 📚 Citation Toolkit")
    st.caption("Pediatric Orthopedics — Tachdjian's")
    st.divider()
    tool = st.radio("Tool", [
        "App 1 — Citation Repair",
        "App 2 — Broken Citation Fixer",
        "App 3 — Reference Comparator",
    ], label_visibility="collapsed")
    st.divider()
    st.markdown("""
    <div style="font-size:0.72rem;color:#3a4a5a;line-height:1.8">
    <b style="color:#4fc3f7">App 1</b> — Fill missing citation<br>
    placeholders from library<br><br>
    <b style="color:#4fc3f7">App 2</b> — Fix broken citations<br>
    not recognized by EndNote<br><br>
    <b style="color:#4fc3f7">App 3</b> — Compare two ref lists,<br>
    find missing/extra refs
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# APP 2 — BROKEN CITATION FIXER
# ─────────────────────────────────────────────────────────────────────────────

if tool == "App 2 — Broken Citation Fixer":
    st.markdown("## Broken Citation Fixer")
    st.markdown("""
    <div class="instruction-box">
    <b>When to use this tool:</b> EndNote only recognizes some of your citations even though the 
    bibliography shows all references — this means the citation field codes lost their data during 
    document merging or conversion. This tool walks you through the full fix in three stages.
    </div>
    """, unsafe_allow_html=True)

    # ── STEP 0: Extract traveling library ────────────────────────────────────

    with st.expander("📥  Extract traveling library references (start here if EndNote can't find your refs)", expanded=True):
        st.markdown("""
        <div class="step-desc">
        If EndNote only recognizes a few citations even though your bibliography shows all of them,
        the references exist in the <b>traveling library</b> embedded inside the Word file — but not
        in your connected EndNote library. Extract them here and import the result into KAROL.
        </div>
        """, unsafe_allow_html=True)

        tl_file = st.file_uploader("Upload Word document to extract references from",
                                    type=["docx"], key="tl_doc_upload")
        if tl_file:
            tl_bytes = tl_file.read()
            with st.spinner("Extracting traveling library references..."):
                try:
                    tl_xml, tl_count = extract_traveling_library_xml(tl_bytes)
                    st.success(f"✓ Extracted **{tl_count}** references from the traveling library.")

                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(
                            f"⬇ Download EndNote XML ({tl_count} refs)",
                            data=tl_xml.encode('utf-8'),
                            file_name=Path(tl_file.name).stem + "_traveling_library.xml",
                            mime="application/xml",
                            type="primary",
                            use_container_width=True,
                            help="Import this into KAROL via: File → Import → File → select EndNote XML"
                        )

                    st.markdown("""
                    <div class="instruction-box">
                    <b>Import into KAROL:</b><br>
                    1. Open KAROL in EndNote<br>
                    2. <b>File → Import → File</b><br>
                    3. Select the downloaded XML file<br>
                    4. Set <b>Import Option</b> to <code>EndNote XML</code><br>
                    5. Set <b>Duplicates</b> to <code>Discard Duplicates</code><br>
                    6. Click <b>Import</b><br><br>
                    After importing, re-open your Word document with KAROL connected —
                    EndNote should now recognize all the references.
                    </div>
                    """, unsafe_allow_html=True)

                except Exception as e:
                    st.error(f"Could not extract references: {e}")

    st.divider()

        # ── STAGE 1: Upload & analyze ──────────────────────────────────────────

    stage1_done = st.session_state.fix_analysis is not None

    st.markdown(f"""
    <div class="step-card {'done' if stage1_done else 'active'}">
        <div class="step-header">
            <span class="step-num {'done' if stage1_done else ''}">{'✓' if stage1_done else 'STEP 1'}</span>
            <span class="step-title">Upload your document — detect and repair broken fields</span>
        </div>
        <div class="step-desc">
            Upload your Word document. The tool will scan all citation fields, identify which ones 
            have lost their data, and automatically recover it from the document's backup storage.
            No library needed for this step.
        </div>
    </div>
    """, unsafe_allow_html=True)

    if not stage1_done:
        doc_file = st.file_uploader("Word document (.docx)", type=["docx"], key="fix_doc_upload")

        if doc_file:
            col1, col2 = st.columns([3,1])
            with col2:
                run_s1 = st.button("Analyze & repair", type="primary", use_container_width=True)

            if run_s1:
                with st.spinner("Scanning citation fields..."):
                    docx_bytes = doc_file.read()
                    analysis   = analyze_docx_citations(docx_bytes)

                st.session_state.fix_analysis   = analysis
                st.session_state.fix_docx_bytes = docx_bytes
                st.session_state.fix_doc_name   = doc_file.name

                with st.spinner("Repairing broken fields..."):
                    fixed_xml, n_fixed = fix_broken_fields(analysis['raw'])
                    st.session_state.fix_after_stage1 = fixed_xml
                    st.session_state.fix_raw_xml = fixed_xml

                st.rerun()

    else:
        a = st.session_state.fix_analysis
        fixed_xml = st.session_state.fix_after_stage1
        n_fixed = a['broken_empty']

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total citation fields", a['total_fields'])
        col2.metric("Already working", a['working'])
        col3.metric("Were broken", a['broken_empty'], help="Had lost their XML data")
        col4.metric("Recovered", n_fixed, help="Restored from backup storage in the document")

        if n_fixed > 0:
            st.success(f"✓ {n_fixed} broken citation field(s) recovered from internal backup data.")
        else:
            st.info("No broken fields found — all citation data was intact.")

        # Download Stage 1 result
        stage1_bytes = build_fixed_docx(st.session_state.fix_docx_bytes, fixed_xml)
        st.download_button(
            "⬇ Download Stage 1 result (.docx)",
            data=stage1_bytes,
            file_name=Path(st.session_state.fix_doc_name).stem + "_stage1_fixed.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            help="Download now if you only needed to fix broken fields. For the db-id fix, continue to Stage 2."
        )

        if st.button("↺ Start over with a different document"):
            for k in ['fix_analysis','fix_raw_xml','fix_docx_bytes','fix_after_stage1',
                      'fix_after_stage2','fix_karol_db_id','fix_karol_rec_nums','fix_missing_refs']:
                st.session_state[k] = defaults[k]
            st.rerun()

    st.divider()

    # ── STAGE 2: Get KAROL db-id ───────────────────────────────────────────

    stage2_done = st.session_state.fix_karol_db_id is not None

    st.markdown(f"""
    <div class="step-card {'done' if stage2_done else ('active' if stage1_done else 'waiting')}">
        <div class="step-header">
            <span class="step-num {'done' if stage2_done else ''}">{'✓' if stage2_done else 'STEP 2'}</span>
            <span class="step-title">Provide your KAROL library — get its unique fingerprint</span>
        </div>
        <div class="step-desc">
            Every EndNote library has a unique ID (db-id). Citations in your document point to 
            their original library's ID, not KAROL's. We need KAROL's ID to re-point them.<br><br>
            <b>You need two things from EndNote:</b>
        </div>
        <div class="step-action">
            Option A (easiest): In EndNote with KAROL open → <b>File → Export</b> → 
            set format to <b>XML</b> → export a few records → upload that .xml file below<br><br>
            Option B: Upload your KAROL <b>.enl file</b> directly (for checking which refs are missing)
        </div>
    </div>
    """, unsafe_allow_html=True)

    if stage1_done and not stage2_done:
        col_a, col_b = st.columns(2)
        with col_a:
            xml_export = st.file_uploader(
                "EndNote XML export (gets db-id) — export any refs from KAROL as XML",
                type=["xml"], key="fix_xml_export"
            )
        with col_b:
            enl_file = st.file_uploader(
                "KAROL .enl library (checks for missing refs, optional)",
                type=["enl"], key="fix_enl"
            )

        if xml_export:
            db_id = extract_karol_db_id(xml_export.read())
            if db_id:
                st.session_state.fix_karol_db_id = db_id
                st.success(f"✓ KAROL library fingerprint found: `{db_id[:16]}...`")
            else:
                st.error("Could not find a db-id in this XML. Make sure you exported from EndNote as XML format (not RIS or plain text).")

        if enl_file:
            with st.spinner("Reading KAROL library..."):
                rec_nums = get_karol_rec_nums(enl_file.read())
                st.session_state.fix_karol_rec_nums = rec_nums
                # Check for missing refs
                missing = check_missing_from_karol(
                    st.session_state.fix_raw_xml or "",
                    set(rec_nums.keys())
                )
                st.session_state.fix_missing_refs = missing
                st.success(f"✓ KAROL library read — {len(rec_nums):,} references.")

                if missing:
                    st.warning(f"⚠ {len(missing)} reference(s) in the document are NOT in KAROL:")
                    for rn in missing:
                        st.markdown(f"- RecNum **{rn}** — not found in KAROL (needs to be added manually)")

        # Manual db-id entry option
        with st.expander("Enter db-id manually (advanced)"):
            manual_id = st.text_input("Paste KAROL db-id here", placeholder="e.g. s5pa559ekdxfr0esvw85...")
            if st.button("Use this db-id") and manual_id.strip():
                st.session_state.fix_karol_db_id = manual_id.strip()
                st.rerun()

    elif stage2_done:
        db_id = st.session_state.fix_karol_db_id
        st.success(f"✓ KAROL db-id: `{db_id[:20]}...`")
        if st.session_state.fix_karol_rec_nums:
            st.caption(f"KAROL library: {len(st.session_state.fix_karol_rec_nums):,} references loaded")
        if st.session_state.fix_missing_refs:
            st.warning(f"⚠ {len(st.session_state.fix_missing_refs)} ref(s) not in KAROL — see Stage 3 for details")

    st.divider()

    # ── STAGE 3: Check missing refs & apply full fix ───────────────────────

    stage3_done = st.session_state.fix_after_stage2 is not None
    stage3_ready = stage1_done and stage2_done

    st.markdown(f"""
    <div class="step-card {'done' if stage3_done else ('active' if stage3_ready else 'waiting')}">
        <div class="step-header">
            <span class="step-num {'done' if stage3_done else ''}">{'✓' if stage3_done else 'STEP 3'}</span>
            <span class="step-title">Add missing references to KAROL, then apply full fix</span>
        </div>
        <div class="step-desc">
            Any references that exist in the document but not in KAROL need to be added to KAROL 
            manually before the final fix. Then this tool updates all citation fingerprints to 
            point to KAROL, and generates the VBA macro for the final re-link in Word.
        </div>
    </div>
    """, unsafe_allow_html=True)

    if stage3_ready and not stage3_done:
        missing_refs = st.session_state.fix_missing_refs

        if missing_refs:
            st.markdown("### References to add to KAROL before continuing")
            st.markdown("""
            <div class="instruction-box">
            In EndNote with KAROL open: <b>References → New Reference</b> for each ref below.<br>
            After adding them all, re-export a small XML from KAROL and re-upload in Step 2 
            to confirm they're there — then come back and click Apply Fix below.
            </div>
            """, unsafe_allow_html=True)

            for rn in missing_refs:
                karol_recs = st.session_state.fix_karol_rec_nums
                info = karol_recs.get(rn, {})
                st.markdown(f"""
                <div class="ref-item missing">
                    <b>RecNum {rn}</b> — not in KAROL &nbsp;
                    {f'| {info.get("author","")[:40]} ({info.get("year","")}) — {info.get("title","")[:60]}' if info else ''}
                </div>
                """, unsafe_allow_html=True)

            st.markdown("")
            col1, col2 = st.columns(2)
            with col1:
                proceed_anyway = st.checkbox("I've added the missing refs to KAROL (or I'll add them later)")
        else:
            st.success("✓ All references in the document are present in KAROL.")
            proceed_anyway = True

        if proceed_anyway:
            if st.button("Apply full fix & generate files", type="primary"):
                with st.spinner("Patching db-ids in all citation fields..."):
                    old_db_ids = st.session_state.fix_analysis['db_ids']
                    new_db_id  = st.session_state.fix_karol_db_id
                    patched_xml, n_replaced = patch_db_ids(
                        st.session_state.fix_raw_xml,
                        old_db_ids,
                        new_db_id
                    )
                    st.session_state.fix_after_stage2 = patched_xml

                st.rerun()

    elif stage3_done:
        patched_xml   = st.session_state.fix_after_stage2
        docx_bytes    = st.session_state.fix_docx_bytes
        doc_name      = st.session_state.fix_doc_name

        # Build final fixed docx
        final_bytes = build_fixed_docx(docx_bytes, patched_xml)

        # Count final state
        working_after = len(re.findall(r'&lt;EndNote&gt;', patched_xml))
        old_ids       = st.session_state.fix_analysis['db_ids']
        new_id        = st.session_state.fix_karol_db_id

        st.success("✓ Full fix applied.")

        col1, col2, col3 = st.columns(3)
        col1.metric("Citation fields fixed", st.session_state.fix_analysis['broken_empty'])
        col2.metric("db-ids updated", len([d for d in old_ids if d != new_id]))
        col3.metric("Total citations now linked", working_after)

        st.markdown("### Download your files")

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.download_button(
                "⬇ Download fixed document (.docx)",
                data=final_bytes,
                file_name=Path(doc_name).stem + "_fully_fixed.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
        with col_d2:
            macro_text = generate_vba_macro(doc_name)
            st.download_button(
                "⬇ Download VBA macro (.bas)",
                data=macro_text.encode(),
                file_name="RelinkEndNoteCitations.bas",
                mime="text/plain",
                use_container_width=True
            )

        st.divider()
        st.markdown("### Final step — re-link in Word")
        st.markdown(f"""
        <div class="instruction-box">
        <b>You're almost done. One final step in Word + EndNote:</b><br><br>
        1. Open the downloaded <b>{Path(doc_name).stem}_fully_fixed.docx</b> in Word<br>
        2. Make sure <b>KAROL is open in EndNote</b><br>
        3. Press <b>Alt+F11</b> in Word → Insert → Module → paste the downloaded <b>.bas</b> macro<br>
        4. Press <b>Alt+F8</b> → select <code>RelinkAllCitations</code> → <b>Run</b><br>
        5. EndNote will convert citations to temporary format then re-link them all to KAROL<br>
        6. The bibliography will rebuild automatically<br><br>
        <b>Why this last step?</b> The db-id fix tells EndNote which library each citation 
        belongs to. The macro then asks EndNote to formally re-link each one, which syncs 
        the citation formatting and makes future updates work correctly.
        </div>
        """, unsafe_allow_html=True)

        if st.session_state.fix_missing_refs:
            st.markdown("### ⚠ Still missing from KAROL")
            st.markdown("""
            <div class="instruction-box" style="border-color:#ffb300">
            These references were in the document but not in your KAROL library.
            Add them manually in EndNote before running the macro, 
            otherwise they will appear unresolved (highlighted yellow) in the document.
            </div>
            """, unsafe_allow_html=True)
            for rn in st.session_state.fix_missing_refs:
                st.markdown(f"- **RecNum {rn}** — add to KAROL manually")

# ─────────────────────────────────────────────────────────────────────────────
# APP 1 — CITATION REPAIR
# ─────────────────────────────────────────────────────────────────────────────

elif tool == "App 1 — Citation Repair":
    st.markdown("## Citation Repair")
    st.markdown('<div class="step-desc">Find missing citation placeholders in your manuscript and match them to your EndNote library.</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        doc_file = st.file_uploader("Word document (.docx)", type=["docx"], key="doc_upload")
    with col2:
        lib_file = st.file_uploader("EndNote library (.xml)", type=["xml"], key="lib_upload")

    cfg1, cfg2 = st.columns(2)
    with cfg1:
        mode = st.selectbox("Mode", ["Interactive Review", "Auto-Insert", "Report Only"])
    with cfg2:
        use_pubmed = st.toggle("PubMed fallback for low matches", value=False)

    with st.expander("Custom citation markers"):
        custom = st.text_input("Additional markers (comma-separated)", placeholder="e.g. ??, [TBD]")
        if custom:
            extras = [re.escape(m.strip()) for m in custom.split(",") if m.strip()]
            CITATION_MARKERS = re.compile('|'.join(MISSING_PATTERNS + extras), re.IGNORECASE)

    run_btn = st.button("Scan Document", type="primary", disabled=not (doc_file and lib_file))

    if run_btn and doc_file and lib_file:
        for k in ["flagged","current_idx","decisions","doc_obj","refs","vec","mat","repair_done"]:
            st.session_state[k] = defaults[k]
        with st.spinner("Parsing..."):
            db = doc_file.read()
            flagged, doc_obj = extract_flagged(db)
            refs = parse_endnote_xml_bytes(lib_file.read())
            if refs:
                corpora = tuple(r["corpus"] for r in refs)
                vec, mat = build_tfidf(corpora)
                st.session_state.update(dict(flagged=flagged, doc_obj=doc_obj,
                    refs=refs, vec=vec, mat=mat, current_idx=0, decisions=[], repair_done=False))
        if not refs:
            st.error("No references found in XML.")
        elif not flagged:
            st.warning("No citation markers found.")
        else:
            st.success(f"Found **{len(flagged)}** missing citation(s) across **{len(refs)}** library references.")

    flagged = st.session_state.flagged
    if flagged and not st.session_state.repair_done:
        st.divider()
        idx = st.session_state.current_idx
        total = len(flagged); done = len(st.session_state.decisions)
        st.markdown(f'<div class="progress-outer"><div class="progress-inner" style="width:{int(done/total*100)}%"></div></div>', unsafe_allow_html=True)
        st.caption(f"{done} of {total} reviewed")

        if idx < total:
            item  = flagged[idx]
            cands = match_sentence(item["sentence"], st.session_state.vec, st.session_state.mat, st.session_state.refs)
            best  = cands[0]["score"] if cands else 0

            if mode == "Auto-Insert" and best >= TFIDF_THRESHOLD:
                label = author_label(cands[0]["ref"])
                para  = st.session_state.doc_obj.paragraphs[item["para_idx"]]
                insert_superscript(para, item["marker"], label)
                st.session_state.decisions.append({**item,"action":"accepted","ref":cands[0]["ref"],"score":best,"candidates":cands})
                st.session_state.current_idx += 1
                st.rerun()
            elif mode == "Report Only":
                st.session_state.decisions.append({**item,"action":"skipped","candidates":cands})
                st.session_state.current_idx += 1
                st.rerun()
            else:
                st.markdown(f"""
                <div class="match-card">
                    <span class="match-marker">{item['marker']}</span>
                    <div class="match-sentence">"{item['sentence'][:280]}"</div>
                    <div style="font-size:0.72rem;color:#3a4a5a">Para {item['para_idx']+1}</div>
                </div>""", unsafe_allow_html=True)
                st.markdown('<div class="section-label">Top Matches</div>', unsafe_allow_html=True)
                chosen_idx = None
                for j, c in enumerate(cands):
                    ca, cb = st.columns([1,8])
                    with ca:
                        if st.button("Use", key=f"pick_{idx}_{j}"): chosen_idx = j
                    with cb:
                        sc = score_class(c["score"])
                        st.markdown(f'<div style="display:flex;align-items:center;gap:8px;padding:5px 0"><span class="score-pill {sc}">{c["score"]:.3f}</span><span style="font-size:0.83rem;color:#c8d0db">{fmt_ref(c["ref"])}</span></div>', unsafe_allow_html=True)
                    if chosen_idx == j:
                        label = author_label(cands[chosen_idx]["ref"])
                        para  = st.session_state.doc_obj.paragraphs[item["para_idx"]]
                        insert_superscript(para, item["marker"], label)
                        st.session_state.decisions.append({**item,"action":"accepted","ref":cands[chosen_idx]["ref"],"score":cands[chosen_idx]["score"],"candidates":cands})
                        st.session_state.current_idx += 1
                        st.rerun()
                ba, bb, bc = st.columns(3)
                with ba:
                    if st.button("Skip", key=f"skip_{idx}"):
                        st.session_state.decisions.append({**item,"action":"skipped","candidates":cands})
                        st.session_state.current_idx += 1; st.rerun()
                with bb:
                    if use_pubmed and st.button("Search PubMed", key=f"pm_{idx}"):
                        words = [w for w in re.sub(r'[^\w\s]','',item["sentence"]).split() if len(w)>4][:8]
                        with st.spinner("Searching..."):
                            try:
                                r = requests.get(PUBMED_ESEARCH,params=dict(db="pubmed",term=" ".join(words),retmax=PUBMED_MAX,retmode="json"),timeout=8)
                                ids = r.json().get("esearchresult",{}).get("idlist",[])
                                if ids:
                                    r2 = requests.get(PUBMED_ESUM,params=dict(db="pubmed",id=",".join(ids),retmode="json"),timeout=8)
                                    data = r2.json().get("result",{})
                                    for pid in ids:
                                        art = data.get(pid,{}); a = (art.get("authors",[{}])[0]).get("name","Unknown")
                                        st.markdown(f"**{a} ({art.get('pubdate','')[:4]})**. {art.get('title','')}. *{art.get('fulljournalname','')}*")
                                        st.caption(f"[PMID {pid}](https://pubmed.ncbi.nlm.nih.gov/{pid}/)")
                            except: st.warning("PubMed search failed.")
                with bc:
                    if done > 0 and st.button("Finish", key=f"fin_{idx}"):
                        st.session_state.repair_done = True; st.rerun()
        else:
            st.session_state.repair_done = True; st.rerun()

    if st.session_state.repair_done and st.session_state.decisions:
        st.divider()
        decisions = st.session_state.decisions
        accepted = sum(1 for d in decisions if d["action"]=="accepted")
        skipped  = sum(1 for d in decisions if d["action"]=="skipped")
        col1,col2,col3,col4 = st.columns(4)
        col1.metric("Total",len(decisions)); col2.metric("Accepted",accepted)
        col3.metric("Skipped",skipped)
        cd1,cd2 = st.columns(2)
        with cd1:
            st.download_button("⬇ Download repaired document", data=doc_to_bytes(st.session_state.doc_obj),
                file_name="manuscript_repaired.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")
        with cd2:
            st.download_button("⬇ Download decision report", data=write_repair_report(decisions),
                file_name="citation_report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# ─────────────────────────────────────────────────────────────────────────────
# APP 3 — REFERENCE COMPARATOR
# ─────────────────────────────────────────────────────────────────────────────

elif tool == "App 3 — Reference Comparator":
    st.markdown("## Reference List Comparator")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**List A**")
        file_a = st.file_uploader("List A (.xml, .docx, .txt)", type=["xml","docx","txt"], key="comp_a", label_visibility="collapsed")
    with col2:
        st.markdown("**List B**")
        file_b = st.file_uploader("List B (.xml, .docx, .txt)", type=["xml","docx","txt"], key="comp_b", label_visibility="collapsed")

    st.markdown("**Manuscript (optional)** — locates refs in text")
    ms_file = st.file_uploader("Manuscript .docx", type=["docx"], key="comp_ms", label_visibility="collapsed")

    def load_ref_file_st(f):
        name = f.name; data = f.read()
        if name.endswith(".xml"):
            refs = parse_endnote_xml_bytes(data); return refs, name
        elif name.endswith(".docx"):
            doc = Document(io.BytesIO(data)); refs = []; in_refs = False
            pat = re.compile(r'^\s*\d+[\.\)]\s+(.+)')
            for para in doc.paragraphs:
                text = para.text.strip()
                if not text: continue
                if re.match(r'^(references?|bibliography)$', text, re.IGNORECASE): in_refs=True; continue
                if in_refs or pat.match(text):
                    in_refs=True; m=pat.match(text); rt=m.group(1) if m else text
                    ym=re.search(r'\b(19|20)\d{2}\b',rt)
                    refs.append(dict(authors=[],title=rt[:200],journal="",year=ym.group(0) if ym else "",corpus=rt,id=str(len(refs)+1)))
            return refs, name
        elif name.endswith(".txt"):
            content=data.decode("utf-8",errors="replace"); refs=[]
            for i,block in enumerate(re.split(r'\n{2,}',content)):
                block=block.strip()
                if len(block)<20: continue
                m=re.match(r'^\d+[\.\)]\s+(.*)',block,re.DOTALL); text=m.group(1) if m else block
                ym=re.search(r'\b(19|20)\d{2}\b',text)
                refs.append(dict(authors=[],title=text[:200],journal="",year=ym.group(0) if ym else "",corpus=text,id=str(i)))
            return refs, name
        return [], name

    comp_btn = st.button("Compare Lists", type="primary", disabled=not (file_a and file_b))

    if comp_btn:
        with st.spinner("Comparing..."):
            refs_a, label_a = load_ref_file_st(file_a)
            refs_b, label_b = load_ref_file_st(file_b)
            if refs_a and refs_b:
                all_c = [r["corpus"] for r in refs_a] + [r["corpus"] for r in refs_b]
                vec = TfidfVectorizer(ngram_range=(1,2), sublinear_tf=True, max_features=50000)
                vec.fit(all_c)
                emb_a = vec.transform([r["corpus"] for r in refs_a])
                emb_b = vec.transform([r["corpus"] for r in refs_b])
                matrix = cosine_similarity(emb_a, emb_b)
                matched, only_a, fuzzy, matched_b = [], [], [], set()
                for i, ra in enumerate(refs_a):
                    bj=int(matrix[i].argmax()); bs=float(matrix[i][bj])
                    if bs>=MATCH_THRESHOLD: matched.append((ra,refs_b[bj],bs)); matched_b.add(bj)
                    elif bs>=FUZZY_THRESHOLD: fuzzy.append((ra,refs_b[bj],bs))
                    else: only_a.append(ra)
                only_b = [refs_b[j] for j in range(len(refs_b)) if j not in matched_b]
                result = dict(matched=matched,only_in_a=only_a,only_in_b=only_b,fuzzy=fuzzy)
                usage = {}
                if ms_file:
                    ms_doc = Document(io.BytesIO(ms_file.read()))
                    ms_paras = [p.text for p in ms_doc.paragraphs if p.text.strip()]
                    for ref in only_a+only_b+[x[0] for x in fuzzy]:
                        words=[w for w in ref["title"].split() if len(w)>5][:5]
                        found=[p[:200] for p in ms_paras if sum(1 for w in words if w.lower() in p.lower())>=min(3,len(words))]
                        usage[ref.get("id","")]=found
                st.session_state.update(dict(comp_result=result,comp_usage=usage,comp_labels=(label_a,label_b),comp_refs=(refs_a,refs_b)))

    if st.session_state.comp_result:
        result=st.session_state.comp_result; usage=st.session_state.comp_usage
        label_a,label_b=st.session_state.comp_labels; refs_a,refs_b=st.session_state.comp_refs
        st.divider()
        col1,col2,col3,col4=st.columns(4)
        col1.metric("Matched",len(result["matched"]))
        col2.metric("Only in A",len(result["only_in_a"]))
        col3.metric("Only in B",len(result["only_in_b"]))
        col4.metric("Review needed",len(result["fuzzy"]))
        tabs=st.tabs([f"Only in A ({len(result['only_in_a'])})",f"Only in B ({len(result['only_in_b'])})",
                      f"Review ({len(result['fuzzy'])})",f"Matched ({len(result['matched'])})"])
        def ref_card_st(ref,color,locs=None):
            loc_html=""
            if locs:
                loc_html=f'<div style="margin-top:6px;font-size:0.75rem;color:#4fc3f7">Found in {len(locs)} paragraph(s)</div>'
                for loc in locs[:2]: loc_html+=f'<div style="font-size:0.74rem;color:#607080;font-style:italic">{loc[:150]}...</div>'
            st.markdown(f'<div class="ref-item {color}">{fmt_ref(ref)}{loc_html}</div>', unsafe_allow_html=True)
        with tabs[0]:
            if result["only_in_a"]:
                st.caption(f"In **{label_a}** but not **{label_b}**")
                for ref in result["only_in_a"]: ref_card_st(ref,"missing",usage.get(ref.get("id",""),[]))
            else: st.success("None")
        with tabs[1]:
            if result["only_in_b"]:
                st.caption(f"In **{label_b}** but not **{label_a}**")
                for ref in result["only_in_b"]: ref_card_st(ref,"warn",usage.get(ref.get("id",""),[]))
            else: st.success("None")
        with tabs[2]:
            for ra,rb,score in result["fuzzy"]:
                st.markdown(f'<div class="ref-item warn"><b>[{score:.3f}]</b> A: {fmt_ref(ra,True)}<br>B: {fmt_ref(rb,True)}</div>', unsafe_allow_html=True)
        with tabs[3]:
            for ra,rb,score in result["matched"]:
                st.markdown(f'<div class="ref-item ok">[{score:.3f}] {fmt_ref(ra,True)}</div>', unsafe_allow_html=True)

