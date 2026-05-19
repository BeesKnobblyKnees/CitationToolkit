"""
Citation Toolkit v3 — Complete
All 8 apps in one file.
Run: streamlit run app_v2_new.py
"""
import base64
import html as html_module
import io
import re
import sqlite3
import tempfile
import os
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

st.set_page_config(page_title="Citation Toolkit", page_icon="📚",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');
html,body,[class*="css"]{font-family:'IBM Plex Sans',sans-serif;}
section[data-testid="stSidebar"]{background:#0f1117;border-right:1px solid #1e2530;}
section[data-testid="stSidebar"] *{color:#c8d0db !important;}
.main .block-container{padding-top:2rem;max-width:1000px;}
.section-label{font-family:'IBM Plex Mono',monospace;font-size:0.7rem;letter-spacing:0.18em;
  text-transform:uppercase;color:#607080;border-bottom:1px solid #1e2530;padding-bottom:0.4rem;margin-bottom:1rem;}
.step-card{background:#141820;border:1px solid #1e2a38;border-radius:6px;padding:1.2rem 1.4rem;margin-bottom:1rem;}
.step-card.active{border-left:4px solid #4fc3f7;}
.step-card.done{border-left:4px solid #43a047;background:#101a10;}
.step-card.waiting{border-left:4px solid #2a3a4a;opacity:0.6;}
.step-num{font-family:'IBM Plex Mono',monospace;font-size:0.7rem;background:#1e2a38;
  color:#4fc3f7;padding:0.15rem 0.6rem;border-radius:3px;}
.step-num.done{background:#1b3a2a;color:#66bb6a;}
.step-title{font-size:0.95rem;font-weight:500;color:#c8d0db;}
.step-desc{font-size:0.83rem;color:#607080;line-height:1.6;margin-bottom:0.6rem;}
.step-action{font-size:0.8rem;color:#90a0b0;background:#0d1420;border:1px solid #1e2a38;
  border-radius:4px;padding:0.5rem 0.8rem;margin-top:0.5rem;font-family:'IBM Plex Mono',monospace;}
.ref-item{padding:0.5rem 0.8rem;border-radius:4px;margin-bottom:4px;font-size:0.82rem;border-left:3px solid;}
.ref-item.ok{border-color:#43a047;background:#0d1a0d;color:#90b090;}
.ref-item.missing{border-color:#e53935;background:#1a0d0d;color:#c09090;}
.ref-item.warn{border-color:#ffb300;background:#1a1500;color:#c0a060;}
.match-card{background:#141820;border:1px solid #1e2a38;border-left:3px solid #4fc3f7;
  border-radius:4px;padding:1rem 1.2rem;margin-bottom:0.8rem;}
.match-card.accepted{border-left-color:#43a047;background:#101a10;}
.match-card.skipped{border-left-color:#e53935;background:#1a1010;}
.match-sentence{font-size:0.85rem;color:#90a0b0;font-style:italic;margin-bottom:0.5rem;line-height:1.6;}
.match-marker{font-family:'IBM Plex Mono',monospace;font-size:0.75rem;background:#1e2a38;
  color:#4fc3f7;padding:0.1rem 0.5rem;border-radius:3px;display:inline-block;margin-bottom:0.5rem;}
.score-pill{font-family:'IBM Plex Mono',monospace;font-size:0.7rem;padding:0.1rem 0.5rem;
  border-radius:2px;min-width:3.5rem;text-align:center;}
.score-high{background:#1b3a2a;color:#66bb6a;}
.score-mid{background:#2a2a1a;color:#ffca28;}
.score-low{background:#2a1a1a;color:#ef5350;}
.progress-outer{height:4px;background:#1e2530;border-radius:2px;margin-bottom:1rem;overflow:hidden;}
.progress-inner{height:100%;background:#4fc3f7;transition:width 0.3s;}
.instruction-box{background:#0d1525;border:1px solid #1e3050;border-radius:6px;
  padding:1rem 1.2rem;margin:0.8rem 0;font-size:0.85rem;color:#90a8c8;line-height:1.8;}
.instruction-box b{color:#c8d0db;}
.instruction-box code{background:#1a2535;padding:0.1rem 0.4rem;border-radius:3px;
  font-family:'IBM Plex Mono',monospace;font-size:0.8rem;color:#4fc3f7;}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
TOP_N = 5
TFIDF_THRESHOLD = 0.12
PUBMED_MAX = 5
MISSING_PATTERNS = [
    r'\[CITATION\]', r'\[REF\]', r'\[ref\]', r'\[\?\]',
    r'\[citation needed\]', r'\bXXX\b', r'\[#\]', r'<citation>', r'\[ *\]',
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
    return " ".join(p.strip() for p in ([node.text or ""] + [c.text or "" for c in node]) if p.strip())

def fmt_ref(ref, short=False):
    aa = ref.get("authors", [])
    if aa:
        a = (aa[0].split(",")[0] if len(aa)==1 else
             f"{aa[0].split(',')[0]} & {aa[1].split(',')[0]}" if len(aa)==2 else
             f"{aa[0].split(',')[0]} et al.")
    else: a = "Unknown"
    y = ref.get("year","n.d.")
    t = ref.get("title","")[:90] + ("..." if len(ref.get("title",""))>90 else "")
    return f"{a} ({y}) — {t}" if short else f"{a} ({y}). {t}. {ref.get('journal','')}"

def score_class(s):
    return "score-high" if s>=0.20 else "score-mid" if s>=0.10 else "score-low"

def doc_to_bytes(doc):
    buf = io.BytesIO(); doc.save(buf); buf.seek(0); return buf.read()

def _enl_table(cursor):
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
    tables = [t[0] for t in cursor.fetchall()]
    return 'refs' if 'refs' in tables else 'enl_refs'

# ─────────────────────────────────────────────────────────────────────────────
# APP 2 LOGIC — BROKEN CITATION FIXER
# ─────────────────────────────────────────────────────────────────────────────
def analyze_docx_citations(docx_bytes):
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        raw = z.read('word/document.xml').decode('utf-8')
    total    = raw.count('ADDIN EN.CITE')
    working  = len(re.findall(r'&lt;EndNote&gt;', raw))
    empty    = len(re.findall(r'<w:instrText[^>]*> ADDIN EN\.CITE </w:instrText>', raw))
    flddata  = raw.count('<w:fldData')
    db_ids   = list(set(re.findall(r'&lt;key[^&]*db-id=&quot;([^&]+)&quot;', raw)))
    return dict(raw=raw, total_fields=total, working=working,
                broken_empty=empty, flddata_count=flddata, db_ids=db_ids)

def fix_broken_fields(raw_xml):
    pattern = re.compile(
        r'(<w:instrText[^>]*>) ADDIN EN\.CITE (</w:instrText>)'
        r'([\s\S]{0,2000}?)<w:fldData[^>]*>([\s\S+?]+?)</w:fldData>', re.DOTALL)
    fixes = 0
    def rep(m):
        nonlocal fixes
        io_, ic, between, b64r = m.group(1),m.group(2),m.group(3),m.group(4)
        b64 = b64r.replace('\r','').replace('\n','').replace(' ','')
        pad = (4-len(b64)%4)%4
        try:
            dec = base64.b64decode(b64+'='*pad).decode('utf-8',errors='replace')
            dec = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]','',dec)
            esc = dec.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')
            fixes += 1
            return f'{io_} ADDIN EN.CITE {esc}{ic}{between}<w:fldData xml:space="preserve">{b64r}</w:fldData>'
        except: return m.group(0)
    result = pattern.sub(rep, raw_xml)
    result = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]','',result)
    return result, fixes

def extract_karol_db_id(xml_bytes):
    try:
        content = xml_bytes.decode('utf-8',errors='replace')
        m = re.search(r'db-id="([a-z0-9]{20,45})"',content,re.IGNORECASE)
        return m.group(1) if m else None
    except: return None

def get_karol_rec_nums(enl_bytes):
    with tempfile.NamedTemporaryFile(suffix='.enl',delete=False) as f:
        f.write(enl_bytes); tmp=f.name
    try:
        conn=sqlite3.connect(tmp); cursor=conn.cursor()
        tbl=_enl_table(cursor)
        cursor.execute(f'SELECT id,author,year,title FROM {tbl} WHERE trash_state=0 OR trash_state IS NULL')
        rows=cursor.fetchall(); conn.close()
        return {str(r[0]):{'id':str(r[0]),'author':r[1] or '','year':str(r[2] or ''),'title':r[3] or ''} for r in rows}
    except: return {}
    finally: os.unlink(tmp)

def check_missing_from_karol(raw_xml, karol_ids):
    all_rns = set(re.findall(r'&lt;RecNum&gt;(\d+)&lt;/RecNum&gt;', raw_xml))
    return [rn for rn in all_rns if rn not in karol_ids]

def patch_db_ids(raw_xml, old_ids, new_id):
    result=raw_xml; replaced=0
    for old in old_ids:
        if old!=new_id:
            count=result.count(old); result=result.replace(old,new_id); replaced+=count
    return result, replaced

def build_fixed_docx(original_bytes, fixed_xml):
    with zipfile.ZipFile(io.BytesIO(original_bytes)) as z:
        all_files={n:z.read(n) for n in z.namelist()}
    all_files['word/document.xml']=fixed_xml.encode('utf-8')
    buf=io.BytesIO()
    with zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as zout:
        for n,d in all_files.items(): zout.writestr(n,d)
    buf.seek(0); return buf.read()

def extract_traveling_library_xml(docx_bytes):
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        raw=z.read('word/document.xml').decode('utf-8')
    all_cx=[]
    for m in re.findall(r'ADDIN EN\.CITE &lt;EndNote&gt;([\s\S]+?)&lt;/EndNote&gt;',raw):
        all_cx.append(html_module.unescape(f'<EndNote>{m}</EndNote>'))
    for b64r in re.findall(r'<w:fldData[^>]*>([\s\S+?]+?)</w:fldData>',raw):
        b64=b64r.replace('\r','').replace('\n','').replace(' ','')
        pad=(4-len(b64)%4)%4
        try:
            dec=base64.b64decode(b64+'='*pad).decode('utf-8',errors='replace').replace('\x00','')
            if '<EndNote>' in dec: all_cx.append(dec)
        except: pass
    from lxml import etree
    import html as hmod
    traveling={}
    for cx in all_cx:
        try:
            if not cx.startswith('<EndNote>'): cx=f'<EndNote>{cx}</EndNote>'
            root=etree.fromstring(cx.encode('utf-8'))
            for cite in root.findall('.//Cite'):
                rn=cite.findtext('RecNum') or ''
                rec=cite.find('record')
                if rn and rec is not None and rn not in traveling:
                    traveling[rn]=rec
        except: pass
    def gt(elem,tag):
        n=elem.find(f'.//{tag}')
        return ''.join(n.itertext()).strip() if n is not None else ''
    output='<?xml version="1.0" encoding="UTF-8"?>\n<xml>\n  <records>\n'
    for rn,record in sorted(traveling.items(),key=lambda x:int(x[0]) if x[0].isdigit() else 9999):
        rte=record.find('.//ref-type')
        rtn=rte.get('name','Journal Article') if rte is not None else 'Journal Article'
        rtv=rte.text if rte is not None else '17'
        authors=[''.join(a.itertext()).strip() for a in record.findall('.//contributors/authors/author')]
        secauths=[''.join(a.itertext()).strip() for a in record.findall('.//contributors/secondary-authors/author')]
        r=f'    <record>\n      <rec-number>{rn}</rec-number>\n      <ref-type name="{hmod.escape(rtn)}">{rtv}</ref-type>\n'
        if authors:
            r+='      <contributors>\n        <authors>\n'
            for a in authors:
                if a: r+=f'          <author>{hmod.escape(a)}</author>\n'
            r+='        </authors>\n'
            if secauths:
                r+='        <secondary-authors>\n'
                for a in secauths:
                    if a: r+=f'          <author>{hmod.escape(a)}</author>\n'
                r+='        </secondary-authors>\n'
            r+='      </contributors>\n'
        r+='      <titles>\n'
        for tag,xmltag in [('title','title'),('secondary-title','secondary-title'),('tertiary-title','tertiary-title')]:
            v=gt(record,tag)
            if v: r+=f'        <{xmltag}>{hmod.escape(v)}</{xmltag}>\n'
        r+='      </titles>\n'
        for tag,xmltag in [('year','dates><year'),('volume','volume'),('number','number'),
                           ('pages','pages'),('edition','edition'),('publisher','publisher'),
                           ('pub-location','pub-location'),('abstract','abstract')]:
            v=gt(record,tag)
            if v:
                if xmltag=='dates><year': r+=f'      <dates><year>{hmod.escape(v)}</year></dates>\n'
                else: r+=f'      <{xmltag}>{hmod.escape(v)}</{xmltag}>\n'
        kws=[(''.join(k.itertext()).strip()) for k in record.findall('.//keyword')]
        if any(kws):
            r+='      <keywords>\n'
            for kw in kws:
                if kw: r+=f'        <keyword>{hmod.escape(kw)}</keyword>\n'
            r+='      </keywords>\n'
        r+='    </record>\n'
        output+=r
    output+='  </records>\n</xml>\n'
    return output, len(traveling)

def remap_traveling_citations(docx_bytes, enl_bytes):
    from lxml import etree
    import html as hm
    with tempfile.NamedTemporaryFile(suffix='.enl',delete=False) as f:
        f.write(enl_bytes); tmp=f.name
    try:
        conn=sqlite3.connect(tmp); cursor=conn.cursor()
        tbl=_enl_table(cursor)
        cursor.execute(f'SELECT id,author,year,title FROM {tbl} WHERE trash_state=0 OR trash_state IS NULL')
        karol_rows=cursor.fetchall(); conn.close()
    finally: os.unlink(tmp)
    def norm(s): return re.sub(r'[^a-z0-9]','',s.lower()) if s else ''
    def alast(s): return norm(s.split('\r')[0].split('\n')[0].strip().split(',')[0]) if s else ''
    karol_ids=set(str(r[0]) for r in karol_rows)
    by_ay={}
    for row in karol_rows:
        k=(alast(row[1] or ''),norm(str(row[2] or '')))
        by_ay.setdefault(k,[]).append(row)
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        raw=z.read('word/document.xml').decode('utf-8')
        all_files={n:z.read(n) for n in z.namelist()}
    report=[]; remap={}
    for cx_esc in re.findall(r'ADDIN EN\.CITE &lt;EndNote&gt;([\s\S]+?)&lt;/EndNote&gt;',raw):
        cx=hm.unescape(f'<EndNote>{cx_esc}</EndNote>')
        try:
            root=etree.fromstring(cx.encode('utf-8'))
            for cite in root.findall('.//Cite'):
                rn=cite.findtext('RecNum') or ''
                au=cite.findtext('Author') or ''
                yr=cite.findtext('Year') or ''
                if rn in karol_ids or rn in remap: continue
                title=''
                rec=cite.find('record')
                if rec is not None:
                    te=rec.find('.//title')
                    if te is not None: title=''.join(te.itertext())
                key=(alast(au),norm(yr)); cands=by_ay.get(key,[])
                matched=None
                if len(cands)==1: matched=str(cands[0][0])
                elif len(cands)>1:
                    for c in cands:
                        if norm(title)[:25] and norm(str(c[3] or ''))[:25]==norm(title)[:25]:
                            matched=str(c[0]); break
                    if not matched: matched=str(cands[0][0])
                remap[rn]=matched
                report.append({'status':'remapped' if matched else 'not_found',
                               'old_rec_num':rn,'new_rec_num':matched,
                               'author':au,'year':yr,'title':title[:80]})
        except: pass
    if not remap: return docx_bytes, report
    fixed=raw
    for old,new in remap.items():
        if new:
            fixed=fixed.replace(f'&lt;RecNum&gt;{old}&lt;/RecNum&gt;',
                                 f'&lt;RecNum&gt;{new}&lt;/RecNum&gt;')
    def fix_fld(m):
        b64r=m.group(1)
        b64=b64r.replace('\r','').replace('\n','').replace(' ','')
        pad=(4-len(b64)%4)%4
        try:
            dec=base64.b64decode(b64+'='*pad).decode('utf-8',errors='replace').replace('\x00','')
            mod=dec
            for old,new in remap.items():
                if new: mod=mod.replace(f'<RecNum>{old}</RecNum>',f'<RecNum>{new}</RecNum>')
            if mod!=dec:
                nb64=base64.b64encode(mod.encode('utf-8')).decode('ascii')
                wrapped='\r\n'.join(nb64[i:i+76] for i in range(0,len(nb64),76))
                return f'<w:fldData xml:space="preserve">{wrapped}</w:fldData>'
        except: pass
        return m.group(0)
    fixed=re.compile(r'<w:fldData[^>]*>([\s\S+?]+?)</w:fldData>').sub(fix_fld,fixed)
    all_files['word/document.xml']=fixed.encode('utf-8')
    buf=io.BytesIO()
    with zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as zout:
        for n,d in all_files.items(): zout.writestr(n,d)
    buf.seek(0)
    return buf.read(), report

def generate_vba_macro(doc_name):
    return f"""' EndNote Citation Relinker — Word VBA Macro
' Generated: {datetime.now():%Y-%m-%d %H:%M}
' HOW TO USE:
'  1. Open "{doc_name}" in Word with KAROL open in EndNote
'  2. Press Alt+F11 → Insert → Module → paste this macro
'  3. Press Alt+F8 → select RelinkAllCitations → Run
Sub RelinkAllCitations()
    Dim oDoc As Document
    Set oDoc = ActiveDocument
    If oDoc Is Nothing Then MsgBox "No document open.", vbCritical: Exit Sub
    Dim sBackup As String
    sBackup = oDoc.Path & "\\{Path(doc_name).stem}_BACKUP_" & Format(Now,"YYYYMMDD_HHMMSS") & ".docx"
    oDoc.SaveAs2 sBackup, wdFormatXMLDocument
    MsgBox "Backup saved: " & sBackup, vbInformation
    On Error Resume Next
    Application.Run "EndNote.UnformatAll"
    If Err.Number <> 0 Then Err.Clear: Application.Run "EndNote.FormatAll": GoTo done
    Err.Clear
    Application.Run "EndNote.FormatAll"
    If Err.Number <> 0 Then Err.Clear: Application.Run "EndNote.UpdateAll"
    done:
    On Error GoTo 0
    oDoc.Save
    MsgBox "Done! Citations re-linked to KAROL." & vbCrLf & _
           "If any citations are yellow, right-click → Edit Citation → Find.", vbInformation
End Sub
"""

# ─────────────────────────────────────────────────────────────────────────────
# APP 1 LOGIC — CITATION REPAIR
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def parse_endnote_xml_bytes(xml_bytes):
    root=ET.fromstring(xml_bytes); refs=[]
    for rec in root.iter("record"):
        authors=[]
        for a in rec.findall(".//contributors/authors/author"):
            name=" ".join(p.strip() for p in ([a.text or ""]+[c.text or "" for c in a]) if p.strip())
            if name: authors.append(name)
        title=xml_text(rec,".//titles/title")
        journal=xml_text(rec,".//periodical/full-title") or xml_text(rec,".//periodical/abbr-1")
        year=xml_text(rec,".//dates/year"); abstract=xml_text(rec,".//abstract")
        if not title: continue
        corpus=" ".join(filter(None,[title,abstract,journal,year]))
        refs.append(dict(authors=authors,title=title,journal=journal,year=year,corpus=corpus))
    return refs

@st.cache_resource(show_spinner=False)
def build_tfidf(corpora_tuple):
    vec=TfidfVectorizer(ngram_range=(1,2),sublinear_tf=True,max_features=50000)
    mat=vec.fit_transform(list(corpora_tuple)); return vec,mat

def match_sentence(sentence,vec,mat,refs,top_n=TOP_N):
    sv=vec.transform([sentence]); sims=cosine_similarity(sv,mat)[0]
    idx=sims.argsort()[::-1][:top_n]
    return [dict(ref=refs[i],score=float(sims[i])) for i in idx]

def extract_flagged(docx_bytes):
    doc=Document(io.BytesIO(docx_bytes)); flagged=[]
    for pi,para in enumerate(doc.paragraphs):
        text=para.text
        if not text.strip(): continue
        for m in CITATION_MARKERS.finditer(text):
            sents=re.split(r'(?<=[.!?])\s+',text); cum,target=0,text
            for s in sents:
                if cum+len(s)>=m.start(): target=s; break
                cum+=len(s)+1
            flagged.append(dict(para_idx=pi,sentence=target,marker=m.group(),para_text=text))
    return flagged,doc

def author_label(ref):
    aa=ref.get("authors",[]); last=aa[0].split(",")[0].strip().split()[-1] if aa else "Ref"
    return (last+" "+ref.get("year","")).strip()

def insert_superscript(para,marker,label):
    if marker not in para.text: return False
    combined,run_map="",[]
    for run in para.runs:
        s=len(combined); combined+=run.text; run_map.append((s,s+len(run.text),run))
    pos=combined.find(marker)
    if pos==-1: return False
    for (s,e,run) in run_map:
        if s<=pos<e:
            before=run.text[:pos-s]; after=run.text[pos-s+len(marker):]
            run.text=before
            rPr=OxmlElement("w:rPr"); va=OxmlElement("w:vertAlign")
            va.set(qn("w:val"),"superscript"); rPr.append(va)
            nr=OxmlElement("w:r"); nr.append(deepcopy(rPr))
            t=OxmlElement("w:t"); t.text=f"[{label}]"
            t.set("{http://www.w3.org/XML/1998/namespace}space","preserve"); nr.append(t)
            run._r.addnext(nr)
            if after:
                tr=OxmlElement("w:r"); tt=OxmlElement("w:t")
                tt.text=after; tt.set("{http://www.w3.org/XML/1998/namespace}space","preserve")
                tr.append(tt); nr.addnext(tr)
            return True
    return False

def write_repair_report(decisions):
    doc=Document(); doc.add_heading("Citation Repair Report",0)
    doc.add_paragraph(f"Generated: {datetime.now():%Y-%m-%d %H:%M}")
    accepted=[d for d in decisions if d["action"]=="accepted"]
    skipped=[d for d in decisions if d["action"]=="skipped"]
    doc.add_paragraph(f"Total:{len(decisions)} | Accepted:{len(accepted)} | Skipped:{len(skipped)}")
    if accepted:
        doc.add_heading("Accepted",1)
        for d in accepted:
            p=doc.add_paragraph(style="List Bullet")
            p.add_run(f"Marker:{d['marker']}\n").bold=True
            p.add_run(f"Context:{d['sentence'][:200]}\n")
            p.add_run(f"Inserted:{fmt_ref(d['ref'])}\nScore:{d['score']:.3f}")
    if skipped:
        doc.add_heading("Skipped",1)
        for d in skipped:
            p=doc.add_paragraph(style="List Bullet")
            r=p.add_run("NEEDS REVIEW\n"); r.font.color.rgb=RGBColor(0xC0,0,0)
            p.add_run(f"Marker:{d['marker']}\nContext:{d['sentence'][:200]}\n")
    return doc_to_bytes(doc)

# ─────────────────────────────────────────────────────────────────────────────
# APP 3 LOGIC — REFERENCE COMPARATOR
# ─────────────────────────────────────────────────────────────────────────────
def load_ref_file(f):
    name=f.name; data=f.read()
    if name.endswith(".xml"):
        refs=parse_endnote_xml_bytes(data); return refs,name
    elif name.endswith(".docx"):
        doc=Document(io.BytesIO(data)); refs=[]; in_refs=False
        pat=re.compile(r'^\s*\d+[\.\)]\s+(.+)')
        for para in doc.paragraphs:
            text=para.text.strip()
            if not text: continue
            if re.match(r'^(references?|bibliography)$',text,re.IGNORECASE): in_refs=True; continue
            if in_refs or pat.match(text):
                in_refs=True; m=pat.match(text); rt=m.group(1) if m else text
                ym=re.search(r'\b(19|20)\d{2}\b',rt)
                refs.append(dict(authors=[],title=rt[:200],journal="",
                                 year=ym.group(0) if ym else "",corpus=rt,id=str(len(refs)+1)))
        return refs,name
    elif name.endswith(".txt"):
        content=data.decode("utf-8",errors="replace"); refs=[]
        for i,block in enumerate(re.split(r'\n{2,}',content)):
            block=block.strip()
            if len(block)<20: continue
            m=re.match(r'^\d+[\.\)]\s+(.*)',block,re.DOTALL); text=m.group(1) if m else block
            ym=re.search(r'\b(19|20)\d{2}\b',text)
            refs.append(dict(authors=[],title=text[:200],journal="",
                             year=ym.group(0) if ym else "",corpus=text,id=str(i)))
        return refs,name
    return [],name

# ─────────────────────────────────────────────────────────────────────────────
# APP 4 LOGIC — DOCUMENT MERGER
# ─────────────────────────────────────────────────────────────────────────────
def analyze_merge_damage(merged_bytes):
    with zipfile.ZipFile(io.BytesIO(merged_bytes)) as z:
        raw=z.read('word/document.xml').decode('utf-8')
    total=raw.count('ADDIN EN.CITE')
    working=len(re.findall(r'&lt;EndNote&gt;',raw))
    empty=len(re.findall(r'<w:instrText[^>]*> ADDIN EN\.CITE </w:instrText>',raw))
    begins=raw.count('fldCharType="begin"')
    separates=raw.count('fldCharType="separate"')
    ends=raw.count('fldCharType="end"')
    has_tracked='<w:ins ' in raw or '<w:del ' in raw
    ins_count=raw.count('<w:ins ')
    del_count=raw.count('<w:del ')
    db_ids=list(set(re.findall(r'&lt;key[^&]*db-id=&quot;([^&]+)&quot;',raw)))
    return dict(raw=raw,total_en=total,with_data=working,empty_cite=empty,
                begins=begins,separates=separates,ends=ends,
                balanced=(begins==separates==ends),
                has_tracked=has_tracked,ins_count=ins_count,del_count=del_count,
                db_ids=db_ids)

def repair_post_merge_citations(merged_bytes,original_bytes=None):
    analysis=analyze_merge_damage(merged_bytes)
    raw=analysis['raw']
    with zipfile.ZipFile(io.BytesIO(merged_bytes)) as z:
        all_files={n:z.read(n) for n in z.namelist()}
    report=dict(analysis=analysis,steps=[],citations_before=analysis['with_data'])
    fixed=raw
    # Accept tracked changes safely
    if analysis['has_tracked']:
        def rescue_del(m):
            dc=m.group(0)
            if 'ADDIN EN.CITE' in dc:
                runs=re.findall(r'<w:r[^>]*>.*?</w:r>',dc,re.DOTALL)
                cite_runs=[r.replace('<w:delText','<w:t').replace('</w:delText>','</w:t>')
                           for r in runs if any(x in r for x in ['fldChar','instrText','fldData','ADDIN'])]
                if cite_runs: return ''.join(cite_runs)
            return ''
        fixed=re.compile(r'<w:del\b[^>]*>[\s\S]*?</w:del>',re.DOTALL).sub(rescue_del,fixed)
        def accept_ins(m):
            inner=re.sub(r'^<w:ins[^>]*>','',m.group(0)); inner=re.sub(r'</w:ins>$','',inner)
            return inner
        fixed=re.compile(r'<w:ins\b[^>]*>[\s\S]*?</w:ins>',re.DOTALL).sub(accept_ins,fixed)
        report['steps'].append('track_changes_accepted')
    # Restore broken fields
    if analysis['empty_cite']>0:
        fixed,n_fixed=fix_broken_fields(fixed)
        if n_fixed>0: report['steps'].append(f'restored_{n_fixed}_fields')
    # Compare against original
    if original_bytes:
        with zipfile.ZipFile(io.BytesIO(original_bytes)) as z:
            orig_raw=z.read('word/document.xml').decode('utf-8')
        orig_rns=set(re.findall(r'&lt;RecNum&gt;(\d+)&lt;/RecNum&gt;',orig_raw))
        merged_rns=set(re.findall(r'&lt;RecNum&gt;(\d+)&lt;/RecNum&gt;',fixed))
        lost=orig_rns-merged_rns
        report['lost_rec_nums']=list(lost)
        if lost: report['steps'].append(f'{len(lost)}_citations_lost')
    report['citations_after']=len(re.findall(r'&lt;EndNote&gt;',fixed))
    fixed=fixed.replace('\x00','')
    all_files['word/document.xml']=fixed.encode('utf-8')
    buf=io.BytesIO()
    with zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as zout:
        for n,d in all_files.items(): zout.writestr(n,d)
    buf.seek(0)
    return buf.read(),report

# ─────────────────────────────────────────────────────────────────────────────
# APP 5 LOGIC — CITATION RENUMBERING
# ─────────────────────────────────────────────────────────────────────────────
def renumber_citations(docx_bytes):
    doc=Document(io.BytesIO(docx_bytes))
    seen={}
    for para in doc.paragraphs:
        for run in para.runs:
            rpr=run._r.find(qn('w:rPr'))
            if rpr is None: continue
            va=rpr.find(qn('w:vertAlign'))
            if va is None or va.get(qn('w:val'))!='superscript': continue
            text=run.text.strip()
            parts=[p.strip() for p in re.split(r'[,;]',text) if p.strip().isdigit()]
            for ps in parts:
                num=int(ps)
                if num not in seen: seen[num]=len(seen)+1
    if not seen: return docx_bytes,{}
    for para in doc.paragraphs:
        for run in para.runs:
            rpr=run._r.find(qn('w:rPr'))
            if rpr is None: continue
            va=rpr.find(qn('w:vertAlign'))
            if va is None or va.get(qn('w:val'))!='superscript': continue
            text=run.text; sep=';' if ';' in text else ','
            parts=re.split(r'[,;]',text)
            new_parts=[str(seen.get(int(p.strip()),int(p.strip()))) if p.strip().isdigit() else p for p in parts]
            run.text=sep.join(new_parts)
    ref_pat=re.compile(r'^\s*(\d+)\.\s+')
    bib=[(pi,int(ref_pat.match(para.text).group(1))) for pi,para in enumerate(doc.paragraphs) if ref_pat.match(para.text)]
    if bib:
        for pi,old_num in bib:
            new_num=seen.get(old_num)
            if new_num is None: continue
            for run in doc.paragraphs[pi].runs:
                if f"{old_num}." in run.text:
                    run.text=run.text.replace(f"{old_num}.",f"{new_num}.",1); break
        if len(bib)>1:
            W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            body=doc.element.body; all_p=list(body.findall(f'{{{W}}}p'))
            bib_data=sorted([(seen.get(old,old),all_p[pi]) for pi,old in bib],key=lambda x:x[0])
            anchor=all_p[bib[0][0]]
            for _,elem in bib_data: anchor.addprevious(elem)
    buf=io.BytesIO(); doc.save(buf); buf.seek(0); return buf.read(),seen

# ─────────────────────────────────────────────────────────────────────────────
# APP 6 LOGIC — FIGURE INVENTORY
# ─────────────────────────────────────────────────────────────────────────────
def scan_figures(docx_bytes):
    doc=Document(io.BytesIO(docx_bytes)); items=[]
    cap_pats=[(re.compile(r'^\s*Fig(?:ure)?\.?\s*(\d+[\w\-\.]*)',re.IGNORECASE),'Figure'),
              (re.compile(r'^\s*Table\s*(\d+[\w\-\.]*)',re.IGNORECASE),'Table'),
              (re.compile(r'^\s*Box\s*(\d+[\w\-\.]*)',re.IGNORECASE),'Box'),
              (re.compile(r'^\s*Plate\s*(\d+[\w\-\.]*)',re.IGNORECASE),'Plate'),
              (re.compile(r'^\s*Video\s*(\d+[\w\-\.]*)',re.IGNORECASE),'Video'),
              (re.compile(r'^\s*Appendix\s*(\d+[\w\-\.]*)',re.IGNORECASE),'Appendix')]
    DN='http://schemas.openxmlformats.org/drawingml/2006/main'
    for pi,para in enumerate(doc.paragraphs):
        text=para.text.strip()
        has_img=para._p.find(f'.//{{{DN}}}blip') is not None
        style=para.style.name if para.style else ''
        itype=None; inum=None
        for pat,pt in cap_pats:
            m=pat.match(text)
            if m: itype=pt; inum=m.group(1); break
        if itype or has_img or 'caption' in style.lower():
            items.append(dict(para_idx=pi,type=itype or ('Image' if has_img else 'Caption'),
                              number=inum or '',caption=text[:250],has_image=has_img,style=style))
    return items

def cross_ref_excel(items, excel_bytes):
    import openpyxl
    try:
        wb=openpyxl.load_workbook(io.BytesIO(excel_bytes)); ws=wb.active
        rows=list(ws.iter_rows(values_only=True))
        if not rows: return [(i,'no_data','') for i in items]
        hdrs=[str(c or '').lower().strip() for c in rows[0]]
        nc=next((i for i,h in enumerate(hdrs) if any(w in h for w in ['num','#'])),None)
        lc=next((i for i,h in enumerate(hdrs) if any(w in h for w in ['name','caption','title','new','label'])),None)
        if lc is None: return [(i,'no_label_col','') for i in items]
        lookup={}
        for row in rows[1:]:
            if not row: continue
            num=str(row[nc] or '').strip() if nc is not None else ''
            label=str(row[lc] or '').strip()
            if num and label: lookup[num.lower()]=label
        results=[]
        for item in items:
            exp=lookup.get(item['number'].lower(),'')
            if not exp: status='not_in_excel'
            elif exp.lower() in item['caption'].lower(): status='match'
            else: status='mismatch'
            results.append((item,status,exp))
        return results
    except Exception as e: return [(i,f'error','') for i in items]

# ─────────────────────────────────────────────────────────────────────────────
# APP 7 LOGIC — PUBMED SEARCH
# ─────────────────────────────────────────────────────────────────────────────
def pubmed_search_full(query,date_from='',date_to='',journal_filter='',max_results=20):
    try:
        term=query
        if date_from and date_to: term+=f' AND {date_from}:{date_to}[dp]'
        elif date_from: term+=f' AND {date_from}:3000[dp]'
        if journal_filter: term+=f' AND "{journal_filter}"[ta]'
        r=requests.get(PUBMED_ESEARCH,params={'db':'pubmed','term':term,
            'retmax':max_results,'retmode':'json','sort':'relevance'},timeout=10)
        ids=r.json().get('esearchresult',{}).get('idlist',[])
        if not ids: return []
        r2=requests.get('https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi',
            params={'db':'pubmed','id':','.join(ids),'retmode':'xml','rettype':'abstract'},timeout=15)
        root=ET.fromstring(r2.content); results=[]
        for art in root.findall('.//PubmedArticle'):
            pmid=art.findtext('.//PMID') or ''
            title=art.findtext('.//ArticleTitle') or ''
            abstract=' '.join(t.text or '' for t in art.findall('.//AbstractText'))
            journal=art.findtext('.//Journal/Title') or art.findtext('.//ISOAbbreviation') or ''
            year=art.findtext('.//PubDate/Year') or (art.findtext('.//PubDate/MedlineDate') or '')[:4]
            volume=art.findtext('.//Volume') or ''; issue=art.findtext('.//Issue') or ''
            pages=art.findtext('.//MedlinePgn') or ''
            doi=next((a.text for a in art.findall('.//ArticleId') if a.get('IdType')=='doi'),'')
            pmc=next((a.text for a in art.findall('.//ArticleId') if a.get('IdType')=='pmc'),'')
            authors=[]
            for auth in art.findall('.//Author')[:6]:
                last=auth.findtext('LastName') or ''; fore=auth.findtext('Initials') or ''
                if last: authors.append(f"{last} {fore}".strip())
            results.append(dict(pmid=pmid,title=title,authors=authors,journal=journal,
                year=year,volume=volume,issue=issue,pages=pages,
                abstract=abstract[:600],doi=doi,
                pubmed_url=f'https://pubmed.ncbi.nlm.nih.gov/{pmid}/',
                pmc_url=f'https://www.ncbi.nlm.nih.gov/pmc/articles/{pmc}/' if pmc else '',
                doi_url=f'https://doi.org/{doi}' if doi else ''))
        return results
    except: return []

def results_to_xml(results):
    import html as hm
    xml='<?xml version="1.0" encoding="UTF-8"?>\n<xml>\n  <records>\n'
    for i,r in enumerate(results,1):
        xml+=f'    <record>\n      <rec-number>{i}</rec-number>\n      <ref-type name="Journal Article">17</ref-type>\n'
        if r['authors']:
            xml+='      <contributors>\n        <authors>\n'
            for a in r['authors']: xml+=f'          <author>{hm.escape(a)}</author>\n'
            xml+='        </authors>\n      </contributors>\n'
        xml+=f'      <titles>\n        <title>{hm.escape(r["title"])}</title>\n        <secondary-title>{hm.escape(r["journal"])}</secondary-title>\n      </titles>\n'
        xml+=f'      <dates><year>{hm.escape(r["year"])}</year></dates>\n'
        for k,t in [('volume','volume'),('issue','number'),('pages','pages')]:
            if r[k]: xml+=f'      <{t}>{hm.escape(r[k])}</{t}>\n'
        if r['abstract']: xml+=f'      <abstract>{hm.escape(r["abstract"])}</abstract>\n'
        if r['doi']: xml+=f'      <electronic-resource-number>{hm.escape(r["doi"])}</electronic-resource-number>\n'
        xml+=f'      <urls><related-urls><url>{hm.escape(r["pubmed_url"])}</url></related-urls></urls>\n'
        xml+='    </record>\n'
    xml+='  </records>\n</xml>\n'; return xml

# ─────────────────────────────────────────────────────────────────────────────
# APP 8 LOGIC — BATCH RENAME
# ─────────────────────────────────────────────────────────────────────────────
def load_rename_pairs(excel_bytes):
    import openpyxl
    try:
        wb=openpyxl.load_workbook(io.BytesIO(excel_bytes)); ws=wb.active
        rows=list(ws.iter_rows(values_only=True))
        if not rows: return []
        hdrs=[str(c or '').lower().strip() for c in rows[0]]
        oc=next((i for i,h in enumerate(hdrs) if any(w in h for w in ['old','current','find','original','from'])),0)
        nc=next((i for i,h in enumerate(hdrs) if any(w in h for w in ['new','replace','final','to','updated'])),1)
        pairs=[]
        for row in rows[1:]:
            if not row or len(row)<=max(oc,nc): continue
            old=str(row[oc] or '').strip(); new=str(row[nc] or '').strip()
            if old and new and old!=new: pairs.append((old,new))
        return pairs
    except: return []

def batch_rename(docx_bytes,pairs,match_case=False,whole_word=False):
    import html as hm
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        raw=z.read('word/document.xml').decode('utf-8')
        all_files={n:z.read(n) for n in z.namelist()}
    report=[]; fixed=raw
    for old,new in pairs:
        oe=hm.escape(old); ne=hm.escape(new)
        flags=0 if match_case else re.IGNORECASE
        pat=re.escape(oe)
        if whole_word: pat=r'(?<![a-zA-Z])'+pat+r'(?![a-zA-Z])'
        count=len(re.findall(pat,fixed,flags=flags))
        if count>0: fixed=re.sub(pat,ne,fixed,flags=flags); report.append({'old':old,'new':new,'count':count,'status':'replaced'})
        else: report.append({'old':old,'new':new,'count':0,'status':'not_found'})
    all_files['word/document.xml']=fixed.encode('utf-8')
    buf=io.BytesIO()
    with zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as zout:
        for n,d in all_files.items(): zout.writestr(n,d)
    buf.seek(0); return buf.read(),report

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────
defaults=dict(flagged=[],current_idx=0,decisions=[],doc_obj=None,repair_done=False,
              refs=[],vec=None,mat=None,
              fix_stage=1,fix_analysis=None,fix_raw_xml=None,fix_docx_bytes=None,
              fix_after_stage1=None,fix_after_stage2=None,fix_karol_db_id=None,
              fix_karol_rec_nums={},fix_missing_refs=[],fix_doc_name="document.docx",
              comp_result=None,comp_usage={},comp_labels=("",""),comp_refs=([],[]))
for k,v in defaults.items():
    if k not in st.session_state: st.session_state[k]=v

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📚 Citation Toolkit")
    st.caption("Pediatric Orthopedics — Tachdjian's")
    st.divider()
    tool=st.radio("Tool",[
        "App 1 — Citation Repair",
        "App 2 — Broken Citation Fixer",
        "App 3 — Reference Comparator",
        "App 4 — Document Merger",
        "App 5 — Citation Renumbering",
        "App 6 — Figure Inventory",
        "App 7 — PubMed Search",
        "App 8 — Batch Rename",
    ],label_visibility="collapsed")
    st.divider()
    st.markdown("""
    <div style="font-size:0.72rem;color:#3a4a5a;line-height:1.8">
    <b style="color:#4fc3f7">App 1</b> — Fill missing citation placeholders<br>
    <b style="color:#4fc3f7">App 2</b> — Fix broken/unrecognized citations<br>
    <b style="color:#4fc3f7">App 3</b> — Compare two reference lists<br>
    <b style="color:#4fc3f7">App 4</b> — Merge documents safely<br>
    <b style="color:#4fc3f7">App 5</b> — Renumber citations sequentially<br>
    <b style="color:#4fc3f7">App 6</b> — Figure/table/plate inventory<br>
    <b style="color:#4fc3f7">App 7</b> — PubMed search + export<br>
    <b style="color:#4fc3f7">App 8</b> — Batch rename from Excel
    </div>
    """,unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# APP 2 UI — BROKEN CITATION FIXER
# ─────────────────────────────────────────────────────────────────────────────
if tool == "App 2 — Broken Citation Fixer":
    st.markdown("## Broken Citation Fixer")
    st.markdown('<div class="instruction-box">Use this when EndNote only recognizes some citations even though the bibliography shows all references — or when citations show as "Traveling Library" instead of your KAROL library.</div>', unsafe_allow_html=True)

    # Step 0 — Extract traveling library
    with st.expander("📥 Step 0 — Extract traveling library references (start here)", expanded=True):
        st.markdown('<div class="step-desc">If EndNote cannot find your refs, extract them from the Word file and import into KAROL.</div>', unsafe_allow_html=True)
        tl_file=st.file_uploader("Word document",type=["docx"],key="tl_doc")
        if tl_file:
            if st.button("Extract references",type="primary",key="tl_run"):
                with st.spinner("Extracting..."):
                    try:
                        tl_xml,tl_count=extract_traveling_library_xml(tl_file.read())
                        st.success(f"Extracted {tl_count} references.")
                        st.download_button(f"⬇ Download EndNote XML ({tl_count} refs)",
                            data=tl_xml.encode('utf-8'),
                            file_name=Path(tl_file.name).stem+"_traveling_library.xml",
                            mime="application/xml",type="primary")
                        st.markdown('<div class="instruction-box">Import into KAROL: <b>File → Import → File</b> → select XML → Import Option: <code>EndNote XML</code> → Duplicates: <code>Discard Duplicates</code> → Import</div>',unsafe_allow_html=True)
                    except Exception as e: st.error(f"Error: {e}")

    st.divider()

    # Step 0b — Remap traveling library citations
    with st.expander("🔄 Step 0b — Remap citations still showing as Traveling Library"):
        st.markdown('<div class="step-desc">After importing the traveling library, some citations may still show as "Traveling Library" because their RecNums don\'t match KAROL\'s record IDs. This remaps them by author+year+title matching.</div>',unsafe_allow_html=True)
        col1,col2=st.columns(2)
        with col1: remap_doc=st.file_uploader("Word document",type=["docx"],key="remap_doc")
        with col2: remap_enl=st.file_uploader("KAROL .enl library",type=["enl"],key="remap_enl")
        if remap_doc and remap_enl:
            if st.button("Find and remap",type="primary",key="remap_run"):
                with st.spinner("Comparing against KAROL..."):
                    try:
                        fixed_bytes,report=remap_traveling_citations(remap_doc.read(),remap_enl.read())
                        remapped=[r for r in report if r['status']=='remapped']
                        not_found=[r for r in report if r['status']=='not_found']
                        col1,col2,col3=st.columns(3)
                        col1.metric("Checked",len(report)); col2.metric("Remapped",len(remapped)); col3.metric("Not found",len(not_found))
                        if remapped:
                            st.success(f"Remapped {len(remapped)} citations.")
                            st.download_button("⬇ Download remapped document",data=fixed_bytes,
                                file_name=Path(remap_doc.name).stem+"_remapped.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",type="primary")
                        if not_found:
                            st.warning(f"{len(not_found)} citations not found in KAROL — add them manually.")
                            with st.expander("Unmatched citations"):
                                for r in not_found:
                                    st.markdown(f'<div class="ref-item missing"><b>RecNum {r["old_rec_num"]}</b> — {r["author"]} ({r["year"]}) {r["title"]}</div>',unsafe_allow_html=True)
                    except Exception as e: st.error(f"Error: {e}")

    st.divider()

    # Stage 1
    stage1_done=st.session_state.fix_analysis is not None
    st.markdown(f'<div class="step-card {"done" if stage1_done else "active"}"><div class="step-header"><span class="step-num {"done" if stage1_done else ""}">{"✓" if stage1_done else "STEP 1"}</span><span class="step-title">Repair broken citation field codes</span></div><div class="step-desc">Recovers citation XML data from the document\'s internal backup storage.</div></div>',unsafe_allow_html=True)
    if not stage1_done:
        doc_file=st.file_uploader("Word document (.docx)",type=["docx"],key="fix_doc_upload")
        if doc_file and st.button("Analyze & repair",type="primary"):
            with st.spinner("Scanning..."):
                docx_bytes=doc_file.read()
                analysis=analyze_docx_citations(docx_bytes)
                fixed_xml,n_fixed=fix_broken_fields(analysis['raw'])
                st.session_state.fix_analysis=analysis; st.session_state.fix_docx_bytes=docx_bytes
                st.session_state.fix_after_stage1=fixed_xml; st.session_state.fix_raw_xml=fixed_xml
                st.session_state.fix_doc_name=doc_file.name
            st.rerun()
    else:
        a=st.session_state.fix_analysis
        col1,col2,col3,col4=st.columns(4)
        col1.metric("Total fields",a['total_fields']); col2.metric("Working",a['working'])
        col3.metric("Were broken",a['broken_empty']); col4.metric("Recovered",a['broken_empty'])
        if a['broken_empty']>0: st.success(f"✓ {a['broken_empty']} broken field(s) recovered.")
        else: st.info("No broken fields found.")
        stage1_bytes=build_fixed_docx(st.session_state.fix_docx_bytes,st.session_state.fix_after_stage1)
        st.download_button("⬇ Download Stage 1 result",data=stage1_bytes,
            file_name=Path(st.session_state.fix_doc_name).stem+"_stage1.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        if st.button("↺ Start over"):
            for k in ['fix_analysis','fix_raw_xml','fix_docx_bytes','fix_after_stage1','fix_after_stage2','fix_karol_db_id','fix_karol_rec_nums','fix_missing_refs']:
                st.session_state[k]=defaults[k]
            st.rerun()

    st.divider()

    # Stage 2
    stage2_done=st.session_state.fix_karol_db_id is not None
    st.markdown(f'<div class="step-card {"done" if stage2_done else ("active" if stage1_done else "waiting")}"><div class="step-header"><span class="step-num {"done" if stage2_done else ""}">{"✓" if stage2_done else "STEP 2"}</span><span class="step-title">Get KAROL library fingerprint (db-id)</span></div><div class="step-desc">Export any refs from KAROL as XML (File → Export → XML), then upload below. The app extracts the library fingerprint automatically.</div></div>',unsafe_allow_html=True)
    if stage1_done and not stage2_done:
        col_a,col_b=st.columns(2)
        with col_a:
            xml_exp=st.file_uploader("EndNote XML export (gets db-id)",type=["xml"],key="fix_xml_export")
        with col_b:
            enl_f=st.file_uploader("KAROL .enl (checks missing refs, optional)",type=["enl"],key="fix_enl")
        if xml_exp:
            db_id=extract_karol_db_id(xml_exp.read())
            if db_id:
                st.session_state.fix_karol_db_id=db_id
                st.success(f"✓ KAROL fingerprint found: `{db_id[:20]}...`"); st.rerun()
            else: st.error("No db-id found. Make sure you exported from EndNote as XML format.")
        if enl_f:
            with st.spinner("Reading KAROL..."):
                rns=get_karol_rec_nums(enl_f.read())
                st.session_state.fix_karol_rec_nums=rns
                missing=check_missing_from_karol(st.session_state.fix_raw_xml or "",set(rns.keys()))
                st.session_state.fix_missing_refs=missing
            st.success(f"✓ {len(rns):,} refs in KAROL.")
            if missing: st.warning(f"⚠ {len(missing)} ref(s) not in KAROL: RecNums {', '.join(missing)}")
        with st.expander("Enter db-id manually"):
            mid=st.text_input("db-id",placeholder="e.g. s5pa559ekdxfr0esvw85...")
            if st.button("Use this db-id") and mid.strip():
                st.session_state.fix_karol_db_id=mid.strip(); st.rerun()
    elif stage2_done:
        st.success(f"✓ KAROL db-id: `{st.session_state.fix_karol_db_id[:20]}...`")

    st.divider()

    # Stage 3
    stage3_done=st.session_state.fix_after_stage2 is not None
    stage3_ready=stage1_done and stage2_done
    st.markdown(f'<div class="step-card {"done" if stage3_done else ("active" if stage3_ready else "waiting")}"><div class="step-header"><span class="step-num {"done" if stage3_done else ""}">{"✓" if stage3_done else "STEP 3"}</span><span class="step-title">Apply full fix and generate files</span></div></div>',unsafe_allow_html=True)
    if stage3_ready and not stage3_done:
        missing=st.session_state.fix_missing_refs
        proceed=True
        if missing:
            st.warning(f"⚠ {len(missing)} ref(s) not in KAROL — add them manually first (EndNote → References → New Reference).")
            for rn in missing: st.markdown(f"- RecNum **{rn}**")
            proceed=st.checkbox("I've added the missing refs (or will add them later)")
        if proceed and st.button("Apply full fix",type="primary"):
            with st.spinner("Patching db-ids..."):
                patched,_=patch_db_ids(st.session_state.fix_raw_xml,
                                       st.session_state.fix_analysis['db_ids'],
                                       st.session_state.fix_karol_db_id)
                st.session_state.fix_after_stage2=patched; st.rerun()
    elif stage3_done:
        patched=st.session_state.fix_after_stage2
        final_bytes=build_fixed_docx(st.session_state.fix_docx_bytes,patched)
        working_after=len(re.findall(r'&lt;EndNote&gt;',patched))
        st.success("✓ Full fix applied.")
        col1,col2=st.columns(2)
        col1.metric("Citations now linked",working_after)
        col_d1,col_d2=st.columns(2)
        with col_d1:
            st.download_button("⬇ Download fixed document",data=final_bytes,
                file_name=Path(st.session_state.fix_doc_name).stem+"_fully_fixed.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",type="primary")
        with col_d2:
            macro=generate_vba_macro(st.session_state.fix_doc_name)
            st.download_button("⬇ Download VBA macro",data=macro.encode(),
                file_name="RelinkEndNoteCitations.bas",mime="text/plain")
        st.markdown('<div class="instruction-box"><b>Final step in Word:</b><br>1. Open the fixed document with KAROL connected<br>2. Open the .bas file in Notepad, copy all<br>3. In Word: Alt+F11 → Insert → Module → paste → Alt+F8 → RelinkAllCitations → Run<br>4. EndNote tab → Update Citations and Bibliography</div>',unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# APP 1 UI — CITATION REPAIR
# ─────────────────────────────────────────────────────────────────────────────
elif tool == "App 1 — Citation Repair":
    st.markdown("## Citation Repair")
    st.markdown('<div class="step-desc">Find missing citation placeholders and match them to your EndNote library.</div>',unsafe_allow_html=True)
    col1,col2=st.columns(2)
    with col1: doc_file=st.file_uploader("Word document (.docx)",type=["docx"],key="doc_upload")
    with col2: lib_file=st.file_uploader("EndNote library (.xml)",type=["xml"],key="lib_upload")
    cfg1,cfg2=st.columns(2)
    with cfg1: mode=st.selectbox("Mode",["Interactive Review","Auto-Insert","Report Only"])
    with cfg2: use_pubmed=st.toggle("PubMed fallback",value=False)
    with st.expander("Custom citation markers"):
        custom=st.text_input("Additional markers (comma-separated)",placeholder="e.g. ??, [TBD]")
        if custom:
            extras=[re.escape(m.strip()) for m in custom.split(",") if m.strip()]
            CITATION_MARKERS=re.compile('|'.join(MISSING_PATTERNS+extras),re.IGNORECASE)
    run_btn=st.button("Scan Document",type="primary",disabled=not(doc_file and lib_file))
    if run_btn and doc_file and lib_file:
        for k in ["flagged","current_idx","decisions","doc_obj","refs","vec","mat","repair_done"]:
            st.session_state[k]=defaults[k]
        with st.spinner("Parsing..."):
            db=doc_file.read(); flagged,doc_obj=extract_flagged(db)
            refs=parse_endnote_xml_bytes(lib_file.read())
            if refs:
                corpora=tuple(r["corpus"] for r in refs); vec,mat=build_tfidf(corpora)
                st.session_state.update(dict(flagged=flagged,doc_obj=doc_obj,refs=refs,
                    vec=vec,mat=mat,current_idx=0,decisions=[],repair_done=False))
        if not refs: st.error("No references found in XML.")
        elif not flagged: st.warning("No citation markers found.")
        else: st.success(f"Found **{len(flagged)}** missing citation(s) across **{len(refs)}** refs.")
    flagged=st.session_state.flagged
    if flagged and not st.session_state.repair_done:
        st.divider()
        idx=st.session_state.current_idx; total=len(flagged); done=len(st.session_state.decisions)
        st.markdown(f'<div class="progress-outer"><div class="progress-inner" style="width:{int(done/total*100)}%"></div></div>',unsafe_allow_html=True)
        st.caption(f"{done} of {total} reviewed")
        if idx<total:
            item=flagged[idx]
            cands=match_sentence(item["sentence"],st.session_state.vec,st.session_state.mat,st.session_state.refs)
            best=cands[0]["score"] if cands else 0
            if mode=="Auto-Insert" and best>=TFIDF_THRESHOLD:
                label=author_label(cands[0]["ref"]); para=st.session_state.doc_obj.paragraphs[item["para_idx"]]
                insert_superscript(para,item["marker"],label)
                st.session_state.decisions.append({**item,"action":"accepted","ref":cands[0]["ref"],"score":best,"candidates":cands})
                st.session_state.current_idx+=1; st.rerun()
            elif mode=="Report Only":
                st.session_state.decisions.append({**item,"action":"skipped","candidates":cands})
                st.session_state.current_idx+=1; st.rerun()
            else:
                st.markdown(f'<div class="match-card"><span class="match-marker">{item["marker"]}</span><div class="match-sentence">"{item["sentence"][:280]}"</div><div style="font-size:0.72rem;color:#3a4a5a">Para {item["para_idx"]+1}</div></div>',unsafe_allow_html=True)
                st.markdown('<div class="section-label">Top Matches</div>',unsafe_allow_html=True)
                chosen_idx=None
                for j,c in enumerate(cands):
                    ca,cb=st.columns([1,8])
                    with ca:
                        if st.button("Use",key=f"pick_{idx}_{j}"): chosen_idx=j
                    with cb:
                        sc=score_class(c["score"])
                        st.markdown(f'<div style="display:flex;align-items:center;gap:8px;padding:5px 0"><span class="score-pill {sc}">{c["score"]:.3f}</span><span style="font-size:0.83rem;color:#c8d0db">{fmt_ref(c["ref"])}</span></div>',unsafe_allow_html=True)
                    if chosen_idx==j:
                        label=author_label(cands[chosen_idx]["ref"])
                        para=st.session_state.doc_obj.paragraphs[item["para_idx"]]
                        insert_superscript(para,item["marker"],label)
                        st.session_state.decisions.append({**item,"action":"accepted","ref":cands[chosen_idx]["ref"],"score":cands[chosen_idx]["score"],"candidates":cands})
                        st.session_state.current_idx+=1; st.rerun()
                ba,bb=st.columns(2)
                with ba:
                    if st.button("Skip",key=f"skip_{idx}"):
                        st.session_state.decisions.append({**item,"action":"skipped","candidates":cands})
                        st.session_state.current_idx+=1; st.rerun()
                with bb:
                    if done>0 and st.button("Finish",key=f"fin_{idx}"):
                        st.session_state.repair_done=True; st.rerun()
        else:
            st.session_state.repair_done=True; st.rerun()
    if st.session_state.repair_done and st.session_state.decisions:
        st.divider()
        decisions=st.session_state.decisions
        accepted=sum(1 for d in decisions if d["action"]=="accepted")
        skipped=sum(1 for d in decisions if d["action"]=="skipped")
        col1,col2,col3=st.columns(3)
        col1.metric("Total",len(decisions)); col2.metric("Accepted",accepted); col3.metric("Skipped",skipped)
        cd1,cd2=st.columns(2)
        with cd1:
            st.download_button("⬇ Repaired document",data=doc_to_bytes(st.session_state.doc_obj),
                file_name="manuscript_repaired.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",type="primary")
        with cd2:
            st.download_button("⬇ Decision report",data=write_repair_report(decisions),
                file_name="citation_report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# ─────────────────────────────────────────────────────────────────────────────
# APP 3 UI — REFERENCE COMPARATOR
# ─────────────────────────────────────────────────────────────────────────────
elif tool == "App 3 — Reference Comparator":
    st.markdown("## Reference List Comparator")
    col1,col2=st.columns(2)
    with col1:
        st.markdown("**List A**")
        file_a=st.file_uploader("List A",type=["xml","docx","txt"],key="comp_a",label_visibility="collapsed")
    with col2:
        st.markdown("**List B**")
        file_b=st.file_uploader("List B",type=["xml","docx","txt"],key="comp_b",label_visibility="collapsed")
    st.markdown("**Manuscript (optional)**")
    ms_file=st.file_uploader("Manuscript",type=["docx"],key="comp_ms",label_visibility="collapsed")
    if st.button("Compare Lists",type="primary",disabled=not(file_a and file_b)):
        with st.spinner("Comparing..."):
            refs_a,label_a=load_ref_file(file_a); refs_b,label_b=load_ref_file(file_b)
            if refs_a and refs_b:
                all_c=[r["corpus"] for r in refs_a]+[r["corpus"] for r in refs_b]
                vec=TfidfVectorizer(ngram_range=(1,2),sublinear_tf=True,max_features=50000)
                vec.fit(all_c)
                emb_a=vec.transform([r["corpus"] for r in refs_a])
                emb_b=vec.transform([r["corpus"] for r in refs_b])
                matrix=cosine_similarity(emb_a,emb_b)
                matched,only_a,fuzzy,matched_b=[],[],[],set()
                for i,ra in enumerate(refs_a):
                    bj=int(matrix[i].argmax()); bs=float(matrix[i][bj])
                    if bs>=MATCH_THRESHOLD: matched.append((ra,refs_b[bj],bs)); matched_b.add(bj)
                    elif bs>=FUZZY_THRESHOLD: fuzzy.append((ra,refs_b[bj],bs))
                    else: only_a.append(ra)
                only_b=[refs_b[j] for j in range(len(refs_b)) if j not in matched_b]
                result=dict(matched=matched,only_in_a=only_a,only_in_b=only_b,fuzzy=fuzzy)
                usage={}
                if ms_file:
                    ms_doc=Document(io.BytesIO(ms_file.read()))
                    ms_paras=[p.text for p in ms_doc.paragraphs if p.text.strip()]
                    for ref in only_a+only_b+[x[0] for x in fuzzy]:
                        words=[w for w in ref["title"].split() if len(w)>5][:5]
                        found=[p[:200] for p in ms_paras if sum(1 for w in words if w.lower() in p.lower())>=min(3,len(words))]
                        usage[ref.get("id","")]=found
                st.session_state.update(dict(comp_result=result,comp_usage=usage,
                    comp_labels=(label_a,label_b),comp_refs=(refs_a,refs_b)))
    if st.session_state.comp_result:
        result=st.session_state.comp_result; usage=st.session_state.comp_usage
        label_a,label_b=st.session_state.comp_labels
        st.divider()
        col1,col2,col3,col4=st.columns(4)
        col1.metric("Matched",len(result["matched"])); col2.metric("Only in A",len(result["only_in_a"]))
        col3.metric("Only in B",len(result["only_in_b"])); col4.metric("Review",len(result["fuzzy"]))
        tabs=st.tabs([f"Only in A ({len(result['only_in_a'])})",f"Only in B ({len(result['only_in_b'])})",
                      f"Review ({len(result['fuzzy'])})",f"Matched ({len(result['matched'])})"])
        def ref_card(ref,color,locs=None):
            loc=""
            if locs: loc=f'<div style="font-size:0.75rem;color:#4fc3f7;margin-top:4px">Found in {len(locs)} paragraph(s)</div>'+"".join(f'<div style="font-size:0.74rem;color:#607080;font-style:italic">{l[:150]}...</div>' for l in locs[:2])
            st.markdown(f'<div class="ref-item {color}">{fmt_ref(ref)}{loc}</div>',unsafe_allow_html=True)
        with tabs[0]:
            if result["only_in_a"]:
                st.caption(f"In **{label_a}** but not **{label_b}**")
                for ref in result["only_in_a"]: ref_card(ref,"missing",usage.get(ref.get("id",""),[]))
            else: st.success("None")
        with tabs[1]:
            if result["only_in_b"]:
                st.caption(f"In **{label_b}** but not **{label_a}**")
                for ref in result["only_in_b"]: ref_card(ref,"warn",usage.get(ref.get("id",""),[]))
            else: st.success("None")
        with tabs[2]:
            for ra,rb,score in result["fuzzy"]:
                st.markdown(f'<div class="ref-item warn"><b>[{score:.3f}]</b> A:{fmt_ref(ra,True)}<br>B:{fmt_ref(rb,True)}</div>',unsafe_allow_html=True)
        with tabs[3]:
            for ra,rb,score in result["matched"]:
                st.markdown(f'<div class="ref-item ok">[{score:.3f}] {fmt_ref(ra,True)}</div>',unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# APP 4 UI — DOCUMENT MERGER
# ─────────────────────────────────────────────────────────────────────────────
elif tool == "App 4 — Document Merger":
    st.markdown("## Document Merger")
    st.markdown('<div class="instruction-box"><b>When to use:</b> You merged two Word documents and EndNote no longer recognizes the citations. This tool accepts tracked changes safely (rescuing any citations inside deleted text) and repairs broken citation field codes.<br><br><b>Best practice for future merges:</b> Before using Word\'s Compare, go to EndNote tab → Convert Citations → Convert to Unformatted Citations. This turns field codes into plain text like {Hall, 1997 #18} which survives merging perfectly. Then after accepting changes, use Update Citations and Bibliography to reformat.</div>',unsafe_allow_html=True)
    col1,col2=st.columns(2)
    with col1:
        st.markdown("**Post-merge document** — the merged file with broken citations")
        merged_file=st.file_uploader("Merged .docx",type=["docx"],key="merge_merged")
    with col2:
        st.markdown("**Original document** (optional) — used to detect lost citations")
        orig_file=st.file_uploader("Original .docx",type=["docx"],key="merge_orig")
    if merged_file:
        if st.button("Analyze & repair",type="primary"):
            with st.spinner("Analyzing citation damage..."):
                merged_bytes=merged_file.read()
                orig_bytes=orig_file.read() if orig_file else None
                analysis=analyze_merge_damage(merged_bytes)
            st.markdown("### Damage report")
            col1,col2,col3,col4=st.columns(4)
            col1.metric("Total citation fields",analysis['total_en'])
            col2.metric("Working",analysis['with_data'])
            col3.metric("Broken fields",analysis['empty_cite'])
            col4.metric("Tracked changes",analysis['ins_count']+analysis['del_count'])
            if not analysis['balanced']:
                st.warning(f"⚠ Unbalanced field markers (begin:{analysis['begins']}/sep:{analysis['separates']}/end:{analysis['ends']}) — citation fields were split during merge.")
            else: st.success("✓ Citation field boundaries intact.")
            if len(analysis['db_ids'])>1:
                st.info(f"Multiple library fingerprints found ({len(analysis['db_ids'])}) — document contains citations from different libraries.")
            with st.spinner("Repairing..."):
                fixed_bytes,report=repair_post_merge_citations(merged_bytes,orig_bytes)
            st.markdown("### Results")
            col1,col2,col3=st.columns(3)
            col1.metric("Citations before",report['citations_before'])
            col2.metric("Citations after",report['citations_after'])
            col3.metric("Steps applied",len(report['steps']))
            if report['steps']:
                for step in report['steps']:
                    if step=='track_changes_accepted': st.markdown("- ✓ Tracked changes accepted safely")
                    elif 'restored' in step: n=step.split('_')[1]; st.markdown(f"- ✓ {n} broken field(s) restored")
                    elif 'lost' in step: n=step.split('_')[0]; st.markdown(f"- ⚠ {n} citation(s) lost during merge")
            if report.get('lost_rec_nums'):
                st.warning(f"⚠ {len(report['lost_rec_nums'])} citation(s) from original not found after merge: RecNums {', '.join(report['lost_rec_nums'])}")
            st.download_button("⬇ Download repaired document",data=fixed_bytes,
                file_name=Path(merged_file.name).stem+"_repaired.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",type="primary")
            st.markdown('<div class="instruction-box"><b>After downloading:</b><br>1. Open in Word with KAROL connected<br>2. EndNote tab → Update Citations and Bibliography<br>3. Still issues? → Use App 2 Remap tool</div>',unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# APP 5 UI — CITATION RENUMBERING
# ─────────────────────────────────────────────────────────────────────────────
elif tool == "App 5 — Citation Renumbering":
    st.markdown("## Citation Renumbering")
    st.markdown('<div class="instruction-box"><b>When to use:</b> After editing, inline citation superscript numbers and the bibliography are out of sequential order (e.g. text jumps 1, 5, 2, 8...). This renumbers everything sequentially by order of first appearance.<br><br><b>Before using:</b> Make sure EndNote has already formatted the bibliography so citations appear as plain superscript numbers — not live field codes.</div>',unsafe_allow_html=True)
    ren_file=st.file_uploader("Word document (.docx) with formatted citations",type=["docx"],key="ren_doc")
    if ren_file:
        if st.button("Renumber citations",type="primary"):
            with st.spinner("Scanning and renumbering..."):
                fixed_bytes,mapping=renumber_citations(ren_file.read())
            if not mapping:
                st.warning("No superscript citation numbers found. Make sure you have used EndNote to format the bibliography first.")
            else:
                changed={k:v for k,v in mapping.items() if k!=v}
                st.success(f"Done. {len(mapping)} unique citations. {len(changed)} numbers changed.")
                col1,col2,col3=st.columns(3)
                col1.metric("Unique citations",len(mapping))
                col2.metric("Numbers changed",len(changed))
                col3.metric("Already in order",len(mapping)-len(changed))
                st.download_button("⬇ Download renumbered document",data=fixed_bytes,
                    file_name=Path(ren_file.name).stem+"_renumbered.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",type="primary")
                if changed:
                    with st.expander(f"Renumbering map ({len(changed)} changes)"):
                        col1,col2=st.columns(2)
                        col1.markdown("**Old #**"); col2.markdown("**New #**")
                        for old in sorted(mapping):
                            new=mapping[old]
                            if old!=new: col1.markdown(str(old)); col2.markdown(str(new))

# ─────────────────────────────────────────────────────────────────────────────
# APP 6 UI — FIGURE INVENTORY
# ─────────────────────────────────────────────────────────────────────────────
elif tool == "App 6 — Figure Inventory":
    st.markdown("## Figure & Table Inventory")
    st.markdown('<div class="instruction-box">Scans your document for all figures, tables, boxes, plates, and videos. Optionally cross-references against your Excel naming sheet to flag mismatches and missing items.</div>',unsafe_allow_html=True)
    col1,col2=st.columns(2)
    with col1: fig_doc=st.file_uploader("Word document (.docx)",type=["docx"],key="fig_doc")
    with col2: fig_excel=st.file_uploader("Excel naming sheet (optional)",type=["xlsx","xls"],key="fig_excel",help="Columns: Number | Name/Caption")
    if fig_doc:
        if st.button("Scan document",type="primary"):
            with st.spinner("Scanning..."):
                fig_bytes=fig_doc.read(); items=scan_figures(fig_bytes)
            if not items:
                st.warning("No captioned items found. Captions must start with Figure/Table/Box/Plate/Video followed by a number.")
            else:
                type_counts={}
                for item in items: type_counts[item['type']]=type_counts.get(item['type'],0)+1
                cols=st.columns(min(len(type_counts),6))
                for i,(t,c) in enumerate(sorted(type_counts.items())): cols[i%len(cols)].metric(t+"s",c)
                st.markdown(f"**{len(items)} total items**")
                if fig_excel:
                    excel_bytes=fig_excel.read(); results=cross_ref_excel(items,excel_bytes)
                    matched=[r for r in results if r[1]=='match']
                    mismatch=[r for r in results if r[1]=='mismatch']
                    not_found=[r for r in results if r[1]=='not_in_excel']
                    col1,col2,col3,col4=st.columns(4)
                    col1.metric("Matched",len(matched)); col2.metric("Mismatch",len(mismatch))
                    col3.metric("Not in Excel",len(not_found))
                    tabs=st.tabs([f"Mismatches ({len(mismatch)})",f"Not in Excel ({len(not_found)})",
                                  f"Matched ({len(matched)})",f"All ({len(results)})"])
                    def show_res(res_list):
                        for item,status,expected in res_list:
                            color='ok' if status=='match' else 'missing' if status=='not_in_excel' else 'warn'
                            exp=""
                            if expected and status=='mismatch': exp=f'<br><span style="font-size:0.75rem;color:#ffb300">Expected: {expected}</span>'
                            st.markdown(f'<div class="ref-item {color}"><b>{item["type"]} {item["number"]}</b>{"🖼" if item["has_image"] else ""} — Para {item["para_idx"]+1}<br><span style="font-size:0.82rem">{item["caption"][:180]}</span>{exp}</div>',unsafe_allow_html=True)
                    with tabs[0]: show_res(mismatch) if mismatch else st.success("No mismatches.")
                    with tabs[1]: show_res(not_found) if not_found else st.success("All found in Excel.")
                    with tabs[2]: show_res(matched) if matched else st.info("No confirmed matches.")
                    with tabs[3]: show_res(results)
                else:
                    type_filter=st.selectbox("Filter by type",["All"]+sorted(type_counts.keys()))
                    for item in items:
                        if type_filter!="All" and item['type']!=type_filter: continue
                        st.markdown(f'<div class="match-card"><span class="match-marker">{item["type"]} {item["number"]}</span>{"🖼" if item["has_image"] else ""}<span style="font-size:0.72rem;color:#3a4a5a"> Para {item["para_idx"]+1}</span><div class="match-sentence">{item["caption"][:200]}</div></div>',unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# APP 7 UI — PUBMED SEARCH
# ─────────────────────────────────────────────────────────────────────────────
elif tool == "App 7 — PubMed Search":
    st.markdown("## PubMed Literature Search")
    st.markdown('<div class="instruction-box">Search PubMed for relevant articles. Results include abstracts and links to free full text where available (PubMed Central). Export results as EndNote XML to import directly into KAROL.</div>',unsafe_allow_html=True)
    query=st.text_input("Search query",placeholder="e.g. arthrogryposis clubfoot Ponseti treatment",help="Supports PubMed syntax: AND, OR, NOT, [MeSH], [ti], [au]")
    col1,col2,col3=st.columns(3)
    with col1: date_from=st.text_input("Year from",placeholder="2019")
    with col2: date_to=st.text_input("Year to",placeholder="2026")
    with col3: max_res=st.slider("Max results",5,50,20)
    journal_filter=st.text_input("Limit to journal (optional)",placeholder="e.g. J Pediatr Orthop")
    if st.button("Search PubMed",type="primary",disabled=not query.strip()):
        with st.spinner(f"Searching PubMed: {query}..."):
            results=pubmed_search_full(query.strip(),date_from.strip(),date_to.strip(),journal_filter.strip(),max_res)
        if not results: st.warning("No results found. Try broadening your search terms.")
        else:
            st.success(f"Found {len(results)} results.")
            xml_export=results_to_xml(results)
            st.download_button(f"⬇ Export all {len(results)} refs as EndNote XML",
                data=xml_export.encode('utf-8'),file_name="pubmed_results.xml",mime="application/xml")
            st.divider()
            for i,r in enumerate(results,1):
                authors_str='; '.join(r['authors'][:3])+(' et al.' if len(r['authors'])>3 else '')
                cit=f"{authors_str} ({r['year']}). *{r['journal']}*"
                if r['volume']: cit+=f" {r['volume']}"
                if r['issue']:  cit+=f"({r['issue']})"
                if r['pages']:  cit+=f":{r['pages']}"
                with st.expander(f"**{i}.** {r['title'][:100]}{'...' if len(r['title'])>100 else ''}"):
                    st.markdown(cit)
                    lc=st.columns(3)
                    with lc[0]: st.markdown(f"[PubMed]({r['pubmed_url']})")
                    with lc[1]:
                        if r['pmc_url']: st.markdown(f"[Free full text (PMC)]({r['pmc_url']})")
                        elif r['doi_url']: st.markdown(f"[DOI]({r['doi_url']})")
                        else: st.markdown("*No free full text*")
                    with lc[2]:
                        if r['doi']: st.caption(f"DOI: {r['doi']}")
                    if r['abstract']: st.markdown("**Abstract:**"); st.markdown(r['abstract'])
                    single=results_to_xml([r])
                    st.download_button("⬇ Export this ref",data=single.encode('utf-8'),
                        file_name=f"ref_{r['pmid']}.xml",mime="application/xml",key=f"exp_{r['pmid']}")

# ─────────────────────────────────────────────────────────────────────────────
# APP 8 UI — BATCH RENAME
# ─────────────────────────────────────────────────────────────────────────────
elif tool == "App 8 — Batch Rename":
    st.markdown("## Batch Rename")
    st.markdown('<div class="instruction-box">Apply bulk find-and-replace across your Word document using an Excel naming sheet. Use for final editing to update chapter names, author names, figure labels, section titles, etc.<br><br><b>Excel format:</b> Two columns — <code>Old Name</code> (find) and <code>New Name</code> (replace). One row per replacement.</div>',unsafe_allow_html=True)
    col1,col2=st.columns(2)
    with col1: ren_doc=st.file_uploader("Word document (.docx)",type=["docx"],key="batch_doc")
    with col2: ren_excel=st.file_uploader("Excel naming sheet (.xlsx)",type=["xlsx","xls"],key="batch_excel")
    col3,col4=st.columns(2)
    with col3: match_case=st.toggle("Match case",value=False)
    with col4: whole_word=st.toggle("Whole word only",value=False)
    pairs=[]
    if ren_excel:
        with st.spinner("Reading naming sheet..."):
            pairs=load_rename_pairs(ren_excel.read())
        if pairs:
            st.success(f"✓ {len(pairs)} rename pairs loaded.")
            with st.expander("Preview"):
                col1,col2=st.columns(2)
                col1.markdown("**Find**"); col2.markdown("**Replace**")
                for old,new in pairs[:20]: col1.markdown(old); col2.markdown(new)
                if len(pairs)>20: st.caption(f"...and {len(pairs)-20} more")
        else: st.error("Could not read pairs. Check Excel has 'Old Name' and 'New Name' columns.")
    if ren_doc and ren_excel and pairs:
        if st.button("Apply batch rename",type="primary"):
            with st.spinner("Applying..."):
                fixed_bytes,report=batch_rename(ren_doc.read(),pairs,match_case,whole_word)
            replaced=[r for r in report if r['status']=='replaced']
            not_found=[r for r in report if r['status']=='not_found']
            total_changes=sum(r['count'] for r in replaced)
            col1,col2,col3=st.columns(3)
            col1.metric("Total pairs",len(report)); col2.metric("Replaced",len(replaced)); col3.metric("Not found",len(not_found))
            st.success(f"✓ {total_changes} replacement(s) across {len(replaced)} pairs.")
            st.download_button("⬇ Download renamed document",data=fixed_bytes,
                file_name=Path(ren_doc.name).stem+"_renamed.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",type="primary")
            if replaced:
                with st.expander(f"Replacements made ({len(replaced)})"):
                    for r in replaced:
                        st.markdown(f'<div class="ref-item ok"><b>{r["old"]}</b> → {r["new"]} <span style="float:right;font-size:0.75rem">{r["count"]}x</span></div>',unsafe_allow_html=True)
            if not_found:
                with st.expander(f"Not found ({len(not_found)})"):
                    for r in not_found:
                        st.markdown(f'<div class="ref-item missing"><b>{r["old"]}</b> — not found in document</div>',unsafe_allow_html=True)
