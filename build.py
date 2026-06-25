#!/usr/bin/env python3
"""build.py – Flexible parser + adaptive data injector for pedagogy-hub.html"""

import json, re, base64, os
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

SCRIPT_DIR = Path(os.path.abspath(__file__)).parent if '__file__' in dir() else Path('/home/claude')

UNICODE_MAP = {
    '\u2018':"'",'\u2019':"'",'\u201c':'"','\u201d':'"',
    '\u2013':'-','\u2014':'--','\u2026':'...','\u00a0':' ',
    '\uf0b7':'','\uf0a7':'','\uf020':'',
}
def clean(text):
    if not text: return ''
    for src,dst in UNICODE_MAP.items(): text=text.replace(src,dst)
    text=re.sub(r'[\ue000-\uf8ff]','',text)
    text=re.sub(r'\s+',' ',text)
    return text.strip()

def para_text(p):
    return clean(''.join(run.text or '' for run in p.runs))

def is_tech(p):
    return bool(re.match(r'^\(Tech\)',para_text(p),re.IGNORECASE))

def strip_tech(t):
    return re.sub(r'^\(Tech\)\s*','',t,flags=re.IGNORECASE).strip()

def para_style(p):
    return (p.style.name or '').strip()

def has_image(p):
    return bool(p._element.findall('.//' + qn('a:blip')))

def extract_image_b64(p, docx_path):
    blips = p._element.findall('.//' + qn('a:blip'))
    if not blips: return None
    rId = blips[0].get(qn('r:embed'))
    if not rId: return None
    doc2 = Document(docx_path)
    for rel in doc2.part.rels.values():
        if rel.reltype.endswith('/image') and rel.rId == rId:
            image_bytes = rel.target_part.blob
            b64 = base64.b64encode(image_bytes).decode('utf-8')
            mime = 'image/png' if image_bytes[:4]==b'\x89PNG' else 'image/jpeg'
            return f'data:{mime};base64,{b64}'
    return None

def build_numbering_lookup(doc):
    numbering_part = None
    for rel in doc.part.rels.values():
        if 'numbering' in rel.reltype:
            numbering_part = rel.target_part
            break
    if not numbering_part: return {}, {}
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    root = etree.fromstring(numbering_part.blob)
    abstract = {}
    for abNum in root.findall(f'{{{ns}}}abstractNum'):
        abId = abNum.get(f'{{{ns}}}abstractNumId')
        abstract[abId] = {}
        for lvl in abNum.findall(f'{{{ns}}}lvl'):
            ilvl = lvl.get(f'{{{ns}}}ilvl')
            lvlText = lvl.find(f'{{{ns}}}lvlText')
            char = lvlText.get(f'{{{ns}}}val','') if lvlText is not None else ''
            abstract[abId][ilvl] = char
    num_to_abstract = {}
    for num in root.findall(f'{{{ns}}}num'):
        numId = num.get(f'{{{ns}}}numId')
        abRef = num.find(f'{{{ns}}}abstractNumId')
        if abRef is not None:
            num_to_abstract[numId] = abRef.get(f'{{{ns}}}val')
    return abstract, num_to_abstract

def get_bullet_level(p, abstract, num_to_abstract):
    numPr = p._element.find('.//' + qn('w:numPr'))
    ilvl_val = 0; numId_val = None
    if numPr is not None:
        ilvl_el = numPr.find(qn('w:ilvl'))
        numId_el = numPr.find(qn('w:numId'))
        if ilvl_el is not None: ilvl_val = int(ilvl_el.get(qn('w:val'),'0'))
        if numId_el is not None: numId_val = numId_el.get(qn('w:val'))
    if ilvl_val >= 1: return 1
    if numId_val and numId_val in num_to_abstract:
        abId = num_to_abstract[numId_val]
        char = abstract.get(abId,{}).get(str(ilvl_val),'')
        if char == 'o': return 1
    pPr = p._element.find(qn('w:pPr'))
    if pPr is not None:
        ind = pPr.find(qn('w:ind'))
        if ind is not None:
            ind_left = ind.get(qn('w:left'))
            if ind_left and int(ind_left) >= 400: return 1
    return 0

def cell_bullets(cell, abstract, num_to_abstract):
    bullets = []
    for p in cell.paragraphs:
        t = para_text(p)
        if not t: continue
        tech = is_tech(p)
        level = get_bullet_level(p, abstract, num_to_abstract)
        bullets.append({'text': strip_tech(t) if tech else t, 'tech': tech, 'level': level})
    return bullets

def cell_plain(cell):
    return clean(cell.text)

# ── aims cell parser ───────────────────────────────────────────────────────
TAG_RE = re.compile(r'^\[([^\]]+)\]')

def parse_aims_cell(cell):
    """
    Parse the aims cell into a list of aim blocks:
    [{'tag': 'Understand', 'text': 'Students deepen...'}, ...]
    Handles both single-para (tag + text on same line) and
    multi-para (tag on own para, text on following paras) formats.
    """
    paras = [clean(p.text) for p in cell.paragraphs]
    paras = [p for p in paras if p]  # drop empty

    blocks = []
    current_tag = None
    current_lines = []

    def flush():
        if current_tag is not None:
            blocks.append({'tag': current_tag, 'text': ' '.join(current_lines).strip()})

    for para in paras:
        m = TAG_RE.match(para)
        if m:
            flush()
            current_tag = m.group(1).strip()
            rest = para[m.end():].strip()
            current_lines = [rest] if rest else []
        else:
            if current_tag is not None:
                current_lines.append(para)
            else:
                # text before any tag — unlikely but safe
                blocks.append({'tag': None, 'text': para})

    flush()
    return blocks if blocks else None

# ── phase table ────────────────────────────────────────────────────────────
def parse_phase_table(tbl, abstract, num_to_abstract):
    rows = tbl.rows
    if not rows: return None
    col_headers = [cell_plain(c) for c in rows[0].cells]
    n_cols = len(col_headers)
    teacher_col = next((i for i,h in enumerate(col_headers) if 'teacher' in h.lower()), 1)
    student_col = next((i for i,h in enumerate(col_headers) if 'student' in h.lower()), 2)
    aims_col    = next((i for i,h in enumerate(col_headers)
                        if any(k in h.lower() for k in ['align','aim','relev'])),
                       n_cols-1 if n_cols>3 else None)
    parsed_rows = []
    for row in rows[1:]:
        cells = row.cells
        if not cells or not any(cell_plain(c) for c in cells): continue
        phase_raw = cell_plain(cells[0])
        phase_label = re.sub(r'\s+',' ',phase_raw).strip()
        phase_label = re.sub(r'^\d+\.\s*','',phase_label)
        teacher_bullets = cell_bullets(cells[teacher_col],abstract,num_to_abstract) if teacher_col<len(cells) else []
        student_bullets = cell_bullets(cells[student_col],abstract,num_to_abstract) if student_col<len(cells) else []
        # aims: parse into structured blocks
        aims_blocks = parse_aims_cell(cells[aims_col]) if aims_col is not None and aims_col<len(cells) else None
        if phase_label and not teacher_bullets and not student_bullets and not aims_blocks: continue
        parsed_rows.append({
            'phase': phase_label,
            'teacher': teacher_bullets,
            'student': student_bullets,
            'aims': aims_blocks,
        })
    return {'col_headers':col_headers,'teacher_col':teacher_col,'student_col':student_col,'aims_col':aims_col,'rows':parsed_rows}

# ── diff tables ────────────────────────────────────────────────────────────
def parse_diff_table_standard(tbl):
    rows = tbl.rows
    if not rows: return {'type':'standard','headers':[],'stages':[],'has_role_col':False}
    headers = [cell_plain(c) for c in rows[0].cells]
    lower_col  = next((i for i,h in enumerate(headers) if 'lower'  in h.lower()),1)
    higher_col = next((i for i,h in enumerate(headers) if 'higher' in h.lower()),2)
    role_col   = next((i for i,h in enumerate(headers) if h.lower()=='role'),None)
    stages = []; UNNAMED = '__UNNAMED__'
    for row in rows[1:]:
        cells = row.cells
        if not cells or not any(cell_plain(c) for c in cells): continue
        stage_val  = cell_plain(cells[0])
        lower_val  = cell_plain(cells[lower_col])  if lower_col <len(cells) else ''
        higher_val = cell_plain(cells[higher_col]) if higher_col<len(cells) else ''
        role_val   = cell_plain(cells[role_col])   if role_col is not None and role_col<len(cells) else None
        if role_col is not None:
            if stage_val and (not stages or stage_val!=stages[-1]['stage']):
                stages.append({'stage':stage_val,'teacher_lower':'','teacher_higher':'','student_lower':'','student_higher':''})
            if stages:
                if role_val and 'teacher' in role_val.lower():
                    stages[-1]['teacher_lower']=lower_val; stages[-1]['teacher_higher']=higher_val
                elif role_val and 'student' in role_val.lower():
                    stages[-1]['student_lower']=lower_val; stages[-1]['student_higher']=higher_val
        else:
            if stage_val:
                stages.append({'stage':stage_val,'lower':lower_val,'higher':higher_val})
            else:
                if not stages: stages.append({'stage':UNNAMED,'lower':lower_val,'higher':higher_val})
                else:
                    stages[-1]['lower']=(stages[-1].get('lower','')+' '+lower_val).strip()
                    stages[-1]['higher']=(stages[-1].get('higher','')+' '+higher_val).strip()
    return {'type':'standard','headers':headers,'stages':stages,'has_role_col':role_col is not None,'unnamed_sentinel':UNNAMED}

def parse_diff_table_ibl(tbl):
    rows = tbl.rows
    if not rows: return {'type':'ibl_matrix','features':[]}
    features = []; seen = {}
    for row in rows[2:]:
        cells = row.cells
        if len(cells)<6: continue
        feature_val=clean(cells[0].text); role_val=clean(cells[1].text)
        level_cells=[clean(cells[i].text) for i in range(2,6)]
        if not feature_val and not role_val: continue
        if feature_val and feature_val not in seen:
            seen[feature_val]=len(features)
            features.append({'name':feature_val,'teacher':[],'students':[]})
        fi=seen.get(feature_val,len(features)-1)
        if fi>=len(features): continue
        if 'teacher' in role_val.lower(): features[fi]['teacher']=level_cells
        elif 'student' in role_val.lower(): features[fi]['students']=level_cells
    return {'type':'ibl_matrix','features':features}

# ── what is X? ─────────────────────────────────────────────────────────────
def parse_what_is(doc, docx_path, abstract, num_to_abstract):
    items=[]; in_what_is=False; body=doc.element.body
    end_markers=['le facilitated','learning experience (le) in the classroom']
    for child in body:
        tag=child.tag.split('}')[-1]
        if tag=='p':
            from docx.text.paragraph import Paragraph as DP
            p=DP(child,doc); text=para_text(p); style=para_style(p); tl=text.lower()
            if 'what is' in tl and any(s in style for s in ('Title','Heading','Normal')):
                in_what_is=True; continue
            if in_what_is:
                if any(m in tl for m in end_markers) and ('Heading' in style or 'Title' in style): break
                if not text and not has_image(p): continue
                if has_image(p):
                    src=extract_image_b64(p,docx_path)
                    if src:
                        caption=re.sub(r'^\.\s*','',text).strip()
                        items.append({'type':'image','src':src,'caption':caption})
                    continue
                if 'Caption' in style:
                    if items and items[-1]['type']=='image' and not items[-1]['caption']:
                        items[-1]['caption']=re.sub(r'^\.\s*','',text).strip()
                    continue
                level=get_bullet_level(p,abstract,num_to_abstract)
                items.append({'type':'bullet','text':text,'tech':is_tech(p),'level':level})
        elif tag=='tbl' and in_what_is:
            from docx.table import Table as DT
            tbl=DT(child,doc); rows=tbl.rows
            if not rows: continue
            hdrs=[clean(c.text) for c in rows[0].cells]
            if len(hdrs)==2 and hdrs[0].lower() in ('component',''):
                for row in rows[1:]:
                    num=clean(row.cells[0].text); activity=clean(row.cells[1].text)
                    if num and activity:
                        items.append({'type':'bullet','text':f'{num}. {activity}','tech':False,'level':0})
    return items

def parse_le_preamble(doc, abstract, num_to_abstract):
    items=[]; in_le=False; body=doc.element.body
    le_markers=['le facilitated','learning experience (le) in the classroom']
    for child in body:
        tag=child.tag.split('}')[-1]
        if tag=='p':
            from docx.text.paragraph import Paragraph as DP
            p=DP(child,doc); text=para_text(p); style=para_style(p); tl=text.lower()
            if any(m in tl for m in le_markers): in_le=True; continue
            if in_le:
                if ('Heading' in style or 'Title' in style) and text: break
                if not text: continue
                level=get_bullet_level(p,abstract,num_to_abstract)
                items.append({'text':text,'tech':is_tech(p),'level':level})
        elif tag=='tbl' and in_le: break
    return items

def parse_key_question(doc):
    for tbl in doc.tables:
        if tbl.rows and tbl.columns:
            text=clean(tbl.rows[0].cells[0].text)
            if text.lower().startswith('key question'): return text
    return None

def parse_diff_preamble(doc):
    items=[]; in_diff=False; body=doc.element.body
    for child in body:
        tag=child.tag.split('}')[-1]
        if tag=='p':
            from docx.text.paragraph import Paragraph as DP
            p=DP(child,doc); text=para_text(p); style=para_style(p)
            if 'differentiating' in text.lower() and 'Heading' in style:
                in_diff=True; continue
            if in_diff:
                if not text: continue
                items.append({'text':text})
        elif tag=='tbl' and in_diff: break
    return items

def parse_doc(docx_path, approach_key):
    doc=Document(docx_path)
    abstract,num_to_abstract=build_numbering_lookup(doc)
    data={}
    data['what_is']      =parse_what_is(doc,docx_path,abstract,num_to_abstract)
    data['le_preamble']  =parse_le_preamble(doc,abstract,num_to_abstract)
    data['diff_preamble']=parse_diff_preamble(doc)
    kq=parse_key_question(doc)
    if kq: data['key_question']=kq
    tables=list(doc.tables)
    if approach_key in ('pbl','ibl'):
        phase_tbl=tables[1] if len(tables)>1 else None
        diff_tbl =tables[2] if len(tables)>2 else None
    else:
        phase_tbl=tables[0] if len(tables)>0 else None
        diff_tbl =tables[1] if len(tables)>1 else None
    if phase_tbl:
        data['phase_table']=parse_phase_table(phase_tbl,abstract,num_to_abstract)
    if diff_tbl:
        if approach_key=='ibl': data['diff_table']=parse_diff_table_ibl(diff_tbl)
        else: data['diff_table']=parse_diff_table_standard(diff_tbl)
    return data

def main():
    docs={'pbl':SCRIPT_DIR/'PBL.docx','ibl':SCRIPT_DIR/'IBL.docx','cbl':SCRIPT_DIR/'CBL.docx','ssi':SCRIPT_DIR/'SSI.docx'}
    approach_data={}
    for key,path in docs.items():
        if not path.exists(): print(f'WARNING: {path} not found'); approach_data[key]={}; continue
        print(f'Parsing {path.name}...')
        approach_data[key]=parse_doc(str(path),key)
        print(f'  Done: {key}')
    template_path=SCRIPT_DIR/'template.html'
    output_path  =SCRIPT_DIR/'pedagogy-hub.html'
    if not template_path.exists(): print(f'ERROR: template.html not found'); return
    template=template_path.read_text(encoding='utf-8')
    for key,data in approach_data.items():
        template=template.replace(f'__{key.upper()}_DATA__',json.dumps(data,ensure_ascii=False))
    output_path.write_text(template,encoding='utf-8')
    print(f'\nOutput written to {output_path}')

if __name__=='__main__':
    main()
