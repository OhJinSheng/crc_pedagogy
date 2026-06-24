#!/usr/bin/env python3
"""
build.py – Flexible parser + adaptive data injector for pedagogy-hub.html
Reads PBL.docx, IBL.docx, CBL.docx, SSI.docx and injects parsed JSON
into template.html to produce pedagogy-hub.html.
"""

import json
import re
import base64
import os
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn

SCRIPT_DIR = Path(os.path.abspath(__file__)) .parent if '__file__' in dir() else Path('/home/claude')

# ── helpers ───────────────────────────────────────────────────────────────

UNICODE_MAP = {
    '\u2018': "'", '\u2019': "'", '\u201c': '"', '\u201d': '"',
    '\u2013': '-', '\u2014': '--', '\u2026': '...', '\u00a0': ' ',
    '\uf0b7': '', '\uf0a7': '', '\uf020': '',
}

def clean(text):
    if not text:
        return ''
    for src, dst in UNICODE_MAP.items():
        text = text.replace(src, dst)
    text = re.sub(r'[\ue000-\uf8ff]', '', text)
    # collapse multiple spaces/newlines
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def para_text(p):
    return clean(''.join(run.text or '' for run in p.runs))

def is_tech(p):
    t = para_text(p)
    return bool(re.match(r'^\(Tech\)', t, re.IGNORECASE))

def strip_tech_prefix(t):
    return re.sub(r'^\(Tech\)\s*', '', t, flags=re.IGNORECASE).strip()

def para_style(p):
    return (p.style.name or '').strip()

def has_image(p):
    return bool(p._element.findall('.//' + qn('a:blip')))

def extract_image_b64(p, docx_path):
    blips = p._element.findall('.//' + qn('a:blip'))
    if not blips:
        return None
    rId = blips[0].get(qn('r:embed'))
    if not rId:
        return None
    doc2 = Document(docx_path)
    for rel in doc2.part.rels.values():
        if rel.reltype.endswith('/image') and rel.rId == rId:
            image_bytes = rel.target_part.blob
            b64 = base64.b64encode(image_bytes).decode('utf-8')
            mime = 'image/png' if image_bytes[:4] == b'\x89PNG' else 'image/jpeg'
            return f'data:{mime};base64,{b64}'
    return None

# ── table cell helpers ────────────────────────────────────────────────────

def cell_bullets(cell):
    """Return list of {text, tech} dicts for each non-empty paragraph in a cell."""
    bullets = []
    for p in cell.paragraphs:
        t = para_text(p)
        if not t:
            continue
        tech = is_tech(p)
        bullets.append({'text': strip_tech_prefix(t) if tech else t, 'tech': tech})
    return bullets

def cell_plain(cell):
    return clean(cell.text)

# ── phase table ───────────────────────────────────────────────────────────

def parse_phase_table(tbl):
    rows = tbl.rows
    if not rows:
        return None

    # Header row
    col_headers = [cell_plain(c) for c in rows[0].cells]
    n_cols = len(col_headers)

    # Detect column roles by header text
    teacher_col = next((i for i, h in enumerate(col_headers) if 'teacher' in h.lower()), 1)
    student_col = next((i for i, h in enumerate(col_headers) if 'student' in h.lower()), 2)
    aims_col    = next((i for i, h in enumerate(col_headers)
                        if any(k in h.lower() for k in ['align', 'aim', 'relev'])),
                       n_cols - 1 if n_cols > 3 else None)

    parsed_rows = []
    prev_phase = ''

    for row in rows[1:]:
        cells = row.cells
        if not cells:
            continue

        phase_raw = cell_plain(cells[0])
        # Build phase label: combine multi-line cell text cleanly
        phase_label = re.sub(r'\s+', ' ', phase_raw).strip()
        # Remove numeric prefix like "4. " from IBL
        phase_label = re.sub(r'^\d+\.\s*', '', phase_label)

        # Skip header-repeat rows or empty rows
        if not any(cell_plain(c) for c in cells):
            continue

        teacher_bullets = cell_bullets(cells[teacher_col]) if teacher_col < len(cells) else []
        student_bullets = cell_bullets(cells[student_col]) if student_col < len(cells) else []
        aims_text       = cell_plain(cells[aims_col]) if aims_col is not None and aims_col < len(cells) else ''

        # Skip rows that are just header labels repeated
        if phase_label and not teacher_bullets and not student_bullets and not aims_text:
            continue

        parsed_rows.append({
            'phase':   phase_label,
            'teacher': teacher_bullets,
            'student': student_bullets,
            'aims':    aims_text,
        })

    return {
        'col_headers': col_headers,
        'teacher_col': teacher_col,
        'student_col': student_col,
        'aims_col':    aims_col,
        'rows':        parsed_rows,
    }

# ── diff table ────────────────────────────────────────────────────────────

def parse_diff_table_standard(tbl, has_role_col=False):
    """PBL (3-col), SSI (3-col), CBL (4-col with Role)."""
    rows = tbl.rows
    if not rows:
        return {'type': 'standard', 'headers': [], 'stages': [], 'has_role_col': has_role_col}

    headers = [cell_plain(c) for c in rows[0].cells]
    n_cols  = len(headers)

    stage_col  = 0
    lower_col  = next((i for i, h in enumerate(headers) if 'lower'  in h.lower()), 1)
    higher_col = next((i for i, h in enumerate(headers) if 'higher' in h.lower()), 2)
    role_col   = next((i for i, h in enumerate(headers) if h.lower() == 'role'), None)

    stages = []

    # SSI has an empty Stage cell in row 1 for "Encountering a Focal Issue"
    # We need to infer the stage name from the phase table if possible.
    # Strategy: if first data row has no stage label, use a sentinel.
    UNNAMED = '__UNNAMED__'

    for row in rows[1:]:
        cells = row.cells
        if not cells or not any(cell_plain(c) for c in cells):
            continue

        stage_val  = cell_plain(cells[stage_col]) if stage_col < len(cells) else ''
        lower_val  = cell_plain(cells[lower_col])  if lower_col  < len(cells) else ''
        higher_val = cell_plain(cells[higher_col]) if higher_col < len(cells) else ''
        role_val   = cell_plain(cells[role_col])   if role_col is not None and role_col < len(cells) else None

        if role_col is not None:
            # CBL: Stage / Role / Lower / Higher
            if stage_val and (not stages or stage_val != stages[-1]['stage']):
                stages.append({'stage': stage_val, 'teacher_lower': '', 'teacher_higher': '',
                               'student_lower': '', 'student_higher': ''})
            if stages:
                if role_val and 'teacher' in role_val.lower():
                    stages[-1]['teacher_lower']  = lower_val
                    stages[-1]['teacher_higher'] = higher_val
                elif role_val and 'student' in role_val.lower():
                    stages[-1]['student_lower']  = lower_val
                    stages[-1]['student_higher'] = higher_val
        else:
            # PBL / SSI: Stage / Lower / Higher
            if stage_val:
                stages.append({'stage': stage_val, 'lower': lower_val, 'higher': higher_val})
            else:
                # row with no stage label — start new unnamed stage if first, else append
                if not stages:
                    stages.append({'stage': UNNAMED, 'lower': lower_val, 'higher': higher_val})
                else:
                    stages[-1]['lower']  = (stages[-1].get('lower','')  + ' ' + lower_val).strip()
                    stages[-1]['higher'] = (stages[-1].get('higher','') + ' ' + higher_val).strip()

    return {
        'type':         'standard',
        'headers':      headers,
        'stages':       stages,
        'has_role_col': role_col is not None,
        'unnamed_sentinel': UNNAMED,
    }

def parse_diff_table_ibl(tbl):
    """IBL 6-col two-axis matrix. Columns: Feature | Role | [4 guidance levels]."""
    rows = tbl.rows
    if not rows:
        return {'type': 'ibl_matrix', 'features': []}

    features = []
    seen_features = {}  # name -> index in features

    for row in rows[2:]:  # skip two header rows
        cells = row.cells
        if len(cells) < 6:
            continue

        feature_val = clean(cells[0].text)
        role_val    = clean(cells[1].text)
        level_cells = [clean(cells[i].text) for i in range(2, 6)]

        if not feature_val and not role_val:
            continue

        if feature_val and feature_val not in seen_features:
            seen_features[feature_val] = len(features)
            features.append({'name': feature_val, 'teacher': [], 'students': []})

        feat_idx = seen_features.get(feature_val, len(features) - 1)
        if feat_idx >= len(features):
            continue

        if 'teacher' in role_val.lower():
            features[feat_idx]['teacher'] = level_cells
        elif 'student' in role_val.lower():
            features[feat_idx]['students'] = level_cells

    return {'type': 'ibl_matrix', 'features': features}

# ── "What is X?" ──────────────────────────────────────────────────────────

def parse_what_is(doc, docx_path):
    items = []
    in_what_is = False
    body = doc.element.body

    end_markers = ['le facilitated', 'learning experience (le) in the classroom']

    for child in body:
        tag = child.tag.split('}')[-1]

        if tag == 'p':
            from docx.text.paragraph import Paragraph as DocxPara
            p = DocxPara(child, doc)
            text  = para_text(p)
            style = para_style(p)
            tl    = text.lower()

            if 'what is' in tl and any(s in style for s in ('Title', 'Heading', 'Normal')):
                in_what_is = True
                continue

            if in_what_is:
                if any(m in tl for m in end_markers) and ('Heading' in style or 'Title' in style):
                    break
                if not text and not has_image(p):
                    continue
                if has_image(p):
                    src = extract_image_b64(p, docx_path)
                    if src:
                        # caption text may be in same paragraph
                        caption = re.sub(r'^\.\s*', '', text).strip()
                        items.append({'type': 'image', 'src': src, 'caption': caption})
                    continue
                if 'Caption' in style:
                    # standalone caption after image — attach to last image if present
                    if items and items[-1]['type'] == 'image' and not items[-1]['caption']:
                        items[-1]['caption'] = re.sub(r'^\.\s*', '', text).strip()
                    continue
                items.append({'type': 'bullet', 'text': text, 'tech': is_tech(p)})

        elif tag == 'tbl' and in_what_is:
            from docx.table import Table as DocxTable
            tbl = DocxTable(child, doc)
            rows = tbl.rows
            if not rows:
                continue
            # IBL component table: flatten to numbered bullets
            headers = [clean(c.text) for c in rows[0].cells]
            if len(headers) == 2 and headers[0].lower() in ('component', ''):
                for row in rows[1:]:
                    num      = clean(row.cells[0].text)
                    activity = clean(row.cells[1].text)
                    if num and activity:
                        items.append({'type': 'bullet', 'text': f'{num}. {activity}', 'tech': False})
            # other tables in what_is: skip

    return items

# ── LE preamble ───────────────────────────────────────────────────────────

def parse_le_preamble(doc):
    items = []
    in_le = False
    body = doc.element.body

    le_markers = ['le facilitated', 'learning experience (le) in the classroom']

    for child in body:
        tag = child.tag.split('}')[-1]

        if tag == 'p':
            from docx.text.paragraph import Paragraph as DocxPara
            p = DocxPara(child, doc)
            text  = para_text(p)
            style = para_style(p)
            tl    = text.lower()

            if any(m in tl for m in le_markers):
                in_le = True
                continue

            if in_le:
                if ('Heading' in style or 'Title' in style) and text:
                    break
                if not text:
                    continue
                items.append({'text': text, 'tech': is_tech(p)})

        elif tag == 'tbl' and in_le:
            break

    return items

# ── Key Question ──────────────────────────────────────────────────────────

def parse_key_question(doc):
    for tbl in doc.tables:
        if tbl.rows and tbl.columns:
            text = clean(tbl.rows[0].cells[0].text)
            if text.lower().startswith('key question'):
                return text
    return None

# ── Diff preamble ─────────────────────────────────────────────────────────

def parse_diff_preamble(doc):
    items = []
    in_diff = False
    body = doc.element.body

    for child in body:
        tag = child.tag.split('}')[-1]

        if tag == 'p':
            from docx.text.paragraph import Paragraph as DocxPara
            p = DocxPara(child, doc)
            text  = para_text(p)
            style = para_style(p)

            if 'differentiating' in text.lower() and 'Heading' in style:
                in_diff = True
                continue

            if in_diff:
                if not text:
                    continue
                items.append({'text': text})

        elif tag == 'tbl' and in_diff:
            break

    return items

# ── Per-doc orchestrator ──────────────────────────────────────────────────

def parse_doc(docx_path, approach_key):
    doc = Document(docx_path)
    data = {}

    data['what_is']    = parse_what_is(doc, docx_path)
    data['le_preamble']= parse_le_preamble(doc)
    data['diff_preamble'] = parse_diff_preamble(doc)

    kq = parse_key_question(doc)
    if kq:
        data['key_question'] = kq

    tables = list(doc.tables)

    if approach_key == 'pbl':
        phase_tbl = tables[1] if len(tables) > 1 else None
        diff_tbl  = tables[2] if len(tables) > 2 else None
    elif approach_key == 'ibl':
        phase_tbl = tables[1] if len(tables) > 1 else None
        diff_tbl  = tables[2] if len(tables) > 2 else None
    else:
        phase_tbl = tables[0] if len(tables) > 0 else None
        diff_tbl  = tables[1] if len(tables) > 1 else None

    if phase_tbl:
        data['phase_table'] = parse_phase_table(phase_tbl)

    if diff_tbl:
        if approach_key == 'ibl':
            data['diff_table'] = parse_diff_table_ibl(diff_tbl)
        else:
            has_role = (approach_key == 'cbl')
            data['diff_table'] = parse_diff_table_standard(diff_tbl, has_role_col=has_role)

    return data

# ── Main ──────────────────────────────────────────────────────────────────

def main():
    docs = {
        'pbl': SCRIPT_DIR / 'PBL.docx',
        'ibl': SCRIPT_DIR / 'IBL.docx',
        'cbl': SCRIPT_DIR / 'CBL.docx',
        'ssi': SCRIPT_DIR / 'SSI.docx',
    }

    approach_data = {}
    for key, path in docs.items():
        if not path.exists():
            print(f'WARNING: {path} not found, skipping {key.upper()}')
            approach_data[key] = {}
            continue
        print(f'Parsing {path.name}...')
        approach_data[key] = parse_doc(str(path), key)
        print(f'  Done: {key}')

    template_path = SCRIPT_DIR / 'template.html'
    output_path   = SCRIPT_DIR / 'pedagogy-hub.html'

    if not template_path.exists():
        print(f'ERROR: template.html not found at {template_path}')
        return

    template = template_path.read_text(encoding='utf-8')

    for key, data in approach_data.items():
        placeholder = f'__{key.upper()}_DATA__'
        json_str    = json.dumps(data, ensure_ascii=False)
        template    = template.replace(placeholder, json_str)

    output_path.write_text(template, encoding='utf-8')
    print(f'\nOutput written to {output_path}')

if __name__ == '__main__':
    main()
