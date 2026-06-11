"""
build.py
========
Reads PBL.docx (and IBL/CBL/SSI when ready), parses content,
and injects into template.html to produce pedagogy-hub.html.

Usage:
    python3 build.py

Requirements:
    pip install python-docx
"""

import json, re
from pathlib import Path
from docx import Document

# ── PATHS ─────────────────────────────────────────────────────────────────
TEMPLATE   = Path('/tmp/template.html')
OUTPUT     = Path('/tmp/pedagogy-hub.html')

# Map approach key → docx path. Set to None when not yet available.
DOCX_PATHS = {
    'pbl': Path('/mnt/user-data/uploads/PBL.docx'),
    'ibl': None,
    'cbl': None,
    'ssi': None,
}

# ── SANITISE ──────────────────────────────────────────────────────────────
def clean(s):
    """Strip/normalise characters that cause rendering or JS problems."""
    if not isinstance(s, str):
        return s
    # Smart quotes → straight
    s = s.replace('\u2018', "'").replace('\u2019', "'")
    s = s.replace('\u201c', '"').replace('\u201d', '"')
    s = s.replace('\u201e', '"').replace('\u201f', '"')
    s = s.replace('\u2032', "'").replace('\u2033', '"')
    # Non-breaking space → regular space
    s = s.replace('\xa0', ' ')
    # Unicode bullet variants → remove (CSS supplies bullets)
    for ch in (
        '\u2022\u2023\u2024\u2025\u2026\u2027\u2043\u204c\u204d'
        '\u25aa\u25ab\u25b8\u25b9\u25cf\u25e6'
        '\u2192\u2193\u2190\u2194\u21b2'
        '\u25ba\u25b6\u25c0\u25c4'
        '\uf0b7\uf0d8\uf0a7\u00b7'
    ):
        s = s.replace(ch, '')
    # Private Use Area → strip
    s = re.sub(r'[\ue000-\uf8ff]', '', s)
    # Control characters (keep tab and newline)
    s = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', s)
    # Caret + control sequences
    s = re.sub(r'\^[\x00-\x1f\x40-\x5f]', '', s)
    # Collapse whitespace within a line
    s = re.sub(r'[^\S\n]+', ' ', s)
    return s.strip()

# ── DOCX PARSER ───────────────────────────────────────────────────────────

def cell_paragraphs(cell):
    """Return non-empty cleaned paragraph texts from a table cell."""
    return [clean(p.text) for p in cell.paragraphs if p.text.strip()]

def parse_teacher_cell(cell):
    """
    Parse teacher cell paragraphs into:
      - teacher: list of {tech: bool, text: str}
      - teacher_prompts: list of str

    Rules (derived from actual docx structure):
      - Paragraphs starting with '(Tech)' → tech bullet
      - Paragraphs that are quoted strings (start with " or ") following
        a 'Use prompts such as:' line → prompts
      - Everything else → normal bullet
    """
    paras = cell_paragraphs(cell)
    teacher = []
    prompts = []
    prompt_mode = False

    for p in paras:
        # Tech bullet
        if p.startswith('(Tech)'):
            prompt_mode = False
            text = clean(p[len('(Tech)'):].strip())
            if text:
                teacher.append({'tech': True, 'text': text})
            continue

        # Prompt trigger line
        if re.search(r'use prompts such as', p, re.IGNORECASE):
            prompt_mode = True
            # Extract anything after the colon on the same line
            after = re.split(r'use prompts such as\s*[:\-]?\s*', p, flags=re.IGNORECASE)[-1].strip()
            if after:
                prompts.append(clean(after))
            continue

        # Quoted line → prompt
        if prompt_mode and (p.startswith('"') or p.startswith('\u201c')):
            prompts.append(clean(p.strip('""\u201c\u201d')))
            continue

        # Tech continuation paragraph (belongs to previous Tech bullet)
        # e.g. "Teachers can also bring students..." after a Tech line
        if teacher and teacher[-1]['tech'] and not p.startswith('(Tech)'):
            # Check: does this look like a standalone normal bullet?
            # Heuristic: starts with a verb or capital → new bullet
            # Otherwise append to previous tech item
            if not re.match(r'^[A-Z][a-z]', p):
                teacher[-1]['text'] += ' ' + p
                continue

        # Normal bullet
        prompt_mode = False
        if p:
            teacher.append({'tech': False, 'text': p})

    return teacher, prompts


def parse_aims_cell(cell):
    """
    Parse aims cell into list of {tag: str, text: str}.
    Tags are [Appreciate], [Understand], [Take Action] at start of paragraph.
    """
    aims = []
    tag_re = re.compile(r'^\[([^\]]+)\]\s*(.*)', re.DOTALL)
    for p in cell_paragraphs(cell):
        m = tag_re.match(p)
        if m:
            aims.append({'tag': clean(m.group(1)), 'text': clean(m.group(2))})
        elif aims:
            # Continuation of previous aim
            aims[-1]['text'] += ' ' + clean(p)
    return aims


def parse_student_cell(cell):
    return [t for t in cell_paragraphs(cell) if t]


def parse_phase_cell(cell):
    """Return (name, subtitle) from phase cell."""
    lines = [clean(l) for l in cell.text.split('\n') if clean(l)]
    name     = lines[0] if lines else ''
    subtitle = lines[1] if len(lines) > 1 else ''
    return name, subtitle


def parse_docx(path):
    """
    Parse a pedagogical approach .docx and return a data dict
    compatible with template.html.
    """
    doc = Document(path)

    # ── Identify tables by size ──────────────────────────────────────────
    key_q_table   = None
    phase_table   = None
    diff_table    = None

    for t in doc.tables:
        cols = len(t.columns)
        rows = len(t.rows)
        header = t.rows[0].cells[0].text.strip()
        if cols == 1 and rows == 1:
            key_q_table = t
        elif cols == 4 and rows > 1:
            phase_table = t
        elif cols == 3 and rows > 1 and header.lower() not in ('phase', 'pbl phase'):
            diff_table = t
        elif cols == 3 and rows > 1:
            diff_table = t

    # ── Pre-table paragraphs ─────────────────────────────────────────────
    # Walk body elements to collect paragraphs before the first table
    from docx.oxml.ns import qn
    what_is_bullets    = []
    when_useful        = []
    char_questions     = []
    le_bullets         = []

    section = None   # 'what' | 'le'
    subsection = None  # 'when' | 'q' | None

    for elem in doc.element.body:
        tag = elem.tag.split('}')[-1]
        if tag == 'tbl':
            break
        if tag != 'p':
            continue
        from docx.text.paragraph import Paragraph
        p = Paragraph(elem, doc)
        t = clean(p.text)
        if not t:
            continue
        style = p.style.name

        # Section detection
        if style == 'Title' and re.search(r'what is', t, re.IGNORECASE):
            section = 'what'
            continue
        if style in ('Heading 1', 'Heading 2') and re.search(r'le facilitated|learning experience', t, re.IGNORECASE):
            section = 'le'
            subsection = None
            continue

        if section == 'what':
            tl = t.lower()
            if 'especially useful' in tl or 'useful when' in tl:
                subsection = 'intro_when'
                what_is_bullets.append(t)
                continue
            if 'following questions' in tl or 'characteristic' in tl:
                subsection = 'q'
                what_is_bullets.append(t)
                continue
            if subsection == 'intro_when' and not t[0].isupper():
                when_useful.append(t.rstrip(';').rstrip('.'))
                continue
            if subsection == 'intro_when' and t[0].isupper():
                subsection = None
            if subsection == 'q':
                char_questions.append(t.strip('""\u201c\u201d'))
                continue
            what_is_bullets.append(t)

        elif section == 'le':
            le_bullets.append(t)

    # ── Key question ─────────────────────────────────────────────────────
    key_question = ''
    key_question_body = ''
    if key_q_table:
        full = clean(key_q_table.rows[0].cells[0].text)
        # "Key Question: <question> <body>"
        m = re.match(r'Key Question[:\s]+(.+?)(?=\s{2,}|\n|Singapore)', full, re.IGNORECASE)
        if m:
            key_question = clean(m.group(1))
            key_question_body = clean(full[m.end():].strip())
        else:
            # Fallback: first sentence is the question
            sentences = full.split('. ')
            key_question = sentences[0].replace('Key Question:', '').replace('Key Question', '').strip(' :')
            key_question_body = '. '.join(sentences[1:])

    # ── Phase table ───────────────────────────────────────────────────────
    phases = []
    if phase_table:
        for row in phase_table.rows[1:]:
            cells = row.cells
            if len(cells) < 4:
                continue
            name, subtitle = parse_phase_cell(cells[0])
            if not name:
                continue
            teacher, prompts = parse_teacher_cell(cells[1])
            student          = parse_student_cell(cells[2])
            aims             = parse_aims_cell(cells[3])
            phases.append({
                'name':            name,
                'subtitle':        subtitle,
                'teacher':         teacher,
                'teacher_prompts': prompts,
                'student':         student,
                'aims':            aims,
            })

    # ── Diff table ────────────────────────────────────────────────────────
    diff_phases = []
    diff_intro  = ''
    if diff_table:
        for row in diff_table.rows[1:]:
            cells = row.cells
            if len(cells) < 3:
                continue
            name = clean(cells[0].text.split('\n')[0])
            if not name:
                continue
            lower  = [t for t in cell_paragraphs(cells[1]) if t]
            higher = [t for t in cell_paragraphs(cells[2]) if t]
            diff_phases.append({'name': name, 'lower': lower, 'higher': higher})

    # ── Diff intro (paragraphs between Heading 2 and diff table) ─────────
    found_diff_heading = False
    diff_intro_lines   = []
    for elem in doc.element.body:
        tag = elem.tag.split('}')[-1]
        if tag == 'tbl':
            if found_diff_heading:
                break
            continue
        if tag != 'p':
            continue
        from docx.text.paragraph import Paragraph
        p = Paragraph(elem, doc)
        t = clean(p.text)
        if not t:
            continue
        if p.style.name in ('Heading 1','Heading 2') and 'differentiat' in t.lower():
            found_diff_heading = True
            continue
        if found_diff_heading:
            diff_intro_lines.append(t)
    diff_intro = ' '.join(diff_intro_lines)

    # ── Assemble definition ───────────────────────────────────────────────
    # First bullet under "What is X?" is the definition paragraph
    definition = ''
    remaining_what = []
    for b in what_is_bullets:
        bl = b.lower()
        if not definition and len(b) > 80 and 'especially useful' not in bl and 'following questions' not in bl:
            definition = b
        else:
            remaining_what.append(b)

    return {
        'whatIs': {
            'definition':           definition,
            'whenUseful':           when_useful,
            'characteristicQuestions': char_questions,
        },
        'leIntro': {
            'bullets':          le_bullets,
            'keyQuestion':      key_question,
            'keyQuestionBody':  key_question_body,
        },
        'diffIntro':   diff_intro,
        'phases':      phases,
        'diffPhases':  diff_phases,
    }


# ── BUILD ─────────────────────────────────────────────────────────────────
def build():
    template = TEMPLATE.read_text(encoding='utf-8')

    for key, docx_path in DOCX_PATHS.items():
        placeholder = f'__{key.upper()}_DATA__'
        if docx_path and docx_path.exists():
            print(f'Parsing {docx_path.name}...')
            data = parse_docx(docx_path)
            template = template.replace(placeholder, json.dumps(data, ensure_ascii=False))
            print(f'  {len(data["phases"])} phases, {len(data["diffPhases"])} diff phases')
        else:
            template = template.replace(placeholder, 'null')
            print(f'{key.upper()}: no docx — placeholder set to null')

    OUTPUT.write_text(template, encoding='utf-8')
    print(f'\nBuilt: {OUTPUT} ({len(template):,} chars)')

if __name__ == '__main__':
    build()
