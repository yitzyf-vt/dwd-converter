#!/usr/bin/env python3
"""
DWD to DOCX Converter  –  DavkaWriter .dwd → Microsoft Word .docx

Usage:
    python dwd_to_docx.py input.dwd [output.docx] [--no-nikud] [--no-trup]

    Nikud and trup are included by default. Use --no-nikud / --no-trup to omit.

DavkaWriter format versions (auto-detected by run signature):
    Format A  49 80 01 00 01 00 00 00 02  — DavkaWriter for Windows (original)
    Format B  08 82 01 00 01 00 00 00 02  — DavkaWriter Gold / newer editions
    Format C  11 81 01 00 01 00 00 00 02  — DavkaWriter alternate edition

All formats share the same Hebrew consonant encoding (DAVKA_MAP).
Hebrew vs English detection: for Format A, the style byte determines the font.
For Formats B/C, content-based detection is used (alef byte 0x60 = definitive Hebrew marker).

Nikud bytes (follow the consonant they vowelize):
    0x9f=holam  0xa1=hataf-segol  0xa2=hataf-kamatz  0xa3=hataf-patach
    0xa4=segol  0xa5=tsere  0xa6=hiriq  0xa7=sheva
    0xa8=kamatz  0xa9=patach  0xaa=kubutz
    (confirmed by byte-alignment vs vocalized Mishna Shabbos 1:1)

Trup bytes (ta'amei hamikra / cantillation marks):
    0x84=QARNEY PARA  0x85=MERKHA KEFULA  0xab=SOF PASUK  0xac=MERKHA
    0xad=TIPEHA       0xae=ETNAHTA        0xaf=DARGA       0xb2=TEVIR
    0xb3=MAHAPAKH     0xb4=MUNAH          0xb5=YETIV       0xb6=YERACH BEN YOMO
    0xb8=QADMA        0xb9=GERESH         0xba=GERSHAYIM   0xbb=ZARQA
    0xbc=SEGOL(trup)  0xbd=ZAQEF QATAN   0xbe=ZAQEF GADOL  0xbf=PAZER
    0xc0=REVIA        0xc1=TELISHA QETANA  0xc2=TELISHA GEDOLA  0xc3=SHALSHELET
    0xc6=PASHTA       0xce=MAQAF
    (confirmed by Sha'ar Hata'amim worksheets, 21368)
"""

import io, re, sys, zipfile
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Consonant map ─────────────────────────────────────────────────────────────
DAVKA_MAP = {
    0x60:'א', 0x61:'ב', 0x62:'ג', 0x63:'ד', 0x64:'ה', 0x65:'ו', 0x66:'ז',
    0x67:'ח', 0x68:'ט', 0x69:'י', 0x6a:'כ', 0x6b:'כ', 0x6c:'ל', 0x6d:'ם',
    0x6e:'מ', 0x6f:'ן', 0x70:'נ', 0x71:'ס', 0x72:'ע', 0x73:'ף', 0x74:'פ',
    0x75:'ץ', 0x76:'צ', 0x77:'ק', 0x78:'ר', 0x79:'ש', 0x7a:'ת',
    # uppercase dagesh variants (confirmed from text)
    0x41:'ב\u05BC', 0x49:'י\u05BC', 0x4d:'כ\u05BC', 0x4e:'ל\u05BC',
    0x4f:'מ\u05BC', 0x51:'ס\u05BC', 0x52:'פ\u05BC',
    0x57:'ש\u05C1',   # shin-dot
    # uppercase alternates
    0x42:'ג', 0x43:'ד', 0x44:'ה', 0x45:'ו', 0x46:'ו', 0x47:'ג', 0x48:'ט',
    0x4a:'כ', 0x4b:'ך', 0x4c:'ל', 0x50:'נ', 0x53:'ש', 0x54:'ק', 0x55:'ש',
    0x56:'צ', 0x58:'ש', 0x59:'ת', 0x5a:'פ',
}

NIKUD_MAP = {
    0x9f:'\u05B9', 0xa1:'\u05B1', 0xa2:'\u05B3', 0xa3:'\u05B2',
    0xa4:'\u05B6', 0xa5:'\u05B5', 0xa6:'\u05B4', 0xa7:'\u05B0',
    0xa8:'\u05B8', 0xa9:'\u05B7', 0xaa:'\u05BB', 0xfb:'\u05B9',
}

TRUP_MAP = {
    0x84:'\u059F', 0x85:'\u05A6', 0xab:'\u05C3', 0xac:'\u05A5',
    0xad:'\u0596', 0xae:'\u0591', 0xaf:'\u05A7', 0xb2:'\u059B',
    0xb3:'\u05A4', 0xb4:'\u05A3', 0xb5:'\u059A', 0xb6:'\u05AA',
    0xb8:'\u05A8', 0xb9:'\u059C', 0xba:'\u059E', 0xbb:'\u0598',
    0xbc:'\u0592', 0xbd:'\u0594', 0xbe:'\u0595', 0xbf:'\u05A1',
    0xc0:'\u0597', 0xc1:'\u05A9', 0xc2:'\u05A0', 0xc3:'\u0593',
    0xc6:'\u0599', 0xce:'\u05BE',
}

# Style constants (Format A only — Formats B/C use content detection)
HEB_STYLES      = {0x20,0x23,0x25,0x28,0x29,0x2b,0x2c}
KEYWORD_HEB_STY = {0x23,0x25}
KEYWORD_DEF_STY = 0x24
HEB_HEADING_STY = 0x2b
MISHNA_STYS     = {0x29,0x20}
SECTION_HDR_STY = 0x27
BOLD_STY        = 0x1f

# ── Format detection ──────────────────────────────────────────────────────────
# Four confirmed DavkaWriter run-signature variants:
_FORMATS = [
    # (run_signature, para_signature, name)
    (bytes([0x49,0x80,0x01,0x00,0x01,0x00,0x00,0x00,0x02]),
     bytes([0x47,0x80,0x01,0x00,0x01]),
     'Format A (DavkaWriter original, 49-80)'),
    (bytes([0x08,0x82,0x01,0x00,0x01,0x00,0x00,0x00,0x02]),
     bytes([0x06,0x82,0x01,0x00,0x01]),
     'Format B (DavkaWriter Gold, 08-82)'),
    (bytes([0x11,0x81,0x01,0x00,0x01,0x00,0x00,0x00,0x02]),
     bytes([0x0f,0x81,0x01,0x00,0x01]),
     'Format C (DavkaWriter alternate, 11-81)'),
    (bytes([0x49,0x81,0x01,0x00,0x01,0x00,0x00,0x00,0x02]),
     bytes([0x47,0x81,0x01,0x00,0x01]),
     'Format D (DavkaWriter poster/style, 49-81)'),
]

_BAD_XML = re.compile(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]')


# ── Parser ────────────────────────────────────────────────────────────────────
def _is_hebrew_content(raw):
    """Content heuristic for Formats B/C.

    Alef (0x60) and nikud/trup bytes are definitively Hebrew.
    Digit or uppercase-initial with no alef/diacritics → English.
    (Windows-1252 special chars like smart quotes don't affect detection.)
    Everything else defaults to Hebrew.
    """
    if not raw: return False
    DIACRIT = set(NIKUD_MAP) | set(TRUP_MAP)
    if any(b in DIACRIT for b in raw): return True   # nikud/trup → Hebrew
    if 0x60 in raw:                    return True   # alef → Hebrew
    # Digit or uppercase-initial with no Hebrew markers → English
    if 0x41 <= raw[0] <= 0x5A or 0x30 <= raw[0] <= 0x39:
        return False
    return True   # default Hebrew


def _build_style_map(events):
    """Two-pass style→language map from unambiguous runs in this file.
    A style confirmed Hebrew anywhere is treated as Hebrew everywhere.
    Uses specific nikud/trup byte-sets rather than generic high-byte check
    to avoid mis-classifying English runs with Windows special characters.
    """
    DIACRIT = set(NIKUD_MAP) | set(TRUP_MAP)
    style_heb = {}
    for ev in events:
        if ev['type'] != 'run': continue
        raw, sty = ev['bytes'], ev['style']
        if any(b in DIACRIT for b in raw) or 0x60 in raw:
            style_heb[sty] = True                  # definitively Hebrew
        elif 0x41 <= raw[0] <= 0x5A and all(0x20 <= b <= 0x7E for b in raw):
            if sty not in style_heb:
                style_heb[sty] = False             # English unless contradicted
    return style_heb


def parse_dwd(data):
    """Parse DWD binary into (events, format_name).

    Returns (events, format_name). Supported DavkaWriter format variants:
      Format A (49-80): DavkaWriter original — style byte determines language
      Format B (08-82): DavkaWriter Gold — content + per-file style-map detection
      Format C (11-81): DavkaWriter alternate — same as Format B
    """
    sig_counts = [data.count(fmt[0]) for fmt in _FORMATS]
    best = max(range(len(sig_counts)), key=lambda i: sig_counts[i])
    run_sig, para_sig, fmt_name = _FORMATS[best]
    use_style_detection = (best == 0)

    # JPEG signatures to detect embedded images
    JPEG_SIGS = (b'\xFF\xD8\xFF\xE0', b'\xFF\xD8\xFF\xE1', b'\xFF\xD8\xFF\xDB')

    events, i, n = [], 0, len(data)
    while i < n - 12:
        # ── Image detection ───────────────────────────────────────────────────
        if data[i:i+3] == b'\xFF\xD8\xFF' and i + 4 < n and data[i:i+4] in JPEG_SIGS:
            eoi = data.find(b'\xFF\xD9', i + 4)
            if eoi > 0 and eoi - i > 5000:   # skip tiny false positives
                raw = data[i:eoi+2]
                try:
                    from PIL import Image as _PIL
                    img = _PIL.open(io.BytesIO(raw))
                    w, h = img.size
                    if w > 50 and h > 50:
                        events.append({'type': 'image', 'raw': raw,
                                       'width': w, 'height': h, 'fmt': 'JPEG'})
                        i = eoi + 2; continue
                except Exception:
                    pass
        # ── Run detection ─────────────────────────────────────────────────────
        if data[i:i+9] == run_sig:
            sty, hi, ln = data[i+9], data[i+10], data[i+11]
            if hi == 0 and ln > 0 and i+12+ln <= n:
                events.append({
                    'type': 'run', 'style': sty,
                    'bytes': data[i+12:i+12+ln],
                    'use_style': use_style_detection,
                })
                i += 12 + ln; continue
        if data[i:i+5] == para_sig:
            events.append({'type': 'para'})
            i += 5; continue
        i += 1

    # For non-Format-A: second pass to tag each event with is_hebrew
    if not use_style_detection:
        style_map = _build_style_map(events)
        for ev in events:
            if ev['type'] == 'run':
                raw, sty = ev['bytes'], ev['style']
                DIACRIT = set(NIKUD_MAP) | set(TRUP_MAP)
                if any(b in DIACRIT for b in raw) or 0x60 in raw:
                    ev['is_hebrew'] = True
                elif sty in style_map:
                    ev['is_hebrew'] = style_map[sty]
                else:
                    ev['is_hebrew'] = _is_hebrew_content(raw)

    return events, fmt_name


# ── Decoder ───────────────────────────────────────────────────────────────────
def _clean(s): return _BAD_XML.sub('', s)

def decode_heb(raw, with_nikud=True, with_trup=True):
    out = []
    for b in raw:
        if b in DAVKA_MAP:
            out.append(DAVKA_MAP[b])
        elif b == 0x20:
            out.append(' ')
        elif b < 0x80 and chr(b) in ',.;:!?\'"()-/0123456789\'':
            out.append(chr(b))
        elif b in NIKUD_MAP and with_nikud and out:
            out.append(NIKUD_MAP[b])
        elif b in TRUP_MAP and with_trup and out:
            out.append(TRUP_MAP[b])
    return _clean(''.join(out))

def decode_run(ev, with_nikud=True, with_trup=True):
    sty, raw = ev['style'], ev['bytes']
    start = 0
    if raw and raw[0] < 0x20:
        start = 1
        if len(raw) > 1 and raw[1] < 0x20:
            start = 2
    content = raw[start:]

    if ev.get('use_style', False):
        if sty in HEB_STYLES:
            return decode_heb(content, with_nikud, with_trup)
        return _clean(raw.decode('ascii', errors='replace'))

    # Formats B & C: use pre-computed is_hebrew tag if available
    is_heb = ev.get('is_hebrew', _is_hebrew_content(content))
    if is_heb:
        return decode_heb(content, with_nikud, with_trup)
    return _clean(content.decode('ascii', errors='replace'))

def has_page_break(ev):
    return (ev['type'] == 'run'
            and ev['style'] in MISHNA_STYS
            and ev['bytes']
            and ev['bytes'][0] == 0x0c)

def is_heb(sty): return sty in HEB_STYLES


# ── Document model ────────────────────────────────────────────────────────────
class Block: pass

class TextBlock(Block):
    def __init__(self, role):
        self.role = role
        self.runs = []
    @property
    def text(self): return ''.join(t for _,t in self.runs)
    def add(self, s, t):
        if t: self.runs.append((s, t))

class KeyWordBlock(Block):
    def __init__(self):
        self.pairs = []

class ImageBlock(Block):
    """An embedded image extracted from the DWD file."""
    def __init__(self, raw, width, height, fmt='JPEG', index=0):
        self.raw    = raw      # bytes
        self.width  = width    # pixels
        self.height = height
        self.fmt    = fmt
        self.index  = index    # sequential image number


def build_model(events, with_nikud=True, with_trup=True):
    blocks = []
    cur_text    = None
    cur_kw      = None
    in_kw_sec   = False
    kw_line_heb = []
    kw_line_eng = []
    para_gap    = 0
    heading_count = 0

    def flush_text():
        nonlocal cur_text
        if cur_text and cur_text.runs:
            blocks.append(cur_text)
        cur_text = None

    def flush_kw_line():
        nonlocal kw_line_heb, kw_line_eng
        heb = ' '.join(h for h in kw_line_heb if h)
        eng = ' '.join(e for e in kw_line_eng if e)
        if heb or eng:
            cur_kw.pairs.append((heb, eng))
        kw_line_heb = []; kw_line_eng = []

    def flush_kw():
        nonlocal cur_kw, in_kw_sec, kw_line_heb, kw_line_eng
        if in_kw_sec: flush_kw_line()
        if cur_kw and cur_kw.pairs: blocks.append(cur_kw)
        cur_kw = None; in_kw_sec = False; kw_line_heb = []; kw_line_eng = []

    img_index = [0]   # mutable counter for image numbering

    for ev in events:
        if ev['type'] == 'image':
            flush_text(); flush_kw()
            blocks.append(ImageBlock(ev['raw'], ev['width'], ev['height'],
                                     ev.get('fmt','JPEG'), img_index[0]))
            img_index[0] += 1
            continue
        if ev['type'] == 'para':
            para_gap += 1
            if in_kw_sec:
                flush_kw_line()
            else:
                if para_gap >= 2:
                    flush_text()
                    if blocks and not (isinstance(blocks[-1], TextBlock)
                                       and blocks[-1].role == 'blank'):
                        blocks.append(TextBlock('blank'))
                    para_gap = 0
                elif cur_text and cur_text.runs:
                    flush_text()
            continue

        para_gap = 0
        sty  = ev['style']
        text = decode_run(ev, with_nikud, with_trup)
        if sty in (SECTION_HDR_STY, HEB_HEADING_STY):
            text = text.strip()
        if not text:
            continue

        if has_page_break(ev):
            flush_text(); flush_kw()
            blocks.append(TextBlock('page_break'))
            rest = decode_run({'type':'run','style':sty,
                               'bytes': ev['bytes'][1:],
                               'use_style': ev.get('use_style', False)},
                              with_nikud, with_trup).strip()
            if rest:
                cur_text = TextBlock('mishna'); cur_text.add(sty, rest)
            continue

        if sty == SECTION_HDR_STY:
            flush_text(); flush_kw()
            hdr = TextBlock('section_hdr'); hdr.add(sty, text)
            blocks.append(hdr)
            if text == 'KEY WORDS':
                in_kw_sec = True; cur_kw = KeyWordBlock()
            continue

        if in_kw_sec:
            if sty == HEB_HEADING_STY:
                flush_kw()
            else:
                if is_heb(sty) or _is_hebrew_content(ev['bytes']):
                    kw_line_heb.append(text)
                else:
                    kw_line_eng.append(text)
                continue

        if sty == HEB_HEADING_STY:
            flush_text(); flush_kw()
            if heading_count > 0:
                blocks.append(TextBlock('page_break'))
            heading_count += 1
            tb = TextBlock('heading'); tb.add(sty, text)
            blocks.append(tb)
            continue

        if sty in MISHNA_STYS:
            if not cur_text or cur_text.role != 'mishna':
                flush_text(); cur_text = TextBlock('mishna')
            cur_text.add(sty, text)
            continue

        if not cur_text or cur_text.role != 'body':
            flush_text(); cur_text = TextBlock('body')
        cur_text.add(sty, text)

    flush_text(); flush_kw()
    return blocks


# ── DOCX XML helpers ──────────────────────────────────────────────────────────
def _pPr(p): return p._p.get_or_add_pPr()

def _bidi(p, on=True):
    pr = _pPr(p)
    el = OxmlElement('w:bidi'); el.set(qn('w:val'), '1' if on else '0'); pr.append(el)
    if on:
        jc = OxmlElement('w:jc'); jc.set(qn('w:val'), 'right'); pr.append(jc)

def _spacing(p, before=0, after=0, line=None):
    pr = _pPr(p)
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), str(before)); sp.set(qn('w:after'), str(after))
    if line: sp.set(qn('w:line'), str(line)); sp.set(qn('w:lineRule'), 'auto')
    pr.append(sp)

def _border_bottom(p, color='1F4E79', sz=12):
    pr = _pPr(p); pb = OxmlElement('w:pBdr')
    b = OxmlElement('w:bottom')
    b.set(qn('w:val'),'single'); b.set(qn('w:sz'),str(sz))
    b.set(qn('w:space'),'2');    b.set(qn('w:color'),color)
    pb.append(b); pr.append(pb)

def _box_border(p, color='2E75B6', fill='DBE9F9'):
    pr = _pPr(p); pb = OxmlElement('w:pBdr')
    for side in ('top','left','bottom','right'):
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'),'single'); b.set(qn('w:sz'),'6')
        b.set(qn('w:space'),'4');    b.set(qn('w:color'),color)
        pb.append(b)
    pr.append(pb)
    sh = OxmlElement('w:shd')
    sh.set(qn('w:val'),'clear'); sh.set(qn('w:color'),'auto'); sh.set(qn('w:fill'),fill)
    pr.append(sh)

def _rtl_run(para, text, font='David', sz=13, bold=False, color=None):
    r = para.add_run(text); r.font.name = font; r.font.size = Pt(sz)
    if bold:  r.font.bold = True
    if color: r.font.color.rgb = color
    rp = r._r.get_or_add_rPr()
    rt = OxmlElement('w:rtl'); rt.set(qn('w:val'),'1'); rp.append(rt)
    rf = rp.find(qn('w:rFonts'))
    if rf is None: rf = OxmlElement('w:rFonts'); rp.insert(0, rf)
    rf.set(qn('w:cs'), font)
    return r

def _ltr_run(para, text, font='Calibri', sz=11, bold=False, color=None):
    r = para.add_run(text); r.font.name = font; r.font.size = Pt(sz)
    if bold:  r.font.bold = True
    if color: r.font.color.rgb = color
    return r

def _inline_heb(para, text, sz=12, bold=False):
    r = para.add_run(text); r.font.name = 'David'; r.font.size = Pt(sz)
    if bold: r.font.bold = True
    rp = r._r.get_or_add_rPr()
    rt = OxmlElement('w:rtl'); rt.set(qn('w:val'),'1'); rp.append(rt)
    rf = rp.find(qn('w:rFonts'))
    if rf is None: rf = OxmlElement('w:rFonts'); rp.insert(0, rf)
    rf.set(qn('w:cs'), 'David')
    return r


# ── KEY WORDS table ───────────────────────────────────────────────────────────
def _add_kw_table(doc, pairs):
    NAVY = RGBColor(0x1F, 0x4E, 0x79)
    tbl = doc.add_table(rows=0, cols=2)
    tbl.style = 'Table Grid'
    tbl_xml = tbl._tbl
    tblPr = tbl_xml.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr'); tbl_xml.insert(0, tblPr)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '9360'); tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)
    tblLayout = OxmlElement('w:tblLayout')
    tblLayout.set(qn('w:type'), 'fixed')
    tblPr.append(tblLayout)

    def _mk_border(side):
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), '4')
        b.set(qn('w:space'), '0');    b.set(qn('w:color'), 'B8C4D4')
        return b

    tblBorders = OxmlElement('w:tblBorders')
    for s in ('top','left','bottom','right','insideH','insideV'):
        tblBorders.append(_mk_border(s))
    tblPr.append(tblBorders)

    for heb, eng in pairs:
        row = tbl.add_row()
        row.cells[0].width = Inches(2.0)
        row.cells[1].width = Inches(4.5)
        c0 = row.cells[0]; c0.width = Inches(2.0)
        p0 = c0.paragraphs[0]
        _bidi(p0, True); _spacing(p0, before=30, after=30)
        p0.paragraph_format.left_indent  = Inches(0.06)
        p0.paragraph_format.right_indent = Inches(0.06)
        tc0Pr = c0._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto')
        shd.set(qn('w:fill'),'EEF3FA'); tc0Pr.append(shd)
        _rtl_run(p0, heb, font='David', sz=12, bold=True, color=NAVY)
        c1 = row.cells[1]; c1.width = Inches(4.5)
        p1 = c1.paragraphs[0]
        _bidi(p1, False); _spacing(p1, before=30, after=30)
        p1.paragraph_format.left_indent = Inches(0.08)
        _ltr_run(p1, eng, sz=11)

    doc.add_paragraph()


# ── DOCX builder ─────────────────────────────────────────────────────────────
FONT_NOTE = (
    "Font note: This document uses the David font (built into Windows) as fallback. "
    "For authentic Davka Stam appearance, install one of these free fonts:\n"
    "  • Stam Ashkenaz CLM  →  hebrewfont.net  (best Davka Stam replacement; full nikud/trup)\n"
    "  • Frank Ruehl CLM    →  hebrewfont.net  (classic Israeli serif body font)\n"
    "  • Ezra SIL            →  software.sil.org/ezrasil/  (scholarly; full cantillation)\n"
    "  • Noto Serif Hebrew   →  fonts.google.com  (modern open serif)\n"
    "After installing: Ctrl+A to select all, then choose the font in the Home tab."
)
NAVY  = RGBColor(0x1F, 0x4E, 0x79)
GRAY  = RGBColor(0x50, 0x50, 0x50)


def build_docx(blocks, out_path):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Inches(8.5); sec.page_height = Inches(11)
    for attr in ('left_margin','right_margin','top_margin','bottom_margin'):
        setattr(sec, attr, Inches(0.9))
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(11)

    in_ysk_body = False

    for blk in blocks:
        if isinstance(blk, TextBlock):
            role = blk.role

            if role == 'page_break':
                in_ysk_body = False
                pb_p = doc.add_paragraph()
                run = pb_p.add_run()
                br = OxmlElement('w:br'); br.set(qn('w:type'), 'page')
                run._r.append(br)

            elif role == 'blank':
                p = doc.add_paragraph(); _spacing(p, before=40, after=40)
                in_ysk_body = False

            elif role == 'section_hdr':
                txt = blk.text.strip()
                p = doc.add_paragraph()
                _bidi(p, False); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _spacing(p, before=180, after=80); _box_border(p)
                _ltr_run(p, txt, font='Calibri', sz=12, bold=True, color=NAVY)
                in_ysk_body = (txt == 'YOU SHOULD KNOW')

            elif role == 'heading':
                p = doc.add_paragraph()
                _bidi(p, True); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _spacing(p, before=200, after=80); _border_bottom(p)
                _rtl_run(p, blk.text.strip(), font='David', sz=18, bold=True, color=NAVY)
                in_ysk_body = False

            elif role == 'mishna':
                p = doc.add_paragraph()
                _bidi(p, True); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                _spacing(p, before=60, after=120, line=320)
                if blk.text.strip():
                    _rtl_run(p, blk.text.strip(), font='David', sz=14)
                in_ysk_body = False

            elif role == 'body':
                p = doc.add_paragraph()
                _bidi(p, False); _spacing(p, before=30, after=50)
                for sty, text in blk.runs:
                    if is_heb(sty) or _is_hebrew_content(text.encode('utf-8', 'ignore')):
                        _inline_heb(p, text, sz=12)
                    else:
                        _ltr_run(p, text, sz=11, bold=(sty == BOLD_STY))

        elif isinstance(blk, ImageBlock):
            # Embed image inline; scale to fit 6-inch page width
            MAX_W = 6.0
            w_in = min(MAX_W, blk.width / 100.0)
            p = doc.add_paragraph()
            _bidi(p, False)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _spacing(p, before=80, after=80)
            try:
                doc.add_picture(io.BytesIO(blk.raw), width=Inches(w_in))
                # python-docx adds the picture to the last paragraph; move caption
                cap = doc.add_paragraph()
                _bidi(cap, False); cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _spacing(cap, before=0, after=80)
                _ltr_run(cap,
                         f'Image {blk.index+1}  ({blk.width}×{blk.height} px)',
                         sz=9, color=GRAY)
            except Exception as _e:
                _ltr_run(p, f'[Image {blk.index+1}: could not embed — {_e}]', sz=10, color=GRAY)
            in_ysk_body = False

        elif isinstance(blk, KeyWordBlock):
            if blk.pairs:
                _add_kw_table(doc, blk.pairs)
            in_ysk_body = False

    # Font note page
    doc.add_page_break()
    note_hdr = doc.add_paragraph()
    _bidi(note_hdr, False); _spacing(note_hdr, before=0, after=80)
    _ltr_run(note_hdr, 'Font Information', sz=13, bold=True, color=NAVY)
    for line in FONT_NOTE.split('\n'):
        if not line.strip(): continue
        np = doc.add_paragraph()
        _bidi(np, False); _spacing(np, before=20, after=20)
        if line.startswith('  •'):
            np.paragraph_format.left_indent = Inches(0.3)
            parts = line.strip().lstrip('• ').split('→')
            _ltr_run(np, '• ', sz=11)
            if len(parts) == 2:
                _ltr_run(np, parts[0].strip() + '  ', sz=11, bold=True)
                _ltr_run(np, parts[1].strip(), sz=10, color=GRAY)
            else:
                _ltr_run(np, line.strip().lstrip('• '), sz=11)
        else:
            _ltr_run(np, line.strip(), sz=11)

    doc.save(out_path)


# ── Entry point ───────────────────────────────────────────────────────────────
def convert(inp, out=None, with_nikud=True, with_trup=True):
    """Convert a DWD file to DOCX (or ZIP containing DOCX + loose images).

    Returns the path of the output file (.docx or .zip).
    A ZIP is produced when images are found — it contains:
      • converted.docx  (with images embedded inline)
      • images/img_01.jpg, img_02.jpg, …  (original-quality loose copies)
    """
    p = Path(inp)
    if not p.exists():
        sys.exit(f'File not found: {inp}')

    print(f'Reading  {inp}  ({p.stat().st_size:,} bytes)')
    data = p.read_bytes()

    print('Parsing …')
    evts, fmt_name = parse_dwd(data)
    n_runs  = sum(1 for e in evts if e['type'] == 'run')
    n_paras = sum(1 for e in evts if e['type'] == 'para')
    n_imgs  = sum(1 for e in evts if e['type'] == 'image')
    print(f'  Format: {fmt_name}')
    print(f'  {n_runs} runs, {n_paras} para-breaks, {n_imgs} image(s)')

    print('Building document model …')
    blocks = build_model(evts, with_nikud, with_trup)
    pages  = sum(1 for b in blocks if isinstance(b, TextBlock) and b.role == 'page_break')
    kwtbls = sum(1 for b in blocks if isinstance(b, KeyWordBlock))
    img_blocks = [b for b in blocks if isinstance(b, ImageBlock)]
    print(f'  {len(blocks)} blocks  ({pages} page-breaks, {kwtbls} tables, {len(img_blocks)} images)')

    # Determine output path
    stem = p.stem
    if out is None:
        out = str(p.with_suffix('.zip' if img_blocks else '.docx'))
    # Force .zip extension when images present
    out_path = Path(out)
    if img_blocks and out_path.suffix.lower() != '.zip':
        out_path = out_path.with_suffix('.zip')

    docx_path = out_path.with_suffix('.docx') if img_blocks else out_path

    print('Building DOCX …')
    build_docx(blocks, str(docx_path))

    if img_blocks:
        print(f'Packaging ZIP with {len(img_blocks)} image(s) …')
        with zipfile.ZipFile(str(out_path), 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.write(str(docx_path), f'{stem}.docx')
            for blk in img_blocks:
                ext = blk.fmt.lower()
                img_name = f'images/img_{blk.index+1:02d}.{ext}'
                zf.writestr(img_name, blk.raw)
                print(f'  Added {img_name}  ({blk.width}×{blk.height} px)')
        docx_path.unlink()   # remove the intermediate .docx
        print(f'Saved → {out_path}')
        return str(out_path)

    print(f'Saved → {out_path}')
    return str(out_path)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(0)
    with_nikud = '--no-nikud' not in sys.argv
    with_trup  = '--no-trup'  not in sys.argv
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    result = convert(args[0], args[1] if len(args) > 1 else None, with_nikud, with_trup)
    print(f'Output: {result}')
