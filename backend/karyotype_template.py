"""
Karyotype Report Generator - PDF Template
==========================================
Generates Anderson Diagnostics Peripheral Blood Karyotyping reports.

Layout (pixel-perfect from pdfplumber analysis of reference PDFs):
  Page size   : 612 x 792 pt  (US Letter)
  Header img  : top=0,   bot=67.8,  x0=1.4,   x1=611.2
  Footer img  : top=743.8, bot=791.8, x0=1.4, x1=612.0
  Stamp img   : top=126.8, bot=205.6, x0=276.4, x1=339.2  (page 1)
  Title       : y_top≈86.1 (centred), GillSansMT-Bold 16pt
  Patient tbl : y_top≈127, rows every ~29pt, left_x=39.6 right_x=373.7
  Test Ind hdg: y_top≈280
  Result hdg  : y_top≈337-380
  ISCN box    : y_top≈377-420  (amber bg, bold orange text)
  Karyogram   : single centred or dual side-by-side
  Metaphase   : page-1 bottom for 2-page layout; page-2 top for 3-page layout
  Signatures  : page-2 for normal, page-3 for abnormal

Layout decision:
  - comments field is blank → 2-page (normal) layout
  - comments field is non-blank → 3-page (abnormal) layout
"""

import os, io, re, base64, sys
from datetime import datetime
from pathlib import Path


def _resource_path(relative: str) -> str:
    """Works frozen (PyInstaller) and unfrozen."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)


from reportlab.pdfgen           import canvas
from reportlab.lib.colors       import Color, black, white, HexColor
from reportlab.lib.utils        import ImageReader
from reportlab.lib.styles       import ParagraphStyle
from reportlab.lib.enums        import TA_JUSTIFY, TA_LEFT, TA_CENTER
from reportlab.platypus         import Paragraph
from reportlab.pdfbase          import pdfmetrics
from reportlab.pdfbase.ttfonts  import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily

import karyotype_assets as _assets

# ─── Colours ──────────────────────────────────────────────────────────────────
# Extracted from reference PDF (PARTHIBAN): (0.122, 0.22, 0.392) = #1F3864
DARK_BLUE  = HexColor('#1F3864')   # section headings, title, "This report..." header
RED        = Color(1, 0, 0)        # ISCN box text — abnormal
GREEN      = HexColor('#00B050')   # ISCN box text — normal (both autosome+sex normal)
AMBER_BG   = HexColor('#F2F2F2')   # ISCN box background  (light gray)
AMBER_BRD  = HexColor('#D9D9D9')   # ISCN box border
GRAY_DIV   = Color(0.6, 0.6, 0.6) # section divider lines
FIELD_BG   = HexColor('#D9D9D9')   # metaphase table background
BLACK      = black
WHITE      = white

# ─── Fonts ────────────────────────────────────────────────────────────────────
_FONT_DIR = _resource_path("fonts")

def _reg(name, filename):
    path = os.path.join(_FONT_DIR, filename)
    if os.path.exists(path):
        try:
            pdfmetrics.registerFont(TTFont(name, path))
            return True
        except Exception:
            pass
    return False

_reg("GillSansMT-Bold",       "GillSansMT-Bold.ttf")
_reg("GillSansMT",            "GillSansMT.ttf")
_reg("GillSansMT-BoldItalic", "GillSansMT-BoldItalic.ttf")
_reg("SegoeUI-Bold",          "SegoeUI-Bold.ttf")
_reg("SegoeUI",               "SegoeUI.ttf")
_reg("SegoeUI-Italic",        "SegoeUI-Italic.ttf")
_reg("Calibri",               "Calibri.ttf")
_reg("Calibri-Bold",          "Calibri-Bold.ttf")
_reg("Calibri-Italic",        "Calibri-Italic.ttf")
_reg("Calibri-BoldItalic",    "Calibri-BoldItalic.ttf")
_reg("Arial",                 "ArialMT.ttf")
_reg("Arial-Bold",            "Arial-BoldMT.ttf")

def _font_ok(n):
    try: pdfmetrics.getFont(n); return True
    except: return False

if _font_ok("Calibri") and _font_ok("Calibri-Bold"):
    registerFontFamily("Calibri", normal="Calibri", bold="Calibri-Bold",
                       italic="Calibri-Italic" if _font_ok("Calibri-Italic") else "Calibri",
                       boldItalic="Calibri-BoldItalic" if _font_ok("Calibri-BoldItalic") else "Calibri-Bold")
if _font_ok("Arial") and _font_ok("Arial-Bold"):
    registerFontFamily("Arial", normal="Arial", bold="Arial-Bold",
                       italic="Arial", boldItalic="Arial-Bold")
if _font_ok("SegoeUI") and _font_ok("SegoeUI-Bold"):
    registerFontFamily("SegoeUI", normal="SegoeUI", bold="SegoeUI-Bold",
                       italic="SegoeUI-Italic" if _font_ok("SegoeUI-Italic") else "SegoeUI",
                       boldItalic="SegoeUI-Bold")
if _font_ok("GillSansMT") and _font_ok("GillSansMT-Bold"):
    registerFontFamily("GillSansMT", normal="GillSansMT", bold="GillSansMT-Bold",
                       italic="GillSansMT", boldItalic="GillSansMT-BoldItalic" if _font_ok("GillSansMT-BoldItalic") else "GillSansMT-Bold")

# Font aliases matching reference PDF fonts exactly
# Reference: GillSansMT-Bold for title/headings; SegoeUI-Bold/SegoeUI for patient table;
#            Calibri/Calibri-Bold for body/ISCN; Calibri-Italic for references italic
F_TITLE  = "GillSansMT-Bold" if _font_ok("GillSansMT-Bold") else "Helvetica-Bold"
F_HDG    = "GillSansMT-Bold" if _font_ok("GillSansMT-Bold") else "Helvetica-Bold"
F_TBL_LBL = "SegoeUI-Bold"  if _font_ok("SegoeUI-Bold")    else "Helvetica-Bold"   # patient table labels
F_TBL_VAL = "SegoeUI"       if _font_ok("SegoeUI")         else "Helvetica"         # patient table values
F_LBL    = "Calibri-Bold"   if _font_ok("Calibri-Bold")    else "Helvetica-Bold"
F_BODY   = "Calibri"        if _font_ok("Calibri")         else "Helvetica"
F_ITALIC = "Calibri-Italic" if _font_ok("Calibri-Italic")  else "Helvetica"
F_BBOLD  = "Calibri-Bold"   if _font_ok("Calibri-Bold")    else "Helvetica-Bold"
F_SIG    = "Calibri"        if _font_ok("Calibri")         else "Helvetica"

# ─── Page geometry (all in ReportLab coords: origin = bottom-left) ─────────
W, H = 612.0, 792.0

# pdfplumber top → RL y = H - pdfplumber_top
def _rl(pdfplumber_top):
    return H - pdfplumber_top

HDR_X, HDR_Y, HDR_W, HDR_H   =  1.4,  _rl(67.8),   609.8, 67.8
FTR_X, FTR_Y, FTR_W, FTR_H   =  1.4,  0.2,         610.6, 48.0
STAMP_X, STAMP_Y, STAMP_W, STAMP_H = 276.4, _rl(216.0), 62.8, 78.8

# Content margins
LX = 39.6    # left content x
RX = 575.0   # right content x
CW = RX - LX # content width  535.4

# Patient table column boundaries (from pdfplumber analysis)
LEFT_VAL_X  = 136.8  # left-side values start here
RIGHT_LBL_X = 373.7  # right-side labels start here
RIGHT_VAL_X = 487.3  # right-side values start here
COLON_X_L   = 130.2  # left colons
COLON_X_R   = 481.8  # right colons

TABLE_ROW_H = 29.2   # approx row height (based on 127.2 → 161.7 = 34.5, etc)

# Divider line x-span
DIV_X0, DIV_X1 = 72.0, 540.0

# ─── Helpers ──────────────────────────────────────────────────────────────────
def _img(b64: str) -> ImageReader:
    return ImageReader(io.BytesIO(base64.b64decode(b64)))

def _divider(c, rl_y, lw=0.48):
    c.setStrokeColor(GRAY_DIV)
    c.setLineWidth(lw)
    c.line(DIV_X0, rl_y, DIV_X1, rl_y)

def _clean(v) -> str:
    s = str(v).strip()
    return "" if s in ("nan", "NaT", "None", "NaN", "") else s

def _fmt_date(v) -> str:
    if not v: return ""
    s = _clean(v)
    if not s: return ""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s.split(" ")[0], fmt.split(" ")[0]).strftime("%d-%m-%Y")
        except Exception:
            pass
    return s

def _wrap_text(c, text, x, y, max_w, font, size, leading=None) -> float:
    """Word-wrap text; returns RL y after the last drawn line."""
    if leading is None:
        leading = size * 1.45
    c.setFont(font, size)   # always set the correct font before drawing
    words = text.split()
    line  = ""
    for w in words:
        trial = (line + " " + w).strip()
        if c.stringWidth(trial, font, size) <= max_w:
            line = trial
        else:
            if line:
                c.drawString(x, y, line)
                y -= leading
            line = w
    if line:
        c.drawString(x, y, line)
        y -= leading
    return y

def _paragraph_height(text, font, size, max_w, leading=None) -> float:
    """Estimate height of word-wrapped text block (in pt)."""
    if not text: return 0.0
    if leading is None:
        leading = size * 1.45
    # use a dummy canvas approach: count line breaks
    try:
        from reportlab.pdfgen.canvas import Canvas as _C
        buf = io.BytesIO()
        dummy = _C(buf, pagesize=(W, H))
        words = text.split()
        line, lines = "", 0
        for w in words:
            trial = (line + " " + w).strip()
            if dummy.stringWidth(trial, font, size) <= max_w:
                line = trial
            else:
                lines += 1
                line = w
        if line: lines += 1
        return lines * leading
    except Exception:
        # rough fallback
        avg_char_w = size * 0.5
        chars_per_line = int(max_w / avg_char_w)
        n_lines = max(1, len(text) // chars_per_line + 1)
        return n_lines * (leading or size * 1.35)

def _draw_section_heading(c, text, rl_y, color=DARK_BLUE, size=16) -> float:
    """Draw a bold section heading and a divider below it. Returns RL y below divider.

    The text is drawn at rl_y - cap_offset so that pdfplumber-measured glyph top
    aligns with the reference PDF.  The return value (used for body placement)
    is kept at rl_y - size - 3 - 4 so downstream body positions are unaffected.
    """
    # cap_offset ≈ size × 0.73 (Calibri cap height fraction); empirically ~11.8 for 16pt
    cap_offset = size * 0.74
    c.setFont(F_HDG, size)
    c.setFillColor(color)
    c.drawString(DIV_X0, rl_y - cap_offset, text)
    div_y = rl_y - size - 3
    _divider(c, div_y)
    return div_y - 4

def _draw_bullet_list(c, items, x, y, max_w, font, size, leading=None) -> float:
    """Draw a bullet list; returns RL y after last item."""
    if leading is None:
        leading = size * 1.45
    bullet = "\u2022"
    indent = 18.0
    for item in items:
        # Use Helvetica for the bullet glyph (guaranteed to have \u2022)
        c.setFont("Helvetica", size)
        c.setFillColor(BLACK)
        c.drawString(x, y, bullet)
        y = _wrap_text(c, item, x + indent, y, max_w - indent, font, size, leading)
    return y


def _draw_justified(c, text: str, x: float, y: float, max_w: float,
                    font: str, size: float, leading: float = None) -> float:
    """Draw justified paragraph text. Returns RL y below the last line."""
    if not text:
        return y
    if leading is None:
        leading = size * 1.41  # matches reference PDF ~15.5pt leading for 11pt body text
    style = ParagraphStyle(
        'body_just',
        fontName=font,
        fontSize=size,
        leading=leading,
        alignment=TA_JUSTIFY,
        wordWrap='LTR',
    )
    para = Paragraph(text, style)
    w, h = para.wrap(max_w, 9999)
    para.drawOn(c, x, y - h)
    return y - h


# ─── Main report generator ───────────────────────────────────────────────────
class KaryotypeReportGenerator:
    """Generate a single patient's Karyotyping PDF.

    Parameters
    ----------
    data_row : dict
        Keys match Excel columns (case-insensitive normalised).
    image_paths : list[str]
        Paths to karyogram image files for this patient (1, 2, or 3 images).
        - 1 image  → centred single karyogram
        - 2 images → side-by-side dual karyogram (mosaic case)
        - 3 images → scatter + zoom-pair on left, full karyogram on right
    output_dir : str
        Directory where the PDF will be saved.
    """

    def __init__(self, data_row: dict, image_paths: list, output_dir: str,
                 include_logo: bool = True):
        self.d      = {k.strip().upper(): _clean(v) for k, v in data_row.items()}
        self.images = [p for p in image_paths if p and os.path.isfile(p)]
        self.out    = output_dir
        self.include_logo = include_logo

        # Derive filename: "{Name} ({SampleNumber}) PBCKT with/without logo.pdf"
        name   = " ".join((self._get("NAME") or "Unknown").split())
        sample = " ".join((self._get("SAMPLE NUMBER") or "NoSN").split())
        # Remove characters not safe for filenames (keep alphanumeric, spaces, hyphens, parens, dots)
        safe_name = re.sub(r'[^\w\s\-\(\)\.]', '', name).strip()
        logo_tag = "with logo" if include_logo else "without logo"
        self.filename = f"{safe_name} ({sample}) PBCKT {logo_tag}.pdf"
        self.filepath = os.path.join(output_dir, self.filename)

        # Determine layout variant
        has_comments = bool(self._get("COMMENTS"))
        self.three_page = has_comments   # True → 3-page layout

    def _is_normal(self) -> bool:
        """Return True if both autosome and sex chromosome are normal."""
        auto = self._get("AUTOSOME").lower()
        sex  = self._get("SEX CHROMOSOME", "SEX CHROMOSOME ").lower()
        return not ("abnormal" in auto or "abnormal" in sex or
                    "variant" in auto or "variant" in sex)

    def _iscn_color(self):
        """Return ISCN text color based on autosome/sex chromosome values.
        Both Normal → GREEN; either Abnormal → RED; either Variant → BLACK."""
        auto = self._get("AUTOSOME").lower()
        sex  = self._get("SEX CHROMOSOME", "SEX CHROMOSOME ").lower()
        if "abnormal" in auto or "abnormal" in sex:
            return RED
        if "variant" in auto or "variant" in sex:
            return BLACK
        return GREEN  # both normal

    # ── accessors ─────────────────────────────────────────────────────────────
    def _get(self, *keys) -> str:
        for k in keys:
            v = self.d.get(k.upper().strip(), "")
            if v: return v
        return ""

    # ── public entry point ────────────────────────────────────────────────────
    def generate(self) -> str:
        os.makedirs(self.out, exist_ok=True)
        c = canvas.Canvas(self.filepath, pagesize=(W, H))
        c.setTitle(f"Karyotype Report - {self._get('NAME')}")

        if self.three_page:
            self._page1(c)
            c.showPage()
            self._page2_abnormal(c)
            c.showPage()
            self._page3_signatures(c)
        else:
            self._page1_with_metaphase(c)
            c.showPage()
            self._page2_normal(c)

        c.save()
        return self.filepath

    # ══════════════════════════════════════════════════════════════════════════
    # PAGE 1 (both layouts share patient info + result + karyograms)
    # ══════════════════════════════════════════════════════════════════════════
    def _page1_common(self, c, page_num=1, total_pages=None) -> float:
        """Draw everything on page 1 up to (but not including) the karyogram area.
        Returns the RL y just above where karyograms should be placed."""
        if total_pages is None:
            total_pages = 3 if self.three_page else 2

        self._draw_chrome(c, page_num, total_pages)

        # ── Title ────────────────────────────────────────────────────────────
        # Draw title so its glyph TOP aligns with pdfplumber top=86.1
        # baseline = glyph_top_RL - cap_height = _rl(86.1) - 18*0.74
        title_y = _rl(86.1) - int(18 * 0.74)   # RL baseline ≈ 692.6
        c.setFont(F_TITLE, 18)
        c.setFillColor(DARK_BLUE)
        c.drawCentredString(W / 2, title_y, "Peripheral Blood Karyotyping")

        # ── Stamp image (drawn FIRST so table text renders on top) ───────────
        c.drawImage(_img(_assets.STAMP),
                    STAMP_X, STAMP_Y, STAMP_W, STAMP_H, mask="auto")

        # ── Patient info table ────────────────────────────────────────────────
        self._draw_patient_table(c)

        # ── Test Indication ───────────────────────────────────────────────────
        # Position just below the patient-table bottom border (RL ≈ 520.5)
        # leaving ~20 pt of breathing room.
        ti_y = _rl(290.0)   # pdfplumber ≈ 290 → RL 502; 20 pt below table bottom
        section_y = _draw_section_heading(c, "Test Indication", ti_y)
        c.setFont(F_BODY, 11)
        c.setFillColor(BLACK)
        indication = self._get("TEST INDICATION") or "To rule out gross chromosomal abnormality"
        section_y = _draw_justified(c, indication, DIV_X0, section_y - 12, DIV_X1 - DIV_X0,
                                    F_BODY, 11)

        # ── Result ────────────────────────────────────────────────────────────
        res_y = section_y - 39
        _divider(c, res_y + 12, lw=0.5)
        res_y = _draw_section_heading(c, "Result", res_y)

        # ISCN result box — tries to fit on single line by reducing font;
        # falls back to word-wrap only if text won't fit even at min font size.
        iscn = self._get("RESULT")
        prefix = "International System for Human Cytogenomic Nomenclature (ISCN 2024):"
        full_text = prefix + "  " + iscn
        box_x0, box_x1 = DIV_X0, DIV_X1 + 7
        pad   = 8   # horizontal padding inside box
        pad_v = 7   # vertical padding top+bottom
        avail_w = box_x1 - box_x0 - pad * 2

        # Try decreasing font sizes to keep text on one line
        font_sz = 12
        for try_sz in [12, 11, 10, 9, 8]:
            c.setFont(F_LBL, try_sz)
            if c.stringWidth(full_text, F_LBL, try_sz) <= avail_w:
                font_sz = try_sz
                lines   = [full_text]
                break
        else:
            # Still doesn't fit on one line — word-wrap at 10pt
            font_sz = 10
            c.setFont(F_LBL, font_sz)
            words = full_text.split()
            lines, cur = [], ""
            for w in words:
                trial = (cur + " " + w).strip()
                if c.stringWidth(trial, F_LBL, font_sz) <= avail_w:
                    cur = trial
                else:
                    if cur: lines.append(cur)
                    cur = w
            if cur: lines.append(cur)

        line_h  = font_sz * 1.3
        n_lines = len(lines)
        box_h   = n_lines * line_h + pad_v * 2
        box_y   = res_y - 8
        box_bot = box_y - box_h

        c.setFillColor(AMBER_BG)
        c.setStrokeColor(AMBER_BRD)
        c.setLineWidth(0.6)
        c.rect(box_x0, box_bot, box_x1 - box_x0, box_h, fill=1, stroke=1)

        iscn_color = self._iscn_color()
        c.setFont(F_LBL, font_sz)
        c.setFillColor(iscn_color)
        box_cx = (box_x0 + box_x1) / 2
        # Draw lines top-to-bottom, vertically centred
        text_block_h = n_lines * line_h
        first_y = box_bot + (box_h + text_block_h) / 2 - line_h + (line_h - font_sz * 0.74) / 2
        for idx, line in enumerate(lines):
            c.drawCentredString(box_cx, first_y - idx * line_h, line)

        karyogram_top_y = box_bot - 10  # RL y at top of karyogram area
        return karyogram_top_y

    def _page1_with_metaphase(self, c):
        """2-page layout: karyogram + metaphase table on page 1."""
        top_y  = self._page1_common(c, page_num=1, total_pages=2)
        meta_h = 19.5 * 2            # 39 pt  (2 rows × row_h)

        # Keep the metaphase table well above the footer + page-number text.
        # FTR_Y + FTR_H = 48.2 pt from bottom; add ~22 pt clearance.
        meta_bot  = FTR_Y + FTR_H + 22   # RL y of table bottom ≈ 70 pt
        kgram_bot = meta_bot + meta_h + 8  # karyogram bottom is above the table

        self._draw_karyograms(c, top_y, kgram_bot)
        self._draw_metaphase_table(c, meta_bot)

    def _page1(self, c):
        """3-page layout: karyogram only on page 1 (no metaphase table)."""
        top_y    = self._page1_common(c, page_num=1, total_pages=3)
        # Leave clearance above footer + page-number line
        bottom_y = FTR_Y + FTR_H + 22   # ≈ 70 pt from bottom
        self._draw_karyograms(c, top_y, bottom_y)

    # ══════════════════════════════════════════════════════════════════════════
    # PAGE 2 — NORMAL (2-page layout)
    # ══════════════════════════════════════════════════════════════════════════
    def _page2_normal(self, c):
        self._draw_chrome(c, 2, 2)
        y = _rl(67.8) - 52  # ~52 pt below header bottom (pdfplumber: heading top=119.6)

        y = _draw_section_heading(c, "Interpretation", y)
        c.setFillColor(BLACK)
        y = _draw_justified(c, self._get("INTERPRETATION"), DIV_X0, y - 10,
                            DIV_X1 - DIV_X0, F_BODY, 11)

        y = self._draw_recommendations_block(c, y - 12)
        y = self._draw_methodology_block(c, y - 12)
        y = self._draw_limitations_block(c, y - 12)
        y = self._draw_references_block(c, y - 8)
        self._draw_signatures(c, y - 11)

    # ══════════════════════════════════════════════════════════════════════════
    # PAGE 2 — ABNORMAL (3-page layout)
    # ══════════════════════════════════════════════════════════════════════════
    def _page2_abnormal(self, c):
        self._draw_chrome(c, 2, 3)
        hdr_bot = _rl(67.8)   # RL y of header bottom = 724.2

        # Metaphase table sits flush below the header.
        # _draw_metaphase_table(rl_y) draws UPWARD from rl_y by row_h*2.
        meta_h   = 19.5 * 2   # 39 pt
        meta_bot = hdr_bot - meta_h - 2   # bottom of table in RL
        self._draw_metaphase_table(c, meta_bot)
        y = meta_bot - 24

        y = _draw_section_heading(c, "Interpretation", y)
        c.setFillColor(BLACK)
        y = _draw_justified(c, self._get("INTERPRETATION"), DIV_X0, y - 10,
                            DIV_X1 - DIV_X0, F_BODY, 11)

        # Comments section (hidden if empty, but rendered if present)
        comments = self._get("COMMENTS")
        if comments:
            y = y - 12
            y = _draw_section_heading(c, "Comments", y)
            c.setFillColor(BLACK)
            y = _draw_justified(c, comments, DIV_X0, y - 10,
                                DIV_X1 - DIV_X0, F_BODY, 11)

        y = self._draw_recommendations_block(c, y - 12)
        y = self._draw_methodology_block(c, y - 12)
        self._draw_limitations_block(c, y - 12)
        # References moved to page 3 with signatures (spacing adjusted there)

    # ══════════════════════════════════════════════════════════════════════════
    # PAGE 3 — SIGNATURES (3-page layout)
    # ══════════════════════════════════════════════════════════════════════════
    def _page3_signatures(self, c):
        """Page 3: References then signatures."""
        self._draw_chrome(c, 3, 3)
        y = _rl(67.8) - 52   # start just below header
        y = self._draw_references_block(c, y)
        self._draw_signatures(c, y - 20)

    # ══════════════════════════════════════════════════════════════════════════
    # COMMON DRAWING SUBROUTINES
    # ══════════════════════════════════════════════════════════════════════════
    def _draw_chrome(self, c, page_num: int, total_pages: int):
        """Header image, footer image, and page number."""
        if self.include_logo:
            c.drawImage(_img(_assets.HEADER), HDR_X, HDR_Y, HDR_W, HDR_H, mask="auto")
            c.drawImage(_img(_assets.FOOTER), FTR_X, FTR_Y, FTR_W, FTR_H, mask="auto")

        # Page number — only shown when logo/header is included
        if self.include_logo:
            c.setFont(F_BODY, 8)
            c.setFillColor(BLACK)
            pg_str = f"{page_num}  |  P a g e"
            c.drawRightString(DIV_X1, FTR_Y + FTR_H + 4, pg_str)

    def _draw_patient_table(self, c):
        """Draw the patient info table rows."""
        rows = [
            ("Patient name",      self._get("NAME"),
             "PIN",               self._get("PIN")),
            ("Gender/ Age",
             self._get("GENDER", "GENDER ") + " / " + self._get("AGE") + " Year" +
             ("s" if str(self._get("AGE")) != "1" else ""),
             "Sample Number",     self._get("SAMPLE NUMBER")),
            ("Specimen",          self._get("SPECIMEN"),
             "Sample collection date",
             _fmt_date(self._get("SAMPLE COLLECTION DATE", "SAMPLE COLLECTION DATE "))),
            ("Referring Clinician",
             self._get("REFERRING CLINICIAN"),
             "Sample receipt date",
             _fmt_date(self._get("SAMPLE RECEIPT DATE", "SAMPLE RECEIPT DATE "))),
            ("Hospital/Clinic",   self._get("HOSPITAL/CLINIC"),
             "Report Date",
             datetime.today().strftime("%d-%m-%Y")),
        ]

        # Starting y in RL (pdfplumber top=127.2 → RL y = 664.8, but text baseline)
        row_top = _rl(120.0)   # RL y of top padding before first row
        row_h   = 29.5
        label_size = 9
        value_size = 10

        for i, (ll, lv, rl, rv) in enumerate(rows):
            baseline = row_top - (i * row_h) - row_h * 0.55

            # Left label (bold)
            c.setFont(F_TBL_LBL, label_size)
            c.setFillColor(BLACK)
            c.drawString(LX, baseline, ll)
            c.drawString(COLON_X_L, baseline, ":")

            # Left value (bold, title-cased since Excel stores all-caps)
            c.setFont(F_TBL_LBL, value_size)
            _wrap_text(c, lv.title(), LEFT_VAL_X, baseline,
                       RIGHT_LBL_X - LEFT_VAL_X - 8, F_TBL_LBL, value_size, leading=11)

            # Right label (bold)
            c.setFont(F_TBL_LBL, label_size)
            c.setFillColor(BLACK)
            c.drawString(RIGHT_LBL_X, baseline, rl)
            c.drawString(COLON_X_R - 2, baseline, ":")

            # Right value (bold)
            c.setFont(F_TBL_LBL, value_size)
            _wrap_text(c, rv, RIGHT_VAL_X, baseline,
                       RX - RIGHT_VAL_X, F_TBL_LBL, value_size, leading=11)

        pass  # no lines around patient table

    def _draw_karyograms(self, c, top_y: float, bottom_y: float):
        """Place 1, 2, or 3 karyogram images in the available vertical band."""
        if not self.images:
            return

        avail_h = top_y - bottom_y
        avail_w = DIV_X1 - DIV_X0

        n = len(self.images)

        if n == 1:
            self._place_image(c, self.images[0],
                              DIV_X0, bottom_y, avail_w, avail_h)

        elif n == 2:
            # Side-by-side mosaic: force both images to render at equal size
            gap   = 8
            img_w = (avail_w - gap) / 2

            # Read both image dimensions to compute a common scale factor
            def _get_dims(path):
                try:
                    from PIL import Image as PILImage
                    with PILImage.open(path) as im:
                        return im.size
                except Exception:
                    return (img_w, avail_h)

            iw0, ih0 = _get_dims(self.images[0])
            iw1, ih1 = _get_dims(self.images[1])
            # Use the same scale (minimum of both) so rendered sizes are equal
            common_scale = min(img_w / iw0, avail_h / ih0,
                               img_w / iw1, avail_h / ih1)
            self._place_image(c, self.images[0],
                              DIV_X0, bottom_y, img_w, avail_h, fixed_scale=common_scale)
            self._place_image(c, self.images[1],
                              DIV_X0 + img_w + gap, bottom_y, img_w, avail_h, fixed_scale=common_scale)

        else:
            # 3 images: scatter (top-left) + zoom pair (bottom-left) + karyogram (right)
            # Proportions from reference PDF pdfplumber analysis
            left_w  = avail_w * 0.37
            right_w = avail_w * 0.59
            gap     = avail_w - left_w - right_w   # ≈ 4% gap

            # Left column: scatter at top, zoom directly below
            scatter_h = avail_h * 0.54
            zoom_h    = avail_h * 0.44

            # Compute scatter's actual rendered height so zoom sits just below it (not below slot)
            try:
                from PIL import Image as PILImage
                with PILImage.open(self.images[0]) as _im:
                    _s_iw, _s_ih = _im.size
                _scatter_scale = min(left_w / _s_iw, scatter_h / _s_ih)
                scatter_dh = _s_ih * _scatter_scale
            except Exception:
                scatter_dh = scatter_h

            scatter_slot_y = bottom_y + avail_h - scatter_h   # bottom of scatter slot (top-aligned)
            scatter_rendered_bottom = scatter_slot_y + scatter_h - scatter_dh  # actual bottom edge
            IMG_GAP = 4
            zoom_slot_y = scatter_rendered_bottom - IMG_GAP - zoom_h

            # All three images have built-in frames — no extra border on any of them
            self._place_image(c, self.images[0],
                              DIV_X0, scatter_slot_y, left_w, scatter_h,
                              valign='top')
            self._place_image(c, self.images[1],
                              DIV_X0, zoom_slot_y, left_w, zoom_h,
                              valign='top')

            # Right column: karyogram has built-in frame — no extra border to avoid double
            # top-aligned so its top edge matches the scatter top edge
            self._place_image(c, self.images[2],
                              DIV_X0 + left_w + gap, bottom_y, right_w, avail_h,
                              valign='top')

    @staticmethod
    def _image_has_border(path: str) -> bool:
        """Return True if the image already has a built-in dark frame.
        Scans insets 0-15px from each edge and takes the minimum average
        brightness found. Built-in borders appear at different inset depths
        across images, so scanning a range catches them all reliably.
        Threshold 230: no-border images stay at 255; framed images drop to ~130-193."""
        try:
            from PIL import Image as PILImage
            with PILImage.open(path) as im:
                rgb = im.convert('RGB')
                w, h = rgb.size
            step_x = max(1, w // 20)
            step_y = max(1, h // 20)
            min_avg = 255.0
            for inset in range(16):
                i = inset
                samples = []
                for px in range(i, w - i, step_x):
                    samples.append(rgb.getpixel((px, i)))
                    samples.append(rgb.getpixel((px, h - 1 - i)))
                for py in range(i, h - i, step_y):
                    samples.append(rgb.getpixel((i, py)))
                    samples.append(rgb.getpixel((w - 1 - i, py)))
                avg = sum(sum(px) / 3 for px in samples) / len(samples)
                if avg < min_avg:
                    min_avg = avg
            return min_avg < 230
        except Exception:
            return False

    def _place_image(self, c, path: str, x: float, y: float, max_w: float, max_h: float,
                     fixed_scale: float = None, valign: str = 'center'):
        """Draw image scaled to fit max_w × max_h, centred horizontally.
        valign : 'center' (default) | 'top' | 'bottom'
        A 1pt border is drawn automatically only when the image has no built-in frame
        (detected by sampling edge-pixel brightness)."""
        try:
            from PIL import Image as PILImage
            with PILImage.open(path) as im:
                iw, ih = im.size
        except Exception:
            iw, ih = max_w, max_h

        if fixed_scale is not None:
            scale = fixed_scale
        else:
            scale = min(max_w / iw, max_h / ih)

        dw, dh = iw * scale, ih * scale
        # Horizontal: always centre
        cx = x + (max_w - dw) / 2
        # Vertical alignment (RL y increases upward)
        if valign == 'top':
            cy = y + max_h - dh   # flush to top of slot
        elif valign == 'bottom':
            cy = y                 # flush to bottom of slot
        else:
            cy = y + (max_h - dh) / 2  # centred

        try:
            c.drawImage(path, cx, cy, dw, dh, mask="auto")
        except Exception:
            pass

        # Draw a thin border only if the image doesn't already have its own frame
        if not self._image_has_border(path):
            c.setStrokeColor(BLACK)
            c.setLineWidth(0.5)
            c.rect(cx, cy, dw, dh, fill=0, stroke=1)

    def _draw_metaphase_table(self, c, rl_y: float) -> float:
        """Draw the 4-cell metaphase/autosome status table. Returns RL y below table."""
        # From pdfplumber: row1 y=72.6, row2 y=91.9 (from top on page 2 in 3-page layout)
        # On page 1 in 2-page layout: y=675.3 row1, y=694.7 row2 → rl_y = 116.7 and 97.3
        # Two rows, 4 cells
        met_val  = self._get("METAPHASE ANALYSED", "METAPHASE ANALYSED ")
        auto_val = self._get("AUTOSOME")
        band_val = self._get("ESTIMATED BAND RESOLUTION")
        sex_val  = self._get("SEX CHROMOSOME", "SEX CHROMOSOME ")

        row_h  = 19.5
        tbl_x0 = LX + 12
        tbl_x1 = RX - 12
        mid_x  = (tbl_x0 + tbl_x1) / 2
        col1_x = tbl_x0 + 8    # left area label x
        col2_x = col1_x + 175  # left area value x
        col3_x = mid_x + 18    # right area label x — extra offset gives breathing room from centre
        col4_x = col3_x + 175  # right area value x

        # Background
        c.setFillColor(FIELD_BG)
        c.rect(tbl_x0, rl_y, tbl_x1 - tbl_x0, row_h * 2, fill=1, stroke=0)

        # No outer border or internal dividers on the metaphase table

        # Vertically centre text: row_h/2 - cap_height/2 from row top
        cap_h9 = 9 * 0.74 / 2
        label_y1 = rl_y + row_h * 2 - row_h / 2 - cap_h9  # row 1 text baseline (centred)
        label_y2 = rl_y + row_h       - row_h / 2 - cap_h9  # row 2 text baseline (centred)

        # Pre-compute aligned colon x for each half based on longest label
        c.setFont(F_BODY, 9)
        left_colon_x  = col1_x + max(c.stringWidth("Metaphase analysed",        F_BODY, 9),
                                     c.stringWidth("Estimated band resolution",  F_BODY, 9)) + 8
        right_colon_x = col3_x + max(c.stringWidth("Autosome",      F_BODY, 9),
                                     c.stringWidth("Sex chromosome", F_BODY, 9)) + 8
        left_val_x    = left_colon_x  + c.stringWidth(":", F_BODY, 9) + 10
        right_val_x   = right_colon_x + c.stringWidth(":", F_BODY, 9) + 10

        def _cell(lbl, val, lx, colon_x, val_x, ly):
            c.setFont(F_BODY, 9); c.setFillColor(BLACK)
            c.drawString(lx, ly, lbl)
            c.drawString(colon_x, ly, ":")
            c.drawString(val_x, ly, val)

        _cell("Metaphase analysed",       met_val,  col1_x, left_colon_x,  left_val_x,  label_y1)
        _cell("Autosome",                 auto_val, col3_x, right_colon_x, right_val_x, label_y1)
        _cell("Estimated band resolution",band_val, col1_x, left_colon_x,  left_val_x,  label_y2)
        _cell("Sex chromosome",           sex_val,  col3_x, right_colon_x, right_val_x, label_y2)

        return rl_y  # caller can do rl_y - something to continue below

    def _draw_recommendations_block(self, c, y: float) -> float:
        y = _draw_section_heading(c, "Recommendations", y)
        c.setFillColor(BLACK)
        if self._is_normal():
            items = [
                "Genetic counseling is recommended to discuss the implications of the result.",
                "Additional genetic testing may be warranted based on the specific phenotypic indication.",
            ]
            y = _draw_bullet_list(c, items, DIV_X0, y - 18,
                                  DIV_X1 - DIV_X0, F_BODY, 11)
        else:
            recs = self._get("RECOMMENDATIONS")
            # Split multiple recommendations on bullet char or newline
            raw_items = [r.strip() for r in re.split(r'[\n\uf0b7\u2022]', recs) if r.strip()]
            # Merge continuation fragments: if an item starts with lowercase it's a
            # continuation of the previous item (e.g. line-break mid-sentence in Excel)
            items = []
            for item in raw_items:
                if items and item and item[0].islower():
                    items[-1] = items[-1] + " " + item
                else:
                    items.append(item)
            if len(items) > 1:
                y = _draw_bullet_list(c, items, DIV_X0, y - 18,
                                      DIV_X1 - DIV_X0, F_BODY, 11)
            elif items:
                y = _draw_justified(c, items[0], DIV_X0, y - 10,
                                    DIV_X1 - DIV_X0, F_BODY, 11)
            else:
                y = _draw_justified(c, recs, DIV_X0, y - 10,
                                    DIV_X1 - DIV_X0, F_BODY, 11)
        return y

    def _draw_methodology_block(self, c, y: float) -> float:
        y = _draw_section_heading(c, "Test Methodology", y)
        text = ("A 72-hour PHA-M stimulated culture of the received peripheral blood sample was "
                "processed, according to a protocol adapted from the AGT Cytogenetics Laboratory "
                "Manual, Third Edition. Numerical and structural chromosomal abnormalities were "
                "ruled out at a banding resolution suitable for the referral indication, in "
                "accordance with current ISCN guidelines.")
        c.setFillColor(BLACK)
        y = _draw_justified(c, text, DIV_X0, y - 10, DIV_X1 - DIV_X0, F_BODY, 11)
        return y

    def _draw_limitations_block(self, c, y: float) -> float:
        y = _draw_section_heading(c, "Limitations", y)
        items = [
            "All genetic disorders cannot be ruled out by conventional karyotyping.",
            "The accuracy of the test is about 99%.",
            "G-banded analysis cannot detect small rearrangements and submicroscopic deletions.",
            "Low-level mosaicism may not be detected.",
        ]
        c.setFillColor(BLACK)
        y = _draw_bullet_list(c, items, DIV_X0, y - 18, DIV_X1 - DIV_X0, F_BODY, 11)
        return y

    def _draw_references_block(self, c, y: float) -> float:
        y = _draw_section_heading(c, "References", y)
        # Each ref: (number, normal_part, italic_part)
        refs = [
            ("1.", "The AGT Cytogenetics Laboratory Manual Third Edition (1997)  ",
             "Editors: Margaret J. Barch, TuridKnutsen, Jack L. Spurbeck."),
            ("2.", "ISCN- An International System for Human Cytogenetic Nomenclature (2024)  ",
             "Editors: Ros J. Hastings, Sarah Moore, Nicole Chia."),
        ]
        c.setFillColor(BLACK)
        indent = 22.0
        ref_w = DIV_X1 - DIV_X0 - indent
        # Use Paragraph for justified text with inline italic via XML markup
        ref_italic = F_ITALIC  # Calibri-Italic (matches reference PDF)
        # First ref needs smaller gap from heading (~14pt); subsequent refs smaller gap (~10pt)
        first_gap = 14
        inter_gap = 10
        for i, (num, normal_part, italic_part) in enumerate(refs):
            y -= (first_gap if i == 0 else inter_gap)
            c.setFont(F_BBOLD, 11)
            c.setFillColor(BLACK)
            c.drawString(DIV_X0, y, num)
            # Build markup: normal text + italic editors
            markup = (normal_part +
                      f'<font name="{ref_italic}" color="black"><i>{italic_part}</i></font>')
            style = ParagraphStyle(
                'ref_just',
                fontName=F_BODY,
                fontSize=11,
                leading=15.5,
                alignment=TA_JUSTIFY,
                wordWrap='LTR',
            )
            para = Paragraph(markup, style)
            pw, ph = para.wrap(ref_w, 9999)
            para.drawOn(c, DIV_X0 + indent, y - ph + 10)
            y = y - ph + 10
        return y

    def _draw_signatures(self, c, y: float):
        """Draw 'This report has been reviewed and approved by:' + signature image."""
        # Heading
        c.setFont(F_HDG, 14)
        c.setFillColor(DARK_BLUE)
        c.drawString(DIV_X0, y - 14, "This report has been reviewed and approved by:")

        # Combined signature image
        sig_y  = y - 14 - 8    # top of sig image in RL (drawImage draws from bottom-left)
        sig_w  = 538.8 - 97.4  # 441.4 pt (from pdfplumber)
        sig_h  = 81.6          # pt
        sig_x  = 97.4
        c.drawImage(_img(_assets.SIGN_DEEPIKA),
                    sig_x,                  sig_y - sig_h,
                    sig_w / 3, sig_h, mask="auto")
        c.drawImage(_img(_assets.SIGN_TEENA),
                    sig_x + sig_w / 3,      sig_y - sig_h,
                    sig_w / 3, sig_h, mask="auto")
        c.drawImage(_img(_assets.SIGN_SURIYA),
                    sig_x + 2 * sig_w / 3,  sig_y - sig_h,
                    sig_w / 3, sig_h, mask="auto")
