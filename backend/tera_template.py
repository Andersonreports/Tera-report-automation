"""
TERA Report Generator - PDF Template (v6 - Pixel-Perfect from Template PDFs)
=============================================================================
All coordinates derived by pdfplumber analysis of the three template PDFs:
  Receptive report template.pdf
  Prereceptive report template.pdf
  Postreceptive report template.pdf

Coordinate system: ReportLab (origin = bottom-left, y increases upward).
Conversion from pdfplumber (origin = top-left): RL_y = H - pdfplumber_y

CONFIRMED LAYOUT (from pdfplumber analysis):
  Header image  : x=72,  RL_y=718.9,  w=481.45, h=72.6
  Footer image  : x=72,  RL_y=12,     w=480.6,  h=36
  Title line 1  : centred, RL_y=698.1  (GillSansMT-Bold 18pt, blue)
  Title line 2  : centred, RL_y=666.8
  Patient table : x=45.84, RL_top=648.22, col widths=[111.26,7.08,205.61,91.22,9.01,109.10]
  Status heading: x=72,   RL_y=430.4,  GillSansMT-Bold 14pt, blue
  Status divider: RL_y=425.35
  Rec heading   : varies per type
  Rec divider   : varies per type

DATA COLUMNS (from TERA automation report Excel):
  Patient Name              → patient name
  Age                       → age (numeric)
  Sample ID                 → sample identifier
  Lab No.                   → lab number
  Biopsy No.                → specimen label (e.g. "Endometrial Biopsy- 1")
  Doctor Name               → referring clinician
  Center name               → hospital/clinic
  Cycle Type                → cycle type
  P4 /hCG injection  date time  → first P4 intake (string datetime)
  Biopsy time in hrs        → biopsy datetime (string, NOT the P+ hours)
  Biopsy time in hrs.1      → P+ hours (numeric, e.g. 120.0)
  TERA result               → receptivity classification
  Time for report           → embryo transfer timing (e.g. "144 + 2")
  Date of Received          → specimen receipt date (Timestamp)
"""

import os, io, re, base64, sys
from datetime import datetime


def _resource_path(relative: str) -> str:
    """Resolve path to a bundled resource.
    Works both in normal Python and when frozen by PyInstaller (sys._MEIPASS).
    """
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)

from reportlab.pdfgen          import canvas
from reportlab.lib.colors      import Color, black, white, HexColor
from reportlab.lib.utils       import ImageReader
from reportlab.lib.styles      import ParagraphStyle
from reportlab.lib.enums       import TA_JUSTIFY, TA_LEFT
from reportlab.platypus        import Paragraph, Table, TableStyle
from reportlab.pdfbase         import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily

import tera_assets

# ─── Colours ──────────────────────────────────────────────────────────────────
BLUE     = Color(0.122, 0.286, 0.49)      # #1F497D  – headings & title
BLUE_HEX = "#1F497D"                      # For PDF Paragraphs
MED_BLUE = Color(0.310, 0.506, 0.741)     # #4F81BD  – "This report reviewed…"
FIELD    = HexColor('#F1F1F7')            # lavender  – patient table background
GRAY_SIG = Color(0.2, 0.2, 0.2)          # #333333   – reviewer names & titles
BLACK    = black
WHITE    = white

# ─── Font registration ────────────────────────────────────────────────────────
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

_reg("GillSansMT-Bold", "GillSansMT-Bold.ttf")
_reg("SegoeUI-Bold",    "SegoeUI-Bold.ttf")
_reg("SegoeUI",         "SegoeUI.ttf")
_reg("DengXian",        "DengXian.ttf")
_reg("DengXian-Bold",   "DengXian_Bold.ttf")
_reg("Arial-Bold",      "Arial-BoldMT.ttf")
_reg("Arial",           "ArialMT.ttf")
_reg("Calibri",         "Calibri.ttf")
_reg("Calibri-Bold",    "Calibri-Bold.ttf")
_reg("SymbolMT",        "SymbolMT.ttf")

def _font_ok(name):
    try:
        pdfmetrics.getFont(name)
        return True
    except Exception:
        return False

# Register Calibri as a font family (same full 1.6 MB TTFs as PGTA uses — confirmed box-free)
if _font_ok("Calibri") and _font_ok("Calibri-Bold"):
    registerFontFamily("Calibri",
        normal="Calibri", bold="Calibri-Bold",
        italic="Calibri", boldItalic="Calibri-Bold")

# Register SegoeUI as a font family so <b> tags in Paragraph work correctly
if _font_ok("SegoeUI") and _font_ok("SegoeUI-Bold"):
    registerFontFamily("SegoeUI",
        normal="SegoeUI", bold="SegoeUI-Bold",
        italic="SegoeUI", boldItalic="SegoeUI-Bold")

# Register DengXian as a font family so <b> tags in Paragraph render bold correctly
if _font_ok("DengXian") and _font_ok("DengXian-Bold"):
    registerFontFamily("DengXian",
        normal="DengXian", bold="DengXian-Bold",
        italic="DengXian", boldItalic="DengXian-Bold")

# Font aliases (fall back to Helvetica variants if TTF not loaded)
F_TITLE  = "GillSansMT-Bold" if _font_ok("GillSansMT-Bold") else "Helvetica-Bold"
F_HDG    = "GillSansMT-Bold" if _font_ok("GillSansMT-Bold") else "Helvetica-Bold"
F_LBL    = "SegoeUI-Bold"    if _font_ok("SegoeUI-Bold")    else "Helvetica-Bold"
# DengXian: matches reference PDF body font exactly
F_BODY   = "DengXian"        if _font_ok("DengXian")        else "Helvetica"
F_BBOLD  = "DengXian-Bold"   if _font_ok("DengXian-Bold")   else "Helvetica-Bold"
F_SIG    = "SegoeUI"         if _font_ok("SegoeUI")         else "Helvetica"
F_SIGB   = "SegoeUI-Bold"    if _font_ok("SegoeUI-Bold")    else "Helvetica-Bold"
# Bullet: DengXian is the body font and reliably renders U+2022 as a filled circle
F_BULLET = "DengXian"        if _font_ok("DengXian")        else ("Calibri" if _font_ok("Calibri") else "Helvetica")

print(f"[tera_template] Fonts: TITLE={F_TITLE}  LBL={F_LBL}  BODY={F_BODY}  BULLET={F_BULLET}")

# ─── Page geometry ────────────────────────────────────────────────────────────
W, H = 612.0, 792.0

# Header: PGTA/Anderson shared header image (1280×193 px).
#   Drawn at x=72, w=468, h = 468*(193/1280) = 70.6 ≈ 71pt  (same as PGTA).
HDR_X, HDR_Y, HDR_W, HDR_H = 72.0, H - 72.0, 468.0, 72.0
# Footer: aligned to content area (same x/w as header).
#   Source image is footer_clean.png: 681×48px (Anderson Genetics white strip removed).
#   Natural height at 481.9pt wide: h = 481.9 × (48 / 681) = 33.97 ≈ 34pt.
FTR_X, FTR_Y, FTR_W, FTR_H = 72.75, 8.0,      481.9, 34.0

# Patient info table
TBL_X          = 45.84
TBL_TOP_RL     = H - 143.78    # 648.22 – RL y of table top edge
TBL_COL_WIDTHS = [111.26, 7.08, 200.61, 91.22, 9.01, 114.10]   # total 533.28; right value col widened 5pt for "Modified Natural Cycle"
TBL_W          = sum(TBL_COL_WIDTHS)
TBL_PAD_TOP    = 9              # vertical padding above text in each cell (3pt less = content 3pt higher)

# Section divider x-span (same on all pages)
DIV_X0, DIV_X1 = 72.75, 554.65

# ─── Per-result-type layout (all y in ReportLab space) ────────────────────────
# Computed directly from pdfplumber measurements of the three template PDFs.
RESULT_CFG = {
    "receptive": {
        # Receptivity chart image
        "chart_x": 411.85, "chart_y": H - 511.95, "chart_w": 141,    "chart_h": 130.5,
        # White text box drawn over the chart (left half)
        "box_x": 72, "box_y": H - 507.65, "box_w": 264.75, "box_h": 111.1,
        # Status paragraph text x and max width
        "status_x": 79.2,  "status_max_w": 257.55,
        # Recommendations section
        "hdg_recom_y":   H - 553.3,    # heading baseline
        "recom_line_y":  H - 562.5,    # divider line below heading
        "has_biopsy2":   False,
        "blast_x": 171.7, "blast_y": H - 613.0,
        "cleave_x":170.4, "cleave_y": H - 670.6,
        "reco_suffix": "post first progesterone intake",
        "recom_max_w": 280,
        # Icon
        "icon_y": H - 706.5,
        # Status text content
        "bold_phrase": "receptive endometrium",
        "displaced":   False,
        "asset": "RECEPTIVE",
    },
    "pre": {
        "chart_x": 334.70, "chart_y": H - 508.80, "chart_w": 218,    "chart_h": 127.3,
        "box_x": 72, "box_y": H - 510.90, "box_w": 250.25, "box_h": 125.6,
        "status_x": 79.2,  "status_max_w": 243.05,
        "hdg_recom_y":   H - 550.1,
        "recom_line_y":  H - 559.3,
        "has_biopsy2":   False,
        "blast_x": 171.7, "blast_y": H - 609.7,
        "cleave_x":170.4, "cleave_y": H - 667.3,
        "reco_suffix": "post first progesterone intake",
        "recom_max_w": 280,
        "icon_y": H - 703.3,
        "bold_phrase": "pre-receptive endometrium",
        "displaced":   True,
        "asset": "PRE_RECEPTIVE",
    },
    "post": {
        "chart_x": 336.00, "chart_y": H - 509.05, "chart_w": 216.85, "chart_h": 127.55,
        "box_x": 72, "box_y": H - 503.90, "box_w": 257.25, "box_h": 123.85,
        "status_x": 79.2,  "status_max_w": 250.05,
        "hdg_recom_y":   H - 520.0,
        "recom_line_y":  H - 530.0,
        "has_biopsy2":   True,
        "blast_x": 171.7, "blast_y": H - 620.0,
        "cleave_x":170.4, "cleave_y": H - 680.0,
        "reco_suffix": "post first P4 intake",
        "recom_max_w": 380,
        "icon_y": H - 715.0,
        "bold_phrase": "post-receptive endometrium",
        "displaced":   True,
        "asset": "POST_RECPTIVE",
    },
}

# ─── Drawing helpers ──────────────────────────────────────────────────────────
def _img(b64: str) -> ImageReader:
    return ImageReader(io.BytesIO(base64.b64decode(b64)))

def _divider(c, y):
    """Thin gray horizontal rule across the content width."""
    c.setStrokeColor(Color(0.6, 0.6, 0.6))
    c.setLineWidth(0.48)
    c.line(DIV_X0, y, DIV_X1, y)

def _wrap(c, text, x, y, max_w, font, size, leading):
    """Word-wrap text, return y after the last drawn line."""
    words = text.split()
    line  = ""
    for w in words:
        trial = line + w + " "
        if c.stringWidth(trial, font, size) <= max_w:
            line = trial
        else:
            if line:
                c.drawString(x, y, line.rstrip())
                y -= leading
            line = w + " "
    if line.strip():
        c.drawString(x, y, line.rstrip())
        y -= leading
    return y


def _wrap_justify(c, text, x, y, max_w, font, size, leading, first_line_indent=0):
    """Word-wrap text with full justification, return y after the last drawn line."""
    words = text.split()
    lines = []
    line = []
    for w in words:
        trial = " ".join(line + [w])
        indent = first_line_indent if not lines else 0
        if c.stringWidth(trial, font, size) <= (max_w - indent):
            line.append(w)
        else:
            if line:
                lines.append(line)
            line = [w]
    if line:
        lines.append(line)

    for idx, l in enumerate(lines):
        line_str = " ".join(l)
        indent = first_line_indent if idx == 0 else 0
        if idx == len(lines) - 1:
            # Last line: left-aligned
            c.drawString(x + indent, y, line_str)
        else:
            # Full justification
            if len(l) > 1:
                total_w = c.stringWidth(line_str, font, size)
                space_to_add = (max_w - indent) - total_w
                extra_space = space_to_add / (len(l) - 1)
                
                curr_x = x + indent
                for w_idx, w in enumerate(l):
                    c.drawString(curr_x, y, w)
                    curr_x += c.stringWidth(w, font, size) + c.stringWidth(" ", font, size) + extra_space
            else:
                c.drawString(x + indent, y, line_str)
        y -= leading
    return y


def _wrap_pm(c, text, x, y, max_w, font, size, leading):
    """Like _wrap but renders the ± word in Helvetica-Bold (Type1 built-in).
    Helvetica-Bold is a standard PDF font guaranteed to render ± (U+00B1).
    All other words are drawn in `font`.
    """
    PM = '\u00b1'
    PM_FONT = 'Helvetica-Bold'
    space_w = c.stringWidth(' ', font, size)

    def word_w(w):
        return c.stringWidth(PM, PM_FONT, size) if w == PM else c.stringWidth(w, font, size)

    def draw_line(words_list, lx, ly):
        cx = lx
        for i, w in enumerate(words_list):
            if i > 0:
                c.setFont(font, size)
                c.drawString(cx, ly, ' ')
                cx += space_w
            if w == PM:
                c.setFont(PM_FONT, size)
            else:
                c.setFont(font, size)
            c.drawString(cx, ly, w)
            cx += word_w(w)

    words = text.split()
    line_words, line_w = [], 0.0
    for w in words:
        ww = word_w(w)
        gap = space_w if line_words else 0.0
        if line_w + gap + ww <= max_w:
            line_words.append(w)
            line_w += gap + ww
        else:
            if line_words:
                draw_line(line_words, x, y)
                y -= leading
            line_words, line_w = [w], ww
    if line_words:
        draw_line(line_words, x, y)
        y -= leading
    return y


def _justified_block(c, text, x, y, max_w, font, size, leading):
    """Draw fully-justified paragraph; returns y position after the last line.

    ``y`` is the baseline of the first line (same convention as ``_wrap``).
    The caller must set the fill colour before calling this function.
    """
    style = ParagraphStyle(
        "JBlock",
        fontName=font, fontSize=size, leading=leading,
        alignment=TA_JUSTIFY,
        spaceAfter=0, spaceBefore=0,
    )
    para = Paragraph(text, style)
    _, h = para.wrap(max_w, 2000)
    # drawOn places the bottom-left corner at (x, bot_y).
    # We want the top of the first line near 'y', so bot_y = y - h + (leading - size).
    # The small offset aligns the first baseline with the legacy _wrap() position.
    offset = leading - size          # typically ≈11 pt for size=11, leading=22
    para.drawOn(c, x, y - h + offset)
    return y - h + offset


# ─── Main class ───────────────────────────────────────────────────────────────
class TERAReportGenerator:

    def __init__(self, data_row: dict, output_dir: str, with_logo: bool = False):
        self.d         = data_row
        self.out       = output_dir
        self.with_logo = with_logo

        # Classify result type from 'TERA result' column
        raw = str(self.d.get("TERA result",
              self.d.get("TERA result ",
              self.d.get("TERA Result", "")))).strip().lower()
        self.result_type = (
            "pre"  if "pre"  in raw else
            "post" if "post" in raw else
            "receptive"
        )
        self.cfg = RESULT_CFG[self.result_type]

        # Build output filename
        # Format: "{Name}_{Nth biopsy}_TERA_report_with logo.pdf"
        #      or "{Name}_{Nth biopsy}_TERA_report_without logo.pdf"
        name = self._s(self.d.get("Patient Name", "Unknown"))
        # Strip common honorifics
        name = re.sub(r'^(Mrs?\.|MRS?\.|Miss\.?|Ms\.?|Dr\.|DR\.)\s*', '', name).strip()
        # Remove filesystem-unsafe characters
        name = re.sub(r'[<>:"/\\|?*]', '_', name)
        bno_raw = self._s(self.d.get("Biopsy No.", self.d.get("Biopsy", "1")))
        bno = self._biopsy_ordinal(bno_raw)
        logo_tag = "with logo" if self.with_logo else "without logo"
        self.filename = f"{name}_{bno}_TERA_report_{logo_tag}.pdf"
        self.filepath = os.path.join(self.out, self.filename)

    # ── Public ────────────────────────────────────────────────────────────────
    def generate(self) -> str:
        c = canvas.Canvas(self.filepath, pagesize=(W, H))
        c.setTitle(self.filename)
        self._page1(c)
        c.showPage()
        self._page2(c)
        c.showPage()
        self._page3(c)
        c.save()
        return self.filepath

    # ═══════════════════════════════════════════════════════════════════════════
    # Shared header / footer
    # ═══════════════════════════════════════════════════════════════════════════
    def _header(self, c):
        """Draw header.
        with_logo=True  → Anderson shared header image (tera_assets.HEADER_LOGO).
        with_logo=False → nothing (no header image, no placeholder).
        """
        if not self.with_logo:
            return
        c.saveState()
        try:
            c.drawImage(_img(tera_assets.HEADER_LOGO),
                        HDR_X, HDR_Y, width=HDR_W, height=HDR_H,
                        mask="auto", preserveAspectRatio=False)
        except Exception as e:
            print(f"[TERA] Header err: {e}")
        c.restoreState()

    def _footer(self, c):
        """Draw footer image only when with_logo=True."""
        if not self.with_logo:
            return
        c.saveState()
        try:
            c.drawImage(_img(tera_assets.FOOTER),
                        FTR_X, FTR_Y, width=FTR_W, height=FTR_H,
                        mask="auto", preserveAspectRatio=False)
        except Exception:
            pass
        c.restoreState()

    # ═══════════════════════════════════════════════════════════════════════════
    # PAGE 1 – Patient report
    # ═══════════════════════════════════════════════════════════════════════════
    def _page1(self, c):
        self._header(c)
        self._footer(c)
        self._title_block(c)
        self._field_table(c)
        self._status_section(c)
        self._recom_section(c)

    def _title_block(self, c):
        """Centred two-line title (GillSansMT-Bold 18pt, blue).
        Exact y from pdfplumber: line1 bottom=93.9 → RL=698.1
                                 line2 bottom=125.2 → RL=666.8
        """
        c.setFont(F_TITLE, 18)
        c.setFillColor(BLUE)
        c.drawCentredString(W / 2, H - 93.9,
                            "Transcriptome based Endometrial Receptivity Assessment")
        c.drawCentredString(W / 2, H - 125.2, "(TERA)")

    def _field_table(self, c):
        """Patient info table (6 rows × 6 cols) with lavender background.
        Column widths from template: [111.26, 7.08, 205.61, 91.22, 9.01, 109.10]
        Table top  : RL y = H-143.78 = 648.22
        Top padding: 12 pt (text starts 12 pt below cell top edge)
        """
        rows = self._patient_rows()

        cell_style = ParagraphStyle(
            "TeraCell",
            fontName=F_LBL, fontSize=10, leading=12,
            textColor=BLACK, spaceAfter=0, spaceBefore=0,
        )

        def P(text):
            return Paragraph(str(text) if text else "", cell_style)

        data = [[P(l1), P(":"), P(v1), P(l2), P(":"), P(v2)]
                for l1, v1, l2, v2 in rows]

        tbl = Table(data, colWidths=TBL_COL_WIDTHS, rowHeights=None, hAlign="LEFT")
        tbl.setStyle(TableStyle([
            ("FONTNAME",      (0, 0), (-1, -1), F_LBL),
            ("FONTSIZE",      (0, 0), (-1, -1), 10),
            ("VALIGN",        (0, 0), (-1, -1), "TOP"),
            ("ALIGN",         (0, 0), (-1, -1), "LEFT"),
            # 2pt left padding shifts content 2pt right
            ("LEFTPADDING",   (0, 0), (-1, -1), 2),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 2),
            # Zero right padding on right-label column (col 3): "First P4 intake date"
            # is 90.5 pt; the column is 91.22 pt — needs all available space.
            ("RIGHTPADDING",  (3, 0), (3, -1), 0),
            ("TOPPADDING",    (0, 0), (-1, -1), TBL_PAD_TOP),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))

        tbl_w, tbl_h = tbl.wrap(TBL_W, 600)
        tbl_bot = TBL_TOP_RL - tbl_h

        # Draw lavender background first, then table (no stroke border)
        c.setFillColor(FIELD)
        c.rect(TBL_X, tbl_bot, tbl_w, tbl_h, fill=True, stroke=False)

        c.saveState()
        c.setStrokeColor(FIELD)          # suppress default table border
        tbl.drawOn(c, TBL_X, tbl_bot)
        c.restoreState()

    def _status_section(self, c):
        cfg = self.cfg

        # 1. Receptivity chart image (drawn first, sits behind white box)
        try:
            asset = getattr(tera_assets, cfg["asset"])
            c.saveState()
            c.setStrokeColor(WHITE)      # suppress any implicit border stroke
            c.setLineWidth(0)
            c.drawImage(_img(asset),
                        cfg["chart_x"], cfg["chart_y"],
                        width=cfg["chart_w"], height=cfg["chart_h"],
                        preserveAspectRatio=False)
            c.restoreState()
        except Exception:
            pass

        # 2. Section heading – GillSansMT-Bold 14pt, blue
        #    Exact baseline y from template: H-361.6 = 430.4
        c.setFont(F_HDG, 14)
        c.setFillColor(BLUE)
        c.drawString(72, H - 361.6, "Endometrial receptivity status")
        _divider(c, H - 366.65)         # exact template divider y

        # 3. White text box over the left part of the chart
        c.setFillColor(WHITE)
        c.rect(cfg["box_x"], cfg["box_y"],
               cfg["box_w"], cfg["box_h"],
               fill=True, stroke=False)

        # 4. Status paragraph – DengXian 12pt with inline bold for result phrase
        bh_int = self._int(self.d.get("Biopsy time in hrs.1", ""))
        bh_lbl = f"P+{bh_int} Hrs" if bh_int is not None else "the biopsy time"

        suffix = (" and therefore represents a displaced window of implantation."
                  if cfg["displaced"] else
                  " and therefore represents a window of implantation.")
        html = (f"The gene expression profile of the endometrial biopsy sample "
                f"performed on {bh_lbl} is indicative of a "
                f"<b>{cfg['bold_phrase']}</b>{suffix}")

        para_style = ParagraphStyle(
            "TeraStatus",
            fontName=F_BODY, fontSize=12, leading=24,
            alignment=TA_JUSTIFY,
            textColor=BLACK, spaceAfter=0, spaceBefore=0,
        )
        para = Paragraph(html, para_style)
        para_w, para_h = para.wrap(cfg["status_max_w"], 300)

        # Position: 6.5 pt top-padding inside the white box
        box_top_rl = cfg["box_y"] + cfg["box_h"]
        para.drawOn(c, cfg["status_x"], box_top_rl - 6.5 - para_h)

    def _recom_section(self, c):
        cfg = self.cfg

        # 1. Recommendation icon image (left margin, behind text)
        try:
            c.saveState()
            c.setStrokeColor(WHITE)      # suppress any implicit border stroke
            c.setLineWidth(0)
            c.drawImage(_img(tera_assets.RECOMENDATION),
                        72, cfg["icon_y"], width=70, height=124,
                        preserveAspectRatio=False)
            c.restoreState()
        except Exception:
            pass

        # 2. Section heading – GillSansMT-Bold 14pt, blue
        c.setFont(F_HDG, 14)
        c.setFillColor(BLUE)
        c.drawString(72, cfg["hdg_recom_y"],
                     "Recommendations for personalized Embryo Transfer (pET)")
        _divider(c, cfg["recom_line_y"])

        # 3. Transfer timing labels from Excel "Time for report" column
        tr_raw = str(self.d.get("Time for report",
                     self.d.get("Time for report ",
                     self.d.get("Corrected time for report ",
                     self.d.get("embryo transfer time in hrs", ""))))).strip()
        blast_lbl, cleave_lbl = self._parse_tr(tr_raw)
        suffix = cfg["reco_suffix"]

        c.setFont(F_BBOLD, 11)
        c.setFillColor(BLACK)

        # 4. Second biopsy note (post-receptive only) – appears between divider and transfer lines
        if cfg["has_biopsy2"]:
            # Robust Justified Direct Drawing: Full justification for a typeset look
            c.setFont(F_LBL, 11)
            draw_x = 72.0
            # Narrower width: align to DIV_X1 (554.65) to avoid crossing underlines
            wrap_total_w = DIV_X1 - draw_x - 5 
            
            # --- Note 1: Justified ---
            n1 = "A Second biopsy at P+98 Hrs and P+120Hrs is strongly recommended to confirm the Window of implantation."
            curr_y = cfg["recom_line_y"] - 14
            curr_y = _wrap_justify(c, n1, draw_x, curr_y, wrap_total_w, F_LBL, 11, 14)
            
            # --- Note 2: prefix Blue, rest Black, Fully Justified ---
            curr_y -= 8 # Gap between paragraphs
            prefix = "Note: "
            rem = "Patients with post-receptive endometria are prone to cycle-to-cycle variation. Hence repeat biopsy is suggested."
            
            c.setFillColor(BLUE)
            c.drawString(draw_x, curr_y, prefix)
            pw = c.stringWidth(prefix, F_LBL, 11)
            c.setFillColor(BLACK)
            
            curr_y = _wrap_justify(c, rem, draw_x, curr_y, wrap_total_w, F_LBL, 11, 14, first_line_indent=pw)

        reco_w = cfg.get("recom_max_w", 380.0)
        # Use Calibri-Bold for text; _wrap_pm switches to Helvetica-Bold for ±
        # (Helvetica-Bold is a built-in Type1 PDF font guaranteed to render ±)
        reco_font = "Calibri-Bold" if _font_ok("Calibri-Bold") else F_BBOLD

        # 5. Blastocyst transfer line
        _wrap_pm(c,
                 f"Blastocyst transfer (Day 5/6 embryo): {blast_lbl} {suffix}",
                 cfg["blast_x"], cfg["blast_y"], reco_w, reco_font, 11, 17)

        # 6. Cleavage stage transfer line
        _wrap_pm(c,
                 f"Cleavage stage transfer (Day 3 embryo): {cleave_lbl} {suffix}",
                 cfg["cleave_x"], cfg["cleave_y"], reco_w, reco_font, 11, 17)

    # ═══════════════════════════════════════════════════════════════════════════
    # PAGE 2 – About TERA + Methodology
    # ═══════════════════════════════════════════════════════════════════════════
    ABOUT_PARAS = [
        ("Embryo implantation is a highly organized process during which the embryo attaches "
         "to the surface of the endometrium. Synchronous structural and functional remodelling "
         "of the uterine endometrium and the blastocyst is essential for successful implantation. "
         "The window of implantation (WOI) is a limited time span during which crosstalk between "
         "a receptive uterine endometrium and a competent blastocyst occurs effectively."),
        ("A displacement in the window of implantation is among the leading causes of recurrent "
         "implantation failure and is observed in 30% of women undergoing ART conception. It is "
         "frequently observed that an endometrium that appears morphologically ready for "
         "implantation may not express appropriate transcriptomic response characteristic of WOI. "
         "Therefore, an accurate molecular description of the endometrial transcriptomic signature "
         "is essential in ensuring implantation of embryos with good development potentials."),
        ("TERA is designed to provide personalized embryo implantation time on the basis cutting-edge "
         "technical expertise based on Next generation Sequencing (NGS) that allows us to study "
         "unique endometrial signature representation of WOI. Highest reproducibility in TERA "
         "results is observed in HRT cycles."),
    ]
    METHOD_BULLETS = [
        ("TERA detects mRNA expression in endometrial tissues using NGS based RNA-seq method "
         "combined with Artificial Intelligence (AI) empowered data analysis platform to discern "
         "endometrial status. The results are used as references for embryo transfer to improve "
         "chances of successful implantation."),
        ("The duration of WOI may vary from patient to patient. The results of this test suggest "
         "the optimal time to transfer embryos and enable accurate clinical recommendations for "
         "embryo transfer."),
    ]

    def _page2(self, c):
        self._header(c)
        self._footer(c)

        # Page content width: from x=72 to divider end x=554.65 = 482.65 pt
        CONTENT_W = DIV_X1 - 72

        # "About TERA" heading
        # pdfplumber: x=72, top=75.2, bottom=89.2 → RL baseline = H-89.2 = 702.8
        c.setFont(F_HDG, 14)
        c.setFillColor(BLUE)
        c.drawString(72, H - 89.2, "About TERA")
        _divider(c, H - 98.85)          # template divider y = H-98.85

        # Body paragraphs – DengXian 11pt, leading 22, justified
        # First line baseline from template: pdfplumber bottom≈125.4 → H-125.4 = 666.6
        y = H - 125.4
        c.setFillColor(BLACK)
        for para in self.ABOUT_PARAS:
            y = _justified_block(c, para, 72, y, CONTENT_W, F_BODY, 11, 22)
            y -= 23     # inter-paragraph gap (matches reference: 23.3 pt bottom-to-top)

        # "Methodology" heading — 8pt below last About TERA paragraph
        meth_y = y - 8
        c.setFont(F_HDG, 14)
        c.setFillColor(BLUE)
        c.drawString(78.9, meth_y, "Methodology")
        _divider(c, meth_y - 9)         # ~9 pt below heading baseline

        # Bullet points – filled circle drawn directly (font-independent), body text 11pt, justified
        y = meth_y - 37
        for bullet in self.METHOD_BULLETS:
            # Draw bullet as solid filled circle centered vertically with the text cap height
            c.setFillColor(BLACK)
            c.circle(92.5, y + 4, 2.5, fill=1, stroke=0)
            y = _justified_block(c, bullet, 108, y, CONTENT_W - 36, F_BODY, 11, 22)
            y -= 10

    # ═══════════════════════════════════════════════════════════════════════════
    # PAGE 3 – References + Signatures
    # ═══════════════════════════════════════════════════════════════════════════
    REFS = [
        "Achache H, Revel A. Hum Reprod Update, 2006, 12(6):731-46.",
        "Teh W T, Mcbain J, Rogers P. Journal of Assisted Reproduction & Genetics, 2016, 33(11):1-12.",
        "Mahajan N. Journal of Human Reproductive Sciences, 2015, 8(3):121-129.",
        "Ruiz-Alonso M, Blesa D, Díaz-Gimeno, Patricia, et al. Fertility and Sterility, 2013, 100(3):818-824.",
    ]

    def _page3(self, c):
        self._header(c)
        self._footer(c)

        # "References" heading
        # pdfplumber: x=78.9, top=75.2 → RL baseline = H-89.2 = 702.8
        c.setFont(F_HDG, 14)
        c.setFillColor(BLUE)
        c.drawString(78.9, H - 89.2, "References")
        _divider(c, H - 98.1)

        # Reference entries – DengXian 11pt, ~27 pt spacing
        # First ref baseline from template: pdfplumber bottom≈112.15 → RL=679.85
        REF_W = DIV_X1 - 93.9           # text width from indent to right edge
        y = H - 112.15
        c.setFont(F_BODY, 11)
        c.setFillColor(BLACK)
        for i, ref in enumerate(self.REFS, 1):
            c.drawString(75.9, y, f"{i}.")
            _wrap(c, ref, 93.9, y, REF_W, F_BODY, 11, 14)
            y -= 27

        # "This report has been reviewed and approved by:"
        # Arial-BoldMT 12pt, medium blue – pdfplumber top=219.4 → RL≈H-231.9=560.1
        c.setFont(F_SIGB, 12)
        c.setFillColor(MED_BLUE)
        c.drawString(75.9, H - 231.9,
                     "This report has been reviewed and approved by:")

        # Signature images
        # Positions from pdfplumber analysis (averaged across three template PDFs)
        sigs = [
            (80.75,  H - 290.95, 71.15,  33.1,  tera_assets.SIVASHANKAR_SIGN),
            (237.75, H - 290.95, 74.25,  33.05, tera_assets.FIONA_SIGN),
            (406.25, H - 297.75, 100.15, 42.3,  getattr(tera_assets, "SACHIN_SIGN", None)),
        ]
        for sx, sy, sw, sh, asset in sigs:
            if asset:
                try:
                    c.drawImage(_img(asset), sx, sy, width=sw, height=sh,
                                preserveAspectRatio=True, mask="auto")
                except Exception:
                    pass
            else:
                c.setStrokeColor(BLACK)
                c.setLineWidth(0.7)
                c.line(sx, sy + 10, sx + sw, sy + 10)

        # Reviewer names – SegoeUI 11pt, dark gray
        # pdfplumber top=301.9, cap_height≈8 → RL baseline = H-309.9 = 482.1
        c.setFont(F_SIG, 11)
        c.setFillColor(GRAY_SIG)
        name_y = H - 309.9
        c.drawString(72.0,  name_y, "S. Sivasankar, Ph. D")
        c.drawString(208.0, name_y, "Fiona D'Souza, Ph. D")
        c.drawString(395.0, name_y, "Sachin D Honguntikar, Ph. D")

        # Reviewer titles – SegoeUI 11pt, dark gray
        # pdfplumber top=320.0, cap_height≈8 → RL baseline = H-328 = 464
        role_y = H - 328.0
        c.drawString(72.0,  role_y, "Molecular Biologist")
        c.drawString(208.0, role_y, "Head -Scientific Operations")
        c.drawString(395.0, role_y, "Head- Clinical Genetics")

    # ═══════════════════════════════════════════════════════════════════════════
    # Data helpers
    # ═══════════════════════════════════════════════════════════════════════════
    def _patient_rows(self):
        d     = self.d
        name  = self._s(d.get("Patient Name", ""))
        # PIN field in report ← Sample ID column (col C of Excel)
        pin   = self._s(d.get("Sample ID", "")) or "Not Provided"
        # Sample Number field in report ← Lab No. column (col H of Excel)
        sid   = self._s(d.get("Lab No.", ""))
        age_r = self._s(d.get("Age", ""))
        age   = f"{age_r} Years" if age_r else "Not Provided"
        doc   = self._s(d.get("Doctor Name", "")) or "Not Provided"
        hosp  = self._s(d.get("Center name", d.get("Hospital", d.get("Hospital ", ""))))
        # Cycle type display:
        #   HRT               → "HRT; P+{N}"  (N = Biopsy column value, e.g. 5)
        #   Modified Natural Cycle → "Modified Natural Cycle"  (no suffix)
        cyc_raw     = self._s(d.get("Cycle Type", d.get("Cycle type", "HRT")))
        biopsy_days = self._int(d.get("Biopsy", ""))
        cyc_upper   = cyc_raw.upper()
        if "HRT" in cyc_upper:
            cyc = (f"HRT; P+{biopsy_days}" if biopsy_days is not None else "HRT")
        elif cyc_raw:
            cyc = cyc_raw          # e.g. "Modified Natural Cycle" — unchanged
        else:
            cyc = "Not Provided"
        bno   = self._s(d.get("Biopsy No.",  d.get("Biopsy", "")))
        # P4 date: 'P4 /hCG injection  date time' column (string datetime)
        p4d   = self._dt(d.get("P4 /hCG injection  date time", ""))
        # Biopsy date: 'Biopsy time in hrs' stores the biopsy event datetime (NOT the P+ hours)
        biod  = self._dt(d.get("Biopsy time in hrs", ""))
        # Receipt date: Timestamp from 'Date of Received'
        rcpt  = self._dt(d.get("Date of Received", ""), date_only=True)
        # Custom Report Date from field, fallback to today
        rep_date_raw = self._s(d.get("Report Date", ""))
        today = rep_date_raw if rep_date_raw else datetime.today().strftime("%d-%m-%Y")

        return [
            ("Patient Name",          name,  "PIN",                  pin),
            ("Date of Birth/ Age",    age,   "Sample Number",        sid),
            ("Referring Clinician",   doc,   "Cycle type",           cyc),
            ("Hospital/Clinic",       hosp,  "First P4 intake date", p4d),
            ("Specimen",              bno,   "Biopsy date",          biod),
            ("Specimen receipt date", rcpt,  "Report date",          today),
        ]

    @staticmethod
    def _biopsy_ordinal(bno_raw: str) -> str:
        """Convert biopsy string to ordinal form.
        'Endometrial Biopsy- 1' → '1st biopsy'
        'Endometrial Biopsy- 2' → '2nd biopsy'
        Fallback: return raw string unchanged.
        """
        m = re.search(r'(\d+)', bno_raw)
        if m:
            n = int(m.group(1))
            if 11 <= (n % 100) <= 13:
                suffix = 'th'
            else:
                suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
            return f"{n}{suffix} biopsy"
        return bno_raw

    @staticmethod
    def _s(val) -> str:
        """Return clean string; empty string for NaN/None/NaT variants."""
        s = str(val).strip()
        return "" if s in ("nan", "NaT", "None", "NaN") else s

    @staticmethod
    def _int(val):
        """Safe integer conversion with mathematical rounding; returns None on failure."""
        if val is None:
            return None
        s = str(val).strip()
        if s in ("", "nan", "NaT", "None", "NaN"):
            return None
        try:
            import math
            f = float(s)
            return math.floor(f + 0.5)
        except Exception:
            return None

    @staticmethod
    def _dt(val, date_only=False) -> str:
        """Format a date/datetime value as 'DD-MM-YYYY HH:MM Hrs'.
        Accepts: pandas Timestamp, ISO datetime strings ('2026-02-02 12:00:00').
        date_only=True forces date-only output regardless of time component.
        """
        if val is None:
            return ""
        # Handle pandas Timestamp / NaT
        try:
            from pandas import Timestamp, NaT as PD_NAT
            if isinstance(val, Timestamp):
                if val is PD_NAT:
                    return ""
                if date_only or (val.hour == 0 and val.minute == 0):
                    return val.strftime("%d-%m-%Y")
                return val.strftime("%d-%m-%Y %H:%M Hrs")
        except Exception:
            pass
        s = str(val).strip()
        if s in ("", "nan", "NaT", "None", "NaN"):
            return ""
        # Try parsing ISO datetime string formats
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%d-%m-%Y %H:%M", "%Y-%m-%d"):
            try:
                dt = datetime.strptime(s, fmt)
                if date_only or (dt.hour == 0 and dt.minute == 0):
                    return dt.strftime("%d-%m-%Y")
                return dt.strftime("%d-%m-%Y %H:%M Hrs")
            except ValueError:
                continue
        return s   # fallback: return as-is

    @staticmethod
    def _parse_tr(raw: str):
        """Parse transfer time string (e.g. '144 + 2') into labelled strings.
        Blastocyst (Day 5/6) uses the base value.
        Cleavage   (Day 3)   = blastocyst - 48 hrs (constant Day 5→Day 3 offset).
        Returns (blast_label, cleave_label).
        """
        if not raw or raw in ("nan", "NaT", "None", ""):
            return "N/A", "N/A"
        m = re.match(r'^\s*(\d+(?:\.\d+)?)\s*\+\s*(\d+)', raw)
        if m:
            base   = round(float(m.group(1)))
            margin = m.group(2)
            return f"{base} \u00b1 {margin} hrs", f"{base - 48} \u00b1 {margin} hrs"
        try:
            base = round(float(raw))
            return f"{base} \u00b1 2 hrs", f"{base - 48} \u00b1 2 hrs"
        except Exception:
            return raw, "N/A"
