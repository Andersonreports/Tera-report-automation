"""
TERA Report Generator - Desktop Application v2
===============================================
Comprehensive GUI for TERA (Transcriptome based Endometrial Receptivity Assessment)
report generation with:
  - Manual patient data entry + live PDF preview
  - Bulk Excel upload with per-row editing
  - Visual PDF comparison tool (auto vs reference)
  - Draft save / load (JSON)
  - QSettings persistence for output folders
"""

import sys
import os
import json
import re
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path


def _resource_path(relative: str) -> str:
    """Resolve path to a bundled resource (works frozen and un-frozen)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)

import pandas as pd

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QLabel, QLineEdit, QPushButton, QFileDialog,
    QTableWidget, QTableWidgetItem, QMessageBox, QProgressBar,
    QGroupBox, QFormLayout, QScrollArea, QComboBox,
    QStyle, QSplitter, QTextBrowser, QDialog, QDialogButtonBox,
    QHeaderView, QSizePolicy, QFrame, QCheckBox, QRadioButton,
    QButtonGroup,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSettings, QTimer, QItemSelectionModel
from PyQt6.QtGui import QPixmap, QFont, QColor, QIcon

from tera_template import TERAReportGenerator

try:
    import pypdfium2 as _pdfium
    PYPDFIUM_OK = True
except ImportError:
    PYPDFIUM_OK = False

# ─── Field definitions ─────────────────────────────────────────────────────────
# Each entry: (display_label, data_key, widget_type, options_or_default_str)
# widget_type: "line" | "combo"
TERA_FIELD_DEFS = [
    ("Patient Name",         "Patient Name",                   "line",  ""),
    ("Age (Years)",          "Age",                            "line",  ""),
    ("Sample ID",            "Sample ID",                      "line",  ""),
    ("Lab No.",              "Lab No.",                        "line",  ""),
    ("Biopsy No.",           "Biopsy No.",                     "line",  "Endometrial Biopsy- 1"),
    ("Doctor Name",          "Doctor Name",                    "line",  ""),
    ("Center / Hospital",    "Center name",                    "line",  ""),
    ("Cycle Type",           "Cycle Type",                     "combo", ["HRT", "Modified Natural Cycle", "Natural Cycle", "Natural", "Stimulated"]),
    ("P+ Day (Biopsy)",      "Biopsy",                         "line",  "5"),
    ("P4/hCG Date & Time",   "P4 /hCG injection  date time",  "line",  ""),
    ("Biopsy Date & Time",   "Biopsy time in hrs",             "line",  ""),
    ("P+ Hours",             "Biopsy time in hrs.1",           "line",  ""),
    ("TERA Result",          "TERA result",                    "combo", ["Receptive", "Pre-receptive", "Post-receptive"]),
    ("Time for Report",      "Time for report",                "line",  ""),
    ("Date Received",        "Date of Received",               "line",  ""),
]

# Columns shown in the bulk upload table (S.No. + name only — details in editor)
BULK_DISPLAY_COLS = ["S. No.", "Patient Name"]


# ─── Helpers ───────────────────────────────────────────────────────────────────
def _clean(v) -> str:
    """Return clean string; empty for NaN/None/NaT."""
    s = str(v).strip()
    return "" if s in ("nan", "NaT", "None", "NaN", "") else s


def _open_folder(path: str):
    """Open a folder in the system file manager."""
    try:
        subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


# ─── Worker: live PDF preview ──────────────────────────────────────────────────
class PreviewWorker(QThread):
    """Generate a temp PDF then render every page to a list of PNG bytes."""
    finished = pyqtSignal(str)   # path to generated PDF
    error    = pyqtSignal(str)

    def __init__(self, data_row: dict, tmp_pdf: str, with_logo: bool = False):
        super().__init__()
        self.data_row  = data_row
        self.tmp_pdf   = tmp_pdf
        self.with_logo = with_logo

    def run(self):
        try:
            tmp_dir = os.path.dirname(self.tmp_pdf)
            gen = TERAReportGenerator(self.data_row, tmp_dir, with_logo=self.with_logo)
            gen.filepath = self.tmp_pdf
            gen.filename = os.path.basename(self.tmp_pdf)
            gen.generate()
            self.finished.emit(self.tmp_pdf)
        except Exception as e:
            import traceback
            self.error.emit(traceback.format_exc())


# ─── Worker: batch report generation ──────────────────────────────────────────
class ReportGeneratorWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(int, list)   # (success_count, [error_messages])

    def __init__(self, rows: list, output_dir: str, with_logo: bool = False):
        super().__init__()
        self.rows       = rows
        self.output_dir = output_dir
        self.with_logo  = with_logo

    def run(self):
        ok, errs = 0, []
        total = len(self.rows)
        for i, row in enumerate(self.rows, 1):
            name = _clean(row.get("Patient Name", f"Row {i}")) or f"Row {i}"
            try:
                self.progress.emit(int((i - 1) / total * 100),
                                   f"Generating {i}/{total}: {name}…")
                gen = TERAReportGenerator(row, self.output_dir,
                                         with_logo=self.with_logo)
                gen.generate()
                ok += 1
            except Exception as e:
                import traceback
                errs.append(f"{name}: {e}\n{traceback.format_exc()}")
        self.progress.emit(100, "Complete")
        self.finished.emit(ok, errs)


# ─── Worker: PDF auto-diff ──────────────────────────────────────────────────────
class PDFDiffWorker(QThread):
    """Compares two TERA PDFs page-by-page using pdfplumber text extraction.

    Emits:
        finished(str)  – HTML summary of differences
        error(str)     – traceback on failure
    """
    finished = pyqtSignal(str)  # HTML report
    error    = pyqtSignal(str)

    # TERA-specific field regions to check on page 1 (y bands in pdfplumber coords)
    # These are approximate bounding boxes [x0, y0, x1, y1] in pt (origin top-left)
    _P1_FIELDS = {
        "Title":          (45,  60, 570, 145),
        "Patient Info":   (45, 144, 570, 250),
        "Status Section": (45, 360, 570, 520),
        "Recommendations":(45, 520, 570, 710),
    }
    _P2_FIELDS = {
        "About TERA":     (45,  45, 570, 420),
        "Methodology":    (45, 420, 570, 760),
    }
    _P3_FIELDS = {
        "Reviewer Block": (45,  45, 570, 760),
    }

    def __init__(self, left_path: str, right_path: str):
        super().__init__()
        self.left_path  = left_path
        self.right_path = right_path

    # ── helpers ────────────────────────────────────────────────────────────────
    @staticmethod
    def _norm(s: str) -> str:
        """Normalise whitespace for comparison."""
        return re.sub(r'\s+', ' ', s).strip()

    @staticmethod
    def _page_text(page) -> str:
        return page.extract_text() or ""

    @staticmethod
    def _region_text(page, bbox) -> str:
        """Extract text from a bounding-box region (x0,y0,x1,y1 pdfplumber)."""
        try:
            crop = page.within_bbox(bbox)
            return crop.extract_text() or ""
        except Exception:
            return ""

    @staticmethod
    def _word_diff(a: str, b: str) -> list[tuple[str, str]]:
        """Return list of (left_word, right_word) pairs where they differ."""
        import difflib
        sm   = difflib.SequenceMatcher(None, a.split(), b.split(), autojunk=False)
        diffs = []
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag != "equal":
                l_chunk = " ".join(a.split()[i1:i2])
                r_chunk = " ".join(b.split()[j1:j2])
                diffs.append((l_chunk, r_chunk))
        return diffs

    # ── main comparison ────────────────────────────────────────────────────────
    def run(self):
        try:
            import pdfplumber
        except ImportError:
            self.error.emit(
                "pdfplumber not installed.\n\nRun: pip install pdfplumber")
            return
        try:
            html = self._compare(pdfplumber)
            self.finished.emit(html)
        except Exception:
            import traceback
            self.error.emit(traceback.format_exc())

    def _compare(self, pdfplumber) -> str:
        sections = []  # list of (label, issues:[str])

        with pdfplumber.open(self.left_path) as ldoc, \
             pdfplumber.open(self.right_path) as rdoc:

            lpages = ldoc.pages
            rpages = rdoc.pages
            n_left  = len(lpages)
            n_right = len(rpages)

            # ── page-count check ──────────────────────────────────────────────
            if n_left != n_right:
                sections.append(("Page Count",
                    [f"Left PDF has <b>{n_left}</b> pages, "
                     f"Right PDF has <b>{n_right}</b> pages."]))
            else:
                sections.append(("Page Count", [f"Both PDFs have {n_left} pages. ✓"]))

            n_common = min(n_left, n_right)

            # ── per-page analysis ─────────────────────────────────────────────
            page_regions = [
                ("Page 1", self._P1_FIELDS),
                ("Page 2", self._P2_FIELDS),
                ("Page 3", self._P3_FIELDS),
            ]
            for pg_idx in range(n_common):
                pg_label, regions = (page_regions[pg_idx]
                                     if pg_idx < len(page_regions)
                                     else (f"Page {pg_idx+1}", {}))
                lp = lpages[pg_idx]
                rp = rpages[pg_idx]

                page_issues = []

                # Full-page text similarity
                lt = self._norm(self._page_text(lp))
                rt = self._norm(self._page_text(rp))
                if lt == rt:
                    page_issues.append("Full page text is identical. ✓")
                else:
                    diffs = self._word_diff(lt, rt)
                    page_issues.append(
                        f"<span style='color:#c0392b'>Full page text differs "
                        f"({len(diffs)} change(s) found).</span>")
                    for (lw, rw) in diffs[:20]:   # cap at 20 per page
                        left_disp  = f"<span style='background:#fde8e8'>{lw or '(empty)'}</span>"
                        right_disp = f"<span style='background:#e8f5e9'>{rw or '(empty)'}</span>"
                        page_issues.append(
                            f"  <tt>Left:</tt> {left_disp}  →  "
                            f"<tt>Right:</tt> {right_disp}")
                    if len(diffs) > 20:
                        page_issues.append(
                            f"  … and {len(diffs)-20} more difference(s). "
                            "Check side-by-side view for details.")

                # Region-level checks
                if regions:
                    for region_name, bbox in regions.items():
                        lr = self._norm(self._region_text(lp, bbox))
                        rr = self._norm(self._region_text(rp, bbox))
                        if lr == rr:
                            page_issues.append(
                                f"  [{region_name}] identical ✓")
                        else:
                            rdiffs = self._word_diff(lr, rr)
                            page_issues.append(
                                f"  <span style='color:#c0392b'>"
                                f"[{region_name}] {len(rdiffs)} difference(s):</span>")
                            for (lw, rw) in rdiffs[:8]:
                                ldisp = f"<span style='background:#fde8e8'>{lw or '(empty)'}</span>"
                                rdisp = f"<span style='background:#e8f5e9'>{rw or '(empty)'}</span>"
                                page_issues.append(
                                    f"    <tt>L:</tt> {ldisp}  →  "
                                    f"<tt>R:</tt> {rdisp}")

                sections.append((pg_label, page_issues))

        return self._build_html(sections)

    def _build_html(self, sections: list) -> str:
        rows = []
        any_diff = False
        for label, issues in sections:
            has_diff = any("color:#c0392b" in i for i in issues)
            if has_diff:
                any_diff = True
            hdr_bg = "#fde8e8" if has_diff else "#e8f5e9"
            hdr_color = "#c0392b" if has_diff else "#196F3D"
            rows.append(
                f"<div style='margin-bottom:14px;border:1px solid #ddd;"
                f"border-radius:6px;overflow:hidden;'>"
                f"<div style='background:{hdr_bg};padding:8px 12px;"
                f"font-weight:bold;color:{hdr_color};font-size:14px;'>{label}</div>"
                f"<div style='padding:8px 14px;font-family:monospace;"
                f"font-size:12px;line-height:1.8;'>")
            for issue in issues:
                rows.append(f"<div>{issue}</div>")
            rows.append("</div></div>")

        summary = (
            "<span style='color:#c0392b;font-weight:bold'>Differences found — "
            "review highlighted items.</span>"
            if any_diff else
            "<span style='color:#196F3D;font-weight:bold'>"
            "No differences detected. PDFs match. ✓</span>"
        )

        return f"""<html><head>
<style>
body {{ font-family:'Segoe UI',Arial,sans-serif; background:#f8f9fa;
        color:#333; padding:16px; }}
.summary {{ background:#fff; border:2px solid #1F497D; border-radius:6px;
           padding:12px 16px; margin-bottom:16px; font-size:15px; }}
</style></head><body>
<div class="summary">{summary}</div>
{"".join(rows)}
</body></html>"""


# ─── (RowEditDialog removed — editing is now inline in the Bulk tab) ───────────


# ─── Main application ──────────────────────────────────────────────────────────
class TERAReportApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.settings       = QSettings("TERA", "ReportGenerator")
        self.bulk_rows      = []          # list of data dicts (possibly edited)
        self._preview_worker = None
        self._gen_worker     = None
        self._tmp_pdf = os.path.join(tempfile.gettempdir(), "tera_live_preview.pdf")

        # Debounce timer for live preview (manual tab)
        self._preview_timer = QTimer()
        self._preview_timer.setSingleShot(True)
        self._preview_timer.setInterval(900)
        self._preview_timer.timeout.connect(self._run_preview)

        # Bulk tab state
        self._bulk_preview_worker  = None
        self._bulk_tmp_pdf = os.path.join(tempfile.gettempdir(), "tera_bulk_preview.pdf")
        self._bulk_editor_inputs   = {}   # key → (widget, wtype)
        self._bulk_current_row     = -1   # currently selected row index
        self._bulk_preview_timer   = QTimer()
        self._bulk_preview_timer.setSingleShot(True)
        self._bulk_preview_timer.setInterval(900)
        self._bulk_preview_timer.timeout.connect(self._bulk_run_preview)

        self._init_ui()
        self._load_settings()

    # ═══════════════════════════════════════════════════════════════════════════
    # UI bootstrap
    # ═══════════════════════════════════════════════════════════════════════════
    def _init_ui(self):
        self.setWindowTitle("TERA Report Generator")
        _icon_path = _resource_path("tera_icon.png")
        if os.path.exists(_icon_path):
            self.setWindowIcon(QIcon(_icon_path))
        self.setMinimumSize(1350, 840)
        self.resize(1500, 900)

        central = QWidget()
        self.setCentralWidget(central)
        vbox = QVBoxLayout(central)

        # App title bar
        title_row = QHBoxLayout()
        lbl = QLabel("TERA Report Generator")
        lbl.setStyleSheet("font-size:22px;font-weight:bold;padding:8px 4px;color:#1F497D;")
        title_row.addWidget(lbl)
        title_row.addStretch()
        vbox.addLayout(title_row)

        # Tabs
        self.tabs = QTabWidget()
        vbox.addWidget(self.tabs)

        self.tabs.addTab(self._create_manual_tab(), "Manual Entry")
        self.tabs.setTabIcon(
            0, self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView))

        self.tabs.addTab(self._create_bulk_tab(), "Bulk Upload")
        self.tabs.setTabIcon(
            1, self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogListView))

        self.tabs.addTab(self._create_comparison_tab(), "Report Comparison")
        self.tabs.setTabIcon(
            2, self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload))

        self.tabs.addTab(self._create_guide_tab(), "User Guide")
        self.tabs.setTabIcon(
            3, self.style().standardIcon(QStyle.StandardPixmap.SP_MessageBoxInformation))

        self.statusBar().showMessage("Ready")

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 1 – Manual Entry
    # ═══════════════════════════════════════════════════════════════════════════
    def _create_manual_tab(self) -> QWidget:
        tab = QWidget()
        outer = QHBoxLayout(tab)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        outer.addWidget(splitter)

        # ── Left panel: input form ────────────────────────────────────────────
        left_widget = QWidget()
        left_vbox   = QVBoxLayout(left_widget)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner  = QWidget()
        form   = QFormLayout(inner)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        scroll.setWidget(inner)
        left_vbox.addWidget(scroll)

        self._manual_inputs = {}
        for display_lbl, key, wtype, opts in TERA_FIELD_DEFS:
            if wtype == "combo":
                w = QComboBox()
                w.addItems(opts)
                w.currentTextChanged.connect(self._schedule_preview)
            else:
                default = opts if isinstance(opts, str) else ""
                w = QLineEdit(default)
                w.textChanged.connect(self._schedule_preview)
            form.addRow(f"{display_lbl}:", w)
            self._manual_inputs[key] = (w, wtype)

        # Draft / clear buttons
        btn_row = QHBoxLayout()
        for text, slot in [("Save Draft", self._manual_save_draft),
                            ("Load Draft", self._manual_load_draft),
                            ("Clear Form", self._manual_clear)]:
            btn = QPushButton(text)
            btn.clicked.connect(slot)
            btn_row.addWidget(btn)
        btn_row.addStretch()
        left_vbox.addLayout(btn_row)

        # Generate group
        gen_grp    = QGroupBox("Generate Report")
        gen_layout = QVBoxLayout(gen_grp)

        out_row = QHBoxLayout()
        self._manual_out_lbl = QLabel("No folder selected")
        self._manual_out_lbl.setStyleSheet(
            "padding:4px;border:1px solid #ccc;background:white;")
        btn_out = QPushButton("Select Output Folder")
        btn_out.setIcon(self.style().standardIcon(
            QStyle.StandardPixmap.SP_DirOpenIcon))
        btn_out.clicked.connect(self._manual_browse_output)
        out_row.addWidget(self._manual_out_lbl, 1)
        out_row.addWidget(btn_out)
        gen_layout.addLayout(out_row)

        # Logo toggle (with / without)
        logo_row = QHBoxLayout()
        logo_row.addWidget(QLabel("Export:"))
        self._manual_logo_grp  = QButtonGroup(self)
        self._manual_no_logo   = QRadioButton("Without Logo")
        self._manual_with_logo = QRadioButton("With Logo")
        self._manual_with_logo.setChecked(True)
        self._manual_logo_grp.addButton(self._manual_no_logo)
        self._manual_logo_grp.addButton(self._manual_with_logo)
        logo_row.addWidget(self._manual_no_logo)
        logo_row.addWidget(self._manual_with_logo)
        logo_row.addStretch()
        gen_layout.addLayout(logo_row)

        self._manual_gen_btn = QPushButton("Generate Report")
        self._manual_gen_btn.setStyleSheet(
            "background-color:#1F497D;color:white;font-weight:bold;padding:8px;")
        self._manual_gen_btn.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay))
        self._manual_gen_btn.clicked.connect(self._manual_generate)
        gen_layout.addWidget(self._manual_gen_btn)

        left_vbox.addWidget(gen_grp)
        splitter.addWidget(left_widget)

        # ── Right panel: live preview ─────────────────────────────────────────
        right_grp  = QGroupBox("Live PDF Preview")
        right_vbox = QVBoxLayout(right_grp)

        prev_toolbar = QHBoxLayout()
        btn_refresh = QPushButton("Refresh Preview")
        btn_refresh.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload))
        btn_refresh.clicked.connect(self._run_preview)
        self._preview_status = QLabel("Fill in patient details and click Refresh Preview")
        self._preview_status.setStyleSheet("color:gray;font-style:italic;")
        self._preview_status.setWordWrap(True)
        prev_toolbar.addWidget(btn_refresh)
        prev_toolbar.addWidget(self._preview_status, 1)
        right_vbox.addLayout(prev_toolbar)

        # Pages scroll area
        prev_scroll = QScrollArea()
        prev_scroll.setWidgetResizable(True)
        self._preview_inner = QWidget()
        self._preview_vbox  = QVBoxLayout(self._preview_inner)
        self._preview_vbox.setAlignment(
            Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        prev_scroll.setWidget(self._preview_inner)
        right_vbox.addWidget(prev_scroll)

        splitter.addWidget(right_grp)
        splitter.setSizes([440, 800])
        return tab

    # ── Preview helpers ────────────────────────────────────────────────────────
    def _schedule_preview(self):
        self._preview_timer.start()

    def _run_preview(self):
        if self._preview_worker and self._preview_worker.isRunning():
            return
        data = self._get_manual_data()
        if not data.get("Patient Name"):
            self._preview_status.setText(
                "Enter a Patient Name to enable preview.")
            return
        self._preview_status.setText("Generating preview…")
        self._preview_worker = PreviewWorker(
            data, self._tmp_pdf,
            with_logo=getattr(self, "_manual_with_logo", None) and
                      self._manual_with_logo.isChecked())
        self._preview_worker.finished.connect(self._on_preview_ready)
        self._preview_worker.error.connect(
            lambda e: self._preview_status.setText(
                f"Preview error: {e.splitlines()[0]}"))
        self._preview_worker.start()

    def _on_preview_ready(self, pdf_path: str):
        if not PYPDFIUM_OK:
            self._preview_status.setText(
                "PDF preview unavailable — pypdfium2 not installed.")
            return
        try:
            from io import BytesIO
            while self._preview_vbox.count():
                item = self._preview_vbox.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            doc = _pdfium.PdfDocument(pdf_path)
            # Target display width: fill preview panel (accounts for scroll bar + padding)
            target_w = max(self._preview_inner.width() - 24, 640)
            for page_idx in range(len(doc)):
                bm  = doc[page_idx].render(scale=2.5)
                pil = bm.to_pil()
                buf = BytesIO()
                pil.save(buf, format="PNG")
                buf.seek(0)
                px = QPixmap()
                px.loadFromData(buf.read())
                # Scale to fill panel width (smooth downscale from high-res render)
                px = px.scaledToWidth(target_w,
                    Qt.TransformationMode.SmoothTransformation)
                lbl = QLabel()
                lbl.setPixmap(px)
                lbl.setAlignment(Qt.AlignmentFlag.AlignHCenter)
                lbl.setStyleSheet("border:1px solid #ccc;margin:4px 0;background:white;")
                self._preview_vbox.addWidget(lbl)
            doc.close()
            self._preview_status.setText(
                f"Preview updated  ({datetime.now().strftime('%H:%M:%S')})")
        except Exception as e:
            self._preview_status.setText(f"Render error: {e}")

    # ── Manual data helpers ────────────────────────────────────────────────────
    def _get_manual_data(self) -> dict:
        d = {}
        for key, (w, wtype) in self._manual_inputs.items():
            d[key] = w.currentText() if wtype == "combo" else w.text()
        return d

    def _set_manual_data(self, data: dict):
        for key, (w, wtype) in self._manual_inputs.items():
            val = _clean(data.get(key, ""))
            if wtype == "combo":
                matched = False
                for i in range(w.count()):
                    if w.itemText(i).lower() == val.lower():
                        w.setCurrentIndex(i)
                        matched = True
                        break
                if not matched:
                    for i in range(w.count()):
                        item = w.itemText(i).lower()
                        if val.lower().startswith(item) or item.startswith(val.lower()):
                            w.setCurrentIndex(i)
                            break
            else:
                w.setText(val)

    def _manual_browse_output(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self._manual_out_lbl.setText(folder)
            self._manual_out_lbl.setStyleSheet("padding:4px;color:black;")

    def _manual_generate(self):
        out_dir = self._manual_out_lbl.text()
        if not out_dir or out_dir == "No folder selected":
            QMessageBox.warning(self, "No Folder",
                                "Please select an output folder first.")
            return
        data = self._get_manual_data()
        if not data.get("Patient Name"):
            QMessageBox.warning(self, "Missing Data",
                                "Patient Name is required.")
            return
        try:
            gen  = TERAReportGenerator(data, out_dir,
                                       with_logo=self._manual_with_logo.isChecked())
            path = gen.generate()
            box  = QMessageBox(self)
            box.setWindowTitle("Report Generated")
            box.setIcon(QMessageBox.Icon.Information)
            box.setText(
                f"Report generated successfully.\n\n"
                f"File: {os.path.basename(path)}\n"
                f"Folder: {out_dir}"
            )
            btn_open = box.addButton("Open Folder",
                                     QMessageBox.ButtonRole.ActionRole)
            box.addButton(QMessageBox.StandardButton.Ok)
            box.exec()
            if box.clickedButton() == btn_open:
                _open_folder(out_dir)
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "Error",
                                 f"Failed to generate report:\n{e}")

    def _manual_save_draft(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Draft", "tera_manual_draft.json", "JSON (*.json)")
        if not path:
            return
        try:
            with open(path, "w") as f:
                json.dump(self._get_manual_data(), f, indent=2,
                          default=str)
            self.statusBar().showMessage(f"Draft saved: {path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not save draft:\n{e}")

    def _manual_load_draft(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Load Draft", "", "JSON (*.json)")
        if not path:
            return
        try:
            with open(path) as f:
                data = json.load(f)
            self._set_manual_data(data)
            self.statusBar().showMessage(f"Draft loaded: {path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not load draft:\n{e}")

    def _manual_clear(self):
        for key, (w, wtype) in self._manual_inputs.items():
            if wtype == "combo":
                w.setCurrentIndex(0)
            else:
                # Restore default
                default = next(
                    (opts for _, k, _, opts in TERA_FIELD_DEFS
                     if k == key and isinstance(opts, str)), "")
                w.setText(default)

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 2 – Bulk Upload
    # ═══════════════════════════════════════════════════════════════════════════
    def _create_bulk_tab(self) -> QWidget:
        tab  = QWidget()
        vbox = QVBoxLayout(tab)

        # ── 1. Excel file ──────────────────────────────────────────────────────
        file_grp = QGroupBox("1. Load Excel File")
        file_row = QHBoxLayout(file_grp)
        self._bulk_file_lbl = QLabel("No file loaded")
        self._bulk_file_lbl.setStyleSheet("color:gray;font-style:italic;padding:2px;")
        btn_browse = QPushButton("Browse…")
        btn_browse.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogStart))
        btn_browse.clicked.connect(self._bulk_load_excel)
        file_row.addWidget(self._bulk_file_lbl, 1)
        file_row.addWidget(btn_browse)
        vbox.addWidget(file_grp)

        # ── 2. Output folder ───────────────────────────────────────────────────
        out_grp = QGroupBox("2. Output Folder")
        out_row = QHBoxLayout(out_grp)
        self._bulk_out_lbl = QLabel("No folder selected")
        self._bulk_out_lbl.setStyleSheet("color:gray;font-style:italic;padding:2px;")
        btn_out = QPushButton("Browse…")
        btn_out.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon))
        btn_out.clicked.connect(self._bulk_browse_output)
        out_row.addWidget(self._bulk_out_lbl, 1)
        out_row.addWidget(btn_out)
        vbox.addWidget(out_grp)

        # ── 3. Three-panel content area ────────────────────────────────────────
        data_grp    = QGroupBox("3. Review & Edit Patients")
        data_layout = QVBoxLayout(data_grp)

        # Toolbar above the panels
        toolbar = QHBoxLayout()
        for text, slot in [("Select All",    self._bulk_select_all),
                            ("Deselect All", self._bulk_deselect_all)]:
            b = QPushButton(text)
            b.clicked.connect(slot)
            toolbar.addWidget(b)
        toolbar.addStretch()
        for text, slot in [("Save Draft", self._bulk_save_draft),
                            ("Load Draft", self._bulk_load_draft)]:
            b = QPushButton(text)
            b.clicked.connect(slot)
            toolbar.addWidget(b)
        data_layout.addLayout(toolbar)

        # ── Search bar ────────────────────────────────────────────────────────
        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("🔍 Search:"))
        self._bulk_search = QLineEdit()
        self._bulk_search.setPlaceholderText("Search patient name…")
        self._bulk_search.setClearButtonEnabled(True)
        self._bulk_search.textChanged.connect(self._bulk_filter_table)
        search_row.addWidget(self._bulk_search, 1)
        data_layout.addLayout(search_row)

        main_splitter = QSplitter(Qt.Orientation.Horizontal)

        # ── LEFT panel: patient table ──────────────────────────────────────────
        self._bulk_table = QTableWidget()
        self._bulk_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self._bulk_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._bulk_table.verticalHeader().setVisible(False)
        self._bulk_table.horizontalHeader().setStretchLastSection(True)
        self._bulk_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents)
        self._bulk_table.setAlternatingRowColors(True)
        self._bulk_table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        self._bulk_table.itemSelectionChanged.connect(self._bulk_on_row_selected)
        main_splitter.addWidget(self._bulk_table)

        # ── MIDDLE panel: inline editor ────────────────────────────────────────
        editor_grp  = QGroupBox("Patient Editor")
        editor_vbox = QVBoxLayout(editor_grp)

        self._bulk_editor_placeholder = QLabel("Click a row in the table to edit")
        self._bulk_editor_placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._bulk_editor_placeholder.setStyleSheet("color:gray;font-style:italic;padding:20px;")
        editor_vbox.addWidget(self._bulk_editor_placeholder)

        editor_scroll = QScrollArea()
        editor_scroll.setWidgetResizable(True)
        editor_inner = QWidget()
        form = QFormLayout(editor_inner)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        editor_scroll.setWidget(editor_inner)
        editor_vbox.addWidget(editor_scroll, 1)
        self._bulk_editor_scroll = editor_scroll

        for display_lbl, key, wtype, opts in TERA_FIELD_DEFS:
            if wtype == "combo":
                w = QComboBox()
                w.addItems(opts)
                w.currentTextChanged.connect(self._bulk_schedule_preview)
            else:
                default = opts if isinstance(opts, str) else ""
                w = QLineEdit(default)
                w.textChanged.connect(self._bulk_schedule_preview)
            form.addRow(f"{display_lbl}:", w)
            self._bulk_editor_inputs[key] = (w, wtype)

        editor_scroll.setVisible(False)

        editor_btn_row = QHBoxLayout()

        save_row_btn = QPushButton("Save Changes")
        save_row_btn.setStyleSheet(
            "background-color:#1F497D;color:white;font-weight:bold;padding:6px;")
        save_row_btn.clicked.connect(self._bulk_save_current_row)
        save_row_btn.setVisible(False)
        editor_btn_row.addWidget(save_row_btn)
        self._bulk_save_row_btn = save_row_btn

        save_draft_btn = QPushButton("Save as Draft")
        save_draft_btn.setStyleSheet(
            "background-color:#6C757D;color:white;font-weight:bold;padding:6px;")
        save_draft_btn.clicked.connect(self._bulk_save_individual_draft)
        save_draft_btn.setVisible(False)
        editor_btn_row.addWidget(save_draft_btn)
        self._bulk_save_draft_btn = save_draft_btn

        editor_vbox.addLayout(editor_btn_row)

        main_splitter.addWidget(editor_grp)

        # ── RIGHT panel: live PDF preview ──────────────────────────────────────
        prev_grp  = QGroupBox("Live Preview")
        prev_vbox = QVBoxLayout(prev_grp)

        prev_top = QHBoxLayout()
        self._bulk_preview_status = QLabel("Select a row to preview")
        self._bulk_preview_status.setStyleSheet("color:gray;font-style:italic;")
        self._bulk_preview_status.setWordWrap(True)
        prev_top.addWidget(self._bulk_preview_status, 1)
        btn_bulk_refresh = QPushButton("Refresh")
        btn_bulk_refresh.setIcon(self.style().standardIcon(
            QStyle.StandardPixmap.SP_BrowserReload))
        btn_bulk_refresh.clicked.connect(self._bulk_run_preview)
        prev_top.addWidget(btn_bulk_refresh)
        prev_vbox.addLayout(prev_top)

        prev_scroll = QScrollArea()
        prev_scroll.setWidgetResizable(True)
        self._bulk_preview_inner = QWidget()
        self._bulk_preview_vbox  = QVBoxLayout(self._bulk_preview_inner)
        self._bulk_preview_vbox.setAlignment(
            Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)
        prev_scroll.setWidget(self._bulk_preview_inner)
        prev_vbox.addWidget(prev_scroll, 1)

        main_splitter.addWidget(prev_grp)
        main_splitter.setSizes([280, 340, 600])

        data_layout.addWidget(main_splitter, 1)
        vbox.addWidget(data_grp, 1)

        # ── 4. Generate ────────────────────────────────────────────────────────
        gen_grp    = QGroupBox("4. Generate Reports")
        gen_layout = QVBoxLayout(gen_grp)

        # Logo toggle
        bulk_logo_row = QHBoxLayout()
        bulk_logo_row.addWidget(QLabel("Export:"))
        self._bulk_logo_grp  = QButtonGroup(self)
        self._bulk_no_logo   = QRadioButton("Without Logo")
        self._bulk_with_logo = QRadioButton("With Logo")
        self._bulk_with_logo.setChecked(True)
        self._bulk_logo_grp.addButton(self._bulk_no_logo)
        self._bulk_logo_grp.addButton(self._bulk_with_logo)
        bulk_logo_row.addWidget(self._bulk_no_logo)
        bulk_logo_row.addWidget(self._bulk_with_logo)
        bulk_logo_row.addStretch()
        gen_layout.addLayout(bulk_logo_row)

        act_row = QHBoxLayout()
        self._bulk_gen_sel_btn = QPushButton("Generate Selected")
        self._bulk_gen_sel_btn.setStyleSheet(
            "background-color:#1F497D;color:white;font-weight:bold;padding:8px;")
        self._bulk_gen_sel_btn.setEnabled(False)
        self._bulk_gen_sel_btn.clicked.connect(self._bulk_generate_selected)

        self._bulk_gen_all_btn = QPushButton("Generate All")
        self._bulk_gen_all_btn.setStyleSheet(
            "background-color:#27AE60;color:white;font-weight:bold;padding:8px;")
        self._bulk_gen_all_btn.setEnabled(False)
        self._bulk_gen_all_btn.clicked.connect(self._bulk_generate_all)

        act_row.addWidget(self._bulk_gen_sel_btn)
        act_row.addWidget(self._bulk_gen_all_btn)
        act_row.addStretch()
        gen_layout.addLayout(act_row)

        self._bulk_prog_lbl = QLabel("")
        self._bulk_prog_lbl.setVisible(False)
        self._bulk_progress = QProgressBar()
        self._bulk_progress.setVisible(False)
        gen_layout.addWidget(self._bulk_prog_lbl)
        gen_layout.addWidget(self._bulk_progress)
        vbox.addWidget(gen_grp)

        self._bulk_show_cols = []
        return tab

    # ── Bulk helpers ───────────────────────────────────────────────────────────
    def _bulk_load_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open TERA Excel File", "",
            "Excel Files (*.xls *.xlsx)")
        if not path:
            return
        try:
            df = pd.read_excel(path)
            df = df.dropna(how="all")
            df.columns = df.columns.str.strip()
            self.bulk_rows = df.to_dict(orient="records")
            self._bulk_file_lbl.setText(os.path.basename(path))
            self._bulk_file_lbl.setStyleSheet("color:black;padding:2px;")
            self._populate_bulk_table()
            self._bulk_gen_sel_btn.setEnabled(True)
            self._bulk_gen_all_btn.setEnabled(True)
            self.statusBar().showMessage(
                f"Loaded {len(self.bulk_rows)} rows from "
                f"{os.path.basename(path)}")
        except Exception as e:
            QMessageBox.critical(self, "Load Error",
                                 f"Failed to load Excel:\n{e}")

    def _populate_bulk_table(self):
        if not self.bulk_rows:
            return
        all_keys = list(self.bulk_rows[0].keys())
        cols = [c for c in BULK_DISPLAY_COLS if c in all_keys]
        if not cols:
            cols = all_keys[:7]
        self._bulk_show_cols = cols

        self._bulk_table.setColumnCount(len(cols))
        self._bulk_table.setHorizontalHeaderLabels(cols)
        self._bulk_table.setRowCount(len(self.bulk_rows))

        for r_idx, row in enumerate(self.bulk_rows):
            for c_idx, col in enumerate(cols):
                val = _clean(row.get(col, ""))
                self._bulk_table.setItem(
                    r_idx, c_idx, QTableWidgetItem(val))

        # Auto-select and populate editor for first row
        self._bulk_current_row = -1
        self._bulk_table.selectRow(0)

    def _bulk_filter_table(self):
        """Show/hide rows based on the search bar text."""
        text = self._bulk_search.text().strip().lower()
        tbl = self._bulk_table
        # Patient Name is always column 1
        for row in range(tbl.rowCount()):
            item = tbl.item(row, 1)
            name = item.text().lower() if item else ""
            tbl.setRowHidden(row, bool(text) and text not in name)

    def _bulk_select_all(self):
        """Select all visible (non-hidden) rows."""
        model = self._bulk_table.selectionModel()
        model.clearSelection()
        for row in range(self._bulk_table.rowCount()):
            if not self._bulk_table.isRowHidden(row):
                idx = self._bulk_table.model().index(row, 0)
                model.select(idx, QItemSelectionModel.SelectionFlag.Select |
                              QItemSelectionModel.SelectionFlag.Rows)

    def _bulk_deselect_all(self):
        self._bulk_table.clearSelection()

    def _bulk_on_row_selected(self):
        """Single-click on a table row → populate inline editor + trigger preview."""
        row_idx = self._bulk_table.currentRow()
        if row_idx < 0 or row_idx >= len(self.bulk_rows):
            return
        self._bulk_populate_editor(row_idx)
        self._bulk_schedule_preview()

    def _bulk_populate_editor(self, row_idx: int):
        """Fill the inline editor widgets from bulk_rows[row_idx]."""
        self._bulk_current_row = row_idx
        data = self.bulk_rows[row_idx]
        # Show editor, hide placeholder
        self._bulk_editor_placeholder.setVisible(False)
        self._bulk_editor_scroll.setVisible(True)
        self._bulk_save_row_btn.setVisible(True)
        self._bulk_save_draft_btn.setVisible(True)
        # Block signals to avoid cascading previews while filling fields
        for key, (w, wtype) in self._bulk_editor_inputs.items():
            w.blockSignals(True)
        for key, (w, wtype) in self._bulk_editor_inputs.items():
            val = _clean(data.get(key, ""))
            if wtype == "combo":
                # Exact match first (case-insensitive) to avoid "receptive"
                # matching "post-receptive" or "pre-receptive" by substring.
                matched = False
                for i in range(w.count()):
                    if w.itemText(i).lower() == val.lower():
                        w.setCurrentIndex(i)
                        matched = True
                        break
                if not matched:
                    # Fallback: item text starts with value or value starts with item text
                    # (avoids "Natural" matching "Modified Natural Cycle")
                    for i in range(w.count()):
                        item = w.itemText(i).lower()
                        if val.lower().startswith(item) or item.startswith(val.lower()):
                            w.setCurrentIndex(i)
                            matched = True
                            break
                if not matched:
                    w.setCurrentIndex(0)
            else:
                w.setText(val)
        for key, (w, wtype) in self._bulk_editor_inputs.items():
            w.blockSignals(False)

    def _bulk_save_current_row(self):
        """Save inline editor values back to bulk_rows and refresh the table row."""
        if self._bulk_current_row < 0 or self._bulk_current_row >= len(self.bulk_rows):
            return
        row_idx = self._bulk_current_row
        d = dict(self.bulk_rows[row_idx])
        for key, (w, wtype) in self._bulk_editor_inputs.items():
            d[key] = w.currentText() if wtype == "combo" else w.text()
        self.bulk_rows[row_idx] = d
        for c_idx, col in enumerate(self._bulk_show_cols):
            val = _clean(d.get(col, ""))
            self._bulk_table.setItem(row_idx, c_idx, QTableWidgetItem(val))
        self._bulk_filter_table()   # re-apply search filter in case name changed
        self.statusBar().showMessage(f"Row {row_idx + 1} saved")
        self._bulk_run_preview()

    def _bulk_save_individual_draft(self):
        """Save the current patient's editor data as a standalone JSON draft file."""
        if self._bulk_current_row < 0 or self._bulk_current_row >= len(self.bulk_rows):
            return
        # Pull latest values from editor widgets first
        d = dict(self.bulk_rows[self._bulk_current_row])
        for key, (w, wtype) in self._bulk_editor_inputs.items():
            d[key] = w.currentText() if wtype == "combo" else w.text()
        # Suggest a filename based on patient name
        name = _clean(d.get("Patient Name", "patient")).replace(" ", "_")
        default = os.path.join(
            self._bulk_out_lbl.text()
            if self._bulk_out_lbl.text() != "No folder selected" else "",
            f"{name}_draft.json"
        )
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Individual Draft", default, "JSON (*.json)")
        if not path:
            return
        try:
            # Convert non-serialisable values to strings
            serialisable = {k: str(v) if not isinstance(v, (str, int, float, bool, type(None)))
                            else v for k, v in d.items()}
            with open(path, "w", encoding="utf-8") as f:
                json.dump(serialisable, f, indent=2, ensure_ascii=False)
            self.statusBar().showMessage(f"Draft saved: {os.path.basename(path)}")
        except Exception as e:
            QMessageBox.warning(self, "Save Error", f"Could not save draft:\n{e}")

    def _bulk_schedule_preview(self):
        if self._bulk_current_row >= 0:
            self._bulk_preview_timer.start()

    def _bulk_run_preview(self):
        if self._bulk_current_row < 0 or self._bulk_current_row >= len(self.bulk_rows):
            return
        if self._bulk_preview_worker and self._bulk_preview_worker.isRunning():
            return
        # Use live editor values (even before Save is clicked)
        d = dict(self.bulk_rows[self._bulk_current_row])
        for key, (w, wtype) in self._bulk_editor_inputs.items():
            d[key] = w.currentText() if wtype == "combo" else w.text()
        if not _clean(d.get("Patient Name", "")):
            self._bulk_preview_status.setText("Enter a Patient Name to enable preview.")
            return
        self._bulk_preview_status.setText("Generating preview…")
        self._bulk_preview_worker = PreviewWorker(
            d, self._bulk_tmp_pdf,
            with_logo=self._bulk_with_logo.isChecked())
        self._bulk_preview_worker.finished.connect(self._bulk_on_preview_ready)
        self._bulk_preview_worker.error.connect(
            lambda e: self._bulk_preview_status.setText(f"Preview error: {e}"))
        self._bulk_preview_worker.start()

    def _bulk_on_preview_ready(self, pdf_path: str):
        if not PYPDFIUM_OK:
            self._bulk_preview_status.setText(
                "PDF preview unavailable — pypdfium2 not installed.")
            return
        try:
            from io import BytesIO
            while self._bulk_preview_vbox.count():
                item = self._bulk_preview_vbox.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            doc = _pdfium.PdfDocument(pdf_path)
            target_w = max(self._bulk_preview_inner.width() - 24, 560)
            for page_idx in range(len(doc)):
                bm  = doc[page_idx].render(scale=2.5)
                pil = bm.to_pil()
                buf = BytesIO()
                pil.save(buf, format="PNG")
                buf.seek(0)
                px = QPixmap()
                px.loadFromData(buf.read())
                px = px.scaledToWidth(target_w,
                    Qt.TransformationMode.SmoothTransformation)
                lbl = QLabel()
                lbl.setPixmap(px)
                lbl.setAlignment(Qt.AlignmentFlag.AlignHCenter)
                lbl.setStyleSheet("border:1px solid #ccc;margin:4px 0;background:white;")
                self._bulk_preview_vbox.addWidget(lbl)
            doc.close()
            self._bulk_preview_status.setText(
                f"Preview updated  ({datetime.now().strftime('%H:%M:%S')})")
        except Exception as e:
            self._bulk_preview_status.setText(f"Render error: {e}")

    def _bulk_browse_output(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Select Output Folder")
        if folder:
            self._bulk_out_lbl.setText(folder)
            self._bulk_out_lbl.setStyleSheet("color:black;padding:2px;")

    def _bulk_generate_selected(self):
        ranges = self._bulk_table.selectedRanges()
        if not ranges:
            QMessageBox.warning(self, "No Selection",
                                "Select rows to generate.")
            return
        rows = []
        for rng in ranges:
            for r in range(rng.topRow(), rng.bottomRow() + 1):
                if r < len(self.bulk_rows) and not self._bulk_table.isRowHidden(r):
                    rows.append(self.bulk_rows[r])
        self._start_bulk_gen(rows)

    def _bulk_generate_all(self):
        if not self.bulk_rows:
            QMessageBox.warning(self, "No Data",
                                "Load an Excel file first.")
            return
        self._start_bulk_gen(self.bulk_rows)

    def _start_bulk_gen(self, rows: list):
        out_dir = self._bulk_out_lbl.text()
        if not out_dir or out_dir == "No folder selected":
            out_dir = QFileDialog.getExistingDirectory(
                self, "Select Output Folder")
            if not out_dir:
                return
            self._bulk_out_lbl.setText(out_dir)
            self._bulk_out_lbl.setStyleSheet("color:black;padding:2px;")

        self._bulk_progress.setValue(0)
        self._bulk_progress.setVisible(True)
        self._bulk_prog_lbl.setVisible(True)
        self._bulk_gen_sel_btn.setEnabled(False)
        self._bulk_gen_all_btn.setEnabled(False)

        self._gen_worker = ReportGeneratorWorker(
            rows, out_dir, with_logo=self._bulk_with_logo.isChecked())
        self._gen_worker.progress.connect(self._on_bulk_progress)
        self._gen_worker.finished.connect(self._on_bulk_finished)
        self._gen_worker.start()

    def _on_bulk_progress(self, pct: int, msg: str):
        self._bulk_progress.setValue(pct)
        self._bulk_prog_lbl.setText(msg)

    def _on_bulk_finished(self, successes: int, errors: list):
        self._bulk_progress.setVisible(False)
        self._bulk_prog_lbl.setVisible(False)
        self._bulk_gen_sel_btn.setEnabled(True)
        self._bulk_gen_all_btn.setEnabled(True)

        if errors:
            box = QMessageBox(self)
            box.setWindowTitle("Generation Complete")
            box.setIcon(QMessageBox.Icon.Warning)
            box.setText(f"{successes} report(s) generated, "
                        f"{len(errors)} failed.")
            box.setDetailedText("\n\n".join(errors))
            box.exec()
        else:
            out_dir = self._bulk_out_lbl.text()
            box = QMessageBox(self)
            box.setWindowTitle("Success")
            box.setIcon(QMessageBox.Icon.Information)
            box.setText(
                f"{successes} report(s) generated successfully.\n\n"
                f"Saved to:\n{out_dir}"
            )
            btn_open = box.addButton("Open Folder",
                                     QMessageBox.ButtonRole.ActionRole)
            box.addButton(QMessageBox.StandardButton.Ok)
            box.exec()
            if box.clickedButton() == btn_open:
                _open_folder(out_dir)

    def _bulk_save_draft(self):
        if not self.bulk_rows:
            QMessageBox.warning(self, "No Data", "No data to save.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Bulk Draft", "tera_bulk_draft.json",
            "JSON (*.json)")
        if not path:
            return
        try:
            serial = [{k: _clean(v) for k, v in row.items()}
                      for row in self.bulk_rows]
            with open(path, "w") as f:
                json.dump(serial, f, indent=2)
            self.statusBar().showMessage(f"Bulk draft saved: {path}")
        except Exception as e:
            QMessageBox.critical(self, "Error",
                                 f"Could not save draft:\n{e}")

    def _bulk_load_draft(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Load Bulk Draft", "", "JSON (*.json)")
        if not path:
            return
        try:
            with open(path) as f:
                data = json.load(f)
            self.bulk_rows = data
            self._bulk_file_lbl.setText(f"Draft: {os.path.basename(path)}")
            self._bulk_file_lbl.setStyleSheet(
                "color:#1F497D;padding:2px;")
            self._populate_bulk_table()
            self._bulk_gen_sel_btn.setEnabled(True)
            self._bulk_gen_all_btn.setEnabled(True)
            self.statusBar().showMessage(
                f"Bulk draft loaded: {len(data)} rows")
        except Exception as e:
            QMessageBox.critical(self, "Error",
                                 f"Could not load draft:\n{e}")

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 3 – Report Comparison
    # ═══════════════════════════════════════════════════════════════════════════
    def _create_comparison_tab(self) -> QWidget:
        tab  = QWidget()
        vbox = QVBoxLayout(tab)

        # Info banner
        info = QLabel(
            "Compare two TERA PDFs — auto-generated (left) vs reference (right). "
            "Click <b>Compare PDFs</b> for a visual side-by-side view AND an "
            "automated text-diff summary that highlights every discrepancy."
        )
        info.setWordWrap(True)
        info.setStyleSheet(
            "color:#444;padding:8px;background:#EAF4FF;"
            "border:1px solid #BDD7EE;border-radius:4px;")
        vbox.addWidget(info)

        # File selectors
        sel_grp    = QGroupBox("Select PDFs")
        sel_layout = QHBoxLayout(sel_grp)

        left_col = QVBoxLayout()
        left_col.addWidget(QLabel("<b>Left PDF  (Auto-generated)</b>"))
        self._cmp_left_lbl = QLabel("No file selected")
        self._cmp_left_lbl.setStyleSheet(
            "color:gray;font-style:italic;"
            "border:1px solid #ccc;padding:4px;")
        btn_left = QPushButton("Select PDF…")
        btn_left.clicked.connect(self._cmp_pick_left)
        left_col.addWidget(self._cmp_left_lbl)
        left_col.addWidget(btn_left)
        sel_layout.addLayout(left_col)

        line = QFrame()
        line.setFrameShape(QFrame.Shape.VLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        sel_layout.addWidget(line)

        right_col = QVBoxLayout()
        right_col.addWidget(QLabel("<b>Right PDF  (Reference / Manual)</b>"))
        self._cmp_right_lbl = QLabel("No file selected")
        self._cmp_right_lbl.setStyleSheet(
            "color:gray;font-style:italic;"
            "border:1px solid #ccc;padding:4px;")
        btn_right = QPushButton("Select PDF…")
        btn_right.clicked.connect(self._cmp_pick_right)
        right_col.addWidget(self._cmp_right_lbl)
        right_col.addWidget(btn_right)
        sel_layout.addLayout(right_col)

        cmp_btn_col = QVBoxLayout()
        btn_compare = QPushButton("  Compare PDFs  ")
        btn_compare.setStyleSheet(
            "background-color:#1F497D;color:white;"
            "font-weight:bold;padding:10px 16px;")
        btn_compare.clicked.connect(self._cmp_run)
        self._cmp_status_lbl = QLabel("")
        self._cmp_status_lbl.setStyleSheet("color:gray;font-style:italic;")
        cmp_btn_col.addWidget(btn_compare)
        cmp_btn_col.addWidget(self._cmp_status_lbl)
        sel_layout.addLayout(cmp_btn_col)

        vbox.addWidget(sel_grp)

        # Main splitter: left = visual viewer, right = diff summary
        main_splitter = QSplitter(Qt.Orientation.Horizontal)

        # ── Left: side-by-side visual viewer ──────────────────────────────────
        viewer_widget = QWidget()
        viewer_vbox   = QVBoxLayout(viewer_widget)
        viewer_vbox.setContentsMargins(0, 0, 0, 0)

        nav = QHBoxLayout()
        self._cmp_prev_btn = QPushButton("◀  Prev Page")
        self._cmp_next_btn = QPushButton("Next Page  ▶")
        self._cmp_page_lbl = QLabel("—")
        self._cmp_page_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._cmp_page_lbl.setMinimumWidth(100)
        self._cmp_prev_btn.clicked.connect(self._cmp_prev)
        self._cmp_next_btn.clicked.connect(self._cmp_next)
        for w in (self._cmp_prev_btn, self._cmp_page_lbl, self._cmp_next_btn):
            nav.addWidget(w)
        nav.insertStretch(0)
        nav.addStretch()
        viewer_vbox.addLayout(nav)

        pages_splitter        = QSplitter(Qt.Orientation.Horizontal)
        self._cmp_left_label  = self._make_cmp_label("Select a PDF on the left")
        self._cmp_right_label = self._make_cmp_label("Select a PDF on the right")
        pages_splitter.addWidget(self._cmp_left_label)
        pages_splitter.addWidget(self._cmp_right_label)
        viewer_vbox.addWidget(pages_splitter, 1)

        main_splitter.addWidget(viewer_widget)

        # ── Right: auto diff summary ───────────────────────────────────────────
        diff_grp  = QGroupBox("Auto Diff Summary")
        diff_vbox = QVBoxLayout(diff_grp)

        self._cmp_diff_browser = QTextBrowser()
        self._cmp_diff_browser.setOpenExternalLinks(False)
        self._cmp_diff_browser.setHtml(
            "<p style='color:gray;font-style:italic;padding:10px;'>"
            "Click <b>Compare PDFs</b> to run the automatic analysis.</p>")
        diff_vbox.addWidget(self._cmp_diff_browser)

        btn_export_diff = QPushButton("Export Diff as HTML…")
        btn_export_diff.clicked.connect(self._cmp_export_diff)
        diff_vbox.addWidget(btn_export_diff)

        main_splitter.addWidget(diff_grp)
        main_splitter.setSizes([700, 500])

        vbox.addWidget(main_splitter, 1)

        self._cmp_left_path   = None
        self._cmp_right_path  = None
        self._cmp_page        = 0
        self._cmp_total_pages = 0
        self._cmp_diff_html   = ""
        self._cmp_diff_worker = None
        return tab

    @staticmethod
    def _make_cmp_label(placeholder: str) -> QLabel:
        lbl = QLabel(placeholder)
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl.setStyleSheet(
            "background:#f5f5f5;border:1px solid #ddd;")
        lbl.setSizePolicy(QSizePolicy.Policy.Expanding,
                          QSizePolicy.Policy.Expanding)
        lbl.setMinimumHeight(400)
        return lbl

    def _cmp_pick_left(self):
        p, _ = QFileDialog.getOpenFileName(
            self, "Select Left PDF", "", "PDF (*.pdf)")
        if p:
            self._cmp_left_path = p
            self._cmp_left_lbl.setText(os.path.basename(p))
            self._cmp_left_lbl.setStyleSheet(
                "color:black;border:1px solid #ccc;padding:4px;")

    def _cmp_pick_right(self):
        p, _ = QFileDialog.getOpenFileName(
            self, "Select Right PDF", "", "PDF (*.pdf)")
        if p:
            self._cmp_right_path = p
            self._cmp_right_lbl.setText(os.path.basename(p))
            self._cmp_right_lbl.setStyleSheet(
                "color:black;border:1px solid #ccc;padding:4px;")

    def _cmp_run(self):
        if not self._cmp_left_path or not self._cmp_right_path:
            QMessageBox.warning(self, "Missing PDFs",
                                "Please select both PDFs first.")
            return

        # Visual side-by-side
        if PYPDFIUM_OK:
            self._cmp_page = 0
            self._cmp_render_current()
        else:
            self._cmp_left_label.setText(
                "pypdfium2 not installed — visual preview unavailable.")
            self._cmp_right_label.setText("")

        # Auto-diff
        if self._cmp_diff_worker and self._cmp_diff_worker.isRunning():
            return
        self._cmp_status_lbl.setText("Running auto-diff…")
        self._cmp_diff_browser.setHtml(
            "<p style='color:gray;font-style:italic;padding:10px;'>"
            "Analysing…</p>")
        self._cmp_diff_worker = PDFDiffWorker(
            self._cmp_left_path, self._cmp_right_path)
        self._cmp_diff_worker.finished.connect(self._cmp_on_diff_done)
        self._cmp_diff_worker.error.connect(self._cmp_on_diff_error)
        self._cmp_diff_worker.start()

    def _cmp_on_diff_done(self, html: str):
        self._cmp_diff_html = html
        self._cmp_diff_browser.setHtml(html)
        self._cmp_status_lbl.setText("Diff complete.")

    def _cmp_on_diff_error(self, msg: str):
        self._cmp_status_lbl.setText("Diff error — see summary panel.")
        self._cmp_diff_browser.setHtml(
            f"<pre style='color:red;padding:10px;'>{msg}</pre>")

    def _cmp_export_diff(self):
        if not self._cmp_diff_html:
            QMessageBox.information(self, "No Diff", "Run a comparison first.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Export Diff", "tera_diff_report.html", "HTML (*.html)")
        if path:
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(self._cmp_diff_html)
                self.statusBar().showMessage(f"Diff exported: {path}")
            except Exception as e:
                QMessageBox.warning(self, "Export Error", str(e))

    def _cmp_render_current(self):
        if not PYPDFIUM_OK:
            return
        from io import BytesIO

        def load_page(pdf_path, page_idx) -> tuple:
            """Returns (QPixmap, page_count)."""
            doc   = _pdfium.PdfDocument(pdf_path)
            n     = len(doc)
            idx   = min(page_idx, n - 1)
            page  = doc[idx]
            bm    = page.render(scale=1.35)
            pil   = bm.to_pil()
            doc.close()
            buf = BytesIO()
            pil.save(buf, format="PNG")
            buf.seek(0)
            px = QPixmap()
            px.loadFromData(buf.read())
            return px, n

        try:
            px_l, n_l = load_page(self._cmp_left_path,  self._cmp_page)
            px_r, n_r = load_page(self._cmp_right_path, self._cmp_page)
        except Exception as e:
            QMessageBox.critical(self, "Render Error",
                                 f"Failed to render PDFs:\n{e}")
            return

        self._cmp_total_pages = max(n_l, n_r)
        self._cmp_page_lbl.setText(
            f"Page {self._cmp_page + 1} / {self._cmp_total_pages}")

        def put_pixmap(lbl, px):
            scaled = px.scaled(
                lbl.width() - 8, lbl.height() - 8,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation)
            lbl.setPixmap(scaled)

        put_pixmap(self._cmp_left_label,  px_l)
        put_pixmap(self._cmp_right_label, px_r)

    def _cmp_prev(self):
        if self._cmp_page > 0:
            self._cmp_page -= 1
            self._cmp_render_current()

    def _cmp_next(self):
        if self._cmp_page < self._cmp_total_pages - 1:
            self._cmp_page += 1
            self._cmp_render_current()

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 4 – User Guide
    # ═══════════════════════════════════════════════════════════════════════════
    def _create_guide_tab(self) -> QWidget:
        tab    = QWidget()
        layout = QVBoxLayout(tab)
        guide  = QTextBrowser()
        guide.setOpenExternalLinks(True)
        guide.setHtml(_GUIDE_HTML)
        layout.addWidget(guide)
        return tab

    # ═══════════════════════════════════════════════════════════════════════════
    # Settings persistence
    # ═══════════════════════════════════════════════════════════════════════════
    def _load_settings(self):
        for attr, key in [("_manual_out_lbl", "manual_output_dir"),
                          ("_bulk_out_lbl",   "bulk_output_dir")]:
            val = self.settings.value(key, "")
            if val and os.path.isdir(val):
                getattr(self, attr).setText(val)
                getattr(self, attr).setStyleSheet(
                    "padding:2px;color:black;")

    def closeEvent(self, event):
        for attr, key in [("_manual_out_lbl", "manual_output_dir"),
                          ("_bulk_out_lbl",   "bulk_output_dir")]:
            lbl = getattr(self, attr)
            val = lbl.text()
            self.settings.setValue(
                key,
                val if val not in ("No folder selected", "No folder selected")
                else "")
        super().closeEvent(event)


# ─── User Guide HTML ───────────────────────────────────────────────────────────
_GUIDE_HTML = """
<html>
<head>
<style>
  body  { font-family:'Segoe UI',Arial,sans-serif; line-height:1.65;
          color:#333; background:#f8f9fa; padding:24px; }
  .hdr  { background:#1F497D; color:white; padding:28px;
          border-radius:8px; margin-bottom:22px; text-align:center; }
  .hdr h1 { margin:0; font-size:24px; }
  .hdr p  { margin:5px 0 0; opacity:.85; font-size:14px; }
  .card { background:white; border-radius:6px; padding:18px 22px;
          margin-bottom:16px; border-left:4px solid #1F497D;
          box-shadow:0 1px 4px rgba(0,0,0,.06); }
  .card h3 { color:#1F497D; margin:0 0 10px;
             border-bottom:1px solid #eee; padding-bottom:8px; }
  ul { padding-left:20px; margin:0; }
  li { padding:4px 0; }
  code { background:#f1f1f1; padding:2px 6px; border-radius:3px;
         font-family:monospace; font-size:13px; }
  .tip  { background:#e7f3ff; border:1px solid #b8daff;
          padding:12px 16px; border-radius:4px; color:#004085;
          margin-top:6px; }
</style>
</head>
<body>
<div class="hdr">
  <h1>TERA Report Generator &mdash; User Guide</h1>
  <p>Transcriptome based Endometrial Receptivity Assessment</p>
</div>

<div class="card">
  <h3>1. Manual Entry</h3>
  <ul>
    <li><b>Fill in the form</b> on the left. Fields auto-trigger a live PDF
        preview after a short pause.</li>
    <li>Click <b>Refresh Preview</b> to immediately render all 3 report pages
        in the right panel (requires pypdfium2).</li>
    <li><b>TERA Result</b>: choose Receptive, Pre-receptive, or
        Post-receptive.</li>
    <li><b>Biopsy Date &amp; Time</b>: enter as
        <code>YYYY-MM-DD HH:MM:SS</code> or <code>DD-MM-YYYY HH:MM</code>.</li>
    <li><b>P+ Hours</b>: numeric hours since P4/hCG injection (e.g.
        <code>120</code>).</li>
    <li><b>Time for Report</b>: blastocyst hours &plusmn; margin —
        e.g. <code>144 + 2</code>.
        Cleavage time is computed automatically as
        blastocyst&nbsp;&minus;&nbsp;48&nbsp;hrs.</li>
    <li><b>Save Draft / Load Draft</b>: stores the form as a JSON file
        so you can reload it later.</li>
    <li>Set the output folder then click <b>Generate Report</b> to create
        the PDF.</li>
  </ul>
</div>

<div class="card">
  <h3>2. Bulk Upload</h3>
  <ul>
    <li>Click <b>Browse</b> to load a <code>.xls</code> or
        <code>.xlsx</code> TERA automation report.</li>
    <li>All rows appear in the <b>left patient table</b>. Click any row to
        load it into the <b>inline editor</b> (middle panel) and generate a
        live PDF preview (right panel) automatically.</li>
    <li>Edit any field in the middle panel and click <b>Save Changes</b> to
        commit the edits back to the row list. The preview refreshes instantly.</li>
    <li>Select one or more rows and click <b>Generate Selected</b>, or use
        <b>Generate All</b> to process every row.</li>
    <li><b>Save Draft</b> serialises the (possibly edited) rows to JSON
        for future reuse. <b>Load Draft</b> restores them.</li>
  </ul>
</div>

<div class="card">
  <h3>3. Report Comparison</h3>
  <ul>
    <li>Select two PDFs — auto-generated (left) and a reference / manually
        prepared version (right).</li>
    <li>Click <b>Compare PDFs</b> to render both pages side by side.</li>
    <li>Use <b>Prev Page / Next Page</b> to navigate all 3 report pages.</li>
    <li>Drag the centre splitter to give more space to either side.</li>
  </ul>
</div>

<div class="card">
  <h3>4. Result Types &amp; Key Fields</h3>
  <ul>
    <li><b>Receptive</b>: endometrium is within the window of implantation.</li>
    <li><b>Pre-receptive</b>: biopsy too early; window not yet open.</li>
    <li><b>Post-receptive</b>: biopsy after the window; second biopsy
        strongly recommended.</li>
    <li><b>Time for Report</b> e.g. <code>144 + 2</code>:<br>
        &nbsp;&nbsp;Blastocyst&nbsp;(Day&nbsp;5/6)&nbsp;=&nbsp;P+144&nbsp;&plusmn;&nbsp;2&nbsp;hrs<br>
        &nbsp;&nbsp;Cleavage&nbsp;(Day&nbsp;3)&nbsp;=&nbsp;P+96&nbsp;&plusmn;&nbsp;2&nbsp;hrs
        &nbsp;(blastocyst&nbsp;&minus;&nbsp;48)</li>
  </ul>
</div>

<div class="card">
  <h3>5. System Requirements &amp; Installation</h3>
  <ul>
    <li>Python 3.10+, PyQt6 &ge; 6.6, ReportLab &ge; 4.0</li>
    <li>pandas, xlrd, openpyxl, Pillow, pypdfium2</li>
    <li><b>Step 1 — Install Python packages:</b><br>
        <code>pip install -r requirements.txt</code></li>
    <li><b>Step 2 — Install bundled fonts (run once):</b><br>
        <code>python install_fonts.py</code><br>
        <small>Copies Calibri, SegoeUI, GillSansMT and DengXian fonts to your system font folder
        (<em>Windows:</em> C:\Windows\Fonts &nbsp;|&nbsp;
         <em>Linux:</em> ~/.local/share/fonts/TERA &nbsp;|&nbsp;
         <em>macOS:</em> ~/Library/Fonts).
        On Windows, run as Administrator if prompted.</small></li>
  </ul>
</div>

<div class="tip">
  <b>Tip:</b> The output folder is remembered between sessions.
  Draft files are plain JSON — you can edit them in a text editor
  before reloading.
</div>
</body>
</html>
"""


# ─── Entry point ───────────────────────────────────────────────────────────────
def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Check optional dependencies that may be missing on a fresh install
    _missing = []
    for _pkg in ("xlrd", "openpyxl", "reportlab", "pdfplumber", "pypdfium2", "PIL"):
        try:
            __import__(_pkg)
        except ImportError:
            _missing.append(_pkg)
    if _missing:
        QMessageBox.critical(
            None, "Missing Dependencies",
            "The following packages are not installed:\n\n"
            + "\n".join(f"  \u2022 {m}" for m in _missing)
            + "\n\nRun:  pip install -r requirements.txt\nthen restart the application."
        )
        sys.exit(1)
    _icon_path = _resource_path("tera_icon.png")
    if os.path.exists(_icon_path):
        app.setWindowIcon(QIcon(_icon_path))
    window = TERAReportApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
