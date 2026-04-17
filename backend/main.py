from fastapi import FastAPI, UploadFile, File, Request, BackgroundTasks
from typing import List
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import shutil

from dotenv import load_dotenv
load_dotenv()


import pandas as pd
import os
import io
import uuid
import math
import re
import base64
import difflib
from datetime import datetime

import pdfplumber

from tera_template import TERAReportGenerator
from pgta_template import PGTAReportTemplate
from fastapi.staticfiles import StaticFiles
from pgta_docx_generator import PGTADocxGenerator
from pgta_classify import auto_map_cnvs
try:
    from supabase_client import supabase, upload_pdf, save_report, upload_pgta_file
except ImportError:
    supabase = None
    upload_pdf = None
    save_report = None


app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
FRONTEND_DIR = os.path.join(os.path.dirname(BASE_DIR), "front end")
ROOT_DIR     = os.path.dirname(BASE_DIR)

TEMP_DIR        = os.path.join(BASE_DIR, "temp")
REPORT_DIR      = os.path.join(BASE_DIR, "reports")
PGTA_REPORT_DIR = os.path.join(BASE_DIR, "reports-pgta")

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(REPORT_DIR, exist_ok=True)
os.makedirs(PGTA_REPORT_DIR, exist_ok=True)

PGTA_CNV_DIR   = os.path.join(BASE_DIR, "uploads", "pgta_cnv")
PGTA_DRAFT_DIR = os.path.join(BASE_DIR, "drafts", "PGTA")
os.makedirs(PGTA_CNV_DIR, exist_ok=True)
os.makedirs(PGTA_DRAFT_DIR, exist_ok=True)

app.mount("/reports", StaticFiles(directory=REPORT_DIR), name="reports")
app.mount("/reports-pgta", StaticFiles(directory=PGTA_REPORT_DIR), name="reports-pgta")
app.mount("/pgta-fonts", StaticFiles(directory=os.path.join(BASE_DIR, "assets", "pgta", "fonts")), name="pgta-fonts")
app.mount("/pgta-assets", StaticFiles(directory=os.path.join(BASE_DIR, "assets", "pgta")), name="pgta-assets")

# Serve root-level static assets (logo, icons, images)
if os.path.exists(ROOT_DIR):
    app.mount("/static", StaticFiles(directory=ROOT_DIR), name="static")


def _resolve_cnv_images(embryos: list) -> tuple:
    """
    Convert any base64-encoded CNV images (cnv_image_b64) to temp files and
    populate cnv_image_path so pgta_template.py can find them via os.path.exists().
    Returns (embryos, list_of_temp_paths_to_cleanup).
    """
    temp_paths = []
    for emb in embryos:
        b64 = (emb.get("cnv_image_b64") or "").strip()
        if not b64:
            continue  # already has a path, or no image
        try:
            # Strip data-URL prefix if present (data:image/png;base64,...)
            if "," in b64:
                b64 = b64.split(",", 1)[1]
            img_bytes = base64.b64decode(b64)
            tmp_name  = f"cnv_{uuid.uuid4().hex}.png"
            tmp_path  = os.path.join(TEMP_DIR, tmp_name)
            with open(tmp_path, "wb") as f:
                f.write(img_bytes)
            emb["cnv_image_path"] = tmp_path
            temp_paths.append(tmp_path)
        except Exception as exc:
            print(f"[cnv_resolve] Failed to decode base64 image for embryo "
                  f"'{emb.get('embryo_id','?')}': {exc}")
    return embryos, temp_paths

@app.api_route("/", methods=["GET", "HEAD"])
def root():
    p = os.path.join(FRONTEND_DIR, "login.html")
    if os.path.exists(p):
        return FileResponse(p, media_type="text/html")
    return {"status": "TERA backend running"}

@app.get("/home")
def home():
    p = os.path.join(FRONTEND_DIR, "home page.html")
    if os.path.exists(p):
        return FileResponse(p, media_type="text/html")
    return {"status": "not found"}

@app.get("/dashboard")
def dashboard():
    p = os.path.join(FRONTEND_DIR, "dashboard_copy.html")
    if os.path.exists(p):
        return FileResponse(p, media_type="text/html")
    return {"status": "not found"}


# -------- Preview Report --------
@app.post("/preview")
async def preview_report(data: dict):

    file_id = str(uuid.uuid4()) + ".pdf"
    filepath = os.path.join(TEMP_DIR, file_id)

    with_logo = data.get("logo_option", "without_logo") == "with_logo"
    gen = TERAReportGenerator(data, TEMP_DIR, with_logo=with_logo)
    gen.filepath = filepath
    gen.filename = file_id

    gen.generate()

    return {"preview_url": f"/preview-file/{file_id}"}
    

@app.get("/preview-file/{filename}")
def preview_file(filename: str):

    path = os.path.join(TEMP_DIR, filename)

    return FileResponse(path, media_type="application/pdf")


# -------- Native Folder Picker --------
@app.get("/open-folder-dialog")
async def open_folder_dialog():
    """Open the OS native folder-picker dialog and return the chosen path.
    Tries tkinter first; falls back to PowerShell FolderBrowserDialog on Windows."""
    # Attempt 1: tkinter (works when display is available)
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes("-topmost", True)
        folder = filedialog.askdirectory(title="Select Export Folder", parent=root)
        root.destroy()
        return {"path": folder or ""}
    except Exception:
        pass

    # Attempt 2: PowerShell FolderBrowserDialog (Windows fallback)
    try:
        import subprocess
        ps = (
            "Add-Type -AssemblyName System.Windows.Forms; "
            "$d = New-Object System.Windows.Forms.FolderBrowserDialog; "
            "$d.Description = 'Select Export Folder'; "
            "$d.ShowNewFolderButton = $true; "
            "if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)"
            "{ Write-Output $d.SelectedPath }"
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
            capture_output=True, text=True, timeout=60
        )
        return {"path": result.stdout.strip()}
    except Exception as e:
        return {"error": str(e)}


# -------- Single Report Generation --------
@app.post("/generate")
async def generate_report(data: dict):

    try:
        with_logo = data.get("logo_option", "without_logo") == "with_logo"
        generator = TERAReportGenerator(data, REPORT_DIR, with_logo=with_logo)
        pdf_path = generator.generate()

        # check if file exists
        if not pdf_path or not os.path.exists(pdf_path):
            return {"error": "PDF not generated"}

        file_name = _build_file_name(data, with_logo)

        # Copy to custom output_dir if provided
        custom_dir = (data.get("output_dir") or "").strip()
        if custom_dir:
            try:
                out_dir = custom_dir if os.path.isabs(custom_dir) else os.path.join(BASE_DIR, custom_dir)
                os.makedirs(out_dir, exist_ok=True)
                import shutil as _shutil
                _shutil.copy2(pdf_path, os.path.join(out_dir, os.path.basename(pdf_path)))
            except Exception as cp_err:
                print(f"output_dir copy failed: {cp_err}")

        # upload to Supabase
        file_url = upload_pdf(pdf_path, file_name)

        # save to DB (non-fatal if table missing)
        try:
            doc_folder = data.get("doctor_name") or data.get("center_name") or "Unknown"
            if save_report:
                save_report(doc_folder, file_url, "pgta")
        except Exception as db_err:
            print(f"DB save skipped: {db_err}")

        return {
            "status": "success",
            "file_url": file_url
        }

    except Exception as e:
        return {"error": str(e)}
        
# -------- Bulk Report Generation --------
@app.post("/generate-bulk")
async def generate_bulk(request: Request):
    data = await request.json()

    output_files = []
    errors = []

    for row in data:
        patient_name = row.get("Patient Name", "Unknown")
        try:
            with_logo = row.get("logo_option", "without_logo") == "with_logo"
            generator = TERAReportGenerator(row, REPORT_DIR, with_logo=with_logo)
            pdf_path = generator.generate()

            file_name = _build_file_name(row, with_logo)

            # Copy to custom output_dir if provided
            custom_dir = (row.get("output_dir") or "").strip()
            if custom_dir:
                try:
                    out_dir = custom_dir if os.path.isabs(custom_dir) else os.path.join(BASE_DIR, custom_dir)
                    os.makedirs(out_dir, exist_ok=True)
                    import shutil as _shutil
                    _shutil.copy2(pdf_path, os.path.join(out_dir, os.path.basename(pdf_path)))
                except Exception as cp_err:
                    print(f"output_dir copy failed for {patient_name}: {cp_err}")

            # upload to Supabase if client available
            file_url = upload_pdf(pdf_path, file_name) if upload_pdf else f"/reports/{file_name}"

            try:
                doc_folder = row.get("doctor_name") or row.get("center_name") or "Unknown"
                if save_report:
                    save_report(doc_folder, file_url, "tera")
            except Exception as db_err:
                print(f"DB save skipped for {patient_name}: {db_err}")

            output_files.append({
                "file_name": os.path.basename(pdf_path),
                "file_url": file_url
            })
        except Exception as e:
            import traceback
            print(f"ERROR for {patient_name}: {e}")
            traceback.print_exc()
            errors.append({"patient": patient_name, "error": str(e)})

    print(f"Bulk done: {len(output_files)} success, {len(errors)} errors")
    return {
        "generated": output_files,
        "errors": errors
    }


# -------- Compare PDFs --------
_P1_FIELDS = {
    "Title":           (45,  60, 570, 145),
    "Patient Info":    (45, 144, 570, 250),
    "Status Section":  (45, 360, 570, 520),
    "Recommendations": (45, 520, 570, 710),
}
_P2_FIELDS = {
    "About TERA":  (45,  45, 570, 420),
    "Methodology": (45, 420, 570, 760),
}
_P3_FIELDS = {
    "Reviewer Block": (45, 45, 570, 760),
}
_PAGE_REGIONS = [_P1_FIELDS, _P2_FIELDS, _P3_FIELDS]

def _biopsy_ordinal(biopsy_no: str) -> str:
    match = re.search(r'(\d+)', str(biopsy_no))
    n = int(match.group(1)) if match else 1
    suffix = {1: "st", 2: "nd", 3: "rd"}.get(n if n < 20 else n % 10, "th")
    return f"{n}{suffix} biopsy"

def _safe_name(name: str) -> str:
    return re.sub(r'[^a-zA-Z0-9 ]', '', str(name).strip()).replace(' ', '_')

def _build_file_name(row: dict, with_logo: bool) -> str:
    patient = _safe_name(row.get("Patient Name", "Unknown"))
    biopsy  = _biopsy_ordinal(row.get("Biopsy No.", "1"))
    logo    = "with_logo" if with_logo else "without_logo"
    return f"TERA/{patient}_{biopsy}_TERA_report_{logo}.pdf"

def _norm(s: str) -> str:
    return re.sub(r'\s+', ' ', s).strip()

def _region_text(page, bbox) -> str:
    try:
        return page.within_bbox(bbox).extract_text() or ""
    except Exception:
        return ""

def _word_diff(a: str, b: str):
    sm = difflib.SequenceMatcher(None, a.split(), b.split(), autojunk=False)
    diffs = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag != "equal":
            diffs.append((" ".join(a.split()[i1:i2]), " ".join(b.split()[j1:j2])))
    return diffs

def _build_diff_html(sections: list) -> str:
    rows = []
    any_diff = False
    for label, issues in sections:
        has_diff = any("color:#c0392b" in i for i in issues)
        if has_diff:
            any_diff = True
        hdr_bg    = "#fde8e8" if has_diff else "#e8f5e9"
        hdr_color = "#c0392b" if has_diff else "#196F3D"
        rows.append(
            f"<div style='margin-bottom:14px;border:1px solid #ddd;"
            f"border-radius:6px;overflow:hidden;'>"
            f"<div style='background:{hdr_bg};padding:8px 12px;"
            f"font-weight:bold;color:{hdr_color};font-size:14px;'>{label}</div>"
            f"<div style='padding:8px 14px;font-family:monospace;font-size:12px;line-height:1.8;'>")
        for issue in issues:
            rows.append(f"<div>{issue}</div>")
        rows.append("</div></div>")

    summary = (
        "<span style='color:#c0392b;font-weight:bold'>Differences found — review highlighted items.</span>"
        if any_diff else
        "<span style='color:#196F3D;font-weight:bold'>No differences detected. PDFs match. ✓</span>"
    )
    return (f"<html><head><style>"
            f"body{{font-family:'Segoe UI',Arial,sans-serif;background:#f8f9fa;color:#333;padding:16px;}}"
            f".summary{{background:#fff;border:2px solid #1F497D;border-radius:6px;padding:12px 16px;"
            f"margin-bottom:16px;font-size:15px;}}"
            f"</style></head><body>"
            f"<div class='summary'>{summary}</div>{''.join(rows)}</body></html>")

@app.post("/compare-pdf")
async def compare_pdf(file1: UploadFile = File(...), file2: UploadFile = File(...)):

    data1 = await file1.read()
    data2 = await file2.read()

    sections = []

    with pdfplumber.open(io.BytesIO(data1)) as ldoc, \
         pdfplumber.open(io.BytesIO(data2)) as rdoc:

        n_left, n_right = len(ldoc.pages), len(rdoc.pages)

        if n_left != n_right:
            sections.append(("Page Count",
                [f"Manual PDF has <b>{n_left}</b> pages, Automated PDF has <b>{n_right}</b> pages."]))
        else:
            sections.append(("Page Count", [f"Both PDFs have {n_left} pages. ✓"]))

        for pg_idx in range(min(n_left, n_right)):
            pg_label = f"Page {pg_idx + 1}"
            regions  = _PAGE_REGIONS[pg_idx] if pg_idx < len(_PAGE_REGIONS) else {}
            lp, rp   = ldoc.pages[pg_idx], rdoc.pages[pg_idx]
            issues   = []

            lt = _norm(lp.extract_text() or "")
            rt = _norm(rp.extract_text() or "")
            if lt == rt:
                issues.append("Full page text is identical. ✓")
            else:
                diffs = _word_diff(lt, rt)
                issues.append(f"<span style='color:#c0392b'>Full page text differs ({len(diffs)} change(s)).</span>")
                for lw, rw in diffs[:20]:
                    issues.append(
                        f"  <tt>Manual:</tt> <span style='background:#fde8e8'>{lw or '(empty)'}</span>"
                        f"  →  <tt>Automated:</tt> <span style='background:#e8f5e9'>{rw or '(empty)'}</span>")
                if len(diffs) > 20:
                    issues.append(f"  … and {len(diffs)-20} more difference(s).")

            for region_name, bbox in regions.items():
                lr = _norm(_region_text(lp, bbox))
                rr = _norm(_region_text(rp, bbox))
                if lr == rr:
                    issues.append(f"  [{region_name}] identical ✓")
                else:
                    rdiffs = _word_diff(lr, rr)
                    issues.append(f"  <span style='color:#c0392b'>[{region_name}] {len(rdiffs)} difference(s):</span>")
                    for lw, rw in rdiffs[:8]:
                        issues.append(
                            f"    <tt>Manual:</tt> <span style='background:#fde8e8'>{lw or '(empty)'}</span>"
                            f"  →  <tt>Automated:</tt> <span style='background:#e8f5e9'>{rw or '(empty)'}</span>")

            sections.append((pg_label, issues))

    html = _build_diff_html(sections)
    return {"html": html, "differences": []}


# -------- Excel Upload --------
@app.post("/upload-excel")
async def upload_excel(file: UploadFile = File(...)):
    try:
        df = pd.read_excel(file.file)
        rows = df.to_dict(orient="records")

        def _to_json_safe(v):
            """Convert any pandas/numpy value to a JSON-serializable Python primitive."""
            # Handle None up-front
            if v is None:
                return None
            # Catches NaN, NaT, None via pd.isna
            try:
                if pd.isna(v):
                    return None
            except (TypeError, ValueError):
                pass
            # numpy scalar types (int64, float64, bool_, etc.)
            if hasattr(v, "item"):
                return v.item()
            # pandas Timestamp or any datetime-like
            if hasattr(v, "isoformat"):
                return str(v)
            return v

        cleaned_rows = [
            {k: _to_json_safe(v) for k, v in row.items()}
            for row in rows
        ]
        return {"rows": cleaned_rows}
    except Exception as e:
        return {"error": str(e), "rows": []}

@app.get("/test-db")
def test_db():
    if not supabase:
        return {"error": "Supabase client not initialized"}
    response = supabase.table("reports").select("*").execute()
    return response.data


# -------- Drafts --------
@app.post("/save-draft/{draft_type}")
async def save_draft(draft_type: str, request: Request):
    data = await request.json()
    if not supabase:
        return {"status": "error", "error": "Supabase not configured"}
    try:
        # check if draft already exists
        existing = supabase.table("drafts").select("id").eq("draft_type", draft_type).execute()
        if existing.data:
            supabase.table("drafts").update({"data": data}).eq("draft_type", draft_type).execute()
        else:
            supabase.table("drafts").insert({"draft_type": draft_type, "data": data}).execute()
        return {"status": "saved"}
    except Exception as e:
        return {"status": "error", "error": str(e)}

@app.get("/list-drafts")
def list_drafts():
    if not supabase:
        return {"drafts": [], "error": "Supabase not configured"}
    try:
        result = supabase.table("drafts").select("draft_type, updated_at").order("updated_at", desc=True).execute()
        return {"drafts": result.data or []}
    except Exception as e:
        return {"drafts": [], "error": str(e)}

@app.get("/load-draft/{draft_type}")
def load_draft(draft_type: str):
    if not supabase:
        return {"data": None, "error": "Supabase not configured"}
    try:
        result = supabase.table("drafts").select("data").eq("draft_type", draft_type).single().execute()
        return {"data": result.data["data"] if result.data else None}
    except Exception as e:
        return {"data": None, "error": str(e)}


# ================================================================
# PGT-A FILE-BASED DRAFT ENDPOINTS
# ================================================================

@app.post("/pgta/draft/save")
async def pgta_save_draft_file(request: Request):
    """Save a single patient's draft data as a JSON file in backend/drafts/PGTA/."""
    try:
        body = await request.json()
        patient = body.get("patient", {})
        patient_name = re.sub(r'[^a-zA-Z0-9 ]', '', str(patient.get("patient_name", "draft"))).replace(" ", "_").strip() or "draft"
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"pgta_bulk_draft_{patient_name}_{ts}.json"
        filepath = os.path.join(PGTA_DRAFT_DIR, filename)
        import json
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump({"patients": [patient], "_type": "pgta_bulk_draft", "_savedAt": datetime.now().isoformat()}, f, indent=2)
        return {"status": "saved", "filename": filename}
    except Exception as e:
        return {"status": "error", "error": str(e)}

@app.get("/pgta/draft/list")
def pgta_list_draft_files():
    """List all draft JSON files saved in backend/drafts/PGTA/."""
    try:
        files = sorted(
            [f for f in os.listdir(PGTA_DRAFT_DIR) if f.endswith(".json")],
            reverse=True
        )
        return {"files": files}
    except Exception as e:
        return {"files": [], "error": str(e)}

@app.delete("/pgta/draft/delete/{filename}")
def pgta_delete_draft_file(filename: str):
    """Delete a draft file from backend/drafts/PGTA/."""
    try:
        safe_name = os.path.basename(filename)
        filepath = os.path.join(PGTA_DRAFT_DIR, safe_name)
        if os.path.exists(filepath):
            os.remove(filepath)
            return {"status": "deleted"}
        return {"status": "not_found"}
    except Exception as e:
        return {"status": "error", "error": str(e)}


# ================================================================
# PGT-A REPORT COMPARISON ENDPOINT
# ================================================================

@app.post("/pgta/compare")
async def pgta_compare_reports(manual: UploadFile = File(...), automated: UploadFile = File(...)):
    """Compare two PGTA report files (PDF or DOCX) and return an HTML diff."""
    import difflib

    def extract_lines(file_bytes: bytes, filename: str):
        fname = filename.lower()
        if fname.endswith(".pdf"):
            try:
                lines = []
                with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                    for page in pdf.pages:
                        text = page.extract_text()
                        if text:
                            lines.extend(text.splitlines())
                return [l.strip() for l in lines if l.strip()]
            except Exception as e:
                return [f"[PDF extraction error: {e}]"]
        elif fname.endswith(".docx"):
            try:
                from docx import Document
                doc = Document(io.BytesIO(file_bytes))
                return [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            except Exception as e:
                return [f"[DOCX extraction error: {e}]"]
        return ["[Unsupported file format — please use PDF or DOCX]"]

    manual_bytes  = await manual.read()
    auto_bytes    = await automated.read()

    manual_lines  = extract_lines(manual_bytes,  manual.filename)
    auto_lines    = extract_lines(auto_bytes,     automated.filename)

    # Build an HTML diff table via the stdlib
    d = difflib.HtmlDiff(wrapcolumn=80)
    html_table = d.make_table(
        manual_lines, auto_lines,
        fromdesc=f"Manual — {manual.filename}",
        todesc=f"Automated — {automated.filename}",
        context=True, numlines=3
    )

    # Count change blocks
    matcher = difflib.SequenceMatcher(None, manual_lines, auto_lines)
    changes = [(tag, i1, i2, j1, j2) for tag, i1, i2, j1, j2 in matcher.get_opcodes() if tag != "equal"]

    return {
        "manual_file":  manual.filename,
        "auto_file":    automated.filename,
        "html_diff":    html_table,
        "total_changes": len(changes),
        "match":        len(changes) == 0,
    }


# ================================================================
# PGT-A REPORT ENDPOINTS
# ================================================================

@app.get("/pgta")
def pgta_page():
    p = os.path.join(FRONTEND_DIR, "pgta.html")
    if os.path.exists(p):
        return FileResponse(p, media_type="text/html")
    return {"error": "pgta.html not found"}


@app.get("/pgta/select-folder")
def pgta_select_folder():
    """Open a native OS folder picker dialog and return the selected path."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.lift()
        root.attributes('-topmost', True)
        folder = filedialog.askdirectory(title="Select Output Folder")
        root.destroy()
        return {"folder": folder or ""}
    except Exception as e:
        return {"folder": "", "error": str(e)}


@app.post("/pgta/upload-cnv")
async def pgta_upload_cnv(file: UploadFile = File(...)):
    """Upload a CNV chart image for an embryo. Returns server-side path."""
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in ['.png', '.jpg', '.jpeg']:
        return {"error": "Only PNG/JPG images allowed"}
    unique_name = str(uuid.uuid4()) + ext
    save_path = os.path.join(PGTA_CNV_DIR, unique_name)
    with open(save_path, "wb") as f:
        f.write(await file.read())
    return {"path": save_path, "name": unique_name, "url": f"/pgta/cnv-image/{unique_name}"}


@app.get("/pgta/cnv-image/{filename}")
def pgta_get_cnv_image(filename: str):
    path = os.path.join(PGTA_CNV_DIR, filename)
    if os.path.exists(path):
        return FileResponse(path)
    return {"error": "Image not found"}


@app.post("/pgta/preview")
async def pgta_preview_report(request: Request):
    """Generate a preview PDF and return its URL."""
    try:
        data = await request.json()
        file_id = str(uuid.uuid4()) + ".pdf"
        filepath = os.path.join(TEMP_DIR, file_id)

        embryos = data.get("embryos_data", [])
        embryos, tmp_paths = _resolve_cnv_images(embryos)

        tmpl = PGTAReportTemplate()
        tmpl.generate_pdf(
            filepath,
            data.get("patient_data", {}),
            embryos,
            show_logo=data.get("show_logo", True),
            show_grid=data.get("show_grid", False)
        )
        # Clean up temp image files (non-blocking — ignore errors)
        for p in tmp_paths:
            try: os.remove(p)
            except Exception: pass
        return {"preview_url": f"/pgta/preview-file/{file_id}"}
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"error": str(e)}


@app.get("/pgta/preview-file/{filename}")
def pgta_preview_file(filename: str):
    path = os.path.join(TEMP_DIR, filename)
    return FileResponse(path, media_type="application/pdf")


@app.post("/pgta/verify-trf")
async def pgta_verify_trf(file: UploadFile = File(...)):
    """Placeholder for TRF data extraction and verification."""
    try:
        # In a real scenario, we'd use pdfplumber/tesseract to extract text
        # For now, we simulate a successful extraction for the UI parity
        await file.read() # Consume the file
        return {
            "status": "success",
            "extracted_data": {
                "patient_name": "Sita Sharma",
                "sample_id": "PS12345",
                "biopsy_date": "20-03-2026",
                "clinician": "Dr. Anderson"
            }
        }
    except Exception as e:
        return {"error": str(e)}


def _upload_in_background(filepath: str, filename: str):
    """Upload a file to Supabase storage in a background task (non-blocking)."""
    try:
        if upload_pgta_file:
            upload_pgta_file(filepath, filename)
    except Exception as e:
        print(f"Background Supabase upload failed for {filename}: {e}")

@app.post("/pgta/generate")
async def pgta_generate_report(request: Request, background_tasks: BackgroundTasks):
    """Generate final PGT-A reports (PDF/DOCX) and handle output preferences."""
    try:
        data = await request.json()
        patient_info = data.get("patient_info") or data.get("patient_data", {})
        embryos = data.get("embryos") or data.get("embryos_data", [])
        options = data.get("options", {})
        custom_output_dir = data.get("output_dir")
        
        show_logo = options.get("show_logo", True)
        show_grid = options.get("show_grid", False)
        formats = options.get("formats", ["pdf"])
        
        gen_pdf = "pdf" in formats
        gen_docx = "docx" in formats

        # Resolve any base64 CNV images → temp files before generating
        embryos, tmp_cnv_paths = _resolve_cnv_images(embryos)

        def _name_parts(raw):
            """Return cleaned list of name words, stripped of special chars."""
            cleaned = re.sub(r'[^a-zA-Z0-9 ]', '', str(raw or '')).strip()
            return [p for p in cleaned.split() if p]

        p_parts = _name_parts(patient_info.get("patient_name", ""))
        s_parts = _name_parts(patient_info.get("spouse_name", ""))

        p_first     = p_parts[0].upper() if p_parts else "UNKNOWN"
        p_last_init = p_parts[-1][0].upper() if len(p_parts) > 1 else (p_parts[0][0].upper() if p_parts else "X")

        # Spouse may be "w/o" or blank — omit initials if meaningless
        s_raw = str(patient_info.get("spouse_name", "") or "").strip().upper()
        include_spouse = s_parts and s_raw not in ("WO", "W/O", "NA", "N/A", "")
        if include_spouse:
            s_first     = s_parts[0].upper()
            s_last_init = s_parts[-1][0].upper() if len(s_parts) > 1 else s_parts[0][0].upper()
            name_segment = f"{p_first}_{s_first}_{p_last_init}_{s_last_init}"
        else:
            name_segment = f"{p_first}_{p_last_init}"

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        logo_tag = "withlogo" if show_logo else "withoutlogo"

        base_filename = f"PGTA_{name_segment}_{timestamp}_{logo_tag}"
        
        # Resolve custom output dir — must be absolute to be usable
        extra_output_dir = None
        if custom_output_dir:
            cand = custom_output_dir if os.path.isabs(custom_output_dir) else os.path.join(BASE_DIR, custom_output_dir)
            try:
                os.makedirs(cand, exist_ok=True)
                extra_output_dir = cand
            except Exception:
                pass

        results = {}

        # 1. Generate PDF — always into PGTA_REPORT_DIR so /reports-pgta/ URL is valid
        if gen_pdf:
            file_name_pdf = f"{base_filename}.pdf"
            filepath_pdf = os.path.join(PGTA_REPORT_DIR, file_name_pdf)
            tmpl = PGTAReportTemplate()
            tmpl.generate_pdf(filepath_pdf, patient_info, embryos, show_logo=show_logo, show_grid=show_grid)
            if extra_output_dir:
                try:
                    shutil.copy2(filepath_pdf, os.path.join(extra_output_dir, file_name_pdf))
                except Exception as cp_err:
                    print(f"Copy to output dir skipped: {cp_err}")
            local_pdf_url = f"/reports-pgta/{file_name_pdf}"
            # Upload to Supabase in background — response returns immediately with local URL
            if upload_pgta_file:
                background_tasks.add_task(_upload_in_background, filepath_pdf, file_name_pdf)
            results["pdf"] = {"file": file_name_pdf, "url": local_pdf_url, "local_url": local_pdf_url}

        # 2. Generate DOCX — same pattern
        if gen_docx:
            file_name_docx = f"{base_filename}.docx"
            filepath_docx = os.path.join(PGTA_REPORT_DIR, file_name_docx)
            docx_gen = PGTADocxGenerator(assets_dir="assets/pgta")
            docx_gen.generate_docx(filepath_docx, patient_info, embryos, show_logo=show_logo, show_grid=show_grid)
            if extra_output_dir:
                try:
                    shutil.copy2(filepath_docx, os.path.join(extra_output_dir, file_name_docx))
                except Exception as cp_err:
                    print(f"Copy to output dir skipped: {cp_err}")
            local_docx_url = f"/reports-pgta/{file_name_docx}"
            if upload_pgta_file:
                background_tasks.add_task(_upload_in_background, filepath_docx, file_name_docx)
            results["docx"] = {"file": file_name_docx, "url": local_docx_url, "local_url": local_docx_url}

        # Clean up temp CNV image files after generation
        for p in tmp_cnv_paths:
            try: os.remove(p)
            except Exception: pass

        return {"status": "success", "results": results}
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"error": str(e)}


async def _parse_pgta_excel_core(contents: bytes):
    """Core logic to parse PGT-A Excel from bytes."""
    try:
        xl = pd.ExcelFile(io.BytesIO(contents))
        sheets = xl.sheet_names
        sheet_names_lower = [s.lower().strip() for s in sheets]

        # 1. Helper functions from desktop app
        def get_clean_value(row, keys, default=''):
            if isinstance(keys, str): keys = [keys]
            for k in keys:
                if k in row:
                    val = row[k]
                    if pd.isna(val): continue
                    s_val = str(val).strip(' \t\r\f\v')
                    if s_val.lower() in ['nan', 'none', 'nat', 'null']: continue
                    if s_val: return s_val
            return default

        def format_date(d_val):
            if not d_val: return ""
            s = str(d_val).split(' ')[0]
            try:
                for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%d.%m.%Y"):
                    try:
                        dt = datetime.strptime(s, fmt)
                        return dt.strftime("%d-%m-%Y")
                    except ValueError: continue
                return s.replace('/', '-')
            except:
                return s.replace('/', '-')

        def normalize_str(s):
            if not s: return ""
            s = str(s).upper().strip()
            prefixes = ['MRS.', 'MR.', 'SMT.', 'DR.', 'MS.', 'MISS.', 'PROF.', 'R.', 'S.', 'K.', 'M.', 'D.', 'P.', 'A.', 'B.', 'C.', 'G.', 'H.', 'J.', 'L.', 'N.', 'T.', 'V.', 'W.']
            for prefix in prefixes:
                if s.startswith(prefix):
                    s = s[len(prefix):].strip()
                    break
            s = re.sub(r'\([^)]*\)', '', s)
            s = re.sub(r'[^A-Z0-9]', '', s)
            return s

        # 2. Find Details and Summary sheets
        details_idx = next((i for i, s in enumerate(sheet_names_lower) if 'detail' in s), None)
        summary_idx = next((i for i, s in enumerate(sheet_names_lower) if 'summary' in s), None)

        if details_idx is None and summary_idx is None and len(sheets) >= 2:
            details_idx, summary_idx = 0, 1
        elif details_idx is None and len(sheets) == 1:
            summary_idx = 0

        details_df = None
        summary_df = None

        if details_idx is not None:
            details_df = xl.parse(sheets[details_idx])
            details_df.columns = [str(c).strip() for c in details_df.columns]
        
        if summary_idx is not None:
            # Header discovery for summary sheet
            try:
                df_full = xl.parse(sheets[summary_idx], header=None)
                header_row_idx = 0
                for r_idx, row in df_full.iterrows():
                    if any('sample name' in str(val).lower() for val in row.values):
                        header_row_idx = r_idx
                        break
                summary_df = xl.parse(sheets[summary_idx], header=header_row_idx)
                summary_df.columns = [str(c).strip() for c in summary_df.columns]
            except:
                summary_df = xl.parse(sheets[summary_idx])
                summary_df.columns = [str(c).strip() for c in summary_df.columns]

        patient_map = {}
        patients = []

        # 3. Process Details (Patients)
        if details_df is not None:
            for _, p_row in details_df.iterrows():
                p_name = get_clean_value(p_row, ['Patient Name', 'patient_name', 'Name'])
                if not p_name: continue
                
                pid = get_clean_value(p_row, ['Sample ID', 'Patient ID', 'sample_id', 'PIN', 'pin'])
                b_date = format_date(get_clean_value(p_row, ['Date of Biopsy', 'Biopsy Date']))
                r_date = format_date(get_clean_value(p_row, ['Date Sample Received', 'Receipt Date', 'Sample Receipt Date']))

                patient = {
                    "patient_name": p_name,
                    "sample_number": get_clean_value(p_row, ['Sample Number', 'Sample No', 'Sample No.', 'SampleNumber', 'Accession Number', 'Acc. No.']) or '',
                    "hospital_clinic": get_clean_value(p_row, ['Center name', 'Center Name', 'center_name', 'Hospital', 'Clinic']),
                    "biopsy_date": b_date,
                    "sample_receipt_date": r_date,
                    "biopsy_performed_by": get_clean_value(p_row, ['EMBRYOLOGIST NAME', 'Embryologist Name', 'Biologist']),
                    "spouse_name": get_clean_value(p_row, ['Spouse Name', 'Husband Name', 'Partner Name', 'spouse_name']) or 'w/o',
                    "pin": pid,
                    "age": get_clean_value(p_row, ['Age', 'age', 'Patient Age']),
                    "referring_clinician": get_clean_value(p_row, ['Referring Clinician', 'referring_clinician', 'Doctor']),
                    "sample_collection_date": b_date,
                    "specimen": get_clean_value(p_row, ['Specimen', 'Specimen Type', 'Sample Type'], 'DAY 5 TROPHECTODERM BIOPSY'),
                    "report_date": datetime.now().strftime('%d-%m-%Y'),
                    "indication": get_clean_value(p_row, ['Indication', 'indication', 'Clinical Indication']),
                    "embryos": []
                }
                patient_map[normalize_str(pid)] = patient
                patient_map[normalize_str(p_name)] = patient
                patients.append(patient)

        # 4. Process Summary (Embryos)
        if summary_df is not None:
            for _, s_row in summary_df.iterrows():
                sample_name = get_clean_value(s_row, ['Sample name', 'Sample Name', 'sample_name', 'Sample ID'])
                if not sample_name: continue
                
                embryo = {
                    "embryo_id": sample_name,
                    "embryo_id_detail": sample_name,
                    "result_summary": get_clean_value(s_row, ['Result', 'result', 'Summary']),
                    "interpretation": get_clean_value(s_row, ['Conclusion', 'Interpretation', 'interpretation', 'Result']),
                    "mtcopy": get_clean_value(s_row, ['MTcopy', 'MT Copy', 'mtcopy', 'MT']),
                    "autosomes": get_clean_value(s_row, ['AUTOSOMES', 'Autosomes', 'autosomes', 'Aneuploidy']),
                    "sex_chromosomes": get_clean_value(s_row, ['SEX', 'Sex Chromosomes', 'sex_chromosomes', 'Sex'], 'Normal'),
                    "result_description": get_clean_value(s_row, ['Result', 'result_description', 'Result Description']),
                    "chromosome_statuses": {},
                    "mosaic_percentages": {},
                    "inconclusive_comment": ""
                }
                
                # Match to patient
                norm_sample = str(normalize_str(sample_name))
                matched = False
                for key, patient in patient_map.items():
                    if key and str(key) in norm_sample:
                        if "embryos" not in patient or not isinstance(patient["embryos"], list):
                            patient["embryos"] = []
                        patient["embryos"].append(embryo)
                        matched = True
                        break
                
                if not matched and patients:
                    if "embryos" not in patients[-1] or not isinstance(patients[-1]["embryos"], list):
                        patients[-1]["embryos"] = []
                    patients[-1]["embryos"].append(embryo)

        return {"patients": patients, "sheet_names": sheets}
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise e

@app.post("/pgta/parse-excel")
async def pgta_parse_excel(file: UploadFile = File(...)):
    """Parse a PGT-A Excel file (Details + summary sheets) and return structured patient/embryo data."""
    try:
        contents = await file.read()
        result = await _parse_pgta_excel_core(contents)
        return result
    except Exception as e:
        return {"error": str(e)}

@app.post("/pgta/parse-excel-bulk")
async def pgta_parse_excel_bulk(files: List[UploadFile] = File(...)):
    """
    Parse multiple files: one Excel file + multiple CNV images.
    Returns structured data with CNV mapping.
    """
    try:
        excel_file = None
        image_files = []
        
        for f in files:
            ext = os.path.splitext(f.filename or "")[1].lower()
            if ext in ('.xlsx', '.xls'):
                excel_file = f
            elif ext in ('.png', '.jpg', '.jpeg'):
                image_files.append(f)
                
        if not excel_file:
            return {"error": "No Excel file found in the uploaded folder."}
            
        # 1. Parse Excel
        excel_contents = await excel_file.read()
        data = await _parse_pgta_excel_core(excel_contents)
        patients = data.get("patients", [])
        
        # 2. Map CNV images
        img_filenames = [f.filename for f in image_files]
        all_embryos = []
        for p in patients:
            all_embryos.extend(p.get("embryos", []))
            
        mapped_count = auto_map_cnvs(all_embryos, img_filenames)
        
        # 3. Save mapped images and set URLs
        for f in image_files:
            # We only save files that were actually mapped to save space
            # Find if this file was mapped
            is_mapped = any(e.get("cnv_image_name") == f.filename for e in all_embryos)
            if is_mapped:
                unique_name = str(uuid.uuid4()) + os.path.splitext(f.filename)[1].lower()
                save_path = os.path.join(PGTA_CNV_DIR, unique_name)
                # Reset file pointer if needed (though it should be fine here)
                content = await f.read()
                with open(save_path, "wb") as out:
                    out.write(content)
                
                # Update all embryos that used this filename
                for e in all_embryos:
                    if e.get("cnv_image_name") == f.filename:
                        e["cnv_image_path"] = save_path
                        e["cnv_image_url"] = f"/pgta/cnv-image/{unique_name}"
                        # We also set b64 for preview in frontend if needed, 
                        # but URL is better for bulk handled in backend.
                        # For consistency with current frontend:
                        e["cnv_image_b64"] = f"data:image/png;base64,{base64.b64encode(content).decode()}"

        return {
            "patients": patients, 
            "sheet_names": data.get("sheet_names", []),
            "mapped_count": mapped_count,
            "total_images": len(image_files)
        }
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"error": str(e)}


@app.post("/pgta/compare")
async def pgta_compare_reports(manual: UploadFile = File(...), automated: UploadFile = File(...)):
    """Compare a manual PGT-A PDF report with an automated one."""
    try:
        from report_comparator import PGTAReportComparator

        m_id = str(uuid.uuid4()) + ".pdf"
        a_id = str(uuid.uuid4()) + ".pdf"
        m_path = os.path.join(TEMP_DIR, m_id)
        a_path = os.path.join(TEMP_DIR, a_id)

        with open(m_path, "wb") as f:
            f.write(await manual.read())
        with open(a_path, "wb") as f:
            f.write(await automated.read())

        try:
            comparator = PGTAReportComparator()
            result = comparator.compare_single_pair(m_path, a_path)
            return result
        finally:
            for p in [m_path, a_path]:
                try:
                    os.unlink(p)
                except:
                    pass
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"error": str(e)}


@app.get("/pgta/storage/list")
async def pgta_storage_list(path: str = "pgta"):
    """
    List items in the Report-inputs Supabase bucket at the given path.
    Uses the service-role key so it bypasses anon RLS restrictions.
    """
    if supabase is None:
        from fastapi import HTTPException
        raise HTTPException(status_code=503, detail="Supabase client not configured")
    # Clean path: remove leading/trailing slashes for Supabase storage list
    clean_path = (path or "").strip("/")
    try:
        result = supabase.storage.from_("Report-inputs").list(
            clean_path,
            {"limit": 500, "sortBy": {"column": "name", "order": "asc"}}
        )
        items = []
        for item in (result or []):
            if isinstance(item, dict):
                items.append({"name": item.get("name"), "id": item.get("id"), "metadata": item.get("metadata")})
            else:
                items.append({
                    "name": getattr(item, "name", None),
                    "id": getattr(item, "id", None),
                    "metadata": getattr(item, "metadata", None),
                })
        return {"items": items, "path": path}
    except Exception as exc:
        from fastapi import HTTPException
        raise HTTPException(status_code=500, detail=str(exc))


# ================================================================
# KARYOTYPE REPORT ENDPOINTS
# ================================================================
from karyotype_template import KaryotypeReportGenerator

KARYOTYPE_REPORT_DIR = os.path.join(BASE_DIR, "reports-karyotype")
KARYOTYPE_CNV_DIR   = os.path.join(BASE_DIR, "uploads", "karyotype_images")
KARYOTYPE_DRAFT_DIR = os.path.join(BASE_DIR, "drafts", "KARYOTYPE")

os.makedirs(KARYOTYPE_REPORT_DIR, exist_ok=True)
os.makedirs(KARYOTYPE_CNV_DIR, exist_ok=True)
os.makedirs(KARYOTYPE_DRAFT_DIR, exist_ok=True)

app.mount("/reports-karyotype", StaticFiles(directory=KARYOTYPE_REPORT_DIR), name="reports-karyotype")

@app.get("/karyotype")
def karyotype_page():
    p = os.path.join(FRONTEND_DIR, "karyotype.html")
    if os.path.exists(p):
        return FileResponse(p, media_type="text/html")
    return {"error": "karyotype.html not found"}

@app.post("/karyotype/upload-image")
async def karyotype_upload_image(file: UploadFile = File(...)):
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in ['.png', '.jpg', '.jpeg']:
        return {"error": "Only PNG/JPG images allowed"}
    unique_name = str(uuid.uuid4()) + ext
    save_path = os.path.join(KARYOTYPE_CNV_DIR, unique_name)
    with open(save_path, "wb") as f:
        f.write(await file.read())
    return {"path": save_path, "name": unique_name, "url": f"/karyotype/image/{unique_name}"}

@app.get("/karyotype/image/{filename}")
def karyotype_get_image(filename: str):
    path = os.path.join(KARYOTYPE_CNV_DIR, filename)
    if os.path.exists(path):
        return FileResponse(path)
    return {"error": "Image not found"}

def _resolve_karyotype_images(image_urls: list) -> tuple:
    """Takes a list of uploaded image URL paths (from /karyotype/upload-image) and returns full absolute paths."""
    resolved_paths = []
    for url in image_urls:
        if isinstance(url, str) and url.startswith("/karyotype/image/"):
            filename = url.split("/")[-1]
            resolved_paths.append(os.path.join(KARYOTYPE_CNV_DIR, filename))
        elif isinstance(url, dict) and url.get("path"):
            resolved_paths.append(url.get("path"))
        elif isinstance(url, str) and os.path.exists(url):
            resolved_paths.append(url)
    return resolved_paths

@app.post("/karyotype/preview")
async def karyotype_preview_report(request: Request):
    try:
        data = await request.json()
        file_id = str(uuid.uuid4()) + ".pdf"
        filepath = os.path.join(TEMP_DIR, file_id)

        patient_data = data.get("patient_data", {})
        images = _resolve_karyotype_images(data.get("images", []))

        gen = KaryotypeReportGenerator(patient_data, images, TEMP_DIR, include_logo=data.get("show_logo", True))
        gen.filepath = filepath
        gen.filename = file_id
        gen.generate()

        return {"preview_url": f"/karyotype/preview-file/{file_id}"}
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"error": str(e)}

@app.get("/karyotype/preview-file/{filename}")
def karyotype_preview_file(filename: str):
    path = os.path.join(TEMP_DIR, filename)
    return FileResponse(path, media_type="application/pdf")

@app.post("/karyotype/generate")
async def karyotype_generate_report(request: Request, background_tasks: BackgroundTasks):
    try:
        data = await request.json()
        patient_info = data.get("patient_data", {})
        images = _resolve_karyotype_images(data.get("images", []))
        options = data.get("options", {})
        
        show_logo = options.get("show_logo", True)

        gen = KaryotypeReportGenerator(patient_info, images, KARYOTYPE_REPORT_DIR, include_logo=show_logo)
        pdf_path = gen.generate()
        file_name = os.path.basename(pdf_path)

        # You can add upload scripts here like PGTA if required
        local_pdf_url = f"/reports-karyotype/{file_name}"
        
        return {"status": "success", "results": {"pdf": {"file": file_name, "url": local_pdf_url, "local_url": local_pdf_url}}}
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"error": str(e)}

@app.post("/karyotype/draft/save")
async def karyotype_save_draft_file(request: Request):
    try:
        body = await request.json()
        patient = body.get("patient", {})
        patient_name = re.sub(r'[^a-zA-Z0-9 ]', '', str(patient.get("NAME", "draft"))).replace(" ", "_").strip() or "draft"
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"karyotype_bulk_draft_{patient_name}_{ts}.json"
        import json
        with open(os.path.join(KARYOTYPE_DRAFT_DIR, filename), "w", encoding="utf-8") as f:
            json.dump({"patients": [patient], "_type": "karyotype_bulk_draft", "_savedAt": datetime.now().isoformat()}, f, indent=2)
        return {"status": "saved", "filename": filename}
    except Exception as e:
        return {"status": "error", "error": str(e)}

@app.get("/karyotype/draft/list")
def karyotype_list_draft_files():
    try:
        files = sorted([f for f in os.listdir(KARYOTYPE_DRAFT_DIR) if f.endswith(".json")], reverse=True)
        return {"files": files}
    except Exception as e:
        return {"files": [], "error": str(e)}

@app.delete("/karyotype/draft/delete/{filename}")
def karyotype_delete_draft_file(filename: str):
    try:
        safe_name = os.path.basename(filename)
        filepath = os.path.join(KARYOTYPE_DRAFT_DIR, safe_name)
        if os.path.exists(filepath):
            os.remove(filepath)
            return {"status": "deleted"}
        return {"status": "not_found"}
    except Exception as e:
        return {"status": "error", "error": str(e)}

