from fastapi import FastAPI, UploadFile, File, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from dotenv import load_dotenv
load_dotenv()


import pandas as pd
import os
import io
import uuid
import math
import re
import difflib
from datetime import datetime

import pdfplumber

from tera_template import TERAReportGenerator
from pgta_template import PGTAReportTemplate
from fastapi.staticfiles import StaticFiles
from supabase_client import supabase
from supabase_client import upload_pdf, save_report




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

PGTA_CNV_DIR = os.path.join(BASE_DIR, "uploads", "pgta_cnv")
os.makedirs(PGTA_CNV_DIR, exist_ok=True)

app.mount("/reports", StaticFiles(directory=REPORT_DIR), name="reports")
app.mount("/reports-pgta", StaticFiles(directory=PGTA_REPORT_DIR), name="reports-pgta")

# Serve root-level static assets (logo, icons, images)
if os.path.exists(ROOT_DIR):
    app.mount("/static", StaticFiles(directory=ROOT_DIR), name="static")

@app.get("/")
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

        # upload to Supabase
        file_url = upload_pdf(pdf_path, file_name)

        # save to DB (non-fatal if table missing)
        try:
            save_report(doctor_folder, file_url, "tera")
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

            file_url = upload_pdf(pdf_path, file_name)
            try:
                save_report(doctor_folder, file_url, "tera")
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
                [f"Left PDF has <b>{n_left}</b> pages, Right PDF has <b>{n_right}</b> pages."]))
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
                        f"  <tt>Left:</tt> <span style='background:#fde8e8'>{lw or '(empty)'}</span>"
                        f"  →  <tt>Right:</tt> <span style='background:#e8f5e9'>{rw or '(empty)'}</span>")
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
                            f"    <tt>L:</tt> <span style='background:#fde8e8'>{lw or '(empty)'}</span>"
                            f"  →  <tt>R:</tt> <span style='background:#e8f5e9'>{rw or '(empty)'}</span>")

            sections.append((pg_label, issues))

    html = _build_diff_html(sections)
    return {"html": html, "differences": []}


# -------- Excel Upload --------
@app.post("/upload-excel")
async def upload_excel(file: UploadFile = File(...)):

    df = pd.read_excel(file.file)

    rows = df.to_dict(orient="records")

    # Convert NaN values to None
    for row in rows:
        for key, value in row.items():
            if isinstance(value, float) and math.isnan(value):
                row[key] = None

    return {"rows": rows}

@app.get("/test-db")
def test_db():
    response = supabase.table("reports").select("*").execute()
    return response.data


# -------- Drafts --------
@app.post("/save-draft/{draft_type}")
async def save_draft(draft_type: str, request: Request):
    data = await request.json()
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
    try:
        result = supabase.table("drafts").select("draft_type, updated_at").order("updated_at", desc=True).execute()
        return {"drafts": result.data or []}
    except Exception as e:
        return {"drafts": [], "error": str(e)}

@app.get("/load-draft/{draft_type}")
def load_draft(draft_type: str):
    try:
        result = supabase.table("drafts").select("data").eq("draft_type", draft_type).single().execute()
        return {"data": result.data["data"] if result.data else None}
    except Exception as e:
        return {"data": None, "error": str(e)}


# ================================================================
# PGT-A REPORT ENDPOINTS
# ================================================================

@app.get("/pgta")
def pgta_page():
    p = os.path.join(FRONTEND_DIR, "pgta.html")
    if os.path.exists(p):
        return FileResponse(p, media_type="text/html")
    return {"error": "pgta.html not found"}


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

        tmpl = PGTAReportTemplate()
        tmpl.generate_pdf(
            filepath,
            data.get("patient_data", {}),
            data.get("embryos_data", []),
            show_logo=data.get("show_logo", True),
            show_grid=data.get("show_grid", False)
        )
        return {"preview_url": f"/pgta/preview-file/{file_id}"}
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"error": str(e)}


@app.get("/pgta/preview-file/{filename}")
def pgta_preview_file(filename: str):
    path = os.path.join(TEMP_DIR, filename)
    return FileResponse(path, media_type="application/pdf")


@app.post("/pgta/generate")
async def pgta_generate_report(request: Request):
    """Generate final PGT-A PDF report and return download URL."""
    try:
        data = await request.json()
        patient_data = data.get("patient_data", {})

        sample_num = re.sub(r'[^a-zA-Z0-9]', '', str(patient_data.get("sample_number", "report")))
        patient_name = re.sub(r'[^a-zA-Z0-9 ]', '', str(patient_data.get("patient_name", "Unknown"))).replace(" ", "_")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"PGTA_{sample_num}_{patient_name}_{timestamp}.pdf"
        filepath = os.path.join(PGTA_REPORT_DIR, file_name)

        tmpl = PGTAReportTemplate()
        tmpl.generate_pdf(
            filepath,
            patient_data,
            data.get("embryos_data", []),
            show_logo=data.get("show_logo", True),
            show_grid=data.get("show_grid", False)
        )

        # Upload to Supabase storage: reports bucket → PGT-A/ folder
        supabase_url = None
        try:
            from supabase_client import _get_client
            client = _get_client()
            storage_path = f"PGT-A/{file_name}"
            with open(filepath, "rb") as f:
                client.storage.from_("reports").upload(
                    storage_path, f,
                    {"upsert": "true", "content-type": "application/pdf"}
                )
            supabase_url = client.storage.from_("reports").get_public_url(storage_path)
        except Exception as sup_err:
            print(f"Supabase upload warning: {sup_err}")

        return {"status": "success", "file": file_name, "url": f"/reports-pgta/{file_name}", "supabase_url": supabase_url}
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"error": str(e)}


@app.post("/pgta/parse-excel")
async def pgta_parse_excel(file: UploadFile = File(...)):
    """Parse a PGT-A Excel file (Details + summary sheets) and return structured patient/embryo data."""
    try:
        contents = await file.read()
        xl = pd.ExcelFile(io.BytesIO(contents))
        sheets = xl.sheet_names

        def clean_val(v):
            if v is None:
                return ""
            if isinstance(v, float) and math.isnan(v):
                return ""
            return str(v).strip()

        details_df = None
        summary_df = None

        for sheet in sheets:
            sname = sheet.lower().strip()
            if 'detail' in sname:
                details_df = xl.parse(sheet)
            elif 'summary' in sname:
                summary_df = xl.parse(sheet)

        # Fallback: use first two sheets
        if details_df is None and summary_df is None and len(sheets) >= 2:
            details_df = xl.parse(sheets[0])
            summary_df = xl.parse(sheets[1])
        elif details_df is None and len(sheets) == 1:
            summary_df = xl.parse(sheets[0])

        patient_map = {}
        patients = []

        if details_df is not None:
            details_df.columns = [str(c).strip() for c in details_df.columns]
            for _, row in details_df.iterrows():
                row = {k: clean_val(v) for k, v in row.items()}
                pid = row.get('Sample ID', '') or row.get('Patient ID', '') or row.get('sample_id', '')
                pname = row.get('Patient Name', '') or row.get('patient_name', '')
                if not pname and not pid:
                    continue
                patient = {
                    "patient_name": pname,
                    "sample_number": pid,
                    "hospital_clinic": row.get('Center name', '') or row.get('Center Name', '') or row.get('center_name', ''),
                    "biopsy_date": row.get('Date of Biopsy', '') or row.get('Biopsy Date', ''),
                    "sample_receipt_date": row.get('Date Sample Received', '') or row.get('Date of Sample Received', ''),
                    "biopsy_performed_by": row.get('EMBRYOLOGIST NAME', '') or row.get('Embryologist Name', '') or row.get('embryologist_name', ''),
                    "spouse_name": row.get('Spouse Name', '') or row.get('spouse_name', 'w/o'),
                    "pin": row.get('PIN', '') or row.get('pin', ''),
                    "age": row.get('Age', '') or row.get('age', ''),
                    "referring_clinician": row.get('Referring Clinician', '') or row.get('referring_clinician', ''),
                    "sample_collection_date": row.get('Sample Collection Date', '') or row.get('Date Collected', ''),
                    "specimen": row.get('Specimen', '') or row.get('specimen', 'DAY 5 TROPHECTODERM BIOPSY'),
                    "report_date": row.get('Report Date', '') or datetime.now().strftime('%d-%m-%Y'),
                    "indication": row.get('Indication', '') or row.get('indication', ''),
                    "embryos": []
                }
                patient_map[pid] = patient
                patients.append(patient)

        if summary_df is not None:
            summary_df.columns = [str(c).strip() for c in summary_df.columns]
            for _, row in summary_df.iterrows():
                row = {k: clean_val(v) for k, v in row.items()}
                sample_name = row.get('Sample name', '') or row.get('Sample Name', '') or row.get('sample_name', '')
                if not sample_name:
                    continue
                embryo = {
                    "embryo_id": sample_name,
                    "embryo_id_detail": sample_name,
                    "result_summary": row.get('Result', '') or row.get('result', ''),
                    "interpretation": row.get('Conclusion', '') or row.get('Interpretation', '') or row.get('interpretation', ''),
                    "mtcopy": row.get('MTcopy', '') or row.get('MT Copy', '') or row.get('mtcopy', ''),
                    "autosomes": row.get('AUTOSOMES', '') or row.get('Autosomes', '') or row.get('autosomes', ''),
                    "sex_chromosomes": row.get('SEX', '') or row.get('Sex Chromosomes', '') or row.get('sex_chromosomes', 'Normal'),
                    "result_description": row.get('Result', '') or row.get('result_description', ''),
                    "chromosome_statuses": {},
                    "mosaic_percentages": {},
                    "inconclusive_comment": ""
                }
                # Match to patient by sample_name prefix
                matched = False
                for pid, patient in patient_map.items():
                    if pid and sample_name.startswith(pid):
                        patient["embryos"].append(embryo)
                        matched = True
                        break
                if not matched and patients:
                    patients[-1]["embryos"].append(embryo)

        return {"patients": patients, "sheet_names": sheets}
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

