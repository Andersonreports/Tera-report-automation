from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import os
import uuid
import pandas as pd
from tera_template import TERAReportGenerator
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(title="TERA Report Backend")

# Enable CORS exactly as before
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMP_DIR = os.path.join(BASE_DIR, "temp")
REPORT_DIR = os.path.join(BASE_DIR, "reports")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(REPORT_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)

@app.get("/")
def root():
    return {"status": "TERA backend running"}

@app.post("/preview")
async def preview_report(data: dict):
    preview_id = str(uuid.uuid4()) + ".pdf"
    filepath = os.path.join(TEMP_DIR, preview_id)

    gen = TERAReportGenerator(data, TEMP_DIR)
    gen.filepath = filepath
    gen.filename = preview_id
    gen.generate()

    return {"preview_url": f"/preview-file/{preview_id}"}


@app.get("/preview-file/{filename}")
def get_preview(filename: str):
    filepath = os.path.join(TEMP_DIR, filename)
    return FileResponse(filepath, media_type="application/pdf")


@app.post("/generate")
async def generate_report(data: dict):
    gen = TERAReportGenerator(data, REPORT_DIR)
    path = gen.generate()
    return {"file": os.path.basename(path)}


@app.post("/upload-excel")
async def upload_excel(file: UploadFile = File(...)):
    filepath = os.path.join(UPLOAD_DIR, file.filename)
    with open(filepath, "wb") as f:
        f.write(await file.read())

    df = pd.read_excel(filepath)
    df = df.dropna(how="all")
    df.columns = df.columns.str.strip()

    return df.to_dict(orient="records")

