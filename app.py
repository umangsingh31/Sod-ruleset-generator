import logging
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from fastapi import Request
from fastapi.staticfiles import StaticFiles
import shutil
import os
import uuid

from generator import generate_sre

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s - %(message)s",
)
logger = logging.getLogger(__name__)

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate")
async def generate(
    template: UploadFile = File(...),
    projects: UploadFile = File(...),
    owners: UploadFile = File(...),
    baseline: UploadFile = File(None),
):
    workdir = f"work_{uuid.uuid4().hex}"
    os.makedirs(workdir, exist_ok=True)

    logger.info("Received generate request")

    template_path = os.path.join(workdir, "template.xlsx")
    projects_path = os.path.join(workdir, "projects.xlsx")
    owners_path = os.path.join(workdir, "owners.xlsx")
    output_xlsx_path = os.path.join(workdir, "output.xlsx")

    baseline_path = None

    # Save uploaded files
    with open(template_path, "wb") as f:
        shutil.copyfileobj(template.file, f)

    with open(projects_path, "wb") as f:
        shutil.copyfileobj(projects.file, f)

    with open(owners_path, "wb") as f:
        shutil.copyfileobj(owners.file, f)

    if baseline:
        # Keep original extension (.xls or .xlsx)
        ext = os.path.splitext(baseline.filename)[1].lower() or ".xlsx"
        baseline_path = os.path.join(workdir, "baseline" + ext)
        with open(baseline_path, "wb") as f:
            shutil.copyfileobj(baseline.file, f)

    logger.info("Starting generation process")

    generate_sre(
        template_path=template_path,
        projects_path=projects_path,
        owners_path=owners_path,
        baseline_path=baseline_path,
        output_path=output_xlsx_path
    )

    logger.info("Generation process completed")

    xls_path = os.path.join(workdir, "output.xls")
    logger.info(f"Sending file to client: {xls_path}")

    return FileResponse(
        path=xls_path,
        filename="output.xls",
        media_type="application/vnd.ms-excel",
        headers={
            "Content-Disposition": 'attachment; filename="output.xls"'
        }
    )
