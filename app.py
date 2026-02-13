from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from fastapi import Request
from fastapi.staticfiles import StaticFiles
import shutil
import os
import uuid

from generator import generate_sre

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

    template_path = os.path.join(workdir, "template.xlsx")
    projects_path = os.path.join(workdir, "projects.xlsx")
    owners_path = os.path.join(workdir, "owners.xlsx")
    output_path = os.path.join(workdir, "output.xlsx")

    baseline_path = None

    # Save uploaded files
    with open(template_path, "wb") as f:
        shutil.copyfileobj(template.file, f)

    with open(projects_path, "wb") as f:
        shutil.copyfileobj(projects.file, f)

    with open(owners_path, "wb") as f:
        shutil.copyfileobj(owners.file, f)

    if baseline:
        baseline_path = os.path.join(workdir, "baseline.xlsx")
        with open(baseline_path, "wb") as f:
            shutil.copyfileobj(baseline.file, f)

    # Run generator
    generate_sre(
        template_path=template_path,
        projects_path=projects_path,
        owners_path=owners_path,
        baseline_path=baseline_path,
        output_path=output_path
    )

    return FileResponse(
        output_path,
        filename="output.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
