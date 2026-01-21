import os
import shutil
import uuid
import json
from pathlib import Path
from typing import Dict, Optional, List
from fastapi import FastAPI, UploadFile, File, BackgroundTasks, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv

# Load context
load_dotenv()

# Import existing pipeline logic
import sys
project_root = Path(__file__).parent.parent
sys.path.append(str(project_root))

from llm_pipeline import LLMHVACPipeline
from extractors.excel_populator import HVACExcelPopulator

app = FastAPI(title="HVAC Extraction API")

# CORS for frontend integration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Storage directories
UPLOAD_DIR = project_root / "uploads"
OUTPUT_DIR = project_root / "output"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# In-memory job tracking (use Redis/DB for production)
jobs: Dict[str, Dict] = {}

def process_hvac_task(job_id: str, pdf_path: str, template_path: Optional[str] = None):
    """Background task to run the extraction pipeline"""
    try:
        jobs[job_id]["status"] = "processing"
        
        # Initialize pipeline
        output_xlsx = OUTPUT_DIR / f"hvac_report_{job_id}.xlsx"
        pipeline = LLMHVACPipeline(
            pdf_path=str(pdf_path),
            output_path=str(output_xlsx),
            job_number="1168",
            project_name="Boeing Arlington R&D"
        )
        
        # 1. Extract data
        jobs[job_id]["step"] = "extracting_data"
        extracted_data = pipeline.extract()
        
        # 2. Save JSON for reference
        json_path = OUTPUT_DIR / f"data_{job_id}.json"
        with open(json_path, "w") as f:
            json.dump(extracted_data, f, indent=2)
            
        # 3. Populate original template if provided
        if template_path and Path(template_path).exists():
            jobs[job_id]["step"] = "populating_template"
            populated_path = OUTPUT_DIR / f"populated_{job_id}.xlsx"
            populator = HVACExcelPopulator(str(template_path))
            populator.populate_all(extracted_data)
            populator.save(str(populated_path))
            jobs[job_id]["populated_file"] = f"populated_{job_id}.xlsx"
        
        # 4. Generate the new report (default behavior of pipeline)
        jobs[job_id]["step"] = "generating_report"
        pipeline.generate_excel()
        
        jobs[job_id]["status"] = "completed"
        jobs[job_id]["step"] = "done"
        jobs[job_id]["result_file"] = f"hvac_report_{job_id}.xlsx"
        jobs[job_id]["data_file"] = f"data_{job_id}.json"
        
    except Exception as e:
        jobs[job_id]["status"] = "failed"
        jobs[job_id]["error"] = str(e)
        print(f"Error processing job {job_id}: {e}")

@app.post("/extract")
async def extract_hvac(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    template: Optional[UploadFile] = File(None)
):
    """Endpoint to start an extraction job"""
    job_id = str(uuid.uuid4())
    pdf_path = UPLOAD_DIR / f"{job_id}_{file.filename}"
    
    with pdf_path.open("wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
        
    template_path = None
    if template:
        template_path = UPLOAD_DIR / f"{job_id}_{template.filename}"
        with template_path.open("wb") as buffer:
            shutil.copyfileobj(template.file, buffer)
            
    jobs[job_id] = {
        "id": job_id,
        "status": "queued",
        "filename": file.filename,
        "timestamp": uuid.uuid4().hex # placeholder for actual time if needed
    }
    
    background_tasks.add_task(process_hvac_task, job_id, pdf_path, template_path)
    
    return {"job_id": job_id, "status": "queued"}

@app.get("/status/{job_id}")
async def get_status(job_id: str):
    """Check status of a job"""
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    
    return jobs[job_id]

@app.get("/download/{filename}")
async def download_file(filename: str):
    """Download a generated file"""
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File not found")
    
    return FileResponse(
        path=file_path,
        filename=filename,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == "__main__":
    import uvicorn
    # Use PORT from environment for Render deployment
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
