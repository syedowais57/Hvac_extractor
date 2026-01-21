import os
import shutil
import uuid
import json
from pathlib import Path
from typing import Dict, Optional, List
from fastapi import FastAPI, UploadFile, File, BackgroundTasks, HTTPException
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse
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

# HTML UI Template
HOME_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>HVAC Data Extractor</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #4F46E5;
            --primary-hover: #4338CA;
            --bg: #0F172A;
            --card: #1E293B;
            --text: #F8FAFC;
            --text-muted: #94A3B8;
        }
        body {
            font-family: 'Inter', sans-serif;
            background-color: var(--bg);
            color: var(--text);
            margin: 0;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
        }
        .container {
            width: 100%;
            max-width: 500px;
            padding: 2rem;
            background: var(--card);
            border-radius: 1.5rem;
            box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
            border: 1px solid rgba(255, 255, 255, 0.1);
        }
        h1 { font-size: 1.875rem; font-weight: 700; margin-bottom: 0.5rem; text-align: center; }
        p { color: var(--text-muted); text-align: center; margin-bottom: 2rem; }
        .field { margin-bottom: 1.5rem; }
        label { display: block; font-size: 0.875rem; font-weight: 600; margin-bottom: 0.5rem; color: var(--text-muted); }
        input[type="file"] {
            width: 100%;
            padding: 0.75rem;
            background: #334155;
            border-radius: 0.75rem;
            border: 2px dashed #475569;
            color: var(--text);
            cursor: pointer;
            box-sizing: border-box;
        }
        button {
            width: 100%;
            padding: 1rem;
            background: var(--primary);
            color: white;
            border: none;
            border-radius: 0.75rem;
            font-size: 1rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
        }
        button:hover { background: var(--primary-hover); transform: translateY(-1px); }
        button:disabled { background: #475569; cursor: not-allowed; }
        #status-area {
            margin-top: 2rem;
            padding: 1rem;
            background: #334155;
            border-radius: 1rem;
            display: none;
        }
        .status-header { font-weight: 600; margin-bottom: 0.5rem; display: flex; justify-content: space-between; }
        .step { font-size: 0.875rem; color: var(--text-muted); }
        .loader {
            height: 4px;
            width: 100%;
            background: #1e293b;
            border-radius: 2px;
            overflow: hidden;
            margin-top: 1rem;
        }
        .loader-bar {
            height: 100%;
            width: 30%;
            background: var(--primary);
            animation: loading 1.5s infinite ease-in-out;
        }
        @keyframes loading {
            0% { transform: translateX(-100%); }
            100% { transform: translateX(400%); }
        }
        .download-links { margin-top: 1rem; display: flex; flex-direction: column; gap: 0.5rem; }
        .download-link {
            display: block;
            padding: 0.75rem;
            background: #10B981;
            color: white;
            text-decoration: none;
            text-align: center;
            border-radius: 0.5rem;
            font-weight: 600;
        }
        .download-link:hover { background: #059669; }
    </style>
</head>
<body>
    <div class="container">
        <h1>HVAC Extractor</h1>
        <p>Extract data from PDF drawings to Excel</p>
        
        <div class="field">
            <label>PDF Drawings (Required)</label>
            <input type="file" id="pdf-file" accept=".pdf">
        </div>
        
        <div class="field">
            <label>Original Excel Template (Optional)</label>
            <input type="file" id="excel-template" accept=".xlsx">
        </div>
        
        <button id="upload-btn">Start Extraction</button>
        
        <div id="status-area">
            <div class="status-header">
                <span id="status-text">Processing...</span>
            </div>
            <div id="step-text" class="step">Initializing...</div>
            <div id="loader" class="loader"><div class="loader-bar"></div></div>
            <div id="download-area" class="download-links"></div>
        </div>
    </div>

    <script>
        const uploadBtn = document.getElementById('upload-btn');
        const pdfInput = document.getElementById('pdf-file');
        const excelInput = document.getElementById('excel-template');
        const statusArea = document.getElementById('status-area');
        const statusText = document.getElementById('status-text');
        const stepText = document.getElementById('step-text');
        const downloadArea = document.getElementById('download-area');
        const loader = document.getElementById('loader');

        uploadBtn.onclick = async () => {
            const pdfFile = pdfInput.files[0];
            if (!pdfFile) { alert('Please select a PDF file'); return; }

            uploadBtn.disabled = true;
            statusArea.style.display = 'block';
            downloadArea.innerHTML = '';
            loader.style.display = 'block';

            const formData = new FormData();
            formData.append('file', pdfFile);
            if (excelInput.files[0]) formData.append('template', excelInput.files[0]);

            try {
                const response = await fetch('/extract', { method: 'POST', body: formData });
                const { job_id } = await response.json();
                pollStatus(job_id);
            } catch (err) {
                alert('Upload failed');
                uploadBtn.disabled = false;
            }
        };

        async function pollStatus(jobId) {
            const interval = setInterval(async () => {
                const res = await fetch(`/status/${jobId}`);
                const data = await res.json();
                
                statusText.innerText = data.status.charAt(0).toUpperCase() + data.status.slice(1);
                stepText.innerText = data.step ? data.step.replace(/_/g, ' ') : 'Wait...';

                if (data.status === 'completed') {
                    clearInterval(interval);
                    loader.style.display = 'none';
                    uploadBtn.disabled = false;
                    
                    let html = `<a href="/download/${data.result_file}" class="download-link">Download New Report</a>`;
                    if (data.populated_file) {
                        html += `<a href="/download/${data.populated_file}" class="download-link">Download Populated Template</a>`;
                    }
                    downloadArea.innerHTML = html;
                } else if (data.status === 'failed') {
                    clearInterval(interval);
                    loader.style.display = 'none';
                    uploadBtn.disabled = false;
                    stepText.innerText = 'Error: ' + data.error;
                }
            }, 3000);
        }
    </script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
async def home():
    """Root endpoint returning the Web UI"""
    return HOME_HTML

if __name__ == "__main__":
    import uvicorn
    # Use PORT from environment for Render deployment
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
