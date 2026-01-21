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
JOBS_FILE = project_root / "jobs.json"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# In-memory job tracking with file persistence
jobs: Dict[str, Dict] = {}

def load_jobs():
    global jobs
    if JOBS_FILE.exists():
        try:
            with open(JOBS_FILE, "r") as f:
                jobs = json.load(f)
                print(f"Loaded {len(jobs)} jobs from storage.")
        except Exception as e:
            print(f"Error loading jobs: {e}")
            jobs = {}

def save_jobs():
    try:
        with open(JOBS_FILE, "w") as f:
            json.dump(jobs, f, indent=2)
    except Exception as e:
        print(f"Error saving jobs: {e}")

# Initial load
load_jobs()

def process_hvac_task(job_id: str, pdf_path: str, template_path: Optional[str] = None):
    """Background task to run the extraction pipeline"""
    try:
        jobs[job_id]["status"] = "processing"
        save_jobs()
        
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
        save_jobs()
        extracted_data = pipeline.extract()
        
        # 2. Save JSON for reference
        json_path = OUTPUT_DIR / f"data_{job_id}.json"
        with open(json_path, "w") as f:
            json.dump(extracted_data, f, indent=2)
            
        # 3. Populate original template if provided
        if template_path and Path(template_path).exists():
            jobs[job_id]["step"] = "populating_template"
            save_jobs()
            populated_path = OUTPUT_DIR / f"populated_{job_id}.xlsx"
            populator = HVACExcelPopulator(str(template_path))
            populator.populate_all(extracted_data)
            populator.save(str(populated_path))
            jobs[job_id]["populated_file"] = f"populated_{job_id}.xlsx"
        
        # 4. Generate the new report (default behavior of pipeline)
        jobs[job_id]["step"] = "generating_report"
        save_jobs()
        pipeline.generate_excel()
        
        jobs[job_id]["status"] = "completed"
        jobs[job_id]["step"] = "done"
        jobs[job_id]["result_file"] = f"hvac_report_{job_id}.xlsx"
        jobs[job_id]["data_file"] = f"data_{job_id}.json"
        save_jobs()
        
    except Exception as e:
        jobs[job_id]["status"] = "failed"
        jobs[job_id]["error"] = str(e)
        save_jobs()
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
        "timestamp": uuid.uuid4().hex 
    }
    save_jobs()
    
    background_tasks.add_task(process_hvac_task, job_id, pdf_path, template_path)
    
    return {"job_id": job_id, "status": "queued"}

@app.get("/status/{job_id}")
async def get_status(job_id: str):
    """Check status of a job"""
    print(jobs, job_id)
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
    <title>HVAC Extractor | Premium HVAC Data Processing</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&display=swap" rel="stylesheet">
    <style>
        body {
            font-family: 'Outfit', sans-serif;
            background: radial-gradient(circle at top right, #1e1b4b, #0f172a);
            color: #f8fafc;
            min-height: 100vh;
        }
        .glass {
            background: rgba(30, 41, 59, 0.7);
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            border: 1px solid rgba(255, 255, 255, 0.1);
        }
        .gradient-text {
            background: linear-gradient(135deg, #818cf8 0%, #c084fc 100%);
            -webkit-background-clip: text;
            background-clip: text;
            -webkit-text-fill-color: transparent;
        }
        .btn-primary {
            background: linear-gradient(135deg, #6366f1 0%, #a855f7 100%);
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }
        .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px -5px rgba(99, 102, 241, 0.4);
        }
        .drag-area.active {
            border-color: #818cf8;
            background: rgba(99, 102, 241, 0.1);
        }
        @keyframes shimmer {
            0% { transform: translateX(-100%); }
            100% { transform: translateX(100%); }
        }
        .animate-shimmer {
            animation: shimmer 2s infinite;
        }
    </style>
</head>
<body class="flex items-center justify-center p-4">
    <div class="container max-w-2xl w-full">
        <div class="glass rounded-3xl p-8 md:p-12 shadow-2xl">
            <!-- Header -->
            <div class="text-center mb-10">
                <div class="inline-block p-3 rounded-2xl bg-indigo-500/10 mb-4">
                    <svg class="w-8 h-8 text-indigo-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 11H5m14 0a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2v-6a2 2 0 012-2m14 0V9a2 2 0 00-2-2M5 11V9a2 2 0 012-2m0 0V5a2 2 0 012-2h6a2 2 0 012 2v4M7 7h10"></path>
                    </svg>
                </div>
                <h1 class="text-4xl font-bold mb-2">HVAC <span class="gradient-text">Extractor</span></h1>
                <p class="text-slate-400">Intelligent PDF to Excel Data Pipeline</p>
            </div>

            <!-- API Configuration -->
            <div class="mb-8 p-4 rounded-xl bg-slate-900/50 border border-slate-700">
                <label class="text-xs font-semibold uppercase tracking-wider text-slate-500 mb-2 block">API Endpoint</label>
                <div class="flex gap-2">
                    <input type="text" id="api-url" value="http://localhost:8000" 
                           class="bg-slate-800 border-none rounded-lg px-4 py-2 w-full text-slate-300 text-sm focus:ring-2 focus:ring-indigo-500 outline-none">
                    <button onclick="saveApiUrl()" class="px-4 py-2 bg-slate-700 hover:bg-slate-600 rounded-lg text-xs font-semibold transition-colors">Save</button>
                </div>
                <p class="text-[10px] text-slate-500 mt-2 italic">Change this after deploying to Render (e.g., https://your-app.onrender.com)</p>
            </div>

            <!-- Upload Forms -->
            <div class="space-y-6">
                <!-- PDF Upload -->
                <div>
                    <label class="text-sm font-semibold text-slate-300 mb-3 block">1. Upload Drawings (PDF)</label>
                    <div id="pdf-drop" class="drag-area border-2 border-dashed border-slate-700 rounded-2xl p-8 text-center cursor-pointer transition-all hover:border-slate-500">
                        <input type="file" id="pdf-input" accept=".pdf" class="hidden">
                        <div id="pdf-info">
                            <svg class="w-10 h-10 text-slate-500 mx-auto mb-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"></path>
                            </svg>
                            <p class="text-slate-400 text-sm">Drop your PDF here or <span class="text-indigo-400 font-semibold">Browse</span></p>
                        </div>
                    </div>
                </div>

                <!-- Template Upload -->
                <div>
                    <label class="text-sm font-semibold text-slate-300 mb-3 block">2. Original Template (Optional XLSX)</label>
                    <div id="template-drop" class="drag-area border-2 border-dashed border-slate-700 rounded-2xl p-6 text-center cursor-pointer transition-all hover:border-slate-500">
                        <input type="file" id="template-input" accept=".xlsx" class="hidden">
                        <div id="template-info">
                            <p class="text-slate-500 text-xs">Click to select Boeing Setup Excel</p>
                        </div>
                    </div>
                </div>

                <!-- Action Button -->
                <button id="start-btn" class="btn-primary w-full py-4 rounded-xl font-bold text-lg shadow-lg disabled:opacity-50 disabled:cursor-not-allowed">
                    Proceed to Extraction
                </button>
            </div>

            <!-- Progress Card -->
            <div id="progress-card" class="hidden mt-10 p-6 rounded-2xl bg-slate-900/50 border border-indigo-500/20">
                <div class="flex items-center justify-between mb-4">
                    <span class="text-sm font-semibold text-indigo-300" id="status-label">Extracting Data...</span>
                    <span class="text-xs text-slate-500" id="percent-label">Working</span>
                </div>
                <div class="h-2 w-full bg-slate-800 rounded-full overflow-hidden mb-6">
                    <div id="progress-bar" class="h-full bg-indigo-500 transition-all duration-500 w-1/3"></div>
                </div>
                
                <div id="results-area" class="hidden space-y-3">
                    <h3 class="text-xs font-bold uppercase tracking-widest text-slate-500 mb-2">Generated Reports</h3>
                    <div id="link-container" class="grid grid-cols-1 gap-3">
                        <!-- Links injected here -->
                    </div>
                </div>
            </div>
        </div>

        <!-- Footer -->
        <p class="text-center mt-8 text-slate-500 text-sm">
            Powered by Gemini 2.0 Flash & FastAPI
        </p>
    </div>

    <script>
        let BASE_URL = localStorage.getItem('hvac_api_url') || window.location.origin;
        document.getElementById('api-url').value = BASE_URL;

        function saveApiUrl() {
            BASE_URL = document.getElementById('api-url').value.replace(/\\/$/, '');
            localStorage.setItem('hvac_api_url', BASE_URL);
            alert('API URL updated: ' + BASE_URL);
        }

        const setupDragDrop = (areaId, inputId, infoId) => {
            const area = document.getElementById(areaId);
            const input = document.getElementById(inputId);
            const info = document.getElementById(infoId);

            area.onclick = () => input.click();
            area.ondragover = (e) => { e.preventDefault(); area.classList.add('active'); };
            area.ondragleave = () => area.classList.remove('active');
            area.ondrop = (e) => {
                e.preventDefault();
                area.classList.remove('active');
                input.files = e.dataTransfer.files;
                updateInfo();
            };
            input.onchange = () => updateInfo();

            function updateInfo() {
                if(input.files.length) {
                    info.innerHTML = `<span class="text-indigo-300 font-semibold">${input.files[0].name}</span>`;
                }
            }
        };

        setupDragDrop('pdf-drop', 'pdf-input', 'pdf-info');
        setupDragDrop('template-drop', 'template-input', 'template-info');

        const startBtn = document.getElementById('start-btn');
        const progressCard = document.getElementById('progress-card');
        const progressBar = document.getElementById('progress-bar');
        const statusLabel = document.getElementById('status-label');
        const resultsArea = document.getElementById('results-area');
        const linkContainer = document.getElementById('link-container');

        startBtn.onclick = async () => {
            const pdfFile = document.getElementById('pdf-input').files[0];
            const templateFile = document.getElementById('template-input').files[0];

            if(!pdfFile) return alert('Please upload a PDF drawing first.');

            startBtn.disabled = true;
            progressCard.classList.remove('hidden');
            resultsArea.classList.add('hidden');
            progressBar.classList.remove('bg-emerald-500');
            progressBar.classList.add('bg-indigo-500');
            progressBar.style.width = '20%';
            statusLabel.innerText = 'Uploading to Server...';

            const formData = new FormData();
            formData.append('file', pdfFile);
            if(templateFile) formData.append('template', templateFile);

            try {
                const response = await fetch(`${BASE_URL}/extract`, {
                    method: 'POST',
                    body: formData
                });
                const data = await response.json();
                pollStatus(data.job_id);
            } catch (err) {
                statusLabel.innerText = 'Upload Error: Verify API URL and CORS';
                startBtn.disabled = false;
            }
        };

        async function pollStatus(jobId) {
            const interval = setInterval(async () => {
                try {
                    const response = await fetch(`${BASE_URL}/status/${jobId}`);
                    const data = await response.json();

                    statusLabel.innerText = `${data.status.toUpperCase()}: ${data.step ? data.step.replace(/_/g, ' ') : ''}`;
                    
                    if(data.status === 'processing') progressBar.style.width = '60%';
                    
                    if(data.status === 'completed') {
                        clearInterval(interval);
                        progressBar.style.width = '100%';
                        progressBar.classList.remove('bg-indigo-500');
                        progressBar.classList.add('bg-emerald-500');
                        statusLabel.innerText = 'Extraction Complete!';
                        showResults(data);
                        startBtn.disabled = false;
                    } else if(data.status === 'failed') {
                        clearInterval(interval);
                        statusLabel.innerText = 'Error: ' + (data.error || 'Pipeline failed');
                        startBtn.disabled = false;
                    }
                } catch (e) {
                    clearInterval(interval);
                    statusLabel.innerText = 'Connection Lost';
                    startBtn.disabled = false;
                }
            }, 3000);
        }

        function showResults(data) {
            resultsArea.classList.remove('hidden');
            let links = `
                <a href="${BASE_URL}/download/${data.result_file}" class="px-4 py-3 bg-indigo-500/10 border border-indigo-500/30 rounded-xl text-indigo-300 hover:bg-indigo-500/20 transition-all font-semibold text-sm flex justify-between items-center">
                    <span>Summary Report</span>
                    <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
                </a>
            `;
            if(data.populated_file) {
                links += `
                    <a href="${BASE_URL}/download/${data.populated_file}" class="px-4 py-3 bg-emerald-500/10 border border-emerald-500/30 rounded-xl text-emerald-300 hover:bg-emerald-500/20 transition-all font-semibold text-sm flex justify-between items-center">
                        <span>Populated Template</span>
                        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
                    </a>
                `;
            }
            linkContainer.innerHTML = links;
        }
    </script>
</body>
</html>
"""

@app.api_route("/", methods=["GET", "HEAD"], response_class=HTMLResponse)
async def home():
    """Root endpoint returning the Web UI"""
    return HOME_HTML

if __name__ == "__main__":
    import uvicorn
    # Use PORT from environment for Render deployment
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
