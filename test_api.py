import requests
import time
import sys
from pathlib import Path

def test_extraction():
    url = "http://127.0.0.1:8000"
    pdf_path = r"D:\SW\new project\Boeing R&D Drawings.pdf"
    template_path = r"D:\SW\new project\Boeing Arlington R&D Setup.xlsx"
    
    print(f"Connecting to {url}...")
    
    # 1. Start extraction
    print("Step 1: Uploading PDF and starting extraction...")
    with open(pdf_path, "rb") as pdf_file, open(template_path, "rb") as template_file:
        files = {
            "file": (Path(pdf_path).name, pdf_file, "application/pdf"),
            "template": (Path(template_path).name, template_file, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        }
        response = requests.post(f"{url}/extract", files=files)
        
    if response.status_code != 200:
        print(f"Error starting extraction: {response.text}")
        return
        
    job_id = response.json()["job_id"]
    print(f"Job started! ID: {job_id}")
    
    # 2. Poll for status
    print("Step 2: Polling for status...")
    while True:
        status_resp = requests.get(f"{url}/status/{job_id}")
        status_data = status_resp.json()
        status = status_data["status"]
        step = status_data.get("step", "queued")
        
        print(f"  Status: {status} | Step: {step}")
        
        if status == "completed":
            print("\nExtraction complete!")
            print(f"Result file: {status_data['result_file']}")
            if "populated_file" in status_data:
                print(f"Populated file: {status_data['populated_file']}")
            break
        elif status == "failed":
            print(f"\nExtraction failed: {status_data.get('error')}")
            break
            
        time.sleep(5)
        
    # 3. Download result
    if status == "completed":
        print("\nStep 3: Downloading result...")
        result_file = status_data["result_file"]
        download_resp = requests.get(f"{url}/download/{result_file}")
        
        output_path = Path("output") / f"test_result_{job_id}.xlsx"
        output_path.parent.mkdir(exist_ok=True)
        with open(output_path, "wb") as f:
            f.write(download_resp.content)
        print(f"Downloaded to: {output_path}")

if __name__ == "__main__":
    test_extraction()
