from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks, Form, Query
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import os
import shutil
import subprocess
import time
import sys
import importlib.util
import asyncio

app = FastAPI(title="Timetable Generation API")

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Create necessary directories if they don't exist
os.makedirs("uploads", exist_ok=True)
os.makedirs("validated", exist_ok=True)
os.makedirs("output", exist_ok=True)
os.makedirs("static", exist_ok=True)

# Mount static files directory
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/", response_class=FileResponse)
def read_root():
    """Serve the main HTML page"""
    return FileResponse("static/index.html")

@app.post("/upload-excel/")
async def upload_excel_file(excel_file: UploadFile = File(...)):
    """Upload Excel file using the original filename, adding _1, _2, etc. if needed."""
    os.makedirs("uploads", exist_ok=True)
    orig_name = os.path.splitext(os.path.basename(excel_file.filename))[0]
    ext = os.path.splitext(excel_file.filename)[1]
    filename = orig_name + ext
    file_path = os.path.join("uploads", filename)
    counter = 1
    while os.path.exists(file_path):
        filename = f"{orig_name}_{counter}{ext}"
        file_path = os.path.join("uploads", filename)
        counter += 1
    with open(file_path, "wb") as f:
        content = await excel_file.read()
        f.write(content)
    return {"message": "Excel file uploaded successfully", "filename": filename}

@app.post("/validate-excel/")
async def validate_excel(filename: str = Form(...)):
    """Validate the uploaded Excel file and generate CSV files"""
    excel_path = os.path.join("uploads", filename)
    if not os.path.exists(excel_path):
        raise HTTPException(status_code=400, detail="Excel file not found. Please upload a file first.")
    # Run the validation script (v.py)
    try:
        # Import the v.py module dynamically
        spec = importlib.util.spec_from_file_location("v", "v.py")
        v_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(v_module)
        # Redirect stdout to capture any print statements
        original_stdout = sys.stdout
        sys.stdout = None
        # Load and validate Excel data
        excel_data = v_module.load_excel(excel_path)
        validated_sheets = v_module.validate_data(excel_data)
        # Clear validated directory
        if os.path.exists("validated"):
            shutil.rmtree("validated")
        os.makedirs("validated", exist_ok=True)
        # Save validated data to CSV files
        v_module.save_to_csv(validated_sheets, "validated")
        # Restore stdout
        sys.stdout = original_stdout
        return {"message": "Excel data validated and CSV files generated successfully"}
    except Exception as e:
        # Restore stdout in case of exception
        if 'original_stdout' in locals():
            sys.stdout = original_stdout
        raise HTTPException(status_code=400, detail=f"Validation error: {str(e)}")

@app.post("/generate-timetable/")
async def generate_timetable(filename: str = Form(...)):
    """Start timetable generation as a background task with lock file, using the uploaded filename."""
    os.makedirs("output", exist_ok=True)
    base_name = os.path.splitext(filename)[0]
    output_name = f"{base_name}_timetable.xlsx"
    lock_path = os.path.join("output", ".generating")
    with open(lock_path, "w") as f:
        f.write("generating")

    async def run_csp3():
        timetable_path = os.path.join("output", output_name)
        if os.path.exists(timetable_path):
            os.remove(timetable_path)
        # Pass input and output filenames to CSP3.py
        try:
            subprocess.run(["python", "CSP3.py", filename, output_name], check=True)
        finally:
            if os.path.exists(lock_path):
                os.remove(lock_path)
    asyncio.create_task(run_csp3())
    return {"message": "Timetable generation started.", "output_file": output_name}

@app.get("/status/")
async def check_status(filename: str = Query(...)):
    base_name = os.path.splitext(filename)[0]
    output_name = f"{base_name}_timetable.xlsx"
    timetable_path = os.path.join("output", output_name)
    lock_path = os.path.join("output", ".generating")
    timetable_ready = os.path.exists(timetable_path) and not os.path.exists(lock_path)
    files_available = [output_name] if timetable_ready else []
    return {
        "timetable_generated": timetable_ready,
        "files_available": files_available
    }

@app.get("/download/{file_name}")
async def download_file(file_name: str):
    """Download generated files"""
    
    file_path = os.path.join("output", file_name)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail=f"File {file_name} not found")
    
    return FileResponse(
        path=file_path,
        filename=file_name,
        media_type="application/octet-stream"
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000) 