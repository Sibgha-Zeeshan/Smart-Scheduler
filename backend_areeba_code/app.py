from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
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
    """Upload Excel file containing timetable data"""
    
    # Clear uploads directory
    if os.path.exists("uploads"):
        shutil.rmtree("uploads")
    os.makedirs("uploads", exist_ok=True)
    
    # Save uploaded Excel file
    excel_path = os.path.join("uploads", "Data.xlsx")
    with open(excel_path, "wb") as f:
        content = await excel_file.read()
        f.write(content)
    
    return {"message": "Excel file uploaded successfully"}

@app.post("/validate-excel/")
async def validate_excel():
    """Validate the uploaded Excel file and generate CSV files"""
    
    excel_path = os.path.join("uploads", "Data.xlsx")
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
async def generate_timetable(background_tasks: BackgroundTasks):
    """Generate timetable based on validated data"""
    
    # Check if validated files exist
    required_files = ["Faculty.csv", "Rooms.csv", "Time Slots.csv", "Courses.csv"]
    for file in required_files:
        if not os.path.exists(os.path.join("validated", file)):
            raise HTTPException(status_code=400, detail=f"Missing validated file: {file}. Please upload and validate Excel file first.")
    
    # Run the timetable generation script
    try:
        # Run the script as a background task
        background_tasks.add_task(run_timetable_generation)
        return {"message": "Timetable generation started. Use the /status endpoint to check progress."}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error during timetable generation: {str(e)}")

def run_timetable_generation():
    """Run the timetable generation script as a subprocess"""
    subprocess.run(["python", "CSP3.py"], check=True)
    
    # Copy output files to output directory
    if os.path.exists("timetable-section.xlsx"):
        shutil.copy("timetable-section.xlsx", os.path.join("output", "timetable-section.xlsx"))
    if os.path.exists("conflicts.csv"):
        shutil.copy("conflicts.csv", os.path.join("output", "conflicts.csv"))

@app.get("/status/")
async def check_status():
    """Check the status of timetable generation"""
    
    status = {
        "timetable_generated": os.path.exists(os.path.join("output", "timetable-section.xlsx")),
        "conflicts_file": os.path.exists(os.path.join("output", "conflicts.csv")),
        "files_available": []
    }
    
    if os.path.exists("output"):
        status["files_available"] = os.listdir("output")
    
    return status

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