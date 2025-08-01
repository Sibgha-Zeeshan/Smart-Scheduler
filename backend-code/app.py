from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks, Form, Query, Body
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import os
import shutil
import subprocess
import time
import sys
import importlib.util
import asyncio
from pymongo import MongoClient
from passlib.context import CryptContext
import datetime
from pydantic import BaseModel
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import io

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

# MongoDB setup
MONGO_URL = "mongodb+srv://dbuser:dbuserpassword@cluster0.xm91idu.mongodb.net/"
client = MongoClient(MONGO_URL)
db = client["smartSchedule"]
register_collection = db["register"]

# Add pending_verifications collection for email verification
pending_verifications_collection = db["pending_verifications"]

import random

def generate_verification_code():
    return str(random.randint(1000, 9999))

# Password hashing context
pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

# --- Admin DB setup ---
admin_client = MongoClient(MONGO_URL)
admin_db = admin_client["admin1"]
admin_login_collection = admin_db["login"]
admin_request_collection = admin_db["request"]
admin_history_collection = admin_db["history"]

# Ensure default admin exists
if not admin_login_collection.find_one({"email": "umtadmin@umt.edu.pk"}):
    admin_login_collection.insert_one({"email": "umtadmin@umt.edu.pk", "password": pwd_context.hash("admin123")})

@app.post("/upload-excel/")
async def upload_excel_file(excel_file: UploadFile = File(...)):
    """Upload Excel file, keeping only the latest file in the uploads directory."""
    uploads_dir = "uploads"
    # Delete all previous files in uploads
    for f in os.listdir(uploads_dir):
        file_path = os.path.join(uploads_dir, f)
        if os.path.isfile(file_path):
            os.remove(file_path)
    orig_name = os.path.splitext(os.path.basename(excel_file.filename))[0]
    ext = os.path.splitext(excel_file.filename)[1]
    filename = orig_name + ext
    file_path = os.path.join("uploads", filename)
    # No need to add _1, _2, etc. since folder is empty
    with open(file_path, "wb") as f:
        content = await excel_file.read()
        f.write(content)
    return {"message": "Excel file uploaded successfully", "filename": filename}

@app.post("/validate-excel/")
async def validate_excel(filename: str = Form(...)):
    """Validate the uploaded Excel file and generate CSV files, returning all validation messages from v.py."""
    excel_path = os.path.join("uploads", filename)
    if not os.path.exists(excel_path):
        raise HTTPException(status_code=400, detail="Excel file not found. Please upload a file first.")
    try:
        # Import the v.py module dynamically
        spec = importlib.util.spec_from_file_location("v", "v.py")
        v_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(v_module)
        # Capture stdout to get all print statements from v.py
        import sys
        import io
        log_capture = io.StringIO()
        original_stdout = sys.stdout
        sys.stdout = log_capture
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
        validation_log = log_capture.getvalue()
        return {
            "message": "Excel data validated and CSV files generated successfully",
            "validation_log": validation_log
        }
    except Exception as e:
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
    return {"message": "Timetable generated successfully.", "output_file": output_name}

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

@app.get("/timetable-data/")
async def get_timetable_data(file_name: str = Query(...)):
    """Return the first 10 rows of the timetable as a list of columns and rows for dashboard viewing."""
    file_path = os.path.join("output", file_name)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Timetable file not found.")
    df = pd.read_excel(file_path, header=1)
    df = df.where(pd.notnull(df), None)
    columns = list(df.columns)
    rows = df.head(10).values.tolist()
    return {"columns": columns, "rows": rows}

@app.get("/latest-timetable-filename/")
async def get_latest_timetable_filename():
    output_dir = "output"
    timetable_files = [
        f for f in os.listdir(output_dir)
        if f.endswith("_timetable.xlsx")
    ]
    if not timetable_files:
        raise HTTPException(status_code=404, detail="No timetable files found.")
    latest_file = max(
        timetable_files,
        key=lambda f: os.path.getmtime(os.path.join(output_dir, f))
    )
    return {"latest_file": latest_file}

@app.post("/register/")
async def register_user(
    email: str = Body(...),
    username: str = Body(...),
    password: str = Body(...),
    role: str = Body(...)
):
    # Check if user already exists (by email or username)
    if register_collection.find_one({"$or": [{"email": email}, {"username": username}]}):
        raise HTTPException(status_code=400, detail="User with this email or username already exists.")
    
    # Check if there's already a pending verification for this email/username
    existing_pending = pending_verifications_collection.find_one({"$or": [{"email": email}, {"username": username}]})
    if existing_pending:
        # Remove the existing pending verification to create a new one
        pending_verifications_collection.delete_one({"_id": existing_pending["_id"]})
    
    # Hash the password
    hashed_password = pwd_context.hash(password)
    # Generate verification code
    code = generate_verification_code()
    # Store in pending_verifications
    pending_verifications_collection.insert_one({
        "email": email,
        "username": username,
        "password": hashed_password,
        "role": role,
        "code": code
    })
    # Send verification email
    try:
        send_email(
            email,
            "Your Verification Code",
            f"Hello {username},\n\nYour verification code is: {code}\n\nPlease enter this code to verify your email and complete signup.\n\nBest regards,\nAdmin Team"
        )
    except Exception as e:
        print("Failed to send verification email:", e)
        raise HTTPException(status_code=500, detail="Failed to send verification email.")
    return {"message": "Verification code sent to your email. Please verify to complete signup."}

class EmailVerificationRequest(BaseModel):
    email: str
    code: str

@app.post("/verify-email/")
async def verify_email(req: EmailVerificationRequest):
    # Find pending verification
    pending = pending_verifications_collection.find_one({"email": req.email})
    if not pending:
        raise HTTPException(status_code=404, detail="No pending verification found for this email.")
    if pending["code"] != req.code:
        raise HTTPException(status_code=400, detail="Invalid verification code.")
    # Move user to register and admin.request
    user_data = {
        "email": pending["email"],
        "username": pending["username"],
        "password": pending["password"],
        "role": pending["role"]
    }
    register_collection.insert_one(user_data)
    admin_request_collection.insert_one({
        "email": pending["email"],
        "username": pending["username"],
        "role": pending["role"]
    })
    # Remove from pending_verifications
    pending_verifications_collection.delete_one({"_id": pending["_id"]})
    return {"message": "Email verified and your request is pending. When accepted or rejected, you will get an email."}

@app.post("/login/")
async def login_user(
    email: str = Body(...),
    password: str = Body(...)
):
    # If user is still in admin.request, their request is pending
    if admin_request_collection.find_one({"email": email}):
        raise HTTPException(status_code=403, detail="Your request is pending.")
    # If user was rejected
    if admin_history_collection.find_one({"email": email, "action": "rejected"}):
        raise HTTPException(status_code=403, detail="You are not eligible.")
    user = register_collection.find_one({"email": email})
    if not user:
        raise HTTPException(status_code=404, detail="User not registered.")
    if not pwd_context.verify(password, user["password"]):
        raise HTTPException(status_code=401, detail="Incorrect password.")
    # Return user's email, username, and role
    return {
        "message": "Login successful.",
        "email": user["email"],
        "username": user["username"],
        "role": user["role"]
    }

@app.post("/admin-login/")
async def admin_login(email: str = Body(...), password: str = Body(...)):
    admin = admin_login_collection.find_one({"email": email})
    if not admin or not pwd_context.verify(password, admin["password"]):
        raise HTTPException(status_code=401, detail="Invalid admin credentials.")
    return {"message": "Admin login successful."}

@app.get("/admin/requests/")
async def get_requests():
    requests = list(admin_request_collection.find({}, {"_id": 0}))
    return {"requests": requests}

class EmailRequest(BaseModel):
    email: str

def send_email(to_email, subject, body):
    sender_email = "hello19213141@gmail.com"
    sender_password = "gkhndyphfhmcurig" 
    from_name = "UMT Admin"
    msg = MIMEMultipart()
    msg["From"] = f"{from_name} <{sender_email}>"
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, to_email, msg.as_string())

@app.post("/admin/approve/")
async def approve_request(req: EmailRequest):
    email = req.email
    req_doc = admin_request_collection.find_one_and_delete({"email": email})
    if not req_doc:
        raise HTTPException(status_code=404, detail="Request not found.")
    admin_history_collection.insert_one({
        "email": req_doc["email"],
        "username": req_doc["username"],
        "role": req_doc["role"],
        "action": "approved"
    })
    # Send approval email
    try:
        send_email(
            req_doc["email"],
            "Your Registration Has Been Approved",
            f"Hello {req_doc['username']},\n\nYour registration has been approved. You can now log in to the system.\n\nBest regards,\nAdmin Team"
        )
    except Exception as e:
        print("Failed to send approval email:", e)
    return {"message": f"User {email} approved and notified by email."}

@app.post("/admin/reject/")
async def reject_request(req: EmailRequest):
    email = req.email
    req_doc = admin_request_collection.find_one_and_delete({"email": email})
    if not req_doc:
        raise HTTPException(status_code=404, detail="Request not found.")
    admin_history_collection.insert_one({
        "email": req_doc["email"],
        "username": req_doc["username"],
        "role": req_doc["role"],
        "action": "rejected"
    })
    register_collection.delete_one({"email": email})
    # Send rejection email
    try:
        send_email(
            req_doc["email"],
            "Your Registration Has Been Rejected",
            f"Hello {req_doc['username']},\n\nWe regret to inform you that your registration request has been rejected.\n\nBest regards,\nAdmin Team"
        )
    except Exception as e:
        print("Failed to send rejection email:", e)
    return {"message": f"User {email} rejected, removed from requests, and notified by email."}

@app.get("/admin/history/")
async def get_history(current_admin_email: str = Query(None)):
    """
    Get user history. If current_admin_email is provided,
    it excludes that user from the history list.
    """
    query = {}
    if current_admin_email:
        query = {"email": {"$ne": current_admin_email}}
    
    history = list(admin_history_collection.find(query, {"_id": 0}))
    return {"history": history}

class HistoryEditRequest(BaseModel):
    email: str
    action: str
    role: str = None  # Optional, for editing role

class HistoryRemoveRequest(BaseModel):
    email: str

@app.post("/admin/history/remove/")
async def remove_history(req: HistoryRemoveRequest):
    # Remove the most recent record for this email
    record = admin_history_collection.find_one({"email": req.email}, sort=[("_id", -1)])
    if not record:
        raise HTTPException(status_code=404, detail="History record not found.")
    result = admin_history_collection.delete_one({"_id": record["_id"]})
    # Also remove the user from the main user collection
    reg_result = register_collection.delete_one({"email": req.email})
    if result.deleted_count == 1:
        return {"message": "History record and user removed."}
    else:
        raise HTTPException(status_code=404, detail="History record not found.")

@app.post("/admin/history/edit/")
async def edit_history(req: HistoryEditRequest):
    # Edit the most recent record for this email
    record = admin_history_collection.find_one({"email": req.email}, sort=[("_id", -1)])
    if not record:
        raise HTTPException(status_code=404, detail="History record not found.")
    old_action = record["action"]
    update_fields = {"action": req.action}
    role_changed = False
    if req.role is not None:
        update_fields["role"] = req.role
        # Also update the role in the main user collection
        reg_result = register_collection.update_one({"email": req.email}, {"$set": {"role": req.role}})
        role_changed = reg_result.modified_count == 1
    result = admin_history_collection.update_one(
        {"_id": record["_id"]},
        {"$set": update_fields}
    )
    if result.modified_count == 1 or role_changed:
        # Send email if action changed from approved to rejected or vice versa
        if old_action != req.action:
            try:
                reg_user = register_collection.find_one({"email": req.email})
                username = reg_user["username"] if reg_user and "username" in reg_user else req.email
                if req.action == "rejected":
                    send_email(
                        req.email,
                        "Your Registration Has Been Rejected",
                        f"Hello {username},\n\nWe regret to inform you that your registration request has been rejected.\n\nBest regards,\nAdmin Team"
                    )
                elif req.action == "approved":
                    send_email(
                        req.email,
                        "Your Registration Has Been Approved",
                        f"Hello {username},\n\nYour registration has been approved. You can now log in to the system.\n\nBest regards,\nAdmin Team"
                    )
            except Exception as e:
                print("Failed to send history edit email:", e)
        return {"message": "History record updated."}
    else:
        raise HTTPException(status_code=404, detail="History record not found or not changed.")

@app.get("/admin/timetables/")
async def list_timetables():
    """List all generated timetable files for the admin."""
    output_dir = "output"
    try:
        all_files = os.listdir(output_dir)
        timetable_files = [f for f in all_files if f.endswith("_timetable.xlsx")]
        
        files_info = []
        for f in timetable_files:
            file_path = os.path.join(output_dir, f)
            files_info.append({
                "filename": f,
                "modified_time": os.path.getmtime(file_path)
            })
        
        # Sort by most recently modified
        files_info.sort(key=lambda x: x["modified_time"], reverse=True)
        
        return {"timetables": files_info}
    except FileNotFoundError:
        return {"timetables": []}


@app.delete("/admin/timetables/{filename}")
async def delete_timetable(filename: str):
    """Delete a specific timetable file for the admin."""
    output_dir = "output"
    file_path = os.path.join(output_dir, filename)
    
    # Basic security check to prevent path traversal
    if not os.path.abspath(file_path).startswith(os.path.abspath(output_dir)):
        raise HTTPException(status_code=400, detail="Invalid filename.")

    if os.path.exists(file_path):
        os.remove(file_path)
        return {"message": f"Timetable '{filename}' deleted successfully."}
    else:
        raise HTTPException(status_code=404, detail="File not found.")

@app.get("/download-template/")
async def download_template():
    """Download the Excel template for input."""
    template_path = os.path.join("template", "timetable_template.xlsx")
    if not os.path.exists(template_path):
        raise HTTPException(status_code=404, detail="Template file not found.")
    return FileResponse(
        path=template_path,
        filename="timetable_template.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

class ResendCodeRequest(BaseModel):
    email: str

@app.post("/resend-code/")
async def resend_code(req: ResendCodeRequest):
    # Find pending verification
    pending = pending_verifications_collection.find_one({"email": req.email})
    if not pending:
        raise HTTPException(status_code=404, detail="No pending verification found for this email.")
    # Generate a new code
    new_code = generate_verification_code()
    # Update the code in the database
    pending_verifications_collection.update_one(
        {"email": req.email},
        {"$set": {"code": new_code}}
    )
    # Send the new code via email
    try:
        send_email(
            req.email,
            "Your New Verification Code",
            f"Hello {pending['username']},\n\nYour new verification code is: {new_code}\n\nPlease enter this code to verify your email and complete signup.\n\nBest regards,\nAdmin Team"
        )
    except Exception as e:
        print("Failed to send new verification email:", e)
        raise HTTPException(status_code=500, detail="Failed to send new verification email.")
    return {"message": "A new verification code has been sent to your email."}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000) 
