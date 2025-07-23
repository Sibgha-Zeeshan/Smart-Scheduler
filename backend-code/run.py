#!/usr/bin/env python3
import os
import uvicorn
import argparse

def main():
    # Parse command line arguments
    parser = argparse.ArgumentParser(description="Run the Timetable Generation API")
    parser.add_argument("--host", default="0.0.0.0", help="Host to run the server on")
    parser.add_argument("--port", default=8000, type=int, help="Port to run the server on")
    parser.add_argument("--reload", action="store_true", help="Enable auto-reload")
    args = parser.parse_args()
    
    # Create necessary directories
    os.makedirs("uploads", exist_ok=True)
    os.makedirs("validated", exist_ok=True)
    os.makedirs("output", exist_ok=True)
    os.makedirs("static", exist_ok=True)
    
    # Run the API server
    print(f"Starting Timetable Generation API on http://{args.host}:{args.port}")
    print("- API Documentation: http://localhost:8000/docs")
    print("- Web Interface: http://localhost:8000/")
    
    uvicorn.run(
        "app:app", 
        host=args.host, 
        port=args.port,
        reload=args.reload
    )

if __name__ == "__main__":
    main() 