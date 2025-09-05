import os
import tempfile
import traceback
from io import BytesIO
from pathlib import Path

import uvicorn
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
from starlette.middleware.cors import CORSMiddleware

# Import your existing logic
from cut_optimizer import optimise_cuts

app = FastAPI(
    title="Cut Sheet Optimizer",
    version="2.0.0",
    description="Optimizes saw cut sheets for fabrication"
)

# CORS for n8n integration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def root():
    """Health check endpoint"""
    return {
        "status": "online",
        "service": "Cut Sheet Optimizer",
        "version": "2.0.0"
    }

@app.get("/health")
def health():
    """Detailed health check for monitoring"""
    return {"status": "healthy", "timestamp": "2025-01-01T00:00:00Z"}

@app.post("/optimize")
async def optimize_cut_sheet(
    file: UploadFile = File(..., description="Fabrication Summary Excel file"),
    kerf: float = Form(0.125, description="Saw blade width in inches"),
    return_name: str = Form("Cut_sheet.xlsx", description="Output filename"),
):
    """
    Optimizes cut sheet from uploaded Fabrication Summary
    
    - **file**: Excel file (.xlsx/.xlsm) with fabrication data
    - **kerf**: Saw blade width in inches (default: 0.125)
    - **return_name**: Name for output file (default: Cut_sheet.xlsx)
    """
    
    print(f"=== DEBUG: Starting optimization ===")
    print(f"File received: {file.filename}")
    print(f"File content type: {file.content_type}")
    print(f"Kerf: {kerf}")
    print(f"Return name: {return_name}")
    
    # Validate file type
    if not file.filename.endswith(('.xlsx', '.xlsm')):
        print(f"ERROR: Invalid file type: {file.filename}")
        raise HTTPException(
            status_code=400, 
            detail="File must be Excel format (.xlsx or .xlsm)"
        )
    
    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            temp_path = Path(temp_dir)
            input_file = temp_path / file.filename
            
            print(f"Temp directory: {temp_path}")
            print(f"Input file path: {input_file}")
            
            # Save uploaded file
            content = await file.read()
            print(f"File content size: {len(content)} bytes")
            
            input_file.write_bytes(content)
            print(f"File saved successfully: {input_file.exists()}")
            print(f"Saved file size: {input_file.stat().st_size} bytes")
            
            # Test imports
            try:
                print("Testing imports...")
                from cut_sheet_loader import load_cut_demand, STICK_LENGTHS
                print("✓ cut_sheet_loader imported successfully")
                print(f"STICK_LENGTHS defined: {len(STICK_LENGTHS)} combinations")
                for key, value in STICK_LENGTHS.items():
                    print(f"  {key}: {value}")
            except Exception as import_error:
                print(f"ERROR importing cut_sheet_loader: {import_error}")
                raise HTTPException(status_code=500, detail=f"Import error: {import_error}")
            
            # Test data loading
            try:
                print("Testing data loading...")
                tidy = load_cut_demand(input_file)
                print(f"✓ Data loaded: {len(tidy)} rows")
                print(f"Columns: {list(tidy.columns)}")
                print(f"First few rows:\n{tidy.head()}")
                print(f"Unique materials: {tidy['Material'].unique()}")
                print(f"Unique diameters: {sorted(tidy['diameter_in'].unique())}")
                print(f"Rows per tab:\n{tidy.groupby('tab').size()}")
            except Exception as load_error:
                print(f"ERROR loading data: {load_error}")
                print(f"Traceback: {traceback.format_exc()}")
                raise HTTPException(status_code=400, detail=f"Data loading error: {load_error}")
            
            # Run optimization
            print("Running optimization...")
            output_path = optimise_cuts(input_file, kerf=kerf)
            print(f"✓ Optimization completed: {output_path}")
            
            if not output_path.exists():
                print(f"ERROR: Output file does not exist: {output_path}")
                raise HTTPException(
                    status_code=500, 
                    detail="Optimization failed - no output generated"
                )
            
            print(f"Output file size: {output_path.stat().st_size} bytes")
            
            # Return optimized file
            result_bytes = BytesIO(output_path.read_bytes())
            print("✓ File ready for download")
            
            return StreamingResponse(
                result_bytes,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                headers={"Content-Disposition": f'attachment; filename="{return_name}"'}
            )
            
        except Exception as e:
            # Detailed error for debugging
            print(f"=== ERROR OCCURRED ===")
            print(f"Error type: {type(e).__name__}")
            print(f"Error message: {str(e)}")
            print(f"Full traceback:\n{traceback.format_exc()}")
            
            error_detail = {
                "error": str(e),
                "type": type(e).__name__,
                "trace": traceback.format_exc()
            }
            raise HTTPException(status_code=400, detail=error_detail)

if __name__ == "__main__":
    # Railway will set PORT environment variable
    port = int(os.getenv("PORT", 8080))
    uvicorn.run("app:app", host="0.0.0.0", port=port)