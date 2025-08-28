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
    
    # Validate file type
    if not file.filename.endswith(('.xlsx', '.xlsm')):
        raise HTTPException(
            status_code=400, 
            detail="File must be Excel format (.xlsx or .xlsm)"
        )
    
    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            temp_path = Path(temp_dir)
            input_file = temp_path / file.filename
            
            # Save uploaded file
            content = await file.read()
            input_file.write_bytes(content)
            
            # Run optimization
            output_path = optimise_cuts(input_file, kerf=kerf)
            
            if not output_path.exists():
                raise HTTPException(
                    status_code=500, 
                    detail="Optimization failed - no output generated"
                )
            
            # Return optimized file
            result_bytes = BytesIO(output_path.read_bytes())
            
            return StreamingResponse(
                result_bytes,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                headers={"Content-Disposition": f'attachment; filename="{return_name}"'}
            )
            
        except Exception as e:
            # Detailed error for debugging
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