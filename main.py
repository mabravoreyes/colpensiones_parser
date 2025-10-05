from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import tempfile
import os
import pandas as pd
from typing import Dict, Any
import logging

# Import your extractor
from pdf_table_extractor import ColpensionesUnifiedExtractor

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="Colpensiones PDF Parser API")

# Configure CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
async def root():
    return {"status": "ok", "message": "Colpensiones PDF Parser API"}

@app.get("/health")
async def health():
    return {"status": "healthy"}

@app.post("/parse-pension-pdf")
async def parse_pension_pdf(file: UploadFile = File(...)):
    """
    Parse a Colpensiones PDF and extract contribution weeks and payments data.
    
    Returns:
        {
            "weeks_data": [...],
            "summary_values": {...},
            "payments_data": [...]
        }
    """
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="File must be a PDF")
    
    # Save uploaded file to temporary location
    temp_pdf_path = None
    try:
        # Create a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
            # Write uploaded content to temp file
            content = await file.read()
            temp_file.write(content)
            temp_pdf_path = temp_file.name
        
        logger.info(f"Processing PDF: {file.filename} (saved to {temp_pdf_path})")
        
        # Extract data using your Python logic
        extractor = ColpensionesUnifiedExtractor()
        weeks_df, summary_values, payments_df = extractor.extract_all_from_pdf(temp_pdf_path)
        
        # Convert DataFrames to JSON-serializable format
        result = {
            "weeks_data": weeks_df.to_dict('records') if not weeks_df.empty else [],
            "summary_values": summary_values,
            "payments_data": payments_df.to_dict('records') if not payments_df.empty else []
        }
        
        logger.info(f"Successfully parsed PDF: {len(result['weeks_data'])} weeks, {len(result['payments_data'])} payments")
        
        return result
        
    except Exception as e:
        logger.error(f"Error parsing PDF: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Error parsing PDF: {str(e)}")
        
    finally:
        # Clean up temporary file
        if temp_pdf_path and os.path.exists(temp_pdf_path):
            try:
                os.unlink(temp_pdf_path)
            except Exception as e:
                logger.warning(f"Could not delete temporary file {temp_pdf_path}: {str(e)}")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
