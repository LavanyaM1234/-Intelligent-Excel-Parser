"""FastAPI application with Excel parsing endpoints"""

from fastapi import FastAPI, File, UploadFile, HTTPException, Query
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import logging
import os
from pathlib import Path

# Load environment variables from .env file
from dotenv import load_dotenv
load_dotenv(Path(__file__).parent.parent / ".env")

from app.parser import ExcelParser
from app.models import ParseResponse

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Create FastAPI app
app = FastAPI(
    title="Intelligent Excel Parser",
    description="AI-powered Excel parsing service for factory operational data",
    version="1.0.0"
)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Initialize parser
parser = ExcelParser()


@app.get("/health")
async def health_check():
    """Health check endpoint"""
    return {"status": "healthy", "service": "Excel Parser API"}


@app.post("/parse", response_model=ParseResponse)
async def parse_excel(
    file: UploadFile = File(...),
    sheet_name: str = Query(None, description="Optional sheet name to parse")
):
    """
    Parse an Excel file and extract structured data.
    
    - **file**: Excel file (.xlsx) to parse
    - **sheet_name**: Optional specific sheet to parse (defaults to first sheet)
    
    Returns JSON with:
    - parsed_data: Array of cells with parameter mapping and parsed values
    - unmapped_columns: Columns that couldn't be mapped to parameters
    - warnings: Any warnings or issues encountered
    """
    
    try:
        # Read file content
        content = await file.read()
        
        if not content:
            raise HTTPException(status_code=400, detail="File is empty")
        
        if not file.filename.lower().endswith('.xlsx'):
            raise HTTPException(status_code=400, detail="File must be .xlsx format")
        
        # Parse the Excel file
        response = parser.parse_file(content, sheet_name)
        
        return response
    
    except HTTPException:
        raise
    except Exception as e:
        logger.exception(f"Error parsing file {file.filename}")
        raise HTTPException(
            status_code=500,
            detail=f"Error processing file: {str(e)}"
        )


@app.get("/registries")
async def get_registries():
    """Get the current parameter and asset registries"""
    from app.registries import PARAMETER_REGISTRY, ASSET_REGISTRY
    return {
        "parameters": PARAMETER_REGISTRY,
        "assets": ASSET_REGISTRY
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
