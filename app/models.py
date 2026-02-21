"""Pydantic models for request/response handling"""

from pydantic import BaseModel
from typing import Optional, List, Dict, Any


class ParsedCell(BaseModel):
    """Represents a parsed cell with all metadata"""
    row: int
    col: int
    param_name: str
    asset_name: Optional[str] = None
    raw_value: Any
    parsed_value: Optional[float] = None
    confidence: str  # "high", "medium", "low"
    notes: Optional[str] = None


class UnmappedColumn(BaseModel):
    """Represents a column that couldn't be mapped"""
    col: int
    header: str
    reason: str


class ParseResponse(BaseModel):
    """API response for Excel parsing"""
    status: str  # "success" or "error"
    header_row: Optional[int] = None
    parsed_data: List[ParsedCell] = []
    unmapped_columns: List[UnmappedColumn] = []
    warnings: List[str] = []
    errors: List[str] = []
    metadata: Dict[str, Any] = {}
    # Added fields for improved clarity
    parameters: Dict[str, List[str]] = {}  # param_name -> [asset_names]
    detected_assets: List[str] = []  # List of detected assets
    units: Dict[str, str] = {}  # param_name -> unit


class MappingResult(BaseModel):
    """Result of LLM mapping for a column header"""
    header: str
    param_name: Optional[str] = None
    asset_name: Optional[str] = None
    confidence: str  # "high", "medium", "low"
    reason: str
