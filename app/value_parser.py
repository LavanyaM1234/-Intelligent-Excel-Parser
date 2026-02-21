"""Functions to parse and convert raw cell values to numeric formats"""

import re
from typing import Optional, Tuple

# Parameters that must be numeric; booleans should not be accepted for these
NUMERIC_PARAMS = [
    "coal_consumption",
    "steam_generation",
    "power_generation",
    "efficiency",
]


def parse_value(raw_value: any, param_name: str = None) -> Tuple[Optional[float], str]:
    """
    Parse a raw cell value to numeric format.
    Returns (parsed_value, confidence_note)
    """
    
    # Handle None/empty values
    if raw_value is None or raw_value == "" or raw_value == "N/A" or raw_value == "NA":
        return None, "null_value"
    
    # If already numeric, return it
    if isinstance(raw_value, (int, float)):
        value = float(raw_value)
        # If this looks like an efficiency given as a percent integer (e.g., 85), convert to 0.85
        if param_name and "efficiency" in param_name.lower():
            if value > 1 and value <= 100:
                return value / 100.0, "numeric_direct_inferred_percentage"
        return value, "numeric_direct"
    
    # Convert to string and clean
    value_str = str(raw_value).strip().upper()
    
    # Handle boolean-like values
    if value_str in ["YES", "TRUE", "Y", "1"]:
        # If this parameter is expected numeric, treat boolean as unexpected
        if param_name and any(p in param_name.lower() for p in NUMERIC_PARAMS):
            return None, "unexpected_boolean"
        return 1.0, "boolean_true"
    if value_str in ["NO", "FALSE", "N", "0"]:
        if param_name and any(p in param_name.lower() for p in NUMERIC_PARAMS):
            return None, "unexpected_boolean"
        return 0.0, "boolean_false"
    
    # Remove percentage sign and convert
    if "%" in value_str:
        try:
            num_str = value_str.replace("%", "").strip()
            value = float(num_str) / 100  # Convert % to decimal
            return value, "percentage"
        except ValueError:
            pass
    
    # Remove commas and convert
    if "," in value_str:
        try:
            cleaned = value_str.replace(",", "").strip()
            value = float(cleaned)
            return value, "numeric_with_commas"
        except ValueError:
            pass
    
    # Try direct float conversion
    try:
        value = float(value_str)
        # If param suggests efficiency and value appears as percentage without % sign, infer
        if param_name and "efficiency" in param_name.lower():
            if value > 1 and value <= 100:
                return value / 100.0, "numeric_string_inferred_percentage"
        return value, "numeric_direct_string"
    except ValueError:
        pass
    
    # Fallback: try to extract first number sequence
    match = re.search(r'[-+]?\d*\.?\d+', value_str)
    if match:
        try:
            value = float(match.group())
            return value, "extracted_number"
        except ValueError:
            pass
    
    return None, "unparseable"


def validate_value(parsed_value: Optional[float], param_name: str = None) -> Tuple[bool, Optional[str]]:
    """
    Validate if a parsed value is reasonable for the given parameter.
    Returns (is_valid, warning_message)
    """
    
    if parsed_value is None:
        return True, None  # Nulls are fine
    
    # Efficiency should be between 0 and 1 (or 0-100 if percentage)
    if param_name and "efficiency" in param_name.lower():
        if parsed_value < 0 or parsed_value > 1:
            if parsed_value > 100:
                return False, f"Efficiency value {parsed_value} exceeds 100%"
            return True, f"Efficiency {parsed_value} - verify if this is decimal or percentage"
    
    # Coal consumption should be positive
    if param_name and "coal" in param_name.lower():
        if parsed_value < 0:
            return False, f"Coal consumption cannot be negative: {parsed_value}"
    
    # Power should be positive
    if param_name and "power" in param_name.lower():
        if parsed_value < 0:
            return False, f"Power generation cannot be negative: {parsed_value}"
    
    # Operating hours should be reasonable (0-24)
    if param_name and "hour" in param_name.lower():
        if parsed_value < 0 or parsed_value > 24:
            return False, f"Operating hours {parsed_value} outside expected range (0-24)"
    
    return True, None
