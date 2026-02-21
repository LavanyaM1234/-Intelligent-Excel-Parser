"""LLM integration for intelligent header mapping and asset detection"""

import json
from typing import Dict, List, Optional
import os
from datetime import datetime
from pathlib import Path


from dotenv import load_dotenv
load_dotenv()



from app.models import MappingResult
from app.registries import get_parameter_context, get_asset_context, PARAMETER_REGISTRY, ASSET_REGISTRY


class ExcelParsingAgent:
    """AI agent for intelligent Excel parsing and mapping"""
    
    def __init__(self, api_key: Optional[str] = None):
        """
        Initialize the parsing agent.
        
        Args:
            api_key: OpenAI API key (optional, can be from env)
        """
        # Prefer Google API key (Gemini). Fall back to OPENAI_API_KEY if present.
        self.api_key = api_key or os.getenv("GOOGLE_API_KEY") or os.getenv("OPENAI_API_KEY")
        self.param_registry = PARAMETER_REGISTRY
        self.asset_registry = ASSET_REGISTRY
        self.llm_available = self.api_key is not None
        
    def create_system_prompt(self) -> str:
        """Create the system prompt with registries"""
        return f"""You are an expert at mapping messy Excel column headers to standard parameter names and detecting assets.

{get_parameter_context()}

{get_asset_context()}

Your task is to:
1. Map a given column header to a parameter name from the registry (use exact parameter keys like "coal_consumption")
2. Detect if the header references a specific asset (use asset IDs from the registry)
3. Provide a confidence score: "high" (exact match), "medium" (fuzzy match), or "low" (uncertain)
4. Explain your reasoning

Return your response as JSON with this exact structure:
{{
    "header": "original header text",
    "param_name": "parameter_key_or_null",
    "asset_name": "asset_id_or_null",
    "confidence": "high|medium|low",
    "reason": "explanation of mapping decision"
}}
"""
    
    def map_headers(self, headers: List[str]) -> Dict[int, MappingResult]:
        """
        Map a list of column headers to parameters and assets.
        
        Args:
            headers: List of column headers from Excel
            
        Returns:
            Dictionary mapping column index to MappingResult
        """
        results = {}
        
        # First, detect obvious non-parameter columns (date, id, notes, comments)
        for col_idx, header in enumerate(headers):
            non_param, reason = self._is_non_parameter(header)
            if non_param:
                results[col_idx] = MappingResult(
                    header=header,
                    param_name=None,
                    asset_name=None,
                    confidence="low",
                    reason=reason
                )
                continue

            if not self.llm_available:
                # Fallback: use fuzzy matching
                results[col_idx] = self._fuzzy_match_header(header)
            else:
                # Use LLM for intelligent mapping
                results[col_idx] = self._map_header_with_llm(header)
        
        return results
    
    def _fuzzy_match_header(self, header: str) -> MappingResult:
        """Fallback fuzzy matching when LLM is not available"""
        from app.registries import get_parameter_by_name
        
        header_lower = header.lower().strip()
        
        best_match = None
        best_match_score = 0
        best_asset = None
        asset_inferred = False  # Track if asset was inferred indirectly
        
        # Try to match parameter
        for param in self.param_registry:
            for alias in param.get("aliases", []):
                alias_lower = alias.lower()
                score = self._string_similarity(header_lower, alias_lower)
                if score > best_match_score:
                    best_match_score = score
                    best_match = param["name"]  # Use 'name' field for ID
        
        # Try to match asset (direct match first)
        for asset in self.asset_registry:
            for alias in asset.get("aliases", []):
                if alias.lower() in header_lower:
                    best_asset = asset["name"]
                    break
        
        # Asset inference: map boiler/turbine patterns to asset IDs
        if not best_asset:
            best_asset, asset_inferred = self._infer_asset_from_text(header_lower)
        
        # Determine confidence (adjusted for asset inference)
        if best_match_score > 0.85:
            confidence = "high"
        elif best_match_score > 0.6:
            confidence = "medium"
        else:
            confidence = "low"
        
        # Reduce confidence if asset was inferred indirectly
        if asset_inferred and best_asset:
            if confidence == "high":
                confidence = "medium"
        
        reason = f"Fuzzy matched with score {best_match_score:.2f}"
        if asset_inferred:
            reason += f"; asset inferred from boiler/turbine mapping"
        
        return MappingResult(
            header=header,
            param_name=best_match,
            asset_name=best_asset,
            confidence=confidence,
            reason=reason
        )
    
    def _map_header_with_llm(self, header: str) -> MappingResult:
        """
        Use Google Gemini (new SDK) to map header intelligently.
        Falls back to fuzzy matching if API call fails.
        """
        try:
            from google import genai
            # Initialize client
            client = genai.Client(api_key=self.api_key)

            system_prompt = self.create_system_prompt()

            prompt = f"""
            {system_prompt}

            Map this Excel column header: '{header}'

            Return JSON only:
            {{
                \"param_name\": \"...\",
                \"asset_name\": \"...\",
                \"confidence\": \"high/medium/low\",
                \"reason\": \"...\"
            }}
            """

            response = client.models.generate_content(
                model="gemini-2.0-flash",
                contents=prompt
            )

            response_text = response.text.strip()

            mapping_json = json.loads(response_text)

            return MappingResult(
                header=header,
                param_name=mapping_json.get("param_name"),
                asset_name=mapping_json.get("asset_name"),
                confidence=mapping_json.get("confidence", "medium"),
                reason=mapping_json.get("reason", "LLM mapping")
            )

        except json.JSONDecodeError as e:
            print(f"⚠️ Failed to parse LLM response as JSON: {e}")
            return self._fuzzy_match_header(header)

        except Exception as e:
            print(f"⚠️ LLM API error: {e}. Falling back to fuzzy matching.")
            return self._fuzzy_match_header(header)

    def _is_non_parameter(self, header: str) -> (bool, Optional[str]):
        """Detect headers that are not parameters (date, id, notes, comments).

        Returns (True, reason) when detected, otherwise (False, None).
        """
        if not header:
            return True, "Empty header"

        h = header.lower()
        non_terms = ["date", "time", "timestamp", "day", "month", "year", "id", "comment", "notes", "remarks", "description"]
        for term in non_terms:
            if term in h:
                return True, f"Non-parameter column detected: '{term}'"

        # If header looks like a sequence number or generic column name
        if h.startswith("column_") or h.strip().lower().startswith("unnamed"):
            return True, "Unnamed or generic column"

        return False, None
    
    @staticmethod
    def _string_similarity(s1: str, s2: str) -> float:
        """Calculate simple string similarity using character overlap"""
        if s1 == s2:
            return 1.0
        if not s1 or not s2:
            return 0.0
        
        # Check if one is substring of other
        if s1 in s2 or s2 in s1:
            return 0.9
        
        # Levenshtein-like distance
        longer = s1 if len(s1) > len(s2) else s2
        shorter = s2 if len(s1) > len(s2) else s1
        
        match_count = sum(1 for c in shorter if c in longer)
        return match_count / len(longer)
    
    def _infer_asset_from_text(self, text: str) -> tuple:
        """Infer asset ID from pattern matching (e.g., 'Boiler 1' -> AFBC-1).
        
        Returns (asset_name, was_inferred) tuple.
        """
        # Boiler patterns -> AFBC asset mapping
        boiler_patterns = {
            "boiler 1": "AFBC-1",
            "boiler1": "AFBC-1",
            "boiler-1": "AFBC-1",
            "afbc-1": "AFBC-1",
            "boiler 2": "AFBC-2",
            "boiler2": "AFBC-2",
            "boiler-2": "AFBC-2",
            "afbc-2": "AFBC-2",
        }
        
        # Turbine patterns -> TG asset mapping
        turbine_patterns = {
            "tg-1": "TG-1",
            "tg1": "TG-1",
            "turbine 1": "TG-1",
            "turbine1": "TG-1",
            "turbine-1": "TG-1",
            "generator 1": "TG-1",
            "tg-2": "TG-2",
            "tg2": "TG-2",
            "turbine 2": "TG-2",
            "turbine2": "TG-2",
            "turbine-2": "TG-2",
            "generator 2": "TG-2",
        }
        
        all_patterns = {**boiler_patterns, **turbine_patterns}
        for pattern, asset_id in all_patterns.items():
            if pattern in text:
                return asset_id, True  # was inferred
        
        return None, False
    
    def detect_header_row(self, excel_data: List[List[any]], max_rows_to_check: int = 5) -> Optional[int]:
        """
        Detect the header row in Excel data by looking for non-empty strings.
        
        Args:
            excel_data: Raw Excel rows
            max_rows_to_check: Number of rows to inspect
            
        Returns:
            Header row index (0-based) or None
        """
        for row_idx in range(min(max_rows_to_check, len(excel_data))):
            row = excel_data[row_idx]
            # Check if this row has meaningful headers (mostly strings, not empty)
            string_count = sum(1 for cell in row if cell and isinstance(cell, str) and len(str(cell).strip()) > 2)
            if string_count >= len(row) * 0.6:  # At least 60% strings
                return row_idx
        return 0  # Default to first row
    
    def generate_data_validation_report(self, mapped_data: List[Dict]) -> Dict:
        """Generate validation report for mapped data"""
        report = {
            "total_cells_parsed": len(mapped_data),
            "cells_by_confidence": {"high": 0, "medium": 0, "low": 0},
            "unmapped_parameters": 0,
            "validation_warnings": []
        }
        
        for item in mapped_data:
            if "confidence" in item:
                report["cells_by_confidence"][item["confidence"]] += 1
            if "param_name" not in item or item["param_name"] is None:
                report["unmapped_parameters"] += 1
        
        return report
