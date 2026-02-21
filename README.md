


# Intelligent Excel Parser 📊



## The Problem

Factories upload operational data as Excel files, but they're **never consistent**:
- Column names vary: "Coal Consumption" vs "COAL CONSMPTN" vs "Coal Used (MT)"
- Asset references embedded in headers: "Coal Consumption AFBC-1"
- Values in different formats: "1,234.56" vs "45%" vs "YES"
- Metadata rows scattered randomly
- Multiple sheets with inconsistent layouts

This service **automatically parses, maps, and validates** that messy data into reliable JSON.

This project uses Google's Gemini LLM (via Gemini API) for intelligent header understanding and semantic mapping.

---

## Features ✨

### Core Capabilities
- ✅ **Intelligent Header Mapping** - Maps messy headers to canonical parameter names using fuzzy matching
- ✅ **Asset Detection** - Automatically detects when headers reference specific assets (AFBC-1, TG-2, etc.)
- ✅ **Header Row Detection** - Automatically skips title rows, empty rows, and finds the actual header
- ✅ **Smart Value Parsing** - Converts "1,234.56" → 1234.56, "45%" → 0.45, "YES" → 1.0, "N/A" → null
- ✅ **Confidence Scoring** - Each parsed cell gets high/medium/low confidence based on match quality
- ✅ **Unmapped Columns** - Flags columns that don't match any known parameter
  

---

## Project Structure

```
latspace/
├── app/
│   ├── main.py              # FastAPI app & endpoints
│   ├── parser.py            # Core Excel parsing logic
│   ├── llm_agent.py         # Header mapping 
│   ├── models.py            # Pydantic response models
│   ├── registries.py        # Parameter & asset definitions
│   ├── value_parser.py      # Value conversion & validation
│   └── __init__.py
├── data/                    # Test Excel files
│   ├── clean_data.xlsx      # Clean baseline (should parse perfectly)
│   ├── messy_data.xlsx      # Messy headers, mixed formats
│   ├── multi_asset_data.xlsx # Multiple assets per parameter
│   └── ...                  # More test cases
├── scripts/
│   └── post_test_parse.py   # Quick API test script
├── Dockerfile              # Container setup
├── docker-compose.yml      # Easy local run
├── requirements.txt        # Dependencies
├── create_test_data.py     # Generate test Excel files
└── README.md              # This file
```

---


## Quick Start 🚀

### Option 1: Docker (Recommended)
```bash
# Start the service
docker-compose up --build
create .env file
GOOGLE_API_KEY=your_google_api_key_here
# Access API docs
open http://localhost:8000/docs
```

### Option 2: Local Python
```bash
# Install dependencies
pip install -r requirements.txt

# Create .env file
echo "GOOGLE_API_KEY=your_key_here" > .env

# Run the app
python -m uvicorn app.main:app --reload

# Access API
open http://localhost:8000/docs
```

---

## API Usage 🔌

### Upload & Parse Excel
```bash
curl -X POST "http://localhost:8000/parse" \
  -F "file=@data/clean_data.xlsx"
```

### Response Example
```json
{
  "status": "success",
  "header_row": 0,
  "parsed_data": [
    {
      "row": 1,
      "col": 0,
      "param_name": "coal_consumption",
      "asset_name": "AFBC-1",
      "raw_value": "1,234.56",
      "parsed_value": 1234.56,
      "confidence": "high"
    }
  ],
  "unmapped_columns": [],
  "warnings": [],
  "detected_assets": ["AFBC-1"],
  "parameters": {"coal_consumption": ["AFBC-1"]}
}
```

### Available Endpoints
- `GET /health` - Service health check
- `POST /parse` - Upload and parse Excel file (main endpoint)
- `GET /registries` - View all parameters & assets

---

## Parameters & Assets 📋

### Supported Parameters (20+)
- **COGEN BOILER**: coal_consumption, steam_generation, efficiency
- **POWER PLANT**: power_generation, power_export, heat_rate
- **UTILITIES**: water_consumption, auxiliary_power
- **EMISSIONS**: co2_emissions, so2_emissions, nox_emissions
- **PRODUCTION**: production_output, fly_ash_generated
- **And more...**

### Supported Assets
- `AFBC-1`, `AFBC-2` (Boilers)
- `TG-1`, `TG-2` (Turbine Generators)
- `KILN-1` (Rotary Kiln)
- `VSF` (Viscose Staple Fiber)

---

## Testing 🧪

### Generate Test Data
```bash
python create_test_data.py
```

Creates 10 test files covering edge cases:
- Clean headers, messy headers, mixed formats
- Multiple assets per parameter
- Special characters, validation errors

### Test the API
```bash
python scripts/post_test_parse.py
```

---

## How It Works 🔧


### Parsing Flow
1. **Load Excel** → Extract sheets, detect header row
2. **Map Headers** → Match column names to known parameters (fuzzy matching)
3. **Detect Assets** → Extract asset names from headers ("Boiler 1" → AFBC-1)
4. **Parse Values** → Convert raw values to numbers with confidence scoring
5. **Validate** → Flag suspicious values and unmapped columns
6. **Return JSON** → Structured response with all metadata

### Header Mapping Strategy
- **Exact Match** (confidence: high) - Perfect parameter name match
- **Fuzzy Match** (confidence: medium) - Close match using aliases
- **Asset Inference** (confidence: medium→low) - Pattern matching for assets
- **Unmapped** (confidence: low) - No match found, flagged for review

<img width="673" height="674" alt="Screenshot 2026-02-21 124422" src="https://github.com/user-attachments/assets/94aa8670-5571-4db9-b2ab-d2454bee23e1" />
---
https://github.com/user-attachments/assets/6bd52d07-69d1-4ae9-947f-14e19673fb42

---

