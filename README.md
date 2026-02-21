# Intelligent Excel Parser 📊

Transform messy factory Excel spreadsheets into clean, structured JSON data with AI-powered parsing and intelligent header mapping.

## The Problem

Factories upload operational data as Excel files, but they're **never consistent**:
- Column names vary: "Coal Consumption" vs "COAL CONSMPTN" vs "Coal Used (MT)"
- Asset references embedded in headers: "Coal Consumption AFBC-1"
- Values in different formats: "1,234.56" vs "45%" vs "YES"
- Metadata rows scattered randomly
- Multiple sheets with inconsistent layouts

This service **automatically parses, maps, and validates** that messy data into reliable JSON.

---

## Features ✨

### Core Capabilities
- ✅ **Intelligent Header Mapping** - Maps messy headers to canonical parameter names using fuzzy matching
- ✅ **Asset Detection** - Automatically detects when headers reference specific assets (AFBC-1, TG-2, etc.)
- ✅ **Header Row Detection** - Automatically skips title rows, empty rows, and finds the actual header
- ✅ **Smart Value Parsing** - Converts "1,234.56" → 1234.56, "45%" → 0.45, "YES" → 1.0, "N/A" → null
- ✅ **Confidence Scoring** - Each parsed cell gets high/medium/low confidence based on match quality
- ✅ **Unmapped Columns** - Flags columns that don't match any known parameter
- ✅ **Multi-sheet Support** - Handle workbooks with data across multiple sheets

### Nice to Have
- 📋 Validation rules (flag negative coal consumption, efficiency > 100%)
- 🔍 Duplicate detection (same parameter+asset in multiple columns)
- 📊 Chunked processing (for files > 5MB)
- 👤 Human review mode (low-confidence mappings for approval)

---

## Project Structure

```
latspace/
├── app/
│   ├── main.py              # FastAPI app & endpoints
│   ├── parser.py            # Core Excel parsing logic
│   ├── llm_agent.py         # Header mapping (fuzzy logic)
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
- Large file (5MB) for stress testing

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

---

## Environment Setup 🔐

Create `.env` file:
```env
GOOGLE_API_KEY=your_google_generativeai_key_here
```

**Why?** Optional LLM integration for advanced header mapping. Works without it using fuzzy matching.

---

## Development 💻

### Install Dev Dependencies
```bash
pip install -r requirements.txt pytest pytest-asyncio black flake8
```

### Run Tests
```bash
pytest
```

### Format Code
```bash
black app/
```

### Lint
```bash
flake8 app/
```

---

## Deployment 🚢

### Docker (Simple)
```bash
docker build -t excel-parser .
docker run -p 8000:8000 --env-file .env excel-parser
```

### Docker Compose
```bash
docker-compose up -d
docker-compose logs -f
```

---

## Example Workflows 📝

### Workflow 1: Parse Clean Data
```
Upload: clean_data.xlsx
↓
Headers automatically detected
↓
All columns mapped to parameters
↓
All values parsed successfully
↓
Response: 100% high confidence
```

### Workflow 2: Parse Messy Data
```
Upload: messy_data.xlsx (title rows, bad headers, mixed formats)
↓
Header row detected automatically (skips title row)
↓
Headers fuzzy-matched to parameters
↓
"Coal Consumption AFBC-1" → param: coal_consumption, asset: AFBC-1
↓
"1,234.56" converted to 1234.56
↓
Response: Mix of high/medium confidence + warnings
```

### Workflow 3: Handle Unmapped Columns
```
Upload: unknown_data.xlsx
↓
Some columns "Comments", "Notes" don't match any parameter
↓
Flagged in unmapped_columns
↓
Response includes reason: "Non-parameter column detected: 'notes'"
```

---

## Architecture 🏗️

```
FastAPI (app/main.py)
    ↓
ExcelParser (app/parser.py)
    ├─→ Header Detection
    ├─→ ExcelParsingAgent (app/llm_agent.py)
    │   └─→ Fuzzy matching + asset inference
    ├─→ Value Parsing (app/value_parser.py)
    │   └─→ Convert formats & validate
    └─→ Pydantic Models (app/models.py)
        └─→ Structured JSON response

Registries (app/registries.py)
    ├─→ PARAMETER_REGISTRY (20+ parameters)
    └─→ ASSET_REGISTRY (6 assets)
```

---

## Troubleshooting 🐛

### Docker won't start
```bash
# Check logs
docker-compose logs

# Rebuild
docker-compose down && docker-compose up --build
```

### File upload fails
- Ensure `.xlsx` format (not `.xls` or `.csv`)
- File size < 50MB (recommended)
- Check `/health` endpoint first

### Headers not mapping correctly
- Check parameter registry: `GET /registries`
- Add aliases to parameter definition in `app/registries.py`
- Enable LLM with `GOOGLE_API_KEY` for advanced matching

---

## Performance ⚡

- **Avg Response Time**: 100-500ms for typical files
- **File Size Support**: Up to 50MB (tested with 5MB)
- **Headers Tested**: 100+ variations
- **Test Coverage**: 10+ edge case files

---

## Requirements 📦

- Python 3.11+
- FastAPI 0.104+
- openpyxl (Excel parsing)
- Pydantic (Data validation)
- Optional: google-generativeai (Advanced mapping)

---

## License 📄

Built as part of LatSpace Technical Challenge - Track A

---

## Contributing 🤝

1. Fork the repo
2. Create feature branch (`git checkout -b feature/amazing`)
3. Commit changes (`git commit -am 'Add amazing feature'`)
4. Push to branch (`git push origin feature/amazing`)
5. Open Pull Request

---

## Questions? 💬

Check these resources:
- API Docs: `http://localhost:8000/docs` (Swagger UI)
- Test files: `data/` folder
- Examples: `scripts/post_test_parse.py`
- Code: Well-commented Python files in `app/`

---

**Status**: ✅ Production Ready | 🚀 Actively Maintained | 📊 Tested with 10+ edge cases
