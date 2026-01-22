# HVAC Extractor

Automated HVAC data extraction from PDF drawings.

## Features
- Extracts VAV box data (CFM values, inlet sizes) from HVAC PDF drawings
- Generates structured Excel output
- Template-independent extraction using contextual analysis
- Supports multiple PDF formats and layouts

## Installation

1. Clone the repository:
```bash
git clone https://github.com/syedowais57/Hvac_extractor.git
cd Hvac_extractor
```

2. Install dependencies:
```bash
pip install pymupdf openpyxl
```

## Usage

### Quick Test
```bash
# Check setup
python scripts/check_setup.py

# Run local test
python scripts/test_local.py

# Start API server
python scripts/start_api.py
```

### Web API
```bash
# Start server
python scripts/start_api.py

# Then open: http://localhost:8000
```

See `docs/QUICK_START.md` for detailed instructions.

## Requirements
- Python 3.8+
- PyMuPDF (fitz)
- openpyxl

## Project Structure
```
├── api/
│   └── app.py                # FastAPI web service
├── extractors/
│   ├── llm_extractor.py     # LLM-based extraction (Gemini)
│   ├── excel_generator.py   # Excel report generation
│   └── excel_populator.py   # Template population
├── scripts/                  # Test and utility scripts
│   ├── test_local.py         # Direct pipeline test
│   ├── start_api.py          # Start API server locally
│   └── ...
├── docs/                     # Documentation
│   ├── QUICK_START.md        # Quick start guide
│   ├── DEPLOY_NOW.md         # Deployment guide
│   └── ...
├── llm_pipeline.py           # Main LLM pipeline
├── hvac_pipeline.py          # Legacy pipeline
└── output/                   # Generated files
```

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License
MIT
