# HVAC Extractor

Automated HVAC data extraction from PDF drawings.

## Features
- Extracts VAV box data (CFM values, inlet sizes) from HVAC PDF drawings
- Generates structured Excel output
- Template-independent extraction using contextual analysis

## Usage

```bash
python hvac_pipeline.py
```

## Requirements
- Python 3.8+
- PyMuPDF (fitz)
- openpyxl

```bash
pip install pymupdf openpyxl
```

## Project Structure
```
├── hvac_pipeline.py          # Main pipeline
├── extractors/
│   └── improved_extractor.py # Core extraction logic
├── debug_scripts/            # Test/debug utilities
└── output/                   # Generated files
```
