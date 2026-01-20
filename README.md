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

```bash
python hvac_pipeline.py
```

## Requirements
- Python 3.8+
- PyMuPDF (fitz)
- openpyxl

## Project Structure
```
├── hvac_pipeline.py          # Main pipeline
├── extractors/
│   └── improved_extractor.py # Core extraction logic
├── debug_scripts/            # Test/debug utilities
└── output/                   # Generated files
```

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License
MIT
