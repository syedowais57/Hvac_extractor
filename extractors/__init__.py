# HVAC Extractors Package
"""
Extractors for HVAC equipment data from PDF drawings.

Available extractors:
- ImprovedVAVExtractor: Regex-based VAV extraction (fallback)
- GeminiHVACExtractor: LLM-based extraction using Gemini Vision
- HVACExcelPopulator: Populates Excel templates with extracted data
"""

from extractors.improved_extractor import ImprovedVAVExtractor, extract_vavs, VAVData

# LLM-based extractor (requires GEMINI_API_KEY)
try:
    from extractors.llm_extractor import GeminiHVACExtractor, extract_hvac_with_llm
except ImportError:
    GeminiHVACExtractor = None
    extract_hvac_with_llm = None

# Excel populator
from extractors.excel_populator import HVACExcelPopulator, populate_excel_from_json

__all__ = [
    'ImprovedVAVExtractor',
    'extract_vavs',
    'VAVData',
    'GeminiHVACExtractor',
    'extract_hvac_with_llm',
    'HVACExcelPopulator',
    'populate_excel_from_json',
]
