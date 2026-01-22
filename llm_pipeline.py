"""
HVAC Data Extraction Pipeline - LLM Version
Uses Gemini Vision API for accurate HVAC data extraction
Generates NEW Excel workbook from scratch (no template needed)
"""
import os
import sys
import json
import argparse
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional

# Load environment variables from .env file
from dotenv import load_dotenv
load_dotenv()

# Add parent to path for imports
sys.path.insert(0, str(Path(__file__).parent))

from extractors.llm_extractor import GeminiHVACExtractor
from extractors.excel_generator import HVACExcelGenerator


class LLMHVACPipeline:
    """
    Complete HVAC extraction pipeline using LLM Vision
    
    Workflow:
    1. Convert PDF pages to images
    2. Send to Gemini for structured extraction
    3. Generate NEW Excel workbook with extracted data
    """
    
    def __init__(
        self,
        pdf_path: str,
        output_path: Optional[str] = None,
        job_number: str = "1168",
        project_name: str = "HVAC Project",
        api_key: Optional[str] = None,
        full_extraction: bool = True  # Extract from ALL pages, not just schedules
    ):
        self.pdf_path = pdf_path
        self.output_path = output_path or str(Path(pdf_path).with_suffix('.xlsx'))
        self.job_number = job_number
        self.project_name = project_name
        self.api_key = api_key or os.environ.get("GEMINI_API_KEY")
        self.full_extraction = full_extraction
        
        self.extracted_data: Dict[str, List] = {}
        
    def extract(self) -> Dict[str, List]:
        """Extract HVAC data from PDF using LLM"""
        print("\n" + "=" * 60)
        print("STEP 1: LLM EXTRACTION")
        print("=" * 60)
        
        if not self.api_key:
            raise ValueError(
                "GEMINI_API_KEY not set. Please set it:\n"
                "  Windows: set GEMINI_API_KEY=your-key\n"
                "  Linux/Mac: export GEMINI_API_KEY=your-key"
            )
        
        extractor = GeminiHVACExtractor(api_key=self.api_key)
        
        if self.full_extraction:
            print("Mode: FULL EXTRACTION (all pages)")
            self.extracted_data = extractor.extract_from_pdf(self.pdf_path)
        else:
            print("Mode: SCHEDULE-ONLY (faster)")
            self.extracted_data = extractor.extract_schedules_only(self.pdf_path)
        
        # Summary
        print("\nExtraction Summary:")
        for key, items in self.extracted_data.items():
            if items:
                print(f"  {key}: {len(items)} items")
        
        return self.extracted_data
    
    def generate_excel(self) -> str:
        """Generate NEW Excel workbook from extracted data"""
        print("\n" + "=" * 60)
        print("STEP 2: EXCEL GENERATION")
        print("=" * 60)
        
        if not self.extracted_data:
            print("No extracted data. Run extract() first.")
            return None
        
        # Create output directory if needed
        Path(self.output_path).parent.mkdir(parents=True, exist_ok=True)
        
        generator = HVACExcelGenerator(
            job_number=self.job_number,
            project_name=self.project_name
        )
        stats = generator.generate_from_data(self.extracted_data)
        
        output = generator.save(self.output_path)
        
        print("\nGeneration Summary:")
        print(f"  VAV sheets: {stats['vavs']}")
        print(f"  Fan sheets: {stats['fans']}")
        print(f"  CRAC sheets: {stats['cracs']}")
        print(f"  Heater sheets: {stats['heaters']}")
        print(f"  Total sheets: {sum(stats.values()) + 1}")  # +1 for summary
        
        return output
    
    def save_json(self, path: Optional[str] = None) -> str:
        """Save extracted data as JSON for debugging"""
        json_path = path or self.pdf_path.replace(".pdf", "_extracted.json")
        
        with open(json_path, "w") as f:
            json.dump(self.extracted_data, f, indent=2)
        
        print(f"  JSON saved to: {json_path}")
        return json_path
    
    def run(self, save_intermediate: bool = True) -> str:
        """
        Run complete pipeline
        
        Args:
            save_intermediate: Save extracted JSON for debugging
            
        Returns:
            Path to generated Excel file
        """
        print("=" * 60)
        print("LLM HVAC EXTRACTION PIPELINE")
        print("=" * 60)
        print(f"PDF Input:    {self.pdf_path}")
        print(f"Output:       {self.output_path}")
        print(f"Job Number:   {self.job_number}")
        print(f"Project:      {self.project_name}")
        print(f"LLM Model:    Gemini 2.0 Flash")
        print(f"Extraction:   {'Full (all pages)' if self.full_extraction else 'Schedules only'}")
        
        # Step 1: Extract
        self.extract()
        
        # Optionally save JSON
        if save_intermediate:
            self.save_json()
        
        # Step 2: Generate Excel
        output = self.generate_excel()
        
        # Done
        print("\n" + "=" * 60)
        print("✓ PIPELINE COMPLETE!")
        print("=" * 60)
        print(f"Output file: {output}")
        
        return output


def main():
    """CLI interface for the pipeline"""
    parser = argparse.ArgumentParser(
        description="Extract HVAC data from PDF and generate Excel report"
    )
    parser.add_argument(
        "--pdf", "-p",
        default=r"D:\SW\new project\Boeing R&D Drawings.pdf",
        help="Path to PDF file"
    )
    parser.add_argument(
        "--output", "-o",
        default=r"D:\SW\new project\output\hvac_report.xlsx",
        help="Output Excel path"
    )
    parser.add_argument(
        "--job", "-j",
        default="1168",
        help="Job number"
    )
    parser.add_argument(
        "--project", "-n",
        default="Boeing Arlington R&D",
        help="Project name"
    )
    parser.add_argument(
        "--api-key", "-k",
        default=None,
        help="Gemini API key (or set GEMINI_API_KEY env var)"
    )
    parser.add_argument(
        "--schedules-only", "-s",
        action="store_true",
        help="Only extract from schedule pages (faster but may miss data)"
    )
    
    args = parser.parse_args()
    
    pipeline = LLMHVACPipeline(
        pdf_path=args.pdf,
        output_path=args.output,
        job_number=args.job,
        project_name=args.project,
        api_key=args.api_key,
        full_extraction=not args.schedules_only
    )
    
    try:
        pipeline.run()
    except ValueError as e:
        print(f"\n❌ Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
