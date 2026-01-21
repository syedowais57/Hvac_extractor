"""
LLM-Based HVAC Extractor using Google Gemini Vision
Extracts HVAC equipment data from PDF drawings
"""
import os
import json
import fitz  # PyMuPDF
import google.generativeai as genai
from dataclasses import dataclass, asdict
from typing import List, Dict, Optional, Any
from pathlib import Path
import base64


@dataclass
class ExtractedEquipment:
    """Base class for extracted equipment"""
    equipment_type: str
    tag: str
    location: str = ""
    
    def to_dict(self) -> Dict:
        return asdict(self)


@dataclass
class VAVData:
    """VAV/Air Terminal Unit data"""
    tag: str
    location: str = ""
    cfm_max: int = 0
    cfm_min: int = 0
    inlet_size: str = ""
    has_reheat: bool = False
    reheat_kw: float = 0.0
    
    def to_dict(self) -> Dict:
        return asdict(self)


@dataclass
class FanData:
    """Exhaust Fan data"""
    tag: str
    location: str = ""
    fan_type: str = ""
    drive: str = ""
    cfm: int = 0
    esp: float = 0.0  # External Static Pressure
    motor_power: str = ""
    rpm: int = 0
    voltage: str = ""
    
    def to_dict(self) -> Dict:
        return asdict(self)


@dataclass
class CRACData:
    """CRAC Unit data"""
    tag: str
    location: str = ""
    cfm: int = 0
    cooling_capacity: str = ""
    
    def to_dict(self) -> Dict:
        return asdict(self)


@dataclass 
class ElectricHeaterData:
    """Electric Duct Heater data"""
    tag: str
    location: str = ""
    cfm: int = 0
    voltage: int = 0
    kw: float = 0.0
    associated_vav: str = ""
    
    def to_dict(self) -> Dict:
        return asdict(self)


class GeminiHVACExtractor:
    """
    Extracts HVAC equipment data from PDF drawings using Gemini Vision API
    """
    
    SCHEDULE_PROMPT = """Analyze this HVAC drawing page VERY CAREFULLY and extract ALL equipment data with COMPLETE details.

Look for:
1. VAV SCHEDULE tables - Contains VAVB5-XX entries with complete specifications
2. FAN SCHEDULE tables - Contains EF-1, EF-2, etc.
3. ELECTRIC DUCT HEATER SCHEDULE - Contains heater data
4. CRAC UNIT data
5. AIR DEVICE SCHEDULE - Supply diffusers, grilles

VAV TAGS follow pattern: VAVB5-01, VAVB5-02, ... VAVB5-99

For each VAV, extract ALL these fields if visible:
- tag: VAV identifier (e.g., VAVB5-01)
- location: Room number (e.g., "Software 443")
- area_served: Area description (e.g., "Large Conference 402")
- inlet_size: Primary air inlet size in inches (e.g., "10")
- total_cfm: Total fan CFM (design value)
- cfm_min: Minimum CFM
- cfm_max: Maximum CFM
- manufacturer: Manufacturer name
- model: Model number
- motor_hp: Motor horsepower
- motor_voltage: Motor voltage
- motor_phase: Motor phase
- motor_amperage: Motor amperage
- has_reheat: true/false if unit has electric reheat
- reheat_kw: Reheat kilowatts

Return a JSON object with this structure:
{
    "fans": [
        {"tag": "EF-1", "location": "407 - RR", "fan_type": "CEILING", "drive": "DIRECT", "cfm": 100, "esp": 0.25, "motor_power": "21 W", "rpm": 1100, "voltage": "120/1/60"}
    ],
    "vavs": [
        {
            "tag": "VAVB5-01",
            "location": "501",
            "area_served": "Office Area",
            "inlet_size": "10",
            "total_cfm": 750,
            "cfm_min": 150,
            "cfm_max": 750,
            "manufacturer": "",
            "model": "",
            "motor_hp": "",
            "motor_voltage": "",
            "has_reheat": true,
            "reheat_kw": 4.0
        }
    ],
    "cracs": [
        {"tag": "CRAC-1", "location": "Server Room", "cfm": 1000, "cooling_capacity": "5 tons"}
    ],
    "heaters": [
        {"tag": "VAVB5-01-H", "location": "501", "cfm": 700, "voltage": 277, "kw": 4.0, "associated_vav": "VAVB5-01"}
    ],
    "air_devices": [
        {"tag": "SD-1", "type": "Supply Diffuser", "cfm": 200, "size": "12x12"}
    ]
}

- Extract ALL rows from ALL tables - do not skip any
- For VAVs, inlet size is typically 6, 8, 10, or 12 inches
- If a field is not visible, use null
- If no equipment is found, return an object with empty lists.
- Return ONLY a valid JSON object. Do not include any conversational text, explanations, or markdown formatting. 
- Ensure every tag follows the expected pattern (e.g., VAVB5-XX, EF-X)."""

    FLOOR_PLAN_PROMPT = """Analyze this HVAC floor plan drawing VERY CAREFULLY.

Look for ALL equipment tags visible anywhere on the drawing:
- VAV boxes labeled VAVB5-XX (e.g., VAVB5-01, VAVB5-02, VAVB5-52, VAVB5-78)
- Exhaust fans labeled EF-X
- CRAC units labeled CRAC-X

Return a JSON object:
{
    "vavs": [
        {"tag": "VAVB5-01", "location": "Room 501", "cfm_max": null, "cfm_min": null, "inlet_size": null, "has_reheat": null, "reheat_kw": null}
    ],
    "fans": [],
    "cracs": [],
    "heaters": [],
    "air_devices": []
}

- Look for tags in circles, rectangles, or next to ductwork
- Include tags that may be partially visible or small
- Tags range from VAVB5-01 to VAVB5-99
- If no tags are found, return the JSON object with empty lists.
- IMPORTANT: Return ONLY valid JSON. Absolutely no conversational text or explaining why you couldn't find anything.
- Return ONLY the JSON object, do not use markdown code blocks."""

    def __init__(self, api_key: Optional[str] = None):
        """Initialize with Gemini API key"""
        self.api_key = api_key or os.environ.get("GEMINI_API_KEY")
        if not self.api_key:
            raise ValueError("GEMINI_API_KEY not provided. Set it as environment variable or pass to constructor.")
        
        genai.configure(api_key=self.api_key)
        self.model = genai.GenerativeModel("gemini-2.0-flash")
        
    def _pdf_page_to_image(self, pdf_path: str, page_num: int, dpi: int = 120) -> bytes:
        """Convert PDF page to PNG image bytes"""
        doc = fitz.open(pdf_path)
        page = doc[page_num]
        
        # Higher DPI for better text recognition
        zoom = dpi / 72
        matrix = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=matrix)
        
        img_bytes = pix.tobytes("png")
        doc.close()
        return img_bytes
    
    def _is_schedule_page(self, pdf_path: str, page_num: int) -> bool:
        """Check if page contains schedule tables"""
        doc = fitz.open(pdf_path)
        text = doc[page_num].get_text().upper()
        doc.close()
        
        schedule_indicators = ["SCHEDULE", "DESIGNATION", "CFM", "AIRFLOW"]
        return any(ind in text for ind in schedule_indicators)
    
    def _extract_with_gemini(self, image_bytes: bytes, prompt: str) -> Dict:
        """Send image to Gemini and extract structured data"""
        import PIL.Image
        import io
        
        # Convert bytes to PIL Image
        image = PIL.Image.open(io.BytesIO(image_bytes))
        
        # Send to Gemini
        response = self.model.generate_content([prompt, image])
        
        # Parse JSON response
        response_text = response.text.strip()
        
        # Enhanced cleanup: Find the first '{' and last '}'
        try:
            start_idx = response_text.find('{')
            end_idx = response_text.rfind('}')
            if start_idx != -1 and end_idx != -1:
                response_text = response_text[start_idx:end_idx + 1]
            
            return json.loads(response_text)
        except (json.JSONDecodeError, ValueError) as e:
            print(f"JSON parse error on page: {e}")
            # Fallback for "I am unable to..." conversational responses
            return {"fans": [], "vavs": [], "cracs": [], "heaters": [], "air_devices": []}
    
    def extract_from_pdf(self, pdf_path: str) -> Dict[str, List]:
        """
        Extract all HVAC equipment from PDF (ALL pages)
        
        Returns:
            Dict with keys: fans, vavs, cracs, heaters, air_devices
        """
        doc = fitz.open(pdf_path)
        num_pages = len(doc)
        doc.close()
        
        all_data = {
            "fans": [],
            "vavs": [],
            "cracs": [],
            "heaters": [],
            "air_devices": []
        }
        
        print(f"Processing ALL {num_pages} pages...")
        
        for page_num in range(num_pages):
            print(f"  Page {page_num + 1}/{num_pages}...", end=" ")
            
            image_bytes = self._pdf_page_to_image(pdf_path, page_num)
            
            if self._is_schedule_page(pdf_path, page_num):
                print("(Schedule page)")
                data = self._extract_with_gemini(image_bytes, self.SCHEDULE_PROMPT)
            else:
                print("(Floor plan)")
                data = self._extract_with_gemini(image_bytes, self.FLOOR_PLAN_PROMPT)
            
            # Merge extracted data
            for key in all_data:
                if key in data and data[key]:
                    all_data[key].extend(data[key])
        
        # Post-process: deduplicate and merge
        all_data = self._deduplicate_and_merge(all_data)
        
        # Generate heater entries from VAVs with reheat
        all_data = self._generate_heaters_from_vavs(all_data)
        
        return all_data
    
    def _deduplicate_and_merge(self, data: Dict[str, List]) -> Dict[str, List]:
        """Deduplicate entries by tag, keeping the one with most data"""
        for key in data:
            if not data[key]:
                continue
                
            # Group by tag
            by_tag = {}
            for item in data[key]:
                tag = item.get("tag")
                if not tag:
                    continue
                    
                if tag not in by_tag:
                    by_tag[tag] = item
                else:
                    # Merge: keep values from item with more non-null fields
                    existing = by_tag[tag]
                    for field, value in item.items():
                        if value is not None and existing.get(field) is None:
                            existing[field] = value
            
            data[key] = list(by_tag.values())
        
        return data
    
    def _generate_heaters_from_vavs(self, data: Dict[str, List]) -> Dict[str, List]:
        """Generate heater entries for VAVs with reheat"""
        existing_heater_tags = {h.get("tag") for h in data.get("heaters", [])}
        
        for vav in data.get("vavs", []):
            if vav.get("has_reheat") and vav.get("tag"):
                heater_tag = f"{vav['tag']}-H"
                
                if heater_tag not in existing_heater_tags:
                    heater = {
                        "tag": heater_tag,
                        "location": vav.get("location", ""),
                        "cfm": vav.get("cfm_max", 0),
                        "voltage": 277,  # Standard voltage for electric heaters
                        "kw": vav.get("reheat_kw", 0),
                        "associated_vav": vav["tag"]
                    }
                    data["heaters"].append(heater)
                    existing_heater_tags.add(heater_tag)
        
        return data
    
    def extract_schedules_only(self, pdf_path: str) -> Dict[str, List]:
        """Extract only from schedule pages (faster, cheaper)"""
        doc = fitz.open(pdf_path)
        num_pages = len(doc)
        doc.close()
        
        all_data = {
            "fans": [],
            "vavs": [],
            "cracs": [],
            "heaters": [],
            "air_devices": []
        }
        
        for page_num in range(num_pages):
            if self._is_schedule_page(pdf_path, page_num):
                print(f"Processing schedule page {page_num + 1}...")
                image_bytes = self._pdf_page_to_image(pdf_path, page_num)
                data = self._extract_with_gemini(image_bytes, self.SCHEDULE_PROMPT)
                
                for key in all_data:
                    if key in data and data[key]:
                        all_data[key].extend(data[key])
        
        # Post-process: deduplicate and merge
        all_data = self._deduplicate_and_merge(all_data)
        
        # Generate heater entries from VAVs with reheat
        all_data = self._generate_heaters_from_vavs(all_data)
        
        return all_data


def extract_hvac_with_llm(pdf_path: str, api_key: Optional[str] = None) -> Dict[str, List]:
    """Convenience function to extract HVAC data using LLM"""
    extractor = GeminiHVACExtractor(api_key=api_key)
    return extractor.extract_schedules_only(pdf_path)


if __name__ == "__main__":
    import sys
    
    # Test extraction
    PDF_PATH = r"D:\SW\new project\Boeing R&D Drawings.pdf"
    
    # Check for API key
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        print("Please set GEMINI_API_KEY environment variable")
        print("Example: set GEMINI_API_KEY=your-api-key-here")
        sys.exit(1)
    
    print("=" * 60)
    print("LLM-Based HVAC Extraction")
    print("=" * 60)
    
    extractor = GeminiHVACExtractor(api_key=api_key)
    data = extractor.extract_schedules_only(PDF_PATH)
    
    print("\n" + "=" * 60)
    print("EXTRACTION RESULTS")
    print("=" * 60)
    
    for equipment_type, items in data.items():
        if items:
            print(f"\n{equipment_type.upper()} ({len(items)} items):")
            for item in items[:5]:  # Show first 5
                print(f"  {item}")
            if len(items) > 5:
                print(f"  ... and {len(items) - 5} more")
    
    # Save to JSON
    output_file = "extracted_hvac_data.json"
    with open(output_file, "w") as f:
        json.dump(data, f, indent=2)
    print(f"\nFull data saved to {output_file}")
