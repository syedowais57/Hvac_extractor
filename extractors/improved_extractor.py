"""
Improved HVAC Extractor - Template Independent
Uses contextual grouping and multiple extraction strategies
"""
import fitz
import re
from dataclasses import dataclass, field
from typing import List, Dict, Tuple, Optional
import math


@dataclass
class VAVData:
    tag: str
    total_cfm: int = 0
    min_cfm: int = 0
    max_cfm: int = 0
    inlet_size: str = ""
    location: str = ""
    area_served: str = ""
    page: int = 0
    x: float = 0
    y: float = 0


class ImprovedVAVExtractor:
    """
    Improved VAV extraction using multiple strategies:
    1. Look for VAV schedule tables first (most accurate)
    2. Use contextual grouping (text blocks near VAV tags)
    3. Fall back to proximity matching
    """
    
    VAV_PATTERN = re.compile(r"VAVB?\d*-\d+|VAV-\d+")
    CFM_PATTERN = re.compile(r"(\d{2,4})\s*(?:CFM|cfm)?")
    SIZE_PATTERN = re.compile(r'(\d+)"?\s*(?:Ø|ø|INCH|inch|IN)?')
    
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        
    def close(self):
        self.doc.close()
    
    def _get_text_blocks_with_positions(self, page_index: int) -> List[Dict]:
        """Get all text blocks with position info"""
        page = self.doc[page_index]
        blocks = []
        
        for block in page.get_text("dict")["blocks"]:
            if "lines" in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        text = span["text"].strip()
                        if text:
                            bbox = span["bbox"]
                            blocks.append({
                                "text": text,
                                "x": bbox[0],
                                "y": bbox[1],
                                "x1": bbox[2],
                                "y1": bbox[3],
                                "font_size": span["size"]
                            })
        return blocks
    
    def _find_vav_with_context(self, page_index: int) -> List[VAVData]:
        """
        Find VAV tags and analyze surrounding text for CFM values
        Uses a contextual window approach
        """
        blocks = self._get_text_blocks_with_positions(page_index)
        vavs = []
        cfm_blocks = []
        
        # First pass: identify VAV tags and CFM values
        for block in blocks:
            text = block["text"]
            
            # Find VAV tags
            if self.VAV_PATTERN.search(text):
                vav_match = self.VAV_PATTERN.search(text)
                vavs.append({
                    "tag": vav_match.group(),
                    "x": block["x"],
                    "y": block["y"],
                    "full_text": text
                })
            
            # Find CFM values
            try:
                if text.isdigit() and 50 <= int(text) <= 5000:
                    cfm_blocks.append({
                        "value": int(text),
                        "x": block["x"],
                        "y": block["y"]
                    })
                elif "CFM" in text.upper():
                    cfm_match = self.CFM_PATTERN.search(text)
                    if cfm_match:
                        cfm_blocks.append({
                            "value": int(cfm_match.group(1)),
                            "x": block["x"],
                            "y": block["y"]
                        })
            except (ValueError, UnicodeError):
                pass  # Skip problematic text
        
        # Second pass: associate CFM to VAVs using proximity
        results = []
        for vav in vavs:
            vav_data = VAVData(
                tag=vav["tag"],
                page=page_index,
                x=vav["x"],
                y=vav["y"]
            )
            
            if cfm_blocks:
                # Find nearest CFM within reasonable distance (300 points)
                distances = []
                for cfm in cfm_blocks:
                    dist = math.sqrt((vav["x"] - cfm["x"])**2 + (vav["y"] - cfm["y"])**2)
                    if dist < 300:  # Only consider CFMs within 300 points
                        distances.append((cfm["value"], dist))
                
                if distances:
                    nearest = min(distances, key=lambda x: x[1])
                    vav_data.total_cfm = nearest[0]
                    vav_data.max_cfm = nearest[0]
                    vav_data.min_cfm = int(nearest[0] * 0.2)  # Estimate
            
            # Estimate inlet size from CFM
            vav_data.inlet_size = self._estimate_inlet_size(vav_data.total_cfm)
            
            results.append(vav_data)
        
        return results
    
    def _find_schedule_data(self) -> Dict[str, VAVData]:
        """
        Look for VAV schedule tables in the PDF
        These typically have columns: VAV Tag, CFM, Size, etc.
        """
        schedule_data = {}
        
        for page_index in range(len(self.doc)):
            page = self.doc[page_index]
            text = page.get_text()
            
            # Look for schedule indicators
            if "SCHEDULE" in text.upper() or "VAV" in text.upper():
                # Try to extract tabular data
                lines = text.split("\n")
                for i, line in enumerate(lines):
                    vav_match = self.VAV_PATTERN.search(line)
                    if vav_match:
                        tag = vav_match.group()
                        # Look for CFM in same line or nearby lines
                        context = " ".join(lines[max(0,i-1):min(len(lines),i+2)])
                        cfm_match = self.CFM_PATTERN.search(context)
                        
                        if cfm_match and tag not in schedule_data:
                            schedule_data[tag] = VAVData(
                                tag=tag,
                                total_cfm=int(cfm_match.group(1)),
                                max_cfm=int(cfm_match.group(1)),
                                min_cfm=int(int(cfm_match.group(1)) * 0.2),
                                page=page_index
                            )
        
        return schedule_data
    
    def _estimate_inlet_size(self, cfm: int) -> str:
        """Estimate inlet size based on CFM"""
        if cfm <= 0:
            return ""
        elif cfm <= 200:
            return '6"'
        elif cfm <= 400:
            return '8"'
        elif cfm <= 700:
            return '10"'
        else:
            return '12"'
    
    def extract_all(self) -> List[VAVData]:
        """
        Main extraction method combining all strategies
        """
        all_vavs = {}
        
        # Strategy 1: Try to find schedule data first
        schedule_vavs = self._find_schedule_data()
        for tag, data in schedule_vavs.items():
            all_vavs[tag] = data
        
        # Strategy 2: Extract from floor plans using contextual grouping
        for page_index in range(len(self.doc)):
            page_vavs = self._find_vav_with_context(page_index)
            for vav in page_vavs:
                if vav.tag not in all_vavs:
                    all_vavs[vav.tag] = vav
                elif vav.total_cfm > 0 and all_vavs[vav.tag].total_cfm == 0:
                    # Update if we found CFM data
                    all_vavs[vav.tag] = vav
        
        return list(all_vavs.values())


def extract_vavs(pdf_path: str) -> List[VAVData]:
    """Convenience function"""
    extractor = ImprovedVAVExtractor(pdf_path)
    try:
        return extractor.extract_all()
    finally:
        extractor.close()


if __name__ == "__main__":
    PDF_PATH = r"D:\SW\new project\Boeing R&D Drawings.pdf"
    
    print("IMPROVED VAV EXTRACTION")
    print("=" * 60)
    
    vavs = extract_vavs(PDF_PATH)
    
    print(f"Extracted {len(vavs)} VAV units:\n")
    
    # Sort by tag
    vavs_sorted = sorted(vavs, key=lambda v: v.tag)
    
    for vav in vavs_sorted[:20]:
        print(f"{vav.tag}:")
        print(f"  Total CFM: {vav.total_cfm}")
        print(f"  Min/Max: {vav.min_cfm}/{vav.max_cfm}")
        print(f"  Inlet Size: {vav.inlet_size}")
        print(f"  Page: {vav.page + 1}")
        print()
