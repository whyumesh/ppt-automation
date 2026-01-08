"""
Rule Discoverer
Compares Excel data with PPT slides to discover business rules, calculations, and formatting logic.
"""

import pandas as pd
from typing import Dict, List, Any, Optional, Tuple
import json
from pathlib import Path
try:
    from .template_extractor import TemplateExtractor
    from .excel_analyzer import ExcelAnalyzer
except ImportError:
    from template_extractor import TemplateExtractor
    from excel_analyzer import ExcelAnalyzer
import re
from decimal import Decimal


class RuleDiscoverer:
    """Discovers business rules by comparing Excel data with PowerPoint slides."""
    
    def __init__(self, excel_path: str, ppt_path: str):
        """
        Initialize the rule discoverer.
        
        Args:
            excel_path: Path to the Excel file
            ppt_path: Path to the PowerPoint file
        """
        self.excel_path = excel_path
        self.ppt_path = ppt_path
        self.excel_analyzer = ExcelAnalyzer(excel_path)
        self.template_extractor = TemplateExtractor(ppt_path)
        self.discovered_rules = {
            "excel_file": excel_path,
            "ppt_file": ppt_path,
            "slide_mappings": [],
            "calculation_rules": [],
            "formatting_rules": [],
            "filtering_rules": [],
            "text_generation_rules": []
        }
    
    def discover_all(self) -> Dict[str, Any]:
        """
        Discover all rules by comparing Excel and PPT.
        
        Returns:
            Dictionary containing discovered rules
        """
        # Extract data from both sources
        excel_info = self.excel_analyzer.analyze_all()
        ppt_info = self.template_extractor.extract_all()
        
        # Discover mappings and rules
        self._discover_slide_mappings(excel_info, ppt_info)
        self._discover_calculation_rules(excel_info, ppt_info)
        self._discover_formatting_rules(excel_info, ppt_info)
        self._discover_filtering_rules(excel_info, ppt_info)
        self._discover_text_generation_rules(excel_info, ppt_info)
        
        return self.discovered_rules
    
    def _discover_slide_mappings(self, excel_info: Dict, ppt_info: Dict):
        """Discover mappings between Excel data and PPT slides."""
        for slide_info in ppt_info.get("slides", []):
            slide_mapping = {
                "slide_number": slide_info["slide_number"],
                "slide_layout": slide_info.get("layout_name", "Unknown"),
                "data_sources": [],
                "shape_mappings": []
            }
            
            # Analyze each shape on the slide
            for shape in slide_info.get("shapes", []):
                if shape.get("type") == "text_box":
                    # Try to find matching data in Excel
                    text_content = self._extract_text_content(shape)
                    excel_match = self._find_excel_match(text_content, excel_info)
                    
                    if excel_match:
                        shape_mapping = {
                            "shape_index": shape["index"],
                            "shape_name": shape.get("name", ""),
                            "excel_match": excel_match,
                            "mapping_type": "text"
                        }
                        slide_mapping["shape_mappings"].append(shape_mapping)
                
                elif shape.get("type") == "table":
                    # Analyze table structure
                    table_match = self._find_table_match(shape, excel_info)
                    if table_match:
                        shape_mapping = {
                            "shape_index": shape["index"],
                            "shape_name": shape.get("name", ""),
                            "excel_match": table_match,
                            "mapping_type": "table"
                        }
                        slide_mapping["shape_mappings"].append(shape_mapping)
            
            if slide_mapping["shape_mappings"]:
                self.discovered_rules["slide_mappings"].append(slide_mapping)
    
    def _extract_text_content(self, shape: Dict) -> str:
        """Extract all text content from a shape."""
        text_parts = []
        for para in shape.get("text_content", []):
            text_parts.append(para.get("text", ""))
        return " ".join(text_parts)
    
    def _find_excel_match(self, text: str, excel_info: Dict) -> Optional[Dict]:
        """Try to find matching data in Excel for given text."""
        # Extract numbers from text
        numbers = self._extract_numbers(text)
        
        if not numbers:
            return None
        
        # Search through Excel sheets for matching values
        for sheet_info in excel_info.get("sheets", []):
            if "error" in sheet_info:
                continue
            
            for col_info in sheet_info.get("columns", []):
                # Check if column contains matching values
                sample_values = col_info.get("sample_values", [])
                for num in numbers:
                    if self._values_match(num, sample_values):
                        return {
                            "sheet": sheet_info["name"],
                            "column": col_info["name"],
                            "matched_value": num
                        }
        
        return None
    
    def _find_table_match(self, shape: Dict, excel_info: Dict) -> Optional[Dict]:
        """Try to find matching table data in Excel."""
        table_info = shape.get("table_info", {})
        if not table_info:
            return None
        
        # Extract table structure
        rows = table_info.get("rows", 0)
        cols = table_info.get("columns", 0)
        
        # Try to find matching table structure in Excel
        for sheet_info in excel_info.get("sheets", []):
            if "error" in sheet_info:
                continue
            
            if sheet_info.get("row_count", 0) >= rows and sheet_info.get("column_count", 0) >= cols:
                # Potential match - extract cell values for comparison
                return {
                    "sheet": sheet_info["name"],
                    "rows": rows,
                    "columns": cols
                }
        
        return None
    
    def _extract_numbers(self, text: str) -> List[float]:
        """Extract numeric values from text."""
        # Pattern to match numbers (including percentages, decimals, etc.)
        patterns = [
            r'\d+\.\d+',  # Decimal numbers
            r'\d+',       # Integers
            r'\d+%',      # Percentages
        ]
        
        numbers = []
        for pattern in patterns:
            matches = re.findall(pattern, text)
            for match in matches:
                try:
                    if match.endswith('%'):
                        num = float(match[:-1]) / 100
                    else:
                        num = float(match)
                    numbers.append(num)
                except ValueError:
                    continue
        
        return numbers
    
    def _values_match(self, value1: float, values: List[Any], tolerance: float = 0.01) -> bool:
        """Check if value1 matches any value in values list."""
        for val in values:
            try:
                val_float = float(val)
                if abs(value1 - val_float) < tolerance:
                    return True
            except (ValueError, TypeError):
                continue
        return False
    
    def _discover_calculation_rules(self, excel_info: Dict, ppt_info: Dict):
        """Discover calculation patterns (aggregations, deltas, percentages, rankings)."""
        # This is a simplified version - in practice, would need more sophisticated analysis
        for slide_info in ppt_info.get("slides", []):
            for shape in slide_info.get("shapes", []):
                text_content = self._extract_text_content(shape)
                
                # Look for percentage patterns
                if '%' in text_content:
                    rule = {
                        "slide_number": slide_info["slide_number"],
                        "shape_index": shape["index"],
                        "rule_type": "percentage_calculation",
                        "description": "Percentage value detected in slide"
                    }
                    self.discovered_rules["calculation_rules"].append(rule)
                
                # Look for delta/change patterns
                if any(word in text_content.lower() for word in ['change', 'delta', 'increase', 'decrease', 'growth']):
                    rule = {
                        "slide_number": slide_info["slide_number"],
                        "shape_index": shape["index"],
                        "rule_type": "delta_calculation",
                        "description": "Delta/change calculation detected"
                    }
                    self.discovered_rules["calculation_rules"].append(rule)
    
    def _discover_formatting_rules(self, excel_info: Dict, ppt_info: Dict):
        """Discover formatting rules (rounding, number formats, conditional formatting)."""
        for slide_info in ppt_info.get("slides", []):
            for shape in slide_info.get("shapes", []):
                # Extract formatting information
                if shape.get("type") == "text_box":
                    for para in shape.get("text_content", []):
                        for run in para.get("runs", []):
                            # Check font colors
                            if "font_color" in run:
                                color = run["font_color"]
                                rule = {
                                    "slide_number": slide_info["slide_number"],
                                    "shape_index": shape["index"],
                                    "rule_type": "color_formatting",
                                    "color": color,
                                    "description": f"Text color formatting: RGB({color.get('r')}, {color.get('g')}, {color.get('b')})"
                                }
                                self.discovered_rules["formatting_rules"].append(rule)
                            
                            # Check font size
                            if run.get("font_size"):
                                rule = {
                                    "slide_number": slide_info["slide_number"],
                                    "shape_index": shape["index"],
                                    "rule_type": "font_size",
                                    "size": run["font_size"],
                                    "description": f"Font size: {run['font_size']}pt"
                                }
                                self.discovered_rules["formatting_rules"].append(rule)
                
                elif shape.get("type") == "table":
                    # Analyze table cell formatting
                    table_info = shape.get("table_info", {})
                    for cell in table_info.get("cells", []):
                        if cell.get("fill_color"):
                            rule = {
                                "slide_number": slide_info["slide_number"],
                                "shape_index": shape["index"],
                                "rule_type": "cell_color",
                                "row": cell["row"],
                                "column": cell["column"],
                                "color": cell["fill_color"],
                                "description": "Cell background color"
                            }
                            self.discovered_rules["formatting_rules"].append(rule)
    
    def _discover_filtering_rules(self, excel_info: Dict, ppt_info: Dict):
        """Discover filtering rules and thresholds."""
        # Compare Excel data volume with PPT content
        for sheet_info in excel_info.get("sheets", []):
            if "error" in sheet_info:
                continue
            
            excel_row_count = sheet_info.get("row_count", 0)
            
            # Try to find corresponding slide
            for slide_info in ppt_info.get("slides", []):
                # Count text elements or table rows in slide
                slide_data_count = self._count_slide_data_elements(slide_info)
                
                if slide_data_count > 0 and slide_data_count < excel_row_count:
                    rule = {
                        "slide_number": slide_info["slide_number"],
                        "excel_sheet": sheet_info["name"],
                        "excel_rows": excel_row_count,
                        "slide_elements": slide_data_count,
                        "rule_type": "filtering",
                        "description": f"Data filtered from {excel_row_count} rows to {slide_data_count} elements"
                    }
                    self.discovered_rules["filtering_rules"].append(rule)
    
    def _count_slide_data_elements(self, slide_info: Dict) -> int:
        """Count data elements (rows, text items) in a slide."""
        count = 0
        
        for shape in slide_info.get("shapes", []):
            if shape.get("type") == "table":
                count += shape.get("table_info", {}).get("rows", 0)
            elif shape.get("type") == "text_box":
                # Count non-empty paragraphs
                for para in shape.get("text_content", []):
                    if para.get("text", "").strip():
                        count += 1
        
        return count
    
    def _discover_text_generation_rules(self, excel_info: Dict, ppt_info: Dict):
        """Discover text generation patterns."""
        for slide_info in ppt_info.get("slides", []):
            for shape in slide_info.get("shapes", []):
                text_content = self._extract_text_content(shape)
                
                # Look for conditional text patterns
                if any(word in text_content.lower() for word in ['improved', 'declined', 'increased', 'decreased']):
                    rule = {
                        "slide_number": slide_info["slide_number"],
                        "shape_index": shape["index"],
                        "rule_type": "conditional_text",
                        "text_pattern": text_content,
                        "description": "Conditional text generation detected"
                    }
                    self.discovered_rules["text_generation_rules"].append(rule)
    
    def save_rules(self, output_path: str):
        """Save discovered rules to a JSON file."""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(self.discovered_rules, f, indent=2, ensure_ascii=False, default=str)


def discover_rules(excel_path: str, ppt_path: str, output_json: Optional[str] = None) -> Dict[str, Any]:
    """
    Convenience function to discover rules from Excel and PPT files.
    
    Args:
        excel_path: Path to Excel file
        ppt_path: Path to PowerPoint file
        output_json: Optional path to save JSON output
    
    Returns:
        Dictionary containing discovered rules
    """
    discoverer = RuleDiscoverer(excel_path, ppt_path)
    rules = discoverer.discover_all()
    
    if output_json:
        discoverer.save_rules(output_json)
    
    return rules


if __name__ == "__main__":
    # Example usage
    import sys
    
    if len(sys.argv) < 3:
        print("Usage: python rule_discoverer.py <excel_file> <ppt_file> [output_json]")
        sys.exit(1)
    
    excel_file = sys.argv[1]
    ppt_file = sys.argv[2]
    output_json = sys.argv[3] if len(sys.argv) > 3 else None
    
    rules = discover_rules(excel_file, ppt_file, output_json)
    print(f"Discovered rules from {excel_file} and {ppt_file}")
    print(f"Found {len(rules['slide_mappings'])} slide mappings")
    print(f"Found {len(rules['calculation_rules'])} calculation rules")
    print(f"Found {len(rules['formatting_rules'])} formatting rules")

