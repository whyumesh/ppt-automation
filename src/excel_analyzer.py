"""
Excel File Analyzer
Analyzes Excel file structures, identifies schemas, and maps files to slides.
"""

import pandas as pd
import os
from typing import Dict, List, Any, Optional, Union
import json
from pathlib import Path
import openpyxl
from pyxlsb import open_workbook as open_xlsb


class ExcelAnalyzer:
    """Analyzes Excel files to understand their structure and content."""
    
    def __init__(self, excel_path: str):
        """
        Initialize the Excel analyzer.
        
        Args:
            excel_path: Path to the Excel file to analyze
        """
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Excel file not found: {excel_path}")
        
        self.excel_path = excel_path
        self.file_extension = os.path.splitext(excel_path)[1].lower()
        self.analysis_info = {
            "file_path": excel_path,
            "file_name": os.path.basename(excel_path),
            "file_type": self.file_extension,
            "sheets": []
        }
    
    def analyze_all(self) -> Dict[str, Any]:
        """
        Analyze all sheets in the Excel file.
        
        Returns:
            Dictionary containing analysis information
        """
        if self.file_extension == '.xlsb':
            self._analyze_xlsb()
        else:
            self._analyze_xlsx()
        
        return self.analysis_info
    
    def _analyze_xlsx(self):
        """Analyze .xlsx file using pandas and openpyxl."""
        try:
            # Get sheet names
            excel_file = pd.ExcelFile(self.excel_path)
            sheet_names = excel_file.sheet_names
            
            for sheet_name in sheet_names:
                try:
                    df = pd.read_excel(self.excel_path, sheet_name=sheet_name, nrows=1000)
                    sheet_info = self._analyze_sheet(df, sheet_name)
                    self.analysis_info["sheets"].append(sheet_info)
                except Exception as e:
                    self.analysis_info["sheets"].append({
                        "name": sheet_name,
                        "error": str(e)
                    })
        except Exception as e:
            self.analysis_info["error"] = str(e)
    
    def _analyze_xlsb(self):
        """Analyze .xlsb file using pyxlsb."""
        try:
            with open_xlsb(self.excel_path) as wb:
                for sheet_name in wb.sheets:
                    try:
                        # Read first 1000 rows for analysis
                        rows = []
                        for i, row in enumerate(wb.get_sheet(sheet_name)):
                            if i >= 1000:
                                break
                            rows.append([cell.v if hasattr(cell, 'v') else None for cell in row])
                        
                        if rows:
                            # Convert to DataFrame
                            df = pd.DataFrame(rows[1:], columns=rows[0] if len(rows) > 1 else None)
                            sheet_info = self._analyze_sheet(df, sheet_name)
                            self.analysis_info["sheets"].append(sheet_info)
                        else:
                            self.analysis_info["sheets"].append({
                                "name": sheet_name,
                                "row_count": 0,
                                "columns": []
                            })
                    except Exception as e:
                        self.analysis_info["sheets"].append({
                            "name": sheet_name,
                            "error": str(e)
                        })
        except Exception as e:
            self.analysis_info["error"] = str(e)
    
    def _analyze_sheet(self, df: pd.DataFrame, sheet_name: str) -> Dict[str, Any]:
        """
        Analyze a single sheet DataFrame.
        
        Args:
            df: DataFrame to analyze
            sheet_name: Name of the sheet
        
        Returns:
            Dictionary containing sheet analysis information
        """
        sheet_info = {
            "name": sheet_name,
            "row_count": len(df),
            "column_count": len(df.columns),
            "columns": []
        }
        
        # Analyze each column
        for col_name in df.columns:
            col_info = self._analyze_column(df, col_name)
            sheet_info["columns"].append(col_info)
        
        # Detect potential header rows
        sheet_info["has_header"] = self._detect_header(df)
        
        # Sample data (first few rows)
        sheet_info["sample_data"] = df.head(5).to_dict('records') if len(df) > 0 else []
        
        # Detect data types
        sheet_info["data_types"] = df.dtypes.astype(str).to_dict()
        
        return sheet_info
    
    def _analyze_column(self, df: pd.DataFrame, col_name: str) -> Dict[str, Any]:
        """
        Analyze a single column.
        
        Args:
            df: DataFrame containing the column
            col_name: Name of the column
        
        Returns:
            Dictionary containing column analysis information
        """
        col_info = {
            "name": str(col_name),
            "data_type": str(df[col_name].dtype),
            "non_null_count": df[col_name].notna().sum(),
            "null_count": df[col_name].isna().sum(),
            "unique_count": df[col_name].nunique()
        }
        
        # Statistical information for numeric columns
        if pd.api.types.is_numeric_dtype(df[col_name]):
            col_info["statistics"] = {
                "min": float(df[col_name].min()) if df[col_name].notna().any() else None,
                "max": float(df[col_name].max()) if df[col_name].notna().any() else None,
                "mean": float(df[col_name].mean()) if df[col_name].notna().any() else None,
                "median": float(df[col_name].median()) if df[col_name].notna().any() else None,
                "std": float(df[col_name].std()) if df[col_name].notna().any() else None
            }
        
        # Sample values
        non_null_values = df[col_name].dropna()
        if len(non_null_values) > 0:
            col_info["sample_values"] = non_null_values.head(5).tolist()
        
        # Detect if column contains dates
        if pd.api.types.is_datetime64_any_dtype(df[col_name]):
            col_info["is_date"] = True
            col_info["date_range"] = {
                "min": str(df[col_name].min()) if df[col_name].notna().any() else None,
                "max": str(df[col_name].max()) if df[col_name].notna().any() else None
            }
        
        # Detect if column contains percentages
        if df[col_name].dtype == 'object':
            sample_str = str(df[col_name].dropna().iloc[0]) if len(df[col_name].dropna()) > 0 else ""
            if '%' in sample_str or (isinstance(sample_str, str) and sample_str.endswith('%')):
                col_info["likely_percentage"] = True
        
        return col_info
    
    def _detect_header(self, df: pd.DataFrame) -> bool:
        """
        Detect if the DataFrame has a header row.
        
        Args:
            df: DataFrame to analyze
        
        Returns:
            True if header is detected, False otherwise
        """
        if len(df) == 0:
            return False
        
        # Simple heuristic: if first row contains mostly strings and subsequent rows contain data
        first_row_types = [type(val).__name__ for val in df.iloc[0]]
        string_count = sum(1 for t in first_row_types if t == 'str')
        
        return string_count > len(df.columns) * 0.5
    
    def get_schema(self) -> Dict[str, Any]:
        """
        Extract schema information from the Excel file.
        
        Returns:
            Dictionary containing schema information
        """
        schema = {
            "file_path": self.excel_path,
            "file_name": os.path.basename(self.excel_path),
            "sheets": []
        }
        
        for sheet_info in self.analysis_info["sheets"]:
            if "error" not in sheet_info:
                sheet_schema = {
                    "name": sheet_info["name"],
                    "columns": [
                        {
                            "name": col["name"],
                            "data_type": col["data_type"],
                            "nullable": col["null_count"] > 0
                        }
                        for col in sheet_info["columns"]
                    ]
                }
                schema["sheets"].append(sheet_schema)
        
        return schema
    
    def save_analysis(self, output_path: str):
        """
        Save analysis information to a JSON file.
        
        Args:
            output_path: Path to save the JSON file
        """
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(self.analysis_info, f, indent=2, ensure_ascii=False, default=str)


def analyze_excel_file(excel_path: str, output_json: Optional[str] = None) -> Dict[str, Any]:
    """
    Convenience function to analyze an Excel file.
    
    Args:
        excel_path: Path to the Excel file
        output_json: Optional path to save JSON output
    
    Returns:
        Dictionary containing analysis information
    """
    analyzer = ExcelAnalyzer(excel_path)
    analysis_info = analyzer.analyze_all()
    
    if output_json:
        analyzer.save_analysis(output_json)
    
    return analysis_info


def analyze_directory(directory_path: str, pattern: str = "*.xlsx") -> List[Dict[str, Any]]:
    """
    Analyze all Excel files in a directory.
    
    Args:
        directory_path: Path to directory containing Excel files
        pattern: File pattern to match (default: "*.xlsx")
    
    Returns:
        List of analysis dictionaries
    """
    results = []
    directory = Path(directory_path)
    
    # Find all matching files
    for file_path in directory.rglob(pattern):
        if file_path.is_file():
            try:
                analyzer = ExcelAnalyzer(str(file_path))
                analysis_info = analyzer.analyze_all()
                results.append(analysis_info)
            except Exception as e:
                results.append({
                    "file_path": str(file_path),
                    "error": str(e)
                })
    
    # Also check for .xlsb files
    for file_path in directory.rglob("*.xlsb"):
        if file_path.is_file():
            try:
                analyzer = ExcelAnalyzer(str(file_path))
                analysis_info = analyzer.analyze_all()
                results.append(analysis_info)
            except Exception as e:
                results.append({
                    "file_path": str(file_path),
                    "error": str(e)
                })
    
    return results


if __name__ == "__main__":
    # Example usage
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python excel_analyzer.py <excel_file> [output_json]")
        print("   or: python excel_analyzer.py <directory> --directory")
        sys.exit(1)
    
    input_path = sys.argv[1]
    output_json = sys.argv[2] if len(sys.argv) > 2 and sys.argv[2] != "--directory" else None
    
    if len(sys.argv) > 2 and sys.argv[2] == "--directory":
        results = analyze_directory(input_path)
        print(f"Analyzed {len(results)} files")
        for result in results:
            print(f"  - {result.get('file_name', result.get('file_path', 'Unknown'))}")
    else:
        info = analyze_excel_file(input_path, output_json)
        print(f"Analyzed {info.get('file_name', 'file')}")
        print(f"Found {len(info.get('sheets', []))} sheets")

