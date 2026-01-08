"""
Data Loader
Loads Excel files, handles multiple sheets, and validates schemas.
"""

import pandas as pd
import os
from typing import Dict, List, Any, Optional, Union
from pathlib import Path
import openpyxl
from pyxlsb import open_workbook as open_xlsb
import yaml


class DataLoader:
    """Loads and validates Excel files."""
    
    def __init__(self, schema_config: Optional[str] = None):
        """
        Initialize the data loader.
        
        Args:
            schema_config: Optional path to schema configuration YAML file
        """
        self.schema_config = schema_config
        self.schemas = {}
        
        if schema_config and os.path.exists(schema_config):
            self._load_schemas()
    
    def _load_schemas(self):
        """Load schema definitions from configuration file."""
        with open(self.schema_config, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
            self.schemas = config.get("schemas", {})
    
    def load_excel(self, excel_path: str, sheet_name: Optional[str] = None, 
                   header_row: int = 0) -> Union[pd.DataFrame, Dict[str, pd.DataFrame]]:
        """
        Load an Excel file.
        
        Args:
            excel_path: Path to the Excel file
            sheet_name: Optional sheet name to load (if None, loads all sheets)
            header_row: Row number to use as header (0-indexed)
        
        Returns:
            DataFrame if sheet_name is specified, otherwise dict of DataFrames
        """
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Excel file not found: {excel_path}")
        
        file_extension = os.path.splitext(excel_path)[1].lower()
        
        if file_extension == '.xlsb':
            return self._load_xlsb(excel_path, sheet_name, header_row)
        else:
            return self._load_xlsx(excel_path, sheet_name, header_row)
    
    def _load_xlsx(self, excel_path: str, sheet_name: Optional[str] = None,
                   header_row: int = 0) -> Union[pd.DataFrame, Dict[str, pd.DataFrame]]:
        """Load .xlsx file."""
        if sheet_name:
            df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header_row)
            return df
        else:
            excel_file = pd.ExcelFile(excel_path)
            sheets = {}
            for sheet in excel_file.sheet_names:
                try:
                    df = pd.read_excel(excel_path, sheet_name=sheet, header=header_row)
                    sheets[sheet] = df
                except Exception as e:
                    print(f"Warning: Could not load sheet '{sheet}': {e}")
            return sheets
    
    def _load_xlsb(self, excel_path: str, sheet_name: Optional[str] = None,
                   header_row: int = 0) -> Union[pd.DataFrame, Dict[str, pd.DataFrame]]:
        """Load .xlsb file."""
        with open_xlsb(excel_path) as wb:
            if sheet_name:
                return self._read_xlsb_sheet(wb, sheet_name, header_row)
            else:
                sheets = {}
                for sheet in wb.sheets:
                    try:
                        df = self._read_xlsb_sheet(wb, sheet, header_row)
                        sheets[sheet] = df
                    except Exception as e:
                        print(f"Warning: Could not load sheet '{sheet}': {e}")
                return sheets
    
    def _read_xlsb_sheet(self, wb, sheet_name: str, header_row: int = 0) -> pd.DataFrame:
        """Read a single sheet from xlsb workbook."""
        rows = []
        for row in wb.get_sheet(sheet_name):
            rows.append([cell.v if hasattr(cell, 'v') else None for cell in row])
        
        if not rows:
            return pd.DataFrame()
        
        # Use header_row as column names if specified
        if header_row < len(rows):
            columns = rows[header_row]
            data_rows = rows[header_row + 1:]
        else:
            columns = None
            data_rows = rows
        
        df = pd.DataFrame(data_rows, columns=columns)
        return df
    
    def validate_schema(self, df: pd.DataFrame, schema_name: str) -> Dict[str, Any]:
        """
        Validate a DataFrame against a schema definition.
        
        Args:
            df: DataFrame to validate
            schema_name: Name of the schema to validate against
        
        Returns:
            Dictionary containing validation results
        """
        if schema_name not in self.schemas:
            return {
                "valid": False,
                "error": f"Schema '{schema_name}' not found in configuration"
            }
        
        schema = self.schemas[schema_name]
        validation_result = {
            "valid": True,
            "errors": [],
            "warnings": []
        }
        
        # Check required columns
        required_columns = schema.get("required_columns", [])
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            validation_result["valid"] = False
            validation_result["errors"].append(f"Missing required columns: {missing_columns}")
        
        # Check data types
        column_types = schema.get("column_types", {})
        for col, expected_type in column_types.items():
            if col in df.columns:
                actual_type = str(df[col].dtype)
                if not self._types_match(actual_type, expected_type):
                    validation_result["warnings"].append(
                        f"Column '{col}' has type {actual_type}, expected {expected_type}"
                    )
        
        return validation_result
    
    def _types_match(self, actual: str, expected: str) -> bool:
        """Check if actual type matches expected type."""
        type_mapping = {
            "int64": ["int", "integer", "int64"],
            "float64": ["float", "numeric", "float64", "number"],
            "object": ["string", "str", "object", "text"],
            "datetime64[ns]": ["date", "datetime", "datetime64[ns]"],
            "bool": ["bool", "boolean"]
        }
        
        for key, values in type_mapping.items():
            if actual == key and expected.lower() in values:
                return True
            if expected.lower() == key and actual in values:
                return True
        
        return False
    
    def load_multiple_files(self, file_paths: List[str], 
                           sheet_names: Optional[Dict[str, str]] = None) -> Dict[str, Union[pd.DataFrame, Dict[str, pd.DataFrame]]]:
        """
        Load multiple Excel files.
        
        Args:
            file_paths: List of file paths to load
            sheet_names: Optional dict mapping file paths to sheet names
        
        Returns:
            Dictionary mapping file paths to loaded DataFrames
        """
        loaded_data = {}
        
        for file_path in file_paths:
            sheet_name = sheet_names.get(file_path) if sheet_names else None
            try:
                data = self.load_excel(file_path, sheet_name=sheet_name)
                loaded_data[file_path] = data
            except Exception as e:
                print(f"Error loading {file_path}: {e}")
                loaded_data[file_path] = None
        
        return loaded_data


if __name__ == "__main__":
    # Example usage
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python data_loader.py <excel_file> [sheet_name] [schema_config]")
        sys.exit(1)
    
    excel_file = sys.argv[1]
    sheet_name = sys.argv[2] if len(sys.argv) > 2 else None
    schema_config = sys.argv[3] if len(sys.argv) > 3 else None
    
    loader = DataLoader(schema_config=schema_config)
    data = loader.load_excel(excel_file, sheet_name=sheet_name)
    
    if isinstance(data, pd.DataFrame):
        print(f"Loaded DataFrame with shape: {data.shape}")
        print(f"Columns: {list(data.columns)}")
    else:
        print(f"Loaded {len(data)} sheets:")
        for sheet_name, df in data.items():
            print(f"  {sheet_name}: {df.shape}")

