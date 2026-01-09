"""
Raw File Processors
Process raw files from Reports folder to create Working File sheets
"""

import pandas as pd
import os
from typing import Dict, Any, Optional
from pathlib import Path
from pyxlsb import open_workbook


class RawFileProcessor:
    """Base class for processing raw files."""
    
    def process(self, file_path: str) -> pd.DataFrame:
        """
        Process raw file and return cleaned DataFrame.
        
        Args:
            file_path: Path to raw file
            
        Returns:
            Processed DataFrame
        """
        raise NotImplementedError("Subclasses must implement process()")
    
    def _load_xlsb_sheet(self, file_path: str, sheet_name: str, header_row: int = 1) -> pd.DataFrame:
        """Load a sheet from .xlsb file."""
        rows = []
        with open_workbook(file_path) as wb:
            for row in wb.get_sheet(sheet_name):
                row_data = [cell.v if hasattr(cell, 'v') else None for cell in row]
                rows.append(row_data)
        
        if not rows:
            return pd.DataFrame()
        
        # Use header_row as column names
        if header_row < len(rows):
            columns = rows[header_row]
            data_rows = rows[header_row + 1:]
        else:
            columns = None
            data_rows = rows
        
        df = pd.DataFrame(data_rows, columns=columns)
        return df


class ConsentedStatusProcessor(RawFileProcessor):
    """Process AIL Consented Status files to create consent sheet."""
    
    def process(self, file_path: str) -> pd.DataFrame:
        """
        Process consented status file.
        
        Raw file structure:
        - Sheet: "Division Summary"
        - Header row: 1
        - Columns: Division Name, # Of DVL, Consent Received/Accepted, # Consent Require, % Consent Required
        
        Working File structure:
        - Columns: Division Name, DVL, # HCP Consent, % Consent Require
        """
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext == '.xlsb':
            # Load Division Summary sheet
            df = self._load_xlsb_sheet(file_path, "Division Summary", header_row=1)
        else:
            df = pd.read_excel(file_path, sheet_name="Division Summary", header=1)
        
        # Clean column names
        df.columns = df.columns.str.strip()
        
        # Select and rename columns to match Working File
        # Map: Division Name -> Division Name, # Of DVL -> DVL, etc.
        column_mapping = {
            'Division Name': 'Division Name',
            '# Of DVL': 'DVL',
            'Consent Received/Accepted': '# HCP Consent',
            '# Consent Require': 'Consent Require',
            '% Consent Required': '% Consent Require'
        }
        
        # Select only columns that exist
        available_cols = {}
        for old_col, new_col in column_mapping.items():
            # Try exact match first
            if old_col in df.columns:
                available_cols[old_col] = new_col
            else:
                # Try case-insensitive match
                for col in df.columns:
                    if str(col).strip().lower() == old_col.lower():
                        available_cols[col] = new_col
                        break
        
        # Select and rename columns
        result_df = df[[col for col in available_cols.keys()]].copy()
        result_df = result_df.rename(columns=available_cols)
        
        # Clean Division Name (remove extra spaces)
        if 'Division Name' in result_df.columns:
            result_df['Division Name'] = result_df['Division Name'].str.strip()
        
        # Convert % Consent Require from decimal to percentage
        if '% Consent Require' in result_df.columns:
            # If values are decimals (0-1), convert to percentage (0-100)
            sample_val = result_df['% Consent Require'].dropna().iloc[0] if len(result_df['% Consent Require'].dropna()) > 0 else None
            if sample_val is not None and sample_val < 1:
                # It's a decimal, convert to percentage
                result_df['% Consent Require'] = result_df['% Consent Require'] * 100
        
        # Remove rows where Division Name is empty/NaN
        if 'Division Name' in result_df.columns:
            result_df = result_df[result_df['Division Name'].notna()]
            result_df = result_df[result_df['Division Name'].str.strip() != '']
        
        # Ensure numeric columns are numeric
        numeric_cols = ['DVL', '# HCP Consent', 'Consent Require', '% Consent Require']
        for col in numeric_cols:
            if col in result_df.columns:
                result_df[col] = pd.to_numeric(result_df[col], errors='coerce')
        
        return result_df


class ChronicMissingProcessor(RawFileProcessor):
    """Process Chronic Missing Report files to create Chronic & Overcalling sheet."""
    
    def process(self, file_path: str, consent_data: Optional[pd.DataFrame] = None) -> pd.DataFrame:
        """
        Process chronic missing report file.
        
        Raw file structure:
        - Sheet: "Visual"
        - Header row: 0
        - Columns: User: Division, Divison Name, #HCPs
        
        Working File structure:
        - Columns: Division, #DVL, #HCPs Missed, % HCP Missed
        - Header row: 4 (row 0 is actual header)
        """
        # Load Chronic Missing Report
        df_chronic = pd.read_excel(file_path, sheet_name="Visual", header=0)
        
        # Clean column names
        df_chronic.columns = df_chronic.columns.str.strip()
        
        # Map columns
        # 'Divison Name' -> 'Division', '#HCPs' -> '#HCPs Missed'
        column_mapping = {
            'Divison Name': 'Division',
            'Division Name': 'Division',  # Handle both spellings
            '#HCPs': '#HCPs Missed'
        }
        
        # Find matching columns
        available_cols = {}
        for old_col, new_col in column_mapping.items():
            for col in df_chronic.columns:
                if str(col).strip().lower() == old_col.lower():
                    available_cols[col] = new_col
                    break
        
        # Select and rename columns
        result_df = df_chronic[[col for col in available_cols.keys()]].copy()
        result_df = result_df.rename(columns=available_cols)
        
        # Clean Division names
        if 'Division' in result_df.columns:
            result_df['Division'] = result_df['Division'].str.strip()
            result_df = result_df[result_df['Division'].notna()]
            result_df = result_df[result_df['Division'].str.strip() != '']
        
        # Get #DVL from consent data if available
        if consent_data is not None and 'Division Name' in consent_data.columns and 'DVL' in consent_data.columns:
            # Create a lookup dictionary from consent data (case-insensitive)
            consent_lookup = {}
            for _, row in consent_data.iterrows():
                div_name = str(row['Division Name']).strip().lower()
                dvl_value = row['DVL']
                consent_lookup[div_name] = dvl_value
            
            # Map DVL values to result_df using lowercase matching
            result_df['#DVL'] = result_df['Division'].str.strip().str.lower().map(consent_lookup)
        else:
            # If no consent data, we'll need to get DVL from another source
            # For now, set to None and it will need to be filled later
            result_df['#DVL'] = None
        
        # Ensure numeric columns are numeric
        numeric_cols = ['#HCPs Missed', '#DVL']
        for col in numeric_cols:
            if col in result_df.columns:
                result_df[col] = pd.to_numeric(result_df[col], errors='coerce')
        
        # Calculate % HCP Missed = (#HCPs Missed / #DVL) * 100
        if '#HCPs Missed' in result_df.columns and '#DVL' in result_df.columns:
            result_df['% HCP Missed'] = (result_df['#HCPs Missed'] / result_df['#DVL'] * 100).fillna(0)
        else:
            result_df['% HCP Missed'] = None
        
        # Rename columns to match Working File structure (as expected by config)
        # Working File uses: 'Slide 9' (Division), 'Unnamed: 1' (#DVL), 'Unnamed: 2' (#HCPs Missed), 'Unnamed: 3' (% HCP Missed)
        column_rename = {
            'Division': 'Slide 9',
            '#DVL': 'Unnamed: 1',
            '#HCPs Missed': 'Unnamed: 2',
            '% HCP Missed': 'Unnamed: 3'
        }
        
        # Only rename columns that exist
        rename_dict = {k: v for k, v in column_rename.items() if k in result_df.columns}
        result_df = result_df.rename(columns=rename_dict)
        
        # Select final columns in correct order
        final_cols = ['Slide 9', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3']
        result_df = result_df[[col for col in final_cols if col in result_df.columns]]
        
        return result_df


class RawFileProcessorFactory:
    """Factory to create appropriate processor for each raw file."""
    
    @staticmethod
    def get_processor(file_path: str) -> Optional[RawFileProcessor]:
        """
        Get appropriate processor for a file based on filename.
        
        Args:
            file_path: Path to raw file
            
        Returns:
            Processor instance or None
        """
        file_name = os.path.basename(file_path).lower()
        
        if 'consented status' in file_name or 'consent' in file_name:
            return ConsentedStatusProcessor()
        elif 'chronic missing' in file_name:
            return ChronicMissingProcessor()
        # elif 'overcalling' in file_name:
        #     return OvercallingProcessor()
        
        return None
    
    @staticmethod
    def can_process(file_path: str) -> bool:
        """Check if we have a processor for this file."""
        return RawFileProcessorFactory.get_processor(file_path) is not None

