"""
Data Normalizer
Normalizes column names, handles missing data, and performs type conversions.
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Any, Optional, Callable
import re


class DataNormalizer:
    """Normalizes and cleans data for consistent processing."""
    
    def __init__(self):
        """Initialize the data normalizer."""
        self.normalization_rules = {}
    
    def normalize_column_names(self, df: pd.DataFrame, 
                               mapping: Optional[Dict[str, str]] = None,
                               case: str = "lower",
                               remove_special: bool = True) -> pd.DataFrame:
        """
        Normalize column names.
        
        Args:
            df: DataFrame to normalize
            mapping: Optional explicit mapping of old names to new names
            case: Case conversion ('lower', 'upper', 'title', None)
            remove_special: Whether to remove special characters
        
        Returns:
            DataFrame with normalized column names
        """
        df = df.copy()
        
        if mapping:
            df = df.rename(columns=mapping)
        
        new_columns = {}
        for col in df.columns:
            new_col = str(col)
            
            # Remove special characters
            if remove_special:
                new_col = re.sub(r'[^a-zA-Z0-9\s_]', '', new_col)
                new_col = re.sub(r'\s+', '_', new_col.strip())
            
            # Case conversion
            if case == "lower":
                new_col = new_col.lower()
            elif case == "upper":
                new_col = new_col.upper()
            elif case == "title":
                new_col = new_col.title()
            
            # Remove leading/trailing underscores
            new_col = new_col.strip('_')
            
            # Ensure uniqueness
            original_new_col = new_col
            counter = 1
            while new_col in new_columns.values():
                new_col = f"{original_new_col}_{counter}"
                counter += 1
            
            new_columns[col] = new_col
        
        df = df.rename(columns=new_columns)
        return df
    
    def handle_missing_data(self, df: pd.DataFrame, 
                           strategy: str = "fill",
                           fill_value: Any = None,
                           drop_threshold: float = 0.5) -> pd.DataFrame:
        """
        Handle missing data in DataFrame.
        
        Args:
            df: DataFrame to process
            strategy: Strategy to use ('fill', 'drop', 'interpolate')
            fill_value: Value to fill with (if strategy is 'fill')
            drop_threshold: Threshold for dropping columns/rows (if strategy is 'drop')
        
        Returns:
            DataFrame with missing data handled
        """
        df = df.copy()
        
        if strategy == "fill":
            if fill_value is not None:
                df = df.fillna(fill_value)
            else:
                # Fill numeric columns with 0, object columns with empty string
                for col in df.columns:
                    if pd.api.types.is_numeric_dtype(df[col]):
                        df[col] = df[col].fillna(0)
                    else:
                        df[col] = df[col].fillna("")
        
        elif strategy == "drop":
            # Drop columns with too many missing values
            col_missing_ratio = df.isnull().sum() / len(df)
            cols_to_drop = col_missing_ratio[col_missing_ratio > drop_threshold].index
            df = df.drop(columns=cols_to_drop)
            
            # Drop rows with too many missing values
            row_missing_ratio = df.isnull().sum(axis=1) / len(df.columns)
            rows_to_drop = row_missing_ratio[row_missing_ratio > drop_threshold].index
            df = df.drop(index=rows_to_drop)
        
        elif strategy == "interpolate":
            # Interpolate numeric columns
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            df[numeric_cols] = df[numeric_cols].interpolate(method='linear')
            df[numeric_cols] = df[numeric_cols].fillna(method='bfill').fillna(method='ffill')
        
        return df
    
    def convert_types(self, df: pd.DataFrame, 
                     type_mapping: Optional[Dict[str, str]] = None,
                     auto_detect: bool = True) -> pd.DataFrame:
        """
        Convert column types.
        
        Args:
            df: DataFrame to convert
            type_mapping: Optional explicit mapping of column names to types
            auto_detect: Whether to auto-detect and convert types
        
        Returns:
            DataFrame with converted types
        """
        # Ensure df is a DataFrame
        if not isinstance(df, pd.DataFrame):
            return df
        
        df = df.copy()
        
        # Apply explicit type mapping
        if type_mapping:
            for col, dtype in type_mapping.items():
                if col in df.columns:
                    try:
                        df[col] = df[col].astype(dtype)
                    except Exception as e:
                        print(f"Warning: Could not convert column '{col}' to {dtype}: {e}")
        
        # Auto-detect and convert types
        if auto_detect:
            df = self._auto_convert_types(df)
        
        return df
    
    def _auto_convert_types(self, df: pd.DataFrame) -> pd.DataFrame:
        """Auto-detect and convert column types."""
        if not isinstance(df, pd.DataFrame):
            return df
        
        if len(df) == 0:
            return df
        
        for col in df.columns:
            try:
                # Get the column as a Series
                col_series = df[col]
                
                # Skip if not a Series (shouldn't happen, but safety check)
                if not isinstance(col_series, pd.Series):
                    continue
                
                # Try to convert to numeric
                if col_series.dtype == 'object':
                    # Check if it's numeric strings
                    numeric_series = pd.to_numeric(col_series, errors='coerce')
                    if len(df) > 0 and numeric_series.notna().sum() / len(df) > 0.8:
                        df[col] = numeric_series
                        continue
                    
                    # Check if it's dates
                    date_series = pd.to_datetime(col_series, errors='coerce')
                    if len(df) > 0 and date_series.notna().sum() / len(df) > 0.8:
                        df[col] = date_series
                        continue
                    
                    # Check if it's boolean strings
                    if col_series.nunique() == 2:
                        unique_vals = col_series.dropna().unique()
                        if set(unique_vals).issubset({'True', 'False', 'true', 'false', '1', '0', 'Yes', 'No', 'yes', 'no'}):
                            df[col] = col_series.map({'True': True, 'False': False, 'true': True, 'false': False,
                                                  '1': True, '0': False, 'Yes': True, 'No': False,
                                                  'yes': True, 'no': False})
            except (KeyError, AttributeError, TypeError) as e:
                # Skip columns that cause errors
                print(f"Warning: Could not auto-convert column '{col}': {e}")
                continue
        
        return df
    
    def normalize_data(self, df: pd.DataFrame, 
                     column_mapping: Optional[Dict[str, str]] = None,
                     missing_strategy: str = "fill",
                     type_mapping: Optional[Dict[str, str]] = None) -> pd.DataFrame:
        """
        Complete normalization pipeline.
        
        Args:
            df: DataFrame to normalize
            column_mapping: Optional column name mapping
            missing_strategy: Strategy for handling missing data
            type_mapping: Optional type conversion mapping
        
        Returns:
            Normalized DataFrame
        """
        df = self.normalize_column_names(df, mapping=column_mapping)
        df = self.handle_missing_data(df, strategy=missing_strategy)
        df = self.convert_types(df, type_mapping=type_mapping)
        
        return df
    
    def standardize_values(self, df: pd.DataFrame, 
                          column: str,
                          mapping: Dict[str, str]) -> pd.DataFrame:
        """
        Standardize values in a column using a mapping.
        
        Args:
            df: DataFrame to process
            column: Column name to standardize
            mapping: Mapping of old values to new values
        
        Returns:
            DataFrame with standardized values
        """
        df = df.copy()
        if column in df.columns:
            df[column] = df[column].map(mapping).fillna(df[column])
        return df


if __name__ == "__main__":
    # Example usage
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python data_normalizer.py <excel_file>")
        sys.exit(1)
    
    excel_file = sys.argv[1]
    
    import pandas as pd
    from data_loader import DataLoader
    
    loader = DataLoader()
    df = loader.load_excel(excel_file)
    
    if isinstance(df, dict):
        df = list(df.values())[0]
    
    normalizer = DataNormalizer()
    normalized_df = normalizer.normalize_data(df)
    
    print(f"Original shape: {df.shape}")
    print(f"Normalized shape: {normalized_df.shape}")
    print(f"\nOriginal columns: {list(df.columns)}")
    print(f"Normalized columns: {list(normalized_df.columns)}")

