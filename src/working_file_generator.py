"""
Working File Generator
Generates Working File structure from processed raw files
"""

import pandas as pd
import os
from typing import Dict, Any, Optional
from src.raw_file_processors import RawFileProcessorFactory, ChronicMissingProcessor


class WorkingFileGenerator:
    """Generate Working File structure from processed raw files."""
    
    def __init__(self):
        self.processed_sheets = {}
    
    def add_processed_sheet(self, sheet_name: str, data: pd.DataFrame):
        """Add a processed sheet to the working file."""
        self.processed_sheets[sheet_name] = data
    
    def generate_from_raw_files(self, raw_files: Dict[str, str]) -> Dict[str, pd.DataFrame]:
        """
        Generate Working File structure from raw files.
        
        Args:
            raw_files: Dictionary mapping sheet names to raw file paths
                      Example: {"consent": "path/to/consent_file.xlsb"}
        
        Returns:
            Dictionary of sheets (like Working File structure)
        """
        working_file = {}
        consent_data = None
        
        # First pass: Process consent file (needed for other processors)
        if 'consent' in raw_files:
            consent_path = raw_files['consent']
            processor = RawFileProcessorFactory.get_processor(consent_path)
            if processor:
                try:
                    consent_data = processor.process(consent_path)
                    working_file['consent'] = consent_data
                    print(f"  ✓ Processed consent from {os.path.basename(consent_path)}")
                except Exception as e:
                    print(f"  ✗ Error processing consent: {e}")
        
        # Second pass: Process other files (may need consent data)
        for sheet_name, file_path in raw_files.items():
            if sheet_name == 'consent':
                continue  # Already processed
            
            processor = RawFileProcessorFactory.get_processor(file_path)
            if processor:
                try:
                    # Pass consent_data to processors that need it
                    if isinstance(processor, ChronicMissingProcessor):
                        processed_data = processor.process(file_path, consent_data=consent_data)
                    else:
                        processed_data = processor.process(file_path)
                    working_file[sheet_name] = processed_data
                    print(f"  ✓ Processed {sheet_name} from {os.path.basename(file_path)}")
                except Exception as e:
                    print(f"  ✗ Error processing {sheet_name}: {e}")
            else:
                print(f"  ⚠ No processor found for {file_path}")
        
        return working_file
    
    def generate(self) -> Dict[str, pd.DataFrame]:
        """
        Generate Working File from already processed sheets.
        
        Returns:
            Dictionary of sheets
        """
        return self.processed_sheets.copy()

