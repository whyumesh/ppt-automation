"""
Main Entry Point
Orchestrates the full pipeline: load Excel → process data → apply rules → generate PPT
"""

import argparse
import os
import sys
from pathlib import Path
from typing import Dict, Any, Optional
import pandas as pd

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from data_loader import DataLoader
from data_normalizer import DataNormalizer
from transformations import DataTransformations
from rules_engine import RulesEngine
from ppt_generator import PPTGenerator


class PPTPipeline:
    """Main pipeline for generating PowerPoint decks from Excel files."""
    
    def __init__(self, config_dir: str = "config", template_path: Optional[str] = None):
        """
        Initialize the pipeline.
        
        Args:
            config_dir: Directory containing configuration files
            template_path: Path to PowerPoint template file
        """
        self.config_dir = config_dir
        self.template_path = template_path or os.path.join("templates", "template.pptx")
        
        # Initialize components
        schema_config = os.path.join(config_dir, "schema.yaml")
        self.data_loader = DataLoader(schema_config=schema_config)
        self.data_normalizer = DataNormalizer()
        self.transformations = DataTransformations()
        
        rules_config = os.path.join(config_dir, "rules.yaml")
        self.rules_engine = RulesEngine(rules_config=rules_config)
        
        slides_config = os.path.join(config_dir, "slides.yaml")
        formatting_config = os.path.join(config_dir, "formatting.yaml")
        self.ppt_generator = PPTGenerator(
            template_path=self.template_path,
            slides_config=slides_config,
            formatting_config=formatting_config
        )
    
    def process_month(self, month_data_dir: str, output_path: str, use_raw_files: bool = False):
        """
        Process a month's data and generate PowerPoint deck.
        
        Args:
            month_data_dir: Directory containing month's Excel files
            output_path: Path to save generated PowerPoint file
            use_raw_files: If True, process raw files from Reports folder instead of Working File
        """
        print(f"Processing data from: {month_data_dir}")
        
        if use_raw_files:
            # Step 1: Process raw files from Reports folder
            print("Step 1: Processing raw files from Reports folder...")
            from src.working_file_generator import WorkingFileGenerator
            
            reports_dir = os.path.join(month_data_dir, "Reports")
            if not os.path.exists(reports_dir):
                print(f"  Warning: Reports folder not found at {reports_dir}")
                use_raw_files = False
        
        if use_raw_files:
            # Find raw files and map to Working File sheets
            raw_file_mappings = self._map_raw_files_to_sheets(reports_dir)
            
            # Generate Working File from raw files
            generator = WorkingFileGenerator()
            loaded_data = generator.generate_from_raw_files(raw_file_mappings)
            
            # Structure as: {"AIL LT Working file": {sheet_name: DataFrame}}
            # This matches what PPT generator expects
            working_file_data = {"AIL LT Working file": loaded_data}
            loaded_data = working_file_data
        else:
            # Step 1: Load Excel files (original approach)
            print("Step 1: Loading Excel files...")
            excel_files = self._find_excel_files(month_data_dir)
            loaded_data = {}
            
            for file_path in excel_files:
                file_name = os.path.basename(file_path)
                print(f"  Loading: {file_name}")
                try:
                    data = self.data_loader.load_excel(file_path)
                    # Use file name as key (without extension)
                    key = os.path.splitext(file_name)[0]
                    loaded_data[key] = data
                except Exception as e:
                    print(f"  Warning: Could not load {file_name}: {e}")
        
        # Step 2: Normalize data
        print("Step 2: Normalizing data...")
        normalized_data = {}
        for key, data in loaded_data.items():
            if isinstance(data, pd.DataFrame):
                normalized_data[key] = self.data_normalizer.normalize_data(data, preserve_names=True)
            elif isinstance(data, dict):
                # Multiple sheets
                normalized_data[key] = {
                    sheet: self.data_normalizer.normalize_data(df, preserve_names=True)
                    for sheet, df in data.items()
                }
            else:
                normalized_data[key] = data
        
        # Step 3: Apply transformations
        print("Step 3: Applying transformations...")
        transformed_data = {}
        for key, data in normalized_data.items():
            # Apply transformations based on configuration
            # For now, pass through - transformations will be applied per slide
            transformed_data[key] = data
        
        # Step 4: Apply business rules
        print("Step 4: Applying business rules...")
        rule_results = self.rules_engine.evaluate_all_rules(transformed_data)
        
        # Step 5: Generate PowerPoint
        print("Step 5: Generating PowerPoint deck...")
        self.ppt_generator.generate(transformed_data, output_path)
        
        print(f"\nPipeline completed successfully!")
        print(f"Output saved to: {output_path}")
    
    def _find_excel_files(self, directory: str) -> list:
        """Find all Excel files in a directory."""
        excel_files = []
        directory_path = Path(directory)
        
        # Look for .xlsx files
        excel_files.extend(directory_path.rglob("*.xlsx"))
        
        # Look for .xlsb files
        excel_files.extend(directory_path.rglob("*.xlsb"))
        
        # Filter out files in Reports subdirectories if "Working file" exists
        working_files = [f for f in excel_files if "Working file" in f.name or "Working file" in str(f)]
        if working_files:
            # Prefer working files
            excel_files = working_files + [f for f in excel_files if f not in working_files]
        
        return [str(f) for f in excel_files]
    
    def _map_raw_files_to_sheets(self, reports_dir: str) -> Dict[str, str]:
        """
        Map raw files to Working File sheet names.
        
        Args:
            reports_dir: Directory containing raw files
            
        Returns:
            Dictionary mapping sheet names to file paths
        """
        mappings = {}
        
        # Find all Excel files in Reports directory
        reports_path = Path(reports_dir)
        raw_files = list(reports_path.glob("*.xlsx")) + list(reports_path.glob("*.xlsb"))
        
        # Map based on filename patterns
        for file_path in raw_files:
            file_name = file_path.name.lower()
            
            if 'consented status' in file_name or 'consent' in file_name:
                mappings['consent'] = str(file_path)
            elif 'chronic missing' in file_name:
                mappings['Chronic & Overcalling'] = str(file_path)
            # Add more mappings as we process more files
            # elif 'input distribution' in file_name:
            #     mappings['INPUT DISTRIBUTION STATUS'] = str(file_path)
        
        return mappings
    
    def analyze_and_discover(self, excel_path: str, ppt_path: str, output_dir: str = "analysis"):
        """
        Analyze Excel and PPT files to discover rules and mappings.
        
        Args:
            excel_path: Path to Excel file
            ppt_path: Path to PowerPoint file
            output_dir: Directory to save analysis results
        """
        print(f"Analyzing Excel: {excel_path}")
        print(f"Analyzing PPT: {ppt_path}")
        
        os.makedirs(output_dir, exist_ok=True)
        
        # Import analysis modules
        from template_extractor import TemplateExtractor
        from excel_analyzer import ExcelAnalyzer
        from rule_discoverer import RuleDiscoverer
        
        # Extract template information
        print("\nExtracting template information...")
        template_extractor = TemplateExtractor(ppt_path)
        template_info = template_extractor.extract_all()
        template_output = os.path.join(output_dir, "template_info.json")
        template_extractor.save_template_info(template_output)
        print(f"  Saved to: {template_output}")
        
        # Analyze Excel file
        print("\nAnalyzing Excel file...")
        excel_analyzer = ExcelAnalyzer(excel_path)
        excel_info = excel_analyzer.analyze_all()
        excel_output = os.path.join(output_dir, "excel_info.json")
        excel_analyzer.save_analysis(excel_output)
        print(f"  Saved to: {excel_output}")
        
        # Discover rules
        print("\nDiscovering rules...")
        rule_discoverer = RuleDiscoverer(excel_path, ppt_path)
        rules = rule_discoverer.discover_all()
        rules_output = os.path.join(output_dir, "discovered_rules.json")
        rule_discoverer.save_rules(rules_output)
        print(f"  Saved to: {rules_output}")
        
        print(f"\nAnalysis complete! Results saved to: {output_dir}")


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Automated PowerPoint Deck Creation from Excel (Rule-Based)"
    )
    
    subparsers = parser.add_subparsers(dest="command", help="Command to execute")
    
    # Generate command
    generate_parser = subparsers.add_parser("generate", help="Generate PowerPoint deck")
    generate_parser.add_argument("month_dir", help="Directory containing month's Excel files")
    generate_parser.add_argument("output", help="Output PowerPoint file path")
    generate_parser.add_argument("--template", help="Path to PowerPoint template file")
    generate_parser.add_argument("--config-dir", default="config", help="Configuration directory")
    generate_parser.add_argument("--use-raw-files", action="store_true", 
                                help="Process raw files from Reports folder instead of Working File")
    
    # Analyze command
    analyze_parser = subparsers.add_parser("analyze", help="Analyze Excel and PPT files")
    analyze_parser.add_argument("excel_file", help="Path to Excel file")
    analyze_parser.add_argument("ppt_file", help="Path to PowerPoint file")
    analyze_parser.add_argument("--output-dir", default="analysis", help="Output directory for analysis")
    analyze_parser.add_argument("--config-dir", default="config", help="Configuration directory")
    
    args = parser.parse_args()
    
    if args.command == "generate":
        pipeline = PPTPipeline(
            config_dir=args.config_dir,
            template_path=args.template
        )
        pipeline.process_month(args.month_dir, args.output, use_raw_files=args.use_raw_files)
    
    elif args.command == "analyze":
        pipeline = PPTPipeline(config_dir=args.config_dir)
        pipeline.analyze_and_discover(args.excel_file, args.ppt_file, args.output_dir)
    
    else:
        parser.print_help()


if __name__ == "__main__":
    main()

