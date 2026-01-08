"""
PPT Generator
Main PowerPoint generation orchestrator.
"""

import os
from pptx import Presentation
from typing import Dict, List, Any, Optional
import yaml
try:
    from .ppt_builder import PPTBuilder
    from .ppt_formatter import PPTFormatter
except ImportError:
    from ppt_builder import PPTBuilder
    from ppt_formatter import PPTFormatter


class PPTGenerator:
    """Generates PowerPoint decks from processed data."""
    
    def __init__(self, template_path: Optional[str] = None,
                 slides_config: Optional[str] = None,
                 formatting_config: Optional[str] = None):
        """
        Initialize the PPT generator.
        
        Args:
            template_path: Path to PowerPoint template file
            slides_config: Path to slides configuration YAML file
            formatting_config: Path to formatting configuration YAML file
        """
        self.template_path = template_path
        self.slides_config = slides_config
        self.formatting_config = formatting_config
        
        # Load template
        if template_path and os.path.exists(template_path):
            self.presentation = Presentation(template_path)
        else:
            self.presentation = Presentation()
        
        # Load configurations
        self.slides_mapping = {}
        self.formatting_rules = {}
        
        if slides_config and os.path.exists(slides_config):
            self._load_slides_config()
        
        if formatting_config and os.path.exists(formatting_config):
            self._load_formatting_config()
        
        # Initialize formatter and builder
        self.formatter = PPTFormatter(self.formatting_rules)
        self.builder = PPTBuilder(self.presentation, self.formatter)
    
    def _load_slides_config(self):
        """Load slides configuration from YAML file."""
        with open(self.slides_config, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
            self.slides_mapping = config.get("slides", [])
    
    def _load_formatting_config(self):
        """Load formatting configuration from YAML file."""
        with open(self.formatting_config, 'r', encoding='utf-8') as f:
            self.formatting_rules = yaml.safe_load(f)
    
    def generate(self, data: Dict[str, Any], output_path: str):
        """
        Generate PowerPoint deck from data.
        
        Args:
            data: Dictionary mapping data source names to DataFrames
            output_path: Path to save the generated PowerPoint file
        """
        # Clear existing slides (whether using template or not)
        # We want to generate fresh slides, not add to existing ones
        while len(self.presentation.slides) > 0:
            try:
                rId = self.presentation.slides._sldIdLst[0].rId
                self.presentation.part.drop_rel(rId)
                del self.presentation.slides._sldIdLst[0]
            except (IndexError, AttributeError):
                break
        
        # Generate slides based on configuration
        for slide_config in self.slides_mapping:
            self._generate_slide(slide_config, data)
        
        # Ensure output directory exists
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        # Save presentation
        self.presentation.save(output_path)
        print(f"PowerPoint deck saved to: {output_path}")
    
    def _generate_slide(self, slide_config: Dict, data: Dict[str, Any]):
        """Generate a single slide based on configuration."""
        slide_number = slide_config.get("slide_number", 1)
        slide_type = slide_config.get("slide_type", "content")
        
        # Get slide layout
        layout = None
        if self.template_path:
            layout_name = slide_config.get("layout_name")
            if layout_name:
                for layout_option in self.presentation.slide_layouts:
                    if layout_option.name == layout_name:
                        layout = layout_option
                        break
        
        # Add slide
        slide = self.builder.add_slide(layout)
        
        # Populate slide based on type
        if slide_type == "title":
            self._generate_title_slide(slide, slide_config, data)
        elif slide_type == "content":
            self._generate_content_slide(slide, slide_config, data)
        elif slide_type == "table":
            self._generate_table_slide(slide, slide_config, data)
        elif slide_type == "bullet_list":
            self._generate_bullet_list_slide(slide, slide_config, data)
        else:
            # Generic slide generation
            self._generate_generic_slide(slide, slide_config, data)
    
    def _generate_title_slide(self, slide, slide_config: Dict, data: Dict[str, Any]):
        """Generate a title slide."""
        title = slide_config.get("title", "Title")
        subtitle = slide_config.get("subtitle", "")
        
        # Get title from data if specified
        if "title_data_source" in slide_config:
            title = self._get_text_from_data(data, slide_config["title_data_source"])
        
        # Get subtitle from data if specified
        if "subtitle_data_source" in slide_config:
            subtitle = self._get_text_from_data(data, slide_config["subtitle_data_source"])
        
        # Update title and subtitle placeholders
        for shape in slide.shapes:
            if shape.has_text_frame:
                if "Title" in shape.name or shape.name == "":
                    shape.text_frame.text = title
                elif "Subtitle" in shape.name:
                    shape.text_frame.text = subtitle
    
    def _generate_content_slide(self, slide, slide_config: Dict, data: Dict[str, Any]):
        """Generate a content slide."""
        title = slide_config.get("title", "")
        content_mappings = slide_config.get("content_mappings", [])
        
        # Set title
        if title:
            for shape in slide.shapes:
                if shape.has_text_frame and ("Title" in shape.name or shape.name == ""):
                    shape.text_frame.text = title
                    break
        
        # Populate content based on mappings
        self.builder.populate_slide_from_mapping(slide, data, {"shape_mappings": content_mappings})
    
    def _generate_table_slide(self, slide, slide_config: Dict, data: Dict[str, Any]):
        """Generate a table slide."""
        title = slide_config.get("title", "")
        table_mapping = slide_config.get("table_mapping", {})
        
        # Set title
        if title:
            for shape in slide.shapes:
                if shape.has_text_frame and ("Title" in shape.name or shape.name == ""):
                    shape.text_frame.text = title
                    break
        
        # Add or update table
        table_data = self.builder._get_table_data(data, table_mapping)
        if table_data is not None:
            # Find existing table or add new one
            table_shape = None
            for shape in slide.shapes:
                if shape.has_table:
                    table_shape = shape
                    break
            
            if table_shape:
                self.builder.update_table_data(slide, slide.shapes.index(table_shape), table_data, 
                                             table_mapping.get("formatting"))
            else:
                # Add new table
                self.builder.add_table(slide, table_data, 1, 2, 8, 4, 
                                     table_mapping.get("formatting"))
    
    def _generate_bullet_list_slide(self, slide, slide_config: Dict, data: Dict[str, Any]):
        """Generate a bullet list slide."""
        title = slide_config.get("title", "")
        items = slide_config.get("items", [])
        
        # Get items from data if specified
        if "items_data_source" in slide_config:
            items = self._get_list_from_data(data, slide_config["items_data_source"])
        
        # Set title
        if title:
            for shape in slide.shapes:
                if shape.has_text_frame and ("Title" in shape.name or shape.name == ""):
                    shape.text_frame.text = title
                    break
        
        # Add bullet list
        self.builder.add_bullet_list(slide, items, 1, 2, 8, 5, 
                                    slide_config.get("formatting"))
    
    def _generate_generic_slide(self, slide, slide_config: Dict, data: Dict[str, Any]):
        """Generate a generic slide based on mappings."""
        mappings = slide_config.get("mappings", {})
        self.builder.populate_slide_from_mapping(slide, data, mappings)
    
    def _get_text_from_data(self, data: Dict[str, Any], data_source_config: Dict) -> str:
        """Extract text value from data."""
        data_source = data_source_config.get("source")
        column = data_source_config.get("column")
        aggregate = data_source_config.get("aggregate", "first")
        
        if data_source in data:
            df = data[data_source]
            if isinstance(df, dict):
                df = list(df.values())[0]
            
            if isinstance(df, pd.DataFrame) and column in df.columns:
                if aggregate == "sum":
                    return str(df[column].sum())
                elif aggregate == "mean":
                    return str(df[column].mean())
                else:
                    return str(df[column].iloc[0])
        
        return data_source_config.get("default", "")
    
    def _get_list_from_data(self, data: Dict[str, Any], data_source_config: Dict) -> List[str]:
        """Extract list of items from data."""
        data_source = data_source_config.get("source")
        column = data_source_config.get("column")
        
        if data_source in data:
            df = data[data_source]
            if isinstance(df, dict):
                df = list(df.values())[0]
            
            if isinstance(df, pd.DataFrame) and column in df.columns:
                return df[column].astype(str).tolist()
        
        return data_source_config.get("default", [])


if __name__ == "__main__":
    # Example usage
    import sys
    
    if len(sys.argv) < 3:
        print("Usage: python ppt_generator.py <template_path> <output_path> [slides_config] [formatting_config]")
        sys.exit(1)
    
    template_path = sys.argv[1]
    output_path = sys.argv[2]
    slides_config = sys.argv[3] if len(sys.argv) > 3 else None
    formatting_config = sys.argv[4] if len(sys.argv) > 4 else None
    
    generator = PPTGenerator(template_path, slides_config, formatting_config)
    
    # Create sample data
    import pandas as pd
    sample_data = {
        "main": pd.DataFrame({
            "Category": ["A", "B", "C"],
            "Value": [10, 20, 30]
        })
    }
    
    generator.generate(sample_data, output_path)

