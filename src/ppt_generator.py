"""
PPT Generator
Main PowerPoint generation orchestrator.
"""

import os
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
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
            print(f"DEBUG: Loaded {len(self.slides_mapping)} slide configurations")
            for idx, slide_config in enumerate(self.slides_mapping, start=1):
                print(f"DEBUG: Slide {idx}: type={slide_config.get('slide_type')}, title='{slide_config.get('title')}', chart_enabled={slide_config.get('chart', {}).get('enabled', False)}")
    
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
        print(f"DEBUG: Generating {len(self.slides_mapping)} slides from configuration")
        for idx, slide_config in enumerate(self.slides_mapping, start=1):
            try:
                print(f"DEBUG: Processing slide {idx} of {len(self.slides_mapping)}")
                self._generate_slide(slide_config, data)
                print(f"DEBUG: Successfully completed slide {idx}")
            except Exception as e:
                import traceback
                error_msg = f"Failed to generate slide {idx}: {str(e)}\n{traceback.format_exc()}"
                print(f"ERROR: {error_msg}")
                # Continue with next slide instead of stopping
                continue
        
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
        
        print(f"DEBUG: Generating slide {slide_number} of type '{slide_type}'")
        
        try:
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
            print(f"DEBUG: Slide {slide_number} created successfully")
            
            # Populate slide based on type
            if slide_type == "title":
                self._generate_title_slide(slide, slide_config, data)
            elif slide_type == "content":
                self._generate_content_slide(slide, slide_config, data, slide_number)
            elif slide_type == "table":
                self._generate_table_slide(slide, slide_config, data, slide_number)
            elif slide_type == "bullet_list":
                self._generate_bullet_list_slide(slide, slide_config, data)
            else:
                # Generic slide generation
                self._generate_generic_slide(slide, slide_config, data)
            
            print(f"DEBUG: Slide {slide_number} populated successfully")
            
        except Exception as e:
            import traceback
            error_msg = f"Error generating slide {slide_number}: {str(e)}\n{traceback.format_exc()}"
            print(f"ERROR: {error_msg}")
            # Still create the slide even if there's an error, with error message
            try:
                if 'slide' not in locals():
                    slide = self.builder.add_slide(None)
                self.builder.add_text_box(
                    slide, f"Error: Could not generate slide {slide_number}\n{str(e)}",
                    left=1, top=3, width=8, height=2,
                    formatting={"font_size": 14, "font_color": "#CC0000", "font_name": "Calibri"}
                )
            except:
                pass  # If even error message fails, continue
    
    def _generate_title_slide(self, slide, slide_config: Dict, data: Dict[str, Any]):
        """Generate a title slide."""
        title = slide_config.get("title", "Title")
        subtitle = slide_config.get("subtitle", "")
        
        # Get formatting from config
        title_formatting = slide_config.get("title_formatting", {})
        subtitle_formatting = slide_config.get("subtitle_formatting", {})
        
        # Default title formatting with colors and sizes
        default_title_formatting = {
            "font_size": title_formatting.get("font_size", 48),
            "bold": title_formatting.get("bold", True),
            "font_color": title_formatting.get("font_color", "#003B55"),  # Dark blue
            "alignment": "center",
            "font_name": title_formatting.get("font_name", "Calibri")
        }
        
        # Default subtitle formatting
        default_subtitle_formatting = {
            "font_size": subtitle_formatting.get("font_size", 28),
            "bold": subtitle_formatting.get("bold", False),
            "font_color": subtitle_formatting.get("font_color", "#666666"),  # Gray
            "alignment": "center",
            "font_name": subtitle_formatting.get("font_name", "Calibri")
        }
        
        # Merge with provided formatting
        title_formatting = {**default_title_formatting, **title_formatting}
        subtitle_formatting = {**default_subtitle_formatting, **subtitle_formatting}
        
        # Get title from data if specified
        if "title_data_source" in slide_config:
            title = self._get_text_from_data(data, slide_config["title_data_source"])
        
        # Get subtitle from data if specified
        if "subtitle_data_source" in slide_config:
            subtitle = self._get_text_from_data(data, slide_config["subtitle_data_source"])
        
        # Try to update existing placeholders first
        title_found = False
        subtitle_found = False
        for shape in slide.shapes:
            if shape.has_text_frame:
                shape_name = shape.name.lower()
                if "title" in shape_name or (shape_name == "" and not title_found):
                    shape.text_frame.text = title
                    # Format title with colors
                    for paragraph in shape.text_frame.paragraphs:
                        paragraph.alignment = PP_ALIGN.CENTER
                        for run in paragraph.runs:
                            run.font.size = Pt(title_formatting["font_size"])
                            run.font.bold = title_formatting["bold"]
                            # Apply color
                            if "font_color" in title_formatting:
                                color_str = title_formatting["font_color"]
                                if color_str.startswith("#"):
                                    color_str = color_str[1:]
                                r, g, b = int(color_str[0:2], 16), int(color_str[2:4], 16), int(color_str[4:6], 16)
                                from pptx.dml.color import RGBColor
                                run.font.color.rgb = RGBColor(r, g, b)
                    title_found = True
                elif "subtitle" in shape_name:
                    shape.text_frame.text = subtitle
                    # Format subtitle with colors
                    for paragraph in shape.text_frame.paragraphs:
                        paragraph.alignment = PP_ALIGN.CENTER
                        for run in paragraph.runs:
                            run.font.size = Pt(subtitle_formatting["font_size"])
                            run.font.bold = subtitle_formatting["bold"]
                            # Apply color
                            if "font_color" in subtitle_formatting:
                                color_str = subtitle_formatting["font_color"]
                                if color_str.startswith("#"):
                                    color_str = color_str[1:]
                                r, g, b = int(color_str[0:2], 16), int(color_str[2:4], 16), int(color_str[4:6], 16)
                                from pptx.dml.color import RGBColor
                                run.font.color.rgb = RGBColor(r, g, b)
                    subtitle_found = True
        
        # If no placeholders found, add text boxes manually
        if not title_found and title:
            self.builder.add_text_box(
                slide, title, 
                left=1, top=3, width=8, height=1,
                formatting=title_formatting
            )
        if not subtitle_found and subtitle:
            self.builder.add_text_box(
                slide, subtitle,
                left=1, top=4.5, width=8, height=0.8,
                formatting=subtitle_formatting
            )
    
    def _generate_content_slide(self, slide, slide_config: Dict, data: Dict[str, Any], slide_number: int = 1):
        """Generate a content slide."""
        title = slide_config.get("title", "")
        subtitle = slide_config.get("subtitle", "")
        content_mappings = slide_config.get("content_mappings", [])
        chart_config = slide_config.get("chart", None)
        
        # Get formatting from config
        title_formatting = slide_config.get("title_formatting", {})
        subtitle_formatting = slide_config.get("subtitle_formatting", {})
        
        # Default title formatting with colors and sizes
        default_title_formatting = {
            "font_size": title_formatting.get("font_size", 36),
            "bold": title_formatting.get("bold", True),
            "font_color": title_formatting.get("font_color", "#003B55"),  # Dark blue
            "font_name": title_formatting.get("font_name", "Calibri")
        }
        
        # Default subtitle formatting
        default_subtitle_formatting = {
            "font_size": subtitle_formatting.get("font_size", 18),
            "bold": subtitle_formatting.get("bold", False),
            "font_color": subtitle_formatting.get("font_color", "#666666"),  # Gray
            "font_name": subtitle_formatting.get("font_name", "Calibri")
        }
        
        # Merge with provided formatting
        title_formatting = {**default_title_formatting, **title_formatting}
        subtitle_formatting = {**default_subtitle_formatting, **subtitle_formatting}
        
        # Try to set title in existing placeholder
        title_found = False
        if title:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    shape_name = shape.name.lower()
                    if "title" in shape_name or (shape_name == "" and not title_found):
                        shape.text_frame.text = title
                        # Format title with colors
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.font.size = Pt(title_formatting["font_size"])
                                run.font.bold = title_formatting["bold"]
                                # Apply color
                                if "font_color" in title_formatting:
                                    color_str = title_formatting["font_color"]
                                    if color_str.startswith("#"):
                                        color_str = color_str[1:]
                                    r, g, b = int(color_str[0:2], 16), int(color_str[2:4], 16), int(color_str[4:6], 16)
                                    from pptx.dml.color import RGBColor
                                    run.font.color.rgb = RGBColor(r, g, b)
                        title_found = True
                        break
        
        # If no placeholder found, add title as text box
        if not title_found and title:
            self.builder.add_text_box(
                slide, title,
                left=0.5, top=0.5, width=9, height=1,
                formatting=title_formatting
            )
        
        # Add subtitle if present
        if subtitle:
            self.builder.add_text_box(
                slide, subtitle,
                left=0.5, top=1.5, width=9, height=0.6,
                formatting=subtitle_formatting
            )
        
        # Track if chart was successfully added
        chart_success = False
        
        # Add chart if configured
        top_offset = 2.5 if subtitle else 1.8
        if chart_config and chart_config.get("enabled", False):
            print(f"DEBUG: Chart is enabled for slide {slide_number}")
            # Build proper chart mapping with all necessary columns (x + y)
            x_column = chart_config.get("x_column")
            y_columns = chart_config.get("y_columns", [])
            
            if x_column and y_columns:
                # Create a mapping for chart data that includes all needed columns
                chart_mapping = {
                    "data_source": chart_config.get("data_source"),
                    "sheet": chart_config.get("sheet"),
                    "header_row": chart_config.get("header_row", 0),
                    "columns": [x_column] + (y_columns if isinstance(y_columns, list) else [y_columns])
                }
                
                chart_data = self.builder._get_table_data(data, chart_mapping, return_column_mapping=True)
                if chart_data is not None:
                    if isinstance(chart_data, tuple):
                        # Returned (DataFrame, column_mapping)
                        chart_data_df, column_mapping = chart_data
                        chart_data = chart_data_df
                    else:
                        # Just DataFrame, create mapping from column order
                        column_mapping = {}
                        actual_columns = list(chart_data.columns)
                        requested_columns = [x_column] + (y_columns if isinstance(y_columns, list) else [y_columns])
                        for i, req_col in enumerate(requested_columns):
                            if i < len(actual_columns):
                                column_mapping[req_col] = actual_columns[i]
                    
                    if len(chart_data) > 0 and len(chart_data.columns) > 0:
                        # Use column mapping to find actual X and Y column names
                        actual_x_column = column_mapping.get(x_column)
                        if not actual_x_column:
                            # Fallback: use first column if mapping not available
                            actual_x_column = list(chart_data.columns)[0]
                        
                        # Map Y columns
                        actual_y_columns = []
                        y_cols_list = y_columns if isinstance(y_columns, list) else [y_columns]
                        for y_col in y_cols_list:
                            actual_y = column_mapping.get(y_col)
                            if actual_y and actual_y != actual_x_column:
                                actual_y_columns.append(actual_y)
                            elif actual_y == actual_x_column:
                                # Skip if it's the same as X column
                                pass
                        
                        # If no Y columns found, use all columns except X
                        if not actual_y_columns:
                            actual_y_columns = [col for col in chart_data.columns if col != actual_x_column]
                        
                        print(f"DEBUG: Chart - Original x_column: '{x_column}', Actual: '{actual_x_column}'")
                        print(f"DEBUG: Chart - Original y_columns: {y_columns}, Actual: {actual_y_columns}")
                        print(f"DEBUG: Chart - Available columns: {list(chart_data.columns)}")
                        
                        if actual_x_column and actual_y_columns:
                            try:
                                self.builder.add_chart(
                                    slide, chart_data,
                                    chart_type=chart_config.get("type", "column"),
                                    left=1, top=top_offset, width=8, height=4,
                                    x_column=actual_x_column,  # Use actual column name from DataFrame
                                    y_columns=actual_y_columns,  # Use actual column names from DataFrame
                                    title=chart_config.get("title", ""),
                                    formatting=chart_config.get("formatting", {})
                                )
                                chart_success = True
                                print(f"DEBUG: Chart successfully added to slide {slide_number}")
                            except Exception as e:
                                import traceback
                                error_msg = str(e)
                                print(f"ERROR: Could not add chart: {error_msg}")
                                print(f"Traceback: {traceback.format_exc()}")
                                chart_success = False
                        else:
                            print(f"Warning: Missing X or Y columns. X: {actual_x_column}, Y: {actual_y_columns}")
                            chart_success = False
                    else:
                        print(f"Warning: Chart data is empty or has no columns")
                        chart_success = False
                else:
                    print(f"Warning: No chart data available. x_column: {x_column}, y_columns: {y_columns}")
                    chart_success = False
            else:
                print(f"Warning: Chart enabled but missing x_column or y_columns. x: {x_column}, y: {y_columns}")
                chart_success = False
        
        # IMPORTANT: Don't add content if chart is enabled (even if chart failed)
        # User wants: if chart is enabled, show chart only, no content mappings
        if chart_config and chart_config.get("enabled", False):
            if chart_success:
                print(f"DEBUG: Chart successfully added to slide {slide_number}, skipping content mappings")
            else:
                print(f"DEBUG: Chart was enabled but failed, skipping content mappings as requested (chart-only slide)")
            return  # Exit early - chart only slide
        
        # Populate content based on mappings
        if content_mappings:
            self.builder.populate_slide_from_mapping(slide, data, {"shape_mappings": content_mappings})
        elif not title and not subtitle and not (chart_config and chart_config.get("enabled", False)):
            # If no content at all, add a placeholder message
            self.builder.add_text_box(
                slide, "No content configured for this slide",
                left=2, top=3, width=6, height=1,
                formatting={
                    "font_size": 18,
                    "font_color": "#999999",
                    "italic": True,
                    "alignment": "center",
                    "font_name": "Calibri"
                }
            )
    
    def _generate_table_slide(self, slide, slide_config: Dict, data: Dict[str, Any], slide_number: int = 1):
        """Generate a table slide."""
        title = slide_config.get("title", "")
        subtitle = slide_config.get("subtitle", "")
        table_mapping = slide_config.get("table_mapping", {})
        chart_config = slide_config.get("chart", None)
        
        # Get formatting from config
        title_formatting = slide_config.get("title_formatting", {})
        subtitle_formatting = slide_config.get("subtitle_formatting", {})
        
        # Default title formatting with colors and sizes
        default_title_formatting = {
            "font_size": title_formatting.get("font_size", 36),
            "bold": title_formatting.get("bold", True),
            "font_color": title_formatting.get("font_color", "#003B55"),  # Dark blue
            "font_name": title_formatting.get("font_name", "Calibri")
        }
        
        # Default subtitle formatting
        default_subtitle_formatting = {
            "font_size": subtitle_formatting.get("font_size", 18),
            "bold": subtitle_formatting.get("bold", False),
            "font_color": subtitle_formatting.get("font_color", "#666666"),  # Gray
            "font_name": subtitle_formatting.get("font_name", "Calibri")
        }
        
        # Merge with provided formatting
        title_formatting = {**default_title_formatting, **title_formatting}
        subtitle_formatting = {**default_subtitle_formatting, **subtitle_formatting}
        
        # Try to set title in existing placeholder
        title_found = False
        if title:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    shape_name = shape.name.lower()
                    if "title" in shape_name or (shape_name == "" and not title_found):
                        shape.text_frame.text = title
                        # Format title with colors
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.font.size = Pt(title_formatting["font_size"])
                                run.font.bold = title_formatting["bold"]
                                # Apply color
                                if "font_color" in title_formatting:
                                    color_str = title_formatting["font_color"]
                                    if color_str.startswith("#"):
                                        color_str = color_str[1:]
                                    r, g, b = int(color_str[0:2], 16), int(color_str[2:4], 16), int(color_str[4:6], 16)
                                    from pptx.dml.color import RGBColor
                                    run.font.color.rgb = RGBColor(r, g, b)
                        title_found = True
                        break
        
        # Calculate table position based on whether we have title/subtitle/chart
        top_offset = 0.5
        if title:
            top_offset = 1.5
        if subtitle:
            top_offset = 2.2
        
        # If no placeholder found, add title as text box
        if not title_found and title:
            self.builder.add_text_box(
                slide, title,
                left=0.5, top=0.5, width=9, height=0.8,
                formatting=title_formatting
            )
            top_offset = 1.5
        
        # Add subtitle if present
        if subtitle:
            self.builder.add_text_box(
                slide, subtitle,
                left=0.5, top=1.5, width=9, height=0.5,
                formatting=subtitle_formatting
            )
            top_offset = 2.2
        
        # Track if chart was successfully added (initialize to False)
        chart_success = False
        
        # Add chart if configured (before table)
        if chart_config and chart_config.get("enabled", False):
            print(f"DEBUG: Chart is enabled for this slide")
            # Build proper chart mapping with all necessary columns (x + y)
            x_column = chart_config.get("x_column")
            y_columns = chart_config.get("y_columns", [])
            
            if x_column and y_columns:
                print(f"DEBUG: Chart config - x_column: '{x_column}', y_columns: {y_columns}")
                # Create a mapping for chart data that includes all needed columns
                chart_mapping = {
                    "data_source": chart_config.get("data_source"),
                    "sheet": chart_config.get("sheet"),
                    "header_row": chart_config.get("header_row", 0),
                    "columns": [x_column] + (y_columns if isinstance(y_columns, list) else [y_columns])
                }
                
                print(f"DEBUG: Chart mapping - data_source: {chart_mapping['data_source']}, sheet: {chart_mapping['sheet']}")
                chart_data = self.builder._get_table_data(data, chart_mapping, return_column_mapping=True)
                if chart_data is not None:
                    if isinstance(chart_data, tuple):
                        # Returned (DataFrame, column_mapping)
                        chart_data_df, column_mapping = chart_data
                        chart_data = chart_data_df
                    else:
                        # Just DataFrame, create mapping from column order
                        column_mapping = {}
                        actual_columns = list(chart_data.columns)
                        requested_columns = [x_column] + (y_columns if isinstance(y_columns, list) else [y_columns])
                        for i, req_col in enumerate(requested_columns):
                            if i < len(actual_columns):
                                column_mapping[req_col] = actual_columns[i]
                    
                    if len(chart_data) > 0 and len(chart_data.columns) > 0:
                        print(f"DEBUG: Chart data retrieved: {len(chart_data)} rows, {len(chart_data.columns)} columns")
                        # Use column mapping to find actual X and Y column names
                        actual_x_column = column_mapping.get(x_column)
                        if not actual_x_column:
                            # Fallback: use first column if mapping not available
                            actual_x_column = list(chart_data.columns)[0]
                        
                        # Map Y columns
                        actual_y_columns = []
                        y_cols_list = y_columns if isinstance(y_columns, list) else [y_columns]
                        for y_col in y_cols_list:
                            actual_y = column_mapping.get(y_col)
                            if actual_y and actual_y != actual_x_column:
                                actual_y_columns.append(actual_y)
                            elif actual_y == actual_x_column:
                                # Skip if it's the same as X column
                                pass
                        
                        # If no Y columns found, use all columns except X
                        if not actual_y_columns:
                            actual_y_columns = [col for col in chart_data.columns if col != actual_x_column]
                        
                        print(f"DEBUG: Chart - Original x_column: '{x_column}', Actual: '{actual_x_column}'")
                        print(f"DEBUG: Chart - Original y_columns: {y_columns}, Actual: {actual_y_columns}")
                        print(f"DEBUG: Chart - Available columns: {list(chart_data.columns)}")
                        
                        if actual_x_column and actual_y_columns:
                            try:
                                self.builder.add_chart(
                                    slide, chart_data,
                                    chart_type=chart_config.get("type", "column"),
                                    left=1, top=top_offset, width=8, height=3.5,
                                    x_column=actual_x_column,  # Use actual column name from DataFrame
                                    y_columns=actual_y_columns,  # Use actual column names from DataFrame
                                    title=chart_config.get("title", ""),
                                    formatting=chart_config.get("formatting", {})
                                )
                                chart_success = True
                                print(f"DEBUG: Chart successfully added to slide {slide_number}")
                            except Exception as e:
                                import traceback
                                error_msg = str(e)
                                print(f"ERROR: Could not add chart: {error_msg}")
                                print(f"Traceback: {traceback.format_exc()}")
                                chart_success = False
                        else:
                            print(f"Warning: Missing X or Y columns. X: {actual_x_column}, Y: {actual_y_columns}")
                            chart_success = False
                    else:
                        print(f"Warning: Chart data is empty or has no columns")
                        chart_success = False
                else:
                    print(f"Warning: No chart data available. x_column: {x_column}, y_columns: {y_columns}")
                    chart_success = False
            else:
                print(f"Warning: Chart enabled but missing x_column or y_columns. x: {x_column}, y: {y_columns}")
                chart_success = False
        
        # IMPORTANT: Don't add table if chart is enabled (even if chart failed)
        # User wants: if chart is enabled, show chart only, no table
        if chart_config and chart_config.get("enabled", False):
            if chart_success:
                print(f"DEBUG: Chart successfully added to slide {slide_number}, skipping table generation")
            else:
                print(f"DEBUG: Chart was enabled but failed, skipping table generation as requested (chart-only slide)")
                # Show error message if chart failed
                try:
                    self.builder.add_text_box(
                        slide, "Chart configuration error: Could not generate chart",
                        left=2, top=top_offset, width=6, height=1,
                        formatting={
                            "font_size": 16,
                            "font_color": "#CC0000",
                            "italic": True,
                            "alignment": "center",
                            "font_name": "Calibri"
                        }
                    )
                except:
                    pass  # If error message fails, continue
            return  # Exit early, don't add table - chart only slide
        
        # Get table data - only if chart is not enabled or chart failed
        table_data = None
        if table_mapping and table_mapping.get("data_source"):
            table_data = self.builder._get_table_data(data, table_mapping, return_column_mapping=False)
            
            # Handle tuple return if it happens
            if isinstance(table_data, tuple):
                table_data, _ = table_data
        
        # Debug output
        if table_data is not None:
            print(f"DEBUG: Table data retrieved: {len(table_data)} rows, columns: {list(table_data.columns)}")
        else:
            print(f"DEBUG: No table data retrieved for mapping: {table_mapping.get('data_source', 'N/A')} / {table_mapping.get('sheet', 'N/A')}")
        
        if table_data is not None and len(table_data) > 0:
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
                # Calculate table dimensions
                num_rows = len(table_data)
                num_cols = len(table_data.columns)
                
                # Adjust table size based on content
                table_width = min(9, max(6, num_cols * 1.2))  # Between 6 and 9 inches
                table_height = min(4.5, max(2, (num_rows + 1) * 0.4))  # Between 2 and 4.5 inches
                table_left = (10 - table_width) / 2  # Center the table
                
                # Ensure table doesn't go below slide
                max_top = 7.5 - table_height
                table_top = min(top_offset, max_top)
                
                # Add new table with better positioning
                self.builder.add_table(
                    slide, table_data, 
                    left=table_left, top=table_top, 
                    width=table_width, height=table_height,
                    formatting=table_mapping.get("formatting")
                )
        else:
            # No data available - add informative message
            message = "No data available for this table"
            if table_mapping.get("data_source"):
                message = f"No data found in {table_mapping.get('data_source')}"
            self.builder.add_text_box(
                slide, message,
                left=2, top=3.5, width=6, height=1,
                formatting={
                    "font_size": 18,
                    "font_color": "#999999",
                    "italic": True,
                    "alignment": "center",
                    "font_name": "Calibri"
                }
            )
    
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

