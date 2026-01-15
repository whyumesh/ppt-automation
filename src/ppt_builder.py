"""
PPT Builder
Slide building utilities for creating PowerPoint slides.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from typing import Dict, List, Any, Optional
import pandas as pd
try:
    from .ppt_formatter import PPTFormatter
except ImportError:
    from ppt_formatter import PPTFormatter


class PPTBuilder:
    """Builds PowerPoint slides from data."""
    
    def __init__(self, presentation: Presentation, formatter: Optional[PPTFormatter] = None):
        """
        Initialize the PPT builder.
        
        Args:
            presentation: PowerPoint presentation object
            formatter: Optional PPTFormatter instance
        """
        self.presentation = presentation
        self.formatter = formatter or PPTFormatter()
    
    def add_slide(self, layout=None) -> Any:
        """
        Add a new slide to the presentation.
        
        Args:
            layout: Optional slide layout to use
        
        Returns:
            New slide object
        """
        if layout:
            slide = self.presentation.slides.add_slide(layout)
        else:
            slide = self.presentation.slides.add_slide(self.presentation.slide_layouts[0])
        
        return slide
    
    def add_text_box(self, slide, text: str, left: float, top: float,
                    width: float, height: float, formatting: Optional[Dict] = None) -> Any:
        """
        Add a text box to a slide.
        
        Args:
            slide: Slide object
            text: Text content
            left: Left position in inches
            top: Top position in inches
            width: Width in inches
            height: Height in inches
            formatting: Optional formatting dictionary
        
        Returns:
            Shape object
        """
        left_inches = Inches(left)
        top_inches = Inches(top)
        width_inches = Inches(width)
        height_inches = Inches(height)
        
        text_box = slide.shapes.add_textbox(left_inches, top_inches, width_inches, height_inches)
        text_frame = text_box.text_frame
        text_frame.text = text
        text_frame.word_wrap = True
        text_frame.margin_bottom = Inches(0.05)
        text_frame.margin_top = Inches(0.05)
        text_frame.margin_left = Inches(0.1)
        text_frame.margin_right = Inches(0.1)
        
        # Apply default formatting if none specified
        if formatting is None:
            formatting = {
                "font_size": 18,
                "font_name": "Calibri",
                "alignment": "left"
            }
        
        if formatting:
            self.formatter.format_text_box(text_frame, formatting)
        
        return text_box
    
    def add_table(self, slide, data: pd.DataFrame, left: float, top: float,
                 width: float, height: float, formatting: Optional[Dict] = None) -> Any:
        """
        Add a table to a slide.
        
        Args:
            slide: Slide object
            data: DataFrame containing table data
            left: Left position in inches
            top: Top position in inches
            width: Width in inches
            height: Height in inches
            formatting: Optional formatting dictionary
        
        Returns:
            Shape object
        """
        # Handle empty DataFrame - create table with headers and "No data" message
        if data is None:
            # Create empty DataFrame with default structure
            data = pd.DataFrame({"Column": ["No data available"]})
        
        if len(data) == 0:
            # Create a single row with "No data available" message
            if len(data.columns) == 0:
                # No columns at all - create a default column
                data = pd.DataFrame({"Data": ["No data available"]})
            else:
                # Has columns but no rows - add one row with message
                no_data_row = {col: "No data available" for col in data.columns}
                data = pd.DataFrame([no_data_row])
        
        rows = len(data) + 1  # +1 for header
        cols = len(data.columns)
        
        left_inches = Inches(left)
        top_inches = Inches(top)
        width_inches = Inches(width)
        height_inches = Inches(height)
        
        table_shape = slide.shapes.add_table(rows, cols, left_inches, top_inches, width_inches, height_inches)
        table = table_shape.table
        
        # Add table borders for clarity
        # Note: Border setting may not work for all cell types, so we wrap in try-except
        try:
            # Set borders for all cells - light gray borders
            for row in table.rows:
                for cell in row.cells:
                    try:
                        # Try to set borders - not all cells support this
                        if hasattr(cell, 'border_top'):
                            cell.border_top.color.rgb = RGBColor(200, 200, 200)
                            cell.border_top.width = Pt(0.5)
                        if hasattr(cell, 'border_bottom'):
                            cell.border_bottom.color.rgb = RGBColor(200, 200, 200)
                            cell.border_bottom.width = Pt(0.5)
                        if hasattr(cell, 'border_left'):
                            cell.border_left.color.rgb = RGBColor(200, 200, 200)
                            cell.border_left.width = Pt(0.5)
                        if hasattr(cell, 'border_right'):
                            cell.border_right.color.rgb = RGBColor(200, 200, 200)
                            cell.border_right.width = Pt(0.5)
                    except (AttributeError, TypeError) as e:
                        # Some cells don't support border attributes - skip them
                        pass
        except Exception as e:
            # Border setting failed - not critical, continue without borders
            print(f"WARNING: Could not set table borders: {e}")
            pass
        
        # Distribute column widths evenly for better appearance
        # Ensure all columns are visible and properly sized - must fit within table width
        if cols > 0:
            # Calculate width per column to fit exactly within table width
            col_width = width / cols
            
            # Adaptive minimum column width based on number of columns
            if cols <= 4:
                min_col_width_inches = 0.8
            elif cols <= 7:
                min_col_width_inches = 0.6
            elif cols <= 10:
                min_col_width_inches = 0.5
            else:
                min_col_width_inches = 0.4
            
            # If calculated width is too small, use minimum (but this will make table wider)
            # Instead, enforce that columns fit within the provided table width
            # All columns should use equal width that fits exactly
            col_width = width / cols  # Force equal distribution
            
            # Set column widths - must sum to table width
            total_set_width = 0.0
            for col_idx in range(cols):
                table.columns[col_idx].width = Inches(col_width)
                total_set_width += col_width
            
            # Verify total width matches
            if abs(total_set_width - width) > 0.01:
                print(f"WARNING: Column width sum ({total_set_width:.2f}) doesn't match table width ({width:.2f})")
            
            print(f"DEBUG: Distributed {cols} columns with width {col_width:.2f} inches each (total: {total_set_width:.2f}, table width: {width:.2f})")
        
        # Merge formatting with defaults
        if formatting is None:
            formatting = {}
        
        # Adaptive font sizes based on number of columns
        # More columns = smaller font for better fit
        if cols <= 4:
            header_font_size = 12
            data_font_size = 11
        elif cols <= 7:
            header_font_size = 11
            data_font_size = 10
        elif cols <= 10:
            header_font_size = 10
            data_font_size = 9
        else:
            header_font_size = 9
            data_font_size = 8
        
        # Default header formatting - use title color #004E6F
        default_header_formatting = {
            "font_size": header_font_size,
            "bold": True,
            "fill_color": "#004E6F",  # Title color shade
            "font_color": "#FFFFFF"   # White text
        }
        
        # Default data formatting - white background, black text
        default_data_formatting = {
            "font_size": data_font_size,
            "bold": False,
            "font_color": "#000000",  # Black text for readability
            "fill_color": "#FFFFFF",  # White background
            "alignment": "left"  # Left align data cells
        }
        
        header_formatting = formatting.get("header_formatting", default_header_formatting)
        data_formatting = formatting.get("data_formatting", default_data_formatting)
        
        # Ensure header formatting has white text
        if "fill_color" in header_formatting:
            header_formatting["font_color"] = "#FFFFFF"
        
        # Populate header row
        for col_idx, col_name in enumerate(data.columns):
            cell = table.cell(0, col_idx)
            # Clean column name for display - handle NaN/None
            import pandas as pd
            if pd.isna(col_name) or col_name is None:
                cell.text = ""
            else:
                cell.text = str(col_name).strip()
            # Enable text wrapping for headers
            cell.text_frame.word_wrap = True
            cell.text_frame.auto_size = None  # Disable auto-size for better control
            self.formatter.format_table_cell(cell, header_formatting)
            # Left align header text (changed from center)
            cell.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
        
        # Populate data rows with number formatting - no alternating colors, use white background
        number_formatting = formatting.get("number_formatting", {})
        
        for row_idx, (_, row_data) in enumerate(data.iterrows(), start=1):
            # Use white background for all rows (no alternating colors)
            row_formatting = data_formatting.copy()
            row_formatting["fill_color"] = "#FFFFFF"  # Always white background
            
            for col_idx, value in enumerate(row_data):
                cell = table.cell(row_idx, col_idx)
                
                # Clean NaN/None values first
                import pandas as pd
                import numpy as np
                if pd.isna(value) or value is None or str(value).lower() in ['nan', 'none', 'nat']:
                    cell.text = ""
                else:
                    # Apply number formatting if specified
                    col_name = data.columns[col_idx]
                    if col_name in number_formatting:
                        format_type = number_formatting[col_name]
                        try:
                            value_float = float(value)
                            if format_type == "percentage":
                                # Check if value is already a percentage (>= 1) or decimal (< 1)
                                # Also check if it's a very large number (likely already multiplied)
                                if value_float < 1:
                                    # It's a decimal (0-1), convert to percentage
                                    cell.text = f"{value_float * 100:.1f}%"
                                elif value_float > 100:
                                    # It's been multiplied already, divide by 100
                                    cell.text = f"{value_float / 100:.1f}%"
                                else:
                                    # It's already a percentage (1-100), just add % sign
                                    cell.text = f"{value_float:.1f}%"
                            elif format_type == "percentage_decimal":
                                # Value is already decimal (0-1), convert to percentage
                                cell.text = f"{value_float * 100:.1f}%"
                            elif format_type == "number":
                                # Remove decimal if it's .0
                                if value_float == int(value_float):
                                    cell.text = f"{int(value_float):,}"
                                else:
                                    cell.text = f"{value_float:,.2f}"
                            elif format_type == "integer":
                                cell.text = f"{int(value_float):,}"
                            elif format_type == "currency":
                                cell.text = f"${value_float:,.2f}"
                            else:
                                cell.text = str(value).strip()
                        except (ValueError, TypeError):
                            cell.text = str(value).strip() if value else ""
                    else:
                        # Default: format numbers nicely, keep text as-is
                        try:
                            value_float = float(value)
                            # If it's a whole number, remove .0
                            if value_float == int(value_float):
                                cell.text = str(int(value_float))
                            else:
                                cell.text = str(value).strip()
                        except (ValueError, TypeError):
                            cell.text = str(value).strip() if value else ""
                
                # Apply formatting
                cell_formatting = row_formatting.copy()
                
                # Apply conditional formatting if specified (overrides row color)
                # Use title color shade for conditional formatting
                if "conditional_colors" in formatting:
                    for cond in formatting["conditional_colors"]:
                        if cond.get("column") == col_name:
                            try:
                                value_float = float(value)
                                condition = cond.get("condition", "<")
                                threshold = cond.get("threshold", 0)
                                if (condition == "<" and value_float < threshold) or \
                                   (condition == ">" and value_float > threshold) or \
                                   (condition == "==" and value_float == threshold):
                                    # Use title color shade instead of red
                                    cell_formatting["font_color"] = cond.get("color", "#004E6F")
                            except (ValueError, TypeError):
                                pass
                
                # Enable text wrapping for data cells
                cell.text_frame.word_wrap = True
                cell.text_frame.auto_size = None  # Disable auto-size for better control
                # Set vertical alignment to middle for better appearance
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                
                self.formatter.format_table_cell(cell, cell_formatting)
                # Left align data
                cell.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
        
        # Apply additional formatting
        if formatting:
            self.formatter.format_table(table, formatting)
        
        # Ensure proper row heights for readability - adaptive based on columns
        # More columns = slightly smaller row heights
        if cols <= 4:
            header_row_height = 0.5
            data_row_height = 0.45
        elif cols <= 7:
            header_row_height = 0.45
            data_row_height = 0.4
        elif cols <= 10:
            header_row_height = 0.4
            data_row_height = 0.35
        else:
            header_row_height = 0.35
            data_row_height = 0.3
        
        for row_idx, row in enumerate(table.rows):
            if row_idx == 0:
                # Header row
                row.height = Inches(header_row_height)
            else:
                # Data rows
                row.height = Inches(data_row_height)
        
        return table_shape
    
    def add_bullet_list(self, slide, items: List[str], left: float, top: float,
                       width: float, height: float, formatting: Optional[Dict] = None) -> Any:
        """
        Add a bullet list to a slide.
        
        Args:
            slide: Slide object
            items: List of bullet items
            left: Left position in inches
            top: Top position in inches
            width: Width in inches
            height: Height in inches
            formatting: Optional formatting dictionary
        
        Returns:
            Shape object
        """
        text_box = self.add_text_box(slide, "", left, top, width, height, formatting)
        text_frame = text_box.text_frame
        text_frame.word_wrap = True
        
        # Add items as paragraphs with bullets
        for i, item in enumerate(items):
            if i == 0:
                paragraph = text_frame.paragraphs[0]
            else:
                paragraph = text_frame.add_paragraph()
            
            paragraph.text = item
            paragraph.level = 0
            paragraph.space_after = Pt(6)
            
            if formatting:
                self.formatter.format_paragraph(paragraph, formatting)
        
        return text_box
    
    def update_text_in_shape(self, slide, shape_index: int, text: str,
                           formatting: Optional[Dict] = None):
        """
        Update text in an existing shape.
        
        Args:
            slide: Slide object
            shape_index: Index of the shape to update
            text: New text content
            formatting: Optional formatting dictionary
        """
        if shape_index < len(slide.shapes):
            shape = slide.shapes[shape_index]
            if shape.has_text_frame:
                shape.text_frame.text = text
                if formatting:
                    self.formatter.format_text_box(shape.text_frame, formatting)
    
    def update_table_data(self, slide, shape_index: int, data: pd.DataFrame,
                         formatting: Optional[Dict] = None):
        """
        Update data in an existing table.
        
        Args:
            slide: Slide object
            shape_index: Index of the shape to update
            data: New DataFrame data
            formatting: Optional formatting dictionary
        """
        if shape_index < len(slide.shapes):
            shape = slide.shapes[shape_index]
            if shape.has_table:
                table = shape.table
                
                # Clear existing data (except header)
                for row_idx in range(len(table.rows) - 1, 0, -1):
                    table._tbl.remove(table.rows[row_idx]._tr)
                
                # Add new rows
                for _, row_data in data.iterrows():
                    new_row = table.rows.add()
                    for col_idx, value in enumerate(row_data):
                        if col_idx < len(new_row.cells):
                            new_row.cells[col_idx].text = str(value)
                            if formatting and "data_formatting" in formatting:
                                self.formatter.format_table_cell(new_row.cells[col_idx], formatting["data_formatting"])
                
                if formatting:
                    self.formatter.format_table(table, formatting)
    
    def find_shape_by_name(self, slide, name: str) -> Optional[Any]:
        """
        Find a shape by name.
        
        Args:
            slide: Slide object
            name: Name of the shape
        
        Returns:
            Shape object or None
        """
        for shape in slide.shapes:
            if shape.name == name:
                return shape
        return None
    
    def populate_slide_from_mapping(self, slide, data: Dict[str, Any],
                                   mapping: Dict[str, Any]):
        """
        Populate a slide based on a mapping configuration.
        
        Args:
            slide: Slide object
            data: Data dictionary
            mapping: Mapping configuration dictionary
        """
        for shape_mapping in mapping.get("shape_mappings", []):
            shape_index = shape_mapping.get("shape_index")
            mapping_type = shape_mapping.get("mapping_type")
            data_source = shape_mapping.get("data_source")
            
            if mapping_type == "text":
                text_value = self._get_text_value(data, shape_mapping)
                if text_value is not None:
                    self.update_text_in_shape(slide, shape_index, str(text_value))
            
            elif mapping_type == "table":
                table_data = self._get_table_data(data, shape_mapping, return_column_mapping=False)
                # Handle tuple return (DataFrame, mapping) if column_mapping was requested
                if isinstance(table_data, tuple):
                    table_data, _ = table_data
                if table_data is not None:
                    self.update_table_data(slide, shape_index, table_data)
    
    def _get_text_value(self, data: Dict[str, Any], mapping: Dict[str, Any]) -> Optional[str]:
        """Extract text value from data based on mapping."""
        data_source = mapping.get("data_source")
        column = mapping.get("column")
        
        if data_source in data:
            df = data[data_source]
            if isinstance(df, pd.DataFrame) and column in df.columns:
                # Return first value or aggregated value
                aggregate = mapping.get("aggregate", "first")
                if aggregate == "sum":
                    return str(df[column].sum())
                elif aggregate == "mean":
                    return str(df[column].mean())
                else:
                    return str(df[column].iloc[0])
        
        return mapping.get("default_value")
    
    def _get_table_data(self, data: Dict[str, Any], mapping: Dict[str, Any], return_column_mapping: bool = False) -> Optional[pd.DataFrame]:
        """Extract table data from data based on mapping."""
        data_source = mapping.get("data_source")
        sheet_name = mapping.get("sheet")
        
        if not data_source:
            print(f"DEBUG: No data_source specified in mapping: {mapping}")
            return None
        
        print(f"DEBUG: Looking for data_source: '{data_source}' (type: {type(data_source)}), available keys: {list(data.keys())}")
        
        # Normalize data_source for matching (strip whitespace)
        data_source_normalized = str(data_source).strip() if data_source else ""
        
        # Try exact match first
        df_source = None
        if data_source_normalized in data:
            df_source = data[data_source_normalized]
            print(f"DEBUG: Found exact match for data_source: '{data_source_normalized}'")
        else:
            # Try case-insensitive match
            data_source_lower = data_source_normalized.lower()
            for key in data.keys():
                key_normalized = str(key).strip()
                if key_normalized.lower() == data_source_lower:
                    df_source = data[key]
                    print(f"DEBUG: Found case-insensitive match: '{key}' for '{data_source_normalized}'")
                    break
            
            # Try partial match (contains)
            if df_source is None:
                for key in data.keys():
                    key_str = str(key).strip().lower()
                    if data_source_lower in key_str or key_str in data_source_lower:
                        df_source = data[key]
                        print(f"DEBUG: Found partial match: '{key}' for '{data_source_normalized}'")
                        break
        
        if df_source is None:
            print(f"WARNING: Could not find data_source '{data_source}' in available data. Available keys: {list(data.keys())[:10]}")
            # Try to use first available data source as fallback
            if data and len(data) > 0:
                first_key = list(data.keys())[0]
                print(f"WARNING: Using first available data source '{first_key}' as fallback")
                df_source = data[first_key]
            else:
                # Return empty DataFrame with structure instead of None
                print(f"WARNING: No data available. Returning empty DataFrame.")
                return pd.DataFrame({"Message": ["No data available"]})
        
        # Handle nested structure: data[file_name][sheet_name]
        if df_source is not None:
            
            # If it's a dict (multiple sheets), get the specific sheet
            if isinstance(df_source, dict):
                if sheet_name and sheet_name in df_source:
                    df = df_source[sheet_name]
                    print(f"DEBUG: Found exact match for sheet: '{sheet_name}'")
                elif sheet_name:
                    # Try case-insensitive match
                    sheet_lower = str(sheet_name).lower().strip()
                    matched_sheet = None
                    for key in df_source.keys():
                        key_str = str(key).lower().strip()
                        if key_str == sheet_lower:
                            matched_sheet = key
                            break
                    
                    if matched_sheet:
                        df = df_source[matched_sheet]
                        print(f"DEBUG: Found case-insensitive match for sheet: '{matched_sheet}' for '{sheet_name}'")
                    else:
                        # Try partial match
                        for key in df_source.keys():
                            key_str = str(key).lower().strip()
                            if sheet_lower in key_str or key_str in sheet_lower:
                                matched_sheet = key
                                break
                        
                        if matched_sheet:
                            df = df_source[matched_sheet]
                            print(f"DEBUG: Found partial match for sheet: '{matched_sheet}' for '{sheet_name}'")
                        else:
                            print(f"WARNING: Sheet '{sheet_name}' not found in {data_source}. Available sheets: {list(df_source.keys())[:10]}")
                            # Return empty DataFrame instead of None
                            return pd.DataFrame({"Message": [f"Sheet '{sheet_name}' not found"]})
                else:
                    # If no sheet specified, use first sheet
                    if df_source:
                        first_sheet = list(df_source.keys())[0]
                        df = df_source[first_sheet]
                        print(f"DEBUG: No sheet specified, using first sheet: '{first_sheet}'")
                    else:
                        df = None
            elif isinstance(df_source, pd.DataFrame):
                df = df_source
                print(f"DEBUG: Data source is a DataFrame, using directly")
            else:
                print(f"WARNING: Data source '{data_source}' has unsupported type: {type(df_source)}")
                # Return empty DataFrame instead of None
                return pd.DataFrame({"Message": ["Data source type not supported"]})
            
            if df is None or not isinstance(df, pd.DataFrame):
                # Return empty DataFrame instead of None
                print(f"WARNING: Data source returned None or invalid type. Returning empty DataFrame.")
                return pd.DataFrame({"Message": ["No data available"]})
            
            # Apply header row offset if needed (data should already be loaded with correct header)
            # But we can re-read if header_row is specified and different
            header_row = mapping.get("header_row")
            if header_row is not None and header_row != 0:
                # Note: This assumes data was loaded with header=0
                # For now, we'll work with the data as-is
                # In future, we might need to re-read with correct header
                pass
            
            # Apply filters if specified
            filters = mapping.get("filters", [])
            result_df = df.copy()
            
            for filter_def in filters:
                column = filter_def.get("column")
                operator = filter_def.get("operator", "!=")
                value = filter_def.get("value")
                
                if column in result_df.columns:
                    if operator == "!=":
                        if value is None:
                            result_df = result_df[result_df[column].notna()]
                        else:
                            result_df = result_df[result_df[column] != value]
                    elif operator == ">=":
                        result_df = result_df[result_df[column] >= value]
                    elif operator == "<=":
                        result_df = result_df[result_df[column] <= value]
                    elif operator == "==":
                        result_df = result_df[result_df[column] == value]
                    elif operator == "notna":
                        result_df = result_df[result_df[column].notna()]
            
            # Select columns if specified
            columns = mapping.get("columns")
            print(f"DEBUG: Column selection - columns parameter: {columns} (type: {type(columns)}, length: {len(columns) if columns else 'N/A'})")
            print(f"DEBUG: Available columns in DataFrame: {list(result_df.columns)}")
            
            # If columns is specified and not empty, use them (preserve user selection and order)
            if columns is not None and len(columns) > 0:
                print(f"DEBUG: User selected {len(columns)} columns, attempting to match: {columns}")
                available_columns = list(result_df.columns)
                matched_cols = []  # Will preserve user's order
                column_mapping_dict = {}  # Maps user column name to actual column name
                
                columns_to_find = columns if isinstance(columns, list) else [columns]
                
                # Enhanced column matching with multiple strategies
                for user_col in columns_to_find:
                    user_col_str = str(user_col).strip()
                    matched_col = None
                    match_type = None
                    
                    # Strategy 1: Exact match (case-sensitive)
                    if user_col_str in available_columns:
                        matched_col = user_col_str
                        match_type = "exact"
                    
                    # Strategy 2: Exact match (case-insensitive, whitespace-insensitive)
                    if not matched_col:
                        user_col_normalized = user_col_str.lower().strip()
                        for avail_col in available_columns:
                            avail_col_normalized = str(avail_col).strip().lower()
                            if avail_col_normalized == user_col_normalized:
                                matched_col = avail_col  # Use original case
                                match_type = "case_insensitive"
                                break
                    
                    # Strategy 3: Partial match (contains)
                    if not matched_col:
                        user_col_normalized = user_col_str.lower().strip()
                        for avail_col in available_columns:
                            avail_col_normalized = str(avail_col).strip().lower()
                            # Check if one contains the other (both directions)
                            if user_col_normalized in avail_col_normalized or avail_col_normalized in user_col_normalized:
                                # Prefer shorter matches for better accuracy
                                if not matched_col or len(avail_col) < len(matched_col):
                                    matched_col = avail_col
                                    match_type = "partial"
                    
                    # Strategy 4: Fuzzy match (handle common variations)
                    if not matched_col:
                        user_col_normalized = user_col_str.lower().strip()
                        # Remove common prefixes/suffixes and special chars for matching
                        user_col_clean = user_col_normalized.replace('_', ' ').replace('-', ' ').replace('.', ' ')
                        user_col_clean = ' '.join(user_col_clean.split())  # Normalize whitespace
                        
                        for avail_col in available_columns:
                            avail_col_normalized = str(avail_col).strip().lower()
                            avail_col_clean = avail_col_normalized.replace('_', ' ').replace('-', ' ').replace('.', ' ')
                            avail_col_clean = ' '.join(avail_col_clean.split())  # Normalize whitespace
                            
                            if user_col_clean == avail_col_clean:
                                matched_col = avail_col
                                match_type = "fuzzy"
                                break
                    
                    if matched_col:
                        # Only add if not already added (avoid duplicates)
                        if matched_col not in matched_cols:
                            matched_cols.append(matched_col)
                            column_mapping_dict[user_col] = matched_col
                            print(f"INFO: Matched column '{user_col}' -> '{matched_col}' ({match_type})")
                        else:
                            print(f"WARNING: Column '{user_col}' matched to '{matched_col}' which was already matched")
                    else:
                        print(f"WARNING: Could not match column '{user_col}' to any available column")
                        print(f"WARNING:   Available columns: {available_columns[:10]}")
                
                # Preserve user's column order and select matched columns
                if matched_cols:
                    # Reorder matched_cols to preserve user's order
                    ordered_matched_cols = []
                    for user_col in columns_to_find:
                        if user_col in column_mapping_dict:
                            actual_col = column_mapping_dict[user_col]
                            if actual_col not in ordered_matched_cols:
                                ordered_matched_cols.append(actual_col)
                    
                    # Use the ordered list to select columns
                    result_df = result_df[ordered_matched_cols]
                    print(f"INFO: Successfully matched {len(ordered_matched_cols)}/{len(columns_to_find)} columns in user's order")
                    print(f"INFO: Selected columns (in order): {list(result_df.columns)}")
                    print(f"DEBUG: Result DataFrame shape: {result_df.shape}")
                    
                    # Create column mapping if requested
                    if return_column_mapping:
                        return (result_df, column_mapping_dict)
                    
                    return result_df
                else:
                    # No columns matched - this is an error, but return all columns as fallback
                    print(f"ERROR: None of the {len(columns_to_find)} selected columns could be matched!")
                    print(f"ERROR: Selected: {columns_to_find}")
                    print(f"ERROR: Available: {available_columns[:20]}")
                    print(f"INFO: Falling back to all columns to prevent data loss")
                    # Continue to return all columns below
            elif columns is not None and len(columns) == 0:
                # Empty columns list - return all columns (user wants all columns)
                print(f"INFO: Empty columns list - returning all columns")
                # Continue to return all columns below
            else:
                # columns is None - return all columns
                print(f"INFO: No columns specified - returning all columns")
                # Continue to return all columns below
            
            # Limit rows if specified
            max_rows = mapping.get("max_rows")
            if max_rows:
                result_df = result_df.head(max_rows)
            
            # If no column selection was done (columns not specified, empty list, or matching failed), return all columns
            print(f"INFO: Returning DataFrame with all columns. Shape: {result_df.shape}, Columns: {list(result_df.columns)}")
            
            # Validate data integrity - ensure we have data
            if len(result_df) == 0:
                print(f"WARNING: DataFrame is empty (no rows). Data source: {data_source}, Sheet: {sheet_name}")
            if len(result_df.columns) == 0:
                print(f"WARNING: DataFrame has no columns. Data source: {data_source}, Sheet: {sheet_name}")
            
            # Log final result
            print(f"INFO: Final DataFrame - {len(result_df)} rows, {len(result_df.columns)} columns")
            if return_column_mapping:
                # Create identity mapping for all columns
                column_mapping = {col: col for col in result_df.columns}
                print(f"INFO: Column mapping created: {len(column_mapping)} mappings")
                return (result_df, column_mapping)
            
            return result_df
        
        # If we get here, something went wrong - return empty DataFrame instead of None
        print(f"WARNING: _get_table_data reached end without returning data. Returning empty DataFrame.")
        return pd.DataFrame({"Message": ["No data available"]})
    
    def add_chart(self, slide, data: pd.DataFrame, chart_type: str = "column",
                  left: float = 1, top: float = 2, width: float = 8, height: float = 4,
                  x_column: Optional[str] = None, y_columns: Optional[List[str]] = None,
                  title: str = "", formatting: Optional[Dict] = None) -> Any:
        """
        Add a chart to a slide.
        
        Args:
            slide: Slide object
            data: DataFrame containing chart data
            chart_type: Type of chart ("column", "bar", "line", "pie")
            left: Left position in inches
            top: Top position in inches
            width: Width in inches
            height: Height in inches
            x_column: Column name for X-axis categories
            y_columns: List of column names for Y-axis values
            title: Chart title
            formatting: Optional formatting dictionary
        
        Returns:
            Shape object
        """
        if data is None or len(data) == 0:
            raise ValueError("Cannot create chart with empty data")
        
        print(f"DEBUG: Creating chart with data shape: {data.shape}, columns: {list(data.columns)}")
        
        # Default to first column as X, remaining as Y
        if x_column is None:
            x_column = data.columns[0] if len(data.columns) > 0 else None
        
        if y_columns is None or len(y_columns) == 0:
            y_columns = [col for col in data.columns if col != x_column]
            if len(y_columns) == 0:
                y_columns = [data.columns[0]] if len(data.columns) > 0 else []
        
        if not x_column or not y_columns:
            raise ValueError(f"Cannot create chart: x_column={x_column}, y_columns={y_columns}")
        
        # Verify columns exist in data
        if x_column not in data.columns:
            # Try to find matching column
            x_col_lower = str(x_column).lower().strip()
            for col in data.columns:
                if str(col).lower().strip() == x_col_lower:
                    x_column = col
                    break
            if x_column not in data.columns:
                raise ValueError(f"X column '{x_column}' not found in data. Available: {list(data.columns)}")
        
        # Filter y_columns to only those that exist
        valid_y_columns = []
        for y_col in (y_columns if isinstance(y_columns, list) else [y_columns]):
            if y_col in data.columns:
                valid_y_columns.append(y_col)
            else:
                # Try case-insensitive match
                y_col_lower = str(y_col).lower().strip()
                for col in data.columns:
                    if str(col).lower().strip() == y_col_lower:
                        valid_y_columns.append(col)
                        print(f"DEBUG: Matched y_column '{y_col}' to '{col}'")
                        break
        
        if not valid_y_columns:
            raise ValueError(f"No valid Y columns found. Requested: {y_columns}, Available: {list(data.columns)}")
        
        y_columns = valid_y_columns
        print(f"DEBUG: Using x_column='{x_column}', y_columns={y_columns}")
        
        # Prepare chart data
        chart_data = CategoryChartData()
        
        # Get categories from X column
        categories = data[x_column].astype(str).tolist()
        chart_data.categories = categories
        print(f"DEBUG: Categories ({len(categories)}): {categories[:5]}...")
        
        # Add series for each Y column
        for y_col in y_columns:
            if y_col in data.columns and y_col != x_column:
                values = data[y_col].tolist()
                # Convert to numeric values
                numeric_values = []
                for val in values:
                    try:
                        # Handle percentage strings, commas, etc.
                        val_str = str(val).replace('%', '').replace(',', '').strip()
                        numeric_values.append(float(val_str))
                    except (ValueError, TypeError):
                        numeric_values.append(0.0)
                
                print(f"DEBUG: Added series '{y_col}' with {len(numeric_values)} values: {numeric_values[:5]}...")
                chart_data.add_series(y_col, numeric_values)
        
        # Map chart type string to enum
        chart_type_map = {
            "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "bar": XL_CHART_TYPE.BAR_CLUSTERED,
            "line": XL_CHART_TYPE.LINE,
            "pie": XL_CHART_TYPE.PIE,
            "area": XL_CHART_TYPE.AREA,
        }
        
        chart_type_enum = chart_type_map.get(chart_type.lower(), XL_CHART_TYPE.COLUMN_CLUSTERED)
        
        # Add chart to slide
        left_inches = Inches(left)
        top_inches = Inches(top)
        width_inches = Inches(width)
        height_inches = Inches(height)
        
        graphic_frame = slide.shapes.add_chart(
            chart_type_enum, left_inches, top_inches, width_inches, height_inches, chart_data
        )
        chart = graphic_frame.chart
        
        # Set chart title if provided
        if title:
            chart.has_title = True
            chart.chart_title.text_frame.text = title
            # Format chart title
            for paragraph in chart.chart_title.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(16)
                    run.font.bold = True
        
        # Apply formatting - use title color shades (#004E6F) for charts
        title_color_base = "#004E6F"  # Title color
        # Generate shades of title color for multiple series
        def get_color_shade(base_hex, index, total):
            """Generate a shade of the base color."""
            if base_hex.startswith("#"):
                base_hex = base_hex[1:]
            r_base = int(base_hex[0:2], 16)
            g_base = int(base_hex[2:4], 16)
            b_base = int(base_hex[4:6], 16)
            
            # Create lighter shades by mixing with white
            # First series: darkest (original), subsequent: progressively lighter
            if total == 1:
                return (r_base, g_base, b_base)
            
            # Mix with white: (1 - factor) * base + factor * 255
            factor = index / (total - 1) * 0.4  # Max 40% lighter
            r = int((1 - factor) * r_base + factor * 255)
            g = int((1 - factor) * g_base + factor * 255)
            b = int((1 - factor) * b_base + factor * 255)
            return (r, g, b)
        
        # Apply colors to chart series - use title color shades
        num_series = len(chart.series)
        for i, series in enumerate(chart.series):
            # Use title color shades if colors not specified, or use provided colors
            if formatting and "colors" in formatting and len(formatting.get("colors", [])) > i:
                color_str = formatting["colors"][i]
                if color_str.startswith("#"):
                    color_str = color_str[1:]
                try:
                    r, g, b = int(color_str[0:2], 16), int(color_str[2:4], 16), int(color_str[4:6], 16)
                except (ValueError, IndexError):
                    r, g, b = get_color_shade(title_color_base, i, num_series)
            else:
                # Default: use title color shades
                r, g, b = get_color_shade(title_color_base, i, num_series)
            
            try:
                fill = series.format.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(r, g, b)
            except (ValueError, AttributeError):
                pass
        
        return graphic_frame

