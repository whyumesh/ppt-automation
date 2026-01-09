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
        # Handle empty DataFrame
        if data is None or len(data) == 0:
            raise ValueError("Cannot create table with empty data")
        
        rows = len(data) + 1  # +1 for header
        cols = len(data.columns)
        
        left_inches = Inches(left)
        top_inches = Inches(top)
        width_inches = Inches(width)
        height_inches = Inches(height)
        
        table_shape = slide.shapes.add_table(rows, cols, left_inches, top_inches, width_inches, height_inches)
        table = table_shape.table
        
        # Add visual enhancements - borders and shadows
        try:
            # Enable table borders
            for row in table.rows:
                for cell in row.cells:
                    # Add borders to cells
                    cell.fill.solid()
                    if cell.fill.fore_color:
                        # Keep existing fill, just add border
                        pass
            
            # Add shadow effect (if supported)
            # table_shape.shadow.inherit = False
        except:
            pass  # Border formatting may not always work
        
        # Merge formatting with defaults
        if formatting is None:
            formatting = {}
        
        # Default header formatting (if not specified) - more vibrant
        default_header_formatting = {
            "font_size": 14,
            "bold": True,
            "fill_color": "#003B55",  # Dark blue
            "font_color": "#FFFFFF"   # White text
        }
        
        # Default data formatting (if not specified) - better readability
        default_data_formatting = {
            "font_size": 12,  # Increased from 11
            "bold": False,
            "font_color": "#333333"  # Dark gray for better contrast
        }
        
        header_formatting = formatting.get("header_formatting", default_header_formatting)
        data_formatting = formatting.get("data_formatting", default_data_formatting)
        
        # Ensure header formatting has white text
        if "fill_color" in header_formatting and header_formatting["fill_color"] in ["#003B55", "#003b55"]:
            header_formatting["font_color"] = "#FFFFFF"
        
        # Populate header row
        for col_idx, col_name in enumerate(data.columns):
            cell = table.cell(0, col_idx)
            # Clean column name for display
            cell.text = str(col_name).strip()
            self.formatter.format_table_cell(cell, header_formatting)
            # Center align header text
            cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Populate data rows with number formatting and alternating row colors
        number_formatting = formatting.get("number_formatting", {})
        alternate_row_color = formatting.get("alternate_row_color", "#F5F5F5")  # Light gray
        
        for row_idx, (_, row_data) in enumerate(data.iterrows(), start=1):
            # Apply alternating row background color for better readability
            row_formatting = data_formatting.copy()
            if row_idx % 2 == 0:  # Even rows get alternate color
                row_formatting["fill_color"] = alternate_row_color
            
            for col_idx, value in enumerate(row_data):
                cell = table.cell(row_idx, col_idx)
                
                # Apply number formatting if specified
                col_name = data.columns[col_idx]
                if col_name in number_formatting:
                    format_type = number_formatting[col_name]
                    if format_type == "percentage":
                        cell.text = f"{float(value) * 100:.1f}%"
                    elif format_type == "percentage_decimal":
                        # Value is already decimal (0-1), convert to percentage
                        try:
                            cell.text = f"{float(value) * 100:.1f}%"
                        except (ValueError, TypeError):
                            cell.text = str(value)
                    elif format_type == "number":
                        try:
                            cell.text = f"{float(value):,.0f}"
                        except (ValueError, TypeError):
                            cell.text = str(value)
                    elif format_type == "currency":
                        try:
                            cell.text = f"${float(value):,.2f}"
                        except (ValueError, TypeError):
                            cell.text = str(value)
                    else:
                        cell.text = str(value)
                else:
                    cell.text = str(value)
                
                # Apply formatting
                cell_formatting = row_formatting.copy()
                
                # Apply conditional formatting if specified (overrides row color)
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
                                    cell_formatting["font_color"] = cond.get("color", "#C00000")
                            except (ValueError, TypeError):
                                pass
                
                self.formatter.format_table_cell(cell, cell_formatting)
                # Left align data
                cell.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
        
        # Apply additional formatting
        if formatting:
            self.formatter.format_table(table, formatting)
        
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
        
        print(f"DEBUG: Looking for data_source: '{data_source}', available keys: {list(data.keys())}")
        
        # Try exact match first
        df_source = None
        if data_source in data:
            df_source = data[data_source]
            print(f"DEBUG: Found exact match for data_source: '{data_source}'")
        else:
            # Try case-insensitive match
            data_source_lower = str(data_source).lower().strip()
            for key in data.keys():
                if str(key).lower().strip() == data_source_lower:
                    df_source = data[key]
                    print(f"DEBUG: Found case-insensitive match: '{key}' for '{data_source}'")
                    break
            
            # Try partial match (contains)
            if df_source is None:
                for key in data.keys():
                    key_str = str(key).lower().strip()
                    if data_source_lower in key_str or key_str in data_source_lower:
                        df_source = data[key]
                        print(f"DEBUG: Found partial match: '{key}' for '{data_source}'")
                        break
        
        if df_source is None:
            print(f"ERROR: Could not find data_source '{data_source}' in available data. Available keys: {list(data.keys())[:10]}")
            return None
        
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
                            print(f"ERROR: Sheet '{sheet_name}' not found in {data_source}. Available sheets: {list(df_source.keys())[:10]}")
                            return None
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
                print(f"ERROR: Data source '{data_source}' has unsupported type: {type(df_source)}")
                return None
            
            if df is None or not isinstance(df, pd.DataFrame):
                return None
            
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
            if columns and len(columns) > 0:
                # Handle column name mapping (e.g., "Unnamed: 1" -> actual column)
                available_columns = list(result_df.columns)
                selected_cols = []
                columns_to_find = columns if isinstance(columns, list) else [columns]
                
                for col in columns_to_find:
                    col_str = str(col).strip()
                    if col_str in available_columns:
                        selected_cols.append(col_str)
                    else:
                        # Try case-insensitive match
                        col_lower = col_str.lower().strip()
                        matched_col = None
                        for avail_col in available_columns:
                            avail_col_str = str(avail_col).strip()
                            if avail_col_str.lower() == col_lower:
                                matched_col = avail_col_str
                                break
                        if matched_col:
                            selected_cols.append(matched_col)
                        else:
                            # Try partial match (contains)
                            for avail_col in available_columns:
                                avail_col_str = str(avail_col).strip()
                                if col_lower in avail_col_str.lower() or avail_col_str.lower() in col_lower:
                                    matched_col = avail_col_str
                                    break
                            if matched_col:
                                selected_cols.append(matched_col)
                                print(f"Info: Matched column '{col}' to '{matched_col}' (partial match)")
                            else:
                                print(f"Warning: Column '{col}' not found. Available columns: {available_columns[:20]}")
                
                if selected_cols:
                    # Only keep columns that actually exist
                    existing_cols = [col for col in selected_cols if col in result_df.columns]
                    if existing_cols:
                        result_df = result_df[existing_cols]
                        print(f"Info: Selected {len(existing_cols)} columns: {existing_cols}")
                        
                        # Create column mapping if requested
                        if return_column_mapping:
                            column_mapping = {}
                            columns_to_find = columns if isinstance(columns, list) else [columns]
                            for i, req_col in enumerate(columns_to_find):
                                if i < len(existing_cols):
                                    column_mapping[req_col] = existing_cols[i]
                            return (result_df, column_mapping)
                        
                        return result_df
                    else:
                        print(f"Warning: None of the selected columns exist in the dataframe. Selected: {selected_cols}, Available: {list(result_df.columns)[:10]}")
                        return None
                else:
                    print(f"Warning: No matching columns found for {columns}. Available: {available_columns[:20]}")
                    return None
            elif columns is not None and len(columns) == 0:
                # Empty columns list means return None
                print(f"Warning: Empty columns list specified")
                return None
            
            # Limit rows if specified
            max_rows = mapping.get("max_rows")
            if max_rows:
                result_df = result_df.head(max_rows)
            
            # If no column selection was done (columns not specified or all columns used), return as-is
            if return_column_mapping:
                # Create identity mapping for all columns
                column_mapping = {col: col for col in result_df.columns}
                return (result_df, column_mapping)
            
            return result_df
        
        return None
    
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
        
        # Apply formatting if specified
        if formatting:
            # Format chart colors if specified
            if "colors" in formatting and chart.has_legend:
                colors = formatting.get("colors", [])
                for i, series in enumerate(chart.series):
                    if i < len(colors):
                        color_str = colors[i]
                        if color_str.startswith("#"):
                            color_str = color_str[1:]
                        try:
                            r, g, b = int(color_str[0:2], 16), int(color_str[2:4], 16), int(color_str[4:6], 16)
                            fill = series.format.fill
                            fill.solid()
                            fill.fore_color.rgb = RGBColor(r, g, b)
                        except (ValueError, IndexError):
                            pass
        
        return graphic_frame

