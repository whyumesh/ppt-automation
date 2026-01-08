"""
PPT Builder
Slide building utilities for creating PowerPoint slides.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
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
        rows = len(data) + 1  # +1 for header
        cols = len(data.columns)
        
        left_inches = Inches(left)
        top_inches = Inches(top)
        width_inches = Inches(width)
        height_inches = Inches(height)
        
        table_shape = slide.shapes.add_table(rows, cols, left_inches, top_inches, width_inches, height_inches)
        table = table_shape.table
        
        # Populate header row
        for col_idx, col_name in enumerate(data.columns):
            cell = table.cell(0, col_idx)
            cell.text = str(col_name)
            if formatting and "header_formatting" in formatting:
                self.formatter.format_table_cell(cell, formatting["header_formatting"])
        
        # Populate data rows
        for row_idx, (_, row_data) in enumerate(data.iterrows(), start=1):
            for col_idx, value in enumerate(row_data):
                cell = table.cell(row_idx, col_idx)
                cell.text = str(value)
                if formatting and "data_formatting" in formatting:
                    self.formatter.format_table_cell(cell, formatting["data_formatting"])
        
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
                table_data = self._get_table_data(data, shape_mapping)
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
    
    def _get_table_data(self, data: Dict[str, Any], mapping: Dict[str, Any]) -> Optional[pd.DataFrame]:
        """Extract table data from data based on mapping."""
        data_source = mapping.get("data_source")
        
        if data_source in data:
            df = data[data_source]
            if isinstance(df, pd.DataFrame):
                # Apply filters if specified
                filters = mapping.get("filters", [])
                result_df = df.copy()
                
                for filter_def in filters:
                    column = filter_def.get("column")
                    operator = filter_def.get("operator", ">=")
                    value = filter_def.get("value")
                    
                    if column in result_df.columns:
                        if operator == ">=":
                            result_df = result_df[result_df[column] >= value]
                        elif operator == "<=":
                            result_df = result_df[result_df[column] <= value]
                        elif operator == "==":
                            result_df = result_df[result_df[column] == value]
                        # Add more operators as needed
                
                # Select columns if specified
                columns = mapping.get("columns")
                if columns:
                    result_df = result_df[columns]
                
                # Limit rows if specified
                max_rows = mapping.get("max_rows")
                if max_rows:
                    result_df = result_df.head(max_rows)
                
                return result_df
        
        return None

