"""
Configuration Builder
Converts frontend form data to YAML configuration format
"""
from typing import List, Dict, Any


class ConfigBuilder:
    """Builds YAML configuration from frontend form data."""
    
    def build_slides_config(self, slides_config: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Build slides.yaml structure from frontend configuration.
        
        Args:
            slides_config: List of slide configurations from frontend
        
        Returns:
            Dictionary ready to be saved as YAML
        """
        return {
            'slides': [
                self._build_slide_config(slide) for slide in slides_config
            ]
        }
    
    def _build_slide_config(self, slide_data: Dict[str, Any]) -> Dict[str, Any]:
        """Build a single slide configuration."""
        config = {
            'slide_number': slide_data.get('slide_number', 1),
            'slide_type': slide_data.get('slide_type', 'content'),
            'title': slide_data.get('title', ''),
            'layout_name': slide_data.get('layout_name', 'Title Only')
        }
        
        # Add subtitle if present
        if slide_data.get('subtitle'):
            config['subtitle'] = slide_data['subtitle']
        
        # Add title formatting if present
        if slide_data.get('title_formatting'):
            config['title_formatting'] = slide_data['title_formatting']
        
        # Add subtitle formatting if present
        if slide_data.get('subtitle_formatting'):
            config['subtitle_formatting'] = slide_data['subtitle_formatting']
        
        # Build chart configuration if enabled
        chart_config = slide_data.get('chart', {})
        if chart_config and chart_config.get('enabled', False):
            config['chart'] = {
                'enabled': True,
                'type': chart_config.get('type', 'column'),
                'title': chart_config.get('title', ''),
                'x_column': chart_config.get('x_column', ''),
                'y_columns': chart_config.get('y_columns', []),
                'data_source': slide_data.get('data_source'),
                'sheet': slide_data.get('sheet'),
                'header_row': slide_data.get('header_row', 0)
            }
        
        # Build table mapping if slide type is table
        if slide_data.get('slide_type') == 'table':
            table_mapping = self._build_table_mapping(slide_data)
            if table_mapping:
                config['table_mapping'] = table_mapping
        
        # Build content mappings for content slides
        if slide_data.get('slide_type') == 'content':
            content_mappings = slide_data.get('content_mappings', [])
            if content_mappings:
                config['content_mappings'] = content_mappings
        
        return config
    
    def _build_table_mapping(self, slide_data: Dict[str, Any]) -> Dict[str, Any]:
        """Build table mapping configuration."""
        # data_source is the file name (without extension) from frontend
        data_source = slide_data.get('data_source')
        sheet = slide_data.get('sheet')
        columns = slide_data.get('columns', [])
        
        # Ensure columns is a list
        if columns is None:
            columns = []
        elif not isinstance(columns, list):
            columns = [columns] if columns else []
        
        # Normalize data_source (strip whitespace)
        if data_source:
            data_source = str(data_source).strip()
        
        if not data_source or not sheet:
            return None
        
        mapping = {
            'data_source': data_source,
            'sheet': sheet,
            'header_row': slide_data.get('header_row', 0),
            'columns': columns
        }
        
        print(f"DEBUG ConfigBuilder: Built mapping - data_source: '{data_source}', sheet: '{sheet}', columns: {columns} (type: {type(columns)}, len: {len(columns)})")
        
        # Add filters if present
        filters = slide_data.get('filters', [])
        if filters:
            mapping['filters'] = filters
        
        # Add max_rows if specified
        if slide_data.get('max_rows'):
            mapping['max_rows'] = slide_data['max_rows']
        
        # Add formatting if present
        formatting = slide_data.get('formatting')
        if formatting:
            mapping['formatting'] = formatting
        
        return mapping

