"""
Flask Backend API for PPT Automation Web Interface
"""
import os
import sys
import json
import uuid
from pathlib import Path
from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from werkzeug.utils import secure_filename
import pandas as pd

# Add parent directory to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from src.data_loader import DataLoader
from src.data_normalizer import DataNormalizer
from src.rules_engine import RulesEngine
from src.ppt_generator import PPTGenerator

# Import config_builder from web directory
sys.path.insert(0, os.path.dirname(__file__))
from config_builder import ConfigBuilder

app = Flask(__name__)
CORS(app)

# Configuration
# Use absolute paths to ensure uploads work from any directory
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'output')
ALLOWED_EXTENSIONS = {'xlsx', 'xlsb', 'xls'}

# Ensure directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max file size


def allowed_file(filename):
    """Check if file extension is allowed."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """Serve the main frontend page."""
    return render_template('index.html')


@app.route('/api/analyze-excel', methods=['POST'])
def analyze_excel():
    """Analyze uploaded Excel file structure."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Only Excel files (.xlsx, .xlsb, .xls) are allowed.'}), 400
    
    try:
        # Save uploaded file
        filename = secure_filename(file.filename)
        file_id = str(uuid.uuid4())
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{file_id}_{filename}")
        file.save(file_path)
        
        # Analyze Excel structure
        loader = DataLoader()
        data = loader.load_excel(file_path)
        
        analysis = {
            'file_id': file_id,
            'filename': filename,
            'sheets': []
        }
        
        def clean_dataframe_for_json(df):
            """Replace NaN, NaT, and other non-JSON-serializable values."""
            # Convert to dict first, then clean NaN values
            records = df.to_dict(orient='records')
            # Clean records: replace any NaN/NaT/inf values with None (which becomes null in JSON)
            for record in records:
                for key, value in record.items():
                    # Check for NaN/NaT
                    if pd.isna(value):
                        record[key] = None
                    # Check for inf/-inf
                    elif isinstance(value, float) and (value == float('inf') or value == float('-inf')):
                        record[key] = None
            return records
        
        if isinstance(data, pd.DataFrame):
            # Single sheet
            sample_data = clean_dataframe_for_json(data.head(5))
            analysis['sheets'].append({
                'name': 'Sheet1',
                'columns': list(data.columns),
                'row_count': len(data),
                'sample_data': sample_data
            })
        elif isinstance(data, dict):
            # Multiple sheets
            for sheet_name, df in data.items():
                if isinstance(df, pd.DataFrame):
                    sample_data = clean_dataframe_for_json(df.head(5))
                    analysis['sheets'].append({
                        'name': sheet_name,
                        'columns': list(df.columns),
                        'row_count': len(df),
                        'sample_data': sample_data
                    })
        
        return jsonify(analysis)
    
    except Exception as e:
        return jsonify({'error': f'Error analyzing file: {str(e)}'}), 500


@app.route('/api/excel-sheets', methods=['GET'])
def get_excel_sheets():
    """Get list of sheets from an uploaded Excel file."""
    file_id = request.args.get('file_id')
    if not file_id:
        return jsonify({'error': 'file_id required'}), 400
    
    try:
        # Find the file
        upload_dir = Path(app.config['UPLOAD_FOLDER'])
        files = list(upload_dir.glob(f"{file_id}_*"))
        if not files:
            return jsonify({'error': 'File not found'}), 404
        
        file_path = str(files[0])
        
        # Load and analyze
        loader = DataLoader()
        data = loader.load_excel(file_path)
        
        sheets = []
        if isinstance(data, pd.DataFrame):
            sheets.append({'name': 'Sheet1', 'row_count': len(data)})
        elif isinstance(data, dict):
            for sheet_name, df in data.items():
                if isinstance(df, pd.DataFrame):
                    sheets.append({'name': sheet_name, 'row_count': len(df)})
        
        return jsonify({'sheets': sheets})
    
    except Exception as e:
        return jsonify({'error': f'Error getting sheets: {str(e)}'}), 500


@app.route('/api/excel-columns', methods=['GET'])
def get_excel_columns():
    """Get columns from a specific sheet."""
    file_id = request.args.get('file_id')
    sheet_name = request.args.get('sheet')
    
    if not file_id:
        return jsonify({'error': 'file_id required'}), 400
    
    try:
        # Find the file
        upload_dir = Path(app.config['UPLOAD_FOLDER'])
        files = list(upload_dir.glob(f"{file_id}_*"))
        if not files:
            return jsonify({'error': 'File not found'}), 404
        
        file_path = str(files[0])
        
        # Load specific sheet
        loader = DataLoader()
        data = loader.load_excel(file_path)
        
        if isinstance(data, pd.DataFrame):
            df = data
        elif isinstance(data, dict):
            if sheet_name not in data:
                return jsonify({'error': f'Sheet "{sheet_name}" not found'}), 404
            df = data[sheet_name]
        else:
            return jsonify({'error': 'Invalid data structure'}), 500
        
        columns = [{'name': col, 'dtype': str(df[col].dtype)} for col in df.columns]
        
        return jsonify({'columns': columns})
    
    except Exception as e:
        return jsonify({'error': f'Error getting columns: {str(e)}'}), 500


@app.route('/api/templates', methods=['GET'])
def get_templates():
    """Get list of available PowerPoint templates."""
    templates_dir = Path('templates')
    templates = []
    
    if templates_dir.exists():
        for template_file in templates_dir.glob('*.pptx'):
            templates.append({
                'name': template_file.stem,
                'path': str(template_file)
            })
    
    return jsonify({'templates': templates})


@app.route('/api/generate-ppt', methods=['POST'])
def generate_ppt():
    """Generate PowerPoint deck from configuration."""
    try:
        data = request.get_json()
        
        # Extract configuration
        uploaded_files_info = data.get('uploaded_files', {})
        affiliate = data.get('affiliate')
        slides_config = data.get('slides_config', [])
        
        if not affiliate:
            return jsonify({'error': 'Affiliate is required'}), 400
        
        if not uploaded_files_info:
            return jsonify({'error': 'No files provided'}), 400
        
        if not slides_config:
            return jsonify({'error': 'No slides configured'}), 400
        
        # Always use Template folder template
        # Try relative path first (from project root)
        template_path = os.path.join('Template', 'Template.pptx')
        if not os.path.exists(template_path):
            # Try from web directory perspective
            web_dir = os.path.dirname(__file__)
            project_root = os.path.dirname(web_dir)
            template_path = os.path.join(project_root, 'Template', 'Template.pptx')
            if not os.path.exists(template_path):
                return jsonify({'error': f'Template file not found. Checked: Template/Template.pptx and {template_path}'}), 404
        
        # Load all uploaded files
        upload_dir = Path(app.config['UPLOAD_FOLDER'])
        all_loaded_data = {}
        
        for file_id, file_info in uploaded_files_info.items():
            # Find the uploaded file
            files = list(upload_dir.glob(f"{file_id}_*"))
            if not files:
                continue  # Skip if file not found
            
            file_path = str(files[0])
            
            # Load and process data
            loader = DataLoader()
            loaded_data = loader.load_excel(file_path)
            
            # Normalize data
            normalizer = DataNormalizer()
            # Use file name without extension as key (normalize for matching)
            file_key = os.path.splitext(file_info['name'])[0].strip()
            
            print(f"DEBUG: Loading file '{file_info['name']}' with key '{file_key}'")
            
            if isinstance(loaded_data, pd.DataFrame):
                # Preserve column names for accurate matching with user selections
                normalized_df = normalizer.normalize_data(loaded_data, preserve_names=True)
                all_loaded_data[file_key] = normalized_df
                print(f"DEBUG: Loaded single DataFrame with {len(normalized_df)} rows, {len(normalized_df.columns)} columns")
                print(f"DEBUG:   Columns: {list(normalized_df.columns)}")
            elif isinstance(loaded_data, dict):
                # Preserve column names for accurate matching with user selections
                all_loaded_data[file_key] = {
                    sheet: normalizer.normalize_data(df, preserve_names=True)
                    for sheet, df in loaded_data.items()
                }
                print(f"DEBUG: Loaded multi-sheet file with sheets: {list(all_loaded_data[file_key].keys())}")
                for sheet_name, df in all_loaded_data[file_key].items():
                    print(f"DEBUG:   Sheet '{sheet_name}': {len(df)} rows, {len(df.columns)} columns")
                    print(f"DEBUG:     Columns: {list(df.columns)}")
        
        # Debug: Print loaded data keys
        print(f"DEBUG: All loaded data keys: {list(all_loaded_data.keys())}")
        for key, value in all_loaded_data.items():
            if isinstance(value, dict):
                print(f"DEBUG:   '{key}': dict with sheets {list(value.keys())}")
            else:
                print(f"DEBUG:   '{key}': DataFrame with columns {list(value.columns)[:5] if hasattr(value, 'columns') else 'N/A'}")
        
        # Build configuration from frontend data
        config_builder = ConfigBuilder()
        slides_yaml = config_builder.build_slides_config(slides_config)
        
        # Debug: Print slide configs
        print(f"DEBUG: Generated {len(slides_yaml.get('slides', []))} slide configs")
        
        # Validate data availability for all slides
        validation_errors = []
        validation_warnings = []
        
        for i, slide_config in enumerate(slides_yaml.get('slides', [])):
            table_mapping = slide_config.get('table_mapping', {})
            if table_mapping:
                data_source = table_mapping.get('data_source')
                sheet_name = table_mapping.get('sheet')
                columns = table_mapping.get('columns', [])
                
                print(f"DEBUG: Slide {i+1} - data_source: '{data_source}', sheet: '{sheet_name}', columns: {columns}")
                
                # Validate data source exists
                if data_source:
                    data_source_normalized = str(data_source).strip()
                    found = False
                    
                    # Try exact match
                    if data_source_normalized in all_loaded_data:
                        found = True
                    else:
                        # Try case-insensitive match
                        data_source_lower = data_source_normalized.lower()
                        for key in all_loaded_data.keys():
                            if str(key).strip().lower() == data_source_lower:
                                found = True
                                break
                    
                    if not found:
                        warning_msg = f"Slide {i+1}: Data source '{data_source}' not found. Available: {list(all_loaded_data.keys())[:3]}"
                        validation_warnings.append(warning_msg)
                        print(f"WARNING: {warning_msg}")
                    
                    # Validate sheet if data source found
                    if found and sheet_name:
                        df_source = all_loaded_data.get(data_source_normalized)
                        if df_source is None:
                            # Try case-insensitive match
                            for key in all_loaded_data.keys():
                                if str(key).strip().lower() == data_source_normalized.lower():
                                    df_source = all_loaded_data[key]
                                    break
                        
                        if isinstance(df_source, dict):
                            sheet_found = False
                            sheet_lower = str(sheet_name).lower().strip()
                            for sheet_key in df_source.keys():
                                if str(sheet_key).lower().strip() == sheet_lower:
                                    sheet_found = True
                                    break
                            
                            if not sheet_found:
                                warning_msg = f"Slide {i+1}: Sheet '{sheet_name}' not found in '{data_source}'. Available: {list(df_source.keys())[:3]}"
                                validation_warnings.append(warning_msg)
                                print(f"WARNING: {warning_msg}")
                        
                        # Validate columns if sheet found
                        if columns and len(columns) > 0:
                            df = None
                            if isinstance(df_source, dict):
                                for sheet_key in df_source.keys():
                                    if str(sheet_key).lower().strip() == str(sheet_name).lower().strip():
                                        df = df_source[sheet_key]
                                        break
                            elif isinstance(df_source, pd.DataFrame):
                                df = df_source
                            
                            if df is not None and hasattr(df, 'columns'):
                                missing_cols = []
                                available_cols = [str(c) for c in df.columns]
                                for col in columns:
                                    col_str = str(col).strip()
                                    if col_str not in available_cols:
                                        # Try case-insensitive
                                        found_col = False
                                        for avail_col in available_cols:
                                            if avail_col.lower().strip() == col_str.lower():
                                                found_col = True
                                                break
                                        if not found_col:
                                            missing_cols.append(col)
                                
                                if missing_cols:
                                    warning_msg = f"Slide {i+1}: Columns {missing_cols} not found. Available: {available_cols[:5]}"
                                    validation_warnings.append(warning_msg)
                                    print(f"WARNING: {warning_msg}")
        
        # Log validation results
        if validation_warnings:
            print(f"INFO: Validation completed with {len(validation_warnings)} warnings. Generation will proceed with fallbacks.")
        else:
            print(f"INFO: All data sources validated successfully.")
        
        # Save temporary config
        config_id = str(uuid.uuid4())
        temp_config_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{config_id}_slides.yaml")
        with open(temp_config_path, 'w') as f:
            import yaml
            yaml.dump(slides_yaml, f)
        
        # Generate PPT
        generator = PPTGenerator(
            template_path=template_path,
            slides_config=temp_config_path,
            affiliate=affiliate
        )
        
        output_id = str(uuid.uuid4())
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{output_id}.pptx")
        
        generator.generate(all_loaded_data, output_path)
        
        return jsonify({
            'success': True,
            'output_id': output_id,
            'message': 'PowerPoint deck generated successfully'
        })
    
    except Exception as e:
        import traceback
        return jsonify({'error': f'Error generating PPT: {str(e)}\n{traceback.format_exc()}'}), 500


@app.route('/api/download/<output_id>', methods=['GET'])
def download_ppt(output_id):
    """Download generated PowerPoint file."""
    try:
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{output_id}.pptx")
        
        if not os.path.exists(output_path):
            return jsonify({'error': 'File not found'}), 404
        
        return send_file(
            output_path,
            as_attachment=True,
            download_name=f'generated_presentation_{output_id}.pptx',
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
    
    except Exception as e:
        return jsonify({'error': f'Error downloading file: {str(e)}'}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)

