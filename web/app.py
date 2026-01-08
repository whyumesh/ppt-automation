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
UPLOAD_FOLDER = 'web/uploads'
OUTPUT_FOLDER = 'web/output'
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
        
        if isinstance(data, pd.DataFrame):
            # Single sheet
            analysis['sheets'].append({
                'name': 'Sheet1',
                'columns': list(data.columns),
                'row_count': len(data),
                'sample_data': data.head(5).to_dict(orient='records')
            })
        elif isinstance(data, dict):
            # Multiple sheets
            for sheet_name, df in data.items():
                if isinstance(df, pd.DataFrame):
                    analysis['sheets'].append({
                        'name': sheet_name,
                        'columns': list(df.columns),
                        'row_count': len(df),
                        'sample_data': df.head(5).to_dict(orient='records')
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
        template_path = data.get('template_path', 'templates/template.pptx')
        slides_config = data.get('slides_config', [])
        
        if not uploaded_files_info:
            return jsonify({'error': 'No files provided'}), 400
        
        if not slides_config:
            return jsonify({'error': 'No slides configured'}), 400
        
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
            file_key = os.path.splitext(file_info['name'])[0]
            
            if isinstance(loaded_data, pd.DataFrame):
                all_loaded_data[file_key] = normalizer.normalize_data(loaded_data)
            elif isinstance(loaded_data, dict):
                all_loaded_data[file_key] = {
                    sheet: normalizer.normalize_data(df)
                    for sheet, df in loaded_data.items()
                }
        
        # Build configuration from frontend data
        config_builder = ConfigBuilder()
        slides_yaml = config_builder.build_slides_config(slides_config)
        
        # Save temporary config
        config_id = str(uuid.uuid4())
        temp_config_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{config_id}_slides.yaml")
        with open(temp_config_path, 'w') as f:
            import yaml
            yaml.dump(slides_yaml, f)
        
        # Generate PPT
        generator = PPTGenerator(
            template_path=template_path,
            slides_config=temp_config_path
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

