# PPT Automation Web Frontend

Web-based interface for generating PowerPoint presentations from Excel data.

## Setup

1. **Install Dependencies:**
   ```bash
   cd web
   pip install -r requirements.txt
   ```

2. **Run the Server:**
   ```bash
   python app.py
   ```

3. **Access the Interface:**
   Open your browser and go to: `http://localhost:5000`

## Features

- **Excel File Upload**: Upload and analyze Excel files (.xlsx, .xlsb, .xls)
- **Template Selection**: Choose from available PowerPoint templates
- **Slide Configuration**: 
  - Add/remove slides
  - Configure slide types (title, table, content)
  - Map data sources to slides
  - Select columns for tables
- **Generate & Download**: Generate PowerPoint deck and download it

## Usage

1. Upload an Excel file
2. Select a template (optional)
3. Configure slides:
   - Add slides as needed
   - For each slide, select:
     - Slide type
     - Title
     - Data source (sheet)
     - Columns to include
4. Click "Generate PowerPoint Deck"
5. Download the generated file

## API Endpoints

- `POST /api/analyze-excel` - Analyze uploaded Excel file
- `GET /api/excel-sheets` - Get sheets from Excel file
- `GET /api/excel-columns` - Get columns from specific sheet
- `GET /api/templates` - List available templates
- `POST /api/generate-ppt` - Generate PowerPoint deck
- `GET /api/download/<output_id>` - Download generated PPT

## File Structure

```
web/
├── app.py              # Flask backend server
├── config_builder.py   # Configuration builder
├── requirements.txt    # Python dependencies
├── templates/
│   └── index.html      # Frontend HTML
└── static/
    ├── css/
    │   └── style.css   # Styling
    └── js/
        └── app.js      # Frontend JavaScript
```

## Notes

- Uploaded files are stored in `web/uploads/`
- Generated PPTs are stored in `web/output/`
- The frontend builds YAML configuration from user selections
- Configuration is passed to the existing PPT generation pipeline

