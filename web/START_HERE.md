# 🚀 Web Frontend - Quick Start Guide

## ✅ What's Been Built

A complete web-based interface for generating PowerPoint presentations from Excel data!

## 📋 Setup Instructions

### 1. Install Dependencies

```bash
cd web
pip install -r requirements.txt
```

### 2. Start the Server

```bash
python app.py
```

The server will start on `http://localhost:5000`

### 3. Open in Browser

Navigate to: **http://localhost:5000**

## 🎯 How to Use

1. **Upload Excel File**
   - Click "Choose File" or drag & drop
   - System will analyze the file structure

2. **Select Template** (optional)
   - Choose from available templates
   - Or use default template

3. **Configure Slides**
   - Click "+ Add Slide" to add slides
   - For each slide:
     - Choose slide type (Title, Table, Content)
     - Set title
     - Select data source (sheet)
     - Choose columns for tables
     - Set header row if needed

4. **Generate & Download**
   - Click "Generate PowerPoint Deck"
   - Wait for generation to complete
   - Click "Download PowerPoint"

## 📁 File Structure

```
web/
├── app.py              # Flask backend (main server)
├── config_builder.py   # Converts form data to YAML
├── requirements.txt    # Python dependencies
├── templates/
│   └── index.html      # Frontend UI
├── static/
│   ├── css/
│   │   └── style.css   # Styling
│   └── js/
│       └── app.js      # Frontend logic
├── uploads/            # Uploaded Excel files (auto-created)
└── output/            # Generated PPTs (auto-created)
```

## 🔧 API Endpoints

- `GET /` - Main interface
- `POST /api/analyze-excel` - Analyze Excel file
- `GET /api/excel-sheets` - Get sheets from file
- `GET /api/excel-columns` - Get columns from sheet
- `GET /api/templates` - List templates
- `POST /api/generate-ppt` - Generate PowerPoint
- `GET /api/download/<id>` - Download generated PPT

## ⚠️ Important Notes

1. **Empty Slides Issue**: The frontend is built, but the underlying empty slides issue still needs to be fixed separately. The frontend will work, but generated PPTs may be empty until the data mapping is fixed.

2. **Template Path**: Make sure `templates/template.pptx` exists in the root directory.

3. **File Storage**: Uploaded files are stored in `web/uploads/` and generated PPTs in `web/output/`. These directories are created automatically.

## 🐛 Troubleshooting

**Server won't start:**
- Check if port 5000 is available
- Install all dependencies: `pip install -r requirements.txt`

**Can't upload files:**
- Check file size (max 100MB)
- Ensure file is .xlsx, .xlsb, or .xls format

**Generation fails:**
- Check console for error messages
- Verify template file exists
- Ensure Excel file structure is valid

## 🎨 Features

- ✅ Modern, responsive UI
- ✅ Drag & drop file upload
- ✅ Dynamic slide configuration
- ✅ Column selection for tables
- ✅ Real-time validation
- ✅ Progress indicators
- ✅ Download generated files

## 📝 Next Steps

1. Test the frontend with your Excel files
2. Fix the empty slides issue (data mapping)
3. Customize styling if needed
4. Add more features as required

---

**Ready to use! Start the server and open http://localhost:5000**

