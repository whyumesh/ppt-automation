# Demo Guide - PPT Automation Tool

## Sample Data Files

Three sample Excel files have been created in `web/sample_data/`:

1. **Sales_Data.xlsx** - Product sales information
   - Columns: Product, Sales, Growth %, Region
   - 5 rows of data

2. **Marketing_Data.xlsx** - Marketing campaign data
   - Columns: Campaign, Budget, ROI %, Status
   - 4 rows of data

3. **Financial_Data.xlsx** - Monthly financial summary
   - Columns: Month, Revenue, Expenses, Profit
   - 5 rows of data

## Demo Flow

1. **Start the server:**
   ```bash
   cd web
   python app.py
   ```

2. **Open browser:** http://localhost:5000

3. **Step 1:** Select number of slides (e.g., 3)

4. **Step 2:** For each slide:
   - Click "Click to upload Excel file"
   - Navigate to `web/sample_data/` folder
   - Upload one of the sample files
   - See data preview appear
   - Configure:
     - Select sheet
     - Choose columns
     - Set title

5. **Step 3:** Select template

6. **Step 4:** Generate and download

## Demo Script

**For Slide 1:**
- Upload: `Sales_Data.xlsx`
- Title: "Q1 Sales Performance"
- Select columns: Product, Sales, Growth %

**For Slide 2:**
- Upload: `Marketing_Data.xlsx`
- Title: "Marketing Campaigns"
- Select columns: Campaign, Budget, ROI %

**For Slide 3:**
- Upload: `Financial_Data.xlsx`
- Title: "Financial Overview"
- Select columns: Month, Revenue, Profit

## UI Features to Highlight

- ✨ Modern, vibrant gradient design
- 📊 Real-time data preview
- 🎯 Step-by-step workflow
- 📁 Per-slide file upload
- ✅ Visual status indicators
- 🚀 Smooth animations

## Tips for Demo

1. Show the data preview feature - it's impressive!
2. Demonstrate uploading different files for different slides
3. Show the column selection interface
4. Highlight the modern UI design
5. Show the generation process

