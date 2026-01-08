# Updated Frontend Flow

## New Workflow

The frontend has been updated to match your requirements:

### Step 1: Select Number of Slides
- User enters how many slides they need (1-50)
- This determines how many slide configurations will be shown

### Step 2: Upload Excel Files
- User can upload **multiple Excel files**
- All uploaded files are stored and listed
- User can remove files if needed
- Files are analyzed to show their structure (sheets, columns)

### Step 3: Select Template
- Choose from available PowerPoint templates
- Or use default template

### Step 4: Configure Each Slide
- For **each slide**, the user:
  1. **Selects which Excel file** to use (dropdown of uploaded files)
  2. Chooses slide type (Title, Table, Content)
  3. Sets the title
  4. For tables:
     - Selects which sheet from the chosen file
     - Sets header row
     - Selects which columns to include

### Step 5: Generate & Download
- Generate the PowerPoint deck
- Download the generated file

## Key Changes

1. **Multiple File Support**: Users can upload multiple Excel files
2. **Per-Slide File Selection**: Each slide can use a different Excel file
3. **Fixed Slide Count**: Number of slides is set upfront, not added dynamically
4. **Better Organization**: Clear workflow from number of slides → files → template → configuration

## Example Usage

1. User wants 5 slides
2. Uploads 3 Excel files:
   - `Sales Data.xlsx`
   - `Marketing Data.xlsx`
   - `Financial Data.xlsx`
3. Selects template
4. Configures:
   - Slide 1: Uses `Sales Data.xlsx`, Sheet "Q1 Sales", Columns: [Product, Revenue]
   - Slide 2: Uses `Marketing Data.xlsx`, Sheet "Campaigns", Columns: [Campaign, ROI]
   - Slide 3: Uses `Financial Data.xlsx`, Sheet "Summary", Columns: [Month, Profit]
   - Slide 4: Uses `Sales Data.xlsx`, Sheet "Q2 Sales", Columns: [Product, Revenue]
   - Slide 5: Title slide (no data)
5. Generate and download

## Technical Details

- Frontend stores all uploaded files in `uploadedFiles` object
- Each slide has a `file_id` that references the uploaded file
- Backend loads all files and maps them by file name (without extension)
- Configuration builder creates YAML with proper `data_source` references

