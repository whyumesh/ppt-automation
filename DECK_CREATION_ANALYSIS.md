# Detailed Deck Creation Analysis
## How PowerPoint Decks are Created - File by File, Slide by Slide

This document analyzes how the April 2025 PowerPoint deck is created from Excel files.

---

## Excel Files Used

### Primary File: `AIL LT Working file.xlsx`

This is the main data source with 12 sheets:

1. **INPUT DISTRIBUTION STATUS** - Used for Slide 3
2. **Input distribution** - Used for Slide 3, 5
3. **WITHOLD EXPENSE** - Used for Slide 8
4. **Input with hold** - Used for Slide 6
5. **consent** - Used for Slide 4, 10
6. **Chronic & Overcalling** - Used for Slide 9
7. **WORKING 3** - Reference data (540 rows)
8. **FREQUENT DEFAULTERS 2** - Reference data (544 rows)
9. **WITHHOLD COMMON EMPLOYEES 1** - Reference data (709 rows)
10. **Overlapping drs** - Used for Slide 7
11. **Sheet2** - Reference data
12. **CLT** - Alternative data for Slide 3

### Additional Files in Reports Folder

- `AIL - Speaker Allocation Summary - 04-April-25.xlsx`
- `AIL Consented Status HCP's_02.04.2025.xlsb`
- `AIL Input Distribution Apr'24 - Mar'25 01042025.xlsb`
- `Chronic Missing Report AIL - Jan to Mar.xlsx`
- `Overcalling Report AIL - Jan to Mar.xlsx`
- And others...

---

## Slide-by-Slide Analysis

### SLIDE 1: Title Slide
**Layout:** Title Slide Layout

**Content:**
- Title: "AIL LT April'25"
- Subtitle: "Commercial insights DASHBOARD BUSINESS EFFECTIVENESS"

**Data Source:** 
- **Static text** - No Excel data required
- Month/year ("April'25") could be extracted from file name or date

**Operations:**
- None - Static content

---

### SLIDE 2: Business Effectiveness
**Layout:** Title Only

**Content:**
- Title: "Business effectiveness"
- Contains chart/visualization (think-cell embedded object)

**Data Source:**
- **Multiple Excel sheets** - Likely aggregated data from various sources
- May use data from multiple sheets for chart visualization

**Operations:**
- Data aggregation
- Chart generation (external tool - think-cell)

---

### SLIDE 3: PROJECT SHIELD - Input Distribution Status
**Layout:** Title Slide

**Content:**
- Title: "PROJECT SHIELD"
- Subtitle: "Data as on 4th Apr'25"
- **Table** with Division and Distribution percentages

**Data Source:**
- **Excel Sheet:** `INPUT DISTRIBUTION STATUS`
- **Header Row:** Row 2 (0-indexed: 2)
- **Columns Used:**
  - `SLIDE 3` (or `Division`) - Division names
  - `Unnamed: 1` (or ` Total Dis `) - Distribution percentages

**Operations:**
1. **Read Excel:** Load `INPUT DISTRIBUTION STATUS` sheet
2. **Skip Header:** Start from row 2 (header is on row 2)
3. **Extract Columns:** Get Division names and percentages
4. **Filter:** Remove empty rows
5. **Format:** 
   - Percentages displayed as percentages (e.g., 97.31%)
   - Division names as text
6. **Limit Rows:** Show top ~13 divisions

**Transformation:**
```
Excel Row → PPT Table Row
Division Name → First Column
Percentage Value → Second Column (formatted as %)
```

**Example:**
- Excel: `GI Prima` | `97.30524610779793`
- PPT: `GI Prima` | `97.31%`

---

### SLIDE 4: PROJECT REACH - Consent Status
**Layout:** Blank

**Content:**
- Title: "PROJECT REACH - CONSENT STATUS"
- Subtitle: "43% HCP require consent for digital engagement"
- **Table** with Division, DVL, Consent metrics

**Data Source:**
- **Excel Sheet:** `consent`
- **Header Row:** Row 0
- **Columns Used:**
  - `Division Name` - Division names
  - `DVL` - Doctor Visit List count
  - `# HCP Consent` - Number of HCPs with consent
  - `% Consent Require` - Percentage requiring consent

**Operations:**
1. **Read Excel:** Load `consent` sheet
2. **Extract Columns:** Get all 4 columns
3. **Format:**
   - Numbers formatted appropriately
   - Percentages shown as percentages
4. **Limit Rows:** Show ~10 divisions

**Transformation:**
```
Excel → PPT Table
Division Name → Column 1
DVL → Column 2 (number)
# HCP Consent → Column 3 (number)
% Consent Require → Column 4 (percentage)
```

---

### SLIDE 5: Input Distribution Summary
**Layout:** Title Only

**Content:**
- Title: "Input Distribution"
- **Text/Summary:** Shows aggregated percentage (e.g., "93%")

**Data Source:**
- **Excel Sheet:** `Input distribution`
- **Header Row:** Row 0 (but has empty rows at top)
- **Columns Used:**
  - `Unnamed: 2` - Division names
  - `Unnamed: 3` - YTD Distribution %
  - `Unnamed: 4` - SPL Distribution %

**Operations:**
1. **Read Excel:** Load `Input distribution` sheet
2. **Calculate Aggregate:** 
   - Sum or average of YTD Distribution % or SPL Distribution %
   - Result: ~93%
3. **Format:** Display as percentage with text

**Transformation:**
```
Excel Data → Calculation → PPT Text
Multiple rows with percentages → Average/Sum → "93% YTD Distribution"
```

---

### SLIDE 6: Input with Hold
**Layout:** Title Only

**Content:**
- Title: "Input with Hold"
- **Table** with Division, Active HC, Monthly data, Percentage, Common TBMs

**Data Source:**
- **Excel Sheet:** `Input with hold`
- **Header Row:** Row 2
- **Columns Used:**
  - `Divisions` - Division names
  - `Active HC` - Active Healthcare count
  - `Nov'23` through `Apr'24` - Monthly data columns
  - `%` - Percentage column
  - `Common TBMs` - Common Territory Business Managers

**Operations:**
1. **Read Excel:** Load `Input with hold` sheet
2. **Skip Header:** Start from row 2
3. **Extract Columns:** Get all specified columns
4. **Format Dates:** Monthly columns (Nov'23, Dec'23, etc.)
5. **Format Numbers:** Percentages and counts
6. **Limit Rows:** Show ~23 divisions

**Transformation:**
```
Excel → PPT Table
Multiple monthly columns → Displayed as separate columns
Percentage calculated/displayed
```

---

### SLIDE 7: Overlapping Doctors
**Layout:** Title Only

**Content:**
- Title: "Overlapping Doctors"
- **Table** with overlapping doctor information

**Data Source:**
- **Excel Sheet:** `Overlapping drs` (or from Reports folder)
- May also use data from `AIL - Same Doctor in Multiple TBM Status Mar'25.xlsb`

**Operations:**
1. **Read Excel:** Load overlapping doctors data
2. **Filter:** Identify doctors appearing in multiple territories
3. **Format:** Display doctor names, territories, counts

---

### SLIDE 8: Withhold Expense
**Layout:** Title Only

**Content:**
- Title: "Withhold Expense"
- **Table** with Division and expense status data

**Data Source:**
- **Excel Sheet:** `WITHOLD EXPENSE`
- **Header Row:** Row 4
- **Columns Used:**
  - `Slide 5` - Division names
  - `Unnamed: 1-6` - May/June status data columns

**Operations:**
1. **Read Excel:** Load `WITHOLD EXPENSE` sheet
2. **Skip Header:** Start from row 4
3. **Extract Columns:** Get division and status columns
4. **Format:** Display expense status information
5. **Limit Rows:** Show ~16 divisions

---

### SLIDE 9: Chronic & Overcalling
**Layout:** Title Only

**Content:**
- Title: "Chronic & Overcalling"
- **Table** with Division, DVL count, Missed count, % Missed

**Data Source:**
- **Excel Sheet:** `Chronic & Overcalling`
- **Header Row:** Row 4
- **Columns Used:**
  - `Slide 9` - Division names
  - `Unnamed: 1` - #DVL (Doctor Visit List count)
  - `Unnamed: 2` - #Missed (Number missed)
  - `Unnamed: 3` - % Missed (as decimal, needs conversion)

**Operations:**
1. **Read Excel:** Load `Chronic & Overcalling` sheet
2. **Skip Header:** Start from row 4
3. **Calculate:** May calculate % Missed from counts
4. **Format:**
   - Numbers as integers
   - Percentages formatted as percentages
5. **Limit Rows:** Show ~16 divisions

**Transformation:**
```
Excel → PPT
Decimal (0.15) → Percentage (15%)
Counts displayed as numbers
```

---

### SLIDE 10: Consent Status (Detailed)
**Layout:** Title Only

**Content:**
- Title: "Consent Status" or similar
- **Table** with detailed consent information

**Data Source:**
- **Excel Sheet:** `consent` (same as Slide 4, possibly different view)
- Or from `AIL Consented Status HCP's_02.04.2025.xlsb`

**Operations:**
1. **Read Excel:** Load consent data
2. **Filter/Transform:** May show different columns or filtered view
3. **Format:** Display consent metrics

---

### SLIDE 11: Blank/End Slide
**Layout:** Blank

**Content:**
- Empty slide (end of presentation)

**Data Source:**
- None

**Operations:**
- None

---

## Common Operations Across Slides

### 1. **Data Reading**
- Load Excel file: `AIL LT Working file.xlsx`
- Select appropriate sheet
- Identify header row (varies: 0, 2, 4)

### 2. **Data Filtering**
- Remove empty rows
- Filter by division if needed
- Limit number of rows displayed

### 3. **Data Transformation**
- **Percentage Formatting:** Convert decimals to percentages (0.15 → 15%)
- **Number Formatting:** Round to appropriate decimal places
- **Date Formatting:** Format monthly columns (Nov'23, etc.)
- **Text Cleaning:** Remove extra spaces, normalize names

### 4. **Calculations**
- **Aggregations:** Sum, average, count
- **Percentages:** Calculate from counts
- **Rankings:** Sort by values

### 5. **Table Generation**
- Map Excel columns to PPT table columns
- Apply formatting (bold headers, colors)
- Set column widths
- Apply number formatting

### 6. **Text Generation**
- Extract titles from Excel sheet names or config
- Generate subtitles with dates
- Create summary text from calculations

---

## Data Flow Summary

```
Excel File (AIL LT Working file.xlsx)
    ↓
Select Sheet (e.g., "INPUT DISTRIBUTION STATUS")
    ↓
Read Data (with header row offset)
    ↓
Filter/Transform Data
    - Remove empty rows
    - Select columns
    - Calculate percentages
    - Format numbers
    ↓
Map to PPT Structure
    - Create table
    - Set headers
    - Populate rows
    ↓
Apply Formatting
    - Fonts, colors, sizes
    - Number formats
    - Alignment
    ↓
Insert into Slide
    - Add to specific slide
    - Position correctly
    ↓
Final PPT Deck
```

---

## Key Insights

1. **Single Primary File:** Most data comes from `AIL LT Working file.xlsx`
2. **Multiple Sheets:** Different sheets feed different slides
3. **Header Row Variations:** Header rows vary (0, 2, 4) - need to detect
4. **Column Naming:** Many columns are "Unnamed: X" - need positional mapping
5. **Percentage Handling:** Percentages stored as decimals, need conversion
6. **Row Limiting:** Most slides show limited rows (10-23)
7. **Static Content:** Some slides (title, end) have no Excel data
8. **Calculations:** Some slides require aggregations/calculations

---

## Next Steps for Automation

1. **Map each slide** to its Excel sheet and columns
2. **Identify header rows** for each sheet
3. **Document transformations** (percentage conversion, formatting)
4. **Create configuration** for each slide's data mapping
5. **Implement filters** and row limits
6. **Handle calculations** (aggregations, percentages)

