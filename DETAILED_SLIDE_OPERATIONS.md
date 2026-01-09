# Detailed Slide Operations - File by File, Slide by Slide

## Overview

This document provides a **detailed, step-by-step** breakdown of how each slide in the PowerPoint deck is created from Excel files. It shows exactly what operations are performed on which files and how data is transformed.

---

## Excel File Structure

### Primary File: `AIL LT Working file.xlsx`

**Location:** `Data/Apr 2025/AIL LT Working file.xlsx`

**Sheets:**
1. INPUT DISTRIBUTION STATUS
2. Input distribution  
3. WITHOLD EXPENSE
4. Input with hold
5. consent
6. Chronic & Overcalling
7. WORKING 3 (reference data)
8. FREQUENT DEFAULTERS 2 (reference data)
9. WITHHOLD COMMON EMPLOYEES 1 (reference data)
10. Overlapping drs
11. Sheet2
12. CLT

---

## SLIDE 1: Title Slide

### Content
- **Title:** "AIL LT April'25"
- **Subtitle:** "Commercial insights DASHBOARD BUSINESS EFFECTIVENESS"

### Data Source
- **None** - Static text content
- Month/year could be extracted from file name or system date

### Operations
1. **No Excel operations** - Static text
2. **Text formatting:** Apply title/subtitle styles
3. **Date extraction (optional):** Extract "April'25" from file name or date

### Transformation
```
Static Text → PPT Title/Subtitle
"AIL LT" + "April'25" → Title placeholder
"Commercial insights DASHBOARD" → Subtitle placeholder
```

---

## SLIDE 2: Business Effectiveness

### Content
- **Title:** "Business effectiveness"
- **Chart/Visualization:** Embedded think-cell chart object

### Data Source
- **Multiple Excel sheets** (likely aggregated)
- Possibly from `Input distribution` or calculated from multiple sources

### Operations
1. **Data aggregation** from multiple sheets
2. **Chart generation** (external tool - think-cell)
3. **No direct table mapping** - chart is pre-generated

### Transformation
```
Multiple Excel Sheets → Aggregated Data → Chart → PPT Chart Object
```

---

## SLIDE 3: PROJECT SHIELD - Input Distribution Status

### Content
- **Title:** "PROJECT SHIELD"
- **Subtitle:** "Data as on 4th Apr'25"
- **Table:** Division names and Distribution percentages

### Excel File & Sheet
- **File:** `AIL LT Working file.xlsx`
- **Sheet:** `INPUT DISTRIBUTION STATUS`
- **Header Row:** Row 2 (0-indexed: row index 2)

### Excel Structure
```
Row 0: [Empty/NaN]
Row 1: [Empty/NaN]  
Row 2: ["SLIDE 3" (or "Division"), " Total Dis "]
Row 3: ["GI Prima", 97.30524610779793]
Row 4: ["GI Maxima", 95.93341954405594]
Row 5: ["Metabolics", 94.28066312714758]
...
```

### Operations Performed

#### Step 1: Read Excel File
```python
df = pd.read_excel('AIL LT Working file.xlsx', 
                   sheet_name='INPUT DISTRIBUTION STATUS', 
                   header=None)
```

#### Step 2: Identify Header Row
- Header is on **row 2** (0-indexed)
- Columns: `SLIDE 3` (or `Division`) and `Unnamed: 1` (or ` Total Dis `)

#### Step 3: Re-read with Correct Header
```python
df = pd.read_excel('AIL LT Working file.xlsx', 
                   sheet_name='INPUT DISTRIBUTION STATUS', 
                   header=2)  # Header on row 2
```

#### Step 4: Clean Data
- Remove rows where Division is NaN/empty
- Remove rows where percentage is NaN
- Result: ~12-13 data rows

#### Step 5: Select Columns
- Column 1: Division names (e.g., "GI Prima", "GI Maxima")
- Column 2: Distribution percentages (e.g., 97.31, 95.93)

#### Step 6: Format Percentages
- Convert decimal to percentage format
- Example: `97.30524610779793` → `97.31%`
- Round to 2 decimal places

#### Step 7: Limit Rows
- Show top 12-13 divisions (or all if less than 15)

#### Step 8: Create PPT Table
- Create table with 2 columns: Division | Distribution %
- First row: Headers ("Division", "Total Distribution %")
- Subsequent rows: Data rows
- Apply formatting: Bold headers, percentage format for second column

### Transformation Flow
```
Excel File (AIL LT Working file.xlsx)
    ↓
Sheet: INPUT DISTRIBUTION STATUS
    ↓
Read with header=2
    ↓
Filter: Remove empty rows
    ↓
Select: Division column, Total Dis column
    ↓
Transform: Format percentages (decimal → %)
    ↓
Limit: Top 12-13 rows
    ↓
PPT Table: 2 columns, formatted
```

### Example Data Flow
```
Excel Row 3: ["GI Prima", 97.30524610779793]
    ↓
After filtering: ["GI Prima", 97.30524610779793]
    ↓
After formatting: ["GI Prima", "97.31%"]
    ↓
PPT Table Row: | GI Prima | 97.31% |
```

---

## SLIDE 4: PROJECT REACH - Consent Status

### Content
- **Title:** "PROJECT REACH - CONSENT STATUS"
- **Subtitle:** "43% HCP require consent for digital engagement"
- **Table:** Division, DVL, # HCP Consent, % Consent Require

### Excel File & Sheet
- **File:** `AIL LT Working file.xlsx`
- **Sheet:** `consent`
- **Header Row:** Row 0 (standard header)

### Excel Structure
```
Row 0: ["Division Name", "DVL", "# HCP Consent", "% Consent Require"]
Row 1: ["GI Prima", 1250, 850, 0.68]
Row 2: ["GI Maxima", 980, 720, 0.73]
Row 3: ["Metabolics", 1100, 900, 0.82]
...
```

### Operations Performed

#### Step 1: Read Excel File
```python
df = pd.read_excel('AIL LT Working file.xlsx', 
                   sheet_name='consent', 
                   header=0)
```

#### Step 2: Select Columns
- `Division Name` → Column 1
- `DVL` → Column 2 (number)
- `# HCP Consent` → Column 3 (number)
- `% Consent Require` → Column 4 (percentage)

#### Step 3: Format Data
- **DVL:** Display as integer (no decimals)
- **# HCP Consent:** Display as integer
- **% Consent Require:** Convert decimal to percentage
  - Example: `0.68` → `68%`
  - Example: `0.73` → `73%`

#### Step 4: Filter/Limit
- Show top 10 divisions (or all if less)

#### Step 5: Create PPT Table
- 4 columns: Division | DVL | # HCP Consent | % Consent Require
- Format: Numbers as integers, percentages as percentages

### Transformation Flow
```
Excel File → Sheet: consent → Read header=0
    ↓
Select 4 columns
    ↓
Format: Numbers (int), Percentages (decimal → %)
    ↓
Limit: Top 10 rows
    ↓
PPT Table: 4 columns, formatted
```

### Example Data Flow
```
Excel Row 1: ["GI Prima", 1250, 850, 0.68]
    ↓
After formatting: ["GI Prima", 1250, 850, "68%"]
    ↓
PPT Table Row: | GI Prima | 1250 | 850 | 68% |
```

---

## SLIDE 5: Input Distribution Summary

### Content
- **Title:** "Input Distribution"
- **Text/Summary:** Shows aggregated percentage (e.g., "93% YTD Distribution")

### Excel File & Sheet
- **File:** `AIL LT Working file.xlsx`
- **Sheet:** `Input distribution`
- **Header Row:** Row 0 (but has empty rows at top)

### Excel Structure
```
Row 0: [NaN, NaN, "Division", "YTD Distribution % ", "SPL Distribution % "]
Row 1: [NaN, NaN, "GenNext", 85.59, 93.87]
Row 2: [NaN, NaN, "GI Maxima", 75.65, 92.23]
Row 3: [NaN, NaN, "GI Optima", 92.92, 91.02]
...
```

### Operations Performed

#### Step 1: Read Excel File
```python
df = pd.read_excel('AIL LT Working file.xlsx', 
                   sheet_name='Input distribution', 
                   header=None)
```

#### Step 2: Find Actual Header
- Header text is in row 0, columns 2-4
- Actual data starts from row 1

#### Step 3: Re-read with Correct Structure
```python
df = pd.read_excel('AIL LT Working file.xlsx', 
                   sheet_name='Input distribution', 
                   header=0)
# Then select columns: Unnamed: 2, Unnamed: 3, Unnamed: 4
```

#### Step 4: Select Percentage Column
- Use `Unnamed: 3` (YTD Distribution %) or `Unnamed: 4` (SPL Distribution %)

#### Step 5: Calculate Aggregate
- **Option A:** Average of all percentages
  ```python
  average = df['Unnamed: 3'].mean()  # e.g., 93.45
  ```
- **Option B:** Weighted average
- **Option C:** Sum (if appropriate)

#### Step 6: Format Result
- Round to whole number or 1 decimal
- Example: `93.45` → `93%` or `93.5%`

#### Step 7: Generate Text
- Create text: "93% YTD Distribution" or similar
- Insert into PPT text box

### Transformation Flow
```
Excel File → Sheet: Input distribution
    ↓
Select percentage column (YTD or SPL)
    ↓
Calculate: Average/Sum of percentages
    ↓
Format: Round to percentage
    ↓
Generate: Text string with percentage
    ↓
PPT Text Box: Display aggregated percentage
```

### Example Calculation
```
Excel Data:
  GenNext: 85.59%
  GI Maxima: 75.65%
  GI Optima: 92.92%
  GI Prima: 96.16%
  ...
  
Calculation:
  Average = (85.59 + 75.65 + 92.92 + 96.16 + ...) / N
  Result = 93.45%
  
PPT Display:
  "93% YTD Distribution"
```

---

## SLIDE 6: Input with Hold

### Content
- **Title:** "Input with Hold"
- **Table:** Division, Active HC, Monthly columns (Nov'23 - Apr'24), %, Common TBMs

### Excel File & Sheet
- **File:** `AIL LT Working file.xlsx`
- **Sheet:** `Input with hold`
- **Header Row:** Row 2

### Excel Structure
```
Row 0: [Empty]
Row 1: [Empty]
Row 2: ["Divisions", "Active HC", "Nov'23", "Dec'23", ..., "Apr'24", "%", "Common TBMs"]
Row 3: ["GI Prima", 125, 95, 98, ..., 102, 0.85, "TBM1, TBM2"]
Row 4: ["GI Maxima", 110, 88, 92, ..., 96, 0.87, "TBM3, TBM4"]
...
```

### Operations Performed

#### Step 1: Read Excel File
```python
df = pd.read_excel('AIL LT Working file.xlsx', 
                   sheet_name='Input with hold', 
                   header=2)
```

#### Step 2: Select All Columns
- Division names
- Active HC (integer)
- Monthly columns (Nov'23 through Apr'24) - 6 columns
- Percentage column
- Common TBMs (text)

#### Step 3: Format Monthly Columns
- Display as integers (counts)
- Keep month labels (Nov'23, Dec'23, etc.)

#### Step 4: Format Percentage
- Convert decimal to percentage
- Example: `0.85` → `85%`

#### Step 5: Format Common TBMs
- Keep as text (comma-separated names)

#### Step 6: Create PPT Table
- Multiple columns (8-10 columns total)
- Format: Numbers as integers, percentage as percentage

### Transformation Flow
```
Excel File → Sheet: Input with hold → header=2
    ↓
Select all columns (Division, Active HC, Months, %, TBMs)
    ↓
Format: Numbers (int), Percentage (decimal → %), Text (as-is)
    ↓
Limit: Top 23 rows
    ↓
PPT Table: Multiple columns, formatted
```

---

## SLIDE 7: Overlapping Doctors

### Content
- **Title:** "Overlapping Doctors"
- **Table:** Doctor names, territories, overlap information

### Excel File & Sheet
- **File:** `AIL LT Working file.xlsx`
- **Sheet:** `Overlapping drs`
- OR from Reports folder: `AIL - Same Doctor in Multiple TBM Status Mar'25.xlsb`

### Operations Performed

#### Step 1: Read Excel File
```python
# Option 1: From main file
df = pd.read_excel('AIL LT Working file.xlsx', 
                   sheet_name='Overlapping drs')

# Option 2: From Reports folder
df = pd.read_excel('Data/Apr 2025/Reports/AIL - Same Doctor in Multiple TBM Status Mar\'25.xlsb')
```

#### Step 2: Filter Overlapping Doctors
- Identify doctors appearing in multiple territories
- Filter rows where territory count > 1

#### Step 3: Select Columns
- Doctor name
- Territories (list)
- Territory count
- Other relevant columns

#### Step 4: Format Data
- Doctor names as text
- Territory lists formatted
- Counts as integers

#### Step 5: Create PPT Table
- Display overlapping doctors with their territories

---

## SLIDE 8: Withhold Expense

### Content
- **Title:** "Withhold Expense"
- **Table:** Division and expense status data

### Excel File & Sheet
- **File:** `AIL LT Working file.xlsx`
- **Sheet:** `WITHOLD EXPENSE`
- **Header Row:** Row 4

### Operations Performed

#### Step 1: Read Excel File
```python
df = pd.read_excel('AIL LT Working file.xlsx', 
                   sheet_name='WITHOLD EXPENSE', 
                   header=4)
```

#### Step 2: Select Columns
- Division names (`Slide 5` column)
- Status columns (`Unnamed: 1-6`)

#### Step 3: Format Data
- Division names as text
- Status indicators formatted appropriately

#### Step 4: Create PPT Table
- Display division and expense status

---

## SLIDE 9: Chronic & Overcalling

### Content
- **Title:** "Chronic & Overcalling"
- **Table:** Division, #DVL, #Missed, % Missed

### Excel File & Sheet
- **File:** `AIL LT Working file.xlsx`
- **Sheet:** `Chronic & Overcalling`
- **Header Row:** Row 4

### Operations Performed

#### Step 1: Read Excel File
```python
df = pd.read_excel('AIL LT Working file.xlsx', 
                   sheet_name='Chronic & Overcalling', 
                   header=4)
```

#### Step 2: Select Columns
- `Slide 9` - Division names
- `Unnamed: 1` - #DVL (count)
- `Unnamed: 2` - #Missed (count)
- `Unnamed: 3` - % Missed (decimal)

#### Step 3: Format Data
- **#DVL:** Integer
- **#Missed:** Integer
- **% Missed:** Convert decimal to percentage
  - Example: `0.15` → `15%`

#### Step 4: Create PPT Table
- 4 columns: Division | #DVL | #Missed | % Missed
- Format: Numbers as integers, percentage as percentage

### Transformation Flow
```
Excel File → Sheet: Chronic & Overcalling → header=4
    ↓
Select 4 columns
    ↓
Format: Counts (int), Percentage (decimal → %)
    ↓
PPT Table: 4 columns, formatted
```

---

## SLIDE 10: Consent Status (Detailed)

### Content
- **Title:** "Consent Status" or similar
- **Table:** Detailed consent information

### Excel File & Sheet
- **File:** `AIL LT Working file.xlsx`
- **Sheet:** `consent` (same as Slide 4, possibly different view)
- OR from Reports: `AIL Consented Status HCP's_02.04.2025.xlsb`

### Operations
- Similar to Slide 4, but may show different columns or filtered view

---

## SLIDE 11: Blank/End Slide

### Content
- Empty slide

### Operations
- None

---

## Summary of Common Operations

### 1. File Reading
```python
df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
```

### 2. Data Cleaning
- Remove empty rows: `df.dropna(axis=0, how='all')`
- Remove empty columns: `df.dropna(axis=1, how='all')`
- Filter by condition: `df[df['column'] != value]`

### 3. Column Selection
- By name: `df[['Column1', 'Column2']]`
- By position: `df.iloc[:, [0, 1, 2]]`

### 4. Data Transformation
- **Percentage:** `df['col'] * 100` then format as `"{:.2f}%".format(value)`
- **Integer:** `df['col'].astype(int)`
- **Round:** `df['col'].round(2)`

### 5. Aggregations
- **Average:** `df['col'].mean()`
- **Sum:** `df['col'].sum()`
- **Count:** `df['col'].count()`

### 6. Row Limiting
- **Top N:** `df.head(n)`
- **Filtered:** `df[condition].head(n)`

### 7. PPT Table Creation
- Map DataFrame rows to PPT table rows
- Apply formatting (bold headers, number formats)
- Set column widths and alignment

---

## Key Patterns Identified

1. **Header Row Detection:** Critical - varies from 0 to 4
2. **Column Naming:** Many "Unnamed: X" columns - need positional mapping
3. **Percentage Handling:** Always convert decimals to percentages
4. **Row Filtering:** Most slides limit to 10-23 rows
5. **Multi-sheet Usage:** Some slides may use data from multiple sheets
6. **Calculations:** Some slides require aggregations (Slide 5)

---

This analysis provides the foundation for automating the deck creation process.

