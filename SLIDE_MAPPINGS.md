# Detailed Slide Mappings - Excel to PowerPoint

Based on analysis of April 2025 data, here are the precise mappings:

## Slide 1: Title Slide
**Layout:** Title Slide Layout
**Type:** Text
**Content:** 
- Title: "AIL LT"
- Subtitle: "April|25" (month/year)
- Subtitle: "Commercial insights DASHBOARD"

**Excel Source:** None (static title slide)
**Notes:** Month/year may need to be extracted from file name or date

---

## Slide 2: Business Effectiveness
**Layout:** Title Only
**Type:** Text/Bullet Points
**Content:** Business effectiveness metrics and percentages

**Excel Source:** Likely from multiple sheets (needs manual verification)
**Notes:** Contains percentage calculations

---

## Slide 3: Project FMV Status / Input Distribution
**Layout:** Title Slide
**Type:** Table
**Content:** 
- Title: "PROJECT SHIELD"
- Subtitle: "Data as on 4th Apr'25"
- Table with division data

**Excel Source:** 
- **Primary:** `Input distribution` sheet
  - Columns: `SLIDE - 3` (division names), `Unnamed: 2` (Division), `Unnamed: 3` (YTD Distribution %), `Unnamed: 4` (SPL Distribution %)
  - Table structure: ~11 rows x 5 columns
  - **Note:** First few rows are empty/headers, actual data starts around row 5
  - **Column Mapping:**
    - Column 1: Division names (from `Unnamed: 2`)
    - Column 2: YTD Distribution % (from `Unnamed: 3`) - percentages like 85.59%, 75.65%, etc.
    - Column 3: SPL Distribution % (from `Unnamed: 4`) - percentages like 93.87%, 92.23%, etc.

**Alternative Source:** `INPUT DISTRIBUTION STATUS` sheet
  - Columns: `SLIDE 3` (division names), `Unnamed: 1` (Total Dis % - percentages like 97.31%, 95.93%)

**Calculations:** 
- Percentages are already calculated in Excel
- Values appear to be formatted as percentages (e.g., 97.31% = 97.31)

**Formatting:**
- Title font: 24pt, color RGB(0, 176, 240)
- Table cells may have header row formatting

---

## Slide 4: PROJECT REACH - CONSENT STATUS
**Layout:** Blank
**Type:** Text + Table
**Content:**
- Title: "PROJECT REACH - CONSENT STATUS"
- Text: "43% HCP require consent for digital engagement"
- Table with consent data

**Excel Source:** 
- **Primary:** `consent` sheet
  - Columns: `Division Name`, `# of HCPs`, `Consent Received/Accepted`, `% Consent Require` (in `Unnamed: 9`)
  - **Column Mapping:**
    - Column 1: Division Name
    - Column 2: # of HCPs (or DVL)
    - Column 3: Consent Received/Accepted (or # HCP Consent)
    - Column 4: % Consent Require (from `Unnamed: 9` - values like 33.85%, 30.40%)

**Alternative Source:** `Input distribution` sheet (table structure matches)

**Calculations:**
- Percentage calculation: % Consent Require = (Consent Require / DVL) * 100
- Values in `Unnamed: 9` are already calculated percentages

**Formatting:**
- Title font: 24pt
- Table header row: Dark blue background RGB(0, 59, 85)
- Text font: 9pt

---

## Slide 5: Input Distribution Summary
**Layout:** Blank
**Type:** Text + Table
**Content:**
- Title: "YTD March'25 Input distribution is 93%"
- Table with distribution data

**Excel Source:**
- **Primary:** `Input distribution` sheet
  - Similar to Slide 3 but may show aggregated/summary data
- **Alternative:** `INPUT DISTRIBUTION STATUS` sheet

**Calculations:**
- Overall percentage calculation (93% aggregate)
- Individual division percentages

**Formatting:**
- Title font: 24pt
- Footer font: 10pt

---

## Slide 6: Input Distribution with Hold
**Layout:** Title Only
**Type:** Table
**Content:**
- Title: "Input Distribution with Hold"
- Large table with division data across months

**Excel Source:**
- **Primary:** `Input with hold` sheet
  - Columns: `Slide 10` (Divisions), `Unnamed: 1` (Active HC), `Unnamed: 2` through `Unnamed: 7` (months: Nov'23, Dec'23, Jan'24, Feb'24, Mar'24, Apr'24), `Unnamed: 9` (%), `Unnamed: 10` (Common TBMs)
  - Table structure: ~12 rows x 11 columns
  - **Column Mapping:**
    - Column 1: Divisions (from `Slide 10`)
    - Column 2: Active HC (from `Unnamed: 1`)
    - Columns 3-8: Monthly data (Nov'23 through Apr'24)
    - Column 9: % (from `Unnamed: 9` - appears to be 0 for most)
    - Column 10: Common TBMs in all 3 months (from `Unnamed: 10`)

**Calculations:**
- Percentage calculations for each division
- Common TBMs count

**Formatting:**
- Title font: 24pt, color RGB(0, 156, 222)
- Footer font: 10pt

---

## Slide 7: Input Distribution Status
**Layout:** Title Only
**Type:** Table
**Content:**
- Title: "Input Distribution Status"
- Table with division and distribution percentages

**Excel Source:**
- **Primary:** `INPUT DISTRIBUTION STATUS` sheet
  - Columns: `SLIDE 3` (Division), `Unnamed: 1` (Total Dis %)
  - Table structure: ~11 rows x 2 columns
  - **Column Mapping:**
    - Column 1: Division (from `SLIDE 3`)
    - Column 2: Total Dis % (from `Unnamed: 1` - percentages like 97.31%, 95.93%, 94.28%)

**Alternative Source:** `CLT` sheet
  - Similar structure: `SLIDE 3` (Division), `Unnamed: 1` (Total Dis %)

**Calculations:**
- Percentages are pre-calculated in Excel

**Formatting:**
- Title font: 24pt

---

## Slide 8: Withhold Expense
**Layout:** Title Slide
**Type:** Table
**Content:**
- Title: "Withhold Expense Status"
- Table with employee expense data

**Excel Source:**
- **Primary:** `WITHOLD EXPENSE` sheet
  - Columns: `Slide 5` (Division), `Unnamed: 1` (May Status header), `Unnamed: 2` (Allowed exception #), `Unnamed: 3` (Final # employees), `Unnamed: 4` (June Status header), `Unnamed: 5` (Allowed exception #), `Unnamed: 6` (Final # employees)
  - Table structure: ~12 rows x 7 columns
  - **Column Mapping:**
    - Column 1: Division (from `Slide 5`)
    - Columns 2-4: May Status data (Feb+Mar+Apr)
    - Columns 5-7: June Status data (Mar+Apr+May)
  - **Note:** First few rows contain headers like "May Status (Feb+Mar+Apr) data)", "#employees in the original withhold list", etc.

**Calculations:**
- Delta calculations (change between May and June status)
- Employee counts

**Formatting:**
- Title font: 26pt/24pt
- Text color: RGB(192, 0, 0) for certain values (red)
- Footer font: 10pt

---

## Slide 9: Chronic & Overcalling
**Layout:** Blank
**Type:** Table
**Content:**
- Title: "#HCPs Missed in last 3 Months"
- Table with chronic missed data

**Excel Source:**
- **Primary:** `Chronic & Overcalling` sheet
  - Columns: `Slide 9` (Division/headers), `Unnamed: 1` (#DVL), `Unnamed: 2` (#HCPs Missed), `Unnamed: 3` (% HCP Missed)
  - Table structure: ~13 rows x 4 columns
  - **Column Mapping:**
    - Column 1: Division (from `Slide 9`)
    - Column 2: #DVL (from `Unnamed: 1`)
    - Column 3: #HCPs Missed (from `Unnamed: 2`)
    - Column 4: % HCP Missed (from `Unnamed: 3` - decimal values like 0.0486 = 4.86%)

**Alternative Source:** `Sheet2` sheet
  - Similar structure for chronic missed trend

**Calculations:**
- Percentage: % HCP Missed = (#HCPs Missed / #DVL) * 100
- Values in `Unnamed: 3` are decimals (need to multiply by 100 for percentage)

**Formatting:**
- Title font: 20pt, color RGB(0, 176, 240)
- Table header row: Dark blue background RGB(0, 59, 85)
- Table footer row: Dark blue background RGB(0, 59, 85)

---

## Slide 10: Consent Status (Detailed)
**Layout:** Title Slide
**Type:** Text/Bullet Points
**Content:**
- Title: "PROJECT REACH - CONSENT STATUS"
- Detailed consent metrics

**Excel Source:**
- **Primary:** `consent` sheet
  - Similar to Slide 4 but more detailed
  - Columns: `Division Name`, `# of HCPs`, `Consent Received/Accepted`, `% Consent Require`

**Calculations:**
- Percentage calculations for consent rates

**Formatting:**
- Title font: 24pt, color RGB(0, 176, 240)
- Text font: 12pt/11pt

---

## Slide 11: (Need to verify - may be missing from analysis)
**Layout:** Unknown
**Type:** Unknown

---

## Key Findings:

### Column Name Patterns:
- Many columns are named "Unnamed: X" - these need to be identified by position and content
- Some sheets have explicit column names like "Division Name", "# of HCPs"
- Slide number indicators in column names (e.g., "SLIDE 3", "Slide 5", "Slide 10")

### Data Types:
- **Tables:** Slides 3, 4, 6, 7, 8, 9
- **Text/Bullets:** Slides 1, 2, 5, 10
- **Mixed:** Some slides have both text and tables

### Calculations Identified:
1. **Percentages:** 
   - Distribution percentages (already calculated in Excel)
   - Consent percentages (calculated or in `Unnamed: 9`)
   - HCP Missed percentages (decimal format, need * 100)

2. **Deltas/Changes:**
   - Month-over-month comparisons
   - Status changes (May vs June)

3. **Aggregations:**
   - Sums by division
   - Counts of employees/HCPs

### Formatting Rules:
- **Font Sizes:** 10pt (footer), 12pt (body), 14pt (subheadings), 20-24pt (titles), 26pt (large titles)
- **Colors:**
  - Title blue: RGB(0, 176, 240)
  - Dark blue header: RGB(0, 59, 85)
  - Red warning: RGB(192, 0, 0)
  - Text blue: RGB(0, 156, 222)
- **Table Headers:** Dark blue background RGB(0, 59, 85)
- **Table Footers:** Dark blue background RGB(0, 59, 85)

### Data Filtering:
- Most slides show filtered/aggregated data from larger Excel sheets
- Need to identify filtering criteria (e.g., top N, threshold-based, date ranges)

---

## Next Steps for Configuration:

1. **Identify exact row ranges** for each table (skip header rows, identify data start)
2. **Map "Unnamed" columns** to actual meaning by examining sample data
3. **Document filtering rules** (which rows/columns to include)
4. **Specify calculation formulas** for derived values
5. **Define formatting rules** per slide/section

