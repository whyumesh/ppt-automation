# Configuration Guide - Based on Analysis

## Summary of Discovered Mappings

### Excel File Structure
**File:** `AIL LT Working file.xlsx`
**Key:** Use `"AIL LT Working file"` (without .xlsx) as data_source

### Sheet-to-Slide Mappings

| Slide | Sheet Name | Header Row | Key Columns | Type |
|-------|-----------|------------|-------------|------|
| 1 | None | - | Static title | Text |
| 2 | Multiple | - | Chart data | Chart/Visual |
| 3 | INPUT DISTRIBUTION STATUS | 2 | Division, Total Dis % | Table |
| 4 | consent | 0 | Division Name, DVL, # HCP Consent, % Consent Require | Table |
| 5 | Input distribution | 0 | Aggregated summary | Text/Table |
| 6 | Input with hold | 2 | Divisions, Active HC, Months (Nov'23-Apr'24), % | Table |
| 7 | INPUT DISTRIBUTION STATUS | 2 | Division, Total Dis % | Table |
| 8 | WITHOLD EXPENSE | 4 | Division, May Status cols, June Status cols | Table |
| 9 | Chronic & Overcalling | 4 | Division, #DVL, #HCPs Missed, % HCP Missed | Table |
| 10 | consent | 0 | Similar to Slide 4 | Table/Text |
| 11 | TBD | - | - | - |

## Column Name Mappings

### Sheet: INPUT DISTRIBUTION STATUS
- **Header Row:** 2 (0-indexed, so row 3 in Excel)
- **Columns:**
  - `Division` - Division names
  - ` Total Dis ` - Total Distribution percentage (already calculated)

### Sheet: Input distribution
- **Header Row:** 0 (but first few rows are empty)
- **Columns:**
  - `SLIDE - 3` - Empty/header indicator
  - `Unnamed: 1` - Empty
  - `Unnamed: 2` - Division names (starts around row 5)
  - `Unnamed: 3` - YTD Distribution % (percentages like 85.59%)
  - `Unnamed: 4` - SPL Distribution % (percentages like 93.87%)

### Sheet: Input with hold
- **Header Row:** 2
- **Columns:**
  - `Divisions` - Division names
  - `Active HC` - Active headcount (integers)
  - `Nov'23`, `Dec'23`, `Jan'24`, `Feb'24`, `Mar'24`, `Apr'24` - Monthly data
  - `%` - Percentage column (mostly 0)
  - `Common TBMs in all 3 months` - Count of common TBMs

### Sheet: WITHOLD EXPENSE
- **Header Row:** 4 (headers are in rows 4-5)
- **Columns:**
  - `Slide 5` - Division names (data starts after header rows)
  - `Unnamed: 1` - May Status: #employees in original withhold list
  - `Unnamed: 2` - May Status: Allowed exception #
  - `Unnamed: 3` - May Status: Final # employees with expense withheld
  - `Unnamed: 4` - June Status: #employees in original withhold list
  - `Unnamed: 5` - June Status: Allowed exception #
  - `Unnamed: 6` - June Status: Final # employees with expense withheld

### Sheet: consent
- **Header Row:** 0
- **Columns:**
  - `Division Name` - Division names
  - `# of HCPs` or `DVL` - Total HCP count
  - `Consent Received/Accepted` or `# HCP Consent` - Consent count
  - `% Consent Require` - Percentage requiring consent (already calculated)

### Sheet: Chronic & Overcalling
- **Header Row:** 4
- **Columns:**
  - `Slide 9` - Division names
  - `Unnamed: 1` - #DVL (total DVL count)
  - `Unnamed: 2` - #HCPs Missed (count)
  - `Unnamed: 3` - % HCP Missed (decimal format, e.g., 0.0486 = 4.86%)

## Calculations Required

### 1. Percentage Formatting
- **Distribution Percentages:** Already calculated in Excel, format as "XX.X%"
- **Consent Percentages:** Already in `% Consent Require` column, format as "XX.X%"
- **HCP Missed Percentages:** Decimal format (0.0486), multiply by 100 and format as "X.X%"

### 2. Delta Calculations
- **Withhold Expense:** Calculate change between May and June status
  - Delta = June Final - May Final
  - Format negative values in red

### 3. Aggregations
- **Slide 5:** Calculate overall input distribution percentage (93% aggregate)
- **Common TBMs:** Count from `Common TBMs in all 3 months` column

## Formatting Rules

### Font Sizes
- **Titles:** 24pt (some 26pt for Slide 8)
- **Subheadings:** 20pt
- **Table Headers:** 14pt, bold
- **Table Data:** 12pt (some 9pt for Slide 4)
- **Footers:** 10pt

### Colors
- **Title Blue:** RGB(0, 176, 240) = #00B0F0
- **Dark Blue Header:** RGB(0, 59, 85) = #003B55
- **Red Warning:** RGB(192, 0, 0) = #C00000
- **Text Blue:** RGB(0, 156, 222) = #009CDE

### Table Formatting
- **Header Rows:** Dark blue background RGB(0, 59, 85), white text, bold
- **Footer Rows:** Dark blue background RGB(0, 59, 85)
- **Data Rows:** White background, black text
- **Conditional Formatting:** Red text for negative values in certain columns

## Data Filtering Rules

### General Rules
1. **Skip Empty Rows:** Filter out rows where key column (usually Division) is null/empty
2. **Skip Header Rows:** Use `header_row` parameter to skip to actual data
3. **Skip Text Rows:** Filter out rows where Division = "Division" (header text)

### Specific Filters

**Slide 3 & 7:**
- Filter: `Division != null AND Division != "Division"`

**Slide 6:**
- Filter: `Divisions != null`
- Limit to rows with actual data (exclude empty rows at top)

**Slide 8:**
- Filter: `Slide 5 != null AND Slide 5 != "Division" AND Slide 5 != "May Status..."`

**Slide 9:**
- Filter: `Slide 9 != null AND Slide 9 != "Division" AND Slide 9 != "#HCPs Missed..."`

## Number Formatting

### Percentage Columns
- ` Total Dis `: Format as "XX.X%" (values are already percentages)
- `% Consent Require`: Format as "XX.X%" (values are already percentages)
- `Unnamed: 3` (HCP Missed %): Multiply by 100, format as "X.X%" (values are decimals)

### Integer Columns
- `Active HC`: Format as integer (no decimals)
- `#DVL`, `#HCPs Missed`: Format as integer with commas
- Employee counts: Format as integer

### Decimal Columns
- Monthly data (Nov'23, Dec'23, etc.): Format as integer (they appear to be counts)

## Next Steps

1. **Test Slide 3** first (simplest table mapping)
2. **Verify header rows** by opening Excel and checking actual row numbers
3. **Test Slide 7** (same data as Slide 3, different layout)
4. **Configure Slide 6** (most complex table with multiple columns)
5. **Add calculations** for deltas and aggregations
6. **Test formatting** matches manual PPT

## Notes

- Many sheets have empty rows at the top - use `header_row` parameter
- Some columns are "Unnamed: X" - identify by position and sample data
- Percentages may be stored as decimals (0.0486) or percentages (4.86) - check and convert accordingly
- Some slides may reference external files (e.g., "Refer file Chronic missed data") - these need manual handling

