# Slide Operations Summary
## Quick Reference: File → Operations → Slide

This is a concise summary showing which Excel file/sheet feeds which slide and what operations are performed.

---

## Excel File: `AIL LT Working file.xlsx`

### Sheet: `INPUT DISTRIBUTION STATUS`
**Used for:** Slide 3

**Operations:**
1. Read sheet with `header=2` (header on row 2)
2. Select columns: `Division`, ` Total Dis `
3. Filter empty rows
4. Format: Convert decimal to percentage (97.31 → 97.31%)
5. Limit: Top 12-13 rows
6. Create PPT table with 2 columns

**Example:**
- Excel: `["GI Prima", 97.305246]`
- PPT: `| GI Prima | 97.31% |`

---

### Sheet: `consent`
**Used for:** Slide 4, Slide 10

**Operations:**
1. Read sheet with `header=0`
2. Select columns: `Division Name`, `DVL`, `# HCP Consent`, `% Consent Require`
3. Format: 
   - DVL: Integer
   - # HCP Consent: Integer
   - % Consent Require: Decimal → Percentage (0.68 → 68%)
4. Limit: Top 10 rows
5. Create PPT table with 4 columns

**Example:**
- Excel: `["GI Prima", 1250, 850, 0.68]`
- PPT: `| GI Prima | 1250 | 850 | 68% |`

---

### Sheet: `Input distribution`
**Used for:** Slide 3 (alternative), Slide 5

**Operations (Slide 5):**
1. Read sheet (header detection needed)
2. Select percentage column: `Unnamed: 3` (YTD Distribution %) or `Unnamed: 4` (SPL Distribution %)
3. **Calculate:** Average of all percentages
4. Format: Round to whole number (93.45 → 93%)
5. Generate text: "93% YTD Distribution"
6. Insert into PPT text box

**Example:**
- Excel: Multiple rows with percentages [85.59, 75.65, 92.92, ...]
- Calculation: Average = 93.45%
- PPT: Text box showing "93% YTD Distribution"

---

### Sheet: `Input with hold`
**Used for:** Slide 6

**Operations:**
1. Read sheet with `header=2`
2. Select columns: `Divisions`, `Active HC`, `Nov'23`, `Dec'23`, `Jan'24`, `Feb'24`, `Mar'24`, `Apr'24`, `%`, `Common TBMs`
3. Format:
   - Monthly columns: Integers
   - Percentage: Decimal → Percentage
   - TBMs: Text (as-is)
4. Limit: Top 23 rows
5. Create PPT table with 10 columns

**Example:**
- Excel: `["GI Prima", 125, 95, 98, ..., 102, 0.85, "TBM1, TBM2"]`
- PPT: `| GI Prima | 125 | 95 | 98 | ... | 102 | 85% | TBM1, TBM2 |`

---

### Sheet: `WITHOLD EXPENSE`
**Used for:** Slide 8

**Operations:**
1. Read sheet with `header=4`
2. Select columns: `Slide 5` (Division), `Unnamed: 1-6` (Status columns)
3. Format: Division as text, status indicators formatted
4. Limit: Top 16 rows
5. Create PPT table

---

### Sheet: `Chronic & Overcalling`
**Used for:** Slide 9

**Operations:**
1. Read sheet with `header=4`
2. Select columns: `Slide 9` (Division), `Unnamed: 1` (#DVL), `Unnamed: 2` (#Missed), `Unnamed: 3` (% Missed)
3. Format:
   - #DVL: Integer
   - #Missed: Integer
   - % Missed: Decimal → Percentage (0.15 → 15%)
4. Limit: Top 16 rows
5. Create PPT table with 4 columns

**Example:**
- Excel: `["GI Prima", 1250, 188, 0.15]`
- PPT: `| GI Prima | 1250 | 188 | 15% |`

---

### Sheet: `Overlapping drs`
**Used for:** Slide 7

**Operations:**
1. Read sheet
2. Filter: Doctors with multiple territories (territory count > 1)
3. Select columns: Doctor name, Territories, Count
4. Format: Text and integers
5. Create PPT table

---

## Additional Files (Reports Folder)

### File: `AIL - Speaker Allocation Summary - 04-April-25.xlsx`
**Used for:** Possibly Slide 2 (Business Effectiveness chart)

### File: `AIL Consented Status HCP's_02.04.2025.xlsb`
**Used for:** Slide 4 or Slide 10 (alternative consent data)

### File: `Chronic Missing Report AIL - Jan to Mar.xlsx`
**Used for:** Slide 9 (alternative chronic missing data)

### File: `Overcalling Report AIL - Jan to Mar.xlsx`
**Used for:** Slide 9 (overcalling data)

---

## Operation Patterns

### Pattern 1: Simple Table (Slides 3, 4, 6, 8, 9)
```
Excel Sheet → Read with header → Select columns → Format → Limit rows → PPT Table
```

### Pattern 2: Aggregated Text (Slide 5)
```
Excel Sheet → Read → Select column → Calculate average → Format → PPT Text
```

### Pattern 3: Filtered Table (Slide 7)
```
Excel Sheet → Read → Filter condition → Select columns → Format → PPT Table
```

### Pattern 4: Static Content (Slide 1, 11)
```
No Excel → Static text → PPT Text
```

---

## Common Transformations

| Excel Format | PPT Format | Operation |
|-------------|------------|-----------|
| `97.305246` | `97.31%` | Decimal → Percentage, Round 2 decimals |
| `0.68` | `68%` | Decimal → Percentage, Multiply by 100 |
| `1250.0` | `1250` | Float → Integer |
| `"GI Prima"` | `"GI Prima"` | Text (as-is) |
| `[85.59, 75.65, ...]` | `"93%"` | Average → Percentage |

---

## Header Row Detection

| Sheet | Header Row (0-indexed) | Notes |
|-------|------------------------|-------|
| INPUT DISTRIBUTION STATUS | 2 | Has empty rows at top |
| consent | 0 | Standard header |
| Input distribution | 0 | But has empty rows, actual data starts later |
| Input with hold | 2 | Has empty rows at top |
| WITHOLD EXPENSE | 4 | Has multiple empty rows |
| Chronic & Overcalling | 4 | Has multiple empty rows |

---

## Column Mapping Reference

### Slide 3
- Excel: `Division` (or `SLIDE 3`), ` Total Dis ` (or `Unnamed: 1`)
- PPT: Column 1 = Division, Column 2 = Distribution %

### Slide 4
- Excel: `Division Name`, `DVL`, `# HCP Consent`, `% Consent Require`
- PPT: Column 1 = Division, Column 2 = DVL, Column 3 = # HCP Consent, Column 4 = % Consent

### Slide 5
- Excel: `Unnamed: 3` (YTD Distribution %) or `Unnamed: 4` (SPL Distribution %)
- PPT: Text showing aggregated percentage

### Slide 6
- Excel: `Divisions`, `Active HC`, `Nov'23`, `Dec'23`, `Jan'24`, `Feb'24`, `Mar'24`, `Apr'24`, `%`, `Common TBMs`
- PPT: 10 columns with monthly data

### Slide 9
- Excel: `Slide 9` (Division), `Unnamed: 1` (#DVL), `Unnamed: 2` (#Missed), `Unnamed: 3` (% Missed)
- PPT: Column 1 = Division, Column 2 = #DVL, Column 3 = #Missed, Column 4 = % Missed

---

## Quick Decision Tree

```
For each slide:
1. Is it static? (Slide 1, 11)
   → No Excel operations
   
2. Is it a table?
   → Read Excel sheet
   → Identify header row
   → Select columns
   → Format data
   → Limit rows
   → Create PPT table
   
3. Is it aggregated text? (Slide 5)
   → Read Excel sheet
   → Select percentage column
   → Calculate average
   → Format as percentage
   → Generate text
   → Insert into PPT
   
4. Is it filtered data? (Slide 7)
   → Read Excel sheet
   → Apply filter condition
   → Select columns
   → Format
   → Create PPT table
```

---

This summary provides a quick reference for understanding the deck creation process.

