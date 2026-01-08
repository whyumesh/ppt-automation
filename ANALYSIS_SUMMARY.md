# Analysis Summary - Excel to PPT Mappings

## ✅ Completed Analysis

I've analyzed your Excel and PPT files and created detailed mappings. Here's what was discovered:

## 📊 Excel File: `AIL LT Working file.xlsx`

### Sheets Identified (12 total):

1. **INPUT DISTRIBUTION STATUS** - Slide 3, 7
   - Header Row: 2
   - Columns: `Division`, ` Total Dis ` (percentages)
   - 13 data rows

2. **Input distribution** - Slide 3, 5
   - Header Row: 0 (but empty rows at top)
   - Columns: `Unnamed: 2` (Division), `Unnamed: 3` (YTD %), `Unnamed: 4` (SPL %)
   - 15 rows total

3. **WITHOLD EXPENSE** - Slide 8
   - Header Row: 4
   - Columns: `Slide 5` (Division), `Unnamed: 1-6` (May/June status data)
   - 16 rows total

4. **Input with hold** - Slide 6
   - Header Row: 2
   - Columns: `Divisions`, `Active HC`, `Nov'23` through `Apr'24`, `%`, `Common TBMs`
   - 23 data rows

5. **consent** - Slide 4, 10
   - Header Row: 0
   - Columns: `Division Name`, `DVL`, `# HCP Consent`, `% Consent Require`
   - 10 rows

6. **Chronic & Overcalling** - Slide 9
   - Header Row: 4
   - Columns: `Slide 9` (Division), `Unnamed: 1` (#DVL), `Unnamed: 2` (#Missed), `Unnamed: 3` (% Missed - decimal)
   - 16 rows

7. **WORKING 3** - Reference data (540 rows)
8. **FREQUENT DEFAULTERS 2** - Reference data (544 rows)
9. **WITHHOLD COMMON EMPLOYEES 1** - Reference data (709 rows)
10. **Overlapping drs** - Slide 7 reference
11. **Sheet2** - Slide 8 reference
12. **CLT** - Alternative data for Slide 3

## 📄 PowerPoint: 11 Slides

### Slide Breakdown:

| Slide | Type | Primary Excel Source | Key Data |
|-------|------|---------------------|----------|
| 1 | Title | None | Static: "AIL LT April'25" |
| 2 | Chart | Multiple | Business effectiveness metrics |
| 3 | Table | INPUT DISTRIBUTION STATUS | Division + Distribution % |
| 4 | Table | consent | Division + Consent metrics |
| 5 | Text/Table | Input distribution | Aggregated summary (93%) |
| 6 | Table | Input with hold | Division + Monthly data + % |
| 7 | Table | INPUT DISTRIBUTION STATUS | Same as Slide 3 |
| 8 | Table | WITHOLD EXPENSE | Division + May/June status |
| 9 | Table | Chronic & Overcalling | Division + Missed HCP data |
| 10 | Table/Text | consent | Similar to Slide 4 |
| 11 | Unknown | TBD | Need to verify |

## 🔍 Key Discoveries

### Column Naming Patterns:
- ✅ Some sheets have proper column names: `Division`, `Division Name`, `# of HCPs`
- ⚠️ Many columns are `Unnamed: X` - identified by position and content
- 📍 Slide indicators in column names: `SLIDE 3`, `Slide 5`, `Slide 9`, `Slide 10`

### Data Types:
- **Tables:** Slides 3, 4, 6, 7, 8, 9 (6 slides)
- **Text/Bullets:** Slides 1, 2, 5, 10 (4 slides)
- **Charts:** Slide 2 (1 slide - may need manual handling)

### Calculations Found:
1. **Percentages:**
   - Distribution %: Already calculated (97.31%, 95.93%, etc.)
   - Consent %: Already calculated (33.85%, 30.40%, etc.)
   - HCP Missed %: Decimal format (0.0486 = 4.86%) - needs conversion

2. **Deltas:**
   - Withhold expense: Change between May and June status
   - Month-over-month comparisons

3. **Aggregations:**
   - Overall percentages (e.g., 93% aggregate)
   - Counts and sums by division

### Formatting Identified:
- **Font Sizes:** 9pt, 10pt, 12pt, 14pt, 20pt, 24pt, 26pt
- **Colors:** Blue titles RGB(0,176,240), Dark blue headers RGB(0,59,85), Red warnings RGB(192,0,0)
- **Table Headers:** Dark blue background with white bold text
- **Table Footers:** Dark blue background

## 📝 Configuration Files Created

1. **config/slides.yaml** - Complete slide mappings with:
   - Data sources
   - Column mappings
   - Header row positions
   - Filtering rules
   - Formatting specifications

2. **SLIDE_MAPPINGS.md** - Detailed documentation of each slide

3. **CONFIGURATION_GUIDE.md** - Step-by-step configuration guide

4. **analysis/precise_excel_analysis.json** - Precise Excel structure analysis

## 🎯 Ready for Testing

The configuration is ready! You can now:

1. **Test Slide 3** (simplest):
   ```bash
   python main.py generate "Data/Apr 2025" "output/test_slide3.pptx" --template "templates/template.pptx"
   ```

2. **Validate against manual:**
   ```bash
   python -m src.validator "Data/Apr 2025/AIL LT - April'25.pptx" "output/test_slide3.pptx" "validation/slide3_report.json"
   ```

3. **Refine and iterate** based on validation results

## ⚠️ Notes & Warnings

1. **Header Rows:** Many sheets have headers on row 2 or 4 (not row 0) - configured correctly
2. **Empty Rows:** Some sheets have empty rows at top - filters configured
3. **Unnamed Columns:** Identified by position - may need adjustment if Excel structure changes
4. **Slide 2:** Contains chart/visualization - may need manual configuration or chart generation
5. **External References:** Some slides reference external files - these need manual handling
6. **Percentage Formats:** Some are decimals (0.0486), some are percentages (4.86) - conversion rules configured

## 📚 Files to Review

- `config/slides.yaml` - Main configuration (READY TO USE)
- `SLIDE_MAPPINGS.md` - Detailed slide-by-slide documentation
- `CONFIGURATION_GUIDE.md` - Configuration reference guide
- `analysis/precise_excel_analysis.json` - Excel structure details

## 🚀 Next Steps

1. ✅ Review `config/slides.yaml` - mappings are configured
2. ✅ Test generation for one slide (start with Slide 3)
3. ✅ Validate output
4. ✅ Refine configuration based on results
5. ✅ Expand to all slides

The system is ready to generate PowerPoint decks! Start testing with Slide 3.

