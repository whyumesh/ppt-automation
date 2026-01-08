# Getting Started - Step-by-Step Guide

This guide will walk you through setting up and using the PPT automation system.

## Prerequisites

- Python 3.8 or higher installed
- Access to historical Excel files and PowerPoint decks
- Basic familiarity with command line/terminal

---

## Step 1: Install Dependencies

Open your terminal/command prompt in the project directory and run:

```bash
pip install -r requirements.txt
```

**What this does:** Installs all required Python packages (pandas, python-pptx, openpyxl, etc.)

**Expected output:** Packages install successfully without errors

---

## Step 2: Choose a Reference Month

Select one month from your historical data to use as a reference. For example:
- `Data/Apr 2025/` - Contains both Excel files and the manual PPT

**Why:** We'll use this month to discover the patterns and rules.

---

## Step 3: Run Reverse Engineering Analysis

This step analyzes your Excel and PPT files to discover how data maps to slides.

### 3.1: Analyze a Single Month

Run this command (replace paths with your actual files):

```bash
python main.py analyze "Data/Apr 2025/AIL LT Working file.xlsx" "Data/Apr 2025/AIL LT - April'25.pptx" --output-dir analysis
```

**What this does:**
- Extracts template structure from PPT
- Analyzes Excel file structure
- Compares them to discover mappings
- Saves results to `analysis/` folder

**Expected output:**
- `analysis/template_info.json` - PPT structure
- `analysis/excel_info.json` - Excel structure  
- `analysis/discovered_rules.json` - Initial rule discoveries

### 3.2: Review the Analysis Results

Open the JSON files in `analysis/` folder and examine:
- **template_info.json**: See what slides exist, what shapes/text are on each slide
- **excel_info.json**: See what sheets/columns exist in Excel
- **discovered_rules.json**: See initial mappings discovered

**What to look for:**
- Which Excel columns appear in which slides?
- What calculations are being done?
- What formatting is applied?

---

## Step 4: Extract PPT Template

Create a clean template file from one of your existing PPTs:

```bash
python -c "from src.template_extractor import TemplateExtractor; t = TemplateExtractor('Data/Apr 2025/AIL LT - April\'25.pptx'); t.create_template_copy('templates/template.pptx')"
```

Or create the templates directory and copy a PPT file manually:

```bash
mkdir templates
copy "Data\Apr 2025\AIL LT - April'25.pptx" templates\template.pptx
```

**What this does:** Creates a template file that will be used as the base for generated PPTs

---

## Step 5: Configure Slide Mappings

Edit `config/slides.yaml` to map Excel data to PPT slides.

### Example Configuration:

```yaml
slides:
  # Slide 1: Title Slide
  - slide_number: 1
    slide_type: "title"
    title: "AIL LT Report"
    subtitle: "April 2025"
    layout_name: "Title Slide"  # Name from your PPT template

  # Slide 2: Summary Table
  - slide_number: 2
    slide_type: "table"
    title: "Summary Data"
    table_mapping:
      data_source: "AIL LT Working file"  # Key from Excel file name
      sheet: "Summary"  # Sheet name in Excel
      columns: ["Category", "Value", "Previous Value"]  # Columns to include
      filters:
        - column: "Value"
          operator: ">="
          value: 0
      max_rows: 10
      formatting:
        header_formatting:
          font_size: 14
          bold: true
        data_formatting:
          font_size: 12

  # Slide 3: Bullet Points
  - slide_number: 3
    slide_type: "bullet_list"
    title: "Key Insights"
    items_data_source:
      source: "AIL LT Working file"
      sheet: "Insights"
      column: "Insight"
      default: []
```

**How to fill this:**
1. Look at your manual PPT - count how many slides there are
2. For each slide, identify:
   - What Excel file/sheet provides the data?
   - What columns are used?
   - Is it a table, bullet list, or text?
3. Document this in `slides.yaml`

---

## Step 6: Configure Business Rules

Edit `config/rules.yaml` to define calculations and transformations.

### Example Configuration:

```yaml
rules:
  # Calculate growth rate
  calculate_growth:
    type: "calculation"
    operation: "percentage_change"
    params:
      current: "Value"
      previous: "Previous Value"
    data_source: "AIL LT Working file"

  # Filter top performers
  filter_top_10:
    type: "filter"
    filter_type: "top_n"
    params:
      column: "Value"
      n: 10
    data_source: "AIL LT Working file"

  # Generate performance text
  performance_status:
    type: "text_generation"
    template: "Performance {status} by {percentage}%"
    params:
      status:
        type: "conditional"
        condition:
          type: "compare"
          left:
            type: "data_column"
            column: "growth_rate"
          operator: ">="
          right:
            type: "literal"
            value: 0
        true_value: "Improved"
        false_value: "Declined"
      percentage:
        type: "data_column"
        column: "growth_rate"
        aggregate: "mean"
```

**How to fill this:**
1. Review your Excel data - what calculations are done?
2. Review your PPT - what derived values appear?
3. Document the formulas/logic in `rules.yaml`

---

## Step 7: Configure Formatting

Edit `config/formatting.yaml` to match your PPT styling:

```yaml
formatting:
  fonts:
    default_size: 12
    title_size: 24
    heading_size: 18
    default_name: "Calibri"
  colors:
    positive: "#00FF00"  # Green for positive values
    negative: "#FF0000"   # Red for negative values
    neutral: "#000000"    # Black for neutral
  number_formats:
    percentage: "{:.1f}%"
    currency: "${:,.2f}"
    integer: "{:,}"
```

**How to fill this:**
- Check your manual PPT for font sizes, colors, number formats
- Match them in the config file

---

## Step 8: Test Generation (First Run)

Generate a PPT for the same month you analyzed:

```bash
python main.py generate "Data/Apr 2025" "output/Apr_2025_Generated.pptx" --template "templates/template.pptx"
```

**What this does:**
- Loads all Excel files from the month directory
- Applies transformations and rules
- Generates PowerPoint deck
- Saves to output file

**Expected output:**
- Console messages showing progress
- Generated PPT file in `output/` folder

---

## Step 9: Validate Output

Compare the generated PPT with your manual version:

```bash
python -m src.validator "Data/Apr 2025/AIL LT - April'25.pptx" "output/Apr_2025_Generated.pptx" "validation/Apr_2025_report.json"
```

**What this does:**
- Compares slide-by-slide
- Checks text, tables, formatting
- Generates validation report

**Review the report:**
- Check accuracy percentage
- Review mismatches
- Identify what needs fixing

---

## Step 10: Refine Configuration

Based on validation results, update your configs:

1. **If numbers don't match:**
   - Check calculation rules in `rules.yaml`
   - Verify Excel column names match

2. **If slides are missing:**
   - Add missing slides to `slides.yaml`

3. **If formatting is wrong:**
   - Update `formatting.yaml`
   - Check slide-specific formatting in `slides.yaml`

4. **If data is wrong:**
   - Verify Excel file paths
   - Check sheet names
   - Verify column mappings

---

## Step 11: Iterate and Improve

Repeat steps 8-10 until accuracy is acceptable (>95%):

```bash
# Generate
python main.py generate "Data/Apr 2025" "output/Apr_2025_Generated_v2.pptx" --template "templates/template.pptx"

# Validate
python -m src.validator "Data/Apr 2025/AIL LT - April'25.pptx" "output/Apr_2025_Generated_v2.pptx" "validation/Apr_2025_v2_report.json"

# Review and refine
# ... update configs ...

# Repeat
```

---

## Step 12: Test on Other Months

Once one month works well, test on other historical months:

```bash
# Test on May 2025
python main.py generate "Data/May 2025" "output/May_2025_Generated.pptx" --template "templates/template.pptx"
python -m src.validator "Data/May 2025/AIL LT - May'25.pptx" "output/May_2025_Generated.pptx" "validation/May_2025_report.json"

# Test on June 2025
python main.py generate "Data/June 2025" "output/June_2025_Generated.pptx" --template "templates/template.pptx"
python -m src.validator "Data/June 2025/AIL LT - June'25.pptx" "output/June_2025_Generated.pptx" "validation/June_2025_report.json"
```

**What to check:**
- Do all months work?
- Are there month-specific differences?
- Do you need conditional logic for different months?

---

## Step 13: Document Discovered Rules

Update documentation files:

1. **docs/mappings.md**: Document Excel → Slide mappings
2. **docs/business_rules.md**: Document all business logic

**Why:** This helps future maintenance and onboarding

---

## Step 14: Production Use

Once validated, use for new months:

```bash
# For a new month (e.g., Dec 2025)
python main.py generate "Data/Dec 2025" "output/Dec_2025_Generated.pptx" --template "templates/template.pptx"
```

**Workflow:**
1. Place new month's Excel files in `Data/Dec 2025/`
2. Run generate command
3. Review generated PPT
4. Make manual adjustments if needed (and update configs for next time)

---

## Troubleshooting

### Issue: "Excel file not found"
**Solution:** Check file path, ensure Excel files are in the month directory

### Issue: "Sheet not found"
**Solution:** Verify sheet names in Excel match what's in `slides.yaml`

### Issue: "Column not found"
**Solution:** Check column names in Excel match config (case-sensitive)

### Issue: Generated PPT looks wrong
**Solution:** 
1. Run validation to see specific mismatches
2. Check template file matches your manual PPT structure
3. Verify formatting config matches manual PPT

### Issue: Numbers don't match
**Solution:**
1. Check calculation rules in `rules.yaml`
2. Verify Excel data is correct
3. Check if rounding/formats are applied correctly

---

## Quick Reference Commands

```bash
# Analyze files
python main.py analyze <excel_file> <ppt_file> --output-dir analysis

# Generate PPT
python main.py generate <month_dir> <output_file> --template <template_file>

# Validate
python -m src.validator <manual_ppt> <generated_ppt> <report_file>

# Extract template info only
python -c "from src.template_extractor import TemplateExtractor; t = TemplateExtractor('path/to/ppt.pptx'); t.extract_all(); t.save_template_info('template_info.json')"

# Analyze Excel only
python -c "from src.excel_analyzer import ExcelAnalyzer; a = ExcelAnalyzer('path/to/excel.xlsx'); a.analyze_all(); a.save_analysis('excel_info.json')"
```

---

## Next Steps After Setup

1. **Automate monthly runs**: Create a batch script/scheduler
2. **Add more rules**: As you discover new patterns
3. **Extend for new slide types**: If needed
4. **Create unit tests**: For critical rules (in `tests/` folder)

---

## Need Help?

- Check `README.md` for architecture overview
- Review `docs/` folder for detailed documentation
- Examine example configs in `config/` folder
- Review analysis results in `analysis/` folder

