# Next Steps After Analysis

## ✅ Completed
- ✓ Analysis of Excel and PPT files
- ✓ Template file created

## 📋 Step-by-Step Configuration Guide

### Step 1: Review Your Excel Structure

Your Excel file has these sheets (check `analysis/excel_info.json` for details):
- Review which sheets contain data for which slides
- Note column names and data types

### Step 2: Review Your PPT Structure

Your PPT has 11 slides (check `analysis/template_info.json`):
- Slide 1: Title Slide
- Slides 2-11: Content slides
- Review what content appears on each slide

### Step 3: Configure Slide Mappings

Edit `config/slides.yaml` to map Excel data to PPT slides.

**Example for Slide 1 (Title Slide):**
```yaml
slides:
  - slide_number: 1
    slide_type: "title"
    title: "AIL LT"
    subtitle: "April 2025"
    layout_name: "Title Slide Layout"
```

**Example for a Table Slide:**
```yaml
  - slide_number: 3
    slide_type: "table"
    title: "Input Distribution Status"
    table_mapping:
      data_source: "AIL LT Working file"
      sheet: "INPUT DISTRIBUTION STATUS"
      columns: ["SLIDE 3", "Unnamed: 1"]  # Adjust based on actual column names
      max_rows: 15
      formatting:
        header_formatting:
          font_size: 14
          bold: true
        data_formatting:
          font_size: 12
```

### Step 4: Manual Review Process

1. **Open your manual PPT** (`Data/Apr 2025/AIL LT - April'25.pptx`)
2. **Go through each slide** and note:
   - What data appears on this slide?
   - Which Excel sheet/columns provide this data?
   - Is it a table, bullet list, or text?
   - What calculations are done?

3. **Open your Excel file** (`Data/Apr 2025/AIL LT Working file.xlsx`)
   - Check each sheet
   - Note column names (they may have "Unnamed" - you'll need to identify them)
   - See what calculations/formulas exist

### Step 5: Update Configuration Files

Based on your review, update:

1. **config/slides.yaml** - Map each slide to its data source
2. **config/rules.yaml** - Define any calculations needed
3. **config/formatting.yaml** - Match your PPT styling

### Step 6: Test Generation

Once configured, test with:

```bash
python main.py generate "Data/Apr 2025" "output/Apr_2025_Generated.pptx" --template "templates/template.pptx"
```

### Step 7: Validate Output

Compare generated PPT with manual version:

```bash
python -m src.validator "Data/Apr 2025/AIL LT - April'25.pptx" "output/Apr_2025_Generated.pptx" "validation/Apr_2025_report.json"
```

### Step 8: Refine and Iterate

- Review validation report
- Fix mismatches in config files
- Re-generate and validate until accuracy is acceptable

## 🔍 Quick Reference

**Check Excel sheets:**
```bash
python -c "import json; data = json.load(open('analysis/excel_info.json')); [print(f'{s[\"name\"]}: {s[\"row_count\"]} rows') for s in data['sheets']]"
```

**Check PPT slides:**
```bash
python -c "import json; data = json.load(open('analysis/template_info.json')); print(f'Total slides: {data[\"slide_count\"]}'); [print(f'Slide {s[\"slide_number\"]}: {s[\"layout_name\"]}') for s in data['slides']]"
```

## 💡 Tips

1. **Start Simple**: Configure one slide at a time, test, then move to the next
2. **Column Names**: Excel columns may be "Unnamed: X" - you'll need to identify what they contain
3. **Data Source Key**: Use the Excel file name (without extension) as the `data_source` key
4. **Sheet Names**: Must match exactly (case-sensitive)
5. **Test Frequently**: Generate and validate after each major configuration change

## 📝 Configuration Examples

See `GETTING_STARTED.md` for detailed configuration examples.

