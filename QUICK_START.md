# Quick Start Guide - What to Do Next

## ✅ Current Status

- ✅ System is working - PPT was generated successfully!
- ⚠️ Generated PPT has 22 slides (should be 11) - likely using template with existing slides
- ⚠️ Some slides may be empty or have placeholder data

## 🎯 Immediate Next Steps

### 1. Open and Compare the PPTs

**Open both files side-by-side:**
- Manual: `Data/Apr 2025/AIL LT - April'25.pptx`
- Generated: `output/test.pptx`

**Check:**
- How many slides are in each?
- Do the slides have data in tables?
- Are the numbers correct?
- Is the formatting similar?

### 2. Fix the Template Issue

The template likely contains all slides from the original PPT. You have two options:

**Option A: Use a clean template**
1. Open `templates/template.pptx` in PowerPoint
2. Delete all slides except keep the layouts
3. Save it
4. Re-generate

**Option B: Fix the code to clear template slides**

The code should clear existing slides, but it might not be working. Check if slides are being added on top of template slides.

### 3. Verify Data is Loading

Check if Excel data is being loaded correctly:

```bash
python -c "import sys; sys.path.insert(0, 'src'); from data_loader import DataLoader; loader = DataLoader(); data = loader.load_excel('Data/Apr 2025/AIL LT Working file.xlsx', 'INPUT DISTRIBUTION STATUS'); print(f'Loaded {len(data)} rows'); print(data.head())"
```

### 4. Check Slide Configuration

Review `config/slides.yaml`:
- Are slide numbers correct?
- Are data sources pointing to the right Excel file?
- Are sheet names matching exactly?

### 5. Test One Slide at a Time

**Temporarily modify `config/slides.yaml` to only generate Slide 3:**

```yaml
slides:
  # Comment out other slides, keep only:
  - slide_number: 3
    slide_type: "table"
    # ... rest of config
```

Then regenerate and check if Slide 3 has the correct data.

## 🔧 Common Issues & Quick Fixes

### Issue: Slides are empty

**Check:**
1. Excel file name matches `data_source` in config
2. Sheet name matches exactly
3. Column names are correct
4. Header row is correct

### Issue: Wrong data

**Check:**
1. `header_row` parameter (0-indexed)
2. Column names (especially "Unnamed: X")
3. Filters might be excluding data

### Issue: Too many slides

**Fix:** The template has existing slides. Either:
- Clean the template (remove all slides, keep layouts)
- Or fix the code to properly clear template slides

## 📝 Recommended Workflow

1. **Clean the template** - Remove all slides, keep layouts
2. **Test Slide 1** - Title slide (simplest)
3. **Test Slide 3** - Table slide
4. **Verify data** - Check numbers match
5. **Fix formatting** - Adjust fonts, colors
6. **Add remaining slides** - One at a time
7. **Final validation** - Compare with manual version

## 🚀 Quick Commands

**Generate:**
```bash
python main.py generate "Data/Apr 2025" "output/test.pptx" --template "templates/template.pptx"
```

**Validate:**
```bash
python validate_output.py "Data/Apr 2025/AIL LT - April'25.pptx" "output/test.pptx" "validation/report.json"
```

**Check Excel data:**
```bash
python extract_precise_mappings.py
```

## 💡 What to Look For

When comparing PPTs:

1. **Slide count** - Should be 11 slides
2. **Table data** - Should match Excel data
3. **Numbers** - Should be exact matches
4. **Formatting** - Fonts, colors, sizes
5. **Text content** - Titles, labels, footers

## 📚 Documentation

- `NEXT_STEPS_AFTER_GENERATION.md` - Detailed next steps
- `CONFIGURATION_GUIDE.md` - Configuration reference
- `SLIDE_MAPPINGS.md` - Detailed slide mappings
- `ANALYSIS_SUMMARY.md` - Analysis findings

---

**Start by opening both PPT files and visually comparing them. Note what's different, then we can fix the configuration accordingly.**

