# Next Steps After First Generation

## ✅ What Just Happened

You successfully generated your first PowerPoint deck! The file is at: `output/test.pptx`

## 🔍 Step 1: Review the Generated PPT

**Open the generated file:**
```
output/test.pptx
```

**Compare it with the manual version:**
```
Data/Apr 2025/AIL LT - April'25.pptx
```

**Check:**
- ✅ Are slides present?
- ✅ Do tables have data?
- ✅ Are numbers correct?
- ✅ Is formatting close?
- ❌ What's missing?
- ❌ What's wrong?

## 📊 Step 2: Validate Programmatically

Run the validator:

```bash
python validate_output.py "Data/Apr 2025/AIL LT - April'25.pptx" "output/test.pptx" "validation/report.json"
```

This will create a detailed report showing:
- Which slides match
- Which slides have mismatches
- What text/numbers differ
- Accuracy percentage

## 🔧 Step 3: Common Issues & Fixes

### Issue: Slides are empty or missing data

**Fix:** Check `config/slides.yaml`:
- Verify `data_source` matches Excel file name (without .xlsx)
- Verify `sheet` name matches exactly (case-sensitive)
- Check `header_row` is correct
- Verify column names match

### Issue: Wrong data in tables

**Fix:**
- Check if `header_row` needs adjustment
- Verify column names (especially "Unnamed: X" columns)
- Check if filters are excluding needed data
- Verify `max_rows` isn't too restrictive

### Issue: Numbers don't match

**Fix:**
- Check if calculations are needed (percentages, deltas)
- Verify number formatting (decimals vs percentages)
- Check if data needs aggregation

### Issue: Formatting is wrong

**Fix:** Update `config/slides.yaml`:
- Adjust `font_size` values
- Update `fill_color` for headers
- Check `number_formatting` settings

## 📝 Step 4: Refine Configuration

Based on your review, update `config/slides.yaml`:

### Example: Fix Slide 3 Table

If Slide 3 table is empty or wrong:

1. **Check the actual Excel data:**
   ```bash
   python extract_precise_mappings.py
   ```
   Review the output for "INPUT DISTRIBUTION STATUS" sheet

2. **Update slides.yaml:**
   - Adjust `header_row` if needed
   - Verify column names
   - Check filters

3. **Re-generate:**
   ```bash
   python main.py generate "Data/Apr 2025" "output/test_v2.pptx" --template "templates/template.pptx"
   ```

4. **Validate again:**
   ```bash
   python validate_output.py "Data/Apr 2025/AIL LT - April'25.pptx" "output/test_v2.pptx" "validation/report_v2.json"
   ```

## 🎯 Step 5: Iterate Slide by Slide

**Recommended approach:**

1. **Start with Slide 1** (title slide - simplest)
   - Verify title and subtitle are correct
   - Fix month/year extraction if needed

2. **Then Slide 3** (table slide)
   - Verify table data matches
   - Check formatting
   - Fix column mappings if needed

3. **Continue with other slides** one at a time

4. **Test after each fix**

## 📋 Step 6: Add Missing Features

### If slides need calculations:

Add rules to `config/rules.yaml`:

```yaml
rules:
  calculate_overall_percentage:
    type: "calculation"
    operation: "mean"
    params:
      column: " Total Dis "
    data_source: "AIL LT Working file"
    sheet: "INPUT DISTRIBUTION STATUS"
```

### If slides need conditional formatting:

Update `config/slides.yaml` formatting section:

```yaml
formatting:
  conditional_colors:
    - column: "Value"
      condition: "<"
      threshold: 0
      color: "#C00000"  # Red for negative
```

## 🔄 Step 7: Complete Workflow

Once one slide works perfectly:

1. **Document the working configuration**
2. **Apply same patterns to other slides**
3. **Test all slides together**
4. **Validate final output**

## 📚 Useful Commands

**Generate PPT:**
```bash
python main.py generate "Data/Apr 2025" "output/test.pptx" --template "templates/template.pptx"
```

**Validate:**
```bash
python validate_output.py "Data/Apr 2025/AIL LT - April'25.pptx" "output/test.pptx" "validation/report.json"
```

**Review Excel structure:**
```bash
python extract_precise_mappings.py
```

**Review analysis:**
```bash
python review_analysis.py
```

## 💡 Tips

1. **Work incrementally** - Fix one slide at a time
2. **Test frequently** - Generate and validate after each change
3. **Keep notes** - Document what works and what doesn't
4. **Check validation reports** - They show exactly what differs
5. **Compare side-by-side** - Open both PPTs and compare visually

## 🎉 Success Criteria

You're done when:
- ✅ All slides are generated
- ✅ Numbers match manual version (>95% accuracy)
- ✅ Formatting matches manual version
- ✅ Tables have correct data
- ✅ Text content is correct

## 🚀 Next: Production Use

Once validated:
1. Test on other months (May, June, etc.)
2. Create a batch script for monthly automation
3. Document any month-specific variations
4. Set up scheduled runs (if needed)

---

**Start by opening both PPT files and comparing them visually. Note what's different, then update the configuration accordingly.**

