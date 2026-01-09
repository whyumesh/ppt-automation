# Validation Report - Generated Slides vs Manual PPT

## Summary
- **Slide 4 (Consent)**: ✗ FAIL - Data present but needs fixes
- **Slide 9 (Chronic & Overcalling)**: ✗ FAIL - Data present but needs fixes

## Issues Found

### Slide 4 (Consent) Issues:

1. **Missing Column**: 
   - Manual has 5 columns: `Division Name`, `DVL`, `# HCP Consent`, `Consent Require`, `% Consent Require`
   - Generated has 4 columns: Missing `Consent Require` column

2. **Row Order**:
   - Manual starts with "AIL" (aggregate row)
   - Generated starts with "GenNext"
   - Need to sort or reorder rows to match manual

3. **Number Formatting**:
   - Generated: `37068.0`, `34.63634401640229`
   - Manual: `37068`, `35%`
   - Issues:
     - Numbers have ".0" suffix (should be integers)
     - Percentages are decimals (should be formatted as percentages like "35%")

4. **Data Values**:
   - Values appear correct but need verification after fixing formatting

### Slide 9 (Chronic & Overcalling) Issues:

1. **Header Row Structure**:
   - Manual: `['#HCPs Missed in last 3 Months', '', '', '']` (title in first cell)
   - Generated: `['Slide 9', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3']` (column names)
   - Need to handle header row differently

2. **Row Order**:
   - Manual: GenNext, GI Maxima, GI Optima, GI Prima, GI Prospera, Metabolics, NeuroLife, Vaccines, WH- Mitera
   - Generated: Different order
   - Need to sort rows to match manual

3. **Number Formatting**:
   - Generated: `58870.0`, `4.183794802106336`
   - Manual: `57,328`, `4.30%`
   - Issues:
     - Numbers need comma separators
     - Percentages need "%" symbol and rounding (2 decimal places)

4. **DVL Values**:
   - Some slight differences (e.g., Manual: `33,447` vs Generated: `37068.0`)
   - Need to verify source of DVL data

## Required Fixes

### Priority 1 (Critical):
1. ✅ Add "Consent Require" column to Slide 4
2. ✅ Format numbers (remove .0, add commas)
3. ✅ Format percentages (add %, round to 2 decimals)
4. ✅ Sort rows to match manual order

### Priority 2 (Important):
5. ✅ Fix header row structure for Slide 9
6. ✅ Verify DVL data source for Slide 9

### Priority 3 (Nice to have):
7. ✅ Add "AIL" aggregate row to Slide 4 (if needed)
8. ✅ Improve validation script to handle formatting differences

## Next Steps

1. Fix number formatting in PPT builder
2. Add missing "Consent Require" column to Slide 4
3. Implement row sorting to match manual order
4. Fix percentage formatting
5. Re-run validation

