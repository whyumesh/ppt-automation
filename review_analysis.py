"""
Helper script to review analysis results
"""
import json
import os

print("="*60)
print("ANALYSIS REVIEW")
print("="*60)

# Review Excel structure
if os.path.exists("analysis/excel_info.json"):
    print("\n📊 EXCEL FILE STRUCTURE:")
    print("-" * 60)
    with open("analysis/excel_info.json", 'r', encoding='utf-8') as f:
        excel_data = json.load(f)
    
    print(f"File: {excel_data['file_name']}")
    print(f"Total Sheets: {len(excel_data['sheets'])}\n")
    
    for sheet in excel_data['sheets']:
        if 'error' not in sheet:
            print(f"  Sheet: {sheet['name']}")
            print(f"    Rows: {sheet['row_count']}, Columns: {sheet['column_count']}")
            print(f"    Columns: {', '.join([col['name'] for col in sheet['columns'][:5]])}")
            if len(sheet['columns']) > 5:
                print(f"    ... and {len(sheet['columns']) - 5} more columns")
            print()

# Review PPT structure
if os.path.exists("analysis/template_info.json"):
    print("\n📄 POWERPOINT STRUCTURE:")
    print("-" * 60)
    with open("analysis/template_info.json", 'r', encoding='utf-8') as f:
        ppt_data = json.load(f)
    
    print(f"File: {os.path.basename(ppt_data['file_path'])}")
    print(f"Total Slides: {ppt_data['slide_count']}\n")
    
    for slide in ppt_data['slides'][:5]:  # Show first 5 slides
        print(f"  Slide {slide['slide_number']}: {slide['layout_name']}")
        print(f"    Shapes: {len(slide['shapes'])}")
        for shape in slide['shapes'][:3]:  # Show first 3 shapes
            shape_type = shape.get('type', 'unknown')
            if shape_type == 'text_box':
                text_preview = shape.get('text_content', [{}])[0].get('text', '')[:50]
                print(f"      - {shape_type}: {text_preview}...")
            elif shape_type == 'table':
                table_info = shape.get('table_info', {})
                print(f"      - {shape_type}: {table_info.get('rows', 0)}x{table_info.get('columns', 0)}")
            else:
                print(f"      - {shape_type}")
        print()
    
    if len(ppt_data['slides']) > 5:
        print(f"  ... and {len(ppt_data['slides']) - 5} more slides\n")

# Review discovered rules
if os.path.exists("analysis/discovered_rules.json"):
    print("\n🔍 DISCOVERED MAPPINGS:")
    print("-" * 60)
    with open("analysis/discovered_rules.json", 'r', encoding='utf-8') as f:
        rules_data = json.load(f)
    
    print(f"Slide Mappings Found: {len(rules_data.get('slide_mappings', []))}")
    print(f"Calculation Rules: {len(rules_data.get('calculation_rules', []))}")
    print(f"Formatting Rules: {len(rules_data.get('formatting_rules', []))}")
    
    if rules_data.get('slide_mappings'):
        print("\n  Sample Mappings:")
        for mapping in rules_data['slide_mappings'][:3]:
            print(f"    Slide {mapping['slide_number']}: {len(mapping.get('shape_mappings', []))} shape mappings")

print("\n" + "="*60)
print("Next: Review these files and configure config/slides.yaml")
print("See NEXT_STEPS.md for detailed instructions")
print("="*60)

