"""
Simple validation script to compare generated PPT with manual version
"""
import sys
import os

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from validator import validate_ppt

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python validate_output.py <manual_ppt> <generated_ppt> [report_file]")
        sys.exit(1)
    
    manual_ppt = sys.argv[1]
    generated_ppt = sys.argv[2]
    report_file = sys.argv[3] if len(sys.argv) > 3 else "validation/report.json"
    
    # Create validation directory if needed
    os.makedirs(os.path.dirname(report_file) if os.path.dirname(report_file) else ".", exist_ok=True)
    
    validate_ppt(manual_ppt, generated_ppt, report_file)

