"""
Quick script to create a template from existing PPT
"""
from src.template_extractor import TemplateExtractor
import os

ppt_path = "Data/Apr 2025/AIL LT - April'25.pptx"
template_path = "templates/template.pptx"

if not os.path.exists("templates"):
    os.makedirs("templates")

# Create a copy as template (you can manually clean it later)
import shutil
shutil.copy(ppt_path, template_path)
print(f"Template created at: {template_path}")
print("Note: You may want to manually remove content from this file, keeping only the layout structure.")

