import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io
import os
import datetime as dt

# ========== CONFIG ==========
INPUT_FILE = "employee_performance_data.xlsx"
OUTPUT_FILE = f"Employee_Performance_Report_{dt.datetime.now().strftime('%Y%m%d_%H%M')}.pptx"
LOGO_PATH = None  # Optional: "logo.png" if available
TITLE_COLOR = RGBColor(0, 51, 102)
TEXT_COLOR = RGBColor(60, 60, 60)
CHART_COLOR = ['#1f77b4', '#2ca02c', '#ff7f0e']
# ============================


def create_performance_ppt(df):
    prs = Presentation()

    # Create summary slide first
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(1), Inches(9), Inches(1.5))
    tf = title_box.text_frame
    tf.text = "Employee Performance Report"
    tf.paragraphs[0].font.size = Pt(40)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = TITLE_COLOR
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER

    sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(9), Inches(1))
    sf = sub_box.text_frame
    sf.text = f"Generated on: {dt.datetime.now().strftime('%d %B %Y, %I:%M %p')}"
    sf.paragraphs[0].font.size = Pt(18)
    sf.paragraphs[0].alignment = PP_ALIGN.CENTER
    sf.paragraphs[0].font.color.rgb = TEXT_COLOR

    if LOGO_PATH and os.path.exists(LOGO_PATH):
        slide.shapes.add_picture(LOGO_PATH, Inches(8), Inches(0.5), Inches(1.5), Inches(1.5))

    # Identify performance metric columns
    metric_candidates = [c for c in df.columns if c.lower() in ['target', 'achieved', 'rating']]
    if not metric_candidates:
        print("⚠️ No numeric metric columns found — slides will be text only.")

    # Create slides per employee
    for idx, row in df.iterrows():
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # Employee title
        emp_name = str(row.get('Employee Name', f"Employee {idx+1}"))
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = emp_name
        title_frame.paragraphs[0].font.size = Pt(32)
        title_frame.paragraphs[0].font.bold = True
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        title_frame.paragraphs[0].font.color.rgb = TITLE_COLOR

        # Subheader (Designation + Department)
        designation = str(row.get('Designation', ''))
        department = str(row.get('Department', ''))
        sub_info = " | ".join(filter(None, [designation, department]))
        if sub_info:
            sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(9), Inches(0.6))
            sf = sub_box.text_frame
            sf.text = sub_info
            sf.paragraphs[0].font.size = Pt(18)
            sf.paragraphs[0].alignment = PP_ALIGN.CENTER
            sf.paragraphs[0].font.color.rgb = TEXT_COLOR

        # Performance Chart
        metrics = [m for m in ['Target', 'Achieved', 'Rating'] if m in df.columns]
        values = [row[m] for m in metrics if pd.notna(row[m])]
        if len(metrics) == len(values) and values:
            fig, ax = plt.subplots(figsize=(4, 3))
            ax.bar(metrics, values, color=CHART_COLOR[:len(values)])
            ax.set_title("Performance Overview", fontsize=12)
            ax.tick_params(axis='x', labelrotation=0)
            plt.tight_layout()

            img_stream = io.BytesIO()
            plt.savefig(img_stream, format='png', dpi=150)
            plt.close(fig)
            img_stream.seek(0)
            slide.shapes.add_picture(img_stream, Inches(3), Inches(2), Inches(4), Inches(3))

        # Remarks
        if 'Remarks' in df.columns and pd.notna(row['Remarks']):
            remark_box = slide.shapes.add_textbox(Inches(0.5), Inches(5.5), Inches(9), Inches(1))
            rf = remark_box.text_frame
            rf.text = f"Remarks: {row['Remarks']}"
            rf.paragraphs[0].font.size = Pt(16)
            rf.paragraphs[0].font.color.rgb = TEXT_COLOR

        # Optional logo on each slide
        if LOGO_PATH and os.path.exists(LOGO_PATH):
            slide.shapes.add_picture(LOGO_PATH, Inches(8.2), Inches(0.3), Inches(1.2), Inches(1.2))

    prs.save(OUTPUT_FILE)
    print(f"✅ PowerPoint generated successfully: {OUTPUT_FILE}")


# ========== MAIN EXECUTION ==========
if __name__ == "__main__":
    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(f"❌ Input file not found: {INPUT_FILE}")

    df = pd.read_excel(INPUT_FILE)
    print(f"📊 Loaded {len(df)} employee records.")
    create_performance_ppt(df)
