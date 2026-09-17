import os
import re
import pandas as pd
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Table, TableStyle

# --- Paths ---
EXCEL_PATH = "vehicle/Vehicle Control Masterlist.xlsx"
PHOTOS_DIR = "vehicle/photos"          # Root folder containing your photos (or subfolders)
OUTPUT_DIR = "vehicle/vicariate_pdfs"  # Where PDFs will be saved
VALID_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}

os.makedirs(OUTPUT_DIR, exist_ok=True)

# 1. Load Excel Masterlist
df = pd.read_excel(EXCEL_PATH)
df.columns = [c.strip() for c in df.columns]
df["Control No."] = df["Control No."].astype(int)

# 2. Map photos to control numbers (handles 1..80 and second sequence 1..12 -> 81..92)
all_photos = {}

for root, _, files in os.walk(PHOTOS_DIR):
    for file in files:
        base_name, ext = os.path.splitext(file)
        if ext.lower() not in VALID_EXTS:
            continue

        full_path = os.path.join(root, file)
        numbers = [int(n) for n in re.findall(r"\d+", base_name)]
        if not numbers:
            continue

        num = numbers[0]

        # Check if the photo is from the second batch
        rel_path = os.path.relpath(full_path, PHOTOS_DIR).lower()
        is_second_batch = any(
            tag in rel_path for tag in ["part 2", "part2", "batch 2", "batch2", "set 2", "set2", "/2/", "\\2\\"]
        ) or (len(numbers) > 1 or any(sym in base_name for sym in ["(", "_", "-", "copy"]))

        if is_second_batch and 1 <= num <= 30:
            control_no = 80 + num
        else:
            control_no = num

        all_photos[control_no] = full_path

# 3. PDF Layout & Styling Setup
page_w, page_h = A4
margin = 28

styles = getSampleStyleSheet()
title_style = ParagraphStyle(
    "VicTitle",
    parent=styles["Heading2"],
    fontSize=13,
    leading=17,
    textColor=colors.HexColor("#1A237E"),
    alignment=1,
)
cell_style = ParagraphStyle(
    "CellText",
    parent=styles["Normal"],
    fontSize=8,
    leading=10,
)
cell_header = ParagraphStyle(
    "CellHeader",
    parent=styles["Normal"],
    fontSize=8.5,
    leading=11,
    fontName="Helvetica-Bold",
    textColor=colors.white,
    alignment=1,
)

# 4. Generate PDFs grouped by Vicariate
for vicariate, group in df.groupby("Vicariate", sort=False):
    safe_name = re.sub(r'[\\/*?:"<>|]', "", str(vicariate)).strip()
    pdf_path = os.path.join(OUTPUT_DIR, f"{safe_name}.pdf")

    c = canvas.Canvas(pdf_path, pagesize=A4)
    photo_count = 0

    # 1 photo per page (Portrait)
    for _, row in group.iterrows():
        ctrl_no = row["Control No."]
        if ctrl_no in all_photos:
            img_path = all_photos[ctrl_no]
            with Image.open(img_path) as img:
                img_w, img_h = img.size

            scale = min((page_w - 2 * margin) / img_w, (page_h - 2 * margin) / img_h)
            draw_w = img_w * scale
            draw_h = img_h * scale

            x = (page_w - draw_w) / 2
            y = (page_h - draw_h) / 2

            c.drawImage(img_path, x, y, width=draw_w, height=draw_h)
            c.showPage()
            photo_count += 1

    # Summary table page
    table_data = [[
        Paragraph("Control No.", cell_header),
        Paragraph("Parish / Station", cell_header),
        Paragraph("Vehicle Type", cell_header),
    ]]

    for _, row in group.iterrows():
        ctrl_txt = f"<b>#{row['Control No.']}</b>"
        parish_txt = str(row["Parish"])
        veh_col = "Type of Vehicle" if "Type of Vehicle" in row else "Type Vehicle"
        veh_txt = str(row[veh_col]) if pd.notna(row.get(veh_col)) else "-"
        if veh_txt == "nan":
            veh_txt = "-"

        table_data.append([
            Paragraph(ctrl_txt, ParagraphStyle("Ctrl", parent=cell_style, alignment=1)),
            Paragraph(parish_txt, cell_style),
            Paragraph(veh_txt, cell_style),
        ])

    table_usable_w = page_w - 2 * margin
    ctrl_col_w = 80
    remaining_w = table_usable_w - ctrl_col_w
    col_widths = [ctrl_col_w, remaining_w * 0.62, remaining_w * 0.38]

    table = Table(table_data, colWidths=col_widths)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A237E")),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D0D0D0")),
        ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#1A237E")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#F8F9FA"), colors.white]),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
    ]))

    title_p = Paragraph(f"<b>{vicariate}</b><br/><font size=9 color='#555555'>Control Number Directory</font>", title_style)
    title_w, title_h = title_p.wrap(table_usable_w, 50)
    title_p.drawOn(c, margin, page_h - margin - title_h)

    table_w, table_h = table.wrap(table_usable_w, page_h - title_h - 3 * margin)
    table.drawOn(c, margin, page_h - margin - title_h - 12 - table_h)

    c.showPage()
    c.save()
    print(f"Created: {safe_name}.pdf ({photo_count} photos + table page) [Portrait (1/page)]")