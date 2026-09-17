import os
import re
import pandas as pd
from PIL import Image as PILImage

# ReportLab for Master PDF
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# ============================================================
# FILE PATHS
# ============================================================

INPUT_FILE = "Parishes/Processed_Data_Cleaned.xlsx" 
CYM_INPUT_FILE = "CYM/Marian_Youth_Vigil_2026_Cleaned_School_CYM.xlsx"

OUTPUT_FOLDER = "Parishes/Sorted_Data_Delegates"

DELEGATION_HEADS_FILE = os.path.join(OUTPUT_FOLDER, "Delegation Heads.xlsx")
MASTER_PDF_PATH = os.path.join(OUTPUT_FOLDER, "Masterlist_All_Parishes.pdf")


# ============================================================
# PAGE SETTINGS — MASTER PDF
# ============================================================

PAGE_WIDTH = 8.5 * inch
PAGE_HEIGHT = 13.0 * inch
PAGE_SIZE = (PAGE_WIDTH, PAGE_HEIGHT)


# ============================================================
# COLORS
# ============================================================

COLOR_YELLOW_ACCENT = colors.HexColor("#FDB813")
COLOR_PRIMARY_BLUE = colors.HexColor("#0D6EFD")
COLOR_DARK_NAVY = colors.HexColor("#102A43")
COLOR_BG_LIGHT = colors.HexColor("#FFFBEB")
COLOR_BORDER = colors.HexColor("#CBD5E0")


# ============================================================
# BANNER
# ============================================================

def get_banner_info():
    possible_names = ["banner.png", "banner.jpg", "banner.jpeg", "banner"]
    for name in possible_names:
        if os.path.exists(name):
            try:
                with PILImage.open(name) as img:
                    orig_w, orig_h = img.size
                    calc_width = PAGE_WIDTH
                    calc_height = (orig_h / orig_w) * calc_width
                    return name, calc_width, calc_height
            except Exception:
                return name, PAGE_WIDTH, 1.25 * inch
    return None, 0, 0


BANNER_FILE, BANNER_WIDTH, BANNER_HEIGHT = get_banner_info()


# ============================================================
# FONTS
# ============================================================

FONT_REGULAR = "Helvetica"
FONT_BOLD = "Helvetica-Bold"


def register_garet_font():
    global FONT_REGULAR, FONT_BOLD
    garet_files = {
        "Garet-Regular": ["Garet-Book.ttf", "Garet-Regular.ttf", "Garet.ttf"],
        "Garet-Bold": ["Garet-Heavy.ttf", "Garet-Bold.ttf", "Garet-ExtraBold.ttf"]
    }

    reg_found = None
    for filename in garet_files["Garet-Regular"]:
        if os.path.exists(filename):
            pdfmetrics.registerFont(TTFont("Garet-Regular", filename))
            reg_found = "Garet-Regular"
            break

    bold_found = None
    for filename in garet_files["Garet-Bold"]:
        if os.path.exists(filename):
            pdfmetrics.registerFont(TTFont("Garet-Bold", filename))
            bold_found = "Garet-Bold"
            break

    if reg_found and bold_found:
        FONT_REGULAR = reg_found
        FONT_BOLD = bold_found
        print("Garet font registered successfully.")
    elif reg_found:
        FONT_REGULAR = reg_found
        FONT_BOLD = reg_found
        print("Garet font registered (Regular).")
    else:
        print("Garet font not found locally; falling back to Helvetica.")


register_garet_font()


# ============================================================
# NUMBERED CANVAS
# ============================================================

class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count):
        self.saveState()
        if BANNER_FILE:
            self.drawImage(
                BANNER_FILE,
                x=0,
                y=PAGE_HEIGHT - BANNER_HEIGHT,
                width=BANNER_WIDTH,
                height=BANNER_HEIGHT,
                preserveAspectRatio=True,
                mask="auto"
            )

        self.setFont(FONT_REGULAR, 8)
        self.setFillColor(colors.HexColor("#718096"))
        self.drawString(36, 20, "Archdiocesan Marian Youth Vigil 2026 — Official Delegation List")
        page_str = f"Page {self._pageNumber} of {page_count}"
        self.drawRightString(PAGE_WIDTH - 36, 20, page_str)
        self.restoreState()


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_name(name):
    name = str(name).strip()
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    return name.rstrip(" .")


def find_nickname_column(df):
    possible_columns = [
        "Nickname", "Nick Name", "Nick_Name", "Alias",
        "Call Sign", "Preferred Name"
    ]
    for col in possible_columns:
        if col in df.columns:
            return col

    for col in df.columns:
        norm = str(col).lower().replace("-", "").replace("_", "").replace(" ", "")
        if "nick" in norm or "alias" in norm:
            return col
    return None


def find_parish_or_school_column(df):
    exact_targets = [
        "school / cym", "school/cym", "school / cy", "school/cy",
        "school", "school name", "parish / church", "parish/church",
        "parish", "church", "organization", "school / organization"
    ]
    
    col_map = {str(col).strip().lower(): col for col in df.columns}
    for target in exact_targets:
        if target in col_map:
            return col_map[target]

    for col in df.columns:
        norm = re.sub(r'[^a-z0-9]', '', str(col).lower())
        if "school" in norm or "cym" in norm or "parish" in norm or "church" in norm:
            return col

    return None


def find_delegation_head(delegation_heads_df, vicariate_name, parish_name):
    if delegation_heads_df.empty:
        return None

    target_vic = str(vicariate_name).strip().casefold()
    target_par = str(parish_name).strip().casefold()
    matched_names = []

    for _, row in delegation_heads_df.iterrows():
        row_vic = str(row.get("Vicariate", "")).strip().casefold()
        if row_vic != target_vic:
            continue

        raw_parish = str(row.get("Parish / Church", "")).strip().casefold()
        row_parishes = [p.strip() for p in raw_parish.split(";") if p.strip()]

        if target_par in row_parishes:
            name = str(row.get("Full Name", "")).strip()
            if name and not pd.isna(name):
                matched_names.append(name)

    if matched_names:
        return " / ".join(matched_names)

    return None


# ============================================================
# PARISH STORY FOR MASTER PDF
# ============================================================

def build_parish_story(parish_name, vicariate_name, records_df, nickname_col, delegation_head_name):
    story_items = []
    content_width = PAGE_WIDTH - 72
    styles = getSampleStyleSheet()

    parish_style = ParagraphStyle(
        "ParishHeading",
        parent=styles["Heading1"],
        fontName=FONT_BOLD,
        fontSize=14.5,
        leading=17.5,
        textColor=COLOR_DARK_NAVY,
        spaceAfter=2
    )
    meta_left_style = ParagraphStyle(
        "MetaLeft",
        parent=styles["Normal"],
        fontName=FONT_BOLD,
        fontSize=9.5,
        leading=12,
        textColor=colors.HexColor("#4A5568")
    )
    meta_right_style = ParagraphStyle(
        "MetaRight",
        parent=styles["Normal"],
        fontName=FONT_BOLD,
        fontSize=9.5,
        leading=12,
        textColor=COLOR_PRIMARY_BLUE,
        alignment=TA_RIGHT
    )
    header_cell_style = ParagraphStyle(
        "HeaderCell",
        parent=styles["Normal"],
        fontName=FONT_BOLD,
        fontSize=8.5,
        leading=11,
        textColor=COLOR_DARK_NAVY
    )
    header_cell_center = ParagraphStyle(
        "HeaderCellCenter",
        parent=header_cell_style,
        alignment=TA_CENTER
    )
    body_cell_style = ParagraphStyle(
        "BodyCell",
        parent=styles["Normal"],
        fontName=FONT_REGULAR,
        fontSize=8.5,
        leading=11,
        textColor=COLOR_DARK_NAVY
    )
    body_cell_center = ParagraphStyle(
        "BodyCellCenter",
        parent=body_cell_style,
        alignment=TA_CENTER
    )

    story_items.append(Paragraph(str(parish_name).upper(), parish_style))

    meta_table_data = [[
        Paragraph(f"VICARIATE: {vicariate_name}", meta_left_style),
        Paragraph(f"TOTAL DELEGATES: {len(records_df)}", meta_right_style)
    ]]

    meta_table = Table(meta_table_data, colWidths=[350, content_width - 350])
    meta_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LINEBELOW", (0, 0), (-1, -1), 2.5, COLOR_YELLOW_ACCENT),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0)
    ]))

    story_items.append(meta_table)
    story_items.append(Spacer(1, 8))

    # Total width = 540 pt (35 + 325 + 180) to span the page width between margins
    col_widths = [35, 325, 180]
    table_data = [[
        Paragraph("#", header_cell_center),
        Paragraph("FULL NAME", header_cell_style),
        Paragraph("NICKNAME", header_cell_style)
    ]]

    for idx, (_, row) in enumerate(records_df.iterrows(), start=1):
        fullname = str(row.get("Full Name", "")).strip() or "-"
        nickname = str(row.get(nickname_col, "")).strip() or "-" if nickname_col else "-"

        table_data.append([
            Paragraph(str(idx), body_cell_center),
            Paragraph(fullname, body_cell_style),
            Paragraph(nickname, body_cell_style)
        ])

    participant_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    t_style = [
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_YELLOW_ACCENT),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, COLOR_BORDER),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6)
    ]

    for i in range(1, len(table_data)):
        bg = COLOR_BG_LIGHT if i % 2 == 0 else colors.white
        t_style.append(("BACKGROUND", (0, i), (-1, i), bg))

    participant_table.setStyle(TableStyle(t_style))
    story_items.append(participant_table)
    story_items.append(Spacer(1, 20))

    signature_title_style = ParagraphStyle(
        "SignatureTitle",
        parent=styles["Normal"],
        fontName=FONT_BOLD,
        fontSize=9.5,
        leading=12,
        textColor=COLOR_DARK_NAVY
    )
    signature_name_style = ParagraphStyle(
        "SignatureName",
        parent=styles["Normal"],
        fontName=FONT_BOLD,
        fontSize=10,
        leading=13,
        textColor=COLOR_DARK_NAVY,
        alignment=TA_CENTER
    )
    signature_label_style = ParagraphStyle(
        "SignatureLabel",
        parent=styles["Normal"],
        fontName=FONT_REGULAR,
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#718096"),
        alignment=TA_CENTER
    )

    display_name = delegation_head_name.upper() if delegation_head_name else "________________________________"

    signature_table = Table([
        [Paragraph("DELEGATION HEAD", signature_title_style)],
        [Spacer(1, 24)],
        [Paragraph(display_name, signature_name_style)],
        [Paragraph("Signature over Printed Name", signature_label_style)],
        [Spacer(1, 8)],
        [Paragraph("Date: ______________________________", signature_label_style)]
    ], colWidths=[250], hAlign="LEFT")

    signature_table.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2)
    ]))

    story_items.append(signature_table)
    return story_items


# ============================================================
# MAIN PROGRAM
# ============================================================

def main():
    print()
    print("==============================================")
    print("MARIAN YOUTH VIGIL 2026")
    print("PARISH & CYM DELEGATION PROCESSOR")
    print("==============================================")

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    # 1. Load Parish Excel
    print(f"Reading Parish Excel: {INPUT_FILE}...")
    df_parish = pd.read_excel(INPUT_FILE)
    df_parish.columns = [str(col).strip() for col in df_parish.columns]
    df_parish["Category"] = "Parish"

    parish_col = find_parish_or_school_column(df_parish)
    if parish_col and parish_col != "Parish / Church":
        df_parish["Parish / Church"] = df_parish[parish_col]

    # 2. Load CYM Excel
    df_cym = pd.DataFrame()
    if os.path.exists(CYM_INPUT_FILE):
        print(f"Reading CYM Excel: {CYM_INPUT_FILE}...")
        df_cym = pd.read_excel(CYM_INPUT_FILE)
        df_cym.columns = [str(col).strip() for col in df_cym.columns]
        df_cym["Category"] = "CYM / School"

        cym_col = find_parish_or_school_column(df_cym)
        print(f"--> Detected CYM School/Org column: '{cym_col}'")

        if cym_col:
            df_cym["Parish / Church"] = df_cym[cym_col]
        else:
            print("⚠️ WARNING: Could not find School/CYM column in CYM file!")

        print(f"Loaded {len(df_cym)} CYM rows.")

    # Combine datasets
    dfs_to_combine = [df_parish]
    if not df_cym.empty:
        dfs_to_combine.append(df_cym)

    df_combined = pd.concat(dfs_to_combine, ignore_index=True)

    for col in ["Full Name", "Parish / Church", "Vicariate"]:
        if col in df_combined.columns:
            df_combined[col] = (
                df_combined[col]
                .fillna("Not Specified")
                .astype(str)
                .str.strip()
                .replace("", "Not Specified")
            )

    # Determine Type column (Participant vs Delegation Head)
    type_col = None
    for col in df_combined.columns:
        norm = str(col).lower().replace("_", "").replace("-", "").replace(" ", "")
        if norm == "type":
            type_col = col
            break

    if type_col:
        all_participants = df_combined[
            df_combined[type_col].astype(str).str.strip().str.casefold() == "participant"
        ].copy()
        all_delegation_heads = df_combined[
            df_combined[type_col].astype(str).str.strip().str.casefold().isin([
                "delegation head", "delegationhead", "head"
            ])
        ].copy()
    else:
        print("WARNING: No Type column found. Treating all rows as participants.")
        all_participants = df_combined.copy()
        all_delegation_heads = pd.DataFrame()

    all_participants = all_participants[all_participants["Full Name"] != ""].copy()

    nickname_col = find_nickname_column(df_combined)

    if not all_delegation_heads.empty:
        all_delegation_heads.to_excel(DELEGATION_HEADS_FILE, index=False)
        print(f"Delegation Heads Excel created: {DELEGATION_HEADS_FILE}")

    # Parish individual folder Excels
    parish_participants = all_participants[all_participants["Category"] == "Parish"].copy()
    parish_groups = parish_participants.groupby(["Vicariate", "Parish / Church"], dropna=False)
    for ((vicariate, parish), p_df) in parish_groups:
        safe_vicariate = clean_name(vicariate)
        safe_parish = clean_name(parish)
        vic_folder = os.path.join(OUTPUT_FOLDER, safe_vicariate)
        os.makedirs(vic_folder, exist_ok=True)
        out_file = os.path.join(vic_folder, f"{safe_parish}.xlsx")
        p_df.to_excel(out_file, index=False)

    # Master PDF (Parishes)
    master_pdf_story = []
    sorted_parishes = parish_participants.sort_values(
        by=["Vicariate", "Parish / Church", "Full Name"],
        key=lambda col: col.astype(str).str.casefold()
    )
    grouped = sorted_parishes.groupby(["Vicariate", "Parish / Church"], sort=False)

    first_parish = True
    for ((vicariate, parish), records_df) in grouped:
        if not first_parish:
            master_pdf_story.append(PageBreak())
        first_parish = False

        delegation_head_name = find_delegation_head(all_delegation_heads, vicariate, parish)
        master_pdf_story.extend(
            build_parish_story(
                parish_name=parish,
                vicariate_name=vicariate,
                records_df=records_df,
                nickname_col=nickname_col,
                delegation_head_name=delegation_head_name
            )
        )

    if master_pdf_story:
        doc = SimpleDocTemplate(
            MASTER_PDF_PATH,
            pagesize=PAGE_SIZE,
            rightMargin=36,
            leftMargin=36,
            topMargin=(BANNER_HEIGHT + 18),
            bottomMargin=36
        )
        doc.build(master_pdf_story, canvasmaker=NumberedCanvas)
        print(f"Master PDF created: {MASTER_PDF_PATH}")

    print()
    print("==============================================")
    print("PROCESS COMPLETED SUCCESSFULLY")
    print("==============================================")


if __name__ == "__main__":
    main()