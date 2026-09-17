import glob
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

# ================= CONFIGURATION =================
PARENT_BADGES_DIR = "./badges"  # Root folder containing parish/vicariate subfolders
BACK_IMAGE_NAME = "back.png"  # Optional local back image name inside subfolder, or root fallback
ROOT_BACK_IMAGE = "./back.png"

# Grid setup (2 columns x 2 rows per A4 sheet)
COLS = 2
ROWS = 2
BADGES_PER_PAGE = COLS * ROWS

# Badge display dimensions on paper (in mm)
CARD_WIDTH_MM = 98
CARD_HEIGHT_MM = 138

# Gutters / Spacing between cards (in mm)
GUTTER_X_MM = 4  # Gap between columns
GUTTER_Y_MM = 6  # Gap between rows

# Border / Cutting guide styling
DRAW_BORDER = True
BORDER_COLOR = colors.HexColor("#D3D3D3")  # Light gray
BORDER_LINE_WIDTH = 0.5  # Points

# Margins & spacing for A4
PAGE_WIDTH, PAGE_HEIGHT = A4
CARD_W = CARD_WIDTH_MM * mm
CARD_H = CARD_HEIGHT_MM * mm
GUTTER_X = GUTTER_X_MM * mm
GUTTER_Y = GUTTER_Y_MM * mm

MARGIN_X = (PAGE_WIDTH - (COLS * CARD_W) - ((COLS - 1) * GUTTER_X)) / 2
MARGIN_Y = (PAGE_HEIGHT - (ROWS * CARD_H) - ((ROWS - 1) * GUTTER_Y)) / 2
# =================================================


def generate_pdf_for_folder(folder_path, folder_name):
    # Gather image files inside this specific subfolder
    exts = ("*.png", "*.jpg", "*.jpeg")
    badge_files = []
    for ext in exts:
        for file in glob.glob(os.path.join(folder_path, ext)):
            # Ignore the back.png file from being treated as a badge front
            if os.path.basename(file).lower() != BACK_IMAGE_NAME.lower():
                badge_files.append(file)

    badge_files.sort()

    if not badge_files:
        print(f"Skipping '{folder_name}': No badge images found.")
        return

    # Check if a custom back.png exists in this folder; otherwise fall back to root back.png
    subfolder_back = os.path.join(folder_path, BACK_IMAGE_NAME)
    if os.path.exists(subfolder_back):
        active_back = subfolder_back
    elif os.path.exists(ROOT_BACK_IMAGE):
        active_back = ROOT_BACK_IMAGE
    else:
        active_back = None

    # Save output PDF inside the current subfolder
    output_pdf = os.path.join(folder_path, f"{folder_name}_badges.pdf")
    c = canvas.Canvas(output_pdf, pagesize=A4)

    # Process badges in chunks of BADGES_PER_PAGE
    for i in range(0, len(badge_files), BADGES_PER_PAGE):
        batch = badge_files[i : i + BADGES_PER_PAGE]

        # ----------------- 1. FRONT PAGE -----------------
        for idx, img_path in enumerate(batch):
            col = idx % COLS
            row = idx // COLS

            x = MARGIN_X + (col * (CARD_W + GUTTER_X))
            y = PAGE_HEIGHT - MARGIN_Y - ((row + 1) * CARD_H) - (row * GUTTER_Y)

            c.drawImage(
                img_path,
                x,
                y,
                width=CARD_W,
                height=CARD_H,
                preserveAspectRatio=True,
                anchor="c",
            )

            # Draw outer border around card
            if DRAW_BORDER:
                c.setStrokeColor(BORDER_COLOR)
                c.setLineWidth(BORDER_LINE_WIDTH)
                c.rect(x, y, CARD_W, CARD_H, stroke=1, fill=0)

        c.showPage()  # Commit front page

        # ----------------- 2. BACK PAGE ------------------
        for idx, img_path in enumerate(batch):
            col = idx % COLS
            row = idx // COLS

            # Horizontal mirror for duplex long-edge printing
            mirrored_col = (COLS - 1) - col

            x = MARGIN_X + (mirrored_col * (CARD_W + GUTTER_X))
            y = PAGE_HEIGHT - MARGIN_Y - ((row + 1) * CARD_H) - (row * GUTTER_Y)

            target_back = active_back if active_back else img_path
            c.drawImage(
                target_back,
                x,
                y,
                width=CARD_W,
                height=CARD_H,
                preserveAspectRatio=True,
                anchor="c",
            )

            # Draw outer border around card
            if DRAW_BORDER:
                c.setStrokeColor(BORDER_COLOR)
                c.setLineWidth(BORDER_LINE_WIDTH)
                c.rect(x, y, CARD_W, CARD_H, stroke=1, fill=0)

        c.showPage()  # Commit back page

    c.save()
    print(f"Generated: {output_pdf} ({len(badge_files)} badges)")


def process_all_subfolders():
    if not os.path.exists(PARENT_BADGES_DIR):
        print(f"Directory '{PARENT_BADGES_DIR}' not found.")
        return

    # List all subdirectories inside badges/
    subfolders = [
        f
        for f in os.listdir(PARENT_BADGES_DIR)
        if os.path.isdir(os.path.join(PARENT_BADGES_DIR, f))
    ]

    if not subfolders:
        print(f"No subfolders found inside '{PARENT_BADGES_DIR}'.")
        return

    for folder_name in sorted(subfolders):
        folder_path = os.path.join(PARENT_BADGES_DIR, folder_name)
        generate_pdf_for_folder(folder_path, folder_name)


if __name__ == "__main__":
    process_all_subfolders()