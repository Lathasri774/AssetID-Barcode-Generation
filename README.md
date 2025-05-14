# AssetID-Barcode-Generation
!pip install reportlab pandas openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.graphics.barcode import code128
from reportlab.lib.units import cm
import pandas as pd
from google.colab import files

# Constants
TAG_WIDTH = 5.5 * cm
TAG_HEIGHT = 2.5 * cm
MARGIN_X = 1 * cm
MARGIN_Y = 1 * cm
GAP_X = 0.5 * cm
GAP_Y = 0.5 * cm
COLUMNS = 3
ROWS = 9

# Upload files
print("Upload your Excel file:")
uploaded_excel = files.upload()
excel_filename = list(uploaded_excel.keys())[0]

print("Upload your logo image:")
uploaded_logo = files.upload()
logo_filename = list(uploaded_logo.keys())[0]

# Read all sheets
def read_file(file_path):
    sheets = pd.read_excel(file_path, sheet_name=None)
    data = pd.concat(sheets.values(), ignore_index=True)
    return data

# Generate PDF
def generate_pdf(data, output_pdf):
    c = canvas.Canvas(output_pdf, pagesize=A4)
    page_width, page_height = A4
    y_offset = page_height - MARGIN_Y
    count = 0

    for _, row in data.iterrows():
        if count % (COLUMNS * ROWS) == 0 and count > 0:
            c.showPage()
            y_offset = page_height - MARGIN_Y

        col = count % COLUMNS
        row_pos = (count // COLUMNS) % ROWS
        x_pos = MARGIN_X + col * (TAG_WIDTH + GAP_X)
        y_pos = y_offset - row_pos * (TAG_HEIGHT + GAP_Y)

        # Draw outer box
        c.rect(x_pos, y_pos - TAG_HEIGHT, TAG_WIDTH, TAG_HEIGHT, stroke=1, fill=0)

        # Left section (1/4 width)
        left_w = TAG_WIDTH * 0.25
        content_x = x_pos + left_w + 0.2 * cm  # Start of right content
        logo_size = 1.5 * cm

        # Vertical center of the left section
        # Vertical center of the left section
        left_center_y = y_pos - TAG_HEIGHT / 2

        c.setFont("Helvetica", 8)
        c.drawCentredString(x_pos+0.3*cm + left_w / 2, left_center_y + 0.2 * cm, "Property Of")

        logo_size = 1.5 * cm
        c.drawImage(logo_filename, x_pos+0.3*cm + (left_w - logo_size) / 2, left_center_y - 0.8* cm,
            width=logo_size, height=logo_size, preserveAspectRatio=True)


        c.setFont("Helvetica", 6)
        c.drawCentredString(x_pos + TAG_WIDTH / 2, y_pos - TAG_HEIGHT + 0.1 * cm, "DO NOT REMOVE")

        asset_name = str(row['Asset Name']).upper()
        name_len = len(asset_name)
        if name_len <= 9:
             c.setFont("Helvetica-Bold", 6)
             c.drawString(content_x + 1.7 * cm, y_pos - 0.5 * cm, asset_name)
        elif 9 < name_len <= 16:
             c.setFont("Helvetica-Bold", 6)
             c.drawString(content_x + 1.2 * cm, y_pos - 0.5 * cm, asset_name)
        else :
             c.setFont("Helvetica-Bold", 6)
             c.drawString(content_x + 0.7 * cm, y_pos - 0.5 * cm, asset_name)


        # Barcode
        barcode = code128.Code128(str(row['Alloted Asset tag']), barWidth=0.5, barHeight=1.0 * cm)
        barcode.drawOn(c, content_x, y_pos - 1.6 * cm)

        # Serial number
        c.setFont("Helvetica", 8)
        serial_text = str(row['Alloted Asset tag'])
        c.drawString(content_x+ 1.0 * cm, y_pos - 1.9 * cm, serial_text)

        count += 1

    c.save()

# Execute
data = read_file(excel_filename)
output_pdf = "Asset_Tags_Final.pdf"
generate_pdf(data, output_pdf)

# Download
files.download(output_pdf)
