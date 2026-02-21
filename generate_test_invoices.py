"""
generate_test_invoices.py

Generates 100 fake test invoices (and mixed-type documents) as PDFs for testing
the invoice processing pipeline. Requires: pip install reportlab

Output directory: C:/Users/fulle/Documents/AI_PDF_Phase_1/test_invoices/

Document types generated:
  - Invoices (standard, multi-page)
  - Purchase Orders
  - Payment Remittances
  - Account Statements
  - Cover Sheets
  - Mixed multi-document PDFs (for grouping/boundary detection tests)

Edge cases:
  - Rotated pages (90, 180, 270 degrees)
  - Slightly skewed pages (±3–8 degrees)
  - Multi-page invoices with long line-item tables
  - Mixed document batches in one PDF
"""

import os
import random
import math
from datetime import date, timedelta
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_CENTER, TA_LEFT
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import io

# ---------------------------------------------------------------------------
# Seed for reproducibility
# ---------------------------------------------------------------------------
random.seed(42)

OUTPUT_DIR = r"C:\Users\fulle\Documents\AI_PDF_Phase_1\test_invoices"  # noqa
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------------------------------------------------------------------------
# Fake data pools
# ---------------------------------------------------------------------------
COMPANY_NAMES = [
    "Apex Industrial Supply Co.", "Blue Ridge Manufacturing LLC",
    "Cornerstone Equipment Inc.", "Delta Logistics Group",
    "Eastern Fabrication Works", "Frontier Mechanical Services",
    "Global Parts Warehouse", "Harbor Tool & Die Co.",
    "Inland Distribution Corp.", "Jetstream Components Ltd.",
    "Keystone Materials Inc.", "Lakeside Printing Solutions",
    "Meridian Office Supplies", "Northern Steel Products",
    "Omega Electrical Contractors", "Pacific Rim Trading Co.",
    "Quality Fasteners Ltd.", "Riverside Chemical Corp.",
    "Summit HVAC Systems", "Tri-State Lumber & Hardware",
    "United Precision Parts", "Valley Fluid Controls",
    "Westside Safety Products", "Xtreme Cutting Tools Inc.",
    "Yellowstone Maintenance Supply", "Zenith Electronics Corp.",
]

STREET_NAMES = [
    "Industrial Blvd", "Commerce Drive", "Enterprise Way",
    "Manufacturing Row", "Business Park Rd", "Warehouse Lane",
    "Trade Center Dr", "Distribution Ave", "Logistics Pkwy",
]

CITIES = [
    ("Springfield", "IL", "62701"), ("Riverside", "CA", "92501"),
    ("Greenville", "SC", "29601"), ("Fairview", "TX", "75069"),
    ("Madison", "WI", "53701"), ("Burlington", "VT", "05401"),
    ("Salem", "OR", "97301"), ("Auburn", "AL", "36830"),
    ("Franklin", "TN", "37064"), ("Milton", "FL", "32570"),
]

PRODUCTS = [
    ("Steel Hex Bolts 3/8\" x 1\"", "BOX", 8.50),
    ("PVC Conduit 1\" x 10'", "EA", 4.25),
    ("Shop Rags (Pack of 50)", "PK", 12.99),
    ("Safety Gloves Medium", "PR", 6.75),
    ("Lubricating Oil 1 Qt", "EA", 11.50),
    ("Wire Rope 3/8\" x 50'", "RL", 42.00),
    ("Flat Washer 1/2\" SS", "BX", 14.20),
    ("Hydraulic Hose 3/8\" x 6'", "EA", 28.95),
    ("Air Filter Element", "EA", 19.50),
    ("Drill Bit Set 29-PC", "ST", 54.00),
    ("Grinding Wheel 4.5\"", "EA", 7.80),
    ("Thread Seal Tape", "RL", 1.99),
    ("Cable Ties 100-Pack", "PK", 8.40),
    ("Hex Key Set Metric", "ST", 16.75),
    ("Spray Paint Gray 12oz", "CN", 5.99),
    ("Silicone Sealant 10oz", "EA", 9.25),
    ("Fluorescent Tube T8 48\"", "EA", 6.50),
    ("Extension Cord 25' 12AWG", "EA", 34.99),
    ("Lock Washer 5/16\" ZN", "BX", 9.10),
    ("Masking Tape 2\" x 60yd", "RL", 7.45),
    ("WD-40 Specialist 11oz", "CN", 12.00),
    ("Pipe Nipple 1\" x 6\" Blk", "EA", 3.85),
    ("Safety Glasses Clear", "PR", 4.50),
    ("Sandpaper 80 Grit 25-PK", "PK", 11.20),
    ("Ball Valve 3/4\" Brass", "EA", 18.60),
    ("Electrical Tape 3/4\" BLK", "RL", 2.75),
    ("Forklift Battery Watering Cap", "EA", 5.00),
    ("Nitrile Gloves Box 100", "BX", 22.50),
    ("Pallet Jack Wheel", "EA", 65.00),
    ("Floor Dry Absorbent 50 lb", "BG", 29.95),
]

PAYMENT_TERMS = ["Net 30", "Net 45", "Net 60", "Due on Receipt", "2/10 Net 30"]

# ---------------------------------------------------------------------------
# Helper utilities
# ---------------------------------------------------------------------------

def random_company():
    return random.choice(COMPANY_NAMES)

def random_address():
    num = random.randint(100, 9999)
    street = random.choice(STREET_NAMES)
    city, state, zipcode = random.choice(CITIES)
    return f"{num} {street}", f"{city}, {state}  {zipcode}"

def random_date(start_offset=-365, end_offset=0):
    today = date.today()
    delta = random.randint(start_offset, end_offset)
    return today + timedelta(days=delta)

def invoice_number():
    return f"INV-{random.randint(10000, 99999)}"

def po_number():
    return f"PO-{random.randint(10000, 99999)}"

def random_line_items(count=None):
    if count is None:
        count = random.randint(2, 8)
    items = []
    for _ in range(count):
        desc, unit, base_price = random.choice(PRODUCTS)
        qty = random.randint(1, 50)
        price = round(base_price * random.uniform(0.85, 1.15), 2)
        ext = round(qty * price, 2)
        items.append((desc, unit, qty, price, ext))
    return items

def get_tax_rates():
    """Return (gst_rate, pst_rate).
    ~65% GST + PST, ~15% GST only, ~20% no tax."""
    r = random.random()
    if r < 0.65:
        return 0.05, 0.07   # GST 5% + PST 7%
    elif r < 0.80:
        return 0.05, 0.0    # GST only
    return 0.0, 0.0

def fmt_currency(val):
    return f"${val:,.2f}"

def get_styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Right', alignment=TA_RIGHT, fontSize=10))
    styles.add(ParagraphStyle(name='Center', alignment=TA_CENTER, fontSize=10))
    styles.add(ParagraphStyle(name='SmallRight', alignment=TA_RIGHT, fontSize=8))
    styles.add(ParagraphStyle(name='Header', fontSize=16, fontName='Helvetica-Bold'))
    styles.add(ParagraphStyle(name='SubHeader', fontSize=11, fontName='Helvetica-Bold'))
    styles.add(ParagraphStyle(name='Small', fontSize=8))
    return styles


# ---------------------------------------------------------------------------
# Logo generator — draws a unique company logo purely in code (no image files)
# ---------------------------------------------------------------------------

# Palette of distinct brand colours
_LOGO_PALETTES = [
    ((41,  128, 185), (255, 255, 255)),   # blue / white
    ((39,  174,  96), (255, 255, 255)),   # green / white
    ((142,  68, 173), (255, 255, 255)),   # purple / white
    ((192,  57,  43), (255, 255, 255)),   # red / white
    ((243, 156,  18), (255, 255, 255)),   # orange / white
    ((26,  188, 156), (255, 255, 255)),   # teal / white
    ((52,  73,  94),  (255, 255, 255)),   # dark slate / white
    ((211,  84,   0), (255, 255, 255)),   # burnt orange / white
    ((22,  160, 133), (255, 255, 255)),   # emerald / white
    ((41,  128, 185), (241, 196,  15)),   # blue / yellow
]

# Shape styles per logo
_LOGO_SHAPES = ['rect', 'rounded', 'circle', 'diamond', 'banner']


def make_logo(company_name, width_pt=120, height_pt=45):
    """
    Generate a company logo as a reportlab Image flowable (PNG in memory).
    The logo is drawn deterministically from the company name:
      - colour pair chosen by hash
      - shape chosen by hash
      - initials (up to 3 letters) drawn in the foreground colour
    Returns a reportlab Image object sized to width_pt x height_pt points.
    """
    from PIL import Image as PILImage, ImageDraw, ImageFont

    # Deterministic choices from company name hash
    h = abs(hash(company_name))
    bg_rgb, fg_rgb = _LOGO_PALETTES[h % len(_LOGO_PALETTES)]
    shape = _LOGO_SHAPES[(h // len(_LOGO_PALETTES)) % len(_LOGO_SHAPES)]

    # Build initials: first letter of each word, max 3
    words = [w for w in company_name.replace('.', '').replace(',', '').split() if w[0].isupper()]
    initials = ''.join(w[0] for w in words)[:3].upper()

    # Render at 3x resolution for crispness, then we'll let reportlab scale down
    scale = 3
    W = int(width_pt * scale)
    H = int(height_pt * scale)
    img = PILImage.new('RGB', (W, H), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    pad = int(H * 0.06)
    x0, y0, x1, y1 = pad, pad, W - pad, H - pad

    if shape == 'rect':
        draw.rectangle([x0, y0, x1, y1], fill=bg_rgb)

    elif shape == 'rounded':
        r = int(H * 0.25)
        draw.rounded_rectangle([x0, y0, x1, y1], radius=r, fill=bg_rgb)

    elif shape == 'circle':
        # Draw a square bg first, then a circle on the left, text on right
        draw.rectangle([x0, y0, x1, y1], fill=bg_rgb)
        cx = x0 + (y1 - y0) // 2
        cy = (y0 + y1) // 2
        cr = int((y1 - y0) * 0.42)
        draw.ellipse([cx - cr, cy - cr, cx + cr, cy + cr], fill=fg_rgb)

    elif shape == 'diamond':
        draw.rectangle([x0, y0, x1, y1], fill=bg_rgb)
        cx, cy = (x0 + x1) // 2, (y0 + y1) // 2
        half = int((y1 - y0) * 0.4)
        diamond = [(cx, cy - half), (cx + half, cy), (cx, cy + half), (cx - half, cy)]
        draw.polygon(diamond, fill=fg_rgb)

    elif shape == 'banner':
        # Solid bar with a darker accent strip on the left
        draw.rectangle([x0, y0, x1, y1], fill=bg_rgb)
        accent = tuple(max(0, c - 50) for c in bg_rgb)
        draw.rectangle([x0, y0, x0 + int((x1 - x0) * 0.08), y1], fill=accent)

    # Try to load a basic font; fall back to default if unavailable
    font_size = int(H * 0.38)
    try:
        font = ImageFont.truetype("arialbd.ttf", font_size)
    except Exception:
        try:
            font = ImageFont.truetype("Arial_Bold.ttf", font_size)
        except Exception:
            font = ImageFont.load_default()

    # For circle shape, draw initials to the right of the circle
    if shape == 'circle':
        cx_circle = x0 + (y1 - y0) // 2
        text_x = cx_circle + int((y1 - y0) * 0.5)
        text_y = (y0 + y1) // 2
        # Draw initials in fg_rgb on the circle
        cr = int((y1 - y0) * 0.42)
        cx = x0 + (y1 - y0) // 2
        cy = (y0 + y1) // 2
        try:
            bb = draw.textbbox((0, 0), initials[0] if initials else 'A', font=font)
            tw = bb[2] - bb[0]; th = bb[3] - bb[1]
        except Exception:
            tw = font_size; th = font_size
        draw.text((cx - tw // 2, cy - th // 2), initials[0] if initials else 'A',
                  fill=bg_rgb, font=font)
        # Draw remaining initials to the right in bg_rgb
        rem = initials[1:] if len(initials) > 1 else ''
        if rem:
            try:
                bb2 = draw.textbbox((0, 0), rem, font=font)
                tw2 = bb2[2] - bb2[0]; th2 = bb2[3] - bb2[1]
            except Exception:
                tw2 = font_size * len(rem); th2 = font_size
            draw.text((text_x, cy - th2 // 2), rem, fill=fg_rgb, font=font)
    else:
        # Centre initials in the full box
        try:
            bb = draw.textbbox((0, 0), initials, font=font)
            tw = bb[2] - bb[0]; th = bb[3] - bb[1]
        except Exception:
            tw = font_size * len(initials); th = font_size
        tx = (W - tw) // 2
        ty = (H - th) // 2
        draw.text((tx, ty), initials, fill=fg_rgb, font=font)

    # Save to bytes buffer
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return Image(buf, width=width_pt, height=height_pt)


# ---------------------------------------------------------------------------
# Page rotation / skew helpers
# ---------------------------------------------------------------------------

def rotate_pdf_page(input_bytes, rotation_degrees):
    """
    Render the first page of a PDF (as bytes) rotated by rotation_degrees.
    Uses canvas transform so the text layer is preserved (OCR-testable).
    Returns bytes of a new single-page PDF with the page rotated.
    """
    import fitz  # PyMuPDF
    src = fitz.open(stream=input_bytes, filetype="pdf")
    page = src[0]
    page.set_rotation(rotation_degrees)
    buf = io.BytesIO()
    src.save(buf)
    return buf.getvalue()


def skew_pdf_page(input_bytes, skew_degrees):
    """
    Render the first page to a bitmap then draw it skewed onto a new PDF.
    This simulates a scanner skew — the resulting page is an image (tests OCR deskew).
    """
    import fitz
    src = fitz.open(stream=input_bytes, filetype="pdf")
    page = src[0]
    mat = fitz.Matrix(2.0, 2.0)  # 2x zoom for quality
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img_bytes = pix.tobytes("png")

    # Create a new PDF with the image drawn skewed via canvas
    buf = io.BytesIO()
    W, H = letter
    c = canvas.Canvas(buf, pagesize=letter)

    # Draw rotated/skewed image centered on page
    img = ImageReader(io.BytesIO(img_bytes))
    c.saveState()
    cx, cy = W / 2, H / 2
    c.translate(cx, cy)
    c.rotate(skew_degrees)
    scale = 0.48  # shrink to fit within page after rotation
    c.drawImage(img, -W * scale, -H * scale, W * scale * 2, H * scale * 2,
                preserveAspectRatio=True)
    c.restoreState()
    c.save()
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Ground-truth capture — populated by invoice builders, read by main()
# ---------------------------------------------------------------------------
_last_invoice_fields = {}   # holds fields from the most recently built invoice

# ---------------------------------------------------------------------------
# Document builders
# ---------------------------------------------------------------------------

def build_invoice(vendor=None, buyer=None, line_item_count=None, layout_variant=0):
    """
    Returns PDF bytes for a single invoice. layout_variant (0-2) changes visual style.
    """
    global _last_invoice_fields
    styles = get_styles()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
                            leftMargin=0.75*inch, rightMargin=0.75*inch,
                            topMargin=0.75*inch, bottomMargin=0.75*inch)

    vendor = vendor or random_company()
    buyer = buyer or random_company()
    v_addr1, v_addr2 = random_address()
    b_addr1, b_addr2 = random_address()
    inv_num = invoice_number()
    inv_date = random_date(-180, 0)
    due_date = inv_date + timedelta(days=random.choice([30, 45, 60]))
    po_ref = po_number() if random.random() > 0.3 else "N/A"
    terms = random.choice(PAYMENT_TERMS)
    items = random_line_items(line_item_count)
    subtotal = round(sum(i[4] for i in items), 2)
    gst_rate, pst_rate = get_tax_rates()
    gst = round(subtotal * gst_rate, 2)
    pst = round(subtotal * pst_rate, 2)
    total = round(subtotal + gst + pst, 2)

    _last_invoice_fields = {
        "supplier_name":  vendor,
        "customer_name":  buyer,
        "invoice_number": inv_num,
        "invoice_date":   inv_date.strftime("%m/%d/%Y"),
        "due_date":       due_date.strftime("%m/%d/%Y"),
        "po_number":      po_ref,
        "payment_terms":  terms,
        "subtotal":       fmt_currency(subtotal),
        "gst_rate":       f"{gst_rate*100:.0f}%" if gst_rate else "",
        "gst_amount":     fmt_currency(gst) if gst else "",
        "pst_rate":       f"{pst_rate*100:.0f}%" if pst_rate else "",
        "pst_amount":     fmt_currency(pst) if pst else "",
        "total_due":      fmt_currency(total),
    }

    logo = make_logo(vendor)
    story = []

    # --- Header ---
    if layout_variant == 0:
        # Logo left, INVOICE title right
        top_row = Table(
            [[logo, Paragraph("INVOICE", styles['Header'])]],
            colWidths=[1.8*inch, 5.0*inch]
        )
        top_row.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
        ]))
        story.append(top_row)
        story.append(Spacer(1, 0.08*inch))
        story.append(Paragraph(f"<b>{vendor}</b>", styles['Normal']))
        story.append(Paragraph(v_addr1, styles['Normal']))
        story.append(Paragraph(v_addr2, styles['Normal']))
        story.append(Spacer(1, 0.15*inch))

        meta_data = [
            ["Invoice #:", inv_num, "Bill To:"],
            ["Invoice Date:", inv_date.strftime("%m/%d/%Y"), buyer],
            ["Due Date:", due_date.strftime("%m/%d/%Y"), b_addr1],
            ["P.O. Number:", po_ref, b_addr2],
            ["Terms:", terms, ""],
        ]
        meta_table = Table(meta_data, colWidths=[1.3*inch, 1.5*inch, 4.0*inch])
        meta_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))
        story.append(meta_table)

    elif layout_variant == 1:
        # Logo + vendor name left, INVOICE details right
        header_data = [
            [Table([[logo, Paragraph(f"<b>{vendor}</b><br/>{v_addr1}<br/>{v_addr2}",
                                     styles['Normal'])]], colWidths=[1.8*inch, 1.7*inch]),
             Paragraph(f"<b>INVOICE</b><br/># {inv_num}<br/>Date: {inv_date.strftime('%B %d, %Y')}",
                       styles['Right'])],
        ]
        ht = Table(header_data, colWidths=[3.5*inch, 3.5*inch])
        ht.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP')]))
        story.append(ht)
        story.append(Spacer(1, 0.1*inch))
        bill_data = [["Bill To:", f"Due Date: {due_date.strftime('%m/%d/%Y')}"],
                     [buyer, f"Terms: {terms}"],
                     [b_addr1, f"P.O.: {po_ref}"],
                     [b_addr2, ""]]
        bt = Table(bill_data, colWidths=[3.5*inch, 3.5*inch])
        bt.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('TOPPADDING', (0, 0), (-1, -1), 1),
        ]))
        story.append(bt)

    else:
        # Minimal / plain style — logo sits above the single-line header
        top_row = Table(
            [[logo, Paragraph("<b>TAX INVOICE</b>", styles['Header'])]],
            colWidths=[1.8*inch, 5.0*inch]
        )
        top_row.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
        ]))
        story.append(top_row)
        story.append(Paragraph(f"From: {vendor}  |  {v_addr1}, {v_addr2}", styles['Small']))
        story.append(Paragraph(f"To: {buyer}  |  {b_addr1}, {b_addr2}", styles['Small']))
        story.append(Paragraph(
            f"Invoice No: {inv_num}   Date: {inv_date.strftime('%Y-%m-%d')}   "
            f"Due: {due_date.strftime('%Y-%m-%d')}   PO: {po_ref}   Terms: {terms}",
            styles['Small']))

    story.append(Spacer(1, 0.2*inch))

    # --- Line items table ---
    col_headers = ["Description", "Unit", "Qty", "Unit Price", "Amount"]
    table_data = [col_headers]
    for desc, unit, qty, price, ext in items:
        table_data.append([desc, unit, str(qty), fmt_currency(price), fmt_currency(ext)])

    # Subtotals
    table_data.append(["", "", "", "Subtotal:", fmt_currency(subtotal)])
    if gst > 0:
        table_data.append(["", "", "", f"GST ({gst_rate*100:.0f}%):", fmt_currency(gst)])
    if pst > 0:
        table_data.append(["", "", "", f"PST ({pst_rate*100:.0f}%):", fmt_currency(pst)])
    table_data.append(["", "", "", "TOTAL DUE:", fmt_currency(total)])

    col_widths = [3.0*inch, 0.5*inch, 0.5*inch, 1.1*inch, 1.1*inch]
    lt = Table(table_data, colWidths=col_widths)
    lt.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('ALIGN', (2, 0), (-1, -1), 'RIGHT'),
        ('ROWBACKGROUNDS', (0, 1), (-1, len(items)),
         [colors.HexColor('#f2f2f2'), colors.white]),
        ('GRID', (0, 0), (-1, len(items)), 0.5, colors.grey),
        ('LINEABOVE', (3, -1), (-1, -1), 1, colors.black),
        ('FONTNAME', (3, -1), (-1, -1), 'Helvetica-Bold'),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(lt)
    story.append(Spacer(1, 0.3*inch))

    # --- Footer notes ---
    notes = random.choice([
        "Please remit payment by the due date. Make checks payable to the vendor above.",
        "Wire transfer details available upon request. Thank you for your business!",
        "Late payments subject to 1.5% monthly finance charge.",
        "Questions? Contact accounts receivable at ar@vendor.example.com",
    ])
    story.append(Paragraph(f"<i>Note: {notes}</i>", styles['Small']))

    doc.build(story)
    return buf.getvalue()


def build_purchase_order():
    styles = get_styles()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
                            leftMargin=0.75*inch, rightMargin=0.75*inch,
                            topMargin=0.75*inch, bottomMargin=0.75*inch)

    buyer = random_company()
    vendor = random_company()
    b_addr1, b_addr2 = random_address()
    v_addr1, v_addr2 = random_address()
    po_num = po_number()
    po_date = random_date(-180, 0)
    req_delivery = po_date + timedelta(days=random.randint(7, 30))
    items = random_line_items()
    subtotal = sum(i[4] for i in items)
    total = round(subtotal, 2)

    story = []
    story.append(Paragraph("PURCHASE ORDER", styles['Header']))
    story.append(Spacer(1, 0.1*inch))

    header_data = [
        ["From (Buyer):", buyer, "P.O. Number:", po_num],
        ["", b_addr1, "P.O. Date:", po_date.strftime("%m/%d/%Y")],
        ["", b_addr2, "Req. Delivery:", req_delivery.strftime("%m/%d/%Y")],
        ["Vendor:", vendor, "Ship Via:", random.choice(["UPS Ground", "FedEx", "Will Call", "Common Carrier"])],
        ["", v_addr1, "FOB:", random.choice(["Destination", "Origin"])],
        ["", v_addr2, "Payment Terms:", random.choice(PAYMENT_TERMS)],
    ]
    ht = Table(header_data, colWidths=[1.1*inch, 2.6*inch, 1.4*inch, 1.6*inch])
    ht.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTNAME', (2, 0), (2, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.grey),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
    ]))
    story.append(ht)
    story.append(Spacer(1, 0.2*inch))

    col_headers = ["Line", "Item Description", "Unit", "Qty Ordered", "Unit Price", "Total"]
    table_data = [col_headers]
    for i, (desc, unit, qty, price, ext) in enumerate(items, 1):
        table_data.append([str(i), desc, unit, str(qty), fmt_currency(price), fmt_currency(ext)])
    table_data.append(["", "", "", "", "TOTAL:", fmt_currency(total)])

    col_widths = [0.4*inch, 2.8*inch, 0.5*inch, 0.9*inch, 1.0*inch, 1.1*inch]
    lt = Table(table_data, colWidths=col_widths)
    lt.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a5276')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (4, -1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('ALIGN', (3, 0), (-1, -1), 'RIGHT'),
        ('GRID', (0, 0), (-1, len(items)), 0.5, colors.grey),
        ('LINEABOVE', (4, -1), (-1, -1), 1, colors.black),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(lt)
    story.append(Spacer(1, 0.3*inch))
    story.append(Paragraph(
        "Authorized by: ________________________________   Date: ___________",
        styles['Normal']))
    story.append(Spacer(1, 0.1*inch))
    story.append(Paragraph(
        "<i>This Purchase Order constitutes the entire agreement between buyer and vendor "
        "for the items listed above. No verbal agreements are binding.</i>",
        styles['Small']))

    doc.build(story)
    return buf.getvalue()


def build_payment_remittance():
    styles = get_styles()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
                            leftMargin=0.75*inch, rightMargin=0.75*inch,
                            topMargin=0.75*inch, bottomMargin=0.75*inch)

    payer = random_company()
    payee = random_company()
    check_num = f"CHK-{random.randint(100000, 999999)}"
    pay_date = random_date(-90, 0)
    num_invoices = random.randint(1, 5)

    story = []
    story.append(Paragraph("PAYMENT REMITTANCE ADVICE", styles['Header']))
    story.append(Spacer(1, 0.1*inch))

    meta = [
        ["Remitted By:", payer, "Check Number:", check_num],
        ["Remitted To:", payee, "Payment Date:", pay_date.strftime("%m/%d/%Y")],
        ["", "", "Payment Method:", random.choice(["Check", "ACH", "Wire Transfer"])],
    ]
    mt = Table(meta, colWidths=[1.2*inch, 2.8*inch, 1.4*inch, 1.3*inch])
    mt.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTNAME', (2, 0), (2, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    story.append(mt)
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph("Invoice Detail:", styles['SubHeader']))
    story.append(Spacer(1, 0.1*inch))

    total_paid = 0
    table_data = [["Invoice #", "Invoice Date", "Invoice Amt", "Discount", "Amount Paid"]]
    for _ in range(num_invoices):
        inv_amt = round(random.uniform(200, 5000), 2)
        disc = round(inv_amt * random.choice([0.0, 0.0, 0.0, 0.02]), 2)
        paid = round(inv_amt - disc, 2)
        total_paid += paid
        inv_d = pay_date - timedelta(days=random.randint(1, 60))
        table_data.append([
            invoice_number(),
            inv_d.strftime("%m/%d/%Y"),
            fmt_currency(inv_amt),
            fmt_currency(disc) if disc > 0 else "-",
            fmt_currency(paid),
        ])
    table_data.append(["", "", "", "TOTAL PAYMENT:", fmt_currency(round(total_paid, 2))])

    col_widths = [1.3*inch, 1.2*inch, 1.3*inch, 1.1*inch, 1.3*inch]
    dt = Table(table_data, colWidths=col_widths)
    dt.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#117a65')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (3, -1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('ALIGN', (2, 0), (-1, -1), 'RIGHT'),
        ('GRID', (0, 0), (-1, len(table_data)-2), 0.5, colors.grey),
        ('LINEABOVE', (3, -1), (-1, -1), 1, colors.black),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(dt)

    doc.build(story)
    return buf.getvalue()


def build_account_statement():
    styles = get_styles()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
                            leftMargin=0.75*inch, rightMargin=0.75*inch,
                            topMargin=0.75*inch, bottomMargin=0.75*inch)

    vendor = random_company()
    customer = random_company()
    v_addr1, v_addr2 = random_address()
    c_addr1, c_addr2 = random_address()
    stmt_date = random_date(-30, 0)
    period_start = stmt_date - timedelta(days=30)
    acct_num = f"ACCT-{random.randint(10000, 99999)}"

    story = []
    story.append(Paragraph("ACCOUNT STATEMENT", styles['Header']))
    story.append(Spacer(1, 0.1*inch))

    hd = [
        [f"<b>{vendor}</b>", f"Account #: {acct_num}"],
        [v_addr1, f"Statement Date: {stmt_date.strftime('%m/%d/%Y')}"],
        [v_addr2, f"Period: {period_start.strftime('%m/%d/%Y')} – {stmt_date.strftime('%m/%d/%Y')}"],
        ["", ""],
        [f"Customer: {customer}", ""],
        [c_addr1, ""],
        [c_addr2, ""],
    ]
    ht = Table(hd, colWidths=[4.0*inch, 3.0*inch])
    ht.setStyle(TableStyle([
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('TOPPADDING', (0, 0), (-1, -1), 1),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
    ]))
    story.append(ht)
    story.append(Spacer(1, 0.2*inch))

    # Transaction rows
    num_tx = random.randint(3, 10)
    running_balance = round(random.uniform(0, 1000), 2)
    table_data = [["Date", "Reference", "Description", "Charges", "Credits", "Balance"]]
    for i in range(num_tx):
        tx_date = period_start + timedelta(days=random.randint(0, 30))
        is_payment = random.random() < 0.3
        if is_payment:
            amt = round(random.uniform(200, 2000), 2)
            running_balance = round(running_balance - amt, 2)
            table_data.append([
                tx_date.strftime("%m/%d/%Y"),
                check_num := f"CHK-{random.randint(100000,999999)}",
                "Payment - Thank You",
                "",
                fmt_currency(amt),
                fmt_currency(max(running_balance, 0)),
            ])
        else:
            amt = round(random.uniform(50, 3000), 2)
            running_balance = round(running_balance + amt, 2)
            table_data.append([
                tx_date.strftime("%m/%d/%Y"),
                invoice_number(),
                "Invoice",
                fmt_currency(amt),
                "",
                fmt_currency(running_balance),
            ])

    table_data.append(["", "", "", "", "Amount Due:", fmt_currency(max(running_balance, 0))])

    col_widths = [0.9*inch, 1.1*inch, 1.8*inch, 0.9*inch, 0.9*inch, 1.0*inch]
    dt = Table(table_data, colWidths=col_widths)
    dt.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#6c3483')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (4, -1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('ALIGN', (3, 0), (-1, -1), 'RIGHT'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.HexColor('#f5eef8'), colors.white]),
        ('GRID', (0, 0), (-1, -2), 0.5, colors.grey),
        ('LINEABOVE', (4, -1), (-1, -1), 1, colors.black),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(dt)
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph(
        "Please contact us at billing@vendor.example.com if you have any questions.",
        styles['Small']))

    doc.build(story)
    return buf.getvalue()


def build_cover_sheet():
    styles = get_styles()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
                            leftMargin=1.0*inch, rightMargin=1.0*inch,
                            topMargin=1.5*inch, bottomMargin=1.0*inch)

    sender = random_company()
    recipient = random_company()
    fax_date = random_date(-90, 0)
    num_pages = random.randint(2, 8)

    story = []
    story.append(Paragraph("FAX COVER SHEET", styles['Header']))
    story.append(Spacer(1, 0.3*inch))

    rows = [
        ["To:", recipient],
        ["From:", sender],
        ["Date:", fax_date.strftime("%B %d, %Y")],
        ["Pages (including cover):", str(num_pages)],
        ["Re:", random.choice([
            "Invoice enclosed for your records",
            "Purchase order follow-up",
            "Remittance advice attached",
            "Account statement – action required",
        ])],
    ]
    t = Table(rows, colWidths=[2.2*inch, 4.6*inch])
    t.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 11),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
        ('LINEBELOW', (0, -1), (-1, -1), 1, colors.black),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.4*inch))
    story.append(Paragraph("Comments:", styles['SubHeader']))
    story.append(Spacer(1, 0.1*inch))
    comments = random.choice([
        "Please process the attached invoice at your earliest convenience.",
        "Kindly acknowledge receipt and confirm payment schedule.",
        "Please review and approve the attached purchase order.",
        "This transmission contains confidential business information.",
    ])
    story.append(Paragraph(comments, styles['Normal']))
    story.append(Spacer(1, 0.5*inch))
    story.append(Paragraph(
        "CONFIDENTIALITY NOTICE: This facsimile contains confidential information "
        "intended only for the use of the individual or entity named above.",
        styles['Small']))

    doc.build(story)
    return buf.getvalue()


def build_multi_page_invoice(line_item_count=25):
    """
    Multi-page invoice. Page 1 has the full header. Continuation pages repeat a
    compact header (vendor, invoice #, page X). Totals/tax appear only on the
    last page after all line items.
    """
    global _last_invoice_fields
    styles = get_styles()
    buf = io.BytesIO()

    vendor = random_company()
    buyer = random_company()
    v_addr1, v_addr2 = random_address()
    b_addr1, b_addr2 = random_address()
    inv_num = invoice_number()
    inv_date = random_date(-180, 0)
    due_date = inv_date + timedelta(days=random.choice([30, 45, 60]))
    po_ref = po_number() if random.random() > 0.3 else "N/A"
    terms = random.choice(PAYMENT_TERMS)
    items = random_line_items(line_item_count)
    subtotal = round(sum(i[4] for i in items), 2)
    gst_rate, pst_rate = get_tax_rates()
    gst = round(subtotal * gst_rate, 2)
    pst = round(subtotal * pst_rate, 2)
    total = round(subtotal + gst + pst, 2)

    _last_invoice_fields = {
        "supplier_name":  vendor,
        "customer_name":  buyer,
        "invoice_number": inv_num,
        "invoice_date":   inv_date.strftime("%m/%d/%Y"),
        "due_date":       due_date.strftime("%m/%d/%Y"),
        "po_number":      po_ref,
        "payment_terms":  terms,
        "subtotal":       fmt_currency(subtotal),
        "gst_rate":       f"{gst_rate*100:.0f}%" if gst_rate else "",
        "gst_amount":     fmt_currency(gst) if gst else "",
        "pst_rate":       f"{pst_rate*100:.0f}%" if pst_rate else "",
        "pst_amount":     fmt_currency(pst) if pst else "",
        "total_due":      fmt_currency(total),
    }

    # Split items across pages: ~12 items fit after the header on page 1,
    # ~18 items fit on continuation pages.
    PAGE1_ITEMS = 12
    CONT_ITEMS = 18
    pages_items = []
    remaining = list(items)
    pages_items.append(remaining[:PAGE1_ITEMS])
    remaining = remaining[PAGE1_ITEMS:]
    while remaining:
        pages_items.append(remaining[:CONT_ITEMS])
        remaining = remaining[CONT_ITEMS:]
    total_pages = len(pages_items)

    col_headers = ["Description", "Unit", "Qty", "Unit Price", "Amount"]
    col_widths = [3.0*inch, 0.5*inch, 0.5*inch, 1.1*inch, 1.1*inch]

    def line_table(page_items, is_last):
        table_data = [col_headers]
        for desc, unit, qty, price, ext in page_items:
            table_data.append([desc, unit, str(qty), fmt_currency(price), fmt_currency(ext)])
        if is_last:
            table_data.append(["", "", "", "Subtotal:", fmt_currency(subtotal)])
            if gst > 0:
                table_data.append(["", "", "", f"GST ({gst_rate*100:.0f}%):", fmt_currency(gst)])
            if pst > 0:
                table_data.append(["", "", "", f"PST ({pst_rate*100:.0f}%):", fmt_currency(pst)])
            table_data.append(["", "", "", "TOTAL DUE:", fmt_currency(total)])
        n_items = len(page_items)
        style_cmds = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('ALIGN', (2, 0), (-1, -1), 'RIGHT'),
            ('ROWBACKGROUNDS', (0, 1), (-1, n_items),
             [colors.HexColor('#f2f2f2'), colors.white]),
            ('GRID', (0, 0), (-1, n_items), 0.5, colors.grey),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]
        if is_last:
            style_cmds += [
                ('LINEABOVE', (3, -1), (-1, -1), 1, colors.black),
                ('FONTNAME', (3, -1), (-1, -1), 'Helvetica-Bold'),
            ]
        t = Table(table_data, colWidths=col_widths)
        t.setStyle(TableStyle(style_cmds))
        return t

    logo = make_logo(vendor)
    logo_small = make_logo(vendor, width_pt=80, height_pt=30)
    story = []

    for page_idx, page_items in enumerate(pages_items):
        is_first = (page_idx == 0)
        is_last = (page_idx == total_pages - 1)
        page_num = page_idx + 1

        if is_first:
            # Full header on page 1 — logo left, INVOICE right
            top_row = Table(
                [[logo, Paragraph("INVOICE", styles['Header'])]],
                colWidths=[1.8*inch, 5.0*inch]
            )
            top_row.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
            ]))
            story.append(top_row)
            story.append(Spacer(1, 0.08*inch))
            story.append(Paragraph(f"<b>{vendor}</b>", styles['Normal']))
            story.append(Paragraph(v_addr1, styles['Normal']))
            story.append(Paragraph(v_addr2, styles['Normal']))
            story.append(Spacer(1, 0.15*inch))
            meta_data = [
                ["Invoice #:", inv_num, "Bill To:"],
                ["Invoice Date:", inv_date.strftime("%m/%d/%Y"), buyer],
                ["Due Date:", due_date.strftime("%m/%d/%Y"), b_addr1],
                ["P.O. Number:", po_ref, b_addr2],
                ["Terms:", terms, ""],
            ]
            meta_table = Table(meta_data, colWidths=[1.3*inch, 1.5*inch, 4.0*inch])
            meta_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ]))
            story.append(meta_table)
        else:
            # Compact continuation header — small logo + vendor | inv# | page
            cont_data = [
                [Table([[logo_small, Paragraph(f"<b>{vendor}</b>", styles['Normal'])]],
                       colWidths=[0.95*inch, 2.05*inch]),
                 Paragraph(f"<b>INVOICE</b>  #{inv_num}", styles['Right']),
                 Paragraph(f"Page {page_num} of {total_pages}", styles['Right'])],
            ]
            ct = Table(cont_data, colWidths=[3.0*inch, 2.5*inch, 1.7*inch])
            ct.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LINEBELOW', (0, 0), (-1, 0), 0.5, colors.grey),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ]))
            story.append(ct)

        story.append(Spacer(1, 0.15*inch))
        story.append(line_table(page_items, is_last))

        if is_last:
            story.append(Spacer(1, 0.3*inch))
            notes = random.choice([
                "Please remit payment by the due date. Make checks payable to the vendor above.",
                "Wire transfer details available upon request. Thank you for your business!",
                "Late payments subject to 1.5% monthly finance charge.",
                "Questions? Contact accounts receivable at ar@vendor.example.com",
            ])
            story.append(Paragraph(f"<i>Note: {notes}</i>", styles['Small']))
        else:
            story.append(Spacer(1, 0.2*inch))
            story.append(Paragraph(
                f"<i>Continued on next page... (Page {page_num} of {total_pages})</i>",
                styles['Small']))
            story.append(PageBreak())

    doc = SimpleDocTemplate(buf, pagesize=letter,
                            leftMargin=0.75*inch, rightMargin=0.75*inch,
                            topMargin=0.75*inch, bottomMargin=0.75*inch)
    doc.build(story)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Rotation / skew application
# ---------------------------------------------------------------------------

def apply_distortion(pdf_bytes, distortion_type):
    """
    distortion_type:
      'none'    - clean digital PDF (text layer intact)
      'rot90'   - 90° rotation (text layer via PyMuPDF set_rotation)
      'rot180'  - 180° rotation
      'rot270'  - 270° rotation
      'skew'    - slight skew (scanned image, tests deskew)
    """
    if distortion_type == 'none':
        return pdf_bytes
    elif distortion_type == 'rot90':
        return rotate_pdf_page(pdf_bytes, 90)
    elif distortion_type == 'rot180':
        return rotate_pdf_page(pdf_bytes, 180)
    elif distortion_type == 'rot270':
        return rotate_pdf_page(pdf_bytes, 270)
    elif distortion_type == 'skew':
        degrees = random.uniform(3, 8) * random.choice([-1, 1])
        return skew_pdf_page(pdf_bytes, degrees)
    return pdf_bytes


# ---------------------------------------------------------------------------
# Main generation
# ---------------------------------------------------------------------------

def main():
    import csv as _csv
    ground_truth_rows = []
    print(f"Generating test invoices in: {OUTPUT_DIR}")

    # -----------------------------------------------------------------------
    # 1. Individual PDFs: 100 total
    #    Distribution:
    #      55 invoices (some multi-page, various layouts)
    #      15 purchase orders
    #      10 payment remittances
    #      10 account statements
    #      10 cover sheets
    #
    #    Distortion mix:
    #      ~50% clean
    #      ~15% rot90
    #      ~10% rot180
    #      ~10% rot270
    #      ~15% skew
    # -----------------------------------------------------------------------

    distortions = (
        ['none'] * 50 +
        ['rot90'] * 15 +
        ['rot180'] * 10 +
        ['rot270'] * 10 +
        ['skew'] * 15
    )
    random.shuffle(distortions)

    doc_specs = []

    # 55 invoices (last 10 are multi-page)
    for i in range(45):
        doc_specs.append(('invoice', i + 1, random.randint(0, 2), None))
    for i in range(10):
        doc_specs.append(('invoice_multipage', 45 + i + 1, random.randint(0, 2), random.randint(18, 30)))

    # 15 purchase orders
    for i in range(15):
        doc_specs.append(('po', i + 1, 0, None))

    # 10 remittances
    for i in range(10):
        doc_specs.append(('remittance', i + 1, 0, None))

    # 10 statements
    for i in range(10):
        doc_specs.append(('statement', i + 1, 0, None))

    # 10 cover sheets
    for i in range(10):
        doc_specs.append(('cover', i + 1, 0, None))

    random.shuffle(doc_specs)

    generated = 0
    for idx, (doc_type, num, variant, extra) in enumerate(doc_specs):
        distortion = distortions[idx % len(distortions)]

        try:
            if doc_type == 'invoice':
                pdf_bytes = build_invoice(layout_variant=variant)
                prefix = "invoice"
            elif doc_type == 'invoice_multipage':
                pdf_bytes = build_multi_page_invoice(line_item_count=extra)
                prefix = "invoice_multipage"
            elif doc_type == 'po':
                pdf_bytes = build_purchase_order()
                prefix = "purchase_order"
            elif doc_type == 'remittance':
                pdf_bytes = build_payment_remittance()
                prefix = "remittance"
            elif doc_type == 'statement':
                pdf_bytes = build_account_statement()
                prefix = "statement"
            elif doc_type == 'cover':
                pdf_bytes = build_cover_sheet()
                prefix = "cover_sheet"
            else:
                continue

            pdf_bytes = apply_distortion(pdf_bytes, distortion)

            filename = f"{prefix}_{num:03d}_{distortion}.pdf"
            filepath = os.path.join(OUTPUT_DIR, filename)
            with open(filepath, 'wb') as f:
                f.write(pdf_bytes)

            # Capture ground truth for invoice types only
            if doc_type in ('invoice', 'invoice_multipage') and _last_invoice_fields:
                row = dict(_last_invoice_fields)
                row['filename'] = filename
                row['doc_type'] = doc_type
                ground_truth_rows.append(row)

            generated += 1
            print(f"  [{generated:3d}/100] {filename}")

        except Exception as e:
            print(f"  [ERROR] {doc_type} #{num}: {e}")

    # -----------------------------------------------------------------------
    # 2. Combined multi-document PDFs (for grouping/boundary tests)
    #    5 PDFs each containing 3-6 mixed document types concatenated
    # -----------------------------------------------------------------------
    print("\nGenerating mixed multi-document PDFs...")
    import fitz

    for batch_num in range(1, 6):
        builders = [
            build_invoice, build_purchase_order, build_payment_remittance,
            build_account_statement, build_cover_sheet
        ]
        num_docs = random.randint(3, 6)
        combined = fitz.open()

        for _ in range(num_docs):
            builder = random.choice(builders)
            pdf_bytes = builder()
            src = fitz.open(stream=pdf_bytes, filetype="pdf")
            combined.insert_pdf(src)

        out_path = os.path.join(OUTPUT_DIR, f"mixed_batch_{batch_num:02d}.pdf")
        combined.save(out_path)
        print(f"  mixed_batch_{batch_num:02d}.pdf  ({num_docs} documents)")

    print(f"\nDone. {generated} individual PDFs + 5 mixed-batch PDFs")
    print(f"Output: {OUTPUT_DIR}")

    # Write ground truth CSV
    csv_path = os.path.join(OUTPUT_DIR, "ground_truth.csv")
    fieldnames = [
        "filename", "doc_type",
        "supplier_name", "customer_name",
        "invoice_number", "invoice_date", "due_date",
        "po_number", "payment_terms",
        "subtotal", "gst_rate", "gst_amount", "pst_rate", "pst_amount", "total_due",
    ]
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = _csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(ground_truth_rows)
    print(f"\nGround truth CSV: {csv_path}  ({len(ground_truth_rows)} invoice rows)")


if __name__ == "__main__":
    main()
