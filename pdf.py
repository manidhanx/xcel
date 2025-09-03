from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

def create_pi_pdf(filename="Correct_Pi_Generated.pdf"):
    # Setup
    doc = SimpleDocTemplate(filename, pagesize=A4,
                            rightMargin=30, leftMargin=30,
                            topMargin=30, bottomMargin=18)
    story = []
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Center", alignment=1, fontSize=11, leading=14))
    styles.add(ParagraphStyle(name="Right", alignment=2, fontSize=10, leading=13))
    styles.add(ParagraphStyle(name="Bold", fontSize=10, leading=13, spaceAfter=6, spaceBefore=6))

    # Header
    story.append(Paragraph("<b>Proforma Invoice</b>", styles["Center"]))
    story.append(Spacer(1, 12))

    # Supplier / Buyer Details Table
    supplier_buyer = [
        ["Supplier Name No. & date of PI", "SAR/LG/0148 Dt. 14-10-2024"],
        ["SAR APPARELS INDIA PVT.LTD.", ""],
        ["ADDRESS : 6, Picaso Bithi, KOLKATA - 700017.", "Landmark order Reference: CPO/47062/25"],
        ["PHONE : 9874173373", "Buyer Name: LANDMARK GROUP"],
        ["FAX : N.A.", "Brand Name: Juniors"],
        ["Consignee:- RNA Resources Group Ltd- Landmark (Babyshop),", "Payment Term: T/T"],
        ["P O Box 25030, Dubai, UAE,", ""],
        ["Tel: 00971 4 8095500,", "Fax: 00971 4 8095555/66"]
    ]
    t1 = Table(supplier_buyer, colWidths=[8*cm, 8*cm])
    t1.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(t1)
    story.append(Spacer(1, 12))

    # Bank Details
    bank = [
        ["Bank Details (Including Swift/IBAN)", "SAR APPARELS INDIA PVT.LTD"],
        ["ACCOUNT NO", "2112819952"],
        ["BANK'S NAME", "KOTAK MAHINDRA BANK LTD"],
        ["BANK ADDRESS", "2 BRABOURNE ROAD, GOVIND BHAVAN, GROUND FLOOR, KOLKATA-700001"],
        ["SWIFT CODE", "KKBKINBBCPC"],
        ["BANK CODE", "0323"],
    ]
    t2 = Table(bank, colWidths=[6*cm, 10*cm])
    t2.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.grey),
    ]))
    story.append(t2)
    story.append(Spacer(1, 12))

    # Shipment Details
    shipment = [
        ["Loading Country:", "India"],
        ["Port of loading:", "Mumbai"],
        ["Agreed Shipment Date:", "07-02-2025"],
        ["REMARKS if ANY:-", ""],
    ]
    t3 = Table(shipment, colWidths=[6*cm, 10*cm])
    t3.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
    ]))
    story.append(t3)
    story.append(Spacer(1, 12))

    # Goods Table
    goods_data = [
        ["STYLE NO.", "ITEM DESCRIPTION", "FABRIC TYPE", "H.S NO", "COMPOSITION", "COUNTRY OF ORIGIN", "QTY", "UNIT PRICE FOB", "AMOUNT"],
        ["SAV001S25", "S/L Bodysuit 7pk", "KNITTED", "61112000", "100% COTTON", "India", "4,107", "6.00", "24642.00"],
        ["SAV002S25", "S/L Bodysuit 7pk", "KNITTED", "61112000", "100% COTTON", "India", "4,593", "6.00", "27558.00"],
        ["SAV003S25", "S/L Bodysuit 7pk", "KNITTED", "61112000", "100% COTTON", "India", "4,593", "6.00", "27558.00"],
        ["", "", "", "", "", "", "", "<b>Total</b>", "79,758.00 USD"],
    ]
    t4 = Table(goods_data, repeatRows=1, colWidths=[2*cm, 3*cm, 2*cm, 2*cm, 3*cm, 2*cm, 2*cm, 2*cm, 3*cm])
    t4.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.black),
        ("ALIGN", (6, 1), (-1, -1), "RIGHT"),
    ]))
    story.append(t4)
    story.append(Spacer(1, 12))

    # Total in words
    story.append(Paragraph("TOTAL US DOLLAR SEVENTY-NINE THOUSAND SEVEN HUNDRED FIFTY-EIGHT DOLLARS", styles["Bold"]))
    story.append(Spacer(1, 24))

    # Signature
    story.append(Paragraph("Signed by …………………….(Affix Stamp here) for RNA Resources Group Ltd-Landmark (Babyshop)", styles["Normal"]))
    story.append(Spacer(1, 12))

    # Terms
    story.append(Paragraph("Terms & Conditions (If Any)", styles["Bold"]))

    # Build PDF
    doc.build(story)

create_pi_pdf()
