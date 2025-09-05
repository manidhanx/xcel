# proforma_v12.9.4_row3row4_and_table_heights.py
import streamlit as st
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import tempfile, os
from datetime import datetime

def number_to_words(n):
    ones = ["","ONE","TWO","THREE","FOUR","FIVE","SIX","SEVEN","EIGHT","NINE",
            "TEN","ELEVEN","TWELVE","THIRTEEN","FOURTEEN","FIFTEEN","SIXTEEN",
            "SEVENTEEN","EIGHTEEN","NINETEEN"]
    tens = ["","","TWENTY","THIRTY","FORTY","FIFTY","SIXTY","SEVENTY","EIGHTY","NINETY"]
    def words(num):
        if num < 20: return ones[num]
        elif num < 100: return tens[num//10] + ("" if num%10==0 else " " + ones[num%10])
        elif num < 1000: return ones[num//100] + " HUNDRED" + ("" if num%100==0 else " " + words(num%100))
        elif num < 1_000_000: return words(num//1000) + " THOUSAND" + ("" if num%1000==0 else " " + words(num%1000))
        elif num < 1_000_000_000: return words(num//1_000_000) + " MILLION" + ("" if num%1_000_000==0 else " " + words(num%1_000_000))
        else: return str(num)
    return words(n)

def amount_to_words(amount):
    whole = int(amount)
    fraction = int(round((amount - whole) * 100))
    words = number_to_words(whole) + " DOLLARS"
    if fraction > 0:
        words += f" AND {number_to_words(fraction)} CENTS"
    return words + " ONLY"

st.set_page_config(page_title="Proforma Invoice Generator", layout="centered")
st.title("📑 Proforma Invoice Generator (v12.9.4)")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

agg_df = None
order_no = made_in = loading_port = ship_date = order_of = texture = country_of_origin = None

if uploaded_file:
    raw_df = pd.read_excel(uploaded_file, header=None)

    for i, row in raw_df.iterrows():
        for j, cell in enumerate(row):
            cell_val = str(cell).strip().lower()
            if cell_val == "order no :":
                try: order_no = row[j+2]
                except: pass
            elif cell_val == "made in country :":
                try:
                    made_in = row[j+1]
                    country_of_origin = row[j+1]
                except: pass
            elif cell_val == "loading port :":
                try: loading_port = row[j+1]
                except: pass
            elif cell_val == "agreed ship date :":
                try: ship_date = row[j+2]
                except: pass
            elif cell_val == "order of":
                try: order_of = row[j+1]
                except: pass
            elif cell_val == "texture :":
                try: texture = row[j+1]
                except: pass

    if isinstance(ship_date, (datetime, pd.Timestamp)):
        ship_date = ship_date.strftime("%d/%m/%Y")

    header_row_idx = None
    for i, row in raw_df.iterrows():
        if row.astype(str).str.strip().str.lower().eq("style").any():
            header_row_idx = i
            break

    if header_row_idx is None:
        st.error("❌ Could not find 'Style' header.")
    else:
        df = pd.read_excel(uploaded_file, header=[header_row_idx, header_row_idx+1])
        df.columns = [" ".join([str(x) for x in col if str(x)!="nan"]).strip() for col in df.columns.values]
        df = df.dropna(how="all")

        style_col = next((c for c in df.columns if str(c).strip().lower().startswith("style")), None)
        qty_col, value_col_index = None, None
        for idx, col in enumerate(df.columns):
            if "value" in str(col).lower():
                value_col_index = idx
                break
        if value_col_index and value_col_index>0:
            qty_col = df.columns[value_col_index-1]
        fob_col = next((c for c in df.columns if "fob" in str(c).lower()), None)

        if not style_col or not qty_col:
            st.error("❌ Could not detect Qty/Style column.")
        else:
            aggregated_data=[]
            for style in df[style_col].dropna().unique():
                rows=df[df[style_col]==style]
                if len(rows)>0:
                    r=rows.iloc[0]
                    desc=r.iloc[1] if len(r)>1 else ""
                    comp=r.iloc[2] if len(r)>2 else ""
                    total_qty=pd.to_numeric(rows[qty_col],errors='coerce').fillna(0).sum()
                    unit_price=0
                    if fob_col and fob_col in rows.columns:
                        prices=pd.to_numeric(rows[fob_col],errors='coerce').fillna(0)
                        nz=prices[prices>0]
                        unit_price=nz.iloc[0] if len(nz)>0 else 0
                    amount=total_qty*unit_price
                    aggregated_data.append([style,desc,texture or "Knitted","61112000",comp,
                        country_of_origin or "India",int(total_qty),f"{unit_price:.2f}",f"{amount:.2f}"])
            agg_df=pd.DataFrame(aggregated_data,columns=[
                "STYLE NO.","ITEM DESCRIPTION","FABRIC TYPE","H.S NO","COMPOSITION","ORIGIN","QTY","FOB","AMOUNT"
            ])
            st.write("### ✅ Parsed Order Data")
            st.dataframe(agg_df)

# inputs & generate
if agg_df is not None:
    st.write("### ✍️ Enter Invoice Details")
    today_str = datetime.today().strftime("%d/%m/%Y")
    pi_no = st.text_input("PI No. & Date", f"SAR/LG/XXXX Dt. {today_str}")
    consignee_name = st.text_input("Consignee Name", "RNA Resource Group Ltd - Landmark (Babyshop)")
    consignee_addr = st.text_area("Consignee Address", "P.O Box 25030, Dubai, UAE")
    consignee_tel = st.text_input("Consignee Tel/Fax", "Tel: 00971 4 8095500, Fax: 00971 4 8095555/66")
    buyer_name = st.text_input("Buyer Name", "LANDMARK GROUP")
    brand_name = st.text_input("Brand Name", "Juniors")
    payment_term_val = st.text_input("Payment Term", "T/T")

    if st.button("Generate Proforma Invoice"):
        with tempfile.NamedTemporaryFile(delete=False,suffix=".pdf") as tmp:
            pdf_file=tmp.name

        doc=SimpleDocTemplate(pdf_file,pagesize=A4,leftMargin=30,rightMargin=30,topMargin=30,bottomMargin=30)
        styles=getSampleStyleSheet()
        normal=styles["Normal"]

        # paragraph styles
        title_style = ParagraphStyle("title", parent=normal, alignment=1, fontSize=7)
        supplier_label = ParagraphStyle("supplier_label", parent=normal, fontName="Helvetica-Bold", fontSize=8)
        supplier_company = ParagraphStyle("supplier_company", parent=normal, fontName="Helvetica-Bold", fontSize=7)
        supplier_small_label = ParagraphStyle("supplier_small_label", parent=normal, fontName="Helvetica", fontSize=6)
        supplier_small_value = ParagraphStyle("supplier_small_value", parent=normal, fontName="Helvetica", fontSize=6)
        right_block_style = ParagraphStyle("right_block", parent=normal, fontName="Helvetica", fontSize=8, leading=10)
        right_top_style = ParagraphStyle("right_top", parent=normal, fontName="Helvetica-Bold", fontSize=8, leading=9)
        row1_font_size = 8
        row1_normal = ParagraphStyle("row1_normal", parent=normal, fontName="Helvetica", fontSize=row1_font_size)
        payment_header_style=ParagraphStyle("payment_header", parent=normal, fontName="Helvetica-Bold", fontSize=7)
        label_small=ParagraphStyle("label_small", parent=normal, fontName="Helvetica-Bold", fontSize=7)
        value_small=ParagraphStyle("value_small", parent=normal, fontName="Helvetica", fontSize=7, leading=8)

        elements=[]
        content_width = A4[0] - 110
        available_width = content_width - 0.5

        # column proportions (same proportions)
        props = [0.125, 0.185, 0.12, 0.10, 0.15, 0.08, 0.07, 0.08, 0.09]
        total_prop = sum(props)
        props = [p/total_prop for p in props]
        col_widths = [available_width * p for p in props]
        diff = available_width - sum(col_widths)
        if abs(diff) > 0:
            col_widths[-1] += diff

        left_width = sum(col_widths[:3])
        right_width = available_width - left_width

        origin_left_absolute = sum(col_widths[:5])
        indent_inside_right = origin_left_absolute - left_width
        items_cell_left_padding = 4
        indent_inside_right_corrected = max(0, indent_inside_right - items_cell_left_padding)
        extra_left_shift = col_widths[6] * 3
        spacer_to_origin = max(0, indent_inside_right_corrected - extra_left_shift)

        # --- Header / blocks ---
        elements.append(Table([[Paragraph("PROFORMA INVOICE", title_style)]], colWidths=[available_width], style=[
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
            ("TOPPADDING",(0,0),(-1,-1),4),
            ("BOTTOMPADDING",(0,0),(-1,-1),4),
        ]))

        supplier_title = Table([
            [Paragraph("Supplier Name:", supplier_label)],
            [Paragraph("SAR APPARELS INDIA PVT.LTD.", supplier_company)]
        ], colWidths=[left_width])
        supplier_title.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),2),("VALIGN",(0,0),(-1,-1),"TOP")]))

        supplier_contact = Table([
            [Paragraph("Address:", supplier_small_label), Paragraph("6, Picaso Bithi, Kolkata - 700017", supplier_small_value)],
            [Paragraph("Phone:", supplier_small_label), Paragraph("9817473373", supplier_small_value)],
            [Paragraph("Fax:", supplier_small_label), Paragraph("N.A.", supplier_small_value)]
        ], colWidths=[left_width*0.30, left_width*0.70])
        supplier_contact.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),("TOPPADDING",(0,0),(-1,-1),1),("BOTTOMPADDING",(0,0),(-1,-1),1),("VALIGN",(0,0),(-1,-1),"TOP")]))

        supplier_stack = Table([[supplier_title],[supplier_contact]], colWidths=[left_width])
        supplier_stack.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),6)]))

        right_top_para = Paragraph(f"No. & date of PI: {pi_no}", right_top_style)
        right_top = Table([[right_top_para]], colWidths=[right_width])
        right_top.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),2),("RIGHTPADDING",(0,0),(-1,-1),3),("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2),("VALIGN",(0,0),(-1,-1),"TOP"),("LINEBELOW",(0,0),(0,0),0.6,colors.black)]))

        right_bottom_para = Paragraph(f"<b>Landmark order Reference:</b> {order_no}<br/><b>Buyer Name:</b> {buyer_name}<br/><b>Brand Name:</b> {brand_name}", right_block_style)
        right_bottom = Table([[right_bottom_para]], colWidths=[right_width])
        right_bottom.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),2),("RIGHTPADDING",(0,0),(-1,-1),3),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),2),("VALIGN",(0,0),(-1,-1),"TOP")]))

        right_stack = Table([[right_top],[right_bottom]], colWidths=[right_width])
        right_stack.setStyle(TableStyle([("VALIGN",(0,0),(0,1),"TOP"),("LEFTPADDING",(0,0),(0,1),2),("RIGHTPADDING",(0,0),(0,1),0)]))

        consignee_para = Paragraph(f"<b>Consignee:</b><br/>{consignee_name}<br/>{consignee_addr}<br/>{consignee_tel}", row1_normal)
        consignee_box = Table([[consignee_para]], colWidths=[left_width])
        consignee_box.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),("VALIGN",(0,0),(-1,-1),"TOP")]))

        pay_label = Paragraph("Payment Term:", label_small)
        pay_value = Paragraph(payment_term_val, value_small)
        pay_term_tbl = Table([[pay_label, pay_value]], colWidths=[right_width*0.28, right_width*0.72])
        pay_term_tbl.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0)]))

        blank_row = Table([[""]], colWidths=[right_width])
        blank_row.setStyle(TableStyle([("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2)]))

        bank_heading = Paragraph("Bank Details (Including Swift/IBAN)", payment_header_style)
        bank_heading_tbl = Table([[bank_heading]], colWidths=[right_width])
        bank_heading_tbl.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),0),("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),2)]))

        colon_w = 9
        label_col_w = max(80, available_width * 0.08)
        remaining = right_width - spacer_to_origin - label_col_w - colon_w - 6
        value_col_w = max(90, remaining)

        bank_rows = []
        def add_bank_row(lbl, val):
            bank_rows.append([Paragraph(lbl, label_small), "", Paragraph(":-", label_small), Paragraph(val, value_small)])

        add_bank_row("Beneficiary", "SAR APPARELS INDIA PVT.LTD")
        add_bank_row("Account No", "2112819952")
        add_bank_row("BANK'S NAME", "KOTAK MAHINDRA BANK LTD")
        add_bank_row("BANK ADDRESS", "2 BRABOURNE ROAD, GOVIND BHAVAN, GROUND FLOOR, KOLKATA-700001")
        add_bank_row("SWIFT CODE", "KKBKINNBCPC")
        add_bank_row("BANK CODE", "0323")

        bank_inner = Table(bank_rows, colWidths=[label_col_w, spacer_to_origin, colon_w, value_col_w])
        bank_inner.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0)]))

        payment_block = Table([[pay_term_tbl],[blank_row],[bank_heading_tbl],[bank_inner]], colWidths=[right_width])
        payment_block.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),2),("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0)]))

        # ROW3 & ROW4 blocks
        left_row3_para = Paragraph(f"<b>Loading Country:</b> {made_in or ''}<br/><b>Port of Loading:</b> {loading_port or ''}<br/><b>Agreed Shipment Date:</b> {ship_date or ''}", row1_normal)
        left_row3_box = Table([[left_row3_para]], colWidths=[left_width])
        left_row3_box.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),("VALIGN",(0,0),(-1,-1),"TOP")]))

        # ROW3 RIGHT: first line L/C Advising Bank, then three blank lines, then Remarks (last)
        right_row3_para = Paragraph(
            f"<b>L/C Advising Bank:</b> (If applicable)<br/><br/><br/><br/>"
            f"<b>Remarks:</b> (if any)",
            row1_normal
        )
        right_row3_box = Table([[right_row3_para]], colWidths=[right_width])
        right_row3_box.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),("VALIGN",(0,0),(-1,-1),"TOP")]))

        left_row4_para = Paragraph(f"<b>Description of goods:</b> {order_of or 'Value Packs'}", row1_normal)
        left_row4_box = Table([[left_row4_para]], colWidths=[left_width])
        left_row4_box.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),("VALIGN",(0,0),(-1,-1),"TOP")]))

        # ROW4 RIGHT: bottom-right align CURRENCY: USD
        right_row4_para = Paragraph("CURRENCY: USD", row1_normal)
        right_row4_box = Table([[right_row4_para]], colWidths=[right_width])
        right_row4_box.setStyle(TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),4),
            ("RIGHTPADDING",(0,0),(-1,-1),4),
            ("VALIGN",(0,0),(-1,-1),"BOTTOM"),     # <-- vertical bottom
            ("ALIGN",(0,0),(-1,-1),"RIGHT")        # <-- horizontal right
        ]))

        # assemble header — keep header bottom line enabled so it sits on top of items
        # increase first two header rows' heights slightly so the header looks visually taller (3-line feel for style table's header)
        header_table = Table([
            [supplier_stack, right_stack],
            [consignee_box, payment_block],
            [left_row3_box, right_row3_box],
            [left_row4_box, right_row4_box]
        ], colWidths=[left_width, right_width], rowHeights=[28, 28, 64, 64])  # first two rows slightly taller; row3/4 increased

        header_table.setStyle(TableStyle([
            ("VALIGN",(0,0),(1,3),"TOP"),
            ("LINEAFTER",(0,0),(0,3),0.75,colors.black),
            ("LINEBELOW",(0,0),(1,0),0.35,colors.black),
            ("LINEBELOW",(0,1),(1,1),0.35,colors.black),
            ("LINEBELOW",(0,2),(1,2),0.35,colors.black),
            ("LINEBELOW",(0,3),(1,3),0.9,colors.black),
            ("LEFTPADDING",(0,0),(-1,-1),0),
            ("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),2),
            ("BOTTOMPADDING",(0,0),(-1,-1),2),
        ]))

        elements.append(header_table)
        # no spacer — header bottom line is top border of items table

        # ---------------- items table ----------------
        data=[list(agg_df.columns)]
        for _,row in agg_df.iterrows(): data.append(list(row))
        total_qty = agg_df["QTY"].sum()
        total_amount = agg_df["AMOUNT"].astype(float).sum()
        data.append(["TOTAL","","","","","",f"{int(total_qty):,}","USD",f"{total_amount:,.2f}"])

        # enforce minimum body row count (12) so table area is visually large enough for printing
        actual_body_rows = len(data) - 1  # excluding header row
        min_body_rows = 12
        body_rows_needed = max(min_body_rows, actual_body_rows)

        # header row height (bigger to create 3-line header feel)
        header_row_height = 36
        body_row_height = 20
        # construct rowHeights list: header + body_rows_needed
        row_heights = [header_row_height] + [body_row_height] * body_rows_needed

        # If actual data has more rows than body_rows_needed, report that ReportLab will auto-expand rows for actual data.
        # For Table we must pass exactly the number of heights equal to rows in table; so if actual_body_rows > min, override:
        if actual_body_rows > min_body_rows:
            # we have actual rows more than min; set row_heights to match actual length: header + actual_body_rows
            row_heights = [header_row_height] + [body_row_height] * actual_body_rows

        items_table = Table(data, colWidths=col_widths, repeatRows=1, rowHeights=row_heights)
        items_style = TableStyle([
            ("LINEBELOW",(0,0),(-1,0),0.25,colors.black),   # header row bottom thin
            ("GRID",(0,1),(-1,-1),0.25,colors.black),
            ("BACKGROUND",(0,0),(-1,0),colors.black),
            ("TEXTCOLOR",(0,0),(-1,0),colors.whitesmoke),
            ("ALIGN",(0,0),(-1,0),"CENTER"),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("FONTSIZE",(0,0),(-1,0),8),
            ("ALIGN",(0,1),(5,-1),"CENTER"),
            ("ALIGN",(6,1),(-1,-1),"RIGHT"),
            ("FONTSIZE",(0,1),(-1,-1),8),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("LEFTPADDING",(0,0),(-1,-1),4),
            ("RIGHTPADDING",(0,0),(-1,-1),4),
        ])
        items_style.add("FONTNAME",(0,len(data)-1),(-1,len(data)-1),"Helvetica-Bold")
        items_style.add("BACKGROUND",(0,len(data)-1),(-1,len(data)-1),colors.lightgrey)
        items_table.setStyle(items_style)

        elements.append(items_table)

        # amount in words
        amount_words = amount_to_words(total_amount)
        words_table = Table([[Paragraph(f"TOTAL  US DOLLAR {amount_words}", normal)]], colWidths=[available_width])
        words_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.black),("FONTSIZE",(0,0),(-1,-1),8),("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4)]))
        elements.append(words_table)

        # terms & signature
        terms_table = Table([[Paragraph("Terms & Conditions (if any):", normal)]], colWidths=[available_width])
        terms_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.black),("FONTSIZE",(0,0),(-1,-1),8),("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4)]))
        elements.append(terms_table)
        elements.append(Spacer(1,12))

        sig_img = "sarsign.png"
        sign_table = Table([[Image(sig_img,width=150,height=50), Paragraph("Signed by ………………… for RNA Resources Group Ltd - Landmark (Babyshop)", normal)]], colWidths=[0.5*available_width,0.5*available_width])
        sign_table.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"MIDDLE"),("ALIGN",(0,0),(0,0),"LEFT"),("ALIGN",(1,0),(1,0),"RIGHT"),("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),("FONTSIZE",(0,0),(-1,-1),8)]))
        elements.append(sign_table)

        outer_table = Table([[e] for e in elements], colWidths=[content_width])
        outer_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.75,colors.black),("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0)]))

        doc.build([outer_table])

        with open(pdf_file, "rb") as f:
            st.download_button("⬇️ Download Proforma Invoice", f, file_name="Proforma_Invoice.pdf", mime="application/pdf")
        os.remove(pdf_file)
