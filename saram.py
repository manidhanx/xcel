# proforma_v12.9.3_final_master_align_v2.py
import streamlit as st
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image, Spacer
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
    s = number_to_words(whole) + " DOLLARS"
    if fraction > 0:
        s += f" AND {number_to_words(fraction)} CENTS"
    return s + " ONLY"

st.set_page_config(page_title="Proforma Invoice Generator", layout="centered")
st.title("📑 Proforma Invoice Generator (v12.9.3)")

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
                    made_in = row[j+1]; country_of_origin = row[j+1]
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
            header_row_idx = i; break

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
                value_col_index = idx; break
        if value_col_index and value_col_index>0:
            qty_col = df.columns[value_col_index-1]
        fob_col = next((c for c in df.columns if "fob" in str(c).lower()), None)

        if not style_col or not qty_col:
            st.error("❌ Could not detect Qty/Style column.")
        else:
            aggregated_data=[]
for style in valid_styles.unique():
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
                        nz=prices[prices>0]; unit_price=nz.iloc[0] if len(nz)>0 else 0
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
        styles=getSampleStyleSheet(); normal=styles["Normal"]

        # Styles
        title_style = ParagraphStyle("title", parent=normal, alignment=1, fontSize=7)
        supplier_label = ParagraphStyle("supplier_label", parent=normal, fontName="Helvetica-Bold", fontSize=8)
        supplier_company = ParagraphStyle("supplier_company", parent=normal, fontName="Helvetica-Bold", fontSize=7)
        supplier_small_label = ParagraphStyle("supplier_small_label", parent=normal, fontName="Helvetica", fontSize=6)
        supplier_small_value = ParagraphStyle("supplier_small_value", parent=normal, fontName="Helvetica", fontSize=6)
        right_block_style = ParagraphStyle("right_block", parent=normal, fontName="Helvetica", fontSize=8, leading=10)
        right_top_style = ParagraphStyle("right_top", parent=normal, fontName="Helvetica-Bold",fontSize=8, leading=9)
        row1_normal = ParagraphStyle("row1_normal", parent=normal, fontName="Helvetica", fontSize=8)
        payment_header_style=ParagraphStyle("payment_header", parent=normal, fontName="Helvetica-Bold", fontSize=7)
        label_small=ParagraphStyle("label_small", parent=normal, fontName="Helvetica-Bold", fontSize=7)
        value_small=ParagraphStyle("value_small", parent=normal, fontName="Helvetica", fontSize=7, leading=8)
        amount_words_style = ParagraphStyle("amount_words_style", parent=normal, fontName="Helvetica-Bold", fontSize=9, leading=10)
        terms_small = ParagraphStyle("terms_small", parent=normal, fontName="Helvetica", fontSize=6, leading=7)

        elements=[]; content_width = A4[0] - 110; available_width = content_width - 0.5

        # Columns
        props = [0.125, 0.185, 0.12, 0.10, 0.15, 0.08, 0.07, 0.08, 0.09]
        total_prop = sum(props); props = [p/total_prop for p in props]
        col_widths = [available_width * p for p in props]
        diff = available_width - sum(col_widths)
        if abs(diff) > 0: col_widths[-1] += diff
        left_width = sum(col_widths[:3]); right_width = available_width - left_width

        # Align “origin” for bank answers
        origin_left_absolute = sum(col_widths[:5])
        indent_inside_right = origin_left_absolute - left_width
        items_cell_left_padding = 4
        indent_inside_right_corrected = max(0, indent_inside_right - items_cell_left_padding)
        extra_left_shift = col_widths[6] * 3
        spacer_to_origin = max(0, indent_inside_right_corrected - extra_left_shift)

        # Build master header table (title row + 4 header blocks) in one two-column table
        title_para = Paragraph("PROFORMA INVOICE", title_style)

        # Left header (supplier)
        supplier_title = Table([
            [Paragraph("Supplier Name:", supplier_label)],
            [Paragraph("SAR APPARELS INDIA PVT.LTD.", supplier_company)]
        ], colWidths=[left_width])
        supplier_title.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),
                                           ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),2),
                                           ("VALIGN",(0,0),(-1,-1),"TOP")]))

        supplier_contact = Table([
            [Paragraph("Address:", supplier_small_label), Paragraph("6, Picaso Bithi, Kolkata - 700017", supplier_small_value)],
            [Paragraph("Phone:", supplier_small_label), Paragraph("9817473373", supplier_small_value)],
            [Paragraph("Fax:", supplier_small_label), Paragraph("N.A.", supplier_small_value)]
        ], colWidths=[left_width*0.30, left_width*0.70])
        supplier_contact.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),
                                              ("TOPPADDING",(0,0),(-1,-1),1),("BOTTOMPADDING",(0,0),(-1,-1),1),
                                              ("VALIGN",(0,0),(-1,-1),"TOP")]))
        supplier_stack = Table([[supplier_title],[supplier_contact]], colWidths=[left_width])
        supplier_stack.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),6)]))

        # Right top (PI)
        right_top_para = Paragraph(f"No. & date of PI: {pi_no}", right_top_style)
        right_top = Table([[right_top_para]], colWidths=[right_width])
        right_top.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),2),("RIGHTPADDING",(0,0),(-1,-1),0),
                                       ("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2),
                                       ("VALIGN",(0,0),(-1,-1),"TOP"),("LINEBELOW",(0,0),(0,0),0.6,colors.black)]))

        right_bottom_para = Paragraph(
            f"<b>Landmark order Reference:</b> {order_no}<br/>"
            f"<b>Buyer Name:</b> {buyer_name}<br/>"
            f"<b>Brand Name:</b> {brand_name}", right_block_style)
        right_bottom = Table([[right_bottom_para]], colWidths=[right_width])
        right_bottom.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),2),("RIGHTPADDING",(0,0),(-1,-1),0),
                                          ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),2),
                                          ("VALIGN",(0,0),(-1,-1),"TOP")]))
        right_stack = Table([[right_top],[right_bottom]], colWidths=[right_width])
        right_stack.setStyle(TableStyle([("VALIGN",(0,0),(0,1),"TOP"),("LEFTPADDING",(0,0),(0,1),2),("RIGHTPADDING",(0,0),(0,1),0)]))

        # Consignee and payment blocks
        consignee_para = Paragraph(f"<b>Consignee:</b><br/>{consignee_name}<br/>{consignee_addr}<br/>{consignee_tel}", row1_normal)
        consignee_box = Table([[consignee_para]], colWidths=[left_width])
        consignee_box.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),
                                          ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),
                                          ("VALIGN",(0,0),(-1,-1),"TOP")]))

        pay_label = Paragraph("Payment Term:", label_small)
        pay_value = Paragraph(payment_term_val, value_small)
        pay_term_tbl = Table([[pay_label, pay_value]], colWidths=[right_width*0.28, right_width*0.72])
        pay_term_tbl.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0)]))
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
        bank_inner.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
                                        ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0)]))
        payment_block = Table([[pay_term_tbl],[bank_heading_tbl],[bank_inner]], colWidths=[right_width])
        payment_block.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),0),
                                           ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0)]))

        # Row 3
        left_row3_para = Paragraph(f"<b>Loading Country:</b> {made_in or ''}<br/><b>Port of Loading:</b> {loading_port or ''}<br/><b>Agreed Shipment Date:</b> {ship_date or ''}", row1_normal)
        left_row3_box = Table([[left_row3_para]], colWidths=[left_width])
        left_row3_box.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),("VALIGN",(0,0),(-1,-1),"TOP")]))

        # right row3: three breaks between lines
        right_row3_para = Paragraph(f"<b>L/C Advising Bank:</b> (If applicable)<br/><br/><br/><b>Remarks:</b> (if any)", row1_normal)
        right_row3_box = Table([[right_row3_para]], colWidths=[right_width])
        right_row3_box.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),("VALIGN",(0,0),(-1,-1),"TOP")]))

        # Row 4
        left_row4_para = Paragraph(f"<b>Description of goods:</b> {order_of or 'Value Packs'}", row1_normal)
        left_row4_box = Table([[left_row4_para]], colWidths=[left_width])
        left_row4_box.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),("VALIGN",(0,0),(-1,-1),"TOP")]))

        # move CURRENCY to start of UNIT PRICE FOB column (index 7)
        unit_price_col_index = 7
        unit_left_rel_to_rightblock = sum(col_widths[:unit_price_col_index]) - left_width
        unit_left_rel_to_rightblock = max(0, unit_left_rel_to_rightblock)
        padding_needed = unit_left_rel_to_rightblock + 2

        currency_para = Paragraph("CURRENCY: USD", row1_normal)
        row4_height = 56
        right_row4_box = Table([[currency_para]], colWidths=[right_width], rowHeights=[row4_height])
        right_row4_box.setStyle(TableStyle([
            ("ALIGN",(0,0),(0,0),"RIGHT"),
            ("VALIGN",(0,0),(0,0),"BOTTOM"),
            ("LEFTPADDING",(0,0),(0,0),padding_needed),
            ("RIGHTPADDING",(0,0),(0,0),2),
            ("TOPPADDING",(0,0),(0,0),0),("BOTTOMPADDING",(0,0),(0,0),2),
        ]))

        # Build master rows: title row spanned across both columns to keep centered, then header blocks
        master_rows = []
        master_rows.append([title_para, ""])  # we will span this row across both cols
        master_rows.append([supplier_stack, right_stack])
        master_rows.append([consignee_box, payment_block])
        master_rows.append([left_row3_box, right_row3_box])
        master_rows.append([left_row4_box, right_row4_box])

        # create master table with 2 columns; adjust rowHeights so title row small
        master_table = Table(master_rows, colWidths=[left_width, right_width],
                             rowHeights=[18, None, None, 56, row4_height])
        master_table_style = [
            ("VALIGN",(0,0),(1,4),"TOP"),
            # center divider now starts from row 1 (so it stops at the line right below title)
            ("LINEAFTER",(0,1),(0,4),0.75,colors.black),
            # underline title and row separators
            ("SPAN",(0,0),(1,0)),  # span title across both cols so it's centered
            ("ALIGN",(0,0),(1,0),"CENTER"),("VALIGN",(0,0),(1,0),"MIDDLE"),
            ("LINEBELOW",(0,0),(1,0),0.9,colors.black),  # title underline (thicker)
            ("LINEBELOW",(0,1),(1,1),0.35,colors.black),
            ("LINEBELOW",(0,2),(1,2),0.35,colors.black),
            ("LINEBELOW",(0,3),(1,3),0.35,colors.black),
            ("LINEBELOW",(0,4),(1,4),0.9,colors.black),
            ("LEFTPADDING",(0,0),(-1,-1),0),
            ("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),0),
            ("BOTTOMPADDING",(0,0),(-1,-1),0),
            ("BOTTOMPADDING",(0,4),(1,4),0),
            ("TOPPADDING",(0,4),(1,4),0),
        ]
        master_table.setStyle(TableStyle(master_table_style))

        # ---------- ITEMS TABLE ----------
        header_labels = [
            "STYLE NO.","ITEM DESCRIPTION",
            "FABRIC TYPE<br/>KNITTED /<br/>WOVEN",
            "H.S NO<br/>(8digit)",
            "COMPOSITION OF<br/>MATERIAL",
            "COUNTRY OF<br/>ORIGIN",
            "QTY","UNIT PRICE<br/>FOB","AMOUNT"
        ]
        header_par_style = ParagraphStyle("tbl_header", parent=normal, alignment=1,
                                          fontName="Helvetica-Bold", fontSize=6.5, leading=8, textColor=colors.black)
        header_row = [Paragraph(lbl, header_par_style) for lbl in header_labels]

        body_rows = [list(row) for _, row in agg_df.iterrows()]
        total_qty = agg_df["QTY"].sum()
        total_amount = agg_df["AMOUNT"].astype(float).sum()

        # keep 10 extra blank rows
        EXTRA_BLANK_ROWS = 10
        for _ in range(EXTRA_BLANK_ROWS):
            body_rows.append([""]*len(header_row))

        data = [header_row] + body_rows
        total_row = ["TOTAL","","","","",f"{int(total_qty):,}","",None,None]
        data.append(total_row)

        # Row heights: add a small extra top breathing to the first body row
        header_row_h = 40
        body_row_h = 12
        first_body_extra = 4  # extra top breathing for first body row
        total_row_h = 16

        body_count = len(body_rows)
        if body_count >= 1:
            # first body row gets extra height
            row_heights = [header_row_h, body_row_h + first_body_extra] + [body_row_h]*(body_count-1) + [total_row_h]
        else:
            row_heights = [header_row_h] + [body_row_h]*body_count + [total_row_h]

        items_table = Table(data, colWidths=col_widths, repeatRows=1, rowHeights=row_heights)
        items_style = TableStyle([
            ("GRID",(0,1),(-1,-2),0.25,colors.white),
            ("LINEBELOW",(0,0),(-1,0),0.5,colors.black),
            ("BACKGROUND",(0,0),(-1,0),colors.white),
            ("TEXTCOLOR",(0,0),(-1,0),colors.black),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("FONTSIZE",(0,0),(-1,0),6.5),
            ("FONTSIZE",(0,1),(-1,-1),7),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("LEFTPADDING",(0,0),(-1,-1),3),("RIGHTPADDING",(0,0),(-1,-1),3),
            ("TOPPADDING",(0,0),(-1,-1),0),
        ])
        ncols = len(col_widths)
        for c in range(ncols-1):
            items_style.add("LINEAFTER",(c,0),(c,0),0.5,colors.black)
            items_style.add("LINEAFTER",(c,1),(c,len(data)-2),0.25,colors.black)

        items_style.add("LINEBELOW",(0,1),(-1,1),0.25,colors.white)

        total_idx = len(data)-1
        items_style.add("SPAN",(0,total_idx),(4,total_idx))
        items_style.add("ALIGN",(0,total_idx),(4,total_idx),"CENTER")
        items_style.add("FONTNAME",(0,total_idx),(4,total_idx),"Helvetica-Bold")
        items_style.add("FONTSIZE",(0,total_idx),(4,total_idx),8)
        items_style.add("VALIGN",(0,total_idx),(4,total_idx),"MIDDLE")

        items_style.add("SPAN",(5,total_idx),(6,total_idx))
        items_style.add("ALIGN",(5,total_idx),(6,total_idx),"CENTER")
        items_style.add("FONTNAME",(5,total_idx),(6,total_idx),"Helvetica-Bold")
        items_style.add("FONTSIZE",(5,total_idx),(6,total_idx),7)
        items_style.add("VALIGN",(5,total_idx),(6,total_idx),"MIDDLE")

        items_style.add("SPAN",(7,total_idx),(8,total_idx))
        items_style.add("ALIGN",(7,total_idx),(8,total_idx),"CENTER")
        items_style.add("FONTNAME",(7,total_idx),(8,total_idx),"Helvetica-Bold")
        items_style.add("FONTSIZE",(7,total_idx),(8,total_idx),7)

        items_style.add("LINEABOVE",(0,total_idx),(-1,total_idx),0.5,colors.black)
        items_style.add("LINEBELOW",(0,total_idx),(-1,total_idx),0.5,colors.black)
        items_style.add("LINEAFTER",(4,total_idx),(4,total_idx),0.6,colors.black)
        items_style.add("LINEAFTER",(6,total_idx),(6,total_idx),0.6,colors.black)

        items_style.add("LINEABOVE",(0,0),(-1,0),0.25,colors.black)

        items_table.setStyle(items_style)

        # update total row values
        data[total_idx][7] = f"USD {total_amount:,.2f}"
        data[total_idx][8] = ""
        data[total_idx][5] = f"{int(total_qty):,}"
        data[total_idx][6] = ""

        items_table = Table(data, colWidths=col_widths, repeatRows=1, rowHeights=row_heights)
        items_table.setStyle(items_style)

        # Stack master_table and items_table flush so there's no gap between them
        stacked = Table([[master_table],[items_table]], colWidths=[available_width], rowHeights=[None, None])
        stacked.setStyle(TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
        ]))

        elements.append(stacked)

        # Amount in words
        words_para = Paragraph(f"<b>TOTAL&nbsp;&nbsp;&nbsp;US DOLLAR {amount_to_words(total_amount)}</b>", amount_words_style)
        words_table = Table([[words_para]], colWidths=[available_width])
        words_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.white),("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4)]))
        elements.append(words_table)

        # Terms
        terms_para = Paragraph("Terms & Conditions (if any):", terms_small)
        terms_table = Table([[terms_para]], colWidths=[available_width])
        terms_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.white),("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4)]))
        elements.append(terms_table)

        # Signature & footer
        sig_img = "sarsign.png"
        try:
            sign_img = Image(sig_img, width=220, height=80)
        except Exception:
            sign_img = Paragraph("", normal)
        sign_row = Table([[sign_img, ""]], colWidths=[0.5*available_width, 0.5*available_width])
        sign_row.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"MIDDLE"),("ALIGN",(0,0),(0,0),"LEFT"),
                                      ("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),
                                      ("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2)]))
        elements.append(sign_row)
        elements.append(Spacer(1,8))

        left_footer = Paragraph("Signed by ……………………. (Affix Stamp here)", ParagraphStyle("fl", parent=normal, fontSize=6))
        right_footer = Paragraph("for RNA Resources Group Ltd-Landmark (Babyshop)", ParagraphStyle("fr", parent=normal, fontSize=6, alignment=2, fontName="Helvetica-Bold"))
        footer_row = Table([[left_footer, right_footer]], colWidths=[0.5*available_width, 0.5*available_width])
        footer_row.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("ALIGN",(0,0),(0,0),"LEFT"),("ALIGN",(1,0),(1,0),"RIGHT"),
                                        ("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),
                                        ("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2)]))
        elements.append(footer_row)

        outer_table = Table([[e] for e in elements], colWidths=[content_width])
        outer_table.setStyle(TableStyle([("BOX",(0,0),(-1,-1),0.75,colors.black),("VALIGN",(0,0),(-1,-1),"TOP"),
                                         ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0)]))

        doc.build([outer_table])

        with open(pdf_file, "rb") as f:
            st.download_button("⬇️ Download Proforma Invoice", f, file_name="Proforma_Invoice.pdf", mime="application/pdf")
        os.remove(pdf_file)
