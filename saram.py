import streamlit as st
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import tempfile, os
from datetime import datetime

# --- number -> words ---
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
st.title("📑 Proforma Invoice Generator (v12.6.1)")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

agg_df = None
order_no = made_in = loading_port = ship_date = order_of = texture = country_of_origin = None

if uploaded_file:
    raw_df = pd.read_excel(uploaded_file, header=None)

    # --- Extract shipment/order info (unchanged) ---
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

    # --- Find header row ---
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

        # --- Detect columns ---
        style_col = next((c for c in df.columns if str(c).strip().lower().startswith("style")), None)
        qty_col, value_col_index = None, None
        for idx, col in enumerate(df.columns):
            if "value" in str(col).lower():
                value_col_index = idx
                break
        if value_col_index and value_col_index>0: qty_col = df.columns[value_col_index-1]
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

# --- Inputs ---
if agg_df is not None:
    st.write("### ✍️ Enter Invoice Details")
    today_str = datetime.today().strftime("%d/%m/%Y")
    pi_no = st.text_input("PI No. & Date", f"SAR/LG/XXXX Dt. {today_str}")
    consignee_name = st.text_input("Consignee Name", "RNA Resource Group Ltd - Landmark (Babyshop)")
    consignee_addr = st.text_area("Consignee Address", "P.O Box 25030, Dubai, UAE")
    consignee_tel = st.text_input("Consignee Tel/Fax", "Tel: 00971 4 8095500, Fax: 00971 4 8095555/66")
    buyer_name = st.text_input("Buyer Name", "LANDMARK GROUP")
    brand_name = st.text_input("Brand Name", "Juniors")
    payment_term = st.text_input("Payment Term", "T/T")

    if st.button("Generate Proforma Invoice"):
        with tempfile.NamedTemporaryFile(delete=False,suffix=".pdf") as tmp:
            pdf_file=tmp.name

        doc=SimpleDocTemplate(pdf_file,pagesize=A4,leftMargin=30,rightMargin=30,topMargin=30,bottomMargin=30)
        styles=getSampleStyleSheet()
        normal=styles["Normal"]

        # typography tweaks
        title_style = ParagraphStyle("title", parent=normal, alignment=1, fontSize=7)
        supplier_label = ParagraphStyle("supplier_label", parent=normal, fontName="Helvetica-Bold", fontSize=8)
        # company single-line smaller so full name fits
        supplier_company = ParagraphStyle("supplier_company", parent=normal, fontName="Helvetica-Bold", fontSize=7)
        supplier_small_label = ParagraphStyle("supplier_small_label", parent=normal, fontName="Helvetica", fontSize=6)
        supplier_small_value = ParagraphStyle("supplier_small_value", parent=normal, fontName="Helvetica", fontSize=6)
        right_block_style = ParagraphStyle("right_block", parent=normal, fontName="Helvetica", fontSize=8, leading=10)
        right_top_style = ParagraphStyle("right_top", parent=normal, fontName="Helvetica-Bold", fontSize=8, leading=9)

        payment_header_style=ParagraphStyle("payment_header", parent=normal, fontName="Helvetica-Bold", fontSize=7)
        label_small=ParagraphStyle("label_small", parent=normal, fontName="Helvetica-Bold", fontSize=6)
        value_small=ParagraphStyle("value_small", parent=normal, fontName="Helvetica", fontSize=6)

        elements=[]
        content_width = A4[0] - 110
        inner_width = content_width - 6
        table_width = inner_width - 6

        # proportions used previously
        style_prop = 0.125
        item_prop = 0.185
        fabric_prop = 0.12

        left_width = table_width * (style_prop + item_prop + fabric_prop)
        right_width = inner_width - left_width

        # compute indent so payment answers begin at ORIGIN column left edge
        before_origin_prop = style_prop + item_prop + fabric_prop + 0.10 + 0.15
        origin_left_absolute = table_width * before_origin_prop
        indent_inside_right = origin_left_absolute - left_width
        if indent_inside_right < 0:
            indent_inside_right = 0
        if indent_inside_right > (right_width * 0.9):
            indent_inside_right = right_width * 0.6

        # ---------------- TITLE ----------------
        elements.append(Table([[Paragraph("PROFORMA INVOICE", title_style)]], colWidths=[content_width], style=[
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
            ("TOPPADDING",(0,0),(-1,-1),4),
            ("BOTTOMPADDING",(0,0),(-1,-1),4),
        ]))

        # ---------------- ROW 1 LEFT - single-line company ----------------
        supplier_title = Table([
            [Paragraph("Supplier Name:", supplier_label)],
            [Paragraph("SAR APPARELS INDIA PVT.LTD.", supplier_company)]
        ], colWidths=[left_width])
        supplier_title.setStyle(TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),0),
            ("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),0),
            ("BOTTOMPADDING",(0,0),(-1,-1),2),
            ("VALIGN",(0,0),(-1,-1),"TOP")
        ]))

        # contacts below (two-col)
        supplier_contact = Table([
            [Paragraph("Address:", supplier_small_label), Paragraph("6, Picaso Bithi, Kolkata - 700017", supplier_small_value)],
            [Paragraph("Phone:", supplier_small_label), Paragraph("9817473373", supplier_small_value)],
            [Paragraph("Fax:", supplier_small_label), Paragraph("N.A.", supplier_small_value)]
        ], colWidths=[left_width*0.30, left_width*0.70])
        supplier_contact.setStyle(TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),0),
            ("RIGHTPADDING",(0,0),(-1,-1),2),
            ("TOPPADDING",(0,0),(-1,-1),1),
            ("BOTTOMPADDING",(0,0),(-1,-1),1),
            ("VALIGN",(0,0),(-1,-1),"TOP")
        ]))

        supplier_stack = Table([[supplier_title],[supplier_contact]], colWidths=[left_width])
        supplier_stack.setStyle(TableStyle([
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LEFTPADDING",(0,0),(-1,-1),0),
            ("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),0),
            ("BOTTOMPADDING",(0,0),(-1,-1),0),
        ]))

        # ---------------- ROW 1 RIGHT - top-aligned, two-row nested with divider under PI ----------------
        right_top_para = Paragraph(f"No. & date of PI: {pi_no}", right_top_style)
        right_top = Table([[right_top_para]], colWidths=[right_width])
        right_top.setStyle(TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),4),
            ("RIGHTPADDING",(0,0),(-1,-1),4),
            ("TOPPADDING",(0,0),(-1,-1),2),
            ("BOTTOMPADDING",(0,0),(-1,-1),2),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LINEBELOW",(0,0),(0,0),0.6,colors.black)
        ]))

        right_bottom_para = Paragraph(
            f"<b>Landmark order Reference:</b> {order_no}<br/>"
            f"<b>Buyer Name:</b> {buyer_name}<br/>"
            f"<b>Brand Name:</b> {brand_name}", right_block_style
        )
        right_bottom = Table([[right_bottom_para]], colWidths=[right_width])
        right_bottom.setStyle(TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),4),
            ("RIGHTPADDING",(0,0),(-1,-1),4),
            ("TOPPADDING",(0,0),(-1,-1),4),
            ("BOTTOMPADDING",(0,0),(-1,-1),2),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
        ]))

        right_stack = Table([[right_top],[right_bottom]], colWidths=[right_width], rowHeights=[None, None])
        right_stack.setStyle(TableStyle([
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LEFTPADDING",(0,0),(-1,-1),0),
            ("RIGHTPADDING",(0,0),(-1,-1),0),
        ]))

        # ---------------- ROW 2 (Consignee left, Payment terms right) ----------------
        consignee_para = Paragraph(f"<b>Consignee:</b><br/>{consignee_name}<br/>{consignee_addr}<br/>{consignee_tel}", normal)
        consignee_box = Table([[consignee_para]], colWidths=[left_width])
        consignee_box.setStyle(TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),2),
            ("RIGHTPADDING",(0,0),(-1,-1),2),
            ("TOPPADDING",(0,0),(-1,-1),4),
            ("BOTTOMPADDING",(0,0),(-1,-1),4),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
        ]))

        # Payment terms block: label + spacer + values aligned to ORIGIN start
        label_col_w = table_width * 0.08
        spacer_w = indent_inside_right
        value_col_w = right_width - label_col_w - spacer_w - 6
        if value_col_w < 50:
            value_col_w = max(50, right_width - label_col_w - 6)
            spacer_w = right_width - label_col_w - value_col_w - 6

        bank_rows = []
        bank_rows.append([Paragraph("Payment Term:", payment_header_style), "", ""])
        bank_rows.append(["", "", ""])
        bank_rows.append(["", "", ""])
        bank_pairs = [
            ("Beneficiary :-", "SAR APPARELS INDIA PVT.LTD"),
            ("Account No :-", "2112819952"),
            ("BANK'S NAME :-", "KOTAK MAHINDRA BANK LTD"),
            ("BANK ADDRESS :-", "2 BRABOURNE ROAD, GOVIND BHAVAN, GROUND FLOOR, KOLKATA-700001"),
            ("SWIFT CODE :-", "KKBKINNBCPC"),
            ("BANK CODE :-", "0323")
        ]
        for lbl, val in bank_pairs:
            bank_rows.append([Paragraph(lbl, value_small), "", Paragraph(val, value_small)])

        bank_inner = Table(bank_rows, colWidths=[label_col_w, spacer_w, value_col_w])
        bank_inner.setStyle(TableStyle([
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LEFTPADDING",(0,0),(-1,-1),2),
            ("RIGHTPADDING",(0,0),(-1,-1),2),
            ("TOPPADDING",(0,0),(-1,-1),1),
            ("BOTTOMPADDING",(0,0),(-1,-1),1),
        ]))
        payment_box = Table([[bank_inner]], colWidths=[right_width])
        payment_box.setStyle(TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),0),
            ("RIGHTPADDING",(0,0),(-1,-1),4),
            ("TOPPADDING",(0,0),(-1,-1),2),
            ("BOTTOMPADDING",(0,0),(-1,-1),2),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
        ]))

        # ---------------- ROW 3 & 4 ----------------
        left_row3_para = Paragraph(
            f"<b>Loading Country:</b> {made_in or ''}<br/>"
            f"<b>Port of Loading:</b> {loading_port or ''}<br/>"
            f"<b>Agreed Shipment Date:</b> {ship_date or ''}", normal
        )
        left_row3_box = Table([[left_row3_para]], colWidths=[left_width])
        left_row3_box.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),("VALIGN",(0,0),(-1,-1),"TOP")]))

        right_row3_para = Paragraph(
            f"<b>Loading Country:</b> {made_in or ''}<br/>"
            f"<b>L/C Advising Bank:</b> (If applicable)<br/>"
            f"<b>Remarks:</b> (if any)", normal
        )
        right_row3_box = Table([[right_row3_para]], colWidths=[right_width])
        right_row3_box.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),("VALIGN",(0,0),(-1,-1),"TOP")]))

        left_row4_para = Paragraph(f"<b>Description of goods:</b> {order_of or 'Value Packs'}", normal)
        left_row4_box = Table([[left_row4_para]], colWidths=[left_width])
        left_row4_box.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),("VALIGN",(0,0),(-1,-1),"TOP")]))

        right_row4_para = Paragraph("CURRENCY: USD", normal)
        right_row4_box = Table([[right_row4_para]], colWidths=[right_width])
        right_row4_box.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),("VALIGN",(0,0),(-1,-1),"TOP")]))

        # assemble header table (4 rows)
        header_table = Table([
            [supplier_stack, right_stack],   # row 1
            [consignee_box, payment_box],    # row 2
            [left_row3_box, right_row3_box], # row 3
            [left_row4_box, right_row4_box]  # row 4
        ], colWidths=[left_width, right_width])

        # KEY: force both columns to top-align across rows so right_stack is pinned to top
        header_table.setStyle(TableStyle([
            ("VALIGN",(0,0),(0,3),"TOP"),     # left column top
            ("VALIGN",(1,0),(1,3),"TOP"),     # right column top  <-- this enforces top alignment for whole right column
            ("LINEAFTER",(0,0),(0,3),0.75,colors.black),
            ("LINEBELOW",(0,0),(1,0),0.35,colors.black),
            ("LINEBELOW",(0,1),(1,1),0.35,colors.black),
            ("LINEBELOW",(0,2),(1,2),0.35,colors.black),
            ("LINEBELOW",(0,3),(1,3),0.5,colors.black),
            ("LEFTPADDING",(0,0),(-1,-1),0),
            ("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),2),
            ("BOTTOMPADDING",(0,0),(-1,-1),2),
        ]))

        elements.append(header_table)
        elements.append(Spacer(1,6))

        # ----------------- Items Table -----------------
        data=[list(agg_df.columns)]
        for _,row in agg_df.iterrows(): data.append(list(row))
        total_qty=agg_df["QTY"].sum()
        total_amount=agg_df["AMOUNT"].astype(float).sum()
        data.append(["TOTAL","","","","","",f"{int(total_qty):,}","USD",f"{total_amount:,.2f}"])

        col_widths = [
            table_width * style_prop,
            table_width * item_prop,
            table_width * fabric_prop,
            table_width * 0.10,
            table_width * 0.15,
            table_width * 0.08,
            table_width * 0.07,
            table_width * 0.08,
            table_width * 0.09
        ]

        table=Table(data,colWidths=col_widths,repeatRows=1)
        style=TableStyle([
            ("GRID",(0,0),(-1,-1),0.25,colors.black),
            ("BACKGROUND",(0,0),(-1,0),colors.black),
            ("TEXTCOLOR",(0,0),(-1,0),colors.whitesmoke),
            ("ALIGN",(0,0),(-1,0),"CENTER"),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("FONTSIZE",(0,0),(-1,0),7),
            ("ALIGN",(0,1),(5,-1),"CENTER"),
            ("ALIGN",(6,1),(-1,-1),"RIGHT"),
            ("FONTSIZE",(0,1),(-1,-1),8),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("LEFTPADDING",(0,0),(-1,-1),4),
            ("RIGHTPADDING",(0,0),(-1,-1),4),
        ])
        style.add("FONTNAME",(0,len(data)-1),(-1,len(data)-1),"Helvetica-Bold")
        style.add("BACKGROUND",(0,len(data)-1),(-1,len(data)-1),colors.lightgrey)
        table.setStyle(style)
        elements.append(table)

        # ----------------- Amount in Words & Footer -----------------
        amount_words=amount_to_words(total_amount)
        words_table=Table([[Paragraph(f"TOTAL  US DOLLAR {amount_words}", normal)]],colWidths=[inner_width])
        words_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.black),
                                         ("FONTSIZE",(0,0),(-1,-1),8),
                                         ("LEFTPADDING",(0,0),(-1,-1),4),
                                         ("RIGHTPADDING",(0,0),(-1,-1),4),
                                         ]))
        elements.append(words_table)

        terms_table=Table([[Paragraph("Terms & Conditions (if any):", normal)]],colWidths=[inner_width])
        terms_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.black),
                                         ("FONTSIZE",(0,0),(-1,-1),8),
                                         ("LEFTPADDING",(0,0),(-1,-1),4),
                                         ("RIGHTPADDING",(0,0),(-1,-1),4),
                                         ]))
        elements.append(terms_table)
        elements.append(Spacer(1,12))

        # signature
        sig_img = "sarsign.png"
        sign_table=Table([
            [Image(sig_img,width=150,height=50), Paragraph("Signed by ………………… for RNA Resources Group Ltd - Landmark (Babyshop)", normal)]
        ],colWidths=[0.5*inner_width,0.5*inner_width])
        sign_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.black),
                                        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                                        ("ALIGN",(0,0),(0,0),"LEFT"),
                                        ("ALIGN",(1,0),(1,0),"RIGHT"),
                                        ("FONTSIZE",(0,0),(-1,-1),8),
                                        ("LEFTPADDING",(0,0),(-1,-1),4),
                                        ("RIGHTPADDING",(0,0),(-1,-1),4),
                                        ]))
        elements.append(sign_table)

        # outer frame
        outer_table = Table([[e] for e in elements], colWidths=[content_width])
        outer_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.75,colors.black),("VALIGN",(0,0),(-1,-1),"TOP"),]))

        doc.build([outer_table])

        with open(pdf_file, "rb") as f:
            st.download_button(
                "⬇️ Download Proforma Invoice",
                f,
                file_name="Proforma_Invoice.pdf",
                mime="application/pdf"
            )
        os.remove(pdf_file)
