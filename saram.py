import streamlit as st
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import tempfile, os
from datetime import datetime

# --- Pure Python number to words ---
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
st.title("📑 Proforma Invoice Generator (v12.3.3 Pure Reference)")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

agg_df, order_no, made_in, loading_port, ship_date, order_of, texture, country_of_origin = [None]*8

if uploaded_file:
    raw_df = pd.read_excel(uploaded_file, header=None)

    # --- Extract shipment info ---
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
            # --- Aggregate ---
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
        bold=ParagraphStyle("bold",parent=normal,fontName="Helvetica-Bold",fontSize=10)
        small_bold=ParagraphStyle("small_bold",parent=normal,fontName="Helvetica-Bold",fontSize=8)
        label_small=ParagraphStyle("label_small",parent=normal,fontName="Helvetica-Bold",fontSize=6)
        value_small=ParagraphStyle("value_small",parent=normal,fontName="Helvetica",fontSize=6)

        elements=[]
        content_width = A4[0] - 110
        inner_width = content_width - 6
        table_width = inner_width - 6

        # proportions
        style_prop = 0.125
        item_prop = 0.185
        fabric_prop = 0.12

        left_width = table_width * (style_prop + item_prop + fabric_prop)
        right_width = inner_width - left_width

        # supplier inner columns
        style_col_width = table_width * style_prop
        supplier_inner_col2 = left_width - style_col_width

        # --- Build header_section table: title row, supplier/payment row, shipment row ---
        title_para = Paragraph("<b>PROFORMA INVOICE</b>", ParagraphStyle("title", parent=normal, alignment=1, fontSize=7))

        supplier_lines = [
            [Paragraph("Supplier Name:", label_small), Paragraph("", value_small)],   # empty second cell; company next line
            [Paragraph("", label_small), Paragraph("SAR APPARELS INDIA PVT.LTD.", small_bold)],
            [Paragraph("Address:", label_small), Paragraph("6, Picaso Bithi, Kolkata - 700017", value_small)],
            [Paragraph("Phone:", label_small), Paragraph("9817473373", value_small)],
            [Paragraph("Fax:", label_small), Paragraph("N.A.", value_small)]
        ]
        supplier_inner = Table(supplier_lines, colWidths=[style_col_width, supplier_inner_col2])
        supplier_inner.setStyle(TableStyle([
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LEFTPADDING",(0,0),(-1,-1),2),
            ("RIGHTPADDING",(0,0),(-1,-1),2),
            ("TOPPADDING",(0,0),(-1,-1),1),
            ("BOTTOMPADDING",(0,0),(-1,-1),1),
        ]))

        supplier_box = Table([[supplier_inner]], colWidths=[left_width])
        supplier_box.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),0),
                                          ("RIGHTPADDING",(0,0),(-1,-1),0),
                                          ("TOPPADDING",(0,0),(-1,-1),0),
                                          ("BOTTOMPADDING",(0,0),(-1,-1),0),]))

        payment_lines = f"PI No.: {pi_no}<br/>Landmark order Reference: {order_no}<br/>Buyer Name: {buyer_name}<br/>Brand Name: {brand_name}"
        payment_para = Paragraph(payment_lines, normal)
        consignee_lines = f"Consignee:<br/>{consignee_name}<br/>{consignee_addr}<br/>{consignee_tel}"
        consignee_para = Paragraph(consignee_lines, normal)

        payment_box = Table([[payment_para]], colWidths=[right_width])
        payment_box.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),4),
                                        ("RIGHTPADDING",(0,0),(-1,-1),4),
                                        ("TOPPADDING",(0,0),(-1,-1),4),
                                        ("BOTTOMPADDING",(0,0),(-1,-1),4),]))
        consignee_box = Table([[consignee_para]], colWidths=[right_width])
        consignee_box.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),4),
                                          ("RIGHTPADDING",(0,0),(-1,-1),4),
                                          ("TOPPADDING",(0,0),(-1,-1),4),
                                          ("BOTTOMPADDING",(0,0),(-1,-1),4),]))

        # right column stack
        right_nested = Table([[payment_box],[consignee_box]], colWidths=[right_width])
        right_nested.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),
                                          ("LEFTPADDING",(0,0),(-1,-1),0),
                                          ("RIGHTPADDING",(0,0),(-1,-1),0),
                                          ("TOPPADDING",(0,0),(-1,-1),0),
                                          ("BOTTOMPADDING",(0,0),(-1,-1),0),]))

        # shipment as two paragraphs so no internal vertical line appears
        left_ship = Paragraph(f"<b>Loading Country:</b> {made_in}<br/><b>Agreed Shipment Date:</b> {ship_date}", normal)
        right_ship = Paragraph(f"<b>Port of Loading:</b> {loading_port}<br/><b>Description of goods:</b> {order_of}", normal)

        # header_section table (3 rows)
        header_section = Table([
            [title_para, ""],
            [supplier_box, right_nested],
            [left_ship, right_ship]
        ], colWidths=[left_width, right_width])
        header_section.setStyle(TableStyle([
            ("SPAN",(0,0),(1,0)),               # title spans both cols
            ("ALIGN",(0,0),(1,0),"CENTER"),
            ("VALIGN",(0,0),(1,0),"MIDDLE"),
            ("LINEAFTER",(0,0),(0,2),0.75,colors.black),  # vertical divider through rows 0..2
            ("LEFTPADDING",(0,0),(-1,-1),0),
            ("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),0),
            ("BOTTOMPADDING",(0,0),(-1,-1),0),
        ]))

        elements.append(header_section)
        elements.append(Spacer(1,8))

        # --- Main Items Table (unchanged) ---
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

        # --- Amount in Words ---
        amount_words=amount_to_words(total_amount)
        words_table=Table([[Paragraph(f"TOTAL  US DOLLAR {amount_words}", normal)]],colWidths=[inner_width])
        words_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.black),
                                         ("FONTSIZE",(0,0),(-1,-1),8),
                                         ("LEFTPADDING",(0,0),(-1,-1),4),
                                         ("RIGHTPADDING",(0,0),(-1,-1),4),
                                         ]))
        elements.append(words_table)

        # --- Terms & Conditions ---
        terms_table=Table([[Paragraph("Terms & Conditions (if any):", normal)]],colWidths=[inner_width])
        terms_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.black),
                                         ("FONTSIZE",(0,0),(-1,-1),8),
                                         ("LEFTPADDING",(0,0),(-1,-1),4),
                                         ("RIGHTPADDING",(0,0),(-1,-1),4),
                                         ]))
        elements.append(terms_table)
        elements.append(Spacer(1,12))

        # --- Signature ---
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

        # --- Outer Frame (thinner) ---
        outer_table = Table([[e] for e in elements], colWidths=[content_width])
        outer_table.setStyle(TableStyle([
            ("GRID",(0,0),(-1,-1),0.75,colors.black),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
        ]))

        doc.build([outer_table])

        # --- Download Button ---
        with open(pdf_file, "rb") as f:
            st.download_button(
                "⬇️ Download Proforma Invoice",
                f,
                file_name="Proforma_Invoice.pdf",
                mime="application/pdf"
            )
        os.remove(pdf_file)
