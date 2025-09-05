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
st.title("📑 Proforma Invoice Generator (v12.4.3 Row1 Supplier Fix)")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

agg_df = None
order_no = made_in = loading_port = ship_date = order_of = texture = country_of_origin = None

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
        label_small=ParagraphStyle("label_small",parent=normal,fontName="Helvetica",fontSize=8)

        elements=[]
        content_width = A4[0] - 110
        inner_width = content_width - 6
        table_width = inner_width - 6

        style_prop = 0.125
        item_prop = 0.185
        fabric_prop = 0.12

        left_width = table_width * (style_prop + item_prop + fabric_prop)
        right_width = inner_width - left_width

        # Title
        title_para = Paragraph("<b>PROFORMA INVOICE</b>", ParagraphStyle("title", parent=normal, alignment=1, fontSize=7))
        elements.append(Table([[title_para]], colWidths=[content_width], style=[
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
            ("TOPPADDING",(0,0),(-1,-1),4),
            ("BOTTOMPADDING",(0,0),(-1,-1),4),
        ]))

        # --- Row 1 Left (Supplier) ---
        supplier_lines = [
            [Paragraph("Supplier Name:", label_small)],
            [Paragraph("SAR APPARELS INDIA PVT.LTD.", label_small)],
            [Paragraph("Address: 6, Picaso Bithi, Kolkata - 700017", label_small)],
            [Paragraph("Phone: 9817473373", label_small)],
            [Paragraph("Fax: N.A.", label_small)]
        ]
        supplier_inner = Table(supplier_lines, colWidths=[left_width])
        supplier_inner.setStyle(TableStyle([
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LEFTPADDING",(0,0),(-1,-1),0),
            ("RIGHTPADDING",(0,0),(-1,-1),0),
            ("TOPPADDING",(0,0),(-1,-1),1),
            ("BOTTOMPADDING",(0,0),(-1,-1),1),
        ]))
        supplier_box = Table([[supplier_inner]], colWidths=[left_width])

        # --- Row 1 Right ---
        pi_label = Paragraph("<b>No. & date of PI:</b> " + str(pi_no), label_small)
        right_top = Table([[pi_label]], colWidths=[right_width], rowHeights=[20])
        right_top.setStyle(TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),4),
            ("RIGHTPADDING",(0,0),(-1,-1),4),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ]))

        right_bottom_text = (
            f"Landmark order Reference: {order_no}<br/>"
            f"Buyer Name: {buyer_name}<br/>"
            f"Brand Name: {brand_name}"
        )
        right_bottom_para = Paragraph(right_bottom_text, label_small)
        right_bottom = Table([[right_bottom_para]], colWidths=[right_width])
        right_bottom.setStyle(TableStyle([
            ("LEFTPADDING",(0,0),(-1,-1),4),
            ("RIGHTPADDING",(0,0),(-1,-1),4),
        ]))

        right_stack = Table([[right_top],[right_bottom]], colWidths=[right_width], rowHeights=[20,None])
        right_stack.setStyle(TableStyle([
            ("LINEBELOW",(0,0),(-1,0),0.35,colors.black),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
        ]))

        # Row 1 final
        header_table = Table([[supplier_box, right_stack]], colWidths=[left_width, right_width])
        header_table.setStyle(TableStyle([
            ("LINEAFTER",(0,0),(0,0),0.75,colors.black),
        ]))

        elements.append(header_table)

        # build rest of PDF normally...
        doc.build([Table([[e] for e in elements], colWidths=[content_width])])

        with open(pdf_file, "rb") as f:
            st.download_button(
                "⬇️ Download Proforma Invoice",
                f,
                file_name="Proforma_Invoice.pdf",
                mime="application/pdf"
            )
        os.remove(pdf_file)
