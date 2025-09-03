import streamlit as st
import pandas as pd
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
)
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
st.title("📑 Proforma Invoice Generator (v10 Pure)")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    raw_df = pd.read_excel(uploaded_file, header=None)

    # --- Extract header info ---
    buyer_name = "LANDMARK GROUP"
    order_no, brand, made_in, loading_port, ship_date, order_of = [None]*6
    texture, country_of_origin = None, None

    for i, row in raw_df.iterrows():
        for j, cell in enumerate(row):
            cell_val = str(cell).strip().lower()
            if cell_val == "order no :":
                try: order_no = row[j+2]
                except: pass
            elif cell_val == "brand :":
                try: brand = row[j+1]
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

    # --- Detect header row ---
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

        # --- Detect key cols ---
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
            aggregated_data = []
            for style in df[style_col].dropna().unique():
                rows = df[df[style_col]==style]
                if len(rows)>0:
                    r = rows.iloc[0]
                    desc = r.iloc[1] if len(r)>1 else ""
                    comp = r.iloc[2] if len(r)>2 else ""
                    total_qty = pd.to_numeric(rows[qty_col], errors='coerce').fillna(0).sum()
                    unit_price = 0
                    if fob_col and fob_col in rows.columns:
                        prices = pd.to_numeric(rows[fob_col], errors='coerce').fillna(0)
                        nz = prices[prices>0]
                        unit_price = nz.iloc[0] if len(nz)>0 else 0
                    amount = total_qty*unit_price
                    aggregated_data.append([style, desc, texture or "Knitted", "61112000", comp, country_of_origin or "India",
                                            int(total_qty), f"{unit_price:.2f}", f"{amount:.2f}"])

            agg_df = pd.DataFrame(aggregated_data, columns=[
                "STYLE NO.","ITEM DESCRIPTION","FABRIC TYPE","H.S NO","COMPOSITION","ORIGIN",
                "QTY","FOB","AMOUNT"
            ])
            st.dataframe(agg_df)

            if st.button("Generate PI PDF"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    pdf_file = tmp.name

                doc = SimpleDocTemplate(pdf_file, pagesize=A4, leftMargin=30, rightMargin=30, topMargin=30, bottomMargin=30)
                styles = getSampleStyleSheet()
                normal = styles["Normal"]
                bold = ParagraphStyle("bold", parent=normal, fontName="Helvetica-Bold")

                elements=[]

                # --- Title ---
                elements.append(Paragraph("<b><font size=14>PROFORMA INVOICE</font></b>", ParagraphStyle("center", alignment=1)))
                elements.append(Spacer(1,12))

                # --- Supplier details ---
                supplier = [
                    "<b>Supplier Name:</b> SAR APPARELS INDIA PVT.LTD.",
                    "Address: 6, Picaso Bithi, Kolkata - 700017",
                    "Phone: 9817473373",
                    "Fax: N.A."
                ]
                for line in supplier: elements.append(Paragraph(line, normal))
                elements.append(Spacer(1,12))

                # --- Consignee ---
                consignee = [
                    "<b>Consignee:</b>",
                    "RNA Resource Group Ltd - Landmark (Babyshop)",
                    "P.O Box 25030, Dubai, UAE",
                    "Tel: 00971 4 8095500, Fax: 00971 4 8095555/66"
                ]
                for line in consignee: elements.append(Paragraph(line, normal))
                elements.append(Spacer(1,12))

                # --- Order details ---
                today = datetime.today().strftime("%d-%m-%Y")
                order_info=[
                    f"<b>Landmark Order Ref:</b> {order_no}",
                    f"<b>Buyer Name:</b> {buyer_name}",
                    f"<b>Brand:</b> {brand}",
                    f"<b>Loading Country:</b> {made_in}",
                    f"<b>Port of Loading:</b> {loading_port}",
                    f"<b>Agreed Shipment Date:</b> {ship_date}",
                    f"<b>Description of Goods:</b> {order_of}",
                    f"<b>Report Date:</b> {today}"
                ]
                for line in order_info: elements.append(Paragraph(line, normal))
                elements.append(Spacer(1,12))

                # --- Bank details ---
                bank=[
                    "<b>Bank Details (Including Swift/IBAN):</b>",
                    "Beneficiary: SAR APPARELS INDIA PVT.LTD",
                    "Account No: 2112819952",
                    "Bank: Kotak Mahindra Bank Ltd",
                    "Address: 2 Brabourne Road, Govind Bhavan, Ground Floor, Kolkata - 700001",
                    "SWIFT: KKBKINBBCPC",
                    "Bank Code: 0323"
                ]
                for line in bank: elements.append(Paragraph(line, normal))
                elements.append(Spacer(1,12))

                # --- Table ---
                data=[list(agg_df.columns)]
                for _,row in agg_df.iterrows(): data.append(list(row))
                total_qty=agg_df["QTY"].sum()
                total_amount=agg_df["AMOUNT"].astype(float).sum()
                data.append(["TOTAL","","","","","",f"{int(total_qty):,}","",f"{total_amount:,.2f}"])

                col_widths=[0.12*A4[0],0.25*A4[0],0.10*A4[0],0.10*A4[0],0.20*A4[0],0.08*A4[0],0.05*A4[0],0.05*A4[0],0.05*A4[0]]
                table=Table(data,colWidths=col_widths,repeatRows=1)
                style=TableStyle([
                    ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#333333")),
                    ("TEXTCOLOR",(0,0),(-1,0),colors.whitesmoke),
                    ("ALIGN",(0,0),(-1,0),"CENTER"),
                    ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                    ("FONTSIZE",(0,0),(-1,0),7),
                    ("GRID",(0,0),(-1,-1),0.25,colors.black),
                    ("ALIGN",(-3,1),(-1,-1),"RIGHT"),
                    ("FONTSIZE",(0,1),(-1,-1),8),
                    ("VALIGN",(0,0),(-1,-1),"MIDDLE")
                ])
                style.add("FONTNAME",(0,len(data)-1),(-1,len(data)-1),"Helvetica-Bold")
                style.add("BACKGROUND",(0,len(data)-1),(-1,len(data)-1),colors.lightgrey)
                table.setStyle(style)
                elements.append(table)
                elements.append(Spacer(1,12))

                # --- Totals ---
                amount_words=amount_to_words(total_amount)
                elements.append(Paragraph(f"<b>Total Amount:</b> USD {total_amount:,.2f}", bold))
                elements.append(Paragraph(f"<b>In Words:</b> {amount_words}", normal))
                elements.append(Spacer(1,24))

                # --- Signature ---
                elements.append(Paragraph("For RNA Resources Group Ltd - Landmark (Babyshop)", normal))
                elements.append(Spacer(1,36))
                elements.append(Paragraph("-------------------------------------", normal))
                elements.append(Paragraph("Authorised Signatory", normal))

                doc.build(elements)
                with open(pdf_file,"rb") as f:
                    st.download_button("⬇️ Download PI PDF", f, file_name="Proforma_Invoice.pdf")
                os.remove(pdf_file)
