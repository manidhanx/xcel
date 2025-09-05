import streamlit as st
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import tempfile, os
from datetime import datetime

# --- Helper functions ---
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
    words = number_to_words(whole) + " DOLLARS"
    return words

# --- Streamlit setup ---
st.set_page_config(page_title="Proforma Invoice Generator", layout="centered")
st.title("Proforma Invoice Generator")

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
        ship_date = ship_date.strftime("%d-%m-%Y")

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

# --- PDF generation ---
if agg_df is not None:
    st.write("### ✍️ Enter Invoice Details")
    today_str = datetime.today().strftime("%d-%m-%Y")
    pi_no = st.text_input("PI No. & Date", f"SAR/LG/XXXX Dt. {today_str}")
    buyer_name = st.text_input("Buyer Name", "LANDMARK GROUP")
    brand_name = st.text_input("Brand Name", "Juniors")
    payment_term = st.text_input("Payment Term", "T/T")

    consignee_name = "RNA Resources Group Ltd- Landmark (Babyshop)"
    consignee_addr = "P O Box 25030, Dubai, UAE"
    consignee_tel = "Tel: 00971 4 8095500, Fax: 00971 4 8095555/66"

    if st.button("Generate Proforma Invoice"):
        with tempfile.NamedTemporaryFile(delete=False,suffix=".pdf") as tmp:
            pdf_file=tmp.name

        doc=SimpleDocTemplate(pdf_file,pagesize=A4,leftMargin=30,rightMargin=30,topMargin=30,bottomMargin=30)
        styles=getSampleStyleSheet()
        normal=styles["Normal"]
        bold=ParagraphStyle("bold",parent=normal,fontName="Helvetica-Bold",fontSize=9)

        elements=[]

        # --- Title ---
        elements.append(Paragraph("<b>Proforma Invoice</b>", bold))
        elements.append(Spacer(1, 12))

        # --- Supplier / Buyer / Consignee Info ---
        elements.append(Paragraph(f"Supplier Name No. & date of PI {pi_no}", normal))
        elements.append(Paragraph("SAR APPARELS INDIA PVT.LTD.", bold))
        elements.append(Paragraph("ADDRESS : 6, Picaso Bithi, KOLKATA - 700017.", normal))
        elements.append(Paragraph(f"Landmark order Reference: {order_no}", normal))
        elements.append(Paragraph("PHONE : 9874173373", normal))
        elements.append(Paragraph(f"Buyer Name: {buyer_name}", normal))
        elements.append(Paragraph("FAX : N.A.", normal))
        elements.append(Paragraph(f"Brand Name: {brand_name}", normal))
        elements.append(Spacer(1, 12))

        elements.append(Paragraph(f"Consignee:- {consignee_name}", normal))
        elements.append(Paragraph(f"Payment Term: {payment_term}", normal))
        elements.append(Paragraph(consignee_addr, normal))
        elements.append(Paragraph(consignee_tel, normal))
        elements.append(Paragraph("Bank Details (Including Swift/IBAN): SAR APPARELS INDIA PVT.LTD", normal))
        elements.append(Paragraph("ACCOUNT NO :- 2112819952", normal))
        elements.append(Paragraph("BANK'S NAME :- KOTAK MAHINDRA BANK LTD", normal))
        elements.append(Paragraph("BANK ADDRESS :- 2 BRABOURNE ROAD, GOVIND BHAVAN, GROUND FLOOR, KOLKATA-700001", normal))
        elements.append(Paragraph("SWIFT :- KKBKINBBCPC", normal))
        elements.append(Paragraph("BANK CODE :- 0323", normal))
        elements.append(Spacer(1, 12))

        elements.append(Paragraph(f"Loading Country: {made_in}", normal))
        elements.append(Paragraph(f"Port of loading: {loading_port}", normal))
        elements.append(Paragraph(f"Agreed Shipment Date: {ship_date}", normal))
        elements.append(Paragraph(f"REMARKS if ANY:-", normal))
        elements.append(Paragraph(f"Description of goods: {order_of}", normal))
        elements.append(Paragraph("CURRENCY: USD", normal))
        elements.append(Spacer(1, 12))

        # --- Items Table ---
        data=[list(agg_df.columns)]
        for _,row in agg_df.iterrows(): data.append(list(row))
        total_qty=agg_df["QTY"].sum()
        total_amount=agg_df["AMOUNT"].astype(float).sum()
        data.append(["TOTAL","","","","","",f"{int(total_qty):,}","",f"{total_amount:,.2f}USD"])

        table=Table(data, repeatRows=1)
        table.setStyle(TableStyle([
            ("GRID",(0,0),(-1,-1),0.5,colors.black),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
            ("FONTSIZE",(0,0),(-1,-1),8)
        ]))
        elements.append(table)
        elements.append(Spacer(1, 12))

        # --- Amount in Words ---
        amount_words=amount_to_words(total_amount)
        elements.append(Paragraph(f"TOTAL US DOLLAR {amount_words}", normal))
        elements.append(Spacer(1, 24))

        # --- Signature ---
        sig_img = "sarsign.png"
        sign_table=Table([
            [Image(sig_img,width=120,height=40),
             Paragraph("Signed by …………………….(Affix Stamp here) for RNA Resources Group Ltd-Landmark (Babyshop)", normal)]
        ],colWidths=[200, 300])
        sign_table.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"MIDDLE")]))
        elements.append(sign_table)

        # --- Terms ---
        elements.append(Spacer(1, 12))
        elements.append(Paragraph("Terms & Conditions (If Any)", normal))

        doc.build(elements)

        with open(pdf_file, "rb") as f:
            st.download_button("⬇️ Download Proforma Invoice", f, file_name="Proforma_Invoice.pdf")
        os.remove(pdf_file)
