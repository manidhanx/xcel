import streamlit as st
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import tempfile, os
from datetime import datetime

# --- Number to words ---
def number_to_words(n):
    ones = ["","ONE","TWO","THREE","FOUR","FIVE","SIX","SEVEN","EIGHT","NINE",
            "TEN","ELEVEN","TWELVE","THIRTEEN","FOURTEEN","FIFTEEN","SIXTEEN",
            "SEVENTEEN","EIGHTEEN","NINETEEN"]
    tens = ["","","TWENTY","THIRTY","FORTY","FIFTY","SIXTY","SEVENTY","EIGHTY","NINETY"]

    def words(num):
        if num < 20: return ones[num]
        elif num < 100: return tens[num//10] + ("" if num%10==0 else "-" + ones[num%10])
        elif num < 1000: return ones[num//100] + " HUNDRED" + ("" if num%100==0 else " " + words(num%100))
        elif num < 1_000_000: return words(num//1000] + " THOUSAND" + ("" if num%1000==0 else " " + words(num%1000))
        else: return str(num)
    return words(n)

def amount_to_words(amount):
    return number_to_words(int(amount)) + " DOLLARS"

# --- Streamlit setup ---
st.set_page_config(page_title="Proforma Invoice Generator", layout="centered")
st.title("Proforma Invoice Generator")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

agg_df, order_no, made_in, loading_port, ship_date, order_of, texture, country_of_origin = [None]*8

if uploaded_file:
    raw_df = pd.read_excel(uploaded_file, header=None)

    # Extract shipment info
    for i, row in raw_df.iterrows():
        for j, cell in enumerate(row):
            val = str(cell).strip().lower()
            if val == "order no :": order_no = row[j+2]
            elif val == "made in country :": made_in = row[j+1]; country_of_origin = row[j+1]
            elif val == "loading port :": loading_port = row[j+1]
            elif val == "agreed ship date :": ship_date = row[j+2]
            elif val == "order of": order_of = row[j+1]
            elif val == "texture :": texture = row[j+1]

    if isinstance(ship_date, (datetime, pd.Timestamp)):
        ship_date = ship_date.strftime("%d-%m-%Y")

    # Find header row
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
        qty_col, fob_col = None, None
        for col in df.columns:
            if "qty" in str(col).lower(): qty_col = col
            if "fob" in str(col).lower(): fob_col = col

        if style_col and qty_col:
            aggregated=[]
            for style in df[style_col].dropna().unique():
                rows = df[df[style_col]==style]
                if len(rows)>0:
                    r=rows.iloc[0]
                    desc=r.iloc[1] if len(r)>1 else ""
                    comp=r.iloc[2] if len(r)>2 else ""
                    qty=pd.to_numeric(rows[qty_col],errors="coerce").fillna(0).sum()
                    price=0
                    if fob_col and fob_col in rows.columns:
                        prices=pd.to_numeric(rows[fob_col],errors="coerce").fillna(0)
                        nz=prices[prices>0]; price=nz.iloc[0] if len(nz)>0 else 0
                    amount=qty*price
                    aggregated.append([style,desc,texture or "Knitted","61112000",comp,
                                       country_of_origin or "India",int(qty),f"{price:.2f}",f"{amount:.2f}"])
            agg_df=pd.DataFrame(aggregated,columns=[
                "STYLE NO.","ITEM DESCRIPTION","FABRIC TYPE","H.S NO",
                "COMPOSITION","COUNTRY OF ORIGIN","QTY","UNIT PRICE FOB","AMOUNT"
            ])
            st.write("### ✅ Parsed Order Data")
            st.dataframe(agg_df)

# --- PDF generation ---
if agg_df is not None:
    today = datetime.today().strftime("%d-%m-%Y")
    pi_no = st.text_input("PI No. & Date", f"SAR/LG/XXXX Dt. {today}")
    buyer_name = st.text_input("Buyer Name", "LANDMARK GROUP")
    brand_name = st.text_input("Brand Name", "Juniors")
    payment_term = st.text_input("Payment Term", "T/T")

    consignee_name = "RNA Resources Group Ltd- Landmark (Babyshop)"
    consignee_addr = "P O Box 25030, Dubai, UAE"
    consignee_tel = "Tel: 00971 4 8095500, Fax: 00971 4 8095555/66"

    if st.button("Generate Proforma Invoice"):
        with tempfile.NamedTemporaryFile(delete=False,suffix=".pdf") as tmp:
            pdf_file = tmp.name

        doc = SimpleDocTemplate(pdf_file, pagesize=A4, leftMargin=20, rightMargin=20, topMargin=20, bottomMargin=20)
        styles=getSampleStyleSheet(); normal=styles["Normal"]
        bold=ParagraphStyle("bold",parent=normal,fontName="Helvetica-Bold",fontSize=10)

        sections=[]

        # Title
        title = Table([[Paragraph("Proforma Invoice", bold)]], colWidths=[500])
        title.setStyle(TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER"),("GRID",(0,0),(-1,-1),0.5,colors.black)]))
        sections.append(title)

        # Supplier
        sup = [
            ["Supplier Name No. & date of PI", pi_no],
            ["SAR APPARELS INDIA PVT.LTD.", ""],
            ["ADDRESS : 6, Picaso Bithi, KOLKATA - 700017.", f"Landmark order Reference: {order_no}"],
            ["PHONE : 9874173373", f"Buyer Name: {buyer_name}"],
            ["FAX : N.A.", f"Brand Name: {brand_name}"]
        ]
        sup_table = Table(sup, colWidths=[250,250])
        sup_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.5,colors.black)]))
        sections.append(sup_table)

        # Consignee + Bank
        con = [
            ["Consignee:-", f"Payment Term: {payment_term}"],
            [consignee_name, "Bank Details (Including Swift/IBAN)"],
            [consignee_addr, "SAR APPARELS INDIA PVT.LTD"],
            [consignee_tel, "ACCOUNT NO :- 2112819952"],
            ["", "BANK'S NAME :- KOTAK MAHINDRA BANK LTD"],
            ["", "BANK ADDRESS :- 2 BRABOURNE ROAD, GOVIND BHAVAN, GROUND FLOOR, KOLKATA-700001"],
            ["", "SWIFT :- KKBKINBBCPC"],
            ["", "BANK CODE :- 0323"]
        ]
        con_table = Table(con, colWidths=[250,250])
        con_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.5,colors.black)]))
        sections.append(con_table)

        # Shipment
        ship = [
            [f"Loading Country: {made_in}", "L/C Advicing Bank (If Payment term LC Applicable )"],
            [f"Port of loading: {loading_port}", ""],
            [f"Agreed Shipment Date: {ship_date}", ""],
            ["REMARKS if ANY:-", ""],
            [f"Description of goods: {order_of}", ""],
            ["CURRENCY: USD", ""]
        ]
        ship_table = Table(ship, colWidths=[250,250])
        ship_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.5,colors.black)]))
        sections.append(ship_table)

        # Items
        data = [[
            "STYLE NO.","ITEM DESCRIPTION","FABRIC TYPE\n(KNITTED / WOVEN)","H.S NO (8digit)",
            "COMPOSITION OF MATERIAL","COUNTRY OF ORIGIN","QTY","UNIT PRICE FOB","AMOUNT"
        ]]
        for _, row in agg_df.iterrows(): data.append(list(row))
        tot_qty = agg_df["QTY"].sum()
        tot_amt = agg_df["AMOUNT"].astype(float).sum()
        data.append(["TOTAL","","","","","",f"{int(tot_qty):,}","",f"{tot_amt:,.2f} USD"])

        col_widths = [55,95,60,60,75,60,45,60,60]  # sum = 500
        items_table = Table(data, colWidths=col_widths, repeatRows=1)
        items_table.setStyle(TableStyle([
            ("GRID",(0,0),(-1,-1),0.5,colors.black),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
            ("FONTSIZE",(0,0),(-1,-1),8),
            ("FONTNAME",(0,-1),(-1,-1),"Helvetica-Bold"),  # bold total row
            ("ALIGN",(6,-1),(-1,-1),"RIGHT")  # align total qty & amount right
        ]))
        sections.append(items_table)

        # Amount in words
        amt_words = amount_to_words(tot_amt)
        amt_table = Table([[
            f"TOTAL US DOLLAR {amt_words}", f"{tot_amt:,.2f} USD"
        ]], colWidths=[350,150])
        amt_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.5,colors.black)]))
        sections.append(amt_table)

        # Terms
        terms = Table([["Terms & Conditions (If Any)"]], colWidths=[500])
        terms.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.5,colors.black)]))
        sections.append(terms)

        # Signature
        sig_img="sarsign.png"
        sign = Table([
            [Image(sig_img,width=120,height=40)],
            ["Signed by …………………….(Affix Stamp here) for RNA Resources Group Ltd-Landmark (Babyshop)"]
        ], colWidths=[500])
        sign.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.5,colors.black)]))
        sections.append(sign)

        # Outer frame
        outer = Table([[s] for s in sections], colWidths=[500])
        outer.setStyle(TableStyle([("GRID",(0,0),(-1,-1),1,colors.black)]))

        doc.build([outer])

        with open(pdf_file,"rb") as f:
            st.download_button("⬇️ Download Proforma Invoice", f, file_name="Proforma_Invoice.pdf")
        os.remove(pdf_file)
