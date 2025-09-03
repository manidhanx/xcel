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
st.title("📑 Proforma Invoice Generator (v11.2 Pure)")

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

# --- Only show inputs if parsing succeeded ---
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
        bold=ParagraphStyle("bold",parent=normal,fontName="Helvetica-Bold")

        elements=[]
        content_width = A4[0] - 100

        # --- Title ---
        title_table = Table([["PROFORMA INVOICE"]], colWidths=[content_width])
        title_table.setStyle(TableStyle([
            ("GRID",(0,0),(-1,-1),0.75,colors.black),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),
            ("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),
            ("FONTSIZE",(0,0),(-1,-1),14),
            ("TOPPADDING",(0,0),(-1,-1),8),
            ("BOTTOMPADDING",(0,0),(-1,-1),8),
        ]))
        elements.append(title_table)
        elements.append(Spacer(1,12))

        # --- Supplier & Consignee Block ---
        sup=[
            ["<b>Supplier Name:</b> SAR APPARELS INDIA PVT.LTD.",pi_no],
            ["Address: 6, Picaso Bithi, Kolkata - 700017","<b>Landmark order Reference:</b> "+str(order_no)],
            ["Phone: 9817473373","<b>Buyer Name:</b> "+buyer_name],
            ["Fax: N.A.","<b>Brand Name:</b> "+brand_name],
        ]
        con=[
            ["<b>Consignee:</b>",payment_term],
            [consignee_name,"<b>Bank Details (Including Swift/IBAN):</b>"],
            [consignee_addr,"Beneficiary: SAR APPARELS INDIA PVT.LTD"],
            [consignee_tel,"Account No: 2112819952"],
            ["","Bank: Kotak Mahindra Bank Ltd"],
            ["","Address: 2 Brabourne Road, Govind Bhavan, Ground Floor, Kolkata - 700001"],
            ["","SWIFT: KKBKINBBCPC"],
            ["","Bank Code: 0323"],
        ]
        info_table=Table(sup+con,colWidths=[0.5*content_width,0.5*content_width])
        info_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.black),
                                        ("VALIGN",(0,0),(-1,-1),"TOP"),
                                        ("FONTSIZE",(0,0),(-1,-1),8)]))
        elements.append(info_table)
        elements.append(Spacer(1,12))

        # --- Shipment Info ---
        ship=[
            ["<b>Loading Country:</b> "+str(made_in),"<b>Port of Loading:</b> "+str(loading_port)],
            ["<b>Agreed Shipment Date:</b> "+str(ship_date),"<b>Description of goods:</b> "+str(order_of)]
        ]
        ship_table=Table(ship,colWidths=[0.5*content_width,0.5*content_width])
        ship_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.black),
                                        ("FONTSIZE",(0,0),(-1,-1),8)]))
        elements.append(ship_table)
        elements.append(Spacer(1,12))

        # --- Main Items Table ---
        data=[list(agg_df.columns)]
        for _,row in agg_df.iterrows(): data.append(list(row))
        total_qty=agg_df["QTY"].sum()
        total_amount=agg_df["AMOUNT"].astype(float).sum()
        data.append(["TOTAL","","","","","",f"{int(total_qty):,}","USD",f"{total_amount:,.2f}"])

        col_widths = [
            content_width * 0.10,  # STYLE NO
            content_width * 0.28,  # ITEM DESCRIPTION
            content_width * 0.09,  # FABRIC TYPE
            content_width * 0.10,  # H.S NO
            content_width * 0.15,  # COMPOSITION
            content_width * 0.08,  # ORIGIN
            content_width * 0.07,  # QTY
            content_width * 0.065, # FOB
            content_width * 0.065  # AMOUNT
        ]

        table=Table(data,colWidths=col_widths,repeatRows=1)
        style=TableStyle([
            ("GRID",(0,0),(-1,-1),0.25,colors.black),
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#333333")),
            ("TEXTCOLOR",(0,0),(-1,0),colors.whitesmoke),
            ("ALIGN",(0,0),(-1,0),"CENTER"),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("FONTSIZE",(0,0),(-1,0),6.5),
            ("ALIGN",(0,1),(5,-1),"LEFT"),
            ("ALIGN",(6,1),(-1,-1),"RIGHT"),
            ("FONTSIZE",(0,1),(-1,-1),8),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("LEFTPADDING",(0,0),(-1,-1),4),
            ("RIGHTPADDING",(0,0),(-1,-1),4),
            ("WORDWRAP",(0,0),(-1,0),"CJK")
        ])
        style.add("FONTNAME",(0,len(data)-1),(-1,len(data)-1),"Helvetica-Bold")
        style.add("BACKGROUND",(0,len(data)-1),(-1,len(data)-1),colors.lightgrey)
        table.setStyle(style)
        elements.append(table)

        # --- Amount in Words Row ---
        amount_words=amount_to_words(total_amount)
        words_table=Table([[f"TOTAL  US DOLLAR {amount_words}"]],colWidths=[content_width])
        words_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.black),
                                         ("FONTSIZE",(0,0),(-1,-1),8)]))
        elements.append(words_table)

        # --- Terms & Conditions ---
        terms_table=Table([["Terms & Conditions (if any):"]],colWidths=[content_width])
        terms_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.black),
                                         ("FONTSIZE",(0,0),(-1,-1),8)]))
        elements.append(terms_table)
        elements.append(Spacer(1,24))

        # --- Signature ---
        sig_img = "sarsign.png"  # constant image stored alongside script
        sign_table=Table([
            [Image(sig_img,width=150,height=50),
             "Signed by ………………… for RNA Resources Group Ltd - Landmark (Babyshop)"]
        ],colWidths=[0.5*content_width,0.5*content_width])
        sign_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.black),
                                        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                                        ("ALIGN",(0,0),(0,0),"LEFT"),
                                        ("ALIGN",(1,0),(1,0),"RIGHT"),
                                        ("FONTSIZE",(0,0),(-1,-1),8)]))
        elements.append(sign_table)

        # --- Outer Frame ---
        outer_table = Table([[e] for e in elements], colWidths=[content_width])
        outer_table.setStyle(TableStyle([
            ("GRID",(0,0),(-1,-1),1.5,colors.black),  # Official stampy frame
            ("VALIGN",(0,0),(-1,-1),"TOP")
        ]))

        doc.build([outer_table])
        with open(pdf_file,"rb") as f:
            st.download_button("⬇️ Download PI PDF", f, file_name="Proforma_Invoice.pdf")
        os.remove(pdf_file)
