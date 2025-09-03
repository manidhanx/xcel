import streamlit as st
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
import tempfile
import os
from datetime import datetime

st.set_page_config(page_title="Excel Style Aggregator", layout="centered")

st.title("👕 Excel → PDF (Style Report)")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    # Load raw Excel (no header)
    raw_df = pd.read_excel(uploaded_file, header=None)

    # --- Extract key info ---
    buyer_name = str(raw_df.iloc[0, 0]) if not pd.isna(raw_df.iloc[0, 0]) else None
    order_no, brand, made_in, loading_port, ship_date, order_of, texture = [None]*7

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
                try: made_in = row[j+1]
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
            elif cell_val == "texture:":
                try: texture = row[j+1]
                except: pass

    # --- Find header row ("Style") ---
    header_row_idx = None
    for i, row in raw_df.iterrows():
        if row.astype(str).str.strip().str.lower().eq("style").any():
            header_row_idx = i
            break

    if header_row_idx is None:
        st.error("❌ Could not find a 'Style' header in the file.")
    else:
        # Use that row as header
        df = pd.read_excel(uploaded_file, header=header_row_idx)
        df = df.dropna(how="all")

        st.write("### Preview of extracted data")
        st.dataframe(df.head())

        if "Style" not in df.columns:
            st.error("❌ No 'Style' column found after parsing the file.")
        else:
            # Identify Qty and FOB columns
            qty_col = None
            fob_col = None
            for col in df.columns:
                if str(col).strip().lower() in ["qty", "quantity"]:
                    qty_col = col
                if "fob" in str(col).strip().lower():
                    fob_col = col

            # Build final table rows
            final_rows = []
            for _, row in df.iterrows():
                style = row.get("Style")
                if pd.isna(style):
                    continue

                item_desc = row.iloc[1] if len(row) > 1 else ""  # Assume description = 2nd col
                composition = row.iloc[2] if len(row) > 2 else ""  # Assume composition = 3rd col

                qty = row.get(qty_col, 0) if qty_col else 0
                price = row.get(fob_col, 0) if fob_col else 0
                try:
                    amount = float(qty) * float(price)
                except:
                    amount = 0

                final_rows.append([
                    style,                                # STYLE NO
                    item_desc,                            # ITEM DESCRIPTION
                    texture if texture else "",           # FABRIC TYPE
                    "61112000",                           # H.S NO
                    composition,                          # COMPOSITION
                    made_in if made_in else "",           # COUNTRY OF ORIGIN
                    qty,                                  # QTY
                    price,                                # FOB PRICE
                    amount                                # AMOUNT
                ])

            # Generate PDF
            if st.button("Generate PDF"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmpfile:
                    pdf_file = tmpfile.name

                doc = SimpleDocTemplate(pdf_file, pagesize=landscape(A4))
                styles = getSampleStyleSheet()
                elements = []

                # --- Header Info ---
                today = datetime.today().strftime("%d-%m-%Y")
                header_info = [
                    f"<b>Buyer:</b> {buyer_name}" if buyer_name else None,
                    f"<b>Order No:</b> {order_no}" if order_no else None,
                    f"<b>Brand:</b> {brand}" if brand else None,
                    f"<b>Made in Country:</b> {made_in}" if made_in else None,
                    f"<b>Loading Port:</b> {loading_port}" if loading_port else None,
                    f"<b>Agreed Ship Date:</b> {ship_date}" if ship_date else None,
                    f"<b>Order Of:</b> {order_of}" if order_of else None,
                    f"<b>Report Date:</b> {today}"
                ]

                for line in header_info:
                    if line:
                        elements.append(Paragraph(line, styles["Normal"]))
                elements.append(Spacer(1, 12))

                # --- Table Data ---
                table_data = [[
                    "STYLE NO.", "ITEM DESCRIPTION", "FABRIC TYPE", "H.S NO (8digit)",
                    "COMPOSITION OF MATERIAL", "COUNTRY OF ORIGIN", "QTY",
                    "UNIT PRICE FOB", "AMOUNT"
                ]] + final_rows

                table = Table(table_data, repeatRows=1)
                table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("GRID", (0, 0), (-1, -1), 1, colors.black),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                ]))

                elements.append(table)
                doc.build(elements)

                with open(pdf_file, "rb") as f:
                    st.download_button("⬇️ Download PDF", f, file_name="style_report.pdf")

                os.remove(pdf_file)
