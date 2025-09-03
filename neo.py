import streamlit as st
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
import tempfile
import os
from datetime import datetime

st.set_page_config(page_title="Excel Style Aggregator", layout="centered")

st.title("👕 Excel → PDF (Style Aggregator - Fixed Qty)")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    # Load raw Excel (no header)
    raw_df = pd.read_excel(uploaded_file, header=None)

    # --- Extract key info ---
    buyer_name = str(raw_df.iloc[0, 0]) if not pd.isna(raw_df.iloc[0, 0]) else None
    order_no, brand, made_in, loading_port, ship_date, order_of = [None]*6
    texture = None
    country_of_origin = None

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

    # --- Find the row index where "Style" appears ---
    header_row_idx = None
    for i, row in raw_df.iterrows():
        if row.astype(str).str.strip().str.lower().eq("style").any():
            header_row_idx = i
            break

    if header_row_idx is None:
        st.error("❌ Could not find a 'Style' header in the file.")
    else:
        # Read with two header rows (to merge split headers like "Total" + "Qty")
        df = pd.read_excel(uploaded_file, header=[header_row_idx, header_row_idx+1])
        df.columns = [
            " ".join([str(x) for x in col if str(x) != "nan"]).strip()
            for col in df.columns.values
        ]
        df = df.dropna(how="all")  # drop completely empty rows

        st.write("### Preview of extracted data")
        st.dataframe(df.head())

        if "Style" not in df.columns:
            st.error("❌ No 'Style' column found after parsing the file.")
        else:
            # Find the "Total Qty" column
            qty_col = None
            for col in df.columns:
                if str(col).strip().lower() == "total qty":
                    qty_col = col
                    break

            # Find the FOB price column
            fob_col = None
            for col in df.columns:
                if "fob" in str(col).lower():
                    fob_col = col
                    break

            # Build aggregated data
            aggregated_data = []
            unique_styles = df['Style'].dropna().unique()
            
            for style in unique_styles:
                style_rows = df[df['Style'] == style]
                if len(style_rows) > 0:
                    first_row = style_rows.iloc[0]

                    # Extract description & composition
                    item_description = first_row.iloc[1] if len(first_row) > 1 else ""
                    composition = first_row.iloc[2] if len(first_row) > 2 else ""

                    # Calculate total qty
                    total_qty = 0
                    if qty_col and qty_col in style_rows.columns:
                        qty_values = pd.to_numeric(style_rows[qty_col], errors='coerce').fillna(0)
                        total_qty = qty_values.sum()

                    # FOB price
                    unit_price = 0
                    if fob_col and fob_col in style_rows.columns:
                        price_values = pd.to_numeric(style_rows[fob_col], errors='coerce').fillna(0)
                        non_zero_prices = price_values[price_values > 0]
                        unit_price = non_zero_prices.iloc[0] if len(non_zero_prices) > 0 else 0

                    # Amount
                    amount = total_qty * unit_price

                    aggregated_data.append({
                        'STYLE NO.': style,
                        'ITEM DESCRIPTION': item_description if str(item_description) != "nan" else "",
                        'FABRIC TYPE (KNITTED/WOVEN)': texture if texture else "",
                        'H.S NO (8digit)': "61112000",
                        'COMPOSITION OF MATERIAL': composition if str(composition) != "nan" else "",
                        'COUNTRY OF ORIGIN': country_of_origin if country_of_origin else "",
                        'QTY': total_qty,
                        'UNIT PRICE FOB': unit_price,
                        'AMOUNT': amount
                    })

            agg_df = pd.DataFrame(aggregated_data)

            st.write("### Aggregated Data (Custom Format)")
            st.dataframe(agg_df)

            # Generate PDF
            if st.button("Generate PDF"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmpfile:
                    pdf_file = tmpfile.name

                doc = SimpleDocTemplate(pdf_file, pagesize=A4)
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
                if len(agg_df) > 0:
                    data = [agg_df.columns.tolist()]
                    for _, row in agg_df.iterrows():
                        formatted_row = []
                        for val in row:
                            if isinstance(val, (int, float)) and not pd.isna(val):
                                formatted_row.append(f"{val:.2f}" if val % 1 != 0 else f"{int(val)}")
                            else:
                                formatted_row.append(str(val) if not pd.isna(val) else "")
                        data.append(formatted_row)

                    table = Table(data)
                    table.setStyle(TableStyle([
                        ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                        ("GRID", (0, 0), (-1, -1), 1, colors.black),
                        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                        ("FONTSIZE", (0, 0), (-1, -1), 8),
                    ]))

                    elements.append(table)
                else:
                    elements.append(Paragraph("No data to display", styles["Normal"]))

                doc.build(elements)

                with open(pdf_file, "rb") as f:
                    st.download_button("⬇️ Download PDF", f, file_name="style_report.pdf")

                os.remove(pdf_file)
