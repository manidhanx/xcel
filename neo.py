import streamlit as st
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from num2words import num2words
import tempfile
import os
from datetime import datetime

st.set_page_config(page_title="Proforma Invoice Generator", layout="centered")

st.title("📑 Proforma Invoice Generator (v9 Final)")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    # Load raw Excel (no header)
    raw_df = pd.read_excel(uploaded_file, header=None)

    # --- Extract key info ---
    buyer_name = str(raw_df.iloc[0, 0]) if not pd.isna(raw_df.iloc[0, 0]) else "LANDMARK GROUP"
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

        # --- Detect columns ---
        style_col = next((col for col in df.columns if str(col).strip().lower().startswith("style")), None)

        qty_col = None
        value_col_index = None
        for idx, col in enumerate(df.columns):
            if "value" in str(col).lower():
                value_col_index = idx
                break
        if value_col_index and value_col_index > 0:
            qty_col = df.columns[value_col_index - 1]

        fob_col = next((col for col in df.columns if "fob" in str(col).lower()), None)

        if not style_col or not qty_col:
            st.error("❌ Could not detect required columns. Please check the Excel format.")
        else:
            # --- Aggregate Data ---
            aggregated_data = []
            unique_styles = df[style_col].dropna().unique()

            for style in unique_styles:
                style_rows = df[df[style_col] == style]
                if len(style_rows) > 0:
                    first_row = style_rows.iloc[0]

                    item_description = first_row.iloc[1] if len(first_row) > 1 else ""
                    composition = first_row.iloc[2] if len(first_row) > 2 else ""

                    total_qty = pd.to_numeric(style_rows[qty_col], errors='coerce').fillna(0).sum()

                    unit_price = 0
                    if fob_col and fob_col in style_rows.columns:
                        price_values = pd.to_numeric(style_rows[fob_col], errors='coerce').fillna(0)
                        non_zero_prices = price_values[price_values > 0]
                        unit_price = non_zero_prices.iloc[0] if len(non_zero_prices) > 0 else 0

                    amount = total_qty * unit_price

                    aggregated_data.append([
                        style,
                        item_description,
                        texture if texture else "",
                        "61112000",
                        composition,
                        country_of_origin if country_of_origin else "",
                        int(total_qty),
                        f"{unit_price:.2f}",
                        f"{amount:.2f}"
                    ])

            agg_df = pd.DataFrame(aggregated_data, columns=[
                "STYLE NO.",
                "ITEM DESCRIPTION",
                "FABRIC TYPE (KNITTED/WOVEN)",
                "H.S NO (8digit)",
                "COMPOSITION OF MATERIAL",
                "COUNTRY OF ORIGIN",
                "QTY",
                "UNIT PRICE FOB",
                "AMOUNT"
            ])

            st.write("### Aggregated Data (Preview)")
            st.dataframe(agg_df)

            # --- Generate PDF ---
            if st.button("Generate PI PDF"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmpfile:
                    pdf_file = tmpfile.name

                doc = SimpleDocTemplate(
                    pdf_file,
                    pagesize=A4,
                    leftMargin=30,
                    rightMargin=30,
                    topMargin=30,
                    bottomMargin=30
                )
                styles = getSampleStyleSheet()
                normal = styles["Normal"]
                bold = ParagraphStyle("bold", parent=styles["Normal"], fontName="Helvetica-Bold")

                elements = []

                # --- Header ---
                today = datetime.today().strftime("%d-%m-%Y")
                header_info = [
                    f"<b>Buyer:</b> {buyer_name}",
                    f"<b>Order No:</b> {order_no}",
                    f"<b>Brand:</b> {brand}",
                    f"<b>Made in Country:</b> {made_in}",
                    f"<b>Loading Port:</b> {loading_port}",
                    f"<b>Agreed Ship Date:</b> {ship_date}",
                    f"<b>Order Of:</b> {order_of}",
                    f"<b>Report Date:</b> {today}"
                ]
                for line in header_info:
                    elements.append(Paragraph(line, normal))
                elements.append(Spacer(1, 12))

                # --- Table ---
                data = [list(agg_df.columns)]
                for _, row in agg_df.iterrows():
                    data.append(list(row))

                # Add total row
                total_qty = agg_df["QTY"].sum()
                total_amount = agg_df["AMOUNT"].astype(float).sum()
                data.append([
                    "TOTAL", "", "", "", "", "",
                    f"{int(total_qty):,}", "", f"{total_amount:,.2f}"
                ])

                col_widths = [
                    0.12 * A4[0], 0.25 * A4[0], 0.10 * A4[0], 0.10 * A4[0],
                    0.20 * A4[0], 0.08 * A4[0], 0.05 * A4[0], 0.05 * A4[0], 0.05 * A4[0]
                ]
                table = Table(data, colWidths=col_widths, repeatRows=1)

                style = TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#333333")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                    ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, 0), 7),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.black),
                    ("ALIGN", (-3, 1), (-1, -1), "RIGHT"),
                    ("FONTSIZE", (0, 1), (-1, -1), 8),
                ])
                style.add("FONTNAME", (0, len(data)-1), (-1, len(data)-1), "Helvetica-Bold")
                style.add("BACKGROUND", (0, len(data)-1), (-1, len(data)-1), colors.lightgrey)

                table.setStyle(style)
                elements.append(table)
                elements.append(Spacer(1, 12))

                # --- Totals in words ---
                amount_words = num2words(total_amount, to="currency", lang="en").upper()
                elements.append(Paragraph(f"<b>Total Amount:</b> USD {total_amount:,.2f}", bold))
                elements.append(Paragraph(f"<b>In Words:</b> {amount_words}", normal))
                elements.append(Spacer(1, 24))

                # --- Signature block ---
                elements.append(Paragraph("For RNA Resources Group Ltd - Landmark (Babyshop)", normal))
                elements.append(Spacer(1, 48))
                elements.append(Paragraph("Authorised Signatory", normal))

                doc.build(elements)

                with open(pdf_file, "rb") as f:
                    st.download_button("⬇️ Download PI PDF", f, file_name="Proforma_Invoice.pdf")

                os.remove(pdf_file)
