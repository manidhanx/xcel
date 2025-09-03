import streamlit as st
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
import tempfile
import os

st.set_page_config(page_title="Excel Style Aggregator", layout="centered")

st.title("👕 Excel → PDF (Style Aggregator)")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    # Load raw Excel (no header)
    raw_df = pd.read_excel(uploaded_file, header=None)

    # --- 1. Find Order No ---
    order_no = None
    for i, row in raw_df.iterrows():
        for j, cell in enumerate(row):
            if str(cell).strip().lower() == "order no :":
                try:
                    order_no = row[j+2]  # 2 cells to the right
                except Exception:
                    order_no = None
                break
        if order_no:
            break

    # --- 2. Find the row index where "Style" appears ---
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
        df = df.dropna(how="all")  # drop completely empty rows

        st.write("### Preview of extracted data")
        st.dataframe(df.head())

        if "Style" not in df.columns:
            st.error("❌ No 'Style' column found after parsing the file.")
        else:
            # Aggregate by Style (sum numeric columns)
            agg = df.groupby("Style").sum(numeric_only=True).reset_index()

            st.write("### Aggregated Data (by Style)")
            st.dataframe(agg)

            # Generate PDF
            if st.button("Generate PDF"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmpfile:
                    pdf_file = tmpfile.name

                doc = SimpleDocTemplate(pdf_file, pagesize=A4)
                styles = getSampleStyleSheet()
                elements = []

                # Add Order No if found
                if order_no:
                    elements.append(Paragraph(f"<b>Order No:</b> {order_no}", styles["Normal"]))
                    elements.append(Spacer(1, 12))

                # Prepare table data
                data = [agg.columns.tolist()] + agg.values.tolist()

                table = Table(data)
                table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("GRID", (0, 0), (-1, -1), 1, colors.black),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ]))

                elements.append(table)
                doc.build(elements)

                with open(pdf_file, "rb") as f:
                    st.download_button("⬇️ Download PDF", f, file_name="style_report.pdf")

                os.remove(pdf_file)
