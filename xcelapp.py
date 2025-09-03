import streamlit as st
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
import tempfile
import os

st.set_page_config(page_title="Excel → PDF Aggregator", layout="centered")

st.title("📊 Excel to PDF Aggregator")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    # Read Excel into DataFrame
    df = pd.read_excel(uploaded_file)
    st.write("### Preview of uploaded data")
    st.dataframe(df.head())

    # Let user pick a column to group by
    group_col = st.selectbox("Select a column to group by", df.columns)

    # Perform aggregation (sum of numeric columns)
    agg = df.groupby(group_col).sum(numeric_only=True).reset_index()
    st.write("### Aggregated Data")
    st.dataframe(agg)

    # Generate PDF
    if st.button("Generate PDF"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmpfile:
            pdf_file = tmpfile.name

        doc = SimpleDocTemplate(pdf_file, pagesize=A4)

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

        doc.build([table])

        with open(pdf_file, "rb") as f:
            st.download_button("⬇️ Download PDF", f, file_name="aggregated_report.pdf")

        # Cleanup temp file
        os.remove(pdf_file)
