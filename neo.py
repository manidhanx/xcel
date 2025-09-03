if st.button("Generate Proforma Invoice"):
    with tempfile.NamedTemporaryFile(delete=False,suffix=".pdf") as tmp:
        pdf_file=tmp.name

    doc=SimpleDocTemplate(pdf_file,pagesize=A4,leftMargin=30,rightMargin=30,topMargin=30,bottomMargin=30)
    styles=getSampleStyleSheet()
    normal=styles["Normal"]
    bold=ParagraphStyle("bold",parent=normal,fontName="Helvetica-Bold")

    elements=[]
    content_width = A4[0] - 110   # tighter fit
    inner_width = content_width - 6
    table_width = inner_width - 6

    # --- Header with logo ---
    logo = Image("sarlogo.jpg", width=100, height=55)
    title_table = Table([
        [Paragraph("<font size=20><b>PROFORMA INVOICE</b></font>", bold), logo]
    ], colWidths=[0.75*inner_width, 0.25*inner_width])
    title_table.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.75,colors.black),
        ("ALIGN",(0,0),(0,0),"CENTER"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("ALIGN",(1,0),(1,0),"RIGHT"),
        ("FONTSIZE",(0,0),(-1,-1),20),
        ("TOPPADDING",(0,0),(-1,-1),8),
        ("BOTTOMPADDING",(0,0),(-1,-1),8),
    ]))
    elements.append(title_table)
    elements.append(Spacer(1,12))

    # --- Supplier & Consignee ---
    sup=[
        [Paragraph("<b>Supplier Name:</b> SAR APPARELS INDIA PVT.LTD.", normal), Paragraph(pi_no, normal)],
        [Paragraph("Address: 6, Picaso Bithi, Kolkata - 700017", normal), Paragraph("<b>Landmark order Reference:</b> "+str(order_no), normal)],
        [Paragraph("Phone: 9817473373", normal), Paragraph("<b>Buyer Name:</b> "+buyer_name, normal)],
        [Paragraph("Fax: N.A.", normal), Paragraph("<b>Brand Name:</b> "+brand_name, normal)],
    ]
    con=[
        [Paragraph("<b>Consignee:</b>", normal), Paragraph(payment_term, normal)],
        [Paragraph(consignee_name, normal), Paragraph("<b>Bank Details (Including Swift/IBAN):</b>", normal)],
        [Paragraph(consignee_addr, normal), Paragraph("Beneficiary: SAR APPARELS INDIA PVT.LTD", normal)],
        [Paragraph(consignee_tel, normal), Paragraph("Account No: 2112819952", normal)],
        ["", Paragraph("Bank: Kotak Mahindra Bank Ltd", normal)],
        ["", Paragraph("Address: 2 Brabourne Road, Govind Bhavan, Ground Floor, Kolkata - 700001", normal)],
        ["", Paragraph("SWIFT: KKBKINBBCPC", normal)],
        ["", Paragraph("Bank Code: 0323", normal)],
    ]
    info_table=Table(sup+con,colWidths=[0.5*inner_width,0.5*inner_width])
    info_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.black),
                                    ("VALIGN",(0,0),(-1,-1),"TOP"),
                                    ("FONTSIZE",(0,0),(-1,-1),8)]))
    elements.append(info_table)
    elements.append(Spacer(1,12))

    # --- Shipment Info ---
    ship=[
        [Paragraph("<b>Loading Country:</b> "+str(made_in), normal), Paragraph("<b>Port of Loading:</b> "+str(loading_port), normal)],
        [Paragraph("<b>Agreed Shipment Date:</b> "+str(ship_date), normal), Paragraph("<b>Description of goods:</b> "+str(order_of), normal)]
    ]
    ship_table=Table(ship,colWidths=[0.5*inner_width,0.5*inner_width])
    ship_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.black),
                                    ("FONTSIZE",(0,0),(-1,-1),8)]))
    elements.append(ship_table)
    elements.append(Spacer(1,12))

    # --- Items Table ---
    data=[list(agg_df.columns)]
    for _,row in agg_df.iterrows(): data.append(list(row))
    total_qty=agg_df["QTY"].sum()
    total_amount=agg_df["AMOUNT"].astype(float).sum()
    data.append(["TOTAL","","","","","",f"{int(total_qty):,}","USD",f"{total_amount:,.2f}"])

    col_widths = [
        table_width * 0.10,
        table_width * 0.21,
        table_width * 0.12,
        table_width * 0.10,
        table_width * 0.15,
        table_width * 0.08,
        table_width * 0.07,
        table_width * 0.08,
        table_width * 0.09
    ]

    table=Table(data,colWidths=col_widths,repeatRows=1)
    style=TableStyle([
        ("GRID",(0,0),(-1,-1),0.25,colors.black),
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#333333")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.whitesmoke),
        ("ALIGN",(0,0),(-1,0),"CENTER"),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,0),6.5),
        ("ALIGN",(0,1),(5,-1),"CENTER"),
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

    # --- Amount in Words ---
    amount_words=amount_to_words(total_amount)
    words_table=Table([[Paragraph(f"TOTAL  US DOLLAR {amount_words}", normal)]],colWidths=[inner_width])
    words_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.black),
                                     ("FONTSIZE",(0,0),(-1,-1),8)]))
    elements.append(words_table)

    # --- Terms & Conditions ---
    terms_table=Table([[Paragraph("Terms & Conditions (if any):", normal)]],colWidths=[inner_width])
    terms_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.black),
                                     ("FONTSIZE",(0,0),(-1,-1),8)]))
    elements.append(terms_table)
    elements.append(Spacer(1,24))

    # --- Signature ---
    sig_img = "sarsign.png"
    sign_table=Table([
        [Image(sig_img,width=150,height=50),
         Paragraph("Signed by ………………… for RNA Resources Group Ltd - Landmark (Babyshop)", normal)]
    ],colWidths=[0.5*inner_width,0.5*inner_width])
    sign_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.black),
                                    ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                                    ("ALIGN",(0,0),(0,0),"LEFT"),
                                    ("ALIGN",(1,0),(1,0),"RIGHT"),
                                    ("FONTSIZE",(0,0),(-1,-1),8)]))
    elements.append(sign_table)

    # --- Outer Frame ---
    outer_table = Table([[e] for e in elements], colWidths=[content_width])
    outer_table.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),1.5,colors.black),
        ("VALIGN",(0,0),(-1,-1),"TOP")
    ]))

    doc.build([outer_table])

    # --- Auto download ---
    with open(pdf_file, "rb") as f:
        pdf_bytes = f.read()
    b64 = base64.b64encode(pdf_bytes).decode()
    href = f"""
    <a id="autodl" href="data:application/pdf;base64,{b64}" download="Proforma_Invoice.pdf"></a>
    <script>document.getElementById('autodl').click();</script>
    """
    st.markdown(href, unsafe_allow_html=True)

    os.remove(pdf_file)
