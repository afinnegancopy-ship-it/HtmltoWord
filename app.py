# Streamlit app: Excel (ID + HTML) -> Word (.docx)
# This reverses your Word->HTML->Excel pipeline as closely as possible

import streamlit as st
from docx import Document
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime
from bs4 import BeautifulSoup

# ============================
# HTML -> Word helpers
# ============================

def add_html_to_doc(doc, html):
    """Convert a limited HTML subset back into Word paragraphs."""
    soup = BeautifulSoup(html, "html.parser")

    for element in soup.contents:
        if element.name == "p":
            p = doc.add_paragraph()
            add_inline_runs(p, element)

        elif element.name == "ul":
            for li in element.find_all("li", recursive=False):
                p = doc.add_paragraph(style="List Bullet")
                add_inline_runs(p, li)


def add_inline_runs(paragraph, element):
    """Handle <b>, <strong>, and plain text."""
    for node in element.descendants:
        if node.name in ("b", "strong"):
            run = paragraph.add_run(node.get_text())
            run.bold = True
        elif node.name is None:
            text = node.strip()
            if text:
                paragraph.add_run(text)


# ============================
# Excel -> Word
# ============================

def excel_to_word(excel_file):
    wb = load_workbook(excel_file)
    ws = wb.active

    doc = Document()

    # Skip header row
    for row in ws.iter_rows(min_row=2, values_only=True):
        product_id, html = row
        if not product_id or not html:
            continue

        # ID as heading / separator
        doc.add_paragraph(str(product_id)).runs[0].bold = True

        add_html_to_doc(doc, html)

        doc.add_page_break()

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# ============================
# Streamlit UI
# ============================

st.title("Excel to Word Converter")
st.write("Upload an Excel file generated from the Word→HTML tool")

uploaded_file = st.file_uploader("Choose an Excel (.xlsx) file", type=["xlsx"])

if uploaded_file:
    st.success(f"File uploaded: {uploaded_file.name}")

    if st.button("Convert to Word"):
        with st.spinner("Converting..."):
            word_file = excel_to_word(uploaded_file)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"reconstructed_{timestamp}.docx"

        st.download_button(
            label="Download Word Document",
            data=word_file,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        st.success("Conversion complete!")
