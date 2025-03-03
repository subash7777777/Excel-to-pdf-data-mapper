import streamlit as st
import pandas as pd
import pdfrw
from io import BytesIO

class PDFProcessor:
    ANNOT_KEY = '/Annots'
    ANNOT_FIELD_KEY = '/T'
    ANNOT_FORM_KEY = '/FT'
    ANNOT_FORM_TEXT = '/Tx'
    ANNOT_FORM_BUTTON = '/Btn'
    SUBTYPE_KEY = '/Subtype'
    WIDGET_SUBTYPE_KEY = '/Widget'

    def __init__(self):
        self.pdf_template = None
        self.excel_data = None
        self.pdf_template_bytes = None  # Store original PDF bytes for reprocessing

    def load_pdf_template(self, uploaded_file):
        """Load a PDF form template."""
        self.pdf_template_bytes = uploaded_file.read()
        self.pdf_template = pdfrw.PdfReader(BytesIO(self.pdf_template_bytes))

    def load_excel_data(self, uploaded_file):
        """Load an Excel file and extract data."""
        self.excel_data = pd.read_excel(uploaded_file, dtype=str)
        self.excel_data.fillna('', inplace=True)  # Replace NaNs with empty strings

    def print_pdf_fields(self):
        """Extract and display form field names from the PDF."""
        if self.pdf_template:
            fields = set()
            for page in self.pdf_template.pages:
                if page[self.ANNOT_KEY]:
                    for annotation in page[self.ANNOT_KEY]:
                        if annotation[self.ANNOT_FIELD_KEY] and annotation[self.SUBTYPE_KEY] == self.WIDGET_SUBTYPE_KEY:
                            raw_key = annotation[self.ANNOT_FIELD_KEY]  
                            key = raw_key.strip('()') if isinstance(raw_key, str) else raw_key
                            fields.add(key)
                            st.write(f"🔍 Found Field: {repr(key)} (Raw: {repr(raw_key)})")

            st.write("\n📝 PDF Form Fields Detected:")
            st.write(sorted(fields))

            if self.excel_data is not None:
                st.write("\n📊 Excel Columns Found:")
                st.write([repr(col) for col in self.excel_data.columns])  # Debug

    def fill_pdf_form(self, row_data):
        """Fill a PDF template with row data from Excel."""
        template = pdfrw.PdfReader(BytesIO(self.pdf_template_bytes))

        for page in template.pages:
            if page[self.ANNOT_KEY]:
                for annotation in page[self.ANNOT_KEY]:
                    if annotation[self.ANNOT_FIELD_KEY] and annotation[self.SUBTYPE_KEY] == self.WIDGET_SUBTYPE_KEY:
                        raw_key = annotation[self.ANNOT_FIELD_KEY]
                        key = raw_key.strip('()') if isinstance(raw_key, str) else raw_key

                        if key in row_data:
                            field_value = str(row_data.get(key, '')).strip() if pd.notna(row_data.get(key, '')) else ''

                            # Ensure zip codes or numerical fields maintain format
                            if key.lower() in ['zipcode', 'postalcode']:
                                field_value = field_value.zfill(5)

                            # Force update form fields
                            if annotation[self.ANNOT_FORM_KEY] == self.ANNOT_FORM_TEXT:
                                annotation.update(pdfrw.PdfDict(V=field_value, Ff=0))  # Ensure it's writable
                            elif annotation[self.ANNOT_FORM_KEY] == self.ANNOT_FORM_BUTTON:
                                annotation.update(pdfrw.PdfDict(V=pdfrw.PdfName(field_value), AS=pdfrw.PdfName(field_value)))

                            st.write(f"✔️ Mapped: {repr(key)} → {repr(field_value)}")  # Debug

        # Force Acrobat to refresh fields
        template.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))
        return template

    def process_pdfs(self):
        """Fill and generate multiple PDFs for each Excel row."""
        if self.excel_data is None or self.pdf_template is None:
            st.error("❌ Please upload both a PDF template and an Excel file.")
            return

        output_pdfs = []
        for index, row in self.excel_data.iterrows():
            filled_pdf = self.fill_pdf_form(row)
            output_buffer = BytesIO()
            pdfrw.PdfWriter().write(output_buffer, filled_pdf)
            output_pdfs.append((row[0], output_buffer))  # Assuming first column is a unique identifier

        return output_pdfs

# Streamlit UI
st.title("📄 PDF Auto-Fill from Excel")

processor = PDFProcessor()

pdf_file = st.file_uploader("📂 Upload PDF Template", type=["pdf"])
excel_file = st.file_uploader("📊 Upload Excel File", type=["xlsx"])

if pdf_file and excel_file:
    processor.load_pdf_template(pdf_file)
    processor.load_excel_data(excel_file)

    if st.button("🔍 Show PDF Field Names"):
        processor.print_pdf_fields()

    if st.button("🚀 Process PDFs"):
        output_pdfs = processor.process_pdfs()
        if output_pdfs:
            for filename, pdf_buffer in output_pdfs:
                st.download_button(
                    label=f"📥 Download {filename}.pdf",
                    data=pdf_buffer.getvalue(),
                    file_name=f"{filename}.pdf",
                    mime="application/pdf",
                )
