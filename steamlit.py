import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
import datetime
import pdfrw

# PDFFormFiller class
class PDFFormFiller:
    def __init__(self):
        self.ANNOT_KEY = '/Annots'
        self.ANNOT_FIELD_KEY = '/T'
        self.ANNOT_FORM_KEY = '/FT'
        self.ANNOT_FORM_TEXT = '/Tx'
        self.ANNOT_FORM_BUTTON = '/Btn'
        self.SUBTYPE_KEY = '/Subtype'
        self.WIDGET_SUBTYPE_KEY = '/Widget'

    def upload_files(self):
        st.markdown("<h2 style='text-align: center; font-family: Arial, sans-serif; color: #1E3A8A;'>Upload Files</h2>", unsafe_allow_html=True)
        
        # Card-like UI for file upload section
        col1, col2 = st.columns([1, 1])

        with col1:
            uploaded_excel = st.file_uploader("Upload Excel File", type=["xlsx", "xls"], label_visibility="collapsed")
            if uploaded_excel:
                try:
                    self.excel_data = pd.read_excel(uploaded_excel)
                    st.success(f"Excel file uploaded: {uploaded_excel.name}")
                except Exception as e:
                    st.error(f"Error reading Excel file: {e}")

        with col2:
            uploaded_pdf = st.file_uploader("Upload PDF Template", type="pdf", label_visibility="collapsed")
            if uploaded_pdf:
                try:
                    self.pdf_template_bytes = uploaded_pdf.read()
                    self.pdf_template = pdfrw.PdfReader(BytesIO(self.pdf_template_bytes))
                    st.success(f"PDF template uploaded: {uploaded_pdf.name}")
                    self.print_pdf_fields()
                except Exception as e:
                    st.error(f"Error reading PDF template: {e}")

    def print_pdf_fields(self):
        if self.pdf_template:
            fields = set()
            for page in self.pdf_template.pages:
                if page[self.ANNOT_KEY]:
                    for annotation in page[self.ANNOT_KEY]:
                        if annotation[self.ANNOT_FIELD_KEY]:
                            if annotation[self.SUBTYPE_KEY] == self.WIDGET_SUBTYPE_KEY:
                                key = annotation[self.ANNOT_FIELD_KEY][1:-1]
                                fields.add(key)
            st.markdown("<h4 style='text-align: center; font-family: Arial, sans-serif; color: #4B5563;'>PDF Form Fields</h4>", unsafe_allow_html=True)
            st.write(", ".join(sorted(fields)))

            if self.excel_data is not None:
                st.markdown("<h4 style='text-align: center; font-family: Arial, sans-serif; color: #4B5563;'>Excel Columns</h4>", unsafe_allow_html=True)
                st.write(", ".join(self.excel_data.columns))

    def fill_pdf_form(self, row_data):
        template = pdfrw.PdfReader(BytesIO(self.pdf_template_bytes))
        for page in template.pages:
            if page[self.ANNOT_KEY]:
                for annotation in page[self.ANNOT_KEY]:
                    if annotation[self.ANNOT_FIELD_KEY]:
                        if annotation[self.SUBTYPE_KEY] == self.WIDGET_SUBTYPE_KEY:
                            key = annotation[self.ANNOT_FIELD_KEY][1:-1]
                            if key in row_data:
                                field_value = str(row_data[key])
                                if pd.isna(field_value) or field_value.lower() == 'nan':
                                    field_value = ''

                                if annotation[self.ANNOT_FORM_KEY] == self.ANNOT_FORM_TEXT:
                                    annotation.update(pdfrw.PdfDict(V=field_value, AP=field_value))
                                elif annotation[self.ANNOT_FORM_KEY] == self.ANNOT_FORM_BUTTON:
                                    annotation.update(pdfrw.PdfDict(V=pdfrw.PdfName(field_value), AS=pdfrw.PdfName(field_value)))
            annotation.update(pdfrw.PdfDict(Ff=1))
        template.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))
        return template

    def process_all_records(self):
        if self.excel_data is None or self.pdf_template is None:
            st.error("Please upload both Excel file and PDF template.")
            return

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_filename = f"filled_forms_{timestamp}.zip"

        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            successful_count = 0
            failed_count = 0

            for index, row in self.excel_data.iterrows():
                try:
                    identifier = row.get('Account_ID', index) 
                    pdf_filename = f"filled_form_{identifier}.pdf"
                    filled_pdf = self.fill_pdf_form(row.to_dict())
                    pdf_buffer = BytesIO()
                    pdfrw.PdfWriter().write(pdf_buffer, filled_pdf)
                    pdf_bytes = pdf_buffer.getvalue()
                    zip_file.writestr(pdf_filename, pdf_bytes)
                    successful_count += 1
                    st.write(f"Processed record {index + 1}/{len(self.excel_data)}: {pdf_filename}")
                except Exception as e:
                    failed_count += 1
                    st.error(f"Error processing record {index}: {str(e)}")

        zip_buffer.seek(0)
        zip_content = zip_buffer.getvalue()

        st.markdown("<h3 style='text-align: center; font-family: Arial, sans-serif; color: #10B981;'>Download Filled Forms</h3>", unsafe_allow_html=True)
        st.download_button(
            label="Download All Filled Forms",
            data=zip_content,
            file_name=zip_filename,
            mime='application/zip',
            use_container_width=True
        )

        st.write(f"\nProcessing complete!")
        st.write(f"Successfully processed: {successful_count} records")
        st.write(f"Failed to process: {failed_count} records")

# Main function to control the app flow
def main():
    st.markdown("""
    <style>
    /* General App Styles */
    body {
        font-family: 'Arial', sans-serif;
        background-color: #F3F4F6;
    }
    
    /* Styling Buttons */
    .stButton button {
        background-color: #3498db;
        color: white;
        padding: 15px 25px;
        font-size: 16px;
        border-radius: 10px;
        border: none;
        cursor: pointer;
        transition: transform 0.3s ease, background-color 0.3s ease;
    }

    .stButton button:hover {
        background-color: #2980b9;
        transform: scale(1.05);
    }

    .stFileUploader {
        margin: 0 auto;
        width: 80%;
        padding: 20px;
    }

    /* Card-like UI */
    .upload-card {
        background-color: white;
        border-radius: 10px;
        padding: 30px;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        margin: 10px;
        transition: transform 0.3s ease;
    }

    .upload-card:hover {
        transform: scale(1.03);
    }

    /* File upload */
    .stFileUploader input[type="file"] {
        background-color: #E5E7EB;
        padding: 15px;
        border-radius: 8px;
        transition: background-color 0.3s ease;
    }

    .stFileUploader input[type="file"]:hover {
        background-color: #D1D5DB;
    }

    /* Download Button */
    .stDownloadButton button {
        background-color: #27ae60;
        color: white;
        padding: 15px 25px;
        font-size: 16px;
        border-radius: 8px;
        border: none;
        cursor: pointer;
        transition: background-color 0.3s ease, transform 0.3s ease;
    }

    .stDownloadButton button:hover {
        background-color: #2ecc71;
        transform: scale(1.05);
    }
    </style>
    """, unsafe_allow_html=True)

    st.title("💼 PDF Form Filler")
    filler = PDFFormFiller()

    st.markdown("<div style='text-align: center;'><br></div>", unsafe_allow_html=True)

    filler.upload_files()

    st.markdown("<br><br>", unsafe_allow_html=True)
    if st.button("Process All Records", use_container_width=True):
        filler.process_all_records()

if __name__ == "__main__":
    main()

