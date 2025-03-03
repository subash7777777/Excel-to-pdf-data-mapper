import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
import datetime
import pdfrw

class PDFFormFiller:
    def __init__(self):
        self.ANNOT_KEY = '/Annots'
        self.ANNOT_FIELD_KEY = '/T'
        self.ANNOT_FORM_KEY = '/FT'
        self.ANNOT_FORM_TEXT = '/Tx'
        self.ANNOT_FORM_BUTTON = '/Btn'
        self.SUBTYPE_KEY = '/Subtype'
        self.WIDGET_SUBTYPE_KEY = '/Widget'

        self.excel_data = None
        self.pdf_template = None
        self.pdf_template_bytes = None

    def upload_files(self):
        uploaded_excel = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
        if uploaded_excel:
            try:
                self.excel_data = pd.read_excel(uploaded_excel, dtype=str)  # Read as string to avoid numeric issues
                self.excel_data.columns = self.excel_data.columns.str.strip()  # Remove unwanted spaces in column names
                st.success(f"Excel file uploaded: {uploaded_excel.name}")
            except Exception as e:
                st.error(f"Error reading Excel file: {e}")

        uploaded_pdf = st.file_uploader("Upload PDF Template", type="pdf")
        if uploaded_pdf:
            try:
                self.pdf_template_bytes = uploaded_pdf.read()
                self.pdf_template = pdfrw.PdfReader(BytesIO(self.pdf_template_bytes))
                st.success(f"PDF template uploaded: {uploaded_pdf.name}")
                self.print_pdf_fields()
            except Exception as e:
                st.error(f"Error reading PDF template: {e}")

    def print_pdf_fields(self):
        """Extract and display form field names from the PDF."""
        if self.pdf_template:
            fields = set()
            for page in self.pdf_template.pages:
                if page[self.ANNOT_KEY]:
                    for annotation in page[self.ANNOT_KEY]:
                        if annotation[self.ANNOT_FIELD_KEY] and annotation[self.SUBTYPE_KEY] == self.WIDGET_SUBTYPE_KEY:
                            key = annotation[self.ANNOT_FIELD_KEY][1:-1].strip()  # Remove leading/trailing spaces
                            fields.add(key)

            st.write("PDF form fields found:")
            st.write([repr(field) for field in sorted(fields)])  # Debug: Show exact field names

            if self.excel_data is not None:
                st.write("\nExcel columns found:")
                st.write([repr(col) for col in self.excel_data.columns])  # Debug: Show exact column names

    def fill_pdf_form(self, row_data):
        """Fill a PDF template with row data from Excel."""
        template = pdfrw.PdfReader(BytesIO(self.pdf_template_bytes))

        for page in template.pages:
            if page[self.ANNOT_KEY]:
                for annotation in page[self.ANNOT_KEY]:
                    if annotation[self.ANNOT_FIELD_KEY] and annotation[self.SUBTYPE_KEY] == self.WIDGET_SUBTYPE_KEY:
                        key = annotation[self.ANNOT_FIELD_KEY][1:-1].strip()  # Trim spaces
                        
                        if key in row_data:
                            field_value = str(row_data.get(key, '')).strip() if pd.notna(row_data.get(key, '')) else ''

                            # Ensure zip codes or other numerical fields maintain format
                            if key.lower() == 'zipcode':
                                field_value = field_value.zfill(5)

                            # Update text fields properly
                            if annotation[self.ANNOT_FORM_KEY] == self.ANNOT_FORM_TEXT:
                                annotation.update(pdfrw.PdfDict(V=field_value))
                            elif annotation[self.ANNOT_FORM_KEY] == self.ANNOT_FORM_BUTTON:
                                annotation.update(pdfrw.PdfDict(V=pdfrw.PdfName(field_value), AS=pdfrw.PdfName(field_value)))

                            st.write(f"Mapping: {key} -> {field_value}")  # Debug: Show field mapping progress

        # Ensure form updates are displayed properly
        template.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))
        return template

    def process_all_records(self):
        """Process all records from Excel and generate filled PDFs."""
        if self.excel_data is None or self.pdf_template is None:
            st.error("Please upload both Excel file and PDF template.")
            return

        if 'Account number' not in self.excel_data.columns:
            st.error("The Excel file must contain a column named 'Account number'.")
            return

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_filename = f"filled_forms_{timestamp}.zip"

        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            successful_count = 0
            failed_count = 0

            for index, row in self.excel_data.iterrows():
                try:
                    account_number = row.get('Account number', '').strip()
                    if not account_number:
                        st.warning(f"Missing 'Account number' for row {index + 1}. Skipping.")
                        failed_count += 1
                        continue

                    # Ensure account number is formatted correctly
                    account_number_str = str(int(float(account_number))) if account_number.isdigit() else account_number
                    pdf_filename = f"{account_number_str}.pdf"

                    filled_pdf = self.fill_pdf_form(row.to_dict())

                    # Save the modified PDF
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

        st.download_button(
            label="Download Filled Forms",
            data=zip_content,
            file_name=zip_filename,
            mime='application/zip'
        )

        st.write(f"\nProcessing complete!")
        st.write(f"Successfully processed: {successful_count} records")
        st.write(f"Failed to process: {failed_count} records")

def main():
    st.title("PDF Form Filler")
    filler = PDFFormFiller()
    filler.upload_files()

    if st.button("Process All Records"):
        filler.process_all_records()

if __name__ == "__main__":
    main()

