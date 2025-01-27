import streamlit as st
import pdfrw
import pandas as pd
from typing import Dict, Any
import logging
import io
import zipfile
from datetime import datetime

class PDFFormFiller:
    def __init__(self):
        """Initialize the PDF Form Filler"""
        # PDF form field constants
        self.ANNOT_KEY = '/Annots'
        self.ANNOT_FIELD_KEY = '/T'
        self.ANNOT_FORM_KEY = '/FT'
        self.ANNOT_FORM_TEXT = '/Tx'
        self.ANNOT_FORM_BUTTON = '/Btn'
        self.SUBTYPE_KEY = '/Subtype'
        self.WIDGET_SUBTYPE_KEY = '/Widget'
        
        # Set up logging
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)

    def get_pdf_fields(self, pdf_template):
        """Get all field names from the PDF template"""
        try:
            fields = set()
            for page in pdf_template.pages:
                if page[self.ANNOT_KEY]:
                    for annotation in page[self.ANNOT_KEY]:
                        if annotation[self.ANNOT_FIELD_KEY]:
                            if annotation[self.SUBTYPE_KEY] == self.WIDGET_SUBTYPE_KEY:
                                key = annotation[self.ANNOT_FIELD_KEY][1:-1]
                                fields.add(key)
            return sorted(list(fields))
        except Exception as e:
            self.logger.error(f"Error getting PDF fields: {str(e)}")
            raise

    def fill_pdf_form(self, template, row_data: Dict[str, Any]) -> pdfrw.PdfReader:
        """Fill a single PDF form with data from one Excel row."""
        try:
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
                                        annotation.update(pdfrw.PdfDict(
                                            V=field_value,
                                            AP=field_value
                                        ))
                                    elif annotation[self.ANNOT_FORM_KEY] == self.ANNOT_FORM_BUTTON:
                                        annotation.update(pdfrw.PdfDict(
                                            V=pdfrw.PdfName(field_value),
                                            AS=pdfrw.PdfName(field_value)
                                        ))
                                    
                                    annotation.update(pdfrw.PdfDict(Ff=1))
            
            template.Root.AcroForm.update(
                pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true'))
            )
            
            return template
            
        except Exception as e:
            self.logger.error(f"Error filling PDF form: {str(e)}")
            raise

def main():
    try:
        st.set_page_config(page_title="PDF Form Filler", page_icon="📄")
        
        st.title("PDF Form Filler")
        st.write("Upload your Excel file and PDF template to generate filled forms.")
        
        # Initialize PDF Form Filler
        pdf_filler = PDFFormFiller()
        
        # File uploaders
        excel_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])
        pdf_template = st.file_uploader("Upload PDF Template", type=['pdf'])
        
        if excel_file and pdf_template:
            # Read the files
            try:
                df = pd.read_excel(excel_file)
                template = pdfrw.PdfReader(pdf_template)
                
                # Display field information
                st.write("### PDF Fields and Excel Columns")
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write("PDF Form Fields:")
                    pdf_fields = pdf_filler.get_pdf_fields(template)
                    st.write(", ".join(pdf_fields))
                
                with col2:
                    st.write("Excel Columns:")
                    st.write(", ".join(df.columns))
                
                if st.button("Process Files"):
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    # Create ZIP file in memory
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        successful_count = 0
                        failed_count = 0
                        
                        for index, row in df.iterrows():
                            try:
                                # Update progress
                                progress = (index + 1) / len(df)
                                progress_bar.progress(progress)
                                status_text.text(f"Processing record {index + 1}/{len(df)}")
                                
                                # Create filename
                                identifier = row.get('Account_ID', index)
                                pdf_filename = f"filled_form_{identifier}.pdf"
                                
                                # Fill the PDF
                                filled_pdf = pdf_filler.fill_pdf_form(pdfrw.PdfReader(pdf_template), row.to_dict())
                                
                                # Convert to bytes and add to ZIP
                                pdf_buffer = io.BytesIO()
                                pdfrw.PdfWriter().write(pdf_buffer, filled_pdf)
                                zip_file.writestr(pdf_filename, pdf_buffer.getvalue())
                                
                                successful_count += 1
                                
                            except Exception as e:
                                failed_count += 1
                                st.error(f"Error processing record {index}: {str(e)}")
                    
                    # Prepare download
                    zip_buffer.seek(0)
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    
                    # Show completion status
                    st.success(f"""
                    Processing complete!
                    - Successfully processed: {successful_count} records
                    - Failed to process: {failed_count} records
                    """)
                    
                    # Download button
                    st.download_button(
                        label="Download ZIP File",
                        data=zip_buffer.getvalue(),
                        file_name=f"filled_forms_{timestamp}.zip",
                        mime="application/zip"
                    )
            
            except Exception as e:
                st.error(f"Error processing files: {str(e)}")
    
    except Exception as e:
        st.error(f"Application error: {str(e)}")

if __name__ == "__main__":
    main()
