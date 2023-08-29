import os
import subprocess
import streamlit as st
from io import BytesIO
import pythoncom


def convert_to_pdf(file_path):
    file_name, file_extension = os.path.splitext(file_path)
    pdf_file_path = file_name + '.pdf'

    if file_path.lower().endswith('.pdf'):
        st.warning('The file is already in PDF format.')
        return file_path

    if os.name == 'nt':  # Windows
        try:
            pythoncom.CoInitialize()  # Initialize the COM library

            from win32com.client import Dispatch

            word = Dispatch('Word.Application')
            doc = word.Documents.Open(file_path)
            doc.SaveAs(pdf_file_path, FileFormat=17)
            doc.Close()
            word.Quit()
        except Exception as e:
            st.error(f'An error occurred while converting: {str(e)}')
            return None

    st.success('Conversion successful.')
    return pdf_file_path

def main():
    st.title('File to PDF Converter')
    file = st.file_uploader('Upload a file')

    if file is not None:
        file_name = file.name
        file_path = os.path.join(os.getcwd(), file_name)
        actual_name, file_type = file_name.split('.')
        with open(file_path, 'wb') as f:
            f.write(file.read())
        st.info(f'Converting file: {file_name}')
        converted_file = convert_to_pdf(file_path)
        if converted_file:
            with open(converted_file, 'rb') as f:
                pdf_bytes = f.read()
            st.success(f'Converted file saved as: {converted_file}')
            st.download_button('Download Converted PDF',
                               pdf_bytes,
                               file_name=f'{actual_name}.pdf')

if __name__ == '__main__':
    main()
