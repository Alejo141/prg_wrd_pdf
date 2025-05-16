import streamlit as st
import subprocess
import tempfile
import os
import zipfile
from io import BytesIO

st.title("Conversor de Word a PDF")

uploaded_files = st.file_uploader(
    "Sube tus archivos Word (.docx)",
    type=["docx"],
    accept_multiple_files=True
)

def convert_to_pdf(docx_path, output_folder):
    """Convierte DOCX a PDF usando LibreOffice"""
    try:
        subprocess.run([
            'libreoffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', output_folder,
            docx_path
        ], check=True)
        return True
    except subprocess.CalledProcessError as e:
        st.error(f"Error de conversión: {str(e)}")
        return False

if st.button("Convertir a PDF"):
    if not uploaded_files:
        st.warning("Por favor sube al menos un archivo Word.")
    else:
        with tempfile.TemporaryDirectory() as temp_dir:
            pdf_paths = []
            
            for uploaded_file in uploaded_files:
                try:
                    # Guardar archivo Word temporal
                    docx_path = os.path.join(temp_dir, uploaded_file.name)
                    with open(docx_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    
                    # Convertir a PDF
                    if convert_to_pdf(docx_path, temp_dir):
                        pdf_filename = os.path.splitext(uploaded_file.name)[0] + ".pdf"
                        pdf_path = os.path.join(temp_dir, pdf_filename)
                        
                        if os.path.exists(pdf_path):
                            pdf_paths.append(pdf_path)
                            st.success(f"Convertido: {uploaded_file.name} -> {pdf_filename}")
                        else:
                            st.error(f"No se pudo generar el PDF para {uploaded_file.name}")
                
                except Exception as e:
                    st.error(f"Error procesando {uploaded_file.name}: {str(e)}")

            # Crear ZIP con los PDFs
            if pdf_paths:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for pdf_path in pdf_paths:
                        with open(pdf_path, "rb") as f:
                            zip_file.writestr(os.path.basename(pdf_path), f.read())
                
                st.download_button(
                    label="Descargar todos los PDFs",
                    data=zip_buffer.getvalue(),
                    file_name="documentos_convertidos.zip",
                    mime="application/zip"
                )