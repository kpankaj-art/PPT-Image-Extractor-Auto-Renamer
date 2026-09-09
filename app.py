import io
import os
import subprocess
import tempfile
import zipfile
from pdf2image import convert_from_path
import streamlit as st

st.set_page_config(page_title="PPT Slide Extractor", layout="centered")
st.title("PPT Slide to Image (With Markings)")

uploaded_file = st.file_uploader("Apni PPTX File Upload Karein", type=["pptx", "ppt"])

if uploaded_file is not None:
    if st.button("Convert & Download Images"):
        with st.spinner("Processing slides with markings..."):
            with tempfile.TemporaryDirectory() as temp_dir:
                ppt_path = os.path.join(temp_dir, "input.pptx")
                
                with open(ppt_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                # Convert PPTX to PDF using LibreOffice
                subprocess.run(
                    ["libreoffice", "--headless", "--convert-to", "pdf", ppt_path, "--outdir", temp_dir],
                    check=True
                )

                pdf_path = os.path.join(temp_dir, "input.pdf")
                
                # Convert PDF pages to high-quality images
                images = convert_from_path(pdf_path, dpi=200)

                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for idx, img in enumerate(images):
                        img_byte_arr = io.BytesIO()
                        img.save(img_byte_arr, format="PNG")
                        zip_file.writestr(f"slide_{idx + 1}_marked.png", img_byte_arr.getvalue())

                st.success(f"Total {len(images)} slides successfully converted!")
                st.download_button(
                    label="Download Merged Images (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="slides_with_markings.zip",
                    mime="application/zip",
                )
