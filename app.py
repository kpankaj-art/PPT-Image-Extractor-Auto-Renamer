import os
import subprocess
import tempfile
import io
import zipfile
import streamlit as st
from pdf2image import convert_from_path

st.title("PPT Slide to Image Converter (With Markings)")

uploaded_file = st.file_uploader("Apni PPTX File Upload Karein", type=["pptx"])

if uploaded_file is not None:
    with st.spinner("Processing PPT..."):
        # Temporary directory create karein
        with tempfile.TemporaryDirectory() as temp_dir:
            ppt_path = os.path.join(temp_dir, "uploaded.pptx")
            pdf_path = os.path.join(temp_dir, "uploaded.pdf")

            # Uploaded PPT file save karein
            with open(ppt_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # PPT ko PDF mein convert karein (LibreOffice se)
            subprocess.run([
                "soffice", "--headless", "--convert-to", "pdf", ppt_path, "--outdir", temp_dir
            ], check=True)

            # PDF ki har slide ko high resolution PNG mein convert karein
            images = convert_from_path(pdf_path, dpi=200)

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for idx, img in enumerate(images):
                    img_byte_arr = io.BytesIO()
                    img.save(img_byte_arr, format='PNG')
                    zip_file.writestr(f"slide_{idx + 1}_marked.png", img_byte_arr.getvalue())

            st.success(f"Kul {len(images)} slides marking ke sath convert ho gayi hain!")
            
            st.download_button(
                label="Marking wali Saari Images ZIP Mein Download Karein",
                data=zip_buffer.getvalue(),
                file_name="marked_slides.zip",
                mime="application/zip",
            )
