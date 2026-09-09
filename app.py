import io
import os
import subprocess
import tempfile
import zipfile
from pdf2image import convert_from_path
import streamlit as st

st.title("PPT Slide to Image Extractor (With Markings)")

uploaded_file = st.file_uploader("Apni PPTX File Upload Karein", type=["pptx"])

if uploaded_file is not None:
    with st.spinner("Processing PPT..."):
        with tempfile.TemporaryDirectory() as temp_dir:
            ppt_path = os.path.join(temp_dir, "uploaded.pptx")

            with open(ppt_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # 'soffice' ki jagah 'unoconv' command
            subprocess.run(["unoconv", "-f", "pdf", ppt_path], check=True)

            pdf_path = os.path.join(temp_dir, "uploaded.pdf")
            images = convert_from_path(pdf_path, dpi=200)

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(
                zip_buffer, "a", zipfile.ZIP_DEFLATED, False
            ) as zip_file:
                for idx, img in enumerate(images):
                    img_byte_arr = io.BytesIO()
                    img.save(img_byte_arr, format="PNG")
                    zip_file.writestr(
                        f"slide_{idx + 1}_marked.png", img_byte_arr.getvalue()
                    )

            st.success(
                f"Kul {len(images)} slides marking ke sath convert ho gayi hain!"
            )
            st.download_button(
                label="Download ZIP",
                data=zip_buffer.getvalue(),
                file_name="marked_slides.zip",
                mime="application/zip",
            )
