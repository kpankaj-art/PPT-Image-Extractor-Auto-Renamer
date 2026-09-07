import os
import gc
import tempfile
import zipfile
import streamlit as st
from pdf2image import convert_from_path, pdfinfo_from_path
import subprocess

st.set_page_config(page_title="PPT Image Extractor", layout="centered")
st.title("✂️ PPT Image Region Extractor")

uploaded_file = st.file_uploader("Upload PPTX File", type=["pptx"])

def pptx_to_pdf(pptx_path, output_dir):
    """Standard headless LibreOffice command."""
    cmd = [
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        pptx_path,
        "--outdir", output_dir
    ]
    subprocess.run(cmd, check=True)
    pdf_name = os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf"
    return os.path.join(output_dir, pdf_name)

if uploaded_file is not None:
    if st.button("Process & Extract Images"):
        with st.spinner("Processing slides..."):
            with tempfile.TemporaryDirectory() as temp_dir:
                # Save uploaded PPTX
                pptx_path = os.path.join(temp_dir, uploaded_file.name)
                with open(pptx_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                try:
                    # Convert PPTX to PDF
                    pdf_path = pptx_to_pdf(pptx_path, temp_dir)
                    
                    # Get page count
                    info = pdfinfo_from_path(pdf_path)
                    total_pages = info["Pages"]

                    extracted_images = []

                    # Process 1 page at a time (RAM efficient)
                    for page_num in range(1, total_pages + 1):
                        images = convert_from_path(
                            pdf_path,
                            first_page=page_num,
                            last_page=page_num,
                            dpi=150,
                            thread_count=1
                        )
                        
                        if images:
                            img = images[0]
                            img_filename = f"slide_{page_num}.jpg"
                            img_path = os.path.join(temp_dir, img_filename)
                            img.save(img_path, "JPEG", quality=90)
                            extracted_images.append((img_filename, img_path))
                            
                            del img
                            del images

                        gc.collect()

                    # Zip generated images
                    zip_path = os.path.join(temp_dir, "extracted_slides.zip")
                    with zipfile.ZipFile(zip_path, "w") as zipf:
                        for fname, fpath in extracted_images:
                            zipf.write(fpath, arcname=fname)

                    with open(zip_path, "rb") as zf:
                        st.download_button(
                            label="📥 Download Extracted Slides (ZIP)",
                            data=zf.read(),
                            file_name="extracted_slides.zip",
                            mime="application/zip"
                        )
                    st.success("Extraction Complete!")

                except Exception as e:
                    st.error(f"Error processing file: {str(e)}")
                finally:
                    gc.collect()
