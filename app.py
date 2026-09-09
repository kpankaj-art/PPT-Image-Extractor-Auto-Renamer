import io
import os
import subprocess
import tempfile
import zipfile
import fitz  # PyMuPDF (Ultra Lightweight)
import streamlit as st

st.set_page_config(page_title="Lite PPT Extractor", layout="centered")
st.title("PPT Slide & Markings Extractor (1GB RAM Optimized)")

uploaded_file = st.file_uploader(
    "Apni PPTX File Upload Karein", type=["pptx", "ppt"]
)

if uploaded_file is not None:
    if st.button("Convert to Images (With Markings)"):
        with st.spinner("Processing slides with markings..."):
            try:
                with tempfile.TemporaryDirectory() as temp_dir:
                    ppt_path = os.path.join(temp_dir, "input.pptx")
                    pdf_path = os.path.join(temp_dir, "input.pdf")

                    # PPTX file temporarily save karein
                    with open(ppt_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                    # Headless/No-GUI mode se PDF banayein (Low RAM)
                    subprocess.run(
                        [
                            "soffice",
                            "--headless",
                            "--convert-to",
                            "pdf",
                            ppt_path,
                            "--outdir",
                            temp_dir,
                        ],
                        check=True,
                    )

                    # PyMuPDF se PDF pages ko image mein convert karein
                    doc = fitz.open(pdf_path)
                    zip_buffer = io.BytesIO()

                    with zipfile.ZipFile(
                        zip_buffer, "a", zipfile.ZIP_DEFLATED, False
                    ) as zip_file:
                        for idx, page in enumerate(doc):
                            # Resolution set karein (1.5x scale low RAM ke liye perfect hai)
                            pix = page.get_pixmap(dpi=150)
                            img_data = pix.tobytes("png")

                            zip_file.writestr(
                                f"slide_{idx + 1}_marked.png", img_data
                            )

                    st.success(
                        f"Total {len(doc)} slides successfully extracted with markings!"
                    )
                    st.download_button(
                        label="Download Merged Images (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name="marked_slides.zip",
                        mime="application/zip",
                    )

            except Exception as e:
                st.error(f"Error: {str(e)}")
