import io
import re
import zipfile
import fitz  # PyMuPDF
from PIL import Image
import streamlit as st

st.set_page_config(page_title="PDF Auto-Crop & Renamer", layout="centered")
st.title("PDF Image Crop & Auto-Renamer")


def extract_metadata_from_text(text_data):
    outlet_name = "OUTLET"
    contact = "0000000000"
    media_type = "NL"
    size = "0x0"

    outlet_match = re.search(
        r"Outlet Name:\s*([^\n\r]+)", text_data, re.IGNORECASE
    )
    if outlet_match:
        outlet_name = (
            outlet_match.group(1).strip().replace(" ", "_").upper()
        )

    contact_match = re.search(
        r"Contact No:\s*(\d{10})", text_data, re.IGNORECASE
    )
    if contact_match:
        contact = contact_match.group(1).strip()

    type_match = re.search(
        r"Type:\s*([A-Za-z0-9]+)", text_data, re.IGNORECASE
    )
    if type_match:
        media_type = type_match.group(1).strip().upper()

    size_match = re.search(
        r"Size:\s*(\d+)\s*x\s*(\d+)", text_data, re.IGNORECASE
    )
    if size_match:
        size = f"{size_match.group(1)}x{size_match.group(2)}"

    return f"{outlet_name}_{contact}_{media_type}_{size}"


uploaded_file = st.file_uploader(
    "Apni PDF File Upload Karein", type=["pdf"]
)

if uploaded_file is not None:
    if st.button("Extract, Crop & Auto-Name Images"):
        with st.spinner("Processing PDF pages..."):
            pdf_bytes = uploaded_file.read()
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            zip_buffer = io.BytesIO()
            extracted_count = 0

            with zipfile.ZipFile(
                zip_buffer, "a", zipfile.ZIP_DEFLATED, False
            ) as zip_file:
                for page_idx, page in enumerate(doc):
                    page_text = page.get_text()
                    filename_prefix = extract_metadata_from_text(page_text)

                    # Quality high karne ke liye zoom factor (3 = 300 DPI approx)
                    zoom = 3
                    mat = fitz.Matrix(zoom, zoom)

                    image_list = page.get_images(full=True)
                    img_idx = 1

                    for img_info in image_list:
                        xref = img_info[0]
                        rects = page.get_image_rects(xref)

                        for rect in rects:
                            # 📍 YEHA PAR AAYEGA 'clip=rect'
                            pix = page.get_pixmap(matrix=mat, clip=rect)

                            img_data = pix.tobytes("png")
                            final_name = (
                                f"{filename_prefix}_{img_idx}.png"
                            )

                            zip_file.writestr(final_name, img_data)
                            img_idx += 1
                            extracted_count += 1

            doc.close()

            if extracted_count > 0:
                st.success(
                    f"Total {extracted_count} images successfully cropped & saved!"
                )
                st.download_button(
                    label="Download Cropped ZIP",
                    data=zip_buffer.getvalue(),
                    file_name="cropped_outlet_images.zip",
                    mime="application/zip",
                )
            else:
                st.warning("PDF me koi images nahi mili.")
