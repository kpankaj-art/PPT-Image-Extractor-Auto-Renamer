import io
import re
import zipfile
import fitz  # PyMuPDF slide rendering
from PIL import Image
from pptx import Presentation
import streamlit as st

st.set_page_config(page_title="PPT Visual Crop & Renamer", layout="centered")
st.title("PPT Right Image (Exact Markup) Extractor")


def extract_metadata_from_slide(slide):
    text_data = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            text_data.append(shape.text_frame.text)

    full_text = " ".join(text_data)
    clean_text = re.sub(r"\s+", " ", full_text)

    outlet_name = "OUTLET"
    contact = "0000000000"
    media_type = "NL"
    size = "0x0"

    outlet_match = re.search(
        r"Outlet\s*Name\s*:\s*([^:\n\r]+?)(?=\s*Address|\s*City|\s*Contact|\s*Type|$)",
        clean_text,
        re.IGNORECASE,
    )
    if outlet_match:
        raw_name = outlet_match.group(1).strip()
        cleaned = re.sub(r"[^\w\s-]", "", raw_name)
        outlet_name = re.sub(r"\s+", "_", cleaned).upper().strip("_")

    contact_match = re.search(
        r"(?:Contact|Mobile|Phone)?\s*:?\s*([6-9]\d{9})",
        clean_text,
        re.IGNORECASE,
    )
    if contact_match:
        contact = contact_match.group(1).strip()

    type_match = re.search(
        r"Type\s*:\s*([A-Za-z0-9_-]+)", clean_text, re.IGNORECASE
    )
    if type_match:
        media_type = type_match.group(1).strip().upper()

    size_match = re.search(
        r"Size\s*:\s*(\d+)\s*[*xX]\s*(\d+)", clean_text, re.IGNORECASE
    )
    if size_match:
        size = f"{size_match.group(1)}x{size_match.group(2)}"

    return f"{outlet_name}_{contact}_{media_type}_{size}"


uploaded_file = st.file_uploader("Apni PPTX File Upload Karein", type=["pptx"])

if uploaded_file is not None:
    if st.button("Extract Right Image with Exact Markup"):
        with st.spinner("Processing PPT & cropping visual markup..."):
            file_bytes = uploaded_file.read()
            prs = Presentation(io.BytesIO(file_bytes))

            # PyMuPDF se PPT Document open karke slides render karenge
            doc = fitz.open(stream=file_bytes, filetype="pptx")

            zip_buffer = io.BytesIO()
            extracted_count = 0

            with zipfile.ZipFile(
                zip_buffer, "a", zipfile.ZIP_DEFLATED, False
            ) as zip_file:
                for idx, page in enumerate(doc):
                    slide = prs.slides[idx]
                    filename_prefix = extract_metadata_from_slide(slide)

                    # Slide ko High DPI (300 DPI equivalent) par render karna
                    zoom = 3
                    mat = fitz.Matrix(zoom, zoom)
                    pix = page.get_pixmap(matrix=mat)

                    # PIL Image me convert karna
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                    # Right Side Far View Box coordinates (Relative Slide Box)
                    # Far View Box Range: Left ~52% to 95%, Top ~32% to 80%
                    w, h = img.size
                    crop_area = (
                        int(w * 0.515),  # Left
                        int(h * 0.315),  # Top
                        int(w * 0.955),  # Right
                        int(h * 0.810),  # Bottom
                    )

                    cropped_img = img.crop(crop_area)

                    # Bytes stream me save karke ZIP me add karna
                    img_byte_arr = io.BytesIO()
                    cropped_img.save(img_byte_arr, format="PNG")

                    final_name = f"{filename_prefix}_1.png"
                    zip_file.writestr(final_name, img_byte_arr.getvalue())
                    extracted_count += 1

            doc.close()

            if extracted_count > 0:
                st.success(
                    f"Total {extracted_count} Images (With Exact Markup) successfully cropped & renamed!"
                )
                st.download_button(
                    label="Download Cropped ZIP",
                    data=zip_buffer.getvalue(),
                    file_name="outlet_far_views_with_markup.zip",
                    mime="application/zip",
                )
            else:
                st.warning("PPT me koi visual slides process nahi ho sake.")
