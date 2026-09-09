import io
import re
import zipfile
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import streamlit as st

st.set_page_config(page_title="PPT Far-View Extractor", layout="centered")
st.title("PPT Right-Side Image (Far View) Extractor & Renamer")


def extract_metadata_from_slide(slide):
    """Slide ke sabhi text elements se metadata robustly read karta hai."""
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

    # 1. Outlet Name Match
    outlet_match = re.search(
        r"Outlet\s*Name\s*:\s*([^:\n\r]+?)(?=\s*Address|\s*City|\s*Contact|\s*Type|$)",
        clean_text,
        re.IGNORECASE,
    )
    if outlet_match:
        raw_name = outlet_match.group(1).strip()
        cleaned = re.sub(r"[^\w\s-]", "", raw_name)
        outlet_name = re.sub(r"\s+", "_", cleaned).upper().strip("_")

    # 2. Contact Match (Exact 10-digit Mobile Number)
    contact_match = re.search(
        r"(?:Contact|Mobile|Phone)?\s*:?\s*([6-9]\d{9})",
        clean_text,
        re.IGNORECASE,
    )
    if contact_match:
        contact = contact_match.group(1).strip()

    # 3. Type Match
    type_match = re.search(
        r"Type\s*:\s*([A-Za-z0-9_-]+)", clean_text, re.IGNORECASE
    )
    if type_match:
        media_type = type_match.group(1).strip().upper()

    # 4. Size Match (e.g. 10 x 2, 10x2)
    size_match = re.search(
        r"Size\s*:\s*(\d+)\s*[*xX]\s*(\d+)", clean_text, re.IGNORECASE
    )
    if size_match:
        size = f"{size_match.group(1)}x{size_match.group(2)}"

    return f"{outlet_name}_{contact}_{media_type}_{size}"


uploaded_file = st.file_uploader(
    "Apni PPTX File Upload Karein", type=["pptx"]
)

if uploaded_file is not None:
    if st.button("Extract Right Image & Auto-Rename"):
        with st.spinner("Processing PPT slides..."):
            prs = Presentation(uploaded_file)
            zip_buffer = io.BytesIO()
            extracted_count = 0

            slide_width = prs.slide_width

            with zipfile.ZipFile(
                zip_buffer, "a", zipfile.ZIP_DEFLATED, False
            ) as zip_file:
                for idx, slide in enumerate(prs.slides):
                    filename_prefix = extract_metadata_from_slide(slide)
                    img_idx = 1

                    for shape in slide.shapes:
                        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                            # 📍 LOGIC: Sirf right-side wali image filtering (Far View)
                            if shape.left > (slide_width / 2):
                                image_bytes = shape.image.blob
                                image_ext = shape.image.ext

                                final_name = (
                                    f"{filename_prefix}_{img_idx}.{image_ext}"
                                )
                                zip_file.writestr(final_name, image_bytes)

                                img_idx += 1
                                extracted_count += 1

            if extracted_count > 0:
                st.success(
                    f"Total {extracted_count} Right Images (Far View) successfully extracted & renamed!"
                )
                st.download_button(
                    label="Download Renamed ZIP",
                    data=zip_buffer.getvalue(),
                    file_name="outlet_far_views.zip",
                    mime="application/zip",
                )
            else:
                st.warning("PPT me koi Right-side Image (Far View) nahi mili.")
