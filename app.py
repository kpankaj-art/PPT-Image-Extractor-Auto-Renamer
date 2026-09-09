import io
import re
import zipfile
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import streamlit as st

st.set_page_config(
    page_title="PPT Auto-Crop & Renamer", layout="centered"
)
st.title("PPT Image Crop & Auto-Renamer")


def extract_metadata(slide):
    text_data = ""
    for shape in slide.shapes:
        if shape.has_text_frame:
            text_data += " " + shape.text_frame.text

    # Default values
    outlet_name = "OUTLET"
    contact = "0000000000"
    media_type = "NL"
    size = "0x0"

    # Regex Extraction
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
    "Apni PPTX File Upload Karein", type=["pptx"]
)

if uploaded_file is not None:
    if st.button("Extract & Auto-Name Images"):
        with st.spinner("Processing slides..."):
            prs = Presentation(uploaded_file)
            zip_buffer = io.BytesIO()
            extracted_count = 0

            with zipfile.ZipFile(
                zip_buffer, "a", zipfile.ZIP_DEFLATED, False
            ) as zip_file:
                for idx, slide in enumerate(prs.slides):
                    filename_prefix = extract_metadata(slide)
                    img_idx = 1

                    for shape in slide.shapes:
                        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                            image_bytes = shape.image.blob
                            image_ext = shape.image.ext

                            # Correct format naming
                            final_name = f"{filename_prefix}_{img_idx}.{image_ext}"
                            zip_file.writestr(final_name, image_bytes)
                            img_idx += 1
                            extracted_count += 1

            if extracted_count > 0:
                st.success(
                    f"Total {extracted_count} images successfully renamed & extracted!"
                )
                st.download_button(
                    label="Download Renamed ZIP",
                    data=zip_buffer.getvalue(),
                    file_name="renamed_outlet_images.zip",
                    mime="application/zip",
                )
            else:
                st.warning("PPT mein koi images nahi mili.")
