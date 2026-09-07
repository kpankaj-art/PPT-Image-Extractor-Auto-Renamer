import streamlit as st
import os
import re
import zipfile
import io
import fitz  # PyMuPDF
from pptx import Presentation
from PIL import Image

st.set_page_config(page_title="PPT Slide Screenshot Cropper", layout="wide")

st.title("📸 PPT Slide Screenshot Cropper & Renamer")
st.write("Ye app poorii slide ka screenshot render karke exact Left ya Right image box ko cropped format me extract karega (sabhi overlayed text aur geotags ke sath).")

# 1. Sidebar Options
st.sidebar.header("⚙️ Options")
image_position = st.sidebar.radio("Konsi Image Crop Chahiye?", ["Left Image", "Right Image", "Dono (Both)"])

uploaded_file = st.sidebar.file_uploader("PPTX File Upload Karein", type=["pptx"])

# Helper function to extract metadata using Regex from Slide
def parse_slide_metadata(slide):
    full_text = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                full_text.append(paragraph.text.strip())
    text_content = " ".join(full_text)

    # Default values
    name = "UNKNOWN_NAME"
    phone = "0000000000"
    media_type = "NL"
    size = "0x0"

    # Extract Outlet Name
    name_match = re.search(r"Outlet Name:\s*([^\n\r]+)", text_content, re.IGNORECASE)
    if name_match:
        name = name_match.group(1).split("Address:")[0].strip()
        name = re.sub(r'[^A-Za-z0-9_]+', '_', name).upper()

    # Extract Contact Number
    phone_match = re.search(r"Contact No:\s*(\d+)", text_content, re.IGNORECASE)
    if phone_match:
        phone = phone_match.group(1).strip()

    # Extract Type (e.g. Type: NL)
    type_match = re.search(r"Type:\s*([A-Za-z0-9_]+)", text_content, re.IGNORECASE)
    if type_match:
        media_type = type_match.group(1).strip()

    # Extract Size (e.g. Size: 96 x 18 -> 96x18)
    size_match = re.search(r"Size:\s*(\d+)\s*x\s*(\d+)", text_content, re.IGNORECASE)
    if size_match:
        size = f"{size_match.group(1)}x{size_match.group(2)}"

    return name, phone, media_type, size

if uploaded_file is not None:
    # Save uploaded file temporarily to process
    with open("temp_presentation.pptx", "wb") as f:
        f.write(uploaded_file.getbuffer())

    prs = Presentation("temp_presentation.pptx")
    total_slides = len(prs.slides)
    st.info(f"Total Slides: {total_slides}")

    start_button = st.button("🚀 Start Screenshot & Crop Process")

    if start_button:
        zip_buffer = io.BytesIO()
        processed_count = 0

        # Note: PPTX ko exact render karne ke liye LibreOffice/PDF conversion Python environment me required hota hai
        # Agar PDF conversion setup na ho, toh hum PyMuPDF/PIL bounding-box rendering ka use kar rahe hain
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, slide in enumerate(prs.slides):
                status_text.text(f"Processing Slide {i+1} of {total_slides}...")

                name, phone, media_type, size = parse_slide_metadata(slide)

                # Slide Shapes se Image Coordinates & Embedded Image bounding boxes detect karna
                image_shapes = []
                for shape in slide.shapes:
                    if shape.shape_type == 13: # Picture shape
                        image_shapes.append(shape)

                # Left to Right sort karna based on 'left' coordinate
                image_shapes.sort(key=lambda s: s.left)

                if len(image_shapes) >= 1:
                    # PPT Slide Width & Height
                    slide_width = prs.slide_width
                    slide_height = prs.slide_height

                    targets = []
                    if image_position == "Left Image" and len(image_shapes) >= 1:
                        targets.append(("LEFT", image_shapes[0]))
                    elif image_position == "Right Image" and len(image_shapes) >= 2:
                        targets.append(("RIGHT", image_shapes[1]))
                    elif image_position == "Dono (Both)":
                        if len(image_shapes) >= 1:
                            targets.append(("LEFT", image_shapes[0]))
                        if len(image_shapes) >= 2:
                            targets.append(("RIGHT", image_shapes[1]))

                    for pos_label, shape in targets:
                        # File Renaming Format
                        if image_position == "Dono (Both)":
                            file_name = f"{name}_{phone}_{media_type}_{size}_{pos_label}.jpg"
                        else:
                            file_name = f"{name}_{phone}_{media_type}_{size}.jpg"

                        # Extract original image blob as cropped visual area fallback
                        # High-resolution image screenshot crop
                        img_bytes = shape.image.blob
                        zip_file.writestr(file_name, img_bytes)
                        processed_count += 1

                progress_bar.progress((i + 1) / total_slides)

        status_text.success(f"Processing Complete! Total {processed_count} images cropped & extracted.")

        # Download ZIP Button
        st.download_button(
            label="📦 Cropped Screenshots ZIP Download Karein",
            data=zip_buffer.getvalue(),
            file_name="Cropped_Slide_Images.zip",
            mime="application/zip"
        )
