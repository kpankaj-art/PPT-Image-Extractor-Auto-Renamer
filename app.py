import streamlit as st
import os
import re
import zipfile
import io
from pptx import Presentation
from PIL import Image
import fitz  # PyMuPDF (agar PPTX to PDF convert karke exact crop karna ho)

st.set_page_config(page_title="PPT Image Extractor & Renamer", layout="wide")

st.title("📸 PPT Image Extractor & Cropper")
st.write("PPTX upload karein, details extract karein aur Left/Right images ko rename karke crop karein.")

# 1. Sidebar Options
st.sidebar.header("⚙️ Options")
image_position = st.sidebar.radio("Konsi Image Extraction chahiye?", ["Left Image", "Right Image", "Dono (Both)"])

uploaded_file = st.sidebar.file_uploader("PPTX File Upload Karein", type=["pptx"])

# Helper function to extract text safely
def extract_text_from_slide(slide):
    text_data = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                text_data.append(paragraph.text.strip())
    return " ".join(text_data)

# Helper function to extract metadata using Regex
def parse_slide_metadata(full_text):
    # Default values
    name = "UNKNOWN_NAME"
    phone = "0000000000"
    media_type = "NL"
    size = "0x0"

    # Extract Outlet Name
    name_match = re.search(r"Outlet Name:\s*([^\n\r]+)", full_text, re.IGNORECASE)
    if name_match:
        name = name_match.group(1).split("Address:")[0].strip()
        name = re.sub(r'[^A-Za-z0-9_]+', '_', name).upper()

    # Extract Contact Number
    phone_match = re.search(r"Contact No:\s*(\d+)", full_text, re.IGNORECASE)
    if phone_match:
        phone = phone_match.group(1).strip()

    # Extract Type (e.g. Type: NL)
    type_match = re.search(r"Type:\s*([A-Za-z0-9_]+)", full_text, re.IGNORECASE)
    if type_match:
        media_type = type_match.group(1).strip()

    # Extract Size (e.g. Size: 96 x 18 -> 96x18)
    size_match = re.search(r"Size:\s*(\d+)\s*x\s*(\d+)", full_text, re.IGNORECASE)
    if size_match:
        size = f"{size_match.group(1)}x{size_match.group(2)}"

    return name, phone, media_type, size

if uploaded_file is not None:
    prs = Presentation(uploaded_file)
    total_slides = len(prs.slides)
    st.info(f"Total Slides: {total_slides}")

    start_button = st.button("🚀 Start Image Extraction")

    if start_button:
        zip_buffer = io.BytesIO()
        processed_count = 0

        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, slide in enumerate(prs.slides):
                status_text.text(f"Processing Slide {i+1} of {total_slides}...")
                
                # Slide se text padhna
                slide_text = extract_text_from_slide(slide)
                name, phone, media_type, size = parse_slide_metadata(slide_text)

                # Slide se images filter karna (X position ke basis par Left/Right decide karna)
                slide_images = []
                for shape in slide.shapes:
                    if shape.shape_type == 13:  # 13 represents Picture shape in python-pptx
                        slide_images.append((shape.left, shape.image))

                # Left to Right sort karna
                slide_images.sort(key=lambda x: x[0])

                if len(slide_images) >= 2:
                    left_img = slide_images[0][1]
                    right_img = slide_images[1][1]

                    # User selection ke basis par process karna
                    targets = []
                    if image_position in ["Left Image", "Dono (Both)"]:
                        targets.append(("LEFT", left_img))
                    if image_position in ["Right Image", "Dono (Both)"]:
                        targets.append(("RIGHT", right_img))

                    for pos_label, img_obj in targets:
                        # Naya Naming Format: NAME_PHONE_MEDIATYPE_SIZE.jpg
                        # Custom Naming format: KAILASH_SHOE_7004609398_NL_8x4.jpg
                        file_name = f"{name}_{phone}_{media_type}_{size}.jpg"
                        
                        # Agardono images le rahe hain toh distinguish karne ke liye suffix:
                        if image_position == "Dono (Both)":
                            file_name = f"{name}_{phone}_{media_type}_{size}_{pos_label}.jpg"

                        img_bytes = img_obj.blob
                        zip_file.writestr(file_name, img_bytes)
                        processed_count += 1
                
                elif len(slide_images) == 1 and image_position != "Right Image":
                    # Single image case
                    file_name = f"{name}_{phone}_{media_type}_{size}.jpg"
                    zip_file.writestr(file_name, slide_images[0][1].blob)
                    processed_count += 1

                progress_bar.progress((i + 1) / total_slides)

        status_text.success(f"Processing Complete! Total {processed_count} images extracted.")

        # Download ZIP Button
        st.download_button(
            label="📦 Extracted Images ZIP Download Karein",
            data=zip_buffer.getvalue(),
            file_name="Extracted_Images.zip",
            mime="application/zip"
        )
