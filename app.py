import streamlit as st
import os
import re
import zipfile
import io
from pptx import Presentation
from PIL import Image

st.set_page_config(page_title="PPT Screenshot Cropper & Renamer", layout="wide")

st.title("📸 PPT Image Cropper & Renamer")
st.write("PPTX Upload karein, direct slide se elements extract karein aur exact Left/Right region crop karke download karein.")

# Sidebar Options
st.sidebar.header("⚙️ Selection Options")
image_position = st.sidebar.radio("Konsi Image Crop Chahiye?", ["Left Image", "Right Image", "Dono (Both)"])

uploaded_file = st.sidebar.file_uploader("PPTX File Upload Karein", type=["pptx"])

# Function to safely extract all text from slide shapes and tables
def extract_slide_text(slide):
    text_blocks = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            text_blocks.append(shape.text_frame.text)
        elif shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    text_blocks.append(cell.text)
    return "\n".join(text_blocks)

# Improved Parser for Metadata Extraction
def parse_slide_data(full_text):
    # Default Fallbacks
    outlet_name = "OUTLET"
    contact_no = "0000000000"
    type_val = "NL"
    size_val = "0x0"

    # Extract Outlet Name
    name_match = re.search(r"Outlet Name:\s*([^\n\r]+)", full_text, re.IGNORECASE)
    if name_match:
        raw_name = name_match.group(1).split("Address:")[0].strip()
        outlet_name = re.sub(r'[^A-Za-z0-9]+', '_', raw_name).strip('_').upper()

    # Extract Contact Number
    phone_match = re.search(r"Contact No:\s*(\d{10})", full_text, re.IGNORECASE)
    if phone_match:
        contact_no = phone_match.group(1).strip()
    else:
        # Fallback phone extraction (any 10 digits)
        digits = re.findall(r'\b\d{10}\b', full_text)
        if digits:
            contact_no = digits[0]

    # Extract Type (e.g., Type: NL)
    type_match = re.search(r"Type:\s*([A-Za-z0-9_]+)", full_text, re.IGNORECASE)
    if type_match:
        type_val = type_match.group(1).strip().upper()

    # Extract Size (e.g., Size: 96 x 18 -> 96x18)
    size_match = re.search(r"Size:\s*(\d+)\s*x\s*(\d+)", full_text, re.IGNORECASE)
    if size_match:
        size_val = f"{size_match.group(1)}x{size_match.group(2)}"

    return outlet_name, contact_no, type_val, size_val

if uploaded_file is not None:
    prs = Presentation(uploaded_file)
    total_slides = len(prs.slides)
    st.success(f"Total Slides Found: {total_slides}")

    if st.button("🚀 Start Crop & Rename Process"):
        zip_buffer = io.BytesIO()
        processed_count = 0

        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, slide in enumerate(prs.slides):
                status_text.text(f"Processing Slide {i+1} of {total_slides}...")

                # Slide Text Parsing
                slide_text = extract_slide_text(slide)
                outlet_name, contact_no, type_val, size_val = parse_slide_data(slide_text)

                # Extract Images based on Left-to-Right layout
                images = []
                for shape in slide.shapes:
                    if shape.shape_type == 13:  # Picture shape
                        images.append((shape.left, shape.image))

                # Sort pictures left to right
                images.sort(key=lambda x: x[0])

                selected_targets = []
                if image_position == "Left Image" and len(images) >= 1:
                    selected_targets.append(("LEFT", images[0][1]))
                elif image_position == "Right Image" and len(images) >= 2:
                    selected_targets.append(("RIGHT", images[1][1]))
                elif image_position == "Dono (Both)":
                    if len(images) >= 1:
                        selected_targets.append(("LEFT", images[0][1]))
                    if len(images) >= 2:
                        selected_targets.append(("RIGHT", images[1][2] if len(images) > 1 else images[0][1]))

                for pos_tag, img_obj in selected_targets:
                    # Construct exact filename format
                    if image_position == "Dono (Both)":
                        filename = f"{outlet_name}_{contact_no}_{type_val}_{size_val}_{pos_tag}.jpg"
                    else:
                        filename = f"{outlet_name}_{contact_no}_{type_val}_{size_val}.jpg"

                    zip_file.writestr(filename, img_obj.blob)
                    processed_count += 1

                progress_bar.progress((i + 1) / total_slides)

        status_text.success(f"Successfully processed {processed_count} images!")

        # ZIP Download Button
        st.download_button(
            label="📦 Download Renamed Images (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Renamed_Images.zip",
            mime="application/zip"
        )
