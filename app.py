import streamlit as st
import os
import re
import zipfile
import io
from pptx import Presentation

st.set_page_config(page_title="PPT Image Extractor", layout="wide")
st.title("📸 PPT Image Extractor & Renamer (Fixed)")

# Sidebar Controls
st.sidebar.header("⚙️ Selection")
image_position = st.sidebar.radio("Konsi Image Extraction chahiye?", ["Left Image", "Right Image", "Dono (Both)"])

uploaded_file = st.sidebar.file_uploader("PPTX File Upload Karein", type=["pptx"])

# Sub-function to extract text even from grouped shapes & tables
def get_shape_text(shape):
    texts = []
    if shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            if paragraph.text.strip():
                texts.append(paragraph.text.strip())
    elif shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    texts.append(cell.text.strip())
    elif shape.shape_type == 6:  # Group Shape
        for sub_shape in shape.shapes:
            texts.extend(get_shape_text(sub_shape))
    return texts

def parse_slide_metadata(slide):
    all_texts = []
    for shape in slide.shapes:
        all_texts.extend(get_shape_text(shape))
    
    full_str = " \n ".join(all_texts)

    # Defaults
    name = "OUTLET"
    phone = "0000000000"
    media_type = "NL"
    size = "0x0"

    # 1. Extract Outlet Name (Looks for "Outlet Name:" or anything before "Address:" or line break)
    name_match = re.search(r"Outlet\s*Name\s*:?\s*([^\n\r]+)", full_str, re.IGNORECASE)
    if name_match:
        raw_name = name_match.group(1)
        raw_name = re.split(r"Address\s*:", raw_name, flags=re.IGNORECASE)[0]
        name = re.sub(r'[^A-Za-z0-9]+', '_', raw_name).strip('_').upper()

    # 2. Extract Phone Number (Looks for "Contact" label OR any 10-digit number)
    phone_match = re.search(r"(?:Contact|Mob|Mobile|Phone)?\s*(?:No|Num|Number)?\s*:?\s*(\d{10})", full_str, re.IGNORECASE)
    if phone_match:
        phone = phone_match.group(1)
    else:
        all_10_digits = re.findall(r'\b\d{10}\b', full_str)
        if all_10_digits:
            phone = all_10_digits[0]

    # 3. Extract Type (Looks for "Type:" or standalone keywords like NL, BL, GSB)
    type_match = re.search(r"Type\s*:?\s*([A-Za-z0-9_-]+)", full_str, re.IGNORECASE)
    if type_match:
        media_type = type_match.group(1).strip().upper()

    # 4. Extract Size (e.g., 96 x 18 or 96x18)
    size_match = re.search(r"Size\s*:?\s*(\d+)\s*x\s*(\d+)", full_str, re.IGNORECASE)
    if size_match:
        size = f"{size_match.group(1)}x{size_match.group(2)}"
    else:
        # Generic pattern like 8x4 or 96x18 anywhere in text
        generic_size = re.search(r'\b(\d{1,3})\s*x\s*(\d{1,3})\b', full_str, re.IGNORECASE)
        if generic_size:
            size = f"{generic_size.group(1)}x{generic_size.group(2)}"

    return name, phone, media_type, size


if uploaded_file is not None:
    prs = Presentation(uploaded_file)
    st.info(f"Total Slides: {len(prs.slides)}")

    if st.button("🚀 Start Extraction & Rename"):
        zip_buffer = io.BytesIO()
        processed_count = 0

        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            progress_bar = st.progress(0)

            for i, slide in enumerate(prs.slides):
                # Extract details
                name, phone, media_type, size = parse_slide_metadata(slide)

                # Fetch all picture objects
                pictures = []
                for shape in slide.shapes:
                    if shape.shape_type == 13:  # Picture shape
                        pictures.append((shape.left, shape.image))

                # Sort pictures from left to right
                pictures.sort(key=lambda x: x[0])

                targets = []
                if image_position == "Left Image" and len(pictures) >= 1:
                    targets.append(("", pictures[0][1]))
                elif image_position == "Right Image" and len(pictures) >= 2:
                    targets.append(("", pictures[1][1]))
                elif image_position == "Right Image" and len(pictures) == 1:
                    targets.append(("", pictures[0][1]))
                elif image_position == "Dono (Both)":
                    if len(pictures) >= 1:
                        targets.append(("_LEFT", pictures[0][1]))
                    if len(pictures) >= 2:
                        targets.append(("_RIGHT", pictures[1][1]))

                for suffix, img_obj in targets:
                    # Duplicate files prevent karne ke liye slide index append
                    file_name = f"{name}_{phone}_{media_type}_{size}{suffix}.jpg"
                    
                    # Store image in ZIP
                    zip_file.writestr(file_name, img_obj.blob)
                    processed_count += 1

                progress_bar.progress((i + 1) / len(prs.slides))

        st.success(f"Success! {processed_count} Images extracted and renamed.")

        st.download_button(
            label="📦 Download Fixed ZIP",
            data=zip_buffer.getvalue(),
            file_name="Extracted_Images.zip",
            mime="application/zip"
        )
