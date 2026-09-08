import streamlit as st
import os
import re
import zipfile
import io
from pptx import Presentation

st.set_page_config(page_title="PPT Image Extractor & Renamer", layout="wide")

st.title("📸 PPT Image Extractor & Exact Renamer")

# Sidebar Controls
st.sidebar.header("⚙️ Settings")
image_position = st.sidebar.radio("Konsi Image Extraction chahiye?", ["Left Image", "Right Image", "Dono (Both)"])

uploaded_file = st.sidebar.file_uploader("PPTX File Upload Karein", type=["pptx"])

def clean_text(text):
    return re.sub(r'\s+', ' ', text).strip()

def parse_exact_metadata(slide):
    # Default Values
    outlet_name = "UNKNOWN_NAME"
    contact_no = "0000000000"
    type_val = "NL"
    size_val = "0x0"

    all_texts = []
    
    # Extract text from all shapes
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                txt = clean_text(paragraph.text)
                if txt:
                    all_texts.append(txt)

    full_content = " | ".join(all_texts)

    # 1. Extract Outlet Name
    name_match = re.search(r"Outlet\s*Name\s*:\s*([^|]+)", full_content, re.IGNORECASE)
    if name_match:
        raw_name = name_match.group(1)
        raw_name = re.sub(r"Address\s*:.*", "", raw_name, flags=re.IGNORECASE)  # Remove Address part
        outlet_name = re.sub(r'[^A-Za-z0-9]+', '_', raw_name).strip('_').upper()

    # 2. Extract Contact Number
    phone_match = re.search(r"Contact\s*(?:No|Number)?\s*:\s*(\d{10})", full_content, re.IGNORECASE)
    if phone_match:
        contact_no = phone_match.group(1)
    else:
        # Direct 10 digit fallback search
        all_digits = re.findall(r'\b\d{10}\b', full_content)
        if all_digits:
            contact_no = all_digits[0]

    # 3. Extract Type (e.g. Type: NL, Type: GSB, Type: BL)
    type_match = re.search(r"Type\s*:\s*([A-Za-z0-9_-]+)", full_content, re.IGNORECASE)
    if type_match:
        type_val = type_match.group(1).strip().upper()

    # 4. Extract Size (e.g. Size: 96 x 18 -> 96x18)
    size_match = re.search(r"Size\s*:\s*(\d+)\s*x\s*(\d+)", full_content, re.IGNORECASE)
    if size_match:
        size_val = f"{size_match.group(1)}x{size_match.group(2)}"

    return outlet_name, contact_no, type_val, size_val


if uploaded_file is not None:
    prs = Presentation(uploaded_file)
    st.info(f"Total Slides: {len(prs.slides)}")

    if st.button("🚀 Start Process & Rename Images"):
        zip_buffer = io.BytesIO()
        processed_count = 0

        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            progress = st.progress(0)
            status = st.empty()

            for i, slide in enumerate(prs.slides):
                status.text(f"Processing Slide {i+1}...")

                # Extract Details
                outlet_name, contact_no, type_val, size_val = parse_exact_metadata(slide)

                # Fetch Pictures
                pictures = []
                for shape in slide.shapes:
                    if shape.shape_type == 13: # Picture
                        pictures.append((shape.left, shape.image))

                # Sort pictures left-to-right
                pictures.sort(key=lambda item: item[0])

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
                    # Final Name Output Format
                    file_name = f"{outlet_name}_{contact_no}_{type_val}_{size_val}{suffix}.jpg"
                    zip_file.writestr(file_name, img_obj.blob)
                    processed_count += 1

                progress.progress((i + 1) / len(prs.slides))

        status.success(f"Successfully Extracted & Renamed {processed_count} Images!")

        st.download_button(
            label="📦 Download Renamed ZIP",
            data=zip_buffer.getvalue(),
            file_name="Renamed_Images_Output.zip",
            mime="application/zip"
        )
