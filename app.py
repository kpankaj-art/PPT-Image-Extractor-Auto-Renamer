import io
import os
import re
import zipfile
from pptx import Presentation
import streamlit as st

st.set_page_config(
    page_title="PPT Left Image Extractor", page_icon="🖼️", layout="centered"
)
st.title("🖼️ PPT Image Extractor & Auto-Renamer")
st.write(
    "Format: **OutletName_MobileNo_Type_Size.jpg** (Only Left Image Extracted)"
)


def clean_text(text):
    """File name safe characters"""
    if not text:
        return ""
    clean = re.sub(r'[\\/*?:"<>|\n\r\t]', " ", text)
    clean = re.sub(r"\s+", " ", clean).strip()
    return clean.replace(" ", "_")


def extract_info_from_slide(slide):
    """Accurate Extraction logic based on actual PPT layout"""
    all_lines = []

    # Slide ke sabhi shapes/text frames se lines collect karein
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                text = paragraph.text.strip()
                if text:
                    all_lines.append(text)

    full_text = "\n".join(all_lines)

    outlet_name = ""
    contact_no = ""
    media_type = ""
    size = ""

    # 1. Outlet Name: 'Outlet Name:' ke aage ka text (e.g., MD SHOE, GIRJA FOOTWEAR)
    outlet_match = re.search(
        r"Outlet\s*Name\s*[:\-]\s*(.*)", full_text, re.IGNORECASE
    )
    if outlet_match:
        raw_name = outlet_match.group(1).strip()
        # Agar address usi line me judi ho toh usko split karein
        outlet_name = re.split(
            r"Address|City|Contact|Installation|Type|Size",
            raw_name,
            flags=re.IGNORECASE,
        )[0].strip()

    # 2. Contact Number: Contact No: 7282007564 ya koi bhi 10 digit number
    contact_match = re.search(
        r"Contact\s*(?:No)?\s*[:\-]?\s*(\d{10})", full_text, re.IGNORECASE
    )
    if contact_match:
        contact_no = contact_match.group(1).strip()
    else:
        num_match = re.search(r"\b[6-9]\d{9}\b", full_text)
        if num_match:
            contact_no = num_match.group(0)

    # 3. Type: Dedicated search for 'Type:' box (e.g., NL, FL) - Strictly ignore 'Outlet'
    type_match = re.search(
        r"\bType\s*[:\-]\s*([A-Za-z0-9_\-]+)", full_text, re.IGNORECASE
    )
    if type_match:
        extracted_type = type_match.group(1).strip()
        if extracted_type.lower() != "outlet":
            media_type = extracted_type

    # 4. Size: Dedicated search for 'Size:' box (e.g., 10 x 3, 8 x 3)
    size_match = re.search(
        r"\bSize\s*[:\-]\s*(\d+\s*x\s*\d+)", full_text, re.IGNORECASE
    )
    if size_match:
        size = size_match.group(1).replace(" ", "").strip()

    return outlet_name, contact_no, media_type, size


uploaded_file = st.file_uploader("PowerPoint File Upload Karein (.pptx)", type=["pptx"])

if uploaded_file is not None:
    prs = Presentation(uploaded_file)
    zip_buffer = io.BytesIO()

    with st.spinner("Processing slides..."):
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for i, slide in enumerate(prs.slides):
                outlet_name, contact_no, media_type, size = (
                    extract_info_from_slide(slide)
                )

                # Slide ke sabhi picture shapes
                pic_shapes = [s for s in slide.shapes if s.shape_type == 13]

                if pic_shapes:
                    # Sirf LEFT side wali image pick karein
                    leftmost_pic = min(pic_shapes, key=lambda s: s.left)

                    if not outlet_name:
                        outlet_name = f"Slide_{i+1}"

                    # Final Naming Array
                    components = []
                    if outlet_name:
                        components.append(clean_text(outlet_name))
                    if contact_no:
                        components.append(clean_text(contact_no))
                    if media_type:
                        components.append(clean_text(media_type))
                    if size:
                        components.append(clean_text(size))

                    base_filename = "_".join(components)
                    ext = leftmost_pic.image.ext
                    final_name = f"{base_filename}.{ext}"

                    zip_file.writestr(final_name, leftmost_pic.image.blob)

    st.success("🎉 Processing Complete!")
    st.download_button(
        label="📥 Download Renamed Images (ZIP)",
        data=zip_buffer.getvalue(),
        file_name="Renamed_Images.zip",
        mime="application/zip",
    )
