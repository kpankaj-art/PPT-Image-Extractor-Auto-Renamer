import io
import os
import re
import zipfile
from pptx import Presentation
import streamlit as st

st.set_page_config(
    page_title="PPT Left Image Extractor", page_icon="🖼️", layout="centered"
)
st.title("🖼️ PPT Left Image Extractor & Renamer")
st.write(
    "PPT upload karein. Har slide se sirf **LEFT SIDE** wali image extract karke rename kar di jayegi."
)


def clean_text(text):
    if not text:
        return ""
    clean = re.sub(r'[\\/*?:"<>|\n\r\t]', " ", text)
    clean = re.sub(r"\s+", " ", clean).strip()
    return clean.replace(" ", "_")


def extract_info(slide):
    all_text = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            all_text.append(shape.text_frame.text)

    full_text = " ".join(all_text)

    # 1. Outlet Name
    outlet_name = ""
    outlet_match = re.search(
        r"Outlet\s*Name\s*[:\-]?\s*([^Address|City|Contact|Installation|Type|Size|\n\r]+)",
        full_text,
        re.IGNORECASE,
    )
    if outlet_match:
        outlet_name = outlet_match.group(1).strip()

    # 2. Contact Number
    contact_no = ""
    contact_match = re.search(
        r"Contact\s*(?:No)?\s*[:\-]?\s*(\d{10})", full_text, re.IGNORECASE
    )
    if contact_match:
        contact_no = contact_match.group(1).strip()
    else:
        num_match = re.search(r"\b[6-9]\d{9}\b", full_text)
        if num_match:
            contact_no = num_match.group(0)

    # 3. Type
    media_type = ""
    type_match = re.search(
        r"Type\s*[:\-]?\s*([A-Za-z0-9_\-]+)", full_text, re.IGNORECASE
    )
    if type_match:
        media_type = type_match.group(1).strip()

    # 4. Size
    size = ""
    size_match = re.search(
        r"Size\s*[:\-]?\s*(\d+\s*x\s*\d+)", full_text, re.IGNORECASE
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
                outlet_name, contact_no, media_type, size = extract_info(slide)

                # Collect all picture shapes
                pic_shapes = [s for s in slide.shapes if s.shape_type == 13]

                if pic_shapes:
                    # Sort by X (left) coordinate to get LEFTMOST image
                    leftmost_pic = min(pic_shapes, key=lambda s: s.left)

                    if not outlet_name:
                        outlet_name = f"Slide_{i+1}"

                    components = [
                        clean_text(p)
                        for p in [outlet_name, contact_no, media_type, size]
                        if p
                    ]
                    base_filename = "_".join(components)

                    ext = leftmost_pic.image.ext
                    final_name = f"{base_filename}.{ext}"

                    # Write image blob directly
                    zip_file.writestr(final_name, leftmost_pic.image.blob)

    st.success("🎉 Sabhi Left Images ready hain!")
    st.download_button(
        label="📥 Download ZIP",
        data=zip_buffer.getvalue(),
        file_name="Renamed_Left_Images.zip",
        mime="application/zip",
    )
