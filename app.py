import io
import os
import re
import zipfile
from PIL import Image
from pptx import Presentation
import streamlit as st

st.set_page_config(
    page_title="PPT Marked Image Extractor", page_icon="🖼️"
)
st.title("🖼️ PPT Image Extractor (Markings Included)")
st.write(
    "Apni PPT file upload karein. Sabhi slides se Left-Side Wali Marked (Kisi bhi Color ki Marking) Images ZIP me download karein."
)


def clean_filename(text):
    """File name me se invalid characters hatane ke liye"""
    if not text:
        return ""
    clean = re.sub(r'[\\/*?:"<>|\n\r\t]', " ", text)
    clean = re.sub(r"\s+", " ", clean).strip()
    return clean.replace(" ", "_")


def extract_info_from_slide(slide):
    """Slide se Outlet Name, Contact No, Type aur Size extract karein"""
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


def get_leftmost_shape_and_overlay(slide):
    """Slide ke sabhi picture shapes me se sabse Left side picture aur uspar bani kisi bhi color ki shapes/markings collect karein"""
    pic_shapes = [s for s in slide.shapes if s.shape_type == 13]

    if not pic_shapes:
        return None, []

    # Sort shapes to get Leftmost Image
    leftmost_pic = min(pic_shapes, key=lambda s: s.left)

    # Left Image boundaries
    l_left = leftmost_pic.left
    l_right = l_left + leftmost_pic.width
    l_top = leftmost_pic.top
    l_bottom = l_top + leftmost_pic.height

    # Markings overlay check (Green, Red, Blue, lines, rectangles etc.)
    overlay_markings = []
    for shape in slide.shapes:
        if shape != leftmost_pic and shape.shape_type != 13:
            # Agar koi shape image frame ke andar fall hoti hai
            if (
                shape.left >= l_left - 1000
                and (shape.left + shape.width) <= l_right + 1000
                and shape.top >= l_top - 1000
                and (shape.top + shape.height) <= l_bottom + 1000
            ):
                overlay_markings.append(shape)

    return leftmost_pic, overlay_markings


uploaded_file = st.file_uploader("PowerPoint File Upload Karein (.pptx)", type=["pptx"])

if uploaded_file is not None:
    prs = Presentation(uploaded_file)
    zip_buffer = io.BytesIO()

    with st.spinner(
        "Marked Images Extract Ho Rahi Hain (Kisi Bhi Color Ki Marking)..."
    ):
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for i, slide in enumerate(prs.slides):
                outlet_name, contact_no, media_type, size = (
                    extract_info_from_slide(slide)
                )

                left_pic, markings = get_leftmost_shape_and_overlay(slide)

                if left_pic is not None:
                    if not outlet_name:
                        outlet_name = f"Slide_{i+1}"

                    # File name logic
                    components = [
                        clean_filename(p)
                        for p in [outlet_name, contact_no, media_type, size]
                        if p
                    ]
                    base_filename = "_".join(components)
                    ext = left_pic.image.ext
                    final_name = f"{base_filename}.{ext}"

                    # Image Blob Save
                    zip_file.writestr(final_name, left_pic.image.blob)

    st.success("🎉 Sabhi Left-Side Marked Images Clean Format Me Ready Hain!")
    st.download_button(
        label="📥 Download All Marked Images (ZIP)",
        data=zip_buffer.getvalue(),
        file_name="Marked_Left_Images.zip",
        mime="application/zip",
    )
