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
    """Invalid file name characters hatane ke liye"""
    if not text:
        return ""
    clean = re.sub(r'[\\/*?:"<>|\n\r\t]', " ", text)
    clean = re.sub(r"\s+", " ", clean).strip()
    return clean.replace(" ", "_")


def extract_info_from_slide(slide):
    """Exact text matching for Outlet Name, Contact, Type, and Size"""
    full_text = ""
    for shape in slide.shapes:
        if shape.has_text_frame:
            full_text += "\n" + shape.text_frame.text

    outlet_name = ""
    contact_no = ""
    media_type = ""
    size = ""

    # 1. Outlet Name: 'Outlet Name:' se lekar agli line/label tak ka exact text
    outlet_match = re.search(
        r"Outlet\s*Name\s*[:\-]\s*(.*?)(?=\n|Address|City|Contact|Installation|Type|Size|$)",
        full_text,
        re.IGNORECASE,
    )
    if outlet_match:
        outlet_name = outlet_match.group(1).strip()

    # 2. Contact No: 10 digit Mobile Number
    contact_match = re.search(
        r"Contact\s*(?:No)?\s*[:\-]?\s*(\d{10})", full_text, re.IGNORECASE
    )
    if contact_match:
        contact_no = contact_match.group(1).strip()
    else:
        num_match = re.search(r"\b[6-9]\d{9}\b", full_text)
        if num_match:
            contact_no = num_match.group(0)

    # 3. Type: 'Type:' ke aage ka text (FL, NL, etc.) - Ignore 'Outlet' word
    type_match = re.search(
        r"(?<!Outlet\s)Type\s*[:\-]\s*([A-Za-z0-9_\-]+)", full_text, re.IGNORECASE
    )
    if type_match:
        media_type = type_match.group(1).strip()

    # 4. Size: 'Size:' ke aage ka text (e.g., 10x4 ya 10 x 4)
    size_match = re.search(
        r"Size\s*[:\-]\s*(\d+\s*x\s*\d+)", full_text, re.IGNORECASE
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

                # Slide ke saare picture shapes
                pic_shapes = [s for s in slide.shapes if s.shape_type == 13]

                if pic_shapes:
                    # Horizontal position (left) ke basis par Leftmost Image select karein
                    leftmost_pic = min(pic_shapes, key=lambda s: s.left)

                    # Agar Name na mile toh slide number fallback
                    if not outlet_name:
                        outlet_name = f"Slide_{i+1}"

                    # Exact Order: Name -> Number -> Type -> Size
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

    st.success("🎉 Process Complete!")
    st.download_button(
        label="📥 Download Renamed Images (ZIP)",
        data=zip_buffer.getvalue(),
        file_name="Renamed_Images.zip",
        mime="application/zip",
    )
