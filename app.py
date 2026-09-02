import io
import os
import re
import zipfile
from pptx import Presentation
import streamlit as st

st.set_page_config(
    page_title="PPT Image Extractor Fixed", page_icon="🖼️", layout="centered"
)
st.title("🖼️ PPT Image Extractor & Renamer (Fixed)")


def clean_text(text):
    if not text:
        return ""
    clean = re.sub(r'[\\/*?:"<>|\n\r\t]', " ", text)
    clean = re.sub(r"\s+", " ", clean).strip()
    return clean.replace(" ", "_")


def extract_info_from_slide(slide):
    all_text_blocks = []

    for shape in slide.shapes:
        if shape.has_text_frame:
            txt = shape.text_frame.text.strip()
            if txt:
                all_text_blocks.append(txt)

    full_text = "\n".join(all_text_blocks)

    outlet_name = ""
    contact_no = ""
    media_type = ""
    size = ""

    # 1. OUTLET NAME EXTRACTION
    outlet_match = re.search(
        r"Outlet\s*Name\s*[:\-]?\s*([^\n\r]+)", full_text, re.IGNORECASE
    )
    if outlet_match:
        raw_name = outlet_match.group(1).strip()
        outlet_name = re.split(
            r"Address|City|Contact|Installation|Type|Size|Qty",
            raw_name,
            flags=re.IGNORECASE,
        )[0].strip()
    else:
        # Fallback for plain header lines if Outlet Name prefix is missing
        for line in all_text_blocks:
            if not any(
                k in line.lower()
                for k in [
                    "qty",
                    "size",
                    "type",
                    "address",
                    "city",
                    "contact",
                    "far view",
                    "close view",
                    "board",
                ]
            ):
                if len(line) > 3:
                    outlet_name = line.split("\n")[0].strip()
                    break

    # 2. CONTACT NUMBER
    contact_match = re.search(
        r"(?:Contact\s*No|Mob|Mobile)?\s*[:\-]?\s*(\d{10})",
        full_text,
        re.IGNORECASE,
    )
    if contact_match:
        contact_no = contact_match.group(1).strip()

    # 3. TYPE (NL, FL, BL, SB, etc.)
    # Handles both 'Type: NL' and 'FLboard_qty...' patterns visible in PPT text frames
    type_match = re.search(
        r"\b(NL|FL|BL|SB|GSB|Non-Lit|Flex)\b", full_text, re.IGNORECASE
    )
    if type_match:
        media_type = type_match.group(1).upper()
    else:
        type_lbl = re.search(
            r"Type\s*[:\-]?\s*([A-Za-z0-9]+)", full_text, re.IGNORECASE
        )
        if type_lbl and type_lbl.group(1).lower() != "outlet":
            media_type = type_lbl.group(1).upper()

    # 4. SIZE EXTRACTION (Captures 10x4, 24x84, 12x84, etc.)
    size_match = re.search(
        r"(\d{1,3}\s*x\s*\d{1,3})", full_text, re.IGNORECASE
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

                # Get images (shape_type == 13)
                pic_shapes = [s for s in slide.shapes if s.shape_type == 13]

                if pic_shapes:
                    # Pick left-most image only
                    leftmost_pic = min(pic_shapes, key=lambda s: s.left)

                    if not outlet_name:
                        outlet_name = f"Outlet_{i+1}"

                    # Construct Filename
                    parts = [clean_text(outlet_name)]
                    if contact_no:
                        parts.append(clean_text(contact_no))
                    if media_type:
                        parts.append(clean_text(media_type))
                    if size:
                        parts.append(clean_text(size))

                    base_filename = "_".join(parts)
                    ext = leftmost_pic.image.ext
                    final_name = f"Slide_{i+1}_{base_filename}.{ext}"

                    zip_file.writestr(final_name, leftmost_pic.image.blob)

    st.success("🎉 Fix Applied & Processing Complete!")
    st.download_button(
        label="📥 Download Corrected ZIP",
        data=zip_buffer.getvalue(),
        file_name="Renamed_Images_Fixed.zip",
        mime="application/zip",
    )
