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
    """File name safe characters remove/cleaner"""
    if not text:
        return ""
    # Space and special characters to underscore
    clean = re.sub(r'[\\/*?:"<>|\n\r\t]', " ", text)
    clean = re.sub(r"\s+", " ", clean).strip()
    return clean.replace(" ", "_")


def extract_info_from_slide(slide):
    all_text_blocks = []

    # 1. Shapes and Text Frames se Text Extract karein
    for shape in slide.shapes:
        if shape.has_text_frame:
            txt = shape.text_frame.text.strip()
            if txt:
                all_text_blocks.append(txt)

        # 2. Table Shapes ke andar se Text Extract karein (Ye aapki PPT ke liye zaroori hai)
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    if cell_text:
                        all_text_blocks.append(cell_text)

    full_text = "\n".join(all_text_blocks)

    outlet_name = ""
    contact_no = ""
    media_type = ""
    size = ""

    # --- OUTLET NAME EXTRACTION ---
    # Pattern 1: Outlet Name: XYZ
    outlet_match = re.search(
        r"Outlet\s*Name\s*[:\-]?\s*([^\n\r]+)", full_text, re.IGNORECASE
    )
    if outlet_match:
        raw_name = outlet_match.group(1).strip()
        cleaned_name = re.split(
            r"Address|City|Contact|Installation|Type|Size|Qty",
            raw_name,
            flags=re.IGNORECASE,
        )[0].strip()
        if cleaned_name:
            outlet_name = cleaned_name

    # Pattern 2: Agar label bina direct text block me Outlet Name likha ho
    if not outlet_name:
        ignore_keywords = [
            "qty",
            "size",
            "type",
            "address",
            "city",
            "contact",
            "far view",
            "close view",
            "board",
            "installation",
            "dealer_code",
            "outlet",
        ]
        for block in all_text_blocks:
            lines = [l.strip() for l in block.split("\n") if l.strip()]
            for line in lines:
                if not any(k in line.lower() for k in ignore_keywords):
                    if len(line) > 2 and not line.isdigit():
                        outlet_name = line
                        break
            if outlet_name:
                break

    # --- CONTACT NUMBER EXTRACTION ---
    contact_match = re.search(r"\b[6-9]\d{9}\b", full_text)
    if contact_match:
        contact_no = contact_match.group(0)

    # --- TYPE EXTRACTION (NL, FL, BL, SB, etc.) ---
    type_match = re.search(
        r"\b(NL|FL|BL|SB|GSB|Non-Lit|Flex)\b", full_text, re.IGNORECASE
    )
    if type_match:
        media_type = type_match.group(1).upper()

    # --- SIZE EXTRACTION (e.g. 10x4, 8x3, 24x84) ---
    size_match = re.search(
        r"(\d{1,3}\s*x\s*\d{1,3})", full_text, re.IGNORECASE
    )
    if size_match:
        size = size_match.group(1).replace(" ", "").lower()

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

                # Slide ki sirf images (Picture shapes) find karein
                pic_shapes = [s for s in slide.shapes if s.shape_type == 13]

                if pic_shapes:
                    # Strictly left side image extraction
                    leftmost_pic = min(pic_shapes, key=lambda s: s.left)

                    # Agar name abhi bhi nahi mil pata tabhi simple Fallback name lagega
                    if not outlet_name:
                        outlet_name = f"Slide_{i+1}"

                    # Filename structure formation
                    components = [clean_text(outlet_name)]
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
