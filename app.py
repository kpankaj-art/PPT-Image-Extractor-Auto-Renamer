import io
import os
import re
import zipfile
import subprocess
from pptx import Presentation
import streamlit as st
from PIL import Image
from pdf2image import convert_from_path

st.set_page_config(
    page_title="PPT Visual Image Extractor", page_icon="🖼️", layout="centered"
)
st.title("🖼️ PPT Visual Image Extractor")

image_option = st.radio(
    "Select Image to Export:",
    ("Image 1 (Left / Close View)", "Image 2 (Right / Far View)"),
    index=1
)

def clean_text(text):
    if not text:
        return ""
    clean = re.sub(r'[\\/*?:"<>|\n\r\t]', " ", text)
    clean = re.sub(r"\s+", " ", clean).strip()
    return clean.replace(" ", "_")

def extract_info_from_slide(slide):
    all_text_blocks = []
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            all_text_blocks.append(shape.text_frame.text.strip())
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        all_text_blocks.append(cell.text.strip())

    full_text = "\n".join(all_text_blocks)
    outlet_name, contact_no, media_type, size = "", "", "", ""

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

    contact_match = re.search(r"\b[6-9]\d{9}\b", full_text)
    if contact_match:
        contact_no = contact_match.group(0)

    type_match = re.search(
        r"\b(NL|FL|BL|SB|GSB|Non-Lit|Flex)\b", full_text, re.IGNORECASE
    )
    if type_match:
        media_type = type_match.group(1).upper()

    size_match = re.search(
        r"(\d{1,3}(?:\.\d+)?\s*x\s*\d{1,3}(?:\.\d+)?)", full_text, re.IGNORECASE
    )
    if size_match:
        size = size_match.group(1).replace(" ", "").lower()

    return outlet_name, contact_no, media_type, size

def crop_image_visually_from_slide(slide, target_shape, slide_image, prs):
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    left = target_shape.left
    top = target_shape.top
    width = target_shape.width
    height = target_shape.height

    img_w, img_h = slide_image.size

    x1 = int((left / slide_width) * img_w)
    y1 = int((top / slide_height) * img_h)
    x2 = int(((left + width) / slide_width) * img_w)
    y2 = int(((top + height) / slide_height) * img_h)

    cropped = slide_image.crop((x1, y1, x2, y2))
    
    out = io.BytesIO()
    cropped.save(out, format="JPEG", quality=95)
    return out.getvalue()

uploaded_file = st.file_uploader("Upload PowerPoint File (.pptx)", type=["pptx"])

if uploaded_file is not None:
    if st.button("▶️ Start Visual Extraction", type="primary", use_container_width=True):
        with open("temp_input.pptx", "wb") as f:
            f.write(uploaded_file.getbuffer())

        prs = Presentation("temp_input.pptx")
        zip_buffer = io.BytesIO()

        with st.spinner("Converting PPT slides to capture red markings..."):
            try:
                # Convert PPTX to PDF via LibreOffice (Linux Server Supported)
                subprocess.run(
                    ["libreoffice", "--headless", "--convert-to", "pdf", "temp_input.pptx"],
                    check=True
                )
                rendered_images = convert_from_path("temp_input.pdf", dpi=200)
            except Exception as e:
                st.error(f"Rendering error: {e}")
                rendered_images = []

        if rendered_images:
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for i, slide in enumerate(prs.slides):
                    outlet_name, contact_no, media_type, size = extract_info_from_slide(slide)
                    
                    pic_shapes = [s for s in slide.shapes if getattr(s, "shape_type", None) == 13 or hasattr(s, "image")]
                    pic_shapes = sorted(pic_shapes, key=lambda s: s.left)

                    if pic_shapes and i < len(rendered_images):
                        target_pic = pic_shapes[-1] if "Image 2" in image_option else pic_shapes[0]
                        
                        final_bytes = crop_image_visually_from_slide(
                            slide, target_pic, rendered_images[i], prs
                        )

                        if not outlet_name:
                            outlet_name = f"Slide_{i+1}"

                        components = [clean_text(outlet_name)]
                        if contact_no:
                            components.append(clean_text(contact_no))
                        if media_type:
                            components.append(clean_text(media_type))
                        if size:
                            components.append(clean_text(size))

                        final_name = f"{'_'.join(components)}.jpg"
                        zip_file.writestr(final_name, final_bytes)

            st.success("🎉 Process Complete with Red Overlay Marks!")
            st.download_button(
                label="📥 Download Visually Marked Images (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Marked_Images.zip",
                mime="application/zip",
            )
