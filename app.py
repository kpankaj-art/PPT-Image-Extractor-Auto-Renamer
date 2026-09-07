import io
import os
import re
import subprocess
import tempfile
import zipfile
import streamlit as st
from pdf2image import convert_from_path
from pptx import Presentation
from PIL import Image

st.set_page_config(page_title="PPT Screen-Crop Extractor", page_icon="✂️", layout="wide")
st.title("✂️ PPT Visual Screen-Crop Extractor (Snipping Method)")
st.caption("Captures EXACT visual state of the slide just like Windows Snipping Tool!")

def clean_text(text):
    if not text:
        return ""
    clean = re.sub(r'[\\/*?:"<>|\n\r\t]', " ", text)
    clean = re.sub(r"\s+", " ", clean).strip()
    return clean.replace(" ", "_")

def extract_info_from_slide(slide):
    all_text_blocks = []
    def collect_text(shapes):
        for shape in shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                all_text_blocks.append(shape.text_frame.text.strip())
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.text.strip():
                            all_text_blocks.append(cell.text.strip())
            if getattr(shape, "shape_type", None) == 6:
                collect_text(shape.shapes)

    collect_text(slide.shapes)
    full_text = "\n".join(all_text_blocks)

    outlet_name, contact_no, media_type, size = "", "", "", ""

    outlet_match = re.search(
        r"(?:Outlet\s*Name|Customer\s*Name|Store\s*Name|Shop\s*Name)\s*[:\-]?\s*([^\n\r]+)",
        full_text,
        re.IGNORECASE,
    )
    if outlet_match:
        raw_name = outlet_match.group(1).strip()
        cleaned_name = re.split(
            r"Address|City|Contact|Installation|Type|Size|Qty|No",
            raw_name,
            flags=re.IGNORECASE,
        )[0].strip()
        if cleaned_name and len(cleaned_name) > 2:
            outlet_name = cleaned_name

    if not outlet_name and all_text_blocks:
        for block in all_text_blocks:
            lines = [l.strip() for l in block.split("\n") if l.strip()]
            for line in lines:
                if not re.search(
                    r"City|Contact|Address|Type|Size|Qty|Slide|Close|Far|Type:|Size:",
                    line,
                    re.IGNORECASE,
                ) and len(line) > 2:
                    outlet_name = line
                    break
            if outlet_name:
                break

    contact_match = re.search(r"\b[6-9]\d{9}\b", full_text)
    if contact_match:
        contact_no = contact_match.group(0)

    type_match = re.search(
        r"\b(NL|FL|BL|SB|GSB|Non-Lit|Flex|Glow-Sign)\b", full_text, re.IGNORECASE
    )
    if type_match:
        media_type = type_match.group(1).upper()

    size_match = re.search(
        r"(\d{1,3}(?:\.\d+)?\s*x\s*\d{1,3}(?:\.\d+)?)", full_text, re.IGNORECASE
    )
    if size_match:
        size = size_match.group(1).replace(" ", "").lower()

    return outlet_name, contact_no, media_type, size

uploaded_file = st.file_uploader("Upload PowerPoint File (.pptx)", type=["pptx"])

if uploaded_file is not None:
    if st.button("▶️ Start Snipping Tool Crop", type="primary", use_container_width=True):
        with tempfile.TemporaryDirectory() as tmpdir:
            pptx_path = os.path.join(tmpdir, "input.pptx")
            with open(pptx_path, "wb") as f:
                f.write(uploaded_file.getvalue())

            prs = Presentation(pptx_path)
            slide_width = prs.slide_width
            slide_height = prs.slide_height

            # Convert PPTX to PDF using LibreOffice (for accurate visual rendering)
            st.info("Rendering visual slides...")
            cmd = f"soffice --headless --convert-to pdf {pptx_path} --outdir {tmpdir}"
            subprocess.run(cmd, shell=True, check=True)

            pdf_path = os.path.join(tmpdir, "input.pdf")
            rendered_images = convert_from_path(pdf_path, dpi=200)

            zip_buffer = io.BytesIO()
            total_slides = len(prs.slides)
            progress_bar = st.progress(0)

            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for i, slide in enumerate(prs.slides):
                    outlet_name, contact_no, media_type, size = extract_info_from_slide(slide)

                    # Find target picture coordinates on the slide
                    pic_shapes = [
                        s for s in slide.shapes
                        if (getattr(s, "shape_type", None) == 13 or hasattr(s, "image"))
                        and s.width > 1000000
                    ]
                    pic_shapes = sorted(pic_shapes, key=lambda s: s.left)

                    if pic_shapes and i < len(rendered_images):
                        target_pic = pic_shapes[-1]
                        
                        slide_img = rendered_images[i]
                        img_w, img_h = slide_img.size

                        # Calculate relative crop rectangle for the target photo
                        crop_x1 = int((target_pic.left / slide_width) * img_w)
                        crop_y1 = int((target_pic.top / slide_height) * img_h)
                        crop_x2 = int(((target_pic.left + target_pic.width) / slide_width) * img_w)
                        crop_y2 = int(((target_pic.top + target_pic.height) / slide_height) * img_h)

                        # Crop visually (exact Snipping Tool behavior)
                        cropped_img = slide_img.crop((crop_x1, crop_y1, crop_x2, crop_y2))

                        out_bytes = io.BytesIO()
                        cropped_img.save(out_bytes, format="JPEG", quality=95)

                        components = []
                        if outlet_name:
                            components.append(clean_text(outlet_name))
                        if contact_no:
                            components.append(clean_text(contact_no))
                        if media_type:
                            components.append(clean_text(media_type))
                        if size:
                            components.append(clean_text(size))

                        if not components:
                            components.append(f"Store_{i+1}")

                        final_name = f"{'_'.join(components)}.jpg"
                        zip_file.writestr(final_name, out_bytes.getvalue())

                    progress_bar.progress((i + 1) / total_slides)

            st.success("🎉 Visual Screen Snipping Completed Successfully!")
            st.download_button(
                label="📥 Download Snipped Marked Images (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Snipped_Marked_Images.zip",
                mime="application/zip",
            )
