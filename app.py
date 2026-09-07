import io
import math
import re
import zipfile
from pptx import Presentation
import streamlit as st
from PIL import Image, ImageDraw

st.set_page_config(
    page_title="PPT Image & Marking Composite Extractor",
    page_icon="🖼️",
    layout="wide",
)
st.title("🖼️ PPT Image & Overlay Marking Extractor")
st.caption(
    "Detects image along with overlaid Red Box shapes, composites them together, and renames smartly."
)

image_option = st.radio(
    "Select Image to Export:",
    ("Image 1 (Left / Close View)", "Image 2 (Right / Far View)"),
    index=1,
)


def clean_text(text):
    if not text:
        return ""
    clean = re.sub(r'[\\/*?:"<>|\n\r\t]', " ", text)
    clean = re.sub(r"\s+", " ", clean).strip()
    return clean.replace(" ", "_")


def extract_info_from_slide(slide, slide_num):
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
            if getattr(shape, "shape_type", None) == 6:  # Group shape
                collect_text(shape.shapes)

    collect_text(slide.shapes)
    full_text = "\n".join(all_text_blocks)

    outlet_name, contact_no, media_type, size = "", "", "", ""

    # Search Outlet Name explicitly
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

    # Fallback to first non-label text block
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


def composite_image_and_marking(slide, target_pic):
    """Parses PPT XML bounding boxes of overlay red shapes and draws them over the base image"""
    image_bytes = target_pic.image.blob
    pil_img = Image.open(io.BytesIO(image_bytes)).convert("RGBA")
    img_w, img_h = pil_img.size

    pic_left = target_pic.left
    pic_top = target_pic.top
    pic_w = target_pic.width
    pic_h = target_pic.height

    # Overlay layer creation
    overlay_layer = Image.new("RGBA", (img_w, img_h), (255, 255, 255, 0))
    draw = ImageDraw.Draw(overlay_layer)

    # Dynamic line width based on photo resolution
    line_thick = max(4, int(min(img_w, img_h) * 0.012))

    for shape in slide.shapes:
        if shape == target_pic:
            continue

        # Skip slide main containers / text boxes
        if shape.has_text_frame and len(shape.text_frame.text.strip()) > 3:
            continue

        s_left, s_top = shape.left, shape.top
        s_w, s_h = shape.width, shape.height

        # Check if shape bounding box overlaps target image
        if (
            s_left + s_w > pic_left - 50000
            and s_left < pic_left + pic_w + 50000
            and s_top + s_h > pic_top - 50000
            and s_top < pic_top + pic_h + 50000
        ):
            # Calculate pixel positions relative to the picture box
            rel_x1 = (s_left - pic_left) / pic_w
            rel_y1 = (s_top - pic_top) / pic_h
            rel_x2 = (s_left + s_w - pic_left) / pic_w
            rel_y2 = (s_top + s_h - pic_top) / pic_h

            px1 = int(rel_x1 * img_w)
            py1 = int(rel_y1 * img_h)
            px2 = int(rel_x2 * img_w)
            py2 = int(rel_y2 * img_h)

            # Ensure box size is valid
            if abs(px2 - px1) >= 2 and abs(py2 - py1) >= 2:
                for t in range(line_thick):
                    draw.rectangle(
                        [px1 - t, py1 - t, px2 + t, py2 + t],
                        outline=(255, 0, 0, 255),
                    )

    # Composite base image with drawn overlay
    final_img = Image.alpha_composite(pil_img, overlay_layer).convert("RGB")

    out = io.BytesIO()
    final_img.save(out, format="JPEG", quality=95)
    return out.getvalue()


uploaded_file = st.file_uploader("Upload PowerPoint File (.pptx)", type=["pptx"])

if uploaded_file is not None:
    if st.button(
        "▶️ Start Composite Extraction", type="primary", use_container_width=True
    ):
        prs = Presentation(uploaded_file)
        zip_buffer = io.BytesIO()

        total_slides = len(prs.slides)
        progress_bar = st.progress(0)

        with zipfile.ZipFile(
            zip_buffer, "a", zipfile.ZIP_DEFLATED, False
        ) as zip_file:
            for i, slide in enumerate(prs.slides):
                outlet_name, contact_no, media_type, size = (
                    extract_info_from_slide(slide, i + 1)
                )

                # Fetch pictures sorted left-to-right
                pic_shapes = [
                    s
                    for s in slide.shapes
                    if getattr(s, "shape_type", None) == 13
                    or hasattr(s, "image")
                ]
                pic_shapes = sorted(pic_shapes, key=lambda s: s.left)

                if pic_shapes:
                    target_pic = (
                        pic_shapes[-1] if "Image 2" in image_option else pic_shapes[0]
                    )

                    # Extract Image + Red Box Layer together
                    final_bytes = composite_image_and_marking(slide, target_pic)

                    components = []
                    if outlet_name:
                        components.append(clean_text(outlet_name))
                    if contact_no:
                        components.append(clean_text(contact_no))
                    if media_type:
                        components.append(clean_text(media_type))
                    if size:
                        components.append(clean_text(size))

                    # Clean filename (no "Slide_" prefix)
                    if not components:
                        components.append(f"Store_{i+1}")

                    final_name = f"{'_'.join(components)}.jpg"
                    zip_file.writestr(final_name, final_bytes)

                progress_bar.progress((i + 1) / total_slides)

        st.success("🎉 Extraction Completed with Markings Combined!")
        st.download_button(
            label="📥 Download Extracted Images with Markings (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Composited_Marked_Images.zip",
            mime="application/zip",
        )
