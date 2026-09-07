import io
import re
import zipfile
from pptx import Presentation
import streamlit as st
from PIL import Image, ImageDraw

st.set_page_config(
    page_title="PPT Visual Image Extractor", page_icon="🖼️", layout="centered"
)
st.title("🖼️ PPT Visual Image Extractor")
st.write("Extracts images with **Red Box Overlay / Markings** rendered directly on them!")

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


def get_red_shapes_over_target(slide, target_pic):
    """Finds all red box shapes overlaid on the target image"""
    pic_left = target_pic.left
    pic_top = target_pic.top
    pic_right = target_pic.left + target_pic.width
    pic_bottom = target_pic.top + target_pic.height

    overlays = []
    for shape in slide.shapes:
        if shape == target_pic:
            continue

        # Check if shape overlaps with target image bounding box
        s_left, s_top = shape.left, shape.top
        s_right, s_bottom = shape.left + shape.width, shape.top + shape.height

        if (s_left >= pic_left - 100000 and s_right <= pic_right + 100000 and
                s_top >= pic_top - 100000 and s_bottom <= pic_bottom + 100000):

            # Relative coordinates percentage wrt target image
            rel_x1 = max(0.0, min(1.0, (s_left - pic_left) / target_pic.width))
            rel_y1 = max(0.0, min(1.0, (s_top - pic_top) / target_pic.height))
            rel_x2 = max(0.0, min(1.0, (s_right - pic_left) / target_pic.width))
            rel_y2 = max(0.0, min(1.0, (s_bottom - pic_top) / target_pic.height))

            overlays.append((rel_x1, rel_y1, rel_x2, rel_y2))

    return overlays


def render_image_with_overlays(target_pic, overlays):
    """Extracts base image and draws red rectangles on top"""
    image_bytes = target_pic.image.blob
    pil_img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    width, height = pil_img.size

    if overlays:
        draw = ImageDraw.Draw(pil_img)
        line_width = max(3, int(min(width, height) * 0.008))  # Dynamic line thickness

        for rel_x1, rel_y1, rel_x2, rel_y2 in overlays:
            x1 = int(rel_x1 * width)
            y1 = int(rel_y1 * height)
            x2 = int(rel_x2 * width)
            y2 = int(rel_y2 * height)

            # Draw Red Outline Box
            for offset in range(line_width):
                draw.rectangle(
                    [x1 - offset, y1 - offset, x2 + offset, y2 + offset],
                    outline="red"
                )

    out = io.BytesIO()
    pil_img.save(out, format="JPEG", quality=95)
    return out.getvalue()


uploaded_file = st.file_uploader("Upload PowerPoint File (.pptx)", type=["pptx"])

if uploaded_file is not None:
    if st.button("▶️ Start Visual Extraction", type="primary", use_container_width=True):
        prs = Presentation(uploaded_file)
        zip_buffer = io.BytesIO()

        with st.spinner("Extracting images and rendering Red Markings..."):
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for i, slide in enumerate(prs.slides):
                    outlet_name, contact_no, media_type, size = extract_info_from_slide(slide)

                    # Find Picture Shapes sorted by left position
                    pic_shapes = [s for s in slide.shapes if getattr(s, "shape_type", None) == 13 or hasattr(s, "image")]
                    pic_shapes = sorted(pic_shapes, key=lambda s: s.left)

                    if pic_shapes:
                        target_pic = pic_shapes[-1] if "Image 2" in image_option else pic_shapes[0]

                        # Find red overlaid shapes on top of this photo
                        overlays = get_red_shapes_over_target(slide, target_pic)

                        # Render final image with red box overlaid
                        final_bytes = render_image_with_overlays(target_pic, overlays)

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

        st.success("🎉 Process Completed Successfully!")
        st.download_button(
            label="📥 Download Visually Marked Images (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Marked_Images.zip",
            mime="application/zip",
        )
