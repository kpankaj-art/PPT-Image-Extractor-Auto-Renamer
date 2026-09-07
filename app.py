import io
import re
import zipfile
from pptx import Presentation
import streamlit as st
from PIL import Image, ImageDraw

st.set_page_config(
    page_title="PPT Visual Marking Image Extractor", page_icon="🖼️", layout="wide"
)
st.title("🖼️ PPT Visual Image Extractor (Marking Fixed)")

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


def draw_markings_on_target_image(slide, target_pic):
    """Guaranteed overlay extraction for all marking box shapes"""
    image_bytes = target_pic.image.blob
    pil_img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    img_w, img_h = pil_img.size

    pic_left = target_pic.left
    pic_top = target_pic.top
    pic_w = target_pic.width
    pic_h = target_pic.height

    draw = ImageDraw.Draw(pil_img)
    line_thickness = max(4, int(min(img_w, img_h) * 0.012))

    for shape in slide.shapes:
        if shape == target_pic:
            continue

        # Ignore shapes with full text labels
        if shape.has_text_frame and len(shape.text_frame.text.strip()) > 2:
            continue

        s_left, s_top = shape.left, shape.top
        s_w, s_h = shape.width, shape.height

        # Ignore main border blue frames & full background frames
        if s_w >= pic_w * 0.85 and s_h >= pic_h * 0.85:
            continue

        # Check if shape lies over target photo region
        if (
            s_left + s_w > pic_left
            and s_left < pic_left + pic_w
            and s_top + s_h > pic_top
            and s_top < pic_top + pic_h
        ):
            # Calculate pixel bounds on image
            rel_x1 = max(0.0, min(1.0, (s_left - pic_left) / pic_w))
            rel_y1 = max(0.0, min(1.0, (s_top - pic_top) / pic_h))
            rel_x2 = max(0.0, min(1.0, (s_left + s_w - pic_left) / pic_w))
            rel_y2 = max(0.0, min(1.0, (s_top + s_h - pic_top) / pic_h))

            px1, py1 = int(rel_x1 * img_w), int(rel_y1 * img_h)
            px2, py2 = int(rel_x2 * img_w), int(rel_y2 * img_h)

            if abs(px2 - px1) > 2 and abs(py2 - py1) > 2:
                # Draw high-visibility Red Box over the exact spot
                for off in range(line_thickness):
                    draw.rectangle(
                        [px1 - off, py1 - off, px2 + off, py2 + off],
                        outline="#FF0000",
                    )

    out = io.BytesIO()
    pil_img.save(out, format="JPEG", quality=95)
    return out.getvalue()


uploaded_file = st.file_uploader("Upload PowerPoint File (.pptx)", type=["pptx"])

if uploaded_file is not None:
    if st.button("▶️ Start Visual Extraction", type="primary", use_container_width=True):
        prs = Presentation(uploaded_file)
        zip_buffer = io.BytesIO()

        total_slides = len(prs.slides)
        progress_bar = st.progress(0)

        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for i, slide in enumerate(prs.slides):
                outlet_name, contact_no, media_type, size = extract_info_from_slide(
                    slide, i + 1
                )

                # Filter real picture shapes (width check eliminates empty blue frames)
                pic_shapes = [
                    s
                    for s in slide.shapes
                    if (getattr(s, "shape_type", None) == 13 or hasattr(s, "image"))
                    and s.width > 1000000
                ]
                pic_shapes = sorted(pic_shapes, key=lambda s: s.left)

                if pic_shapes:
                    target_pic = pic_shapes[-1]

                    # Burn red box marking directly on image
                    final_bytes = draw_markings_on_target_image(slide, target_pic)

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
                    zip_file.writestr(final_name, final_bytes)

                progress_bar.progress((i + 1) / total_slides)

        st.success("🎉 Extraction Finished!")
        st.download_button(
            label="📥 Download Fixed Marked Images (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Final_Marked_Images.zip",
            mime="application/zip",
        )
