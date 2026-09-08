import io
import re
import zipfile
from pptx import Presentation
import streamlit as st
from PIL import Image, ImageDraw

st.set_page_config(
    page_title="PPT Image Extractor with Marking", page_icon="🖼️", layout="wide"
)
st.title("🖼️ PPT Image Extractor (Red Box Marking)")
st.write("Format: **OutletName_MobileNo_Type_Size.jpg**")

image_position = st.sidebar.radio(
    "Konsi Image Extract karni hai?", ["Left Image", "Right Image", "Dono (Both)"]
)

uploaded_file = st.sidebar.file_uploader(
    "PowerPoint File Upload Karein (.pptx)", type=["pptx"]
)


def clean_text(text):
    if not text:
        return ""
    clean = re.sub(r'[^A-Za-z0-9]+', ' ', text).strip()
    return clean.replace(" ", "_").upper()


def extract_info_from_slide(slide):
    all_text_blocks = []

    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                txt = paragraph.text.strip()
                if txt:
                    all_text_blocks.append(txt)
        elif shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    if cell_text:
                        all_text_blocks.append(cell_text)

    full_text = " \n ".join(all_text_blocks)

    outlet_name = ""
    contact_no = ""
    media_type = ""
    size = ""

    # Outlet Name
    outlet_match = re.search(r"Outlet\s*Name\s*[:\-]?\s*([^\n\r]+)", full_text, re.IGNORECASE)
    if outlet_match:
        raw_name = outlet_match.group(1).strip()
        cleaned_name = re.split(r"Address", raw_name, flags=re.IGNORECASE)[0].strip()
        if cleaned_name:
            outlet_name = clean_text(cleaned_name)

    if not outlet_name:
        ignore_keywords = [
            "qty", "size", "type", "address", "city", "contact",
            "far view", "close view", "board", "installation", "outlet"
        ]
        for block in all_text_blocks:
            lines = [l.strip() for l in block.split("\n") if l.strip()]
            for line in lines:
                if not any(k in line.lower() for k in ignore_keywords):
                    if len(line) > 2 and not line.isdigit():
                        outlet_name = clean_text(line)
                        break
            if outlet_name:
                break

    # Contact No
    contact_match = re.search(r"\b[6-9]\d{9}\b", full_text)
    if contact_match:
        contact_no = contact_match.group(0)

    # Type
    type_match = re.search(r"Type\s*[:\-]?\s*([A-Za-z0-9]+)", full_text, re.IGNORECASE)
    if type_match:
        media_type = type_match.group(1).upper()
    else:
        gen_type = re.search(r"\b(NL|FL|BL|SB|GSB|NON-LIT|FLEX)\b", full_text, re.IGNORECASE)
        if gen_type:
            media_type = gen_type.group(1).upper()

    # Size
    size_match = re.search(r"Size\s*[:\-]?\s*(\d{1,3})\s*x\s*(\d{1,3})", full_text, re.IGNORECASE)
    if size_match:
        size = f"{size_match.group(1)}x{size_match.group(2)}"
    else:
        gen_size = re.search(r"(\d{1,3})\s*x\s*(\d{1,3})", full_text, re.IGNORECASE)
        if gen_size:
            size = f"{gen_size.group(1)}x{gen_size.group(2)}"

    return outlet_name, contact_no, media_type, size


def process_image_with_strict_marking(pic_shape, slide):
    """Strict Overlay Bounding - No Random Extra Lines"""
    image_bytes = io.BytesIO(pic_shape.image.blob)
    pil_img = Image.open(image_bytes).convert("RGB")
    draw = ImageDraw.Draw(pil_img)

    pic_left = pic_shape.left
    pic_top = pic_shape.top
    pic_width = pic_shape.width
    pic_height = pic_shape.height

    img_w, img_h = pil_img.size

    for shape in slide.shapes:
        # Check if shape is an overlay rectangle/box placed specifically ON this image
        if shape != pic_shape and shape.shape_type != 13:
            s_left = shape.left
            s_top = shape.top
            s_right = shape.left + shape.width
            s_bottom = shape.top + shape.height

            p_right = pic_left + pic_width
            p_bottom = pic_top + pic_height

            # Strict overlap filter (Ignores outside borders/lines)
            overlap_left = max(s_left, pic_left)
            overlap_top = max(s_top, pic_top)
            overlap_right = min(s_right, p_right)
            overlap_bottom = min(s_bottom, p_bottom)

            # Only consider if shape covers meaningful area inside image
            if overlap_right > overlap_left and overlap_bottom > overlap_top:
                rel_x1 = int(((overlap_left - pic_left) / pic_width) * img_w)
                rel_y1 = int(((overlap_top - pic_top) / pic_height) * img_h)
                rel_x2 = int(((overlap_right - pic_left) / pic_width) * img_w)
                rel_y2 = int(((overlap_bottom - pic_top) / pic_height) * img_h)

                # Ensure dimensions are valid inside image bounds
                if (rel_x2 - rel_x1) > 10 and (rel_y2 - rel_y1) > 10:
                    line_thickness = max(6, int(img_w / 90))
                    
                    # Lock borders strictly within image dimensions
                    rel_x1 = max(line_thickness // 2, rel_x1)
                    rel_y1 = max(line_thickness // 2, rel_y1)
                    rel_x2 = min(img_w - line_thickness // 2, rel_x2)
                    rel_y2 = min(img_h - line_thickness // 2, rel_y2)

                    draw.rectangle(
                        [rel_x1, rel_y1, rel_x2, rel_y2],
                        outline="red",
                        width=line_thickness,
                    )

    out_bytes = io.BytesIO()
    pil_img.save(out_bytes, format="JPEG", quality=95)
    return out_bytes.getvalue()


if uploaded_file is not None:
    prs = Presentation(uploaded_file)
    total_slides = len(prs.slides)
    st.sidebar.success(f"Total Slides: {total_slides}")

    if st.button("🚀 Start Extraction & Merge Red Marking"):
        zip_buffer = io.BytesIO()
        processed_count = 0

        progress_bar = st.progress(0)
        status_text = st.empty()

        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for i, slide in enumerate(prs.slides):
                status_text.text(f"Processing Slide {i+1} of {total_slides}...")
                progress_bar.progress((i + 1) / total_slides)

                outlet_name, contact_no, media_type, size = extract_info_from_slide(slide)
                pic_shapes = [s for s in slide.shapes if s.shape_type == 13]

                if pic_shapes:
                    pic_shapes.sort(key=lambda s: s.left)

                    if not outlet_name:
                        outlet_name = f"SLIDE_{i+1}"

                    targets = []
                    if image_position == "Left Image" and len(pic_shapes) >= 1:
                        targets.append(("", pic_shapes[0]))
                    elif image_position == "Right Image":
                        right_pic = pic_shapes[1] if len(pic_shapes) >= 2 else pic_shapes[0]
                        targets.append(("", right_pic))
                    elif image_position == "Dono (Both)":
                        if len(pic_shapes) >= 1:
                            targets.append(("_LEFT", pic_shapes[0]))
                        if len(pic_shapes) >= 2:
                            targets.append(("_RIGHT", pic_shapes[1]))

                    for suffix, pic in targets:
                        components = [outlet_name]
                        if contact_no:
                            components.append(contact_no)
                        if media_type:
                            components.append(media_type)
                        if size:
                            components.append(size)

                        base_filename = "_".join(components) + suffix
                        final_name = f"{base_filename}.jpg"

                        final_image_data = process_image_with_strict_marking(pic, slide)
                        zip_file.writestr(final_name, final_image_data)
                        processed_count += 1

        status_text.text("Processing Complete!")
        st.success(f"🎉 Success! Extracted {processed_count} images.")
        st.download_button(
            label="📥 Download Renamed Images (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Renamed_Marked_Images.zip",
            mime="application/zip",
        )
