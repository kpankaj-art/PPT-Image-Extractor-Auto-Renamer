import io
import os
import re
import zipfile
import fitz  # PyMuPDF
import aspose.slides as slides
from pptx import Presentation
import streamlit as st

st.set_page_config(
    page_title="PPT Image Extractor with Marking", page_icon="🖼️", layout="wide"
)
st.title("🖼️ PPT Image Extractor (Accurate Red Box via PDF)")
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


def crop_image_from_pdf_page(pdf_doc, slide_index, pic_shape, prs_width, prs_height):
    """Crops exact image bounding box from converted PDF slide page"""
    page = pdf_doc.load_page(slide_index)
    rect = page.rect
    pdf_w, pdf_h = rect.width, rect.height

    # Calculate exact relative crop coordinates
    left_ratio = pic_shape.left / prs_width
    top_ratio = pic_shape.top / prs_height
    width_ratio = pic_shape.width / prs_width
    height_ratio = pic_shape.height / prs_height

    crop_x1 = left_ratio * pdf_w
    crop_y1 = top_ratio * pdf_h
    crop_x2 = (left_ratio + width_ratio) * pdf_w
    crop_y2 = (top_ratio + height_ratio) * pdf_h

    crop_rect = fitz.Rect(crop_x1, crop_y1, crop_x2, crop_y2)

    # High quality render (3x zoom = ~300 DPI)
    zoom = 3.0
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, clip=crop_rect)

    img_data = pix.tobytes("jpeg")
    return img_data


if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    prs = Presentation(io.BytesIO(file_bytes))
    total_slides = len(prs.slides)
    st.sidebar.success(f"Total Slides: {total_slides}")

    if st.button("🚀 Start Extraction & Merge Red Marking"):
        zip_buffer = io.BytesIO()
        processed_count = 0

        status_text = st.empty()
        status_text.text("Step 1/2: Converting PPT to PDF for perfect Red Box alignment...")

        # Save temporary PPTX file
        temp_pptx = "temp_input.pptx"
        temp_pdf = "temp_output.pdf"
        
        with open(temp_pptx, "wb") as f:
            f.write(file_bytes)

        # Convert PPT to PDF using Aspose
        pres = slides.Presentation(temp_pptx)
        pres.save(temp_pdf, slides.export.SaveFormat.PDF)
        pres.dispose()

        # Load PDF using PyMuPDF
        pdf_doc = fitz.open(temp_pdf)

        progress_bar = st.progress(0)

        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for i, slide in enumerate(prs.slides):
                status_text.text(f"Step 2/2: Extracting Slide {i+1} of {total_slides}...")
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

                        final_image_data = crop_image_from_pdf_page(
                            pdf_doc, i, pic, prs.slide_width, prs.slide_height
                        )
                        zip_file.writestr(final_name, final_image_data)
                        processed_count += 1

        pdf_doc.close()

        # Clean temp files
        if os.path.exists(temp_pptx):
            os.remove(temp_pptx)
        if os.path.exists(temp_pdf):
            os.remove(temp_pdf)

        status_text.text("Processing Complete!")
        st.success(f"🎉 Success! Extracted {processed_count} images with 100% accurate Red Box.")
        st.download_button(
            label="📥 Download Renamed Images (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Renamed_Marked_Images.zip",
            mime="application/zip",
        )
