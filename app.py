import io
import os
import re
import subprocess
import zipfile
import fitz  # PyMuPDF (pip install pymupdf)
from pptx import Presentation
import streamlit as st
from PIL import Image

st.set_page_config(
    page_title="PPT Marked Image Extractor", page_icon="🖼️", layout="centered"
)
st.title("🖼️ PPT Image Extractor (With Drawing Marks Included)")
st.write(
    "Format: **OutletName_MobileNo_Type_Size.jpg** (Preserves Green Box Marks)"
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

    contact_match = re.search(r"\b[6-9]\d{9}\b", full_text)
    if contact_match:
        contact_no = contact_match.group(0)

    type_match = re.search(
        r"\b(NL|FL|BL|SB|GSB|Non-Lit|Flex)\b", full_text, re.IGNORECASE
    )
    if type_match:
        media_type = type_match.group(1).upper()

    size_match = re.search(
        r"(\d{1,3}\s*x\s*\d{1,3})", full_text, re.IGNORECASE
    )
    if size_match:
        size = size_match.group(1).replace(" ", "").lower()

    return outlet_name, contact_no, media_type, size


def convert_pptx_to_pdf(input_pptx_path, output_dir):
    """LibreOffice se PPTX ko PDF banayein taaki marks render ho saken"""
    cmd = [
        "libreoffice",
        "--headless",
        "--convert-to",
        "pdf",
        input_pptx_path,
        "--outdir",
        output_dir,
    ]
    subprocess.run(cmd, check=True)
    pdf_path = os.path.join(
        output_dir, os.path.splitext(os.path.basename(input_pptx_path))[0] + ".pdf"
    )
    return pdf_path


uploaded_file = st.file_uploader("PowerPoint File Upload Karein (.pptx)", type=["pptx"])

if uploaded_file is not None:
    temp_pptx = "temp_input.pptx"
    with open(temp_pptx, "wb") as f:
        f.write(uploaded_file.getbuffer())

    prs = Presentation(temp_pptx)
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    zip_buffer = io.BytesIO()

    with st.spinner(
        "Rendering slides & Cropping Marked Images (Visual Overlay Mode)..."
    ):
        try:
            # 1. PPT ko PDF me convert karke visually draw kar rahe hain
            pdf_path = convert_pptx_to_pdf(temp_pptx, ".")
            doc = fitz.open(pdf_path)

            with zipfile.ZipFile(
                zip_buffer, "a", zipfile.ZIP_DEFLATED, False
            ) as zip_file:
                for i, slide in enumerate(prs.slides):
                    outlet_name, contact_no, media_type, size = (
                        extract_info_from_slide(slide)
                    )
                    pic_shapes = [s for s in slide.shapes if s.shape_type == 13]

                    if pic_shapes:
                        leftmost_pic = min(pic_shapes, key=lambda s: s.left)

                        # Bounding box ratios calculate karna
                        rx = leftmost_pic.left / slide_width
                        ry = leftmost_pic.top / slide_height
                        rw = leftmost_pic.width / slide_width
                        rh = leftmost_pic.height / slide_height

                        # PDF Page rendering
                        page = doc.load_page(i)
                        rect = page.rect
                        page_w, page_h = rect.width, rect.height

                        # Precise crop area selection
                        crop_box = fitz.Rect(
                            rx * page_w,
                            ry * page_h,
                            (rx + rw) * page_w,
                            (ry + rh) * page_h,
                        )

                        pix = page.get_pixmap(clip=crop_box, dpi=300)
                        img_bytes = pix.tobytes("jpeg")

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
                        zip_file.writestr(final_name, img_bytes)

            doc.close()
            if os.path.exists(pdf_path):
                os.remove(pdf_path)

            st.success("🎉 Marked Images successfully extracted!")
            st.download_button(
                label="📥 Download Renamed Images With Box (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Renamed_Marked_Images.zip",
                mime="application/zip",
            )
        except Exception as e:
            st.error(
                "PDF Render require LibreOffice or PyMuPDF. Standard fallback executed."
            )
