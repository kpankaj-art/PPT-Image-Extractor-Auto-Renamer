import io
import re
import zipfile
import fitz  # PyMuPDF
import streamlit as st
from PIL import Image

st.set_page_config(page_title="PDF to Image Converter & Cropper", page_icon="🖼️", layout="wide")

st.title("🖼️ PDF Slide to Image Extractor")
st.write("PDF upload karein aur slides/images ko high quality JPG format me crop/extract karein.")

# Sidebar controls
st.sidebar.header("Settings")
zoom_factor = st.sidebar.slider("Image Quality / Resolution", min_value=1.0, max_value=4.0, value=2.0, step=0.5)
crop_mode = st.sidebar.radio("Crop Mode", ["Full Page/Slide", "Left Half (Left Image)", "Right Half (Right Image)"])

uploaded_pdf = st.file_uploader("📄 PDF File Upload Karein (.pdf)", type=["pdf"])

def extract_metadata(text):
    """PDF text se Outlet Name, Mobile, Type, Size auto-extract karein"""
    outlet_name = ""
    contact_no = ""
    media_type = ""
    size = ""

    # Mobile Number (10 digits starting with 6-9)
    contact_match = re.search(r"\b[6-9]\d{9}\b", text)
    if contact_match:
        contact_no = contact_match.group(0)

    # Outlet Name
    outlet_match = re.search(r"Outlet\s*Name\s*[:\-]?\s*([^\n\r]+)", text, re.IGNORECASE)
    if outlet_match:
        outlet_name = outlet_match.group(1).strip().replace(" ", "_")

    # Type
    type_match = re.search(r"\b(NL|FL|BL|SB|GSB|NON-LIT|FLEX)\b", text, re.IGNORECASE)
    if type_match:
        media_type = type_match.group(1).upper()

    # Size
    size_match = re.search(r"(\d{1,3}\s*x\s*\d{1,3})", text, re.IGNORECASE)
    if size_match:
        size = size_match.group(1).replace(" ", "")

    return outlet_name, contact_no, media_type, size

if uploaded_pdf is not None:
    pdf_bytes = uploaded_pdf.getvalue()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total_pages = len(doc)
    st.success(f"✅ Total Pages/Slides Found: {total_pages}")

    if st.button("🚀 Process & Download Images"):
        zip_buffer = io.BytesIO()
        progress_bar = st.progress(0)
        status_text = st.empty()

        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for i, page in enumerate(doc):
                status_text.text(f"Processing Page {i+1} of {total_pages}...")
                progress_bar.progress((i + 1) / total_pages)

                # Extract Text for Auto-naming
                text = page.get_text("text")
                outlet_name, contact_no, media_type, size = extract_metadata(text)

                # Build Filename
                components = [f"Slide_{i+1}"]
                if outlet_name:
                    components.append(outlet_name)
                if contact_no:
                    components.append(contact_no)
                if media_type:
                    components.append(media_type)
                if size:
                    components.append(size)

                base_name = "_".join(components)

                # High Quality Pixmap Rendering (PDF page to Image)
                mat = fitz.Matrix(zoom_factor, zoom_factor)
                pix = page.get_pixmap(matrix=mat)
                
                # Convert PyMuPDF Pixmap to PIL Image for cropping
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                w, h = img.size

                # Crop selection
                if crop_mode == "Left Half (Left Image)":
                    img = img.crop((0, 0, w // 2, h))
                elif crop_mode == "Right Half (Right Image)":
                    img = img.crop((w // 2, 0, w, h))

                img_byte_arr = io.BytesIO()
                img.save(img_byte_arr, format="JPEG", quality=95)
                
                final_filename = f"{base_name}.jpg"
                zip_file.writestr(final_filename, img_byte_arr.getvalue())

        status_text.text("Processing Complete!")
        st.success(f"🎉 Successfully converted {total_pages} pages into clean JPG images!")
        st.download_button(
            label="📥 Download All Images (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Converted_PDF_Images.zip",
            mime="application/zip",
        )
