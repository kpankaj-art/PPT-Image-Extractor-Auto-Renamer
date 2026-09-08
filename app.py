import io
import re
import zipfile
import gc  # Memory optimization ke liye
import fitz  # PyMuPDF
import streamlit as st
from PIL import Image

st.set_page_config(page_title="PDF Target Image Extractor", page_icon="🖼️", layout="wide")

st.title("🖼️ Exact Right Image Cropper")
st.write("PDF upload karein aur slides ke andar se sirf Right Photo ko crop karein.")

# Settings Sidebar
st.sidebar.header("Crop Controls")
target_pos = st.sidebar.radio("Konsi Photo Crop Karni Hai?", ["Right Image Box", "Left Image Box"])

uploaded_pdf = st.file_uploader("📄 PDF File Upload Karein (.pdf)", type=["pdf"])

def extract_metadata(text):
    """PDF text se details auto-extract karein"""
    outlet_name = ""
    contact_no = ""
    media_type = ""
    size = ""

    contact_match = re.search(r"\b[6-9]\d{9}\b", text)
    if contact_match:
        contact_no = contact_match.group(0)

    outlet_match = re.search(r"Outlet\s*Name\s*[:\-]?\s*([^\n\r]+)", text, re.IGNORECASE)
    if outlet_match:
        raw = outlet_match.group(1).strip()
        cleaned = re.split(r"Address", raw, flags=re.IGNORECASE)[0].strip()
        outlet_name = re.sub(r'[^A-Za-z0-9]+', '_', cleaned).strip('_')

    type_match = re.search(r"\b(NL|FL|BL|SB|GSB|NON-LIT|FLEX)\b", text, re.IGNORECASE)
    if type_match:
        media_type = type_match.group(1).upper()

    size_match = re.search(r"(\d{1,3}\s*x\s*\d{1,3})", text, re.IGNORECASE)
    if size_match:
        size = size_match.group(1).replace(" ", "")

    return outlet_name, contact_no, media_type, size

if uploaded_pdf is not None:
    pdf_bytes = uploaded_pdf.getvalue()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total_pages = len(doc)
    st.success(f"✅ Total Pages/Slides Found: {total_pages}")

    if st.button("🚀 Start Precision Cropping"):
        zip_buffer = io.BytesIO()
        progress_bar = st.progress(0)
        status_text = st.empty()

        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for i in range(total_pages):
                page = doc.load_page(i)
                status_text.text(f"Processing Slide {i+1} of {total_pages}...")
                progress_bar.progress((i + 1) / total_pages)

                text = page.get_text("text")
                outlet_name, contact_no, media_type, size = extract_metadata(text)

                # Naming format
                components = []
                if outlet_name:
                    components.append(outlet_name)
                else:
                    components.append(f"SLIDE_{i+1}")
                if contact_no:
                    components.append(contact_no)
                if media_type:
                    components.append(media_type)
                if size:
                    components.append(size)

                final_name = "_".join(components) + ".jpg"

                # High Resolution Render (Matrix 2.5)
                mat = fitz.Matrix(2.5, 2.5)
                pix = page.get_pixmap(matrix=mat)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                w, h = img.size

                # Exact Photo Bounding Box Coordinates (Percentages relative to page size)
                if target_pos == "Right Image Box":
                    # Exact Crop for Right Image (Far View Box)
                    crop_box = (
                        int(w * 0.505), # Left
                        int(h * 0.380), # Top
                        int(w * 0.730), # Right
                        int(h * 0.770)  # Bottom
                    )
                else:
                    # Crop for Left Image (Close View Box)
                    crop_box = (
                        int(w * 0.260), # Left
                        int(h * 0.380), # Top
                        int(w * 0.485), # Right
                        int(h * 0.770)  # Bottom
                    )

                cropped_img = img.crop(crop_box)

                # Save cropped image to buffer
                img_byte_arr = io.BytesIO()
                cropped_img.save(img_byte_arr, format="JPEG", quality=95)
                zip_file.writestr(final_name, img_byte_arr.getvalue())

                # RAM Cleanup
                del pix
                del img
                del cropped_img
                if i % 10 == 0:
                    gc.collect()

        doc.close()
        gc.collect()

        status_text.text("Processing Complete!")
        st.success(f"🎉 Successfully cropped all {total_pages} right images!")
        st.download_button(
            label="📥 Download Cropped Images (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Cropped_Right_Images.zip",
            mime="application/zip",
        )
