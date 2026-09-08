import io
import re
import zipfile
import gc
import fitz  # PyMuPDF
import streamlit as st
from PIL import Image

st.set_page_config(page_title="PDF Precision Cropper", page_icon="🖼️", layout="wide")

st.title("🖼️ Exact Right Image Auto-Cropper (Full PDF Support)")
st.write("Upload PDF to extract exact cropped images without missing any slide.")

st.sidebar.header("Crop Controls")
target_pos = st.sidebar.radio("Konsi Photo Crop Karni Hai?", ["Right Image", "Left Image"])

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

    if st.button("🚀 Process ALL Slides"):
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

                # High Quality Page Pixmap Render
                mat = fitz.Matrix(2.0, 2.0)
                pix = page.get_pixmap(matrix=mat)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                # Slide ke embedded images ki exact positions detect karna
                images_info = page.get_image_info()
                
                cropped_img = None
                
                if len(images_info) >= 2:
                    # Sort images Left to Right
                    sorted_imgs = sorted(images_info, key=lambda x: x['bbox'][0])
                    
                    selected_info = sorted_imgs[1] if target_pos == "Right Image" else sorted_imgs[0]
                    bbox = selected_info['bbox']
                    
                    # Convert bounding box to scaled pixels
                    x0 = int(bbox[0] * 2.0)
                    y0 = int(bbox[1] * 2.0)
                    x1 = int(bbox[2] * 2.0)
                    y1 = int(bbox[3] * 2.0)

                    cropped_img = img.crop((x0, y0, x1, y1))

                # Dynamic fallback agar embedded bbox read na ho paye
                if cropped_img is None or cropped_img.width < 50:
                    w, h = img.size
                    if target_pos == "Right Image":
                        crop_box = (int(w * 0.50), int(h * 0.35), int(w * 0.74), int(h * 0.78))
                    else:
                        crop_box = (int(w * 0.25), int(h * 0.35), int(w * 0.49), int(h * 0.78))
                    cropped_img = img.crop(crop_box)

                img_byte_arr = io.BytesIO()
                cropped_img.save(img_byte_arr, format="JPEG", quality=90)
                zip_file.writestr(final_name, img_byte_arr.getvalue())

                del pix
                del img
                del cropped_img
                if i % 10 == 0:
                    gc.collect()

        doc.close()
        gc.collect()

        status_text.text("Processing Complete!")
        st.success(f"🎉 Successfully processed ALL {total_pages} slides!")
        st.download_button(
            label="📥 Download Cropped Images (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Cropped_Right_Images_All.zip",
            mime="application/zip",
        )
