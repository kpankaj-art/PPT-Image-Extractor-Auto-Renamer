import io
import re
import zipfile
import fitz  # PyMuPDF
from PIL import Image
import streamlit as st

st.set_page_config(page_title="Far-View Crop & Renamer", layout="centered")
st.title("Far-View (Right Image) Crop & Renamer with Markup")

st.info("💡 **Note:** PPT file ko pehle 'Save As PDF' karein aur yahan PDF upload karein. Isse Red Markup Image ke sath permanently merge ho jata hai.")

def extract_metadata_from_page(page):
    """PDF page ke sabhi text blocks se metadata read karta hai."""
    text_blocks = page.get_text("blocks")
    full_text = " ".join([b[4] for b in text_blocks])
    clean_text = re.sub(r"\s+", " ", full_text)

    outlet_name = "OUTLET"
    contact = "0000000000"
    media_type = "NL"
    size = "0x0"

    # 1. Outlet Name Match
    outlet_match = re.search(
        r"Outlet\s*Name\s*:\s*([^:\n\r]+?)(?=\s*Address|\s*City|\s*Contact|\s*Type|$)",
        clean_text,
        re.IGNORECASE,
    )
    if outlet_match:
        raw_name = outlet_match.group(1).strip()
        cleaned = re.sub(r"[^\w\s-]", "", raw_name)
        outlet_name = re.sub(r"\s+", "_", cleaned).upper().strip("_")

    # 2. Contact Match
    contact_match = re.search(
        r"(?:Contact|Mobile|Phone)?\s*:?\s*([6-9]\d{9})",
        clean_text,
        re.IGNORECASE,
    )
    if contact_match:
        contact = contact_match.group(1).strip()

    # 3. Type Match
    type_match = re.search(
        r"Type\s*:\s*([A-Za-z0-9_-]+)", clean_text, re.IGNORECASE
    )
    if type_match:
        media_type = type_match.group(1).strip().upper()

    # 4. Size Match
    size_match = re.search(
        r"Size\s*:\s*(\d+)\s*[*xX]\s*(\d+)", clean_text, re.IGNORECASE
    )
    if size_match:
        size = f"{size_match.group(1)}x{size_match.group(2)}"

    return f"{outlet_name}_{contact}_{media_type}_{size}"


uploaded_file = st.file_uploader("Apni PDF File Upload Karein", type=["pdf"])

if uploaded_file is not None:
    if st.button("Extract Right Image with Markup"):
        with st.spinner("Processing PDF pages..."):
            pdf_bytes = uploaded_file.read()
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            zip_buffer = io.BytesIO()
            extracted_count = 0

            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for page_idx, page in enumerate(doc):
                    filename_prefix = extract_metadata_from_page(page)

                    # Page ko high DPI rendering (300 DPI)
                    zoom = 3
                    mat = fitz.Matrix(zoom, zoom)
                    pix = page.get_pixmap(matrix=mat)

                    # Visual Image convert karna
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    w, h = img.size

                    # Right Side "Far View" Box Crop Coordinates
                    # (Left: 51.5%, Top: 31.5%, Right: 95.5%, Bottom: 81%)
                    crop_area = (
                        int(w * 0.515),
                        int(h * 0.315),
                        int(w * 0.955),
                        int(h * 0.810),
                    )

                    cropped_img = img.crop(crop_area)

                    img_byte_arr = io.BytesIO()
                    cropped_img.save(img_byte_arr, format="PNG")

                    final_name = f"{filename_prefix}_1.png"
                    zip_file.writestr(final_name, img_byte_arr.getvalue())
                    extracted_count += 1

            doc.close()

            if extracted_count > 0:
                st.success(f"Total {extracted_count} Images (With Markup) successfully extracted & renamed!")
                st.download_button(
                    label="Download Cropped ZIP",
                    data=zip_buffer.getvalue(),
                    file_name="outlet_far_views_with_markup.zip",
                    mime="application/zip",
                )
            else:
                st.warning("PDF process nahi ho saki.")
