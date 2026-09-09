import io
import re
import zipfile
import fitz  # PyMuPDF
from PIL import Image
import streamlit as st

st.set_page_config(page_title="Far-View Crop & Renamer", layout="centered")
st.title("Far-View Image Crop & Renamer (With Markup)")

st.info("💡 **Note:** PDF Upload karein. Markup automatic crop hoga aur sahi naam se save hoga.")

def extract_metadata_from_page(page):
    """PDF page ke text se Outlet Name, Contact, Type aur Size extract karta hai."""
    text_blocks = page.get_text("blocks")
    # Clean text extraction with line breaks preserved for exact field matching
    lines = []
    for b in text_blocks:
        lines.extend([line.strip() for line in b[4].split('\n') if line.strip()])
    
    full_text = " ".join(lines)
    clean_text = re.sub(r"\s+", " ", full_text)

    outlet_name = "OUTLET"
    contact = "0000000000"
    media_type = "NL"
    size = "0x0"

    # 1. Outlet Name Match (Multiple patterns for safety)
    outlet_match = re.search(
        r"Outlet\s*Name\s*:?\s*([^:\n\r|]+?)(?=\s*Address|\s*City|\s*Contact|\s*Type|\s*Size|$)",
        clean_text,
        re.IGNORECASE,
    )
    if outlet_match:
        raw_name = outlet_match.group(1).strip()
        # Non-alphanumeric characters hatana
        cleaned = re.sub(r"[^\w\s-]", "", raw_name)
        outlet_name = re.sub(r"\s+", "_", cleaned).upper().strip("_")

    # 2. Contact Match (10-Digit Indian Mobile Number)
    contact_match = re.search(r"\b([6-9]\d{9})\b", clean_text)
    if contact_match:
        contact = contact_match.group(1).strip()
    else:
        # Fallback search if 'Contact No:' label is present
        alt_contact = re.search(r"Contact\s*(?:No)?\s*:?\s*(\d+)", clean_text, re.IGNORECASE)
        if alt_contact:
            contact = alt_contact.group(1).strip()

    # 3. Type Match (NL, SB, GS, etc.)
    type_match = re.search(r"Type\s*:\s*([A-Za-z0-9_-]+)", clean_text, re.IGNORECASE)
    if type_match:
        media_type = type_match.group(1).strip().upper()

    # 4. Size Match (e.g. 12 x 6, 10x2, 12*6)
    size_match = re.search(r"Size\s*:\s*(\d+)\s*[*xX]\s*(\d+)", clean_text, re.IGNORECASE)
    if size_match:
        size = f"{size_match.group(1)}x{size_match.group(2)}"

    # Default fallback protection agar Name empty reh jaye
    if not outlet_name or outlet_name == "":
        outlet_name = "OUTLET"

    return f"{outlet_name}_{contact}_{media_type}_{size}"


uploaded_file = st.file_uploader("Apni PDF File Upload Karein", type=["pdf"])

if uploaded_file is not None:
    if st.button("Extract & Rename Right Images"):
        with st.spinner("Processing PDF pages..."):
            pdf_bytes = uploaded_file.read()
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            zip_buffer = io.BytesIO()
            extracted_count = 0

            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for page_idx, page in enumerate(doc):
                    filename_prefix = extract_metadata_from_page(page)

                    # Page high DPI rendering (300 DPI equivalent)
                    zoom = 3
                    mat = fitz.Matrix(zoom, zoom)
                    pix = page.get_pixmap(matrix=mat)

                    # Visual Image conversion
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    w, h = img.size

                    # Right Side "Far View" Box Crop Area Coordinates
                    crop_area = (
                        int(w * 0.515),  # Left
                        int(h * 0.315),  # Top
                        int(w * 0.955),  # Right
                        int(h * 0.810),  # Bottom
                    )

                    cropped_img = img.crop(crop_area)

                    img_byte_arr = io.BytesIO()
                    cropped_img.save(img_byte_arr, format="PNG")

                    # Final filename formatting
                    final_name = f"{filename_prefix}_1.png"
                    zip_file.writestr(final_name, img_byte_arr.getvalue())
                    extracted_count += 1

            doc.close()

            if extracted_count > 0:
                st.success(f"Total {extracted_count} Images with Markup successfully cropped & renamed!")
                st.download_button(
                    label="Download Renamed Images ZIP",
                    data=zip_buffer.getvalue(),
                    file_name="outlet_far_views_renamed.zip",
                    mime="application/zip",
                )
            else:
                st.warning("PDF process nahi ho saki.")
