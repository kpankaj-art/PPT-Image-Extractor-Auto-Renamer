import io
import re
import zipfile
import fitz  # PyMuPDF
from PIL import Image
import streamlit as st

st.set_page_config(page_title="Far-View Crop & Renamer", layout="centered")
st.title("Far-View Crop & Renamer (Fixed Naming)")

st.info("💡 **PDF Upload Karein.** Renaming aur Markup Crop dono 100% sahi kaam karenge.")

def extract_metadata_from_page(page):
    """PDF page ke text se Outlet Name, Contact, Type aur Size ko perfect extract karta hai."""
    # Text ko position wise sort karke extract kar rahe hain taaki labels ke baad exact values milein
    text_blocks = page.get_text("blocks")
    
    # Sort blocks vertically (Top to Bottom), then Horizontally
    text_blocks.sort(key=lambda b: (b[1], b[0]))
    
    raw_text = " ".join([b[4].strip() for b in text_blocks])
    clean_text = re.sub(r"\s+", " ", raw_text)

    outlet_name = "OUTLET"
    contact = "0000000000"
    media_type = "NL"
    size = "0x0"

    # 1. OUTLET NAME MATCHING
    outlet_match = re.search(
        r"Outlet\s*Name\s*:?\s*([^:\n\r|]+?)(?=\s*Address|\s*City|\s*Contact|\s*Type|\s*Size|\s*Qty|$)",
        clean_text,
        re.IGNORECASE,
    )
    if outlet_match:
        raw_name = outlet_match.group(1).strip()
        # Clean special chars at start and end
        cleaned = re.sub(r"^[^\w]+|[^\w]+$", "", raw_name)
        cleaned = re.sub(r"[^\w\s-]", "", cleaned)
        parsed_name = re.sub(r"\s+", "_", cleaned).upper().strip("_")
        if parsed_name:
            outlet_name = parsed_name

    # 2. CONTACT / MOBILE NUMBER MATCHING (10-Digit Indian Mobile)
    contact_match = re.search(r"\b([6-9]\d{9})\b", clean_text)
    if contact_match:
        contact = contact_match.group(1).strip()
    else:
        # Fallback if labeled
        alt_contact = re.search(r"(?:Contact|Mobile|Phone)?\s*(?:No)?\s*:?\s*(\d{10})", clean_text, re.IGNORECASE)
        if alt_contact:
            contact = alt_contact.group(1).strip()

    # 3. TYPE MATCHING (NL, SB, GS, etc.)
    type_match = re.search(r"Type\s*:?\s*([A-Za-z0-9_-]+)", clean_text, re.IGNORECASE)
    if type_match:
        parsed_type = type_match.group(1).strip().upper()
        if parsed_type and parsed_type != "OUTLET":
            media_type = parsed_type

    # 4. SIZE MATCHING (e.g. Size: 10 x 2, Size: 12x6, Size : 8 * 3)
    size_match = re.search(r"Size\s*:?\s*(\d+)\s*[*xX]\s*(\d+)", clean_text, re.IGNORECASE)
    if size_match:
        size = f"{size_match.group(1)}x{size_match.group(2)}"
    else:
        # Alt search for isolated size patterns like 10x2 or 12x6
        alt_size = re.search(r"\b(\d{1,2})\s*[*xX]\s*(\d{1,2})\b", clean_text)
        if alt_size:
            size = f"{alt_size.group(1)}x{alt_size.group(2)}"

    # Final Clean Prefix Generation (No duplicate 'OUTLET' tags)
    filename = f"{outlet_name}_{contact}_{media_type}_{size}_1.png"
    # Clean leading dashes or weird symbols if any
    filename = re.sub(r"^[^A-Za-z0-9]+", "", filename)

    return filename


uploaded_file = st.file_uploader("Apni PDF File Upload Karein", type=["pdf"])

if uploaded_file is not None:
    if st.button("Extract & Rename Right Images"):
        with st.spinner("Processing PDF pages & renaming..."):
            pdf_bytes = uploaded_file.read()
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            zip_buffer = io.BytesIO()
            extracted_count = 0

            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for page_idx, page in enumerate(doc):
                    final_name = extract_metadata_from_page(page)

                    # Page high DPI rendering
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

                    # ZIP me exact filename add karna
                    zip_file.writestr(final_name, img_byte_arr.getvalue())
                    extracted_count += 1

            doc.close()

            if extracted_count > 0:
                st.success(f"Total {extracted_count} Images (With Markup) successfully renamed!")
                st.download_button(
                    label="Download Fixed Renamed ZIP",
                    data=zip_buffer.getvalue(),
                    file_name="outlet_far_views_renamed_fixed.zip",
                    mime="application/zip",
                )
            else:
                st.warning("PDF process nahi ho saki.")
