import io
import re
import zipfile
import fitz  # PyMuPDF
import streamlit as st

st.set_page_config(page_title="PDF Far-View Crop & Renamer", layout="centered")
st.title("PDF Far-View (Right Image) Crop & Renamer")


def extract_metadata_from_text(page):
    """PDF page blocks se accurate text parse karke Rename metadata extract karta hai."""
    # Method 1: Get raw text blocks for maximum layout coverage
    text_blocks = page.get_text("blocks")
    full_text = " ".join([b[4] for b in text_blocks])

    # Spaces aur internal linebreaks clean karna
    clean_text = re.sub(r"\s+", " ", full_text)

    outlet_name = "OUTLET"
    contact = "0000000000"
    media_type = "NL"
    size = "0x0"

    # 1. Outlet Name Parsing (Flexible match up to Address/City/Contact)
    outlet_match = re.search(
        r"Outlet\s*Name\s*:\s*([^:\n\r]+?)(?=\s*Address|\s*City|\s*Contact|\s*Type|$)",
        clean_text,
        re.IGNORECASE,
    )
    if outlet_match:
        raw_name = outlet_match.group(1).strip()
        cleaned = re.sub(r"[^\w\s-]", "", raw_name)
        outlet_name = re.sub(r"\s+", "_", cleaned).upper().strip("_")

    # 2. Contact Number Parsing (10 Digits exact pattern search)
    contact_match = re.search(
        r"(?:Contact|Mobile|Phone)?\s*:?\s*([6-9]\d{9})",
        clean_text,
        re.IGNORECASE,
    )
    if contact_match:
        contact = contact_match.group(1).strip()

    # 3. Type Parsing (NL, SB, GS, Non-Lit, Non Lit, Lit etc.)
    type_match = re.search(
        r"Type\s*:\s*([A-Za-z0-9_-]+)", clean_text, re.IGNORECASE
    )
    if type_match:
        media_type = type_match.group(1).strip().upper()

    # 4. Size Parsing (e.g. 10 x 2, 10x2, 8 * 3)
    size_match = re.search(
        r"Size\s*:\s*(\d+)\s*[*xX]\s*(\d+)", clean_text, re.IGNORECASE
    )
    if size_match:
        size = f"{size_match.group(1)}x{size_match.group(2)}"

    # File Format standard naming rule
    return f"{outlet_name}_{contact}_{media_type}_{size}"


uploaded_file = st.file_uploader("Apni PDF File Upload Karein", type=["pdf"])

if uploaded_file is not None:
    if st.button("Extract Right Image & Rename"):
        with st.spinner("Processing PDF pages..."):
            pdf_bytes = uploaded_file.read()
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            zip_buffer = io.BytesIO()
            extracted_count = 0

            with zipfile.ZipFile(
                zip_buffer, "a", zipfile.ZIP_DEFLATED, False
            ) as zip_file:
                for page_idx, page in enumerate(doc):
                    # Direct page object pass kar rahe hain to scan all blocks
                    filename_prefix = extract_metadata_from_text(page)

                    page_rect = page.rect
                    mid_x = page_rect.width / 2

                    zoom = 3  # High Quality Output (300 DPI)
                    mat = fitz.Matrix(zoom, zoom)

                    image_list = page.get_images(full=True)
                    img_idx = 1

                    for img_info in image_list:
                        xref = img_info[0]
                        rects = page.get_image_rects(xref)

                        for rect in rects:
                            # Screen ke RIGHT part wali image select karna (Far View)
                            if rect.x0 >= mid_x * 0.7:
                                pix = page.get_pixmap(matrix=mat, clip=rect)
                                img_data = pix.tobytes("png")

                                final_name = f"{filename_prefix}_{img_idx}.png"
                                zip_file.writestr(final_name, img_data)

                                img_idx += 1
                                extracted_count += 1

            doc.close()

            if extracted_count > 0:
                st.success(
                    f"Total {extracted_count} Images cropped and renamed properly!"
                )
                st.download_button(
                    label="Download Cropped ZIP",
                    data=zip_buffer.getvalue(),
                    file_name="outlet_far_views.zip",
                    mime="application/zip",
                )
            else:
                st.warning("PDF me Right-side wali image nahi mil saki.")
