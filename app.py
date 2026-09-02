import io
import os
import re
import zipfile
from pptx import Presentation
import streamlit as st

st.set_page_config(page_title="PPT Image Extractor", page_icon="🖼️")
st.title("🖼️ PPT Image Extractor & Auto-Renamer")
st.write("Apni PowerPoint file upload karein aur automatic renamed images download karein.")


def sanitize(text):
    """File name me invalid characters remove karne ke liye"""
    clean_text = re.sub(r'[\\/*?:"<>|]', "", text).strip()
    return clean_text.replace(" ", "_")


uploaded_file = st.file_uploader("PowerPoint File Upload Karein (.pptx)", type=["pptx"])

if uploaded_file is not None:
    prs = Presentation(uploaded_file)
    zip_buffer = io.BytesIO()

    with st.spinner("Images process ho rahi hain..."):
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for i, slide in enumerate(prs.slides):
                outlet_name = ""
                contact_no = ""
                media_type = ""
                size = ""
                images = []

                # Slide ke shapes scan karein
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        text = shape.text_frame.text
                        for line in text.split("\n"):
                            clean_line = line.strip()

                            if "Outlet Name:" in clean_line:
                                outlet_name = clean_line.replace(
                                    "Outlet Name:", ""
                                ).strip()
                            elif "Contact No:" in clean_line:
                                contact_no = clean_line.replace(
                                    "Contact No:", ""
                                ).strip()
                            elif "Type:" in clean_line:
                                media_type = clean_line.replace(
                                    "Type:", ""
                                ).strip()
                            elif "Size:" in clean_line:
                                size = clean_line.replace("Size:", "").strip()

                    # Image shape identify karein
                    if shape.shape_type == 13:
                        images.append(shape.image)

                if not outlet_name:
                    outlet_name = f"Slide_{i+1}"

                # Naming components combine karein
                name_components = [
                    sanitize(p)
                    for p in [outlet_name, contact_no, media_type, size]
                    if p
                ]
                base_name = "_".join(name_components)

                # Images ko zip file me save karein
                for idx, img in enumerate(images):
                    ext = img.ext
                    suffix = f"_{idx+1}" if len(images) > 1 else ""
                    file_name = f"{base_name}{suffix}.{ext}"

                    # Zip me image stream write karein
                    zip_file.writestr(file_name, img.blob)

    st.success("Sabhi images successfully extract aur rename ho gayi hain!")

    # Download button for ZIP file
    st.download_button(
        label="📥 Download All Images (ZIP)",
        data=zip_buffer.getvalue(),
        file_name="Extracted_Images.zip",
        mime="application/zip",
    )
