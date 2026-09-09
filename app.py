import io
import zipfile
import streamlit as st
from pptx import Presentation

st.set_page_config(page_title="PPT Image Extractor", layout="centered")
st.title("PPT Image Extractor (With Shapes & Markings)")

uploaded_file = st.file_uploader("Apni PPTX File Upload Karein", type=["pptx"])

if uploaded_file is not None:
    with st.spinner("PPT process ho rahi hai..."):
        prs = Presentation(uploaded_file)
        zip_buffer = io.BytesIO()
        image_count = 0

        with zipfile.ZipFile(
            zip_buffer, "a", zipfile.ZIP_DEFLATED, False
        ) as zip_file:
            for slide_index, slide in enumerate(prs.slides):
                for shape_index, shape in enumerate(slide.shapes):
                    # Check if shape contains an image
                    if shape.shape_type == 13 or hasattr(shape, "image"):
                        image = shape.image
                        image_bytes = image.blob
                        image_ext = image.ext

                        # File Name format
                        image_name = f"slide_{slide_index + 1}_img_{shape_index + 1}.{image_ext}"
                        zip_file.writestr(image_name, image_bytes)
                        image_count += 1

        if image_count > 0:
            st.success(f"Kul {image_count} images successfully extract ho gayi hain!")
            st.download_button(
                label="Saari Images ZIP Mein Download Karein",
                data=zip_buffer.getvalue(),
                file_name="ppt_extracted_images.zip",
                mime="application/zip",
            )
        else:
            st.warning("Is PPT mein koi image nahi mili.")
