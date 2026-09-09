import io
import os
import tempfile
import zipfile
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import streamlit as st

st.set_page_config(page_title="Lite PPT Extractor", layout="centered")
st.title("PPT Slide & Markings Extractor (1GB RAM Optimized)")

uploaded_file = st.file_uploader(
    "Apni PPTX File Upload Karein", type=["pptx", "ppt"]
)

if uploaded_file is not None:
    if st.button("Convert to Images"):
        with st.spinner("Processing (Low Memory Mode)..."):
            try:
                prs = Presentation(uploaded_file)
                zip_buffer = io.BytesIO()

                with zipfile.ZipFile(
                    zip_buffer, "a", zipfile.ZIP_DEFLATED, False
                ) as zip_file:
                    for idx, slide in enumerate(prs.slides):
                        slide_width = int(prs.slide_width.pt)
                        slide_height = int(prs.slide_height.pt)

                        # Blank Image Base
                        img = Image.new(
                            "RGB", (slide_width, slide_height), "white"
                        )

                        for shape in slide.shapes:
                            # Correct way to check for picture shapes
                            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                                image_bytes = shape.image.blob
                                shape_img = Image.open(io.BytesIO(image_bytes))

                                # Resize to shape dimensions
                                shape_img = shape_img.resize(
                                    (
                                        max(1, int(shape.width.pt)),
                                        max(1, int(shape.height.pt)),
                                    )
                                )

                                # Paste image on canvas
                                img.paste(
                                    shape_img,
                                    (
                                        int(shape.left.pt),
                                        int(shape.top.pt),
                                    ),
                                )

                        # Save slide as PNG
                        img_byte_arr = io.BytesIO()
                        img.save(img_byte_arr, format="PNG", quality=85)
                        zip_file.writestr(
                            f"slide_{idx + 1}_merged.png",
                            img_byte_arr.getvalue(),
                        )

                st.success("Conversion Complete!")
                st.download_button(
                    label="Download Merged Images (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="marked_slides.zip",
                    mime="application/zip",
                )

            except Exception as e:
                st.error(f"Error processing file: {str(e)}")
