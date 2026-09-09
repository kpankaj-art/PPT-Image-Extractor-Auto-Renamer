import io
import os
import tempfile
import zipfile
from pdf2image import convert_from_path
import streamlit as st

st.set_page_config(page_title="Lite PPT Extractor", layout="centered")
st.title("PPT Slide & Markings Extractor (1GB RAM Optimized)")

uploaded_file = st.file_uploader("Apni PPTX File Upload Karein", type=["pptx", "ppt"])

if uploaded_file is not None:
    if st.button("Convert to Images"):
        with st.spinner("Processing (Low Memory Mode)..."):
            with tempfile.TemporaryDirectory() as temp_dir:
                ppt_path = os.path.join(temp_dir, "input.pptx")
                pdf_path = os.path.join(temp_dir, "input.pdf")

                # Uploaded PPT save
                with open(ppt_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                # LibreOffice ki jagah lightweight python script call
                # Note: Pure memory optimization ke sath pdf rendering
                try:
                    from pptx import Presentation
                    from PIL import Image, ImageDraw

                    prs = Presentation(ppt_path)
                    zip_buffer = io.BytesIO()

                    with zipfile.ZipFile(
                        zip_buffer, "a", zipfile.ZIP_DEFLATED, False
                    ) as zip_file:
                        for idx, slide in enumerate(prs.slides):
                            # Default Slide Dimension
                            slide_width = int(prs.slide_width.pt)
                            slide_height = int(prs.slide_height.pt)

                            # Blank Image base banayein
                            img = Image.new(
                                "RGB", (slide_width, slide_height), "white"
                            )

                            # Slide shapes aur images extract karke base image par place karein
                            for shape in slide.shapes:
                                if shape.has_image:
                                    image_bytes = shape.image.blob
                                    shape_img = Image.open(
                                        io.BytesIO(image_bytes)
                                    )

                                    # Resize to shape dimensions
                                    shape_img = shape_img.resize(
                                        (
                                            int(shape.width.pt),
                                            int(shape.height.pt),
                                        )
                                    )
                                    img.paste(
                                        shape_img,
                                        (
                                            int(shape.left.pt),
                                            int(shape.top.pt),
                                        ),
                                    )

                            # Save as PNG
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
