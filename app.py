import io
import tempfile
import zipfile
import aspose.slides as slides
import streamlit as st

st.set_page_config(page_title="PPT Slide to Image Extractor", layout="centered")
st.title("PPT Slide to Image (With Markings & Shapes)")

uploaded_file = st.file_uploader("Apni PPTX File Upload Karein", type=["pptx", "ppt"])

if uploaded_file is not None:
    if st.button("Extract Images (With Markings)"):
        with st.spinner("Processing PPT and Rendering Slides..."):
            with tempfile.TemporaryDirectory() as temp_dir:
                # Save uploaded file temporarily
                temp_ppt_path = f"{temp_dir}/input.pptx"
                with open(temp_ppt_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                # Load PPTX using Aspose.Slides
                presentation = slides.Presentation(temp_ppt_path)
                zip_buffer = io.BytesIO()

                # High quality scale factor (2x resolution)
                scale_x = 2.0
                scale_y = 2.0

                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for i, slide in enumerate(presentation.slides):
                        # Render slide with markings into image
                        image_stream = io.BytesIO()
                        with slide.get_image(scale_x, scale_y) as slide_image:
                            slide_image.save(image_stream, slides.ImageFormat.JPEG)

                        # Save into zip
                        zip_file.writestr(f"slide_{i + 1}_marked.jpg", image_stream.getvalue())

                st.success(f"Kul {len(presentation.slides)} slides marking ke sath convert ho gayi hain!")
                st.download_button(
                    label="Download All Merged Images (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="slides_with_markings.zip",
                    mime="application/zip",
                )
