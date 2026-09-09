import io
import re
import zipfile
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import streamlit as st

st.set_page_config(page_title="PPT Far-View Extractor", layout="centered")
st.title("PPT Right Image (With Markup) Extractor")


def extract_metadata_from_slide(slide):
    text_data = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            text_data.append(shape.text_frame.text)

    full_text = " ".join(text_data)
    clean_text = re.sub(r"\s+", " ", full_text)

    outlet_name = "OUTLET"
    contact = "0000000000"
    media_type = "NL"
    size = "0x0"

    outlet_match = re.search(
        r"Outlet\s*Name\s*:\s*([^:\n\r]+?)(?=\s*Address|\s*City|\s*Contact|\s*Type|$)",
        clean_text,
        re.IGNORECASE,
    )
    if outlet_match:
        raw_name = outlet_match.group(1).strip()
        cleaned = re.sub(r"[^\w\s-]", "", raw_name)
        outlet_name = re.sub(r"\s+", "_", cleaned).upper().strip("_")

    contact_match = re.search(
        r"(?:Contact|Mobile|Phone)?\s*:?\s*([6-9]\d{9})",
        clean_text,
        re.IGNORECASE,
    )
    if contact_match:
        contact = contact_match.group(1).strip()

    type_match = re.search(
        r"Type\s*:\s*([A-Za-z0-9_-]+)", clean_text, re.IGNORECASE
    )
    if type_match:
        media_type = type_match.group(1).strip().upper()

    size_match = re.search(
        r"Size\s*:\s*(\d+)\s*[*xX]\s*(\d+)", clean_text, re.IGNORECASE
    )
    if size_match:
        size = f"{size_match.group(1)}x{size_match.group(2)}"

    return f"{outlet_name}_{contact}_{media_type}_{size}"


def add_markup_to_image(img_bytes, shape, slide):
    """Image ke upar slide me jude Markup shapes (like Red Box) ko merge karta hai."""
    image = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    draw = ImageDraw.Draw(image)
    img_w, img_h = image.size

    # Slide me jitne non-picture shapes (markup/rectangles) hain, unhe dhundhna
    for s in slide.shapes:
        if s.shape_type != MSO_SHAPE_TYPE.PICTURE and hasattr(s, "left"):
            # Check agar markup shape right-side wali image ke boundary ke andar ya paas hai
            if (
                s.left >= shape.left - 500000
                and (s.left + s.width) <= (shape.left + shape.width) + 500000
            ):
                # Shape coordinates ko image pixel scale me calculate karna
                rel_x1 = max(0, int((s.left - shape.left) / shape.width * img_w))
                rel_y1 = max(0, int((s.top - shape.top) / shape.height * img_h))
                rel_x2 = min(
                    img_w,
                    int(
                        (s.left + s.width - shape.left)
                        / shape.width
                        * img_w
                    ),
                )
                rel_y2 = min(
                    img_h,
                    int(
                        (s.top + s.height - shape.top)
                        / shape.height
                        * img_h
                    ),
                )

                # Red Color markup line draw karna (Width = 6px)
                if rel_x2 > rel_x1 and rel_y2 > rel_y1:
                    draw.rectangle(
                        [rel_x1, rel_y1, rel_x2, rel_y2],
                        outline="red",
                        width=6,
                    )

    output = io.BytesIO()
    image.save(output, format="PNG")
    return output.getvalue()


uploaded_file = st.file_uploader(
    "Apni PPTX File Upload Karein", type=["pptx"]
)

if uploaded_file is not None:
    if st.button("Extract Right Image With Markup"):
        with st.spinner("Processing slides & merging markup..."):
            prs = Presentation(uploaded_file)
            zip_buffer = io.BytesIO()
            extracted_count = 0

            slide_width = prs.slide_width

            with zipfile.ZipFile(
                zip_buffer, "a", zipfile.ZIP_DEFLATED, False
            ) as zip_file:
                for idx, slide in enumerate(prs.slides):
                    filename_prefix = extract_metadata_from_slide(slide)
                    img_idx = 1

                    for shape in slide.shapes:
                        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                            # Right-side Image Check
                            if shape.left > (slide_width / 2):
                                raw_bytes = shape.image.blob

                                # Markup Merge Logic Call
                                final_image_bytes = add_markup_to_image(
                                    raw_bytes, shape, slide
                                )

                                final_name = (
                                    f"{filename_prefix}_{img_idx}.png"
                                )
                                zip_file.writestr(final_name, final_image_bytes)

                                img_idx += 1
                                extracted_count += 1

            if extracted_count > 0:
                st.success(
                    f"Total {extracted_count} Images (With Red Markup) extracted successfully!"
                )
                st.download_button(
                    label="Download ZIP with Markup",
                    data=zip_buffer.getvalue(),
                    file_name="outlet_far_views_with_markup.zip",
                    mime="application/zip",
                )
            else:
                st.warning("PPT me koi Right-side image nahi mili.")
