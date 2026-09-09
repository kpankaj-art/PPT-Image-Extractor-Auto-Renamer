import io
import re
import zipfile
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import streamlit as st
from datetime import datetime

# ✅ Page config (sirf ek baar)
if "page_loaded" not in st.session_state:
    st.set_page_config(
        page_title="PPT Far-View Extractor",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    st.session_state.page_loaded = True

st.title("📸 PPT Image Extractor with Markup Merge")

# ============ FUNCTIONS ============

def extract_metadata_from_slide(slide):
    """Slide se metadata extract karna"""
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
    
    # Outlet Name
    outlet_match = re.search(
        r"Outlet\s*Name\s*:\s*([^:\n\r]+?)(?=\s*Address|\s*City|\s*Contact|\s*Type|$)",
        clean_text,
        re.IGNORECASE,
    )
    if outlet_match:
        raw_name = outlet_match.group(1).strip()
        cleaned = re.sub(r"[^\w\s-]", "", raw_name)
        outlet_name = re.sub(r"\s+", "_", cleaned).upper().strip("_")
    
    # Contact
    contact_match = re.search(
        r"(?:Contact|Mobile|Phone)?\s*:?\s*([6-9]\d{9})",
        clean_text,
        re.IGNORECASE,
    )
    if contact_match:
        contact = contact_match.group(1).strip()
    
    # Type
    type_match = re.search(
        r"Type\s*:\s*([A-Za-z0-9_-]+)", clean_text, re.IGNORECASE
    )
    if type_match:
        media_type = type_match.group(1).strip().upper()
    
    # Size
    size_match = re.search(
        r"Size\s*:\s*(\d+)\s*[*xX]\s*(\d+)", clean_text, re.IGNORECASE
    )
    if size_match:
        size = f"{size_match.group(1)}x{size_match.group(2)}"
    
    return {
        "filename_prefix": f"{outlet_name}_{contact}_{media_type}_{size}",
        "outlet_name": outlet_name,
        "contact": contact,
        "media_type": media_type,
        "size": size
    }


def get_all_slide_text(slide):
    """Slide se sabhi text ko collect karna"""
    all_text = []
    for shape in slide.shapes:
        if shape.has_text_frame and shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            text = shape.text_frame.text.strip()
            if text:
                all_text.append(text)
    return all_text


def add_text_overlay_to_image(image_bytes, text_list):
    """Image ke top-left pe text ko overlay karna"""
    try:
        img = Image.open(io.BytesIO(image_bytes))
        
        # Convert to RGB agar RGBA hai
        if img.mode in ('RGBA', 'LA', 'P'):
            img = img.convert('RGB')
        
        # Create overlay layer
        overlay = Image.new('RGBA', img.size, (0, 0, 0, 0))
        overlay_draw = ImageDraw.Draw(overlay)
        
        # Add text
        y_position = 10
        for text_line in text_list[:10]:  # Max 10 lines
            try:
                # Try to use a font, fallback to default
                try:
                    font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 16)
                except:
                    font = ImageFont.load_default()
                
                # Get text bounding box
                bbox = overlay_draw.textbbox((10, y_position), text_line, font=font)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                
                # Draw semi-transparent background
                overlay_draw.rectangle(
                    [10 - 5, y_position - 5, text_width + 15, y_position + text_height + 5],
                    fill=(0, 0, 0, 180)
                )
                
                # Draw white text
                overlay_draw.text((10, y_position), text_line, font=font, fill=(255, 255, 255, 255))
                
                y_position += text_height + 10
            except:
                pass
        
        # Composite overlay on original image
        img = img.convert('RGBA')
        img.paste(overlay, (0, 0), overlay)
        img = img.convert('RGB')
        
        # Save
        output = io.BytesIO()
        img.save(output, format='PNG', quality=90)
        output.seek(0)
        return output.getvalue()
    
    except Exception as e:
        st.warning(f"Text overlay error: {e}")
        return image_bytes


def process_ppt_file_with_markup(uploaded_file, merge_markup_flag):
    """PPT processing with text/markup merge"""
    try:
        prs = Presentation(uploaded_file)
        zip_buffer = io.BytesIO()
        extracted_data = []
        total_slides = len(prs.slides)
        
        slide_width = prs.slide_width
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for slide_idx, slide in enumerate(prs.slides):
                # Update progress
                progress = (slide_idx + 1) / total_slides
                progress_bar.progress(progress)
                status_text.text(f"Processing: {slide_idx + 1}/{total_slides} slides...")
                
                metadata = extract_metadata_from_slide(slide)
                filename_prefix = metadata["filename_prefix"]
                
                # Get all text from slide
                slide_text = get_all_slide_text(slide)
                
                img_idx = 1
                
                for shape in slide.shapes:
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        # Right-side filter
                        if shape.left > (slide_width / 2):
                            try:
                                image_bytes = shape.image.blob
                                image_ext = shape.image.ext
                                
                                # Add text overlay agar toggle on hai
                                if merge_markup_flag and slide_text:
                                    image_bytes = add_text_overlay_to_image(image_bytes, slide_text)
                                
                                final_name = f"{filename_prefix}_{img_idx}.{image_ext}"
                                zip_file.writestr(final_name, image_bytes)
                                
                                extracted_data.append({
                                    "slide": slide_idx + 1,
                                    "filename": final_name,
                                    "outlet": metadata["outlet_name"],
                                    "contact": metadata["contact"],
                                    "type": metadata["media_type"],
                                    "size": metadata["size"],
                                    "markup": "✅ Yes" if (merge_markup_flag and slide_text) else "❌ No"
                                })
                                
                                img_idx += 1
                            except Exception as e:
                                pass
        
        progress_bar.empty()
        status_text.empty()
        return zip_buffer, extracted_data
    
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        return None, []


# ============ SIDEBAR ============

st.sidebar.header("⚙️ Settings")
merge_markup = st.sidebar.checkbox("✅ Merge Slide Text/Markup with Image", value=True)

st.sidebar.markdown("""
### 📋 Features:
✓ Extract right-side images
✓ Get all slide text/markup
✓ Overlay text on images  
✓ Auto metadata naming
✓ Lightweight & Fast
""")


# ============ MAIN ============

uploaded_file = st.file_uploader("📁 Upload PPTX File", type=["pptx"])

if uploaded_file is not None:
    st.divider()
    
    col1, col2 = st.columns([3, 1])
    with col1:
        st.info(f"📄 {uploaded_file.name} | {uploaded_file.size / (1024*1024):.2f} MB")
    
    with col2:
        if st.button("🚀 Extract", use_container_width=True, type="primary"):
            if uploaded_file.size > 500 * 1024 * 1024:
                st.error("❌ File > 500MB!")
            else:
                zip_buffer, extracted_data = process_ppt_file_with_markup(uploaded_file, merge_markup)
                
                if extracted_data:
                    st.success(f"✅ {len(extracted_data)} images extracted with markup overlays!")
                    
                    # Table
                    st.subheader("📊 Results")
                    table_data = [{
                        "Slide": i["slide"],
                        "Outlet": i["outlet"],
                        "Contact": i["contact"],
                        "Type": i["type"],
                        "Size": i["size"],
                        "Markup": i["markup"],
                        "File": i["filename"]
                    } for i in extracted_data]
                    
                    st.dataframe(table_data, use_container_width=True, hide_index=True)
                    
                    # Download
                    st.divider()
                    st.download_button(
                        "⬇️ Download ZIP",
                        zip_buffer.getvalue(),
                        f"images_marked_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                        "application/zip",
                        use_container_width=True,
                        type="primary"
                    )
                    
                    # Stats
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Images", len(extracted_data))
                    with col2:
                        st.metric("Unique Outlets", len(set([i["outlet"] for i in extracted_data])))
                    with col3:
                        with_markup = sum(1 for i in extracted_data if i["markup"] == "✅ Yes")
                        st.metric("With Markup", with_markup)
                
                else:
                    st.warning("⚠️ No images found!")

else:
    st.info("""
    ### 👋 Welcome to PPT Image Extractor!
    
    **What this tool does:**
    ✅ Extract right-side images from PowerPoint
    ✅ Get all text/markup from slide
    ✅ Overlay text on extracted images
    ✅ Auto-name based on metadata  
    ✅ Download as ZIP
    
    **How to use:**
    1. Upload PPTX file
    2. Toggle "Merge Slide Text/Markup" in settings (default: ON)
    3. Click "Extract" button
    4. Download ZIP file
    
    **Result:**
    - Images with text overlay on top-left
    - Professional markup integration
    - All metadata preserved
    """)

st.divider()
st.caption("Made with ❤️ | PPT Extractor v3.1 - Text Overlay Edition")
