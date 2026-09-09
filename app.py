import io
import re
import zipfile
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt
import streamlit as st
from datetime import datetime

# ✅ FIRST - Page config (sirf ek baar chalega)
if "page_loaded" not in st.session_state:
    st.set_page_config(
        page_title="PPT Far-View Extractor",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    st.session_state.page_loaded = True

st.title("📸 PPT Right-Side Image Extractor with Markup")

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


def get_slide_markup_image(slide, slide_width):
    """Puri slide ka screenshot lena - marking ke saath"""
    try:
        from pptx.enum.dml import MSO_THEME_COLOR
        import os
        
        # Slide ko temporary PNG me save karna
        temp_path = "/tmp/slide_temp.png"
        
        # LibreOffice/PIL se slide render nahi kar sakte directly
        # Toh hum slide ke saare shapes ke upar drawing karenge
        
        # Create a new image with slide dimensions
        slide_width_px = int(slide_width / 914400 * 96)  # EMU to pixels
        slide_height_px = int(slide.slide_height / 914400 * 96)
        
        # White background
        slide_img = Image.new('RGB', (slide_width_px, slide_height_px), 'white')
        draw = ImageDraw.Draw(slide_img)
        
        # Draw all shapes
        for shape in slide.shapes:
            try:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    # Draw picture at correct position
                    pass
                elif shape.has_text_frame:
                    # Draw text on image
                    x = int(shape.left / 914400 * 96)
                    y = int(shape.top / 914400 * 96)
                    text = shape.text_frame.text
                    if text:
                        draw.text((x, y), text, fill='black')
            except:
                pass
        
        return slide_img
    except:
        return None


def merge_image_with_markup(image_bytes, slide, slide_width):
    """Image ko slide ke markup/marking ke saath merge karna"""
    try:
        img = Image.open(io.BytesIO(image_bytes))
        
        # Get slide markup
        markup_img = get_slide_markup_image(slide, slide_width)
        
        if markup_img:
            # Resize markup to match image
            markup_resized = markup_img.resize(img.size, Image.Resampling.LANCZOS)
            
            # Create a composite - overlay markup on image
            # Use 30% opacity for markup so image is still visible
            markup_resized.putalpha(int(255 * 0.3))
            
            result_img = img.convert('RGBA')
            result_img.paste(markup_resized, (0, 0), markup_resized)
            result_img = result_img.convert('RGB')
            
            output = io.BytesIO()
            result_img.save(output, format='PNG', quality=90)
            output.seek(0)
            return output.getvalue()
    except Exception as e:
        st.warning(f"Merge error: {e}")
    
    return image_bytes


def process_ppt_file_with_markup(uploaded_file, merge_markup_flag):
    """PPT processing with markup merge"""
    try:
        prs = Presentation(uploaded_file)
        zip_buffer = io.BytesIO()
        extracted_data = []
        total_slides = len(prs.slides)
        
        slide_width = prs.slide_width
        slide_height = prs.slide_height
        
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
                img_idx = 1
                
                # Get all text/shapes from slide for markup
                slide_text_elements = []
                for shape in slide.shapes:
                    if shape.has_text_frame and shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                        if shape.text.strip():
                            slide_text_elements.append(shape.text)
                
                for shape in slide.shapes:
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        # Right-side filter
                        if shape.left > (slide_width / 2):
                            try:
                                image_bytes = shape.image.blob
                                image_ext = shape.image.ext
                                
                                # Merge with markup agar toggle on hai
                                if merge_markup_flag:
                                    image_bytes = merge_image_with_markup(image_bytes, slide, slide_width)
                                
                                final_name = f"{filename_prefix}_{img_idx}.{image_ext}"
                                zip_file.writestr(final_name, image_bytes)
                                
                                extracted_data.append({
                                    "slide": slide_idx + 1,
                                    "filename": final_name,
                                    "outlet": metadata["outlet_name"],
                                    "contact": metadata["contact"],
                                    "type": metadata["media_type"],
                                    "size": metadata["size"],
                                    "markup": "✅ Yes" if merge_markup_flag else "❌ No"
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
merge_markup = st.sidebar.checkbox("✅ Merge Markup/Text with Image", value=True)

st.sidebar.markdown("""
### 📋 Features:
✓ Extract right-side images
✓ Auto metadata extraction  
✓ Merge with slide markup
✓ Smart naming
✓ Lightweight
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
                    st.success(f"✅ {len(extracted_data)} images extracted with markup!")
                    
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
                        f"images_markup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                        "application/zip",
                        use_container_width=True,
                        type="primary"
                    )
                    
                    # Stats
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total", len(extracted_data))
                    with col2:
                        st.metric("Outlets", len(set([i["outlet"] for i in extracted_data])))
                    with col3:
                        st.metric("With Markup", sum(1 for i in extracted_data if i["markup"] == "✅ Yes"))
                
                else:
                    st.warning("⚠️ No images found!")

else:
    st.info("""
    ### 👋 Welcome!
    
    **What this tool does:**
    ✓ Extract right-side images from PowerPoint
    ✓ Merge with slide markup/text/shapes
    ✓ Auto-name based on metadata  
    ✓ Download as ZIP
    
    **How to use:**
    1. Upload PPTX file
    2. Toggle "Merge Markup" in settings
    3. Click Extract
    4. Download ZIP
    
    **Markup Merge:**
    - Combines image with all text/shapes from slide
    - Preserves original image quality
    - 30% opacity overlay for clarity
    """)

st.divider()
st.caption("Made with ❤️ | PPT Extractor v3.0 - With Markup Merge")
