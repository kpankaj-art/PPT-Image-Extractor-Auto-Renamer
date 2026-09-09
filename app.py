import io
import re
import zipfile
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import streamlit as st
from datetime import datetime

# Page Configuration
st.set_page_config(
    page_title="PPT Far-View Extractor",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📸 PPT Right-Side Image Extractor & Auto-Renamer")

# ============ FUNCTIONS ============

def extract_metadata_from_slide(slide):
    """Slide ke sabhi text elements se metadata robustly read karta hai."""
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
    
    # 1. Outlet Name Match
    outlet_match = re.search(
        r"Outlet\s*Name\s*:\s*([^:\n\r]+?)(?=\s*Address|\s*City|\s*Contact|\s*Type|$)",
        clean_text,
        re.IGNORECASE,
    )
    if outlet_match:
        raw_name = outlet_match.group(1).strip()
        cleaned = re.sub(r"[^\w\s-]", "", raw_name)
        outlet_name = re.sub(r"\s+", "_", cleaned).upper().strip("_")
    
    # 2. Contact Match (Exact 10-digit Mobile Number)
    contact_match = re.search(
        r"(?:Contact|Mobile|Phone)?\s*:?\s*([6-9]\d{9})",
        clean_text,
        re.IGNORECASE,
    )
    if contact_match:
        contact = contact_match.group(1).strip()
    
    # 3. Type Match
    type_match = re.search(
        r"Type\s*:\s*([A-Za-z0-9_-]+)", clean_text, re.IGNORECASE
    )
    if type_match:
        media_type = type_match.group(1).strip().upper()
    
    # 4. Size Match (e.g. 10 x 2, 10x2)
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


def add_simple_watermark(image_bytes, watermark_text):
    """Lightweight watermark - बिना heavy processing के"""
    try:
        img = Image.open(io.BytesIO(image_bytes))
        
        # Resize करो अगर बहुत बड़ा है (memory save करने के लिए)
        max_size = (1920, 1080)
        if img.size[0] > max_size[0] or img.size[1] > max_size[1]:
            img.thumbnail(max_size, Image.Resampling.LANCZOS)
        
        draw = ImageDraw.Draw(img)
        
        # Simple text without fancy fonts
        text = f"{watermark_text}"
        draw.text((5, 5), text, fill="white")
        
        # Save with compression
        output = io.BytesIO()
        img.save(output, format="PNG", optimize=True, quality=85)
        output.seek(0)
        return output.getvalue()
    except Exception as e:
        st.warning(f"Watermark error: {e}")
        return image_bytes


def process_ppt_file(uploaded_file, add_watermark_flag, progress_bar):
    """PPT file को lightweight तरीके से process करना"""
    try:
        prs = Presentation(uploaded_file)
        zip_buffer = io.BytesIO()
        extracted_data = []
        total_slides = len(prs.slides)
        
        slide_width = prs.slide_width
        
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for slide_idx, slide in enumerate(prs.slides):
                # Progress update
                progress = (slide_idx + 1) / total_slides
                progress_bar.progress(progress, text=f"Processing: {slide_idx + 1}/{total_slides}")
                
                metadata = extract_metadata_from_slide(slide)
                filename_prefix = metadata["filename_prefix"]
                img_idx = 1
                
                for shape in slide.shapes:
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        # Right-side image filter
                        if shape.left > (slide_width / 2):
                            try:
                                image_bytes = shape.image.blob
                                image_ext = shape.image.ext
                                
                                # Add watermark sirf agar toggle on hai
                                if add_watermark_flag:
                                    watermark_text = f"{metadata['outlet_name']}|{metadata['contact']}"
                                    image_bytes = add_simple_watermark(image_bytes, watermark_text)
                                
                                final_name = f"{filename_prefix}_{img_idx}.{image_ext}"
                                zip_file.writestr(final_name, image_bytes)
                                
                                extracted_data.append({
                                    "slide": slide_idx + 1,
                                    "filename": final_name,
                                    "outlet": metadata["outlet_name"],
                                    "contact": metadata["contact"],
                                    "type": metadata["media_type"],
                                    "size": metadata["size"]
                                })
                                
                                img_idx += 1
                            except Exception as e:
                                st.warning(f"Slide {slide_idx + 1} image error: {e}")
        
        return zip_buffer, extracted_data
    
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        return None, []


# ============ SIDEBAR ============

with st.sidebar:
    st.header("⚙️ Settings")
    add_watermark = st.checkbox("✅ Image पर Watermark लगाएं", value=False)
    st.divider()
    st.markdown("### 📋 Features:")
    st.markdown("""
    ✓ Right-side images निकालना
    ✓ Auto metadata extraction
    ✓ Smart image naming
    ✓ Lightweight processing
    """)


# ============ MAIN APP ============

uploaded_file = st.file_uploader(
    "📁 PPTX File Upload करें",
    type=["pptx"]
)

if uploaded_file is not None:
    st.divider()
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.info(f"📄 File: **{uploaded_file.name}** | Size: **{uploaded_file.size / (1024*1024):.2f} MB**")
    
    with col2:
        process_button = st.button(
            "🚀 Extract करें",
            use_container_width=True,
            type="primary"
        )
    
    if process_button:
        if uploaded_file.size > 500 * 1024 * 1024:  # 500MB check
            st.error("❌ File 500MB से बड़ी है! छोटी file upload करें।")
        else:
            progress_bar = st.progress(0, text="Processing...")
            
            zip_buffer, extracted_data = process_ppt_file(uploaded_file, add_watermark, progress_bar)
            
            if extracted_data:
                st.success(f"✅ {len(extracted_data)} images successfully extracted!")
                
                # Show table
                st.subheader("📊 Extracted Images")
                
                table_data = []
                for item in extracted_data:
                    table_data.append({
                        "Slide": item["slide"],
                        "Outlet": item["outlet"],
                        "Contact": item["contact"],
                        "Type": item["type"],
                        "Size": item["size"],
                        "File": item["filename"]
                    })
                
                st.dataframe(table_data, use_container_width=True, hide_index=True)
                
                # Download button
                st.divider()
                st.download_button(
                    label="⬇️ Download ZIP",
                    data=zip_buffer.getvalue(),
                    file_name=f"images_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
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
                    st.metric("Watermark", "✅" if add_watermark else "❌")
            
            else:
                st.warning("⚠️ कोई image नहीं मिली या error हुई।")

else:
    st.info("""
    ### 👋 यह Tool करता है:
    
    ✓ PowerPoint से right-side images extract करना
    ✓ Slide metadata से automatic naming
    ✓ Lightweight processing (500MB तक)
    ✓ ZIP में download करना
    
    **कैसे करें:** File upload करें → Extract बटन दबाएं → Download करें
    """)

st.divider()
st.caption("Made with ❤️ | PPT Extractor v2.0")
