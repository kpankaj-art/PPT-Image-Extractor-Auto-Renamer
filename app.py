import io
import re
import zipfile
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches
import streamlit as st
from datetime import datetime

# Page Configuration
st.set_page_config(
    page_title="PPT Far-View Extractor",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        color: #155724;
        border: 1px solid #c3e6cb;
    }
    .info-box {
        padding: 1rem;
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 0.5rem;
        color: #0c5460;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">📸 PPT Right-Side Image Extractor & Auto-Renamer</div>', unsafe_allow_html=True)

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


def add_watermark_to_image(image_bytes, watermark_text, add_watermark=True):
    """Image ke upar metadata ko watermark ke roop me add karega"""
    try:
        img = Image.open(io.BytesIO(image_bytes))
        
        if add_watermark and watermark_text:
            # Create a copy to draw on
            img_with_watermark = img.copy()
            draw = ImageDraw.Draw(img_with_watermark)
            
            # Try to use a good font, fallback to default if not available
            try:
                font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 30)
            except:
                font = ImageFont.load_default()
            
            # Add semi-transparent background for text
            text_bbox = draw.textbbox((10, 10), watermark_text, font=font)
            bg_padding = 10
            draw.rectangle(
                [text_bbox[0] - bg_padding, text_bbox[1] - bg_padding,
                 text_bbox[2] + bg_padding, text_bbox[3] + bg_padding],
                fill=(0, 0, 0, 128)
            )
            
            # Draw text
            draw.text((10, 10), watermark_text, fill="white", font=font)
            
            # Convert back to bytes
            output = io.BytesIO()
            img_with_watermark.save(output, format=img.format or 'PNG')
            return output.getvalue()
        
        return image_bytes
    except Exception as e:
        st.warning(f"Watermark add करने में error: {e}")
        return image_bytes


def process_ppt_file(uploaded_file, add_watermark_flag, progress_bar):
    """PPT file ko process karke images extract aur rename karega"""
    try:
        prs = Presentation(uploaded_file)
        zip_buffer = io.BytesIO()
        extracted_data = []
        total_slides = len(prs.slides)
        
        slide_width = prs.slide_width
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for slide_idx, slide in enumerate(prs.slides):
                # Update progress bar
                progress = (slide_idx + 1) / total_slides
                progress_bar.progress(progress, text=f"Processing slide {slide_idx + 1}/{total_slides}")
                
                metadata = extract_metadata_from_slide(slide)
                filename_prefix = metadata["filename_prefix"]
                img_idx = 1
                
                for shape in slide.shapes:
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        # 📍 LOGIC: Sirf right-side wali image filtering (Far View)
                        if shape.left > (slide_width / 2):
                            try:
                                image_bytes = shape.image.blob
                                image_ext = shape.image.ext
                                
                                # Add watermark if enabled
                                watermark_text = f"{metadata['outlet_name']} | {metadata['contact']}"
                                if add_watermark_flag:
                                    image_bytes = add_watermark_to_image(image_bytes, watermark_text, True)
                                
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
                                st.warning(f"Slide {slide_idx + 1} की image process करने में error: {e}")
        
        return zip_buffer, extracted_data
    
    except Exception as e:
        st.error(f"❌ PPT file process करने में error: {str(e)}")
        return None, []


# ============ SIDEBAR ============

with st.sidebar:
    st.header("⚙️ Settings")
    add_watermark = st.checkbox("✅ Image पर Metadata Watermark add करें", value=True)
    st.divider()
    st.markdown("### 📋 यह App करता है:")
    st.markdown("""
    - PPT slides से right-side images extract करना
    - Slide से metadata automatically निकालना
    - Images को smart naming देना
    - Optional watermark add करना
    """)


# ============ MAIN APP ============

col1, col2 = st.columns([2, 1])

with col1:
    uploaded_file = st.file_uploader(
        "📁 अपनी PPTX File यहाँ Upload करें",
        type=["pptx"],
        help="सिर्फ .pptx format की files support होती हैं"
    )

if uploaded_file is not None:
    st.divider()
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.info(f"📄 File: **{uploaded_file.name}** | Size: **{uploaded_file.size / (1024*1024):.2f} MB**")
    
    with col2:
        process_button = st.button(
            "🚀 Extract & Rename करें",
            use_container_width=True,
            type="primary"
        )
    
    if process_button:
        progress_bar = st.progress(0, text="Processing शुरू हो रहा है...")
        
        zip_buffer, extracted_data = process_ppt_file(uploaded_file, add_watermark, progress_bar)
        
        if extracted_data:
            # Success message
            st.markdown(f"""
            <div class="success-box">
            ✅ <b>सफल!</b> कुल {len(extracted_data)} Right Images (Far View) successfully extract & rename हुई हैं!
            </div>
            """, unsafe_allow_html=True)
            
            # Display extracted images info
            st.subheader("📊 Extracted Images का विवरण")
            
            # Create a detailed table
            table_data = []
            for item in extracted_data:
                table_data.append({
                    "Slide": item["slide"],
                    "Outlet": item["outlet"],
                    "Contact": item["contact"],
                    "Type": item["type"],
                    "Size": item["size"],
                    "Filename": item["filename"]
                })
            
            st.dataframe(table_data, use_container_width=True, hide_index=True)
            
            # Download button
            st.divider()
            col1, col2, col3 = st.columns([1, 1, 1])
            
            with col2:
                st.download_button(
                    label="⬇️ Download ZIP फ़ाइल",
                    data=zip_buffer.getvalue(),
                    file_name=f"outlet_images_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                    use_container_width=True,
                    type="primary"
                )
            
            # Summary statistics
            st.divider()
            st.subheader("📈 Summary")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total Images", len(extracted_data))
            
            with col2:
                unique_outlets = len(set([item["outlet"] for item in extracted_data]))
                st.metric("Unique Outlets", unique_outlets)
            
            with col3:
                unique_types = len(set([item["type"] for item in extracted_data]))
                st.metric("Media Types", unique_types)
            
            with col4:
                st.metric("Watermark Added", "✅ Yes" if add_watermark else "❌ No")
        
        else:
            st.warning("⚠️ PPT में कोई right-side image (Far View) नहीं मिली या कोई error हुई।")

else:
    st.markdown("""
    <div class="info-box">
    
    ### 👋 स्वागत है!
    
    **यह Tool क्या करता है:**
    - PowerPoint presentations से automatically images extract करता है
    - Slide के right-side (Far View) की सभी images निकालता है  
    - Slide में लिखे metadata के हिसाब से images को smart naming देता है
    - Optional watermark के साथ images save करता है
    
    **कैसे काम करता है:**
    1. अपनी PPTX file upload करें
    2. Settings में watermark option चुनें
    3. "Extract & Rename करें" button दबाएं
    4. Extracted images की ZIP file download करें
    
    **File Format (Auto Generated):**
    ```
    OUTLET_NAME_PHONENUMBER_MEDIATYPE_SIZE_1.png
    ```
    
    Example: `FLIPKART_9876543210_BILLBOARD_10x5_1.png`
    
    </div>
    """, unsafe_allow_html=True)

# Footer
st.divider()
st.markdown("""
<div style='text-align: center; color: gray; font-size: 0.9rem;'>
Made with ❤️ | PPT Image Extractor v2.0 | All rights reserved
</div>
""", unsafe_allow_html=True)
