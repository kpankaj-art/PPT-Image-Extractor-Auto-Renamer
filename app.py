import io
import json
import zipfile
import gc
import streamlit as st
from pptx import Presentation
from PIL import Image
from google import genai
from google.genai import types

st.set_page_config(page_title="PPT Image Extractor", page_icon="📊", layout="wide")

st.title("📊 PowerPoint (.PPTX) Right Image Extractor")
st.write("Apni PPT file upload karein. Gemini AI automatically red markings merge karke sirf Right Photo crop karega.")

# Read Gemini API Key from Streamlit Secrets
api_key = st.secrets.get("GEMINI_API_KEY") or st.sidebar.text_input("Gemini API Key", type="password")

if not api_key:
    st.error("⚠️ GEMINI_API_KEY Streamlit Secrets me nahi mila! Kripya Settings -> Secrets check karein.")
    st.stop()

client = genai.Client(api_key=api_key)

# STRICT PPTX ONLY UPLOADER
uploaded_ppt = st.file_uploader("📂 Select PowerPoint File (.pptx)", type=["pptx"])

def get_slide_merged_canvas(slide, slide_width, slide_height, scale=2.0):
    """PPT Shapes & Overlays (Red Box) ko ek canvas par merge karna"""
    canvas_w = int(slide_width * scale)
    canvas_h = int(slide_height * scale)
    canvas = Image.new("RGB", (canvas_w, canvas_h), (255, 255, 255))
    
    for shape in slide.shapes:
        if shape.shape_type == 13:  # Picture shape
            try:
                img_data = io.BytesIO(shape.image.blob)
                sub_img = Image.open(img_data).convert("RGBA")
                
                left = int((shape.left / slide_width) * canvas_w)
                top = int((shape.top / slide_height) * canvas_h)
                width = int((shape.width / slide_width) * canvas_w)
                height = int((shape.height / slide_height) * canvas_h)
                
                sub_img = sub_img.resize((max(1, width), max(1, height)), Image.Resampling.LANCZOS)
                canvas.paste(sub_img, (left, top), sub_img)
            except Exception:
                pass

    return canvas

def call_gemini_vision(canvas_img, slide_text, slide_num):
    """Gemini 2.5 Flash se Right Photo Bounding Box aur Filename details lena"""
    prompt = f"""
    This is PowerPoint Slide {slide_num}. Text found: "{slide_text}"

    Analyze this slide image:
    1. Extract "outlet_name": Name of the Shop/Outlet.
    2. Extract "contact_no": 10-digit Mobile Number.
    3. Extract "media_type": Type like NL, FL, BL, GSB, FLEX.
    4. Extract "size": Board dimensions (e.g. 10x2).
    5. Locate ONLY the RIGHT SIDE / FAR VIEW photo box (including red box markings and geotag map stamps inside it).
       Return "box_2d": Bounding box coordinates [ymin, xmin, ymax, xmax] (scale 0 to 1000).

    Return ONLY raw JSON:
    {{
        "outlet_name": "STRING",
        "contact_no": "STRING",
        "media_type": "STRING",
        "size": "STRING",
        "box_2d": [ymin, xmin, ymax, xmax]
    }}
    """

    img_bytes = io.BytesIO()
    canvas_img.save(img_bytes, format="JPEG", quality=85)
    
    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[
                types.Part.from_bytes(data=img_bytes.getvalue(), mime_type="image/jpeg"),
                prompt
            ],
            config=types.GenerateContentConfig(response_mime_type="application/json")
        )
        return json.loads(response.text)
    except Exception:
        return None

if uploaded_ppt is not None:
    prs = Presentation(uploaded_ppt)
    total_slides = len(prs.slides)
    st.sidebar.success(f"✅ Total PPT Slides: {total_slides}")

    slide_w = prs.slide_width
    slide_h = prs.slide_height

    # Batch Range Selector
    start_slide = st.sidebar.number_input("Start Slide", min_value=1, max_value=total_slides, value=1)
    end_slide = st.sidebar.number_input("End Slide", min_value=1, max_value=total_slides, value=min(total_slides, 50))

    if st.button("🚀 Process PPT Slides & Extract Images"):
        zip_buffer = io.BytesIO()
        progress_bar = st.progress(0)
        status_text = st.empty()

        slides_range = range(start_slide - 1, end_slide)
        total_batch = len(slides_range)

        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for idx, i in enumerate(slides_range):
                slide = prs.slides[i]
                status_text.text(f"Processing PPT Slide {i+1} of {total_slides}...")
                progress_bar.progress((idx + 1) / total_batch)

                # PPT Text Extraction
                slide_text = []
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for p in shape.text_frame.paragraphs:
                            if p.text.strip():
                                slide_text.append(p.text.strip())
                full_text = " ".join(slide_text)

                # PPT Image Canvas Render
                canvas_img = get_slide_merged_canvas(slide, slide_w, slide_h, scale=2.0)
                cw, ch = canvas_img.size

                # Call Gemini Vision API
                ai_res = call_gemini_vision(canvas_img, full_text, i+1)

                if ai_res:
                    name = ai_res.get("outlet_name", f"SLIDE_{i+1}").replace(" ", "_")
                    mobile = ai_res.get("contact_no", "")
                    m_type = ai_res.get("media_type", "")
                    m_size = ai_res.get("size", "").replace(" ", "")

                    components = [c for c in [name, mobile, m_type, m_size] if c]
                    filename = "_".join(components) + ".jpg"

                    box = ai_res.get("box_2d")
                    if box and len(box) == 4:
                        ymin, xmin, ymax, xmax = box
                        crop_coords = (
                            int((xmin / 1000.0) * cw),
                            int((ymin / 1000.0) * ch),
                            int((xmax / 1000.0) * cw),
                            int((ymax / 1000.0) * ch)
                        )
                        cropped = canvas_img.crop(crop_coords)
                    else:
                        cropped = canvas_img.crop((int(cw * 0.50), int(ch * 0.35), int(cw * 0.74), int(ch * 0.78)))
                else:
                    filename = f"SLIDE_{i+1}.jpg"
                    cropped = canvas_img.crop((int(cw * 0.50), int(ch * 0.35), int(cw * 0.74), int(ch * 0.78)))

                # Save cropped image into ZIP
                out_b = io.BytesIO()
                cropped.save(out_b, format="JPEG", quality=92)
                zip_file.writestr(filename, out_b.getvalue())

                del canvas_img
                del cropped
                gc.collect()

        status_text.text("Processing Complete!")
        st.success(f"🎉 Processed slides {start_slide} to {end_slide} successfully!")
        st.download_button(
            label="📥 Download PPT Cropped Images (ZIP)",
            data=zip_buffer.getvalue(),
            file_name=f"PPT_Cropped_{start_slide}_to_{end_slide}.zip",
            mime="application/zip",
        )
