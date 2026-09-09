import io
import re
import zipfile
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET

import streamlit as st
from PIL import Image

st.set_page_config(
    page_title="PPT Image + Markup Extractor",
    page_icon="🖼️",
    layout="centered",
)

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}

EMU_PER_INCH = 914400


def local_name(tag):
    return tag.rsplit("}", 1)[-1]


def get_attr(element, name, default=None):
    return element.attrib.get(name, default)


def parse_xfrm(pic):
    """Return x, y, width, height from a PowerPoint picture."""
    sppr = pic.find("p:spPr", NS)
    if sppr is None:
        return None

    xfrm = sppr.find("a:xfrm", NS)
    if xfrm is None:
        return None

    off = xfrm.find("a:off", NS)
    ext = xfrm.find("a:ext", NS)

    if off is None or ext is None:
        return None

    try:
        return (
            int(off.attrib["x"]),
            int(off.attrib["y"]),
            int(ext.attrib["cx"]),
            int(ext.attrib["cy"]),
        )
    except (KeyError, ValueError):
        return None


def bbox_contains(outer, inner, tolerance=0.02):
    """Whether inner's center is inside outer, with a small tolerance."""
    ox, oy, ow, oh = outer
    ix, iy, iw, ih = inner

    cx = ix + iw / 2
    cy = iy + ih / 2

    tx = ow * tolerance
    ty = oh * tolerance

    return (
        ox - tx <= cx <= ox + ow + tx
        and oy - ty <= cy <= oy + oh + ty
    )


def bbox_area(box):
    return box[2] * box[3]


def parse_relationships(xml_bytes):
    root = ET.fromstring(xml_bytes)
    result = {}

    for rel in root:
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        rel_type = rel.attrib.get("Type", "")
        if rid and target:
            result[rid] = {
                "target": target,
                "type": rel_type,
            }

    return result


def normalize_target(target):
    # Slide relationship targets normally look like ../media/image1.png.
    target = target.replace("\\", "/")
    while target.startswith("../"):
        target = target[3:]
    if not target.startswith("ppt/"):
        target = "ppt/" + target
    return target


def get_picture_info(pic):
    nv = pic.find("p:nvPicPr", NS)
    c_nv_pr = nv.find("p:cNvPr", NS) if nv is not None else None

    name = c_nv_pr.attrib.get("name", "") if c_nv_pr is not None else ""

    blip = pic.find(".//a:blip", NS)
    rid = blip.attrib.get(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
    ) if blip is not None else None

    bbox = parse_xfrm(pic)

    return {
        "name": name,
        "rid": rid,
        "bbox": bbox,
        "pic": pic,
    }


def image_from_zip(zf, target):
    data = zf.read(target)
    return Image.open(io.BytesIO(data)).convert("RGBA")


def safe_filename(text):
    text = re.sub(r"[^\w\-. ]+", "_", text, flags=re.UNICODE)
    text = re.sub(r"\s+", "_", text).strip("._ ")
    return text or "image"


def overlay_markup_on_image(base_img, markup_img, base_box, markup_box):
    """
    Place the markup image on the extracted base image using the same
    relative position/size that the objects have on the PowerPoint slide.

    The markup fallback image has transparent pixels, so alpha compositing
    keeps the original photograph intact.
    """
    bx, by, bw, bh = base_box
    mx, my, mw, mh = markup_box

    if bw <= 0 or bh <= 0 or mw <= 0 or mh <= 0:
        return base_img

    # Relative geometry on the base picture.
    rel_x = (mx - bx) / bw
    rel_y = (my - by) / bh
    rel_w = mw / bw
    rel_h = mh / bh

    # Convert relative geometry to actual pixels.
    px = round(rel_x * base_img.width)
    py = round(rel_y * base_img.height)
    pw = max(1, round(rel_w * base_img.width))
    ph = max(1, round(rel_h * base_img.height))

    # The fallback ink image is intended to fill its PowerPoint bbox.
    markup = markup_img.resize((pw, ph), Image.Resampling.LANCZOS)

    # Crop if the overlay extends slightly outside the photo.
    canvas = Image.new("RGBA", base_img.size, (0, 0, 0, 0))
    canvas.alpha_composite(markup, (px, py))

    return Image.alpha_composite(base_img, canvas)


def process_pptx(uploaded_file, progress_callback=None):
    """
    Extract every photo from every slide and merge any PowerPoint Ink
    fallback image whose center lies over that photo.

    Returns:
        results: list of dicts with slide number, image number, name, bytes
        stats: dict
    """
    ppt_bytes = uploaded_file.getvalue()

    results = []
    total_slides = 0
    markup_count = 0
    image_count = 0
    merged_count = 0

    with zipfile.ZipFile(io.BytesIO(ppt_bytes), "r") as zf:
        slide_names = [
            n for n in zf.namelist()
            if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)
        ]
        slide_names.sort(key=lambda x: int(re.search(r"slide(\d+)", x).group(1)))
        total_slides = len(slide_names)

        for slide_index, slide_path in enumerate(slide_names, start=1):
            root = ET.fromstring(zf.read(slide_path))

            rel_path = (
                "ppt/slides/_rels/"
                + Path(slide_path).name
                + ".rels"
            )

            if rel_path not in zf.namelist():
                continue

            rels = parse_relationships(zf.read(rel_path))

            pics = [
                get_picture_info(pic)
                for pic in root.findall(".//p:pic", NS)
            ]

            # Candidate normal pictures. Ink fallback pictures are excluded.
            normal_pics = [
                p for p in pics
                if p["bbox"] is not None
                and p["rid"] is not None
                and not p["name"].lower().startswith("ink")
            ]

            ink_pics = [
                p for p in pics
                if p["bbox"] is not None
                and p["rid"] is not None
                and p["name"].lower().startswith("ink")
            ]

            markup_count += len(ink_pics)

            # Build a list of markups that belong to each base picture.
            assigned = {id(p): [] for p in normal_pics}

            for ink in ink_pics:
                ink_box = ink["bbox"]

                # Find the smallest normal picture containing the ink center.
                candidates = [
                    p for p in normal_pics
                    if bbox_contains(p["bbox"], ink_box)
                ]

                if candidates:
                    owner = min(candidates, key=lambda p: bbox_area(p["bbox"]))
                    assigned[id(owner)].append(ink)

            output_image_no = 0

            for base in normal_pics:
                rid_info = rels.get(base["rid"])
                if not rid_info:
                    continue

                target = normalize_target(rid_info["target"])

                if target not in zf.namelist():
                    continue

                # Only process actual image relationships.
                if "image" not in rid_info["type"].lower():
                    continue

                try:
                    base_img = image_from_zip(zf, target)
                except Exception:
                    continue

                image_count += 1
                output_image_no += 1

                # Merge all ink objects assigned to this picture.
                output_img = base_img
                assigned_inks = assigned.get(id(base), [])

                for ink in assigned_inks:
                    ink_rel = rels.get(ink["rid"])
                    if not ink_rel:
                        continue

                    ink_target = normalize_target(ink_rel["target"])
                    if ink_target not in zf.namelist():
                        continue

                    try:
                        markup_img = image_from_zip(zf, ink_target)
                        output_img = overlay_markup_on_image(
                            output_img,
                            markup_img,
                            base["bbox"],
                            ink["bbox"],
                        )
                        merged_count += 1
                    except Exception:
                        # Keep the original photo if a particular markup
                        # image cannot be decoded.
                        pass

                buffer = io.BytesIO()
                output_img.convert("RGB").save(
                    buffer,
                    format="PNG",
                    optimize=True,
                )

                results.append({
                    "slide": slide_index,
                    "image": output_image_no,
                    "filename": (
                        f"Slide_{slide_index:03d}_"
                        f"Image_{output_image_no:02d}.png"
                    ),
                    "data": buffer.getvalue(),
                    "markup_count": len(assigned_inks),
                })

            if progress_callback:
                progress_callback(slide_index / max(total_slides, 1))

    return results, {
        "slides": total_slides,
        "images": image_count,
        "markups": markup_count,
        "merged": merged_count,
        "outputs": len(results),
    }


def make_zip(results):
    out = io.BytesIO()

    with zipfile.ZipFile(
        out,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=6,
    ) as z:
        for item in results:
            z.writestr(item["filename"], item["data"])

    return out.getvalue()


st.title("🖼️ PPT Image + Markup Extractor")
st.write(
    "PowerPoint upload करें और photos को उनके ऊपर मौजूद "
    "markup के साथ अलग-अलग PNG images में निकालें."
)

st.info(
    "यह tool PowerPoint के Ink/markup के fallback image को पहचानकर "
    "उसी photo पर सही position में merge करता है."
)

uploaded_file = st.file_uploader(
    "PowerPoint File (.pptx)",
    type=["pptx"],
)

if uploaded_file:
    st.success(f"Selected: {uploaded_file.name}")

    if st.button("🚀 Process PPT", type="primary", use_container_width=True):
        progress = st.progress(0)
        status = st.empty()

        def update_progress(value):
            progress.progress(min(max(value, 0.0), 1.0))
            status.text(
                f"Processing slides... {int(value * 100)}%"
            )

        try:
            with st.spinner("PPT process हो रही है..."):
                results, stats = process_pptx(
                    uploaded_file,
                    progress_callback=update_progress,
                )

            progress.progress(1.0)
            status.empty()

            if not results:
                st.error(
                    "कोई image नहीं मिली. कृपया सुनिश्चित करें कि PPTX "
                    "file valid है."
                )
            else:
                col1, col2, col3 = st.columns(3)
                col1.metric("Slides", stats["slides"])
                col2.metric("Images", stats["images"])
                col3.metric("Markup merged", stats["merged"])

                st.success(
                    f"{len(results)} images तैयार हैं."
                )

                zip_bytes = make_zip(results)

                st.download_button(
                    label="⬇️ Download All Images (ZIP)",
                    data=zip_bytes,
                    file_name="PPT_Images_With_Markup.zip",
                    mime="application/zip",
                    use_container_width=True,
                )

                st.subheader("Preview")

                # Show first few generated images.
                preview_items = results[:6]

                for item in preview_items:
                    st.image(
                        item["data"],
                        caption=(
                            f"{item['filename']} — "
                            f"Markup objects: {item['markup_count']}"
                        ),
                        use_container_width=True,
                    )

                if len(results) > 6:
                    st.caption(
                        f"Preview में पहले 6 images दिखाई गई हैं. "
                        f"ZIP में सभी {len(results)} images हैं."
                    )

        except zipfile.BadZipFile:
            st.error("यह valid PPTX file नहीं लग रही है.")
        except Exception as e:
            st.error(f"Processing error: {e}")
            st.exception(e)
