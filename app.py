import os
import subprocess
from google.colab import files
from pdf2image import convert_from_path

# 1. PPT File Upload karein
print("Apni PPT file upload karein:")
uploaded = files.upload()
ppt_file = list(uploaded.keys())[0]

# 2. PPT ko PDF mein convert karein (Automatic Merging)
print("PPT ko PDF mein render kiya ja raha hai...")
subprocess.run(
    ["libreoffice", "--headless", "--convert-to", "pdf", ppt_file], check=True
)
pdf_file = os.path.splitext(ppt_file)[0] + ".pdf"

# 3. PDF Pages ka high-quality Screenshot + Auto-Crop karein
print("Slides ke screenshots aur crop process ho rahe hain...")
images = convert_from_path(pdf_file, dpi=200)

os.makedirs("cropped_slides", exist_ok=True)

for i, image in enumerate(images):
    # Auto-Crop: Safed margins/borders ko automatic crop kar deta hai
    bbox = image.getbbox()  # Get content bounding box
    if bbox:
        cropped_image = image.crop(bbox)
    else:
        cropped_image = image

    # Save final marked image
    output_path = f"cropped_slides/slide_{i+1}_marked.png"
    cropped_image.save(output_path, "PNG")
    print(f"Saved: {output_path}")

# 4. ZIP banakar download karein
subprocess.run(["zip", "-r", "marked_slides.zip", "cropped_slides"])
files.download("marked_slides.zip")
print("Download complete!")
