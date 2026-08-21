import os
from PIL import Image

def generate_hd_pdf():
    img_dir = os.path.abspath("slide_screenshots")
    output_pdf = os.path.abspath("PharmaStackX_PitchDeck_HD.pdf")

    images = []
    for i in range(1, 12):
        img_path = os.path.join(img_dir, f"slide_{i}.png")
        if os.path.exists(img_path):
            img = Image.open(img_path).convert("RGB")
            images.append(img)

    if images:
        images[0].save(
            output_pdf, "PDF", resolution=300.0, save_all=True, append_images=images[1:]
        )
        print(f"SUCCESS! Saved HD Web-Identical PDF to:\n{output_pdf}")
    else:
        print("Error: No slide screenshots found.")

if __name__ == "__main__":
    generate_hd_pdf()
