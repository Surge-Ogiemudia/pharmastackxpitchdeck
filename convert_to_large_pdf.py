import os
import time
from playwright.sync_api import sync_playwright
from PIL import Image

def generate_large_legible_pdf():
    html_path = os.path.abspath("index.html")
    file_url = f"file:///{html_path.replace('\\', '/')}"
    output_pdf = os.path.abspath("PharmaStackX_PitchDeck_Large.pdf")
    img_dir = os.path.abspath("slide_screenshots_large")
    os.makedirs(img_dir, exist_ok=True)

    print(f"Opening browser to capture Large/Legible HD slides from: {file_url}")

    with sync_playwright() as p:
        # Launch browser with crisp scale factor
        browser = p.chromium.launch(headless=True)
        # Use 1440x810 viewport with 2.5 scale factor for ultra-sharp large text rendering
        page = browser.new_page(viewport={"width": 1440, "height": 810}, device_scale_factor=2.5)
        page.goto(file_url)
        page.wait_for_timeout(1000)

        # Inject custom zoom and padding CSS to expand content & maximize text size
        css_injection = """
            document.getElementById('nav').style.display = 'none';
            const style = document.createElement('style');
            style.innerHTML = `
                body { zoom: 1.32 !important; }
                .slide { padding: 18px 24px 30px !important; }
                .s-title { margin-bottom: 20px !important; font-size: 38px !important; }
            `;
            document.head.appendChild(style);
        """
        page.evaluate(css_injection)

        # Get total slides
        total_slides = page.evaluate("TOTAL")
        print(f"Total slides detected: {total_slides}")

        screenshot_paths = []

        for i in range(1, total_slides + 1):
            page.evaluate(f"goTo({i});")
            page.wait_for_timeout(600)  # Wait for animation transition
            
            img_path = os.path.join(img_dir, f"slide_large_{i}.png")
            page.screenshot(path=img_path, full_page=True)
            screenshot_paths.append(img_path)
            print(f"Captured Large Slide {i}/{total_slides} -> {img_path}")

        browser.close()

    # Compile into PDF
    print("Compiling Large/Legible PDF presentation...")
    images = []
    for img_path in screenshot_paths:
        if os.path.exists(img_path):
            img = Image.open(img_path).convert("RGB")
            images.append(img)

    if images:
        images[0].save(
            output_pdf, "PDF", resolution=300.0, save_all=True, append_images=images[1:]
        )
        print(f"\nSUCCESS! Saved Large/Legible Pitch Deck PDF to:\n{output_pdf}")
    else:
        print("Error: No slide screenshots captured.")

if __name__ == "__main__":
    generate_large_legible_pdf()
