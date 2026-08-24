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

    print(f"Opening browser to capture ALL slides & sub-steps from: {file_url}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1440, "height": 810}, device_scale_factor=2.5)
        page.goto(file_url)
        page.wait_for_timeout(1000)

        # Inject CSS to hide nav and optimize zoom/padding for maximum legibility
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

        total_slides = page.evaluate("TOTAL")
        print(f"Total main slides detected: {total_slides}")

        screenshot_paths = []
        shot_count = 1

        for i in range(1, total_slides + 1):
            if i == 2:
                # SLIDE 2 HAS 3 CRITICAL SUB-STEPS (SUSAN OBOH, JOY AFIA ASTHMA, WHATSAPP DATA)
                for step in range(1, 4):
                    page.evaluate(f"goTo(2, {step});")
                    page.wait_for_timeout(600)
                    img_path = os.path.join(img_dir, f"slide_large_2_step{step}.png")
                    page.screenshot(path=img_path, full_page=True)
                    screenshot_paths.append(img_path)
                    print(f"Captured Slide 2 Step {step}/3 (Susan / Asthma / WhatsApp) -> {img_path}")
                    shot_count += 1
            else:
                page.evaluate(f"goTo({i});")
                page.wait_for_timeout(600)
                img_path = os.path.join(img_dir, f"slide_large_{i}.png")
                page.screenshot(path=img_path, full_page=True)
                screenshot_paths.append(img_path)
                print(f"Captured Slide {i}/{total_slides} -> {img_path}")
                shot_count += 1

        browser.close()

    # Compile into PDF
    print(f"Compiling ALL {len(screenshot_paths)} slides & sub-steps into PDF...")
    images = []
    for img_path in screenshot_paths:
        if os.path.exists(img_path):
            img = Image.open(img_path).convert("RGB")
            images.append(img)

    if images:
        images[0].save(
            output_pdf, "PDF", resolution=300.0, save_all=True, append_images=images[1:]
        )
        print(f"\nSUCCESS! Saved Complete PDF (including Asthma & WhatsApp slides) to:\n{output_pdf}")
    else:
        print("Error: No slide screenshots captured.")

if __name__ == "__main__":
    generate_large_legible_pdf()
