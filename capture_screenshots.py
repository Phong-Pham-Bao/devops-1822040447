import os
import sys
from playwright.sync_api import sync_playwright

output_dir = os.path.abspath(r"1822040447_PhamBaoPhong_TKPM\images")
os.makedirs(output_dir, exist_ok=True)

targets = [
    {
        "url": "https://phong-pham-bao.github.io/devops-1822040447/",
        "filename": "img_01_website_live.png",
        "viewport": {"width": 1280, "height": 850}
    },
    {
        "url": "https://github.com/Phong-Pham-Bao/devops-1822040447",
        "filename": "img_02_github_repo.png",
        "viewport": {"width": 1280, "height": 850}
    },
    {
        "url": "https://github.com/Phong-Pham-Bao/devops-1822040447/actions",
        "filename": "img_03_actions_list.png",
        "viewport": {"width": 1280, "height": 850}
    },
    {
        "url": "https://github.com/Phong-Pham-Bao/devops-1822040447/actions/runs/41498855062",
        "filename": "img_04_pipeline_failed.png",
        "viewport": {"width": 1280, "height": 850}
    },
    {
        "url": "https://github.com/Phong-Pham-Bao/devops-1822040447/actions/runs/41499011709",
        "filename": "img_05_pipeline_success.png",
        "viewport": {"width": 1280, "height": 850}
    }
]

with sync_playwright() as p:
    browser = p.chromium.launch()
    for item in targets:
        filepath = os.path.join(output_dir, item["filename"])
        print(f"Capturing {item['url']} -> {item['filename']}...")
        page = browser.new_page(viewport=item["viewport"])
        try:
            page.goto(item["url"], wait_until="networkidle", timeout=30000)
            page.wait_for_timeout(2000)
            page.screenshot(path=filepath, full_page=False)
            print(f"Saved: {filepath}")
        except Exception as e:
            print(f"Error capturing {item['url']}: {e}")
        finally:
            page.close()
    browser.close()

print("All screenshots captured successfully!")
