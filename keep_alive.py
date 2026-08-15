import sys
from playwright.sync_api import sync_playwright

def keep_alive(url):
    with sync_playwright() as p:
        # Launch a headless Chromium browser
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        
        print(f"Navigating to {url}...")
        page.goto(url, wait_until="networkidle", timeout=60000)
        
        # Check if the "wake up" button exists (in case it fell asleep)
        wake_button = page.query_selector('button:has-text("Yes, get this app back up!")')
        if wake_button:
            print("App was asleep! Clicking wake-up button...")
            wake_button.click()
            page.wait_for_timeout(10000)
        else:
            print("App is awake and active.")
            
        browser.close()

if __name__ == "__main__":
    APP_URL = "https://professional-cv-generator.streamlit.app"
    keep_alive(APP_URL)
