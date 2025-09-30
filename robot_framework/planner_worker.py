from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
import os, time, sys, argparse

def download_planner_worker(downloads_folder: str, planner_url: str, final_file_path: str) -> None:
    options = Options()
    options.add_argument("--user-data-dir=" + os.path.join(os.getenv("LOCALAPPDATA"), "Microsoft", "Edge", "User Data"))
    options.add_argument("--start-maximized")
    options.add_argument("--disable-extensions")
    options.add_argument("--profile-directory=Default")
    options.add_experimental_option("prefs", {
        "download.default_directory": downloads_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "browser.show_hub_popup_on_download_start": False
    })

    driver = webdriver.Edge(options=options)
    downloaded_file = None
    try:
        driver.get(planner_url)
        wait = WebDriverWait(driver, 20)  # shorter waits, repeat if needed
        wait.until(EC.element_to_be_clickable((By.XPATH, "//i[@data-icon-name='plannerChevronDownSmall']"))).click()
        wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//button[.//span[normalize-space()='Eksportér plan til Excel' or normalize-space()='Export plan to Excel']]"
        ))).click()

        initial = set(os.listdir(downloads_folder))
        start = time.time()
        while True:
            new = [f for f in (set(os.listdir(downloads_folder)) - initial) if f.lower().endswith(".xlsx")]
            if new:
                downloaded_file = os.path.join(downloads_folder, sorted(new)[0])
                break
            if time.time() - start > 60:
                raise TimeoutError("No .xlsx detected within 60s")
            time.sleep(1)

        time.sleep(2)
        os.replace(downloaded_file, final_file_path)
    finally:
        try: driver.quit()
        except Exception: pass

if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("--downloads", required=True)
    p.add_argument("--url", required=True)
    p.add_argument("--out", required=True)
    args = p.parse_args()
    try:
        download_planner_worker(args.downloads, args.url, args.out)
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
