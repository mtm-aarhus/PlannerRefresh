from selenium import webdriver
from selenium.common.exceptions import InvalidSessionIdException, NoSuchWindowException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
import argparse
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path
from urllib.parse import parse_qs, urlparse

try:
    import winreg
except ImportError:
    winreg = None


APP_NAME = "PlannerRefresh"


def get_local_app_data_dir() -> str:
    """Return the current user's local app data directory."""
    local_app_data = os.getenv("LOCALAPPDATA")
    if local_app_data:
        return local_app_data

    return os.path.join(os.path.expanduser("~"), "AppData", "Local")


def get_edge_user_data_dir() -> str:
    """Return the Edge profile root used by Selenium."""
    configured_dir = os.getenv("PLANNER_EDGE_USER_DATA_DIR") or os.getenv("EDGE_USER_DATA_DIR")
    if configured_dir:
        return os.path.expandvars(os.path.expanduser(configured_dir))

    app_data_dir = os.path.join(get_local_app_data_dir(), APP_NAME)
    automation_dir = os.path.join(app_data_dir, "EdgeAutomation")
    previous_dir = os.path.join(app_data_dir, "EdgeUserData")
    # Preserve profiles already used by older versions of the worker.
    if not os.path.exists(automation_dir) and os.path.isfile(os.path.join(previous_dir, "Local State")):
        return previous_dir

    return automation_dir


def get_edge_binary_path() -> str:
    """Find msedge.exe for logging and explicit Selenium configuration."""
    configured_path = os.getenv("PLANNER_EDGE_BINARY")
    if configured_path:
        return os.path.expandvars(os.path.expanduser(configured_path))

    candidate_paths = [
        os.path.join(os.getenv("ProgramFiles(x86)", ""), "Microsoft", "Edge", "Application", "msedge.exe"),
        os.path.join(os.getenv("ProgramFiles", ""), "Microsoft", "Edge", "Application", "msedge.exe"),
        os.path.join(get_local_app_data_dir(), "Microsoft", "Edge", "Application", "msedge.exe"),
    ]
    for candidate_path in candidate_paths:
        if candidate_path and os.path.exists(candidate_path):
            return candidate_path

    return shutil.which("msedge") or ""


def is_truthy(value: str | None) -> bool:
    """Return whether an environment setting should be treated as enabled."""
    return value is not None and value.strip().lower() in ("1", "true", "yes", "on")


def get_edge_policy_values(policy_name: str) -> list[tuple[str, object]]:
    """Read Edge policy values that commonly block WebDriver startup."""
    if winreg is None:
        return []

    values = []
    seen = set()
    key_path = r"SOFTWARE\Policies\Microsoft\Edge"
    hives = (
        ("HKLM", winreg.HKEY_LOCAL_MACHINE),
        ("HKCU", winreg.HKEY_CURRENT_USER),
    )
    access_modes = [("default", winreg.KEY_READ)]
    if hasattr(winreg, "KEY_WOW64_64KEY"):
        access_modes.append(("64-bit", winreg.KEY_READ | winreg.KEY_WOW64_64KEY))
    if hasattr(winreg, "KEY_WOW64_32KEY"):
        access_modes.append(("32-bit", winreg.KEY_READ | winreg.KEY_WOW64_32KEY))

    for hive_name, hive in hives:
        for access_name, access_mode in access_modes:
            try:
                with winreg.OpenKey(hive, key_path, 0, access_mode) as key:
                    value, _ = winreg.QueryValueEx(key, policy_name)
            except FileNotFoundError:
                continue
            except OSError:
                continue

            identity = (hive_name, value)
            if identity in seen:
                continue

            seen.add(identity)
            values.append((f"{hive_name} {access_name}", value))

    return values


def assert_edge_webdriver_allowed_by_policy() -> None:
    """Fail clearly if Microsoft Edge policy blocks WebDriver's DevTools connection."""
    for hive_name, value in get_edge_policy_values("DeveloperToolsAvailability"):
        try:
            policy_value = int(value)
        except (TypeError, ValueError):
            continue

        if policy_value == 2:
            raise RuntimeError(
                "Microsoft Edge WebDriver is blocked by policy: "
                f"{hive_name}\\SOFTWARE\\Policies\\Microsoft\\Edge "
                "DeveloperToolsAvailability=2. Set it to 0 or 1 for the robot account/server."
            )

    for hive_name, value in get_edge_policy_values("ProfilePickerOnStartupAvailability"):
        try:
            policy_value = int(value)
        except (TypeError, ValueError):
            continue

        if policy_value == 2:
            raise RuntimeError(
                "Microsoft Edge profile picker is forced by policy: "
                f"{hive_name}\\SOFTWARE\\Policies\\Microsoft\\Edge "
                "ProfilePickerOnStartupAvailability=2. Set it to 0 or 1 for the robot account/server."
            )


def validate_planner_url(planner_url: str) -> str:
    """Fail before Edge starts if the queue supplied an unusable URL."""
    if not isinstance(planner_url, str) or not planner_url.strip():
        raise ValueError("Planner URL is empty")

    planner_url = planner_url.strip()
    parsed_url = urlparse(planner_url)
    if parsed_url.scheme not in ("http", "https") or not parsed_url.netloc:
        raise ValueError(f"Planner URL is invalid: {planner_url!r}")

    return planner_url


def fragment_has_plan_identity(fragment: str) -> bool:
    """Return whether a Planner fragment points at a specific plan."""
    if "?" not in fragment:
        return False

    query_string = fragment.split("?", 1)[1]
    query_params = parse_qs(query_string)
    return bool(query_params.get("groupId") or query_params.get("planId"))


def restore_planner_fragment_after_redirect(driver, original_url: str) -> None:
    """Reapply Planner's plan fragment when Microsoft's redirect drops it."""
    original = urlparse(original_url)
    current = urlparse(driver.current_url)

    if current.hostname not in (original.hostname, "planner.cloud.microsoft"):
        return
    if not original.fragment or not fragment_has_plan_identity(original.fragment):
        return

    if fragment_has_plan_identity(current.fragment):
        return

    restored_url = current._replace(fragment=original.fragment).geturl()
    driver.get(restored_url)


def log_page_diagnostics(driver, context: str) -> None:
    """Log enough browser state to make blank Selenium wait errors diagnosable."""
    try:
        page = urlparse(driver.current_url)
        print(f"{context} current page: {page._replace(query='', fragment='').geturl()}", file=sys.stderr, flush=True)
    except Exception:
        pass

    try:
        print(f"{context} title: {driver.title}", file=sys.stderr, flush=True)
    except Exception:
        pass


def is_edge_profile_sign_in(url: str) -> bool:
    """Identify browser profile sign-in windows before navigating."""
    page = urlparse(url)
    return page.scheme in ("edge", "chrome") and page.hostname in ("force-signin", "profile-picker")


def open_planner_in_ready_window(driver, planner_url: str) -> None:
    """Wait for a browser window that survives Edge's startup transitions."""
    deadline = time.monotonic() + 30
    last_error = None
    saw_profile_sign_in = False
    try:
        while time.monotonic() < deadline:
            for handle in driver.window_handles:
                try:
                    driver.switch_to.window(handle)
                    if is_edge_profile_sign_in(driver.current_url):
                        saw_profile_sign_in = True
                        continue
                    driver.get(planner_url)
                    if not is_edge_profile_sign_in(driver.current_url):
                        return
                    saw_profile_sign_in = True
                except NoSuchWindowException as error:
                    last_error = error
                except WebDriverException as error:
                    message = (error.msg or "").lower()
                    if "target frame detached" not in message and "unable to receive message from renderer" not in message:
                        raise
                    last_error = error
            time.sleep(0.5)
    except InvalidSessionIdException as error:
        if saw_profile_sign_in:
            raise RuntimeError("Edge closed its session while waiting for browser profile sign-in.") from error
        raise

    if saw_profile_sign_in:
        raise RuntimeError(
            "Edge did not finish browser profile sign-in within 30 seconds. "
            "Check implicit sign-in for the robot's Windows session."
        ) from last_error
    raise RuntimeError("Edge did not provide a stable browser window for Planner.") from last_error


def read_devtools_endpoint(port_file: Path) -> tuple[int, str] | None:
    """Read the endpoint Edge writes for this browser instance."""
    try:
        lines = port_file.read_text(encoding="utf-8").splitlines()
        port = int(lines[0])
    except (OSError, ValueError, IndexError):
        return None
    if 0 < port <= 65535 and len(lines) >= 2 and lines[1].startswith("/devtools/browser/"):
        return port, lines[1]
    return None


def wait_for_devtools_endpoint(
    browser: subprocess.Popen, port_file: Path, previous_endpoint: tuple[int, str] | None
) -> str:
    """Wait for the new Edge process to publish a fresh loopback endpoint."""
    deadline = time.monotonic() + 30
    while time.monotonic() < deadline:
        if browser.poll() is not None:
            raise RuntimeError(
                "Microsoft Edge exited before it could accept WebDriver. "
                "Close any existing Edge window using the robot's profile before running."
            )
        endpoint = read_devtools_endpoint(port_file)
        if endpoint is not None and endpoint != previous_endpoint:
            return f"127.0.0.1:{endpoint[0]}"
        time.sleep(0.25)
    raise RuntimeError(
        "Microsoft Edge did not publish a new DevTools endpoint within 30 seconds. "
        "Check the effective Edge remote debugging and user data directory policies."
    )


def download_planner_worker(downloads_folder: str, planner_url: str, final_file_path: str) -> None:
    planner_url = validate_planner_url(planner_url)
    os.makedirs(downloads_folder, exist_ok=True)

    edge_user_data_dir = get_edge_user_data_dir()
    edge_profile_directory = os.getenv("PLANNER_EDGE_PROFILE_DIRECTORY", "").strip()
    edge_binary_path = get_edge_binary_path()
    if not edge_binary_path:
        raise RuntimeError("Could not find Microsoft Edge to launch the robot's browser profile.")

    assert_edge_webdriver_allowed_by_policy()
    os.makedirs(edge_user_data_dir, exist_ok=True)
    port_file = Path(edge_user_data_dir) / "DevToolsActivePort"
    previous_endpoint = read_devtools_endpoint(port_file)
    command = [edge_binary_path, f"--user-data-dir={edge_user_data_dir}", "--remote-debugging-port=0"]
    if edge_profile_directory:
        command.append(f"--profile-directory={edge_profile_directory}")
    if is_truthy(os.getenv("PLANNER_EDGE_HEADLESS")):
        command.extend(["--headless=new", "--window-size=1920,1080"])

    browser = subprocess.Popen(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    driver = None
    try:
        options = Options()
        options.binary_location = edge_binary_path
        options.debugger_address = wait_for_devtools_endpoint(browser, port_file, previous_endpoint)
        log_path = os.path.join(edge_user_data_dir, "msedgedriver.log")
        service = Service(service_args=["--verbose"], log_output=log_path)
        print(f"EdgeDriver log file {log_path}", file=sys.stderr, flush=True)
        try:
            driver = webdriver.Edge(options=options, service=service)
        except WebDriverException as error:
            raise RuntimeError(f"EdgeDriver could not attach to Microsoft Edge: {error.msg}") from error

        driver.set_page_load_timeout(60)
        try:
            open_planner_in_ready_window(driver, planner_url)
        except WebDriverException as error:
            raise RuntimeError(f"Edge lost its browser window during Planner navigation: {error.msg}") from error
        if driver.current_url in ("about:blank", "data:,"):
            raise RuntimeError("Edge stayed on a blank page after navigating to the Planner URL")

        restore_planner_fragment_after_redirect(driver, planner_url)

        wait = WebDriverWait(driver, 60)
        try:
            # Open the menu (handles both Danish + English)
            wait.until(EC.element_to_be_clickable((
                By.XPATH,
                "//button[@aria-haspopup='true' and (contains(@aria-label, 'Plan options') or contains(@aria-label, 'Planindstillinger'))]"
            ))).click()

            # Click export (handles variations like Eksporter/Export)
            export_button = wait.until(EC.element_to_be_clickable((
                By.XPATH,
                "//*[@role='menuitem' and (contains(@aria-label, 'Export') or contains(@aria-label, 'Eksport') or contains(., 'Export') or contains(., 'Eksport'))]"
            )))
        except Exception:
            log_page_diagnostics(driver, "Planner wait failed")
            raise

        driver.execute_cdp_cmd("Browser.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": os.path.abspath(downloads_folder),
        })
        initial = set(os.listdir(downloads_folder))
        export_button.click()
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
        try:
            if driver is not None:
                try:
                    driver.execute_cdp_cmd("Browser.close", {})
                except WebDriverException:
                    pass
                finally:
                    driver.quit()
        finally:
            try:
                browser.wait(timeout=5)
            except subprocess.TimeoutExpired:
                browser.terminate()
                browser.wait(timeout=5)

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
