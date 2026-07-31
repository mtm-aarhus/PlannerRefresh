from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
import argparse
import json
import os
import shutil
import sys
import time
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

    if is_edge_browser_signin_forced():
        return os.path.join(get_local_app_data_dir(), "Microsoft", "Edge", "User Data")

    return os.path.join(get_local_app_data_dir(), APP_NAME, "EdgeUserData")


def is_valid_profile_directory(profile_directory: object) -> bool:
    """Return whether a profile directory value is safe to pass to Edge."""
    if not isinstance(profile_directory, str) or not profile_directory.strip():
        return False

    profile_directory = profile_directory.strip()
    if os.path.isabs(profile_directory):
        return False

    return os.path.basename(profile_directory) == profile_directory


def get_last_used_edge_profile_directory(edge_user_data_dir: str) -> str | None:
    """Read Edge's last-used profile directory from Local State."""
    local_state_path = os.path.join(edge_user_data_dir, "Local State")
    if not os.path.exists(local_state_path):
        return None

    try:
        with open(local_state_path, "r", encoding="utf-8") as file:
            local_state = json.load(file)
    except (OSError, json.JSONDecodeError):
        return None

    profile_state = local_state.get("profile", {})
    info_cache = profile_state.get("info_cache") or {}

    candidates = [
        profile_state.get("last_used"),
        *(profile_state.get("last_active_profiles") or []),
    ]
    for profile_directory, profile_info in info_cache.items():
        if profile_info.get("user_name") and not profile_info.get("signin_required"):
            candidates.append(profile_directory)

    for profile_directory in candidates:
        profile_info = info_cache.get(profile_directory, {})
        if is_valid_profile_directory(profile_directory) and not profile_info.get("signin_required"):
            return profile_directory.strip()

    return None


def get_edge_profile_directory(edge_user_data_dir: str) -> str:
    """Return a specific profile directory if the caller requested one."""
    configured_profile = os.getenv("PLANNER_EDGE_PROFILE_DIRECTORY")
    if configured_profile and configured_profile.strip():
        return configured_profile.strip()

    return get_last_used_edge_profile_directory(edge_user_data_dir) or "Default"


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


def get_edge_policy_int(policy_name: str) -> int | None:
    """Return the first integer value found for an Edge policy."""
    for _, value in get_edge_policy_values(policy_name):
        try:
            return int(value)
        except (TypeError, ValueError):
            continue

    return None


def is_edge_browser_signin_forced() -> bool:
    """Return whether Edge policy requires browser profile sign-in."""
    return get_edge_policy_int("BrowserSignin") == 2 or get_edge_policy_int("ForceBrowserSignin") == 1


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

    if not original.fragment or not fragment_has_plan_identity(original.fragment):
        return

    if fragment_has_plan_identity(current.fragment):
        return

    restored_url = current._replace(fragment=original.fragment).geturl()
    driver.get(restored_url)


def log_page_diagnostics(driver, context: str) -> None:
    """Log enough browser state to make blank Selenium wait errors diagnosable."""
    try:
        print(f"{context} current URL: {driver.current_url}", file=sys.stderr, flush=True)
    except Exception:
        pass

    try:
        print(f"{context} title: {driver.title}", file=sys.stderr, flush=True)
    except Exception:
        pass


def raise_if_edge_profile_picker(driver) -> None:
    """Fail clearly if Edge blocks automation behind its profile sign-in picker."""
    try:
        current_url = driver.current_url
    except Exception:
        return

    if "profile-picker" not in current_url:
        return

    raise RuntimeError(
        "Microsoft Edge opened the profile sign-in picker instead of Planner. "
        "This machine appears to require browser profile sign-in. Sign in once "
        "to Edge with the robot's Windows account, or set PLANNER_EDGE_USER_DATA_DIR "
        "and PLANNER_EDGE_PROFILE_DIRECTORY to an already signed-in Edge profile."
    )


def newest_completed_xlsx(downloads_folder: str, since: float) -> str | None:
    """Return the newest completed Excel file downloaded after the export click."""
    candidates = []
    for file_name in os.listdir(downloads_folder):
        file_path = os.path.join(downloads_folder, file_name)
        if not os.path.isfile(file_path):
            continue
        if not file_name.lower().endswith(".xlsx"):
            continue
        if os.path.getmtime(file_path) < since - 1:
            continue

        candidates.append((os.path.getmtime(file_path), file_path))

    if not candidates:
        return None

    return max(candidates)[1]


def has_active_download(downloads_folder: str) -> bool:
    """Return whether Edge still has an incomplete download in the folder."""
    active_extensions = (".crdownload", ".tmp")
    return any(file_name.lower().endswith(active_extensions) for file_name in os.listdir(downloads_folder))


def wait_for_completed_xlsx(downloads_folder: str, since: float, timeout_s: int = 60) -> str:
    """Wait until Edge has produced a stable .xlsx file."""
    deadline = time.time() + timeout_s
    last_seen_sizes = {}

    while time.time() < deadline:
        downloaded_file = newest_completed_xlsx(downloads_folder, since)
        if downloaded_file and not has_active_download(downloads_folder):
            file_size = os.path.getsize(downloaded_file)
            if file_size > 0 and last_seen_sizes.get(downloaded_file) == file_size:
                return downloaded_file

            last_seen_sizes[downloaded_file] = file_size

        time.sleep(0.25)

    raise TimeoutError("No completed .xlsx detected within 60s")


def move_file_when_available(source_path: str, target_path: str, timeout_s: int = 10) -> None:
    """Move the downloaded file as soon as Edge releases the file handle."""
    deadline = time.time() + timeout_s
    while True:
        try:
            os.replace(source_path, target_path)
            return
        except PermissionError:
            if time.time() >= deadline:
                raise
            time.sleep(0.25)


def configure_edge_startup_profile(edge_user_data_dir: str, edge_profile_directory: str) -> None:
    """Prefer the automation profile and suppress Edge's profile picker."""
    if not edge_profile_directory:
        return

    os.makedirs(os.path.join(edge_user_data_dir, edge_profile_directory), exist_ok=True)
    local_state_path = os.path.join(edge_user_data_dir, "Local State")

    try:
        if os.path.exists(local_state_path):
            with open(local_state_path, "r", encoding="utf-8") as file:
                local_state = json.load(file)
        else:
            local_state = {}

        profile_state = local_state.setdefault("profile", {})
        profile_state["last_used"] = edge_profile_directory
        profile_state["last_active_profiles"] = [edge_profile_directory]
        profile_state["picker_shown"] = True
        profile_state["show_picker_on_startup"] = False
        profile_state["profile_counts_reported"] = "1"

        profile_picker_state = local_state.setdefault("profile_picker", {})
        profile_picker_state["enabled"] = False

        temp_path = local_state_path + ".tmp"
        with open(temp_path, "w", encoding="utf-8") as file:
            json.dump(local_state, file, separators=(",", ":"))
        os.replace(temp_path, local_state_path)
    except (OSError, json.JSONDecodeError):
        return


def edge_local_state_options(edge_profile_directory: str) -> dict:
    """Return Edge local state preferences for deterministic profile startup."""
    return {
        "profile": {
            "last_used": edge_profile_directory,
            "last_active_profiles": [edge_profile_directory],
            "picker_shown": True,
            "show_picker_on_startup": False,
        },
        "profile_picker": {
            "enabled": False,
        },
    }


def download_planner_worker(downloads_folder: str, planner_url: str, final_file_path: str) -> None:
    planner_url = validate_planner_url(planner_url)
    os.makedirs(downloads_folder, exist_ok=True)

    options = Options()
    edge_user_data_dir = get_edge_user_data_dir()
    edge_profile_directory = get_edge_profile_directory(edge_user_data_dir)
    edge_binary_path = get_edge_binary_path()

    os.makedirs(edge_user_data_dir, exist_ok=True)
    configure_edge_startup_profile(edge_user_data_dir, edge_profile_directory)

    if edge_binary_path:
        options.binary_location = edge_binary_path
    options.add_argument("--user-data-dir=" + edge_user_data_dir)
    if edge_profile_directory:
        options.add_argument("--profile-directory=" + edge_profile_directory)
    options.add_argument(planner_url)
    options.add_argument("--start-maximized")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-features=CalculateNativeWinOcclusion,EnableProfilePickerOnStartup")
    options.add_argument("--disable-backgrounding-occluded-windows")
    options.add_argument("--disable-renderer-backgrounding")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    if is_truthy(os.getenv("PLANNER_EDGE_HEADLESS")):
        options.add_argument("--headless=new")
    options.add_experimental_option("localState", edge_local_state_options(edge_profile_directory))
    options.add_experimental_option("prefs", {
        "download.default_directory": downloads_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "browser.show_hub_popup_on_download_start": False
    })

    assert_edge_webdriver_allowed_by_policy()
    driver = webdriver.Edge(options=options)

    downloaded_file = None
    try:
        driver.set_page_load_timeout(60)
        raise_if_edge_profile_picker(driver)
        driver.get(planner_url)
        raise_if_edge_profile_picker(driver)
        if driver.current_url in ("about:blank", "data:,"):
            raise RuntimeError("Edge stayed on a blank page after navigating to the Planner URL")

        restore_planner_fragment_after_redirect(driver, planner_url)

        wait = WebDriverWait(driver, 45)
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

        download_started_at = time.time()
        export_button.click()
        downloaded_file = wait_for_completed_xlsx(downloads_folder, download_started_at)
        move_file_when_available(downloaded_file, final_file_path)
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
