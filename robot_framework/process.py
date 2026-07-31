"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement

from office365.sharepoint.client_context import ClientContext

import subprocess
import sys
import gc

import os
import json
import shutil
import tempfile
from urllib.parse import urlparse

# pylint: disable-next=unused-argument
def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement, client: ClientContext | None = None) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")

    
    data = json.loads(queue_element.data)
     # Assign each field to a named variable

    file_name = f'{data.get("Name")}.xlsx'
    planner_url = normalize_planner_url(data.get("URL"))
    
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(downloads_folder, exist_ok=True)
    worker_downloads_folder = tempfile.mkdtemp(prefix="PlannerRefresh_", dir=downloads_folder)

    final_file_path = os.path.join(downloads_folder, file_name)
    if os.path.exists(final_file_path):
        os.remove(final_file_path)
    
    sharepoint_folder = "Shared Documents/PowerBi"

    try:
        orchestrator_connection.log_info("Initializing download")
        run_planner_subprocess(worker_downloads_folder, planner_url, final_file_path, timeout_s=300,
                            log_info=orchestrator_connection.log_info,
                            log_error=orchestrator_connection.log_error)

        orchestrator_connection.log_info("Uploading file to SharePoint")
        upload_file_to_sharepoint(client, sharepoint_folder, final_file_path, orchestrator_connection)
        if os.path.exists(final_file_path):
            os.remove(final_file_path)
       
    except Exception as ex:
        gc.collect()
        if os.path.exists(final_file_path):
            os.remove(final_file_path)
        raise ex
    finally:
        shutil.rmtree(worker_downloads_folder, ignore_errors=True)


def normalize_planner_url(planner_url: str | None) -> str:
    """Validate and normalize the Planner URL from the queue payload."""
    if not isinstance(planner_url, str) or not planner_url.strip():
        raise ValueError("Queue element data is missing a non-empty 'URL' value")

    planner_url = planner_url.strip()
    parsed_url = urlparse(planner_url)
    if parsed_url.scheme not in ("http", "https") or not parsed_url.netloc:
        raise ValueError(f"Queue element has an invalid Planner URL: {planner_url!r}")

    return planner_url


def upload_file_to_sharepoint(client: ClientContext, sharepoint_file_url: str, local_file_path: str, orchestrator_connection: OrchestratorConnection):
    """
    Uploads the specified local file back to SharePoint at the given URL.
    Uses the folder path directly to upload files.
    """
    # Extract the root folder, folder path, and file name
    path_parts = sharepoint_file_url.split('/')
    DOCUMENT_LIBRARY = path_parts[0]  # Root folder name (document library)
    FOLDER_PATH = path_parts[1]
    file_name = os.path.basename(local_file_path)  # File name

    # Construct the server-relative folder path (starting with the document library)
    if FOLDER_PATH:
        folder_path = f"{DOCUMENT_LIBRARY}/{FOLDER_PATH}"
    else:
        folder_path = f"{DOCUMENT_LIBRARY}"

    # Get the folder where the file should be uploaded
    target_folder = client.web.get_folder_by_server_relative_url(folder_path)
    client.load(target_folder)
    client.execute_query()
    
    orchestrator_connection.log_info("Uploading file")

    # Upload the file to the correct folder in SharePoint
    with open(local_file_path, "rb") as file_content:
        uploaded_file = target_folder.upload_file(file_name, file_content).execute_query()

    orchestrator_connection.log_info(f"[Ok] file has been uploaded to: {uploaded_file.serverRelativeUrl} on SharePoint")


def run_planner_subprocess(downloads_folder, planner_url, final_file_path, timeout_s, log_info, log_error):
    script = os.path.join(os.path.dirname(__file__), "planner_worker.py")
    cmd = [sys.executable, "-u", script,
           "--downloads", downloads_folder,
           "--url", planner_url,
           "--out", final_file_path]
    env = os.environ.copy()
    env.setdefault("PYDEVD_DISABLE_FILE_VALIDATION", "1")

    # Ensure we can kill the whole tree on Windows
    creationflags = subprocess.CREATE_NEW_PROCESS_GROUP
    proc = subprocess.Popen(
        cmd,
        creationflags=creationflags,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
    )

    try:
        stdout, stderr = proc.communicate(timeout=timeout_s)
    except subprocess.TimeoutExpired:
        log_error("Worker timed out; killing process tree")
        # Kill python child and any spawned msedgedriver/msedge
        subprocess.run(f"taskkill /PID {proc.pid} /T /F", shell=True)
        subprocess.run("taskkill /IM msedgedriver.exe /F /T >NUL 2>&1", shell=True)
        subprocess.run("taskkill /IM msedge.exe /F /T >NUL 2>&1", shell=True)
        try:
            stdout, stderr = proc.communicate(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()
            stdout, stderr = proc.communicate()
        log_worker_output(stdout, stderr, log_info, log_error, failed=True)
        raise RuntimeError("download_planner timed out")

    if proc.returncode != 0:
        log_worker_output(stdout, stderr, log_info, log_error, failed=True)
        details = tail(stderr or stdout)
        raise RuntimeError(f"download_planner failed (exit {proc.returncode}): {details}")

    log_worker_output(stdout, stderr, log_info, log_error, failed=False)


def log_worker_output(stdout: str, stderr: str, log_info, log_error, failed: bool) -> None:
    """Forward worker output to OpenOrchestrator logs."""
    if failed and stderr:
        log_error(f"Planner worker stderr output:\n{tail(stderr)}")


def tail(text: str, max_chars: int = 4000) -> str:
    """Keep subprocess errors readable in the orchestrator log."""
    if len(text) <= max_chars:
        return text.strip()
    return text[-max_chars:].strip()
