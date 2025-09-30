"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement

from office365.sharepoint.client_context import ClientContext

import subprocess, sys
import gc

import os
import json

# pylint: disable-next=unused-argument
def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement, client: ClientContext | None = None) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")

    
    data = json.loads(queue_element.data)
     # Assign each field to a named variable

    file_name = f'{data.get("Name")}.xlsx'
    planner_url = data.get("URL")
    
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")

    final_file_path = os.path.join(downloads_folder, file_name)
    if os.path.exists(final_file_path):
        os.remove(final_file_path)
    
    sharepoint_folder = "Shared Documents/PowerBi"

    try:
        orchestrator_connection.log_info("Initializing download")
        run_planner_subprocess(downloads_folder, planner_url, final_file_path, timeout_s=300,
                            log=orchestrator_connection.log_error)

        orchestrator_connection.log_info("Uploading file to SharePoint")
        upload_file_to_sharepoint(client, sharepoint_folder, final_file_path, orchestrator_connection)
        if os.path.exists(final_file_path):
            os.remove(final_file_path)
       
    except Exception as ex:
        gc.collect()
        if os.path.exists(final_file_path):
            os.remove(final_file_path)
        raise ex


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


def run_planner_subprocess(downloads_folder, planner_url, final_file_path, timeout_s, log):
    script = os.path.join(os.path.dirname(__file__), "planner_worker.py")
    cmd = [sys.executable, "-u", script,
           "--downloads", downloads_folder,
           "--url", planner_url,
           "--out", final_file_path]

    # Ensure we can kill the whole tree on Windows
    creationflags = subprocess.CREATE_NEW_PROCESS_GROUP
    proc = subprocess.Popen(cmd, creationflags=creationflags)

    try:
        proc.wait(timeout=timeout_s)
    except subprocess.TimeoutExpired:
        log("Worker timed out; killing process tree")
        # Kill python child and any spawned msedgedriver/msedge
        subprocess.run(f"taskkill /PID {proc.pid} /T /F", shell=True)
        subprocess.run("taskkill /IM msedgedriver.exe /F /T >NUL 2>&1", shell=True)
        subprocess.run("taskkill /IM msedge.exe /F /T >NUL 2>&1", shell=True)
        raise RuntimeError("download_planner timed out")

    if proc.returncode != 0:
        raise RuntimeError(f"download_planner failed (exit {proc.returncode})")