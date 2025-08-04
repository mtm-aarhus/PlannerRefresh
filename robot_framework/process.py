"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options

from office365.sharepoint.client_context import ClientContext

import os
import time
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
        download_planner(downloads_folder, planner_url, final_file_path, orchestrator_connection)
        orchestrator_connection.log_info("Uploading file to SharePoint")
        upload_file_to_sharepoint(client, sharepoint_folder, final_file_path, orchestrator_connection)
        if os.path.exists(final_file_path):
            os.remove(final_file_path)
    except:
        try:
            os.remove(final_file_path)
        except FileNotFoundError as e:
            print(f"Error: {e}")
        raise
    

def download_planner(downloads_folder, planner_url, final_file_path, orchestrator_connection: OrchestratorConnection):
    # Set up Edge options
    options = Options()
    options.add_argument("--user-data-dir=" + os.path.join(os.getenv("LOCALAPPDATA"), "Microsoft", "Edge", "User Data"))
    options.add_argument("--start-maximized")
    options.add_argument("--disable-extensions")
    options.add_argument("--profile-directory=Default")
    # options.add_argument("--remote-debugging-port=9222")

    prefs = {
        "download.default_directory": downloads_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "browser.show_hub_popup_on_download_start": False
    }
    options.add_experimental_option("prefs", prefs)

    # Initialize Edge WebDriver
    driver = webdriver.Edge(options=options)
    orchestrator_connection.log_info('Driver initialized')
    try:
        # Navigate to Planner URL
        driver.get(planner_url)
             
        orchestrator_connection.log_info("Waiting for dropdown to appear")

        # Wait for the first element to load and interact with it
        wait = WebDriverWait(driver, 60)
        first_element = wait.until(EC.presence_of_element_located((By.XPATH, "//i[@data-icon-name='plannerChevronDownSmall']")))
        first_element.click()
        
        orchestrator_connection.log_info("Waiting for export button to appear")

        # Wait for the second element and click the export button
        export_button = wait.until(EC.presence_of_element_located((By.XPATH, "//button[.//span[text()='Eksportér plan til Excel' or text()='Export plan to Excel']]")))
        export_button.click()

        # Wait for download to complete
        initial_files = set(os.listdir(downloads_folder))
        timeout = 60
        start_time = time.time()
        
        orchestrator_connection.log_info("Waiting for download")

        while True:
            # Get the current list of files
            current_files = set(os.listdir(downloads_folder))
            new_files = current_files - initial_files
            
            # Check if new files have been added
            if new_files:
                # Filter for .xlsx files among the new files
                xlsx_files = [file for file in new_files if file.lower().endswith(".xlsx")]
                if xlsx_files:
                    downloaded_file = os.path.join(downloads_folder, xlsx_files[0])
                    orchestrator_connection.log_info(f"Download completed: {downloaded_file}")
                    break
            
            # Check for timeout
            if time.time() - start_time > timeout:
                orchestrator_connection.log_info("Timeout reached while waiting for a download.")
                break
            
            time.sleep(1)  # Avoid hammering the file system

        os.rename(downloaded_file, final_file_path)

    except:
        try:
            os.remove(final_file_path)
        except FileNotFoundError as e:
             orchestrator_connection.log_info(f"Tried removing downloaded file, didn't exist: {e}")
        driver.quit()
        raise
        


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
