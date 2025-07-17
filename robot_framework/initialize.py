"""This module defines any initial processes to run when the robot starts."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

def sharepoint_client(username: str, password: str, sharepoint_site_url: str, orchestrator_connection: OrchestratorConnection) -> ClientContext:
    """
    Creates and returns a SharePoint client context.
    """
    # Authenticate to SharePoint
    ctx = ClientContext(sharepoint_site_url).with_credentials(UserCredential(username, password))

    # Load and verify connection
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()

    orchestrator_connection.log_info(f"Authenticated successfully. Site Title: {web.properties['Title']}")
    return ctx

def initialize(orchestrator_connection: OrchestratorConnection) -> None:
    """Do all custom startup initializations of the robot."""
    orchestrator_connection.log_trace("Initializing.")
    RobotCredentials = orchestrator_connection.get_credential("Robot365User")
    username = RobotCredentials.username
    password = RobotCredentials.password
    
    
    sharepoint_site_base = orchestrator_connection.get_constant("AarhusKommuneSharePoint").value
    sharepoint_site = f"{sharepoint_site_base}/teams/PlannerPowerBI"

    client = sharepoint_client(username, password, sharepoint_site, orchestrator_connection)
    return client
