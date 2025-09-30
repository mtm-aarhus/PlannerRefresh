"""This module defines any initial processes to run when the robot starts."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from sharepoint import sharepoint_client

def initialize(orchestrator_connection: OrchestratorConnection) -> None:
    """Do all custom startup initializations of the robot."""
    orchestrator_connection.log_trace("Initializing.")
    certification = orchestrator_connection.get_credential("SharePointCert")
    api = orchestrator_connection.get_credential("SharePointAPI")
    
    tenant = api.username
    client_id = api.password
    thumbprint = certification.username
    cert_path = certification.password
    
    sharepoint_site_base = orchestrator_connection.get_constant("AarhusKommuneSharePoint").value
    sharepoint_site = f"{sharepoint_site_base}/teams/PlannerPowerBI"
    
    client = sharepoint_client(tenant, client_id, thumbprint, cert_path, sharepoint_site, orchestrator_connection)


    return client
