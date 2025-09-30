"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement, QueueStatus

from robot_framework.process import process
from robot_framework.initialize import initialize
import os
import json
from typing import Optional

from robot_framework.reset import reset

def make_queue_element_with_payload(
    payload: dict | list,
    queue_name: str,
    reference: Optional[str] = None,
    created_by: Optional[str] = None,
    status: QueueStatus = QueueStatus.NEW, 
) -> QueueElement:
    # Validate & serialize
    data_str = json.dumps(payload, ensure_ascii=False)
    if len(data_str) > 2000:
        raise ValueError("data exceeds 2000 chars (column limit)")

    return QueueElement(
        queue_name=queue_name,
        status=status,
        data=data_str,
        reference=reference,
        created_by=created_by,
    )

# pylint: disable-next=unused-argum
orchestrator_connection = OrchestratorConnection(
    "PlannerRefresh",
    os.getenv("OpenOrchestratorSQL"),
    os.getenv("OpenOrchestratorKey"),
    None,
)

ctx = initialize(orchestrator_connection)

qe = make_queue_element_with_payload(
    payload={
        "Name": "Bystrategi_S7 - Penneo",
        "URL": "https://tasks.office.com/aarhuskommune.onmicrosoft.com/da-DK/Home/Planner/#/plantaskboard?groupId=beacbc2b-a179-4b13-92c3-68727e8adc26&planId=r4KQ8i6trUu2h_nB-ObyI5YAAZNF"
    },
    queue_name="PlannerRefresh",
    reference="Sandbox",
    status=QueueStatus.NEW, 
)

reset(orchestrator_connection)

process(orchestrator_connection, qe, ctx)
