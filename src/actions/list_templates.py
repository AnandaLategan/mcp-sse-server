"""
List available Word templates from SharePoint.
"""

import logging

from ..utils import GraphClient

logger = logging.getLogger(__name__)


async def list_templates_action(
    azure_tenant_id: str,
    azure_client_id: str,
    azure_client_secret: str,
    sharepoint_drive_id: str,          # ← changed
    sharepoint_template_folder: str,
) -> str:
    """
    List all available Word document templates stored in SharePoint.
    Call this when the user wants to see what templates are available
    or when starting a new document.

    Returns:
        A list of available template filenames.
    """
    logger.info("Listing SharePoint templates")

    graph = GraphClient(
        tenant_id=azure_tenant_id,
        client_id=azure_client_id,
        client_secret=azure_client_secret,
    )

    files = await graph.list_sharepoint_files(
        drive_id=sharepoint_drive_id,          # ← changed
        folder_path=sharepoint_template_folder,
    )

    if not files:
        return "No templates found in the SharePoint templates folder."

    template_list = "\n".join([f"  - {f['name']}" for f in files])
    return f"Available templates:\n{template_list}"