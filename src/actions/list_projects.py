"""
List existing document projects from SharePoint output folder.
"""

import logging

from ..utils import GraphClient, get_version_number

logger = logging.getLogger(__name__)


async def list_projects_action(
    azure_tenant_id: str,
    azure_client_id: str,
    azure_client_secret: str,
    sharepoint_drive_id: str,              # ← changed
    sharepoint_output_folder: str,         # ← changed
) -> str:
    """
    List all existing document projects and their latest version from SharePoint.
    Always call this at the start of a conversation to check for existing work
    before asking the user what they want to do.
    A project is a collection of versioned documents for a specific client or purpose
    e.g. tender_BSCGlobal, tender_NedBank.

    Returns:
        A list of document projects with their latest version number,
        or a message indicating no projects exist yet.
    """
    logger.info("Listing document projects from SharePoint")

    graph = GraphClient(
        tenant_id=azure_tenant_id,
        client_id=azure_client_id,
        client_secret=azure_client_secret,
    )

    # List all subfolders in the output folder
    items = await graph.list_sharepoint_folder(            # ← changed
        drive_id=sharepoint_drive_id,                      # ← changed
        folder_path=sharepoint_output_folder,
    )

    if not items:
        return (
            "No existing document projects found. "
            "All documents will start fresh from a template."
        )

    # Each subfolder is a project
    projects = [item["name"] for item in items if item["is_folder"]]

    if not projects:
        return (
            "No existing document projects found. "
            "All documents will start fresh from a template."
        )

    # For each project find the latest version
    result = "Existing document projects:\n"
    for project in sorted(projects):
        project_items = await graph.list_sharepoint_folder(    # ← changed
            drive_id=sharepoint_drive_id,                      # ← changed
            folder_path=f"{sharepoint_output_folder}/{project}",
        )
        docx_files = [
            item["name"] for item in project_items
            if not item["is_folder"] and item["name"].endswith(".docx")
        ]

        if docx_files:
            latest = max(docx_files, key=get_version_number)
            version = get_version_number(latest)
            result += f"  - {project} (latest: v{str(version).zfill(2)})\n"
        else:
            result += f"  - {project} (no versions yet)\n"

    result += "\nTo continue editing a project, provide its name. To start a new project, provide a new name."
    return result