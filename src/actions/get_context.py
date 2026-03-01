"""
Retrieve the saved context for the latest version of a document project.
"""

import logging

from ..utils import (
    GraphClient,
    deserialize_context,
    get_latest_version_filename,
    get_version_number,
)

logger = logging.getLogger(__name__)


async def get_context_action(
    project_name: str,
    azure_tenant_id: str,
    azure_client_id: str,
    azure_client_secret: str,
    sharepoint_drive_id: str,              # ← changed
    sharepoint_output_folder: str,         # ← changed
) -> str:
    """
    Retrieve all currently filled placeholder values for the latest version
    of a document project from SharePoint Memory. Call this when the user wants
    to see what values are already filled in before making edits, or to
    review the current state of a document project.

    Args:
        project_name: The document project name e.g. tender_BSCGlobal

    Returns:
        All currently filled placeholder values for the latest version,
        or a message if no context exists yet.
    """
    logger.info(f"Getting context for project: {project_name}")

    graph = GraphClient(
        tenant_id=azure_tenant_id,
        client_id=azure_client_id,
        client_secret=azure_client_secret,
    )

    # ── Get existing project files ─────────────────────────────────────────────
    project_folder = f"{sharepoint_output_folder}/{project_name}"
    memory_folder = f"{project_folder}/Memory"

    existing_items = await graph.list_sharepoint_folder(   # ← changed
        drive_id=sharepoint_drive_id,                      # ← changed
        folder_path=project_folder,
    )

    if not existing_items:
        return f"No versions found for project '{project_name}'."

    existing_files = [
        item["name"] for item in existing_items
        if not item["is_folder"]
    ]

    # ── Find latest version ────────────────────────────────────────────────────
    latest_docx = get_latest_version_filename(project_name, existing_files)

    if not latest_docx:
        return f"No versions found for project '{project_name}'."

    # ── Load Memory JSON ───────────────────────────────────────────────────────
    memory_filename = latest_docx.replace(".docx", ".json")
    json_str = await graph.download_sharepoint_json(       # ← changed
        drive_id=sharepoint_drive_id,                      # ← changed
        folder_path=memory_folder,
        file_name=memory_filename,
    )

    context = deserialize_context(json_str)

    if not context:
        return (
            f"No saved context found for project '{project_name}' "
            f"version '{latest_docx}'."
        )

    # ── Format output ──────────────────────────────────────────────────────────
    version = get_version_number(latest_docx)
    result = (
        f"Current values for project '{project_name}' "
        f"(v{str(version).zfill(2)}):\n\n"
    )
    for key, value in sorted(context.items()):
        result += f"  - {key}: {value}\n"

    return result