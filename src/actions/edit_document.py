"""
Free-form find and replace on the latest version of a document project.
"""

import logging

from ..utils import (
    GraphClient,
    deserialize_context,
    find_and_replace,
    get_latest_version_filename,
    get_version_number,
    next_version_filename,
    serialize_context,
)

logger = logging.getLogger(__name__)


async def edit_document_action(
    project_name: str,
    find_text: str,
    replace_text: str,
    azure_tenant_id: str,
    azure_client_id: str,
    azure_client_secret: str,
    onedrive_user: str,
    onedrive_output_folder: str,
) -> str:
    """
    Perform a free-form find and replace on the latest version of a document
    project in OneDrive. Use this for quick text edits where placeholders no
    longer exist in the rendered document. Saves the result as the next
    incremented version and carries forward the Memory context.

    Args:
        project_name: The document project name e.g. tender_BSCGlobal
        find_text: The exact text to find in the document
        replace_text: The text to replace it with

    Returns:
        Success message with the new version filename and OneDrive link,
        or a message if the text was not found.
    """
    logger.info(f"Editing document for project: {project_name}")

    graph = GraphClient(
        tenant_id=azure_tenant_id,
        client_id=azure_client_id,
        client_secret=azure_client_secret,
    )

    # ── Get existing project files ─────────────────────────────────────────────
    project_folder = f"{onedrive_output_folder}/{project_name}"
    memory_folder = f"{project_folder}/Memory"

    existing_items = await graph.list_onedrive_folder(
        onedrive_user=onedrive_user,
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

    # ── Download latest .docx from OneDrive ───────────────────────────────────
    docx_bytes = await graph.download_onedrive_file(
        onedrive_user=onedrive_user,
        folder_path=project_folder,
        file_name=latest_docx,
    )

    # ── Perform find & replace ─────────────────────────────────────────────────
    updated_bytes, replacements_made = find_and_replace(
        docx_bytes=docx_bytes,
        find_text=find_text,
        replace_text=replace_text,
    )

    if replacements_made == 0:
        return (
            f"Text '{find_text}' was not found in {latest_docx}. "
            f"No changes were made."
        )

    # ── Calculate next version filename ───────────────────────────────────────
    new_filename = next_version_filename(project_name, existing_files)

    # ── Upload updated .docx to OneDrive ──────────────────────────────────────
    web_url = await graph.upload_onedrive_file(
        onedrive_user=onedrive_user,
        folder_path=project_folder,
        file_name=new_filename,
        content=updated_bytes,
    )

    # ── Carry forward Memory JSON ──────────────────────────────────────────────
    previous_memory = latest_docx.replace(".docx", ".json")
    json_str = await graph.download_json_file(
        onedrive_user=onedrive_user,
        folder_path=memory_folder,
        file_name=previous_memory,
    )
    previous_context = deserialize_context(json_str)

    if previous_context:
        new_memory = new_filename.replace(".docx", ".json")
        await graph.upload_json_file(
            onedrive_user=onedrive_user,
            folder_path=memory_folder,
            file_name=new_memory,
            content=serialize_context(previous_context),
        )

    logger.info(
        f"Edit complete: {replacements_made} replacement(s) made, "
        f"saved as {new_filename}"
    )

    return (
        f"✅ Document updated successfully!\n"
        f"   File: {new_filename}\n"
        f"   Project: {project_name}\n"
        f"   Replaced: '{find_text}' → '{replace_text}'\n"
        f"   Replacements made: {replacements_made}\n"
        f"   Open in OneDrive: {web_url}"
    )