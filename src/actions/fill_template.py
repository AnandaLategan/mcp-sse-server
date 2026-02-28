"""
Fill a Word template with context values and save to OneDrive.
"""

import json
import logging

from ..utils import (
    GraphClient,
    deserialize_context,
    get_latest_version_filename,
    next_version_filename,
    render_template,
    serialize_context,
)

logger = logging.getLogger(__name__)


async def fill_template_action(
    template_name: str,
    project_name: str,
    replacements: dict,
    azure_tenant_id: str,
    azure_client_id: str,
    azure_client_secret: str,
    sharepoint_drive_id: str,              # ← changed
    sharepoint_template_folder: str,
    onedrive_user: str,
    onedrive_output_folder: str,
) -> str:
    """
    Fill a Word template with the provided values and save it to OneDrive.

    - If the project already has versions, loads the previous context from
      OneDrive Memory, merges the new replacements on top, and re-renders
      from the original template with the full merged context.
    - If the project is new, renders directly from the template.
    - Always saves as the next incremented version with a matching
      Memory JSON file.

    Args:
        template_name: The template filename e.g. tender_template.docx
        project_name: The document project name e.g. tender_BSCGlobal
        replacements: Dict of placeholder keys and values to apply or update
                      e.g. {"company_name": "BSC Global", "date": "2026-02-28"}

    Returns:
        Success message with the new version filename and OneDrive link.
    """
    logger.info(f"Filling template for project: {project_name}")

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
    existing_files = [
        item["name"] for item in existing_items
        if not item["is_folder"]
    ]

    # ── Load previous context if exists ───────────────────────────────────────
    latest_docx = get_latest_version_filename(project_name, existing_files)

    if latest_docx:
        # Load previous Memory JSON
        memory_filename = latest_docx.replace(".docx", ".json")
        json_str = await graph.download_json_file(
            onedrive_user=onedrive_user,
            folder_path=memory_folder,
            file_name=memory_filename,
        )
        previous_context = deserialize_context(json_str)
        base_label = f"previous version ({latest_docx})"
        logger.info(f"Loaded previous context from {memory_filename}")
    else:
        previous_context = {}
        base_label = f"template ({template_name})"
        logger.info("No previous version found — starting from template")

    # ── Merge contexts ─────────────────────────────────────────────────────────
    merged_context = {**previous_context, **replacements}

    # ── Download template from SharePoint ─────────────────────────────────────
    template_bytes = await graph.download_sharepoint_file(
        drive_id=sharepoint_drive_id,              # ← changed
        folder_path=sharepoint_template_folder,
        file_name=template_name,
    )

    # ── Render template ────────────────────────────────────────────────────────
    rendered_bytes = render_template(template_bytes, merged_context)

    # ── Calculate next version filename ───────────────────────────────────────
    new_filename = next_version_filename(project_name, existing_files)

    # ── Upload rendered .docx to OneDrive ─────────────────────────────────────
    web_url = await graph.upload_onedrive_file(
        onedrive_user=onedrive_user,
        folder_path=project_folder,
        file_name=new_filename,
        content=rendered_bytes,
    )

    # ── Upload Memory JSON to OneDrive ─────────────────────────────────────────
    memory_filename = new_filename.replace(".docx", ".json")
    await graph.upload_json_file(
        onedrive_user=onedrive_user,
        folder_path=memory_folder,
        file_name=memory_filename,
        content=serialize_context(merged_context),
    )

    logger.info(f"Successfully created {new_filename} for project {project_name}")

    return (
        f"✅ Document saved successfully!\n"
        f"   File: {new_filename}\n"
        f"   Project: {project_name}\n"
        f"   Based on: {base_label}\n"
        f"   Values applied: {list(replacements.keys())}\n"
        f"   Total placeholders filled: {len(merged_context)}\n"
        f"   Open in OneDrive: {web_url}"
    )