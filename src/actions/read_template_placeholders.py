"""
Read and list all placeholders from a Word template stored in SharePoint.
"""

import logging

from ..utils import GraphClient, scan_placeholders

logger = logging.getLogger(__name__)


async def read_template_placeholders_action(
    template_name: str,
    azure_tenant_id: str,
    azure_client_id: str,
    azure_client_secret: str,
    sharepoint_site_url: str,
    sharepoint_template_folder: str,
) -> str:
    """
    Download a Word template from SharePoint and return all Jinja2
    {{ placeholder }} tags found in it. Call this when the user wants
    to start a new document so they know exactly what information
    needs to be provided.

    Args:
        template_name: The filename of the template e.g. tender_template.docx

    Returns:
        A list of all placeholder names found in the template.
    """
    logger.info(f"Reading placeholders from template: {template_name}")

    graph = GraphClient(
        tenant_id=azure_tenant_id,
        client_id=azure_client_id,
        client_secret=azure_client_secret,
    )

    # Download template from SharePoint
    template_bytes = await graph.download_sharepoint_file(
        site_url=sharepoint_site_url,
        folder_path=sharepoint_template_folder,
        file_name=template_name,
    )

    # Scan for placeholders
    placeholders = scan_placeholders(template_bytes)

    if not placeholders:
        return f"No placeholders found in '{template_name}'."

    placeholder_list = "\n".join([f"  - {{{{ {p} }}}}" for p in placeholders])
    return (
        f"Placeholders found in '{template_name}' ({len(placeholders)} total):\n"
        f"{placeholder_list}"
    )