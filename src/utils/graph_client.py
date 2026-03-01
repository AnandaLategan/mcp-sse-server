"""
Microsoft Graph API client for Word MCP Server.
Handles authentication and all SharePoint/OneDrive operations.
"""

import logging
from urllib.parse import quote

import httpx
import msal

logger = logging.getLogger(__name__)


class GraphClient:
    """Handles all Microsoft Graph API operations."""

    GRAPH_BASE = "https://graph.microsoft.com/v1.0"

    def __init__(self, tenant_id: str, client_id: str, client_secret: str):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self._token: str | None = None

    # ── Authentication ─────────────────────────────────────────────────────────

    def _get_token(self) -> str:
        """Acquire an access token using client credentials flow."""
        app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
        )
        result = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        if "access_token" not in result:
            error = result.get("error_description", "Unknown error")
            raise ValueError(f"Failed to acquire Graph API token: {error}")

        logger.debug("Graph API token acquired successfully")
        return result["access_token"]

    def _headers(self) -> dict:
        """Return auth headers for Graph API requests."""
        if not self._token:
            self._token = self._get_token()
        return {
            "Authorization": f"Bearer {self._token}",
            "Content-Type": "application/json",
        }

    def _encode_path(self, path: str) -> str:
        """URL encode a path while preserving forward slashes."""
        return quote(path, safe="/")

    # ── SharePoint ─────────────────────────────────────────────────────────────

    async def download_sharepoint_file(
        self, drive_id: str, folder_path: str, file_name: str
    ) -> bytes:
        """Download a file from SharePoint and return its bytes."""
        file_path = self._encode_path(f"{folder_path}/{file_name}")
        url = f"{self.GRAPH_BASE}/drives/{drive_id}/root:/{file_path}:/content"
        async with httpx.AsyncClient() as client:
            response = await client.get(
                url, headers=self._headers(), follow_redirects=True
            )
            response.raise_for_status()
            logger.info(f"Downloaded SharePoint file: {file_name}")
            return response.content

    async def list_sharepoint_files(
        self, drive_id: str, folder_path: str
    ) -> list[dict]:
        """List files in a SharePoint folder."""
        encoded_path = self._encode_path(folder_path)
        url = f"{self.GRAPH_BASE}/drives/{drive_id}/root:/{encoded_path}:/children"
        async with httpx.AsyncClient() as client:
            response = await client.get(url, headers=self._headers())
            response.raise_for_status()
            items = response.json().get("value", [])
            return [
                {"name": item["name"], "id": item["id"]}
                for item in items
                if not item.get("folder")
            ]

    async def list_sharepoint_folder(
        self, drive_id: str, folder_path: str
    ) -> list[dict]:
        """List files and folders in a SharePoint folder."""
        encoded_path = self._encode_path(folder_path)
        url = f"{self.GRAPH_BASE}/drives/{drive_id}/root:/{encoded_path}:/children"
        async with httpx.AsyncClient() as client:
            response = await client.get(url, headers=self._headers())
            if response.status_code == 404:
                return []
            response.raise_for_status()
            items = response.json().get("value", [])
            return [
                {
                    "name": item["name"],
                    "id": item["id"],
                    "is_folder": "folder" in item,
                }
                for item in items
            ]

    async def upload_sharepoint_file(
        self, drive_id: str, folder_path: str, file_name: str, content: bytes
    ) -> str:
        """
        Upload a .docx file to a SharePoint folder.
        Returns the web URL of the uploaded file.
        """
        upload_path = self._encode_path(f"{folder_path}/{file_name}")
        url = f"{self.GRAPH_BASE}/drives/{drive_id}/root:/{upload_path}:/content"

        headers = self._headers()
        headers["Content-Type"] = (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        async with httpx.AsyncClient() as client:
            response = await client.put(url, headers=headers, content=content)
            response.raise_for_status()
            web_url = response.json().get("webUrl", "")
            logger.info(f"Uploaded file to SharePoint: {folder_path}/{file_name}")
            return web_url

    async def upload_sharepoint_json(
        self, drive_id: str, folder_path: str, file_name: str, content: str
    ) -> str:
        """Upload a JSON file to a SharePoint folder."""
        upload_path = self._encode_path(f"{folder_path}/{file_name}")
        url = f"{self.GRAPH_BASE}/drives/{drive_id}/root:/{upload_path}:/content"

        headers = self._headers()
        headers["Content-Type"] = "application/json"

        async with httpx.AsyncClient() as client:
            response = await client.put(
                url, headers=headers, content=content.encode("utf-8")
            )
            response.raise_for_status()
            logger.info(f"Uploaded JSON to SharePoint: {folder_path}/{file_name}")
            return response.json().get("webUrl", "")

    async def download_sharepoint_json(
        self, drive_id: str, folder_path: str, file_name: str
    ) -> str:
        """Download a JSON file from SharePoint and return its text content."""
        file_path = self._encode_path(f"{folder_path}/{file_name}")
        url = f"{self.GRAPH_BASE}/drives/{drive_id}/root:/{file_path}:/content"

        async with httpx.AsyncClient() as client:
            response = await client.get(
                url, headers=self._headers(), follow_redirects=True
            )
            if response.status_code == 404:
                return ""
            response.raise_for_status()
            return response.text