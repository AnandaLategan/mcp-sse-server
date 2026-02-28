"""
Microsoft Graph API client for Word MCP Server.
Handles authentication and all SharePoint/OneDrive operations.
"""

import logging
from io import BytesIO

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

    # ── SharePoint ─────────────────────────────────────────────────────────────

    async def get_sharepoint_site_id(self, site_url: str) -> str:
        """Get the SharePoint site ID from a site URL."""
        # Extract host and path from URL
        # e.g. https://cyestcorp.sharepoint.com/sites/BSC-Systems
        parts = site_url.replace("https://", "").split("/", 1)
        host = parts[0]
        path = parts[1] if len(parts) > 1 else ""

        url = f"{self.GRAPH_BASE}/sites/{host}:/{path}"
        async with httpx.AsyncClient() as client:
            response = await client.get(url, headers=self._headers())
            response.raise_for_status()
            return response.json()["id"]

    async def download_sharepoint_file(
        self, site_url: str, folder_path: str, file_name: str
    ) -> bytes:
        """Download a file from SharePoint and return its bytes."""
        site_id = await self.get_sharepoint_site_id(site_url)
        file_path = f"{folder_path}/{file_name}"

        url = f"{self.GRAPH_BASE}/sites/{site_id}/drive/root:/{file_path}:/content"
        async with httpx.AsyncClient() as client:
            response = await client.get(
                url, headers=self._headers(), follow_redirects=True
            )
            response.raise_for_status()
            logger.info(f"Downloaded SharePoint file: {file_name}")
            return response.content

    async def list_sharepoint_files(
        self, site_url: str, folder_path: str
    ) -> list[dict]:
        """List files in a SharePoint folder."""
        site_id = await self.get_sharepoint_site_id(site_url)

        url = f"{self.GRAPH_BASE}/sites/{site_id}/drive/root:/{folder_path}:/children"
        async with httpx.AsyncClient() as client:
            response = await client.get(url, headers=self._headers())
            response.raise_for_status()
            items = response.json().get("value", [])
            return [
                {"name": item["name"], "id": item["id"]}
                for item in items
                if not item.get("folder")
            ]

    # ── OneDrive ───────────────────────────────────────────────────────────────

    async def upload_onedrive_file(
        self, onedrive_user: str, folder_path: str, file_name: str, content: bytes
    ) -> str:
        """
        Upload a file to a OneDrive folder.
        Returns the web URL of the uploaded file.
        """
        upload_path = f"{folder_path}/{file_name}"
        url = (
            f"{self.GRAPH_BASE}/users/{onedrive_user}/drive/root:/"
            f"{upload_path}:/content"
        )

        headers = self._headers()
        headers["Content-Type"] = (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        async with httpx.AsyncClient() as client:
            response = await client.put(url, headers=headers, content=content)
            response.raise_for_status()
            web_url = response.json().get("webUrl", "")
            logger.info(f"Uploaded file to OneDrive: {upload_path}")
            return web_url

    async def download_onedrive_file(
        self, onedrive_user: str, folder_path: str, file_name: str
    ) -> bytes:
        """Download a file from OneDrive and return its bytes."""
        file_path = f"{folder_path}/{file_name}"
        url = (
            f"{self.GRAPH_BASE}/users/{onedrive_user}/drive/root:/"
            f"{file_path}:/content"
        )

        async with httpx.AsyncClient() as client:
            response = await client.get(
                url, headers=self._headers(), follow_redirects=True
            )
            response.raise_for_status()
            logger.info(f"Downloaded OneDrive file: {file_name}")
            return response.content

    async def list_onedrive_folder(
        self, onedrive_user: str, folder_path: str
    ) -> list[dict]:
        """List files and folders in a OneDrive folder."""
        url = (
            f"{self.GRAPH_BASE}/users/{onedrive_user}/drive/root:/"
            f"{folder_path}:/children"
        )

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

    async def upload_json_file(
        self, onedrive_user: str, folder_path: str, file_name: str, content: str
    ) -> str:
        """Upload a JSON file to OneDrive."""
        upload_path = f"{folder_path}/{file_name}"
        url = (
            f"{self.GRAPH_BASE}/users/{onedrive_user}/drive/root:/"
            f"{upload_path}:/content"
        )

        headers = self._headers()
        headers["Content-Type"] = "application/json"

        async with httpx.AsyncClient() as client:
            response = await client.put(
                url, headers=headers, content=content.encode("utf-8")
            )
            response.raise_for_status()
            logger.info(f"Uploaded JSON file to OneDrive: {upload_path}")
            return response.json().get("webUrl", "")

    async def download_json_file(
        self, onedrive_user: str, folder_path: str, file_name: str
    ) -> str:
        """Download a JSON file from OneDrive and return its text content."""
        file_path = f"{folder_path}/{file_name}"
        url = (
            f"{self.GRAPH_BASE}/users/{onedrive_user}/drive/root:/"
            f"{file_path}:/content"
        )

        async with httpx.AsyncClient() as client:
            response = await client.get(
                url, headers=self._headers(), follow_redirects=True
            )
            if response.status_code == 404:
                return ""
            response.raise_for_status()
            return response.text