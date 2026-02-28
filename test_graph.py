import asyncio
import httpx
from src.utils.graph_client import GraphClient
from dotenv import load_dotenv
import os

load_dotenv()

async def test():
    graph = GraphClient(
        tenant_id=os.getenv('AZURE_TENANT_ID'),
        client_id=os.getenv('AZURE_CLIENT_ID'),
        client_secret=os.getenv('AZURE_CLIENT_SECRET'),
    )

    drive_id = "b!zUmKjjVj_025fNqBJYFSH0XzDGHRho1KhB8MECPyN5Bf9GHZag7KQ5VfRhpM4AU-"

    # List 7. AI folder
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/7. AI:/children"
    async with httpx.AsyncClient(timeout=30) as client:
        response = await client.get(url, headers=graph._headers())
        response.raise_for_status()
        items = response.json().get("value", [])
        print("7. AI folder contents:")
        for item in items:
            item_type = "📁" if "folder" in item else "📄"
            print(f"  {item_type} {item['name']}")

asyncio.run(test())