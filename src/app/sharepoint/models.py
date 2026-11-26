import asyncio
from loguru import logger

import httpx
from pydantic import BaseModel
from typing import Self
from app.config import settings


class Folder(BaseModel):
    id: str
    name: str

    @classmethod
    def convert_from_json(cls, folder: dict) -> Self:
        return cls(
            id=folder.get("id"),
            name=folder.get("name")
        )

    def to_dict(self) -> dict:
        return {"id": self.id, "type": "folder", "name": self.name}

    async def get_nested_children(
        self,
        client: httpx.AsyncClient,
        *,
        drive_id: str
    ) -> list[dict] | None:
        url = f"drives/{drive_id}/items/{self.id}/children"

        for _ in range(3):
            try:
                response = await client.get(url)
                response.raise_for_status()

            except httpx.TimeoutException:
                logger.error(
                    f"Timeout while getting children for SharePoint site {settings.sharepoint_site_name} | Drive ID: {drive_id} | Folder Name: {self.name} | Folder ID: {self.id}"
                )
                asyncio.sleep(5)
                continue

            except httpx.HTTPStatusError:
                logger.error(
                    f"HTTP error {response.status_code} | When getting children for SharePoint site {settings.sharepoint_site_name} | Drive ID: {drive_id} | Folder Name: {self.name} | Folder ID: {self.id}"
                )
                asyncio.sleep(5)
                continue

            if response.is_success:
                nested_children = response.json()["value"]
                logger.info(
                    f"Retrieved {len(nested_children)} children for folder {self.name} | Folder ID: {self.id}"
                )
                return nested_children

        logger.error(
            f"Failed to get children for folder {self.name} | Folder ID: {self.id} after 3 attempts"
        )
        return None


class Image(BaseModel):
    name: str
    article_id: int
    image: bytes
