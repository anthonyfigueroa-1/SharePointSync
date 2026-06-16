import asyncio
import re
from datetime import datetime
from io import BytesIO
from typing import Self
from zipfile import ZipFile

import docx
import httpx
from loguru import logger
from pydantic import BaseModel
from sqlalchemy import select
from sqlalchemy.orm import Session

from config import settings
from sql_app.models import SharePointSolutionDoc
from sql_app.database import get_db

sem = asyncio.Semaphore(5)


class File(BaseModel):
    id: str
    created_datetime: datetime
    last_modified_datetime: datetime
    name: str
    path: str
    file_extension: str
    web_url: str
    download_url: str

    text_content: str | None = None
    images_bytes: list[list[bytes | str]] | None = None
    images_metadata: list[str] | None = None
    keywords: list[str] | None = None
    summary: str | None = None

    @classmethod
    def convert_json_to_class(cls, file_dict: dict) -> Self:
        parent_reference: dict = file_dict.get("parentReference")
        file_info: dict = file_dict.get("file")
        file = cls(
            id=file_dict.get("id"),
            created_datetime=file_dict.get("createdDateTime"),
            last_modified_datetime=file_dict.get("lastModifiedDateTime"),
            name=file_dict.get("name"),
            path=parent_reference.get("path"),
            file_extension=file_info.get("fileExtension"),
            web_url=file_dict.get("webUrl"),
            download_url=file_dict.get("@microsoft.graph.downloadUrl"),
        )
        file.clean_file_path()
        return file

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "type": "file",
            "created_datetime": datetime.strftime(
                self.created_datetime, "%Y-%m-%d %H:%M:%S"
            ),
            "last_modified_datetime": datetime.strftime(
                self.last_modified_datetime, "%Y-%m-%d %H:%M:%S"
            ),
            "name": self.name,
            "path": self.path,
            "file_extension": self.file_extension,
            "web_url": self.web_url,
            "download_url": self.download_url,
        }

    async def grab_file_content(
        self, client: httpx.AsyncClient, header: dict, drive_id: str
    ) -> None:
        async with sem:
            docx_file = re.search(r"\w+\.docx", self.name)

            if not docx_file:
                logger.warning(
                    f"File {self.name} is not a .docx file, will be skipping over it."
                )
                return

            url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{self.id}/content"

            try:
                response = await client.get(
                    url, headers=header, follow_redirects=True
                )

            except httpx.TimeoutException:
                logger.error(f"Timeout error when getting file {self.name}")

            else:
                if not response.is_success:
                    logger.error(
                        f"{response.status_code} | When getting file {self.name} from SharePoint Drive ID: {drive_id}"
                    )
                    return

                logger.info(
                    f"Got file contents for file {self.path}{self.name}"
                )
                response_bytes = BytesIO(response.content)

                doc = docx.Document(response_bytes)

                doc_text = [
                    par.text for par in doc.paragraphs if par.text.strip()
                ]
                self.text_content = "\n".join(doc_text)

                images = []
                with ZipFile(response_bytes) as zipper:
                    for name in zipper.namelist():
                        image = re.search(r"(?:word\/)media\/(.+)$", name)
                        if image:
                            data = zipper.read(name)
                            images.append([image.group(1), data])

                self.images_bytes = images

    def clean_file_path(self) -> None:
        grab_path = re.search(r"(\/root\:\/.+$)", self.path)
        if grab_path:
            cleaned_path = re.sub(r"\:", "", grab_path.group(1))

            if cleaned_path:
                self.path = cleaned_path

    def add_file_to_database(self, db: Session) -> None:
        file = db.execute(
            select(SharePointSolutionDoc).where(
                SharePointSolutionDoc.sp_id == self.id
            )
        ).scalar_one_or_none()
        if file:
            self.update_file_to_database(file)
            return

        entry = SharePointSolutionDoc()

        entry.sp_id = self.id
        entry.file_type = self.file_extension
        entry.group_id = settings.group_id
        entry.name = self.name
        entry.content_text = self.text_content
        entry.img_metadata = self.images_metadata
        entry.keywords = self.keywords
        entry.summary = self.summary
        entry.file_path = self.path
        entry.web_url = self.web_url
        entry.created_datetime = self.created_datetime
        entry.last_modified = self.last_modified_datetime
        entry.last_synced = datetime.now(settings.tz)

        db.add(entry)
        logger.success(f"Added file {self.name} to the database")

    def update_file_to_database(
        self, file: SharePointSolutionDoc
    ) -> None:
        file.name = self.name
        file.content_text = self.text_content
        file.img_metadata = self.images_metadata
        file.keywords = self.keywords
        file.summary = self.summary
        file.file_path = self.path
        file.web_url = self.web_url
        file.last_modified = self.last_modified_datetime
        file.last_synced = datetime.now(settings.tz)

        logger.success(f"Updated file {self.name} to the database")

    def check_if_should_update(self) -> bool:
        with get_db() as db:
            db_entry_id = db.execute(
                select(SharePointSolutionDoc.id).where(
                    SharePointSolutionDoc.sp_id == self.id
                )
            )
            db_entry_id = db_entry_id.scalar_one_or_none()

            if not db_entry_id:
                return True

            db_entry = db.get(SharePointSolutionDoc, db_entry_id)

        if self.last_modified_datetime >= db_entry.last_synced:
            return True

        last_synced_str = datetime.strftime(
            db_entry.last_synced.astimezone(tz=settings.tz),
            "%a, %b %d, %Y at %H:%M",
        )
        last_modified_str = datetime.strftime(
            db_entry.last_modified_datetime.astimezone(tz=settings.tz),
            "%a, %b %d, %Y at %H:%M",
        )
        logger.info(
            f"Will not update DB file {db_entry.name}, ID: {db_entry.id}, due to last sync occuring {last_synced_str} and last change having happened on {last_modified_str}."
        )
        return False


class Folder(BaseModel):
    id: str
    name: str

    @classmethod
    def convert_json_to_class(cls, folder_dict: dict) -> Self:
        return cls(id=folder_dict.get("id"), name=folder_dict.get("name"))

    def to_dict(self) -> dict:
        return {"id": self.id, "type": "folder", "name": self.name}

    def get_nested_children(
        self, header: dict, drive_id: str
    ) -> list[dict] | None:
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{self.id}/children"

        response = httpx.get(url, headers=header, timeout=(20, 20))

        if response.is_success:
            nested_children = response.json()["value"]
            return nested_children

        logger.error(
            f"{response.status_code} | When getting children for SharePoint Site {settings.sharepoint_site_name} | Drive ID: {drive_id} | Folder Name: {self.name} | Folder ID: {self.id}"
        )
        return None
