import asyncio
import json
import time
from pathlib import Path

import httpx
from azure.identity import InteractiveBrowserCredential
from loguru import logger

from config import settings
from sharepoint.sharepoint_models import File, Folder


async def get_sharepoint_files() -> list[File]:
    token = get_az_token()

    header = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}",
    }

    site_id = get_site_id(header)
    drive_id = get_drive_id(header, site_id)
    root_children = get_root_children(header, drive_id)
    logger.info(
        f"Getting all files from Drive ID {drive_id} located in SharePoint site {settings.sharepoint_site_name}"
    )
    files = grab_all_files(header, drive_id, root_children)

    files = await clean_redundant_files(files)

    if not files:
        return None

    logger.info("Getting file contents from all files retrieved")
    await get_file_contents(header, drive_id, files)

    return files


def get_az_token() -> str:
    token = check_old_token()

    if not token:
        scope = "https://graph.microsoft.com/.default"
        tenant_id = settings.tenant_id.get_secret_value()
        client = InteractiveBrowserCredential(tenant=tenant_id)

        response = client.get_token(scope)

        token = response.token
        expires = response.expires_on

        save_new_token(token, expires)

    return token


def get_site_id(header: dict) -> str | None:
    url = f"https://graph.microsoft.com/v1.0/sites/{settings.sharepoint_domain}.sharepoint.com:/sites/{settings.sharepoint_site_name}"

    try:
        response = httpx.get(url=url, headers=header, timeout=(20, 20))

    except httpx.Timeout:
        logger.error(
            f"A timeout error occured when getting site id for SharePoint Site {settings.sharepoint_site_name}"
        )

    else:
        if response.is_success:
            return response.json().get("id")

        logger.error(
            f"{response.status_code} | When getting site id for SharePoint Site {settings.sharepoint_site_name}"
        )
        return None


def get_drive_id(header: dict, site_id: str) -> str | None:
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"

    try:
        response = httpx.get(url=url, headers=header, timeout=(20, 20))

    except httpx.Timeout:
        logger.error(
            f"A timeout error occured when getting drive id for SharePoint Site {settings.sharepoint_site_name} | Site ID: {site_id}"
        )

    else:
        if not response.is_success:
            logger.error(
                f"{response.status_code} | When getting drive id for SharePoint Site {settings.sharepoint_site_name} | Site ID: {site_id}"
            )
            return None

        response = response.json()["value"]

        for drive in response:
            drive_name: str = drive.get("name")
            if drive_name.lower() == "documents".lower():
                return drive.get("id")


def get_root_children(header: dict, drive_id: str) -> dict | None:
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}{f'{settings.sharepoint_site_root_dir}' if settings.sharepoint_site_root_dir == '/root/children' else f'/root:{settings.sharepoint_site_root_dir}:/children'}"

    try:
        response = httpx.get(url=url, headers=header, timeout=(20, 20))

    except httpx.Timeout:
        logger.error(
            f"A timeout error occured when getting file info for SharePoint Site {settings.sharepoint_site_name} | Drive ID: {drive_id}"
        )

    else:
        if response.is_success:
            root_children = response.json()["value"]
            return root_children

        logger.error(
            f"{response.status_code} | When getting root_children for SharePoint Site {settings.sharepoint_site_name} | Drive ID: {drive_id}"
        )
        return None


def grab_all_files(
    header: dict, drive_id: str, root_children: list[dict]
) -> list[File]:
    files: list[File] = []

    process_filetypes = set(settings.process_filetypes)
    parent_folders: list[Folder] = []
    loops = 1

    while True:
        if not root_children:
            break

        for item in root_children:
            if item.get("file"):
                file = File.convert_json_to_class(item)
                if file.file_extension in process_filetypes:
                    logger.info(f"Found file {file.name}")
                    files.append(file)

            if item.get("folder"):
                folder = Folder.convert_json_to_class(item)
                logger.info(f"Found directory {folder.name}")
                parent_folders.append(folder)

        root_children = []

        if not parent_folders:
            break

        for folder in parent_folders:
            nested_children = folder.get_nested_children(header, drive_id)
            root_children += nested_children

        parent_folders = []

        loops += 1

    logger.info(
        f"Done grabbing all files from {settings.sharepoint_site_name} SharePoint Site"
    )

    return files


async def clean_redundant_files(files: list[File]) -> list[File] | None:
    cleaned_files = []
    for file in files:
        should_add_or_update = await file.check_if_should_update_or_add()
        if should_add_or_update:
            cleaned_files.append(file)

    return cleaned_files


async def get_file_contents(
    header: dict, drive_id: str, files: list[File]
) -> None:
    limits = httpx.Limits(max_connections=10, max_keepalive_connections=5)
    timeout = httpx.Timeout(connect=20, read=60, write=60, pool=20)
    async with (
        httpx.AsyncClient(limits=limits, timeout=timeout) as client,
        asyncio.TaskGroup() as tg,
    ):
        for file in files:
            tg.create_task(file.grab_file_content(client, header, drive_id))


def check_old_token() -> str | None:
    try:
        with Path.open("./mstoken.json", "r") as file:
            info = json.load(file)

    except FileNotFoundError:
        return None

    now = int(time.time())

    expires = info.get("expires")

    if expires < now:
        return None

    return info.get("token")


def save_new_token(token: str, expires: int) -> None:
    payload = {"token": token, "expires": expires}

    with Path.open("./mstoken.json", "w") as file:
        json.dump(payload, file, indent=4)


if __name__ == "__main__":
    get_sharepoint_files()
