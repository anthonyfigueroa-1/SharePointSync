import asyncio
import json
import time
from pathlib import Path

import httpx
from azure.identity import InteractiveBrowserCredential
from loguru import logger

from app.config import settings
from sql_app.models import SharePointSolutionDoc

from app.httpx_client import create_sharepoint_client
from app.sharepoint.models import Folder, Image


async def sync_sharepoint_files() -> list[SharePointSolutionDoc] | None:
    bearer_token = get_az_token()

    async with create_sharepoint_client(bearer_token) as client:
        site_id = await get_site_id(client)
        drive_id = await get_drive_id(client, site_id=site_id)
        root_children = await get_root_children(client, drive_id=drive_id)

        files = await grab_all_files(client, drive_id=drive_id, root_children=root_children)

    # The below just checked if the updated_at time locally was < the cloud updated_at
    # Let's just build this into the add/update for the database
    # files = clean_redundant_files(files)

        if not files:
            return None

        logger.info("Getting file contents from all files retrieved")
        images = await get_file_contents(client, drive_id=drive_id, files=files)

    async with asyncio.TaskGroup() as tg:
        for file in files:
            tg.create_task(file.add_or_update_to_database())

#    async with asyncio.TaskGroup() as tg:
#        for image in images:



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


async def get_site_id(client: httpx.AsyncClient) -> str | None:
    url = f"sites/{settings.sharepoint_domain}.sharepoint.com:/sites/{settings.sharepoint_site_name}"

    for _ in range(3):
        try:
            response = await client.get(url)
            response.raise_for_status()

        except httpx.TimeoutException:
            logger.error(
                f"Timeout while getting site ID for SharePoint site {settings.sharepoint_site_name}"
            )
            asyncio.sleep(5)
            continue

        except httpx.HTTPStatusError:
            logger.error(
                f"HTTP error {response.status_code} | When getting site id for SharePoint Site {settings.sharepoint_site_name}"
            )
            asyncio.sleep(5)
            continue

        if response.is_success:
            logger.info(
                f"Successfully retrieved site ID for SharePoint site {settings.sharepoint_site_name}"
            )
            return response.json().get("id")

    logger.error(
        f"Failed to get site ID for SharePoint site {settings.sharepoint_site_name} after 3 attempts"
    )
    return None


async def get_drive_id(client: httpx.AsyncClient, site_id: str) -> str | None:
    url = f"sites/{site_id}/drives"

    for _ in range(3):
        try:
            response = await client.get(url)
            response.raise_for_status()

        except httpx.TimeoutException:
            logger.error(
                f"A timeout error occured when getting drive id for SharePoint Site {settings.sharepoint_site_name} | Site ID: {site_id}"
            )
            asyncio.sleep(5)
            continue

        except httpx.HTTPStatusError:
            logger.error(
                f"HTTP error {response.status_code} | When getting drive ID for SharePoint Site {settings.sharepoint_site_name} | Site ID: {site_id}"
            )
            asyncio.sleep(5)
            continue

        if response.is_success:
            response = response.json()["value"]

            for drive in response:
                drive_name: str = drive.get("name")
                if drive_name.lower() == "documents".lower():
                    logger.info(
                        f"Successfully retrieved Documents drive for SharePoint site {settings.sharepoint_site_name} | Drive ID: {drive.get('id')}"
                    )
                    return drive.get("id")

    return None


async def get_root_children(client: httpx.AsyncClient, drive_id: str) -> dict | None:
    url = f"drives/{drive_id}{f'{settings.sharepoint_site_root_dir}' if settings.sharepoint_site_root_dir == '/root/children' else f'/root:{settings.sharepoint_site_root_dir}:/children'}"

    for _ in range(3):
        try:
            response = await client.get(url)
            response.raise_for_status()

        except httpx.TimeoutException:
            logger.error(
                f"A timeout error occured when getting file info for SharePoint Site {settings.sharepoint_site_name} | Drive ID: {drive_id}"
            )
            asyncio.sleep(5)
            continue

        except httpx.HTTPStatusError:
            logger.error(
                f"HTTP error {response.status_code} | When getting root children for SharePoint site {settings.sharepoint_site_name} | Drive ID: {drive_id}"
            )
            asyncio.sleep(5)
            continue

        if response.is_success:
            root_children = response.json()["value"]
            logger.info(
                f"Successfully retrieved {len(root_children)} root children for SharePoint site {settings.sharepoint_site_name} | Drive ID: {drive_id}"
            )
            return root_children

    logger.error(
        f"Failed to get root children for SharePoint site {settings.sharepoint_site_name} | Drive ID: {drive_id} after 3 attempts"
    )
    return None


async def grab_all_files(
    client: httpx.AsyncClient, drive_id: str, root_children: list[dict]
) -> list[SharePointSolutionDoc]:
    logger.info(
            f"Getting all files from Drive ID {drive_id} located in SharePoint site {settings.sharepoint_site_name}"
        )

    sharepoint_files: list[SharePointSolutionDoc] = []

    process_filetypes = {".docx"}

    current_children = root_children
    folders_to_process: list[Folder] = []

    """
    This loop keeps searching through the SharePoint folders until there are no
    more folders to check.
    In each pass, it adds supported files to the result and collects any folders it finds.
    It then fetches the children of those folders and checks them in the next pass.
    Clearing the folder list after each pass prevents folders from being processed repeatedly.
    """
    while True:
        if not current_children:
            break

        # We are grabbing all files and folders in the current level of folders we are in
        # E.g. if we start at root, pos. 0, then we check all files on 0, then if more folders
        # exist, we check the next level, 1. And grab all files/folders that live at that level,
        # no matter the folder.
        # We seperate each item based off if it's a file or folder
        for item in current_children:
            if item.get("file"):
                file = SharePointSolutionDoc.convert_from_json(item)
                if file.file_extension in process_filetypes:
                    logger.info(f"Found file {file.name}")
                    sharepoint_files.append(file)

            if item.get("folder"):
                folder = Folder.convert_from_json(item)
                logger.info(f"Found directory {folder.name}")
                folders_to_process.append(folder)

        # Now that the current level has been checked, let's clear it out
        # Cause we will be re-using it and adding to it new stuff, if there is any
        current_children = []

        if not folders_to_process:
            break

        # Now all folders grabbed from the current level will be checked and we will grab all items within them
        for folder in folders_to_process:
            folder_children = await folder.get_nested_children(client, drive_id=drive_id)
            if folder_children:
                current_children.extend(folder_children)

        # Now they've been all processed we must clear the list because it get's reused in the loop
        folders_to_process = []

    logger.info(
        f"Done grabbing all files from {settings.sharepoint_site_name} SharePoint Site"
    )

    return sharepoint_files


async def get_file_contents(
    client: httpx.AsyncClient,
    *,
    drive_id: str,
    files: list[SharePointSolutionDoc]
) -> list[Image]:
    all_images: list[dict] = []
    async with (
        asyncio.TaskGroup() as tg,
    ):
        # Need to switch over to /beta/ cause for some reason /v1.0/ requires an unkown authentication
        # and beta does not need different auth
        for file in files:
            images = tg.create_task(file.grab_file_content(client, drive_id=drive_id))

            images = images.result()

            if images:
                all_images.extend(images)

        images = []
        image_models = []
        if all_images:
            for image in all_images:
                image_models.append(Image.model_validate(image))

        return image_models


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
    asyncio.run(get_and_save_sharepoint_files())
