from __future__ import annotations

import asyncio
import json
import re
from pathlib import Path

from loguru import logger
from openai import AsyncOpenAI

from ai.client import client
from ai.schemas import docx_schema
from config import settings
from sharepoint.sharepoint_models import File

sem = asyncio.Semaphore(5)


async def get_metadata_and_keywords(files: list[File]) -> None:
    async with client as ai_client, asyncio.TaskGroup() as tg:
        for file in files:
            tg.create_task(get_ai_response(ai_client, file))


async def get_ai_response(ai_client: AsyncOpenAI, file: File) -> None:
    async with sem:
        if file.file_extension != ".docx":
            logger.debug(
                f"File {file.name} not of type .docx, will not feed to OPENAI"
            )
            return

        logger.info(f"Sending file {file.name} to OPENAI for processing.")

        cleaned_text = re.sub(r"\\n", "", file.text_content)
        content = [{"type": "input_text", "text": cleaned_text}]

        await create_image_files(file, content)

        openai_input = [{"role": "user", "content": content}]

        response = await ai_client.responses.create(
            model="gpt-4o-mini",
            input=openai_input,
            text=docx_schema,
            instructions=get_instructions(),
        )

        if response:
            logger.success(f"Got response from OPENAI for file {file.name}")
            response = json.loads(response.output_text)
            file.images_metadata = response.get("image_metadata")
            file.keywords = response.get("keywords")
            file.summary = response.get("summary")


async def create_image_files(file: File, content: list[dict]) -> None:
    if file.images_bytes:
        for img in file.images_bytes:
            image_file = await client.files.create(
                file=(img[0], img[1]), purpose="vision"
            )

            content.append({
                "type": "input_image",
                "file_id": image_file.id,
            })


def get_instructions() -> str:
    with Path.open(settings.docx_instructions, "r", encoding="utf-8") as file:
        text_list = list(file)

    if text_list:
        return "".join(text_list)

    return None
