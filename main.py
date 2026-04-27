import asyncio
import sys

from loguru import logger

from ai.docx import get_metadata_and_keywords
from sharepoint.sharepoint import get_sharepoint_files


async def main():
    logger.info("Running SharePoint sync")
    files = await get_sharepoint_files()

    if not files:
        logger.info("No files require updating at this time. Exiting SPS.")
        sys.exit(0)

    await get_metadata_and_keywords(files)

    async with asyncio.TaskGroup() as tg:
        for file in files:
            tg.create_task(file.add_file_to_database())


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("Exiting SPS")
