import asyncio
import sys

from loguru import logger

from app.sharepoint.file_fetcher import sync_sharepoint_files


async def main():
    logger.info("Running SharePoint sync")
    files = await sync_sharepoint_files()

    if not files:
        logger.info("No files require updating at this time. Exiting SPS.")
        sys.exit(0)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("Exiting SPS")
