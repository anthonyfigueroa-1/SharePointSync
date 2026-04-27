from openai import AsyncOpenAI

from config import settings

client = AsyncOpenAI(api_key=settings.openai_key.get_secret_value())
