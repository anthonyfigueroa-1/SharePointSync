from openai import AsyncOpenAI
import openai

from app.config import settings


def create_async_openai_client() -> openai.AsyncOpenAI:
    return AsyncOpenAI(api_key=settings.openai_key.get_secret_value())
