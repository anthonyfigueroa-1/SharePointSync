from typing import Self
from loguru import logger
import openai
import base64
from pydantic import BaseModel


class ImageResponse(BaseModel):
    description: str

    @classmethod
    async def process_image(cls, client: openai.AsyncOpenAI, image: bytes) -> Self | None:
        try:
            response = await client.responses.parse(
                model="gpt-6-luna",
                input={
                    "role": "user",
                    "content": [{
                        "type": "input_image",
                        "image_url": f"data:image/jpeg;base64,{ImageResponse.encode_image(image)}",
                    }]
                },
                instructions="Provide a very detailed description of the image provided. Detailed where the description can explain everything in an image. Keep the descriptions to at most 2 paragraphs.",
                text_format=cls
            )

        except openai.APIConnectionError:
            logger.error(
                "OpenAI API connection error while processing image",
            )
            return None
        except openai.APIStatusError:
            logger.error(
                "OpenAI API status error while processing image",
            )
            return None
        except openai.APIError:
            logger.error(
                "OpenAI API error while processing image",
            )
            return None

        return response.output_parsed

    @staticmethod
    def encode_image(image: bytes) -> str:
        return base64.b64encode(image).decode("utf-8")


class EmbeddingResponse(BaseModel):
    embedding: list[float]

    @classmethod
    async def embed_text(cls, client: openai.AsyncOpenAI, *, text: str) -> Self | None:
        try:
            response = client.embeddings.create(
                model="text-embedding-3-small",
                input=text
            )

        except openai.APIConnectionError:
            logger.error(
                "OpenAI API connection error while embedding text",
            )
            return None
        except openai.APIStatusError:
            logger.error(
                "OpenAI API status error while embedding text",
            )
            return None
        except openai.APIError:
            logger.error(
                "OpenAI API error while embedding text",
            )
            return None

        return cls(embedding=response.data[0].embedding)
