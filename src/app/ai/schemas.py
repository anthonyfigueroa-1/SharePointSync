docx_schema = {
    "format": {
        "type": "json_schema",
        "name": "docx_data",
        "strict": True,
        "schema": {
            "type": "object",
            "properties": {
                "image_metadata": {
                    "type": "array",
                    "description": "One string per image, in the same order the images were provided. Each string must contain a detailed description of that image and end with a final sentence explaining where the image fits into the document text.",
                    "items": {"type": "string"},
                },
                "keywords": {
                    "type": "array",
                    "description": "An array of 1 to 5 high-value keywords or short keyword phrases chosen to maximize semantic document retrieval. Keywords may include strongly implied topics even when the exact words do not appear in the text or images.",
                    "items": {"type": "string"},
                },
                "summary": {
                    "type": ["string", "null"],
                    "description": "A high value summary of the text, and photos, being given.",
                },
            },
            "additionalProperties": False,
            "required": ["image_metadata", "keywords", "summary"],
        },
    }
}
