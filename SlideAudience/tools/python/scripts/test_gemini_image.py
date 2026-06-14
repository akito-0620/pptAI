import argparse
import base64
import json
import os
from pathlib import Path

from dotenv import load_dotenv
from google import genai
from google.genai import types


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("image", type=Path)
    parser.add_argument("--text", default="")
    args = parser.parse_args()

    load_dotenv()
    api_key = os.environ["GEMINI_API_KEY"]
    model = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")

    image_bytes = args.image.read_bytes()
    client = genai.Client(api_key=api_key)
    prompt = (
        "このスライドを見た観客が自然に思いそうな短い日本語コメントを"
        "最大3件、JSONのみで返してください。"
        f"\nスライド内テキスト:\n{args.text}"
    )

    response = client.models.generate_content(
        model=model,
        contents=[
            types.Part.from_text(text=prompt),
            types.Part.from_bytes(data=image_bytes, mime_type="image/png"),
        ],
        config=types.GenerateContentConfig(response_mime_type="application/json"),
    )

    print(json.dumps(json.loads(response.text), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
