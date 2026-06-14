import argparse
import json
import subprocess
import sys
from pathlib import Path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("image_dir", type=Path)
    parser.add_argument("--output", type=Path, default=Path("pregenerated_comments.jsonl"))
    args = parser.parse_args()

    images = sorted(args.image_dir.glob("*.png"))
    with args.output.open("w", encoding="utf-8") as output:
        for image in images:
            result = subprocess.run(
                [sys.executable, str(Path(__file__).with_name("test_gemini_image.py")), str(image)],
                check=True,
                capture_output=True,
                text=True,
                encoding="utf-8",
            )
            output.write(json.dumps({"image": str(image), "result": json.loads(result.stdout)}, ensure_ascii=False))
            output.write("\n")


if __name__ == "__main__":
    main()
