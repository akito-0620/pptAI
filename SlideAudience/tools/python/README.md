# SlideAudience Python Tools

These tools are optional helpers for checking Gemini image prompts, pregenerating comments, and inspecting experiment logs. The PowerPoint add-in does not depend on Python at runtime.

## Setup

```powershell
cd tools/python
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
copy .env.example .env
```

Set `GEMINI_API_KEY` in `.env` before running Gemini checks.

## Scripts

- `scripts/test_gemini_image.py`: send one slide image to Gemini and print generated comments.
- `scripts/pregenerate_comments.py`: generate comments for a folder of PNG files.
- `scripts/analyze_logs.py`: summarize JSONL experiment logs.
