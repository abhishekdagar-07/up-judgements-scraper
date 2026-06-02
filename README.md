# UP Government Orders — Judgement Downloader & Analyser

A Python tool that scrapes all revision petition orders from the [Uttar Pradesh Invest UP portal](https://invest.up.gov.in/gos/) and creates a structured Excel index using AI-powered OCR.

## What it does

1. **Scrapes** all pages of government orders from invest.up.gov.in/gos/
2. **Downloads** every judgement PDF into a local folder (`UP_Judgements/`)
3. **OCR + Analysis** — converts scanned PDF pages to images and sends them to Claude AI (vision) for extraction
4. **Creates an Excel file** (`UP_Judgements_Index.xlsx`) with these columns:

| Column | Description |
|---|---|
| Case Name | Full party names (e.g. "Sanjay Rastogi vs GNIDA") |
| Date of Judgement | Date the order was passed |
| Author of Order | Signing officer and designation |
| Lawyers Who Appeared | Advocates for both sides |
| Brief Facts | 3–5 sentence summary of the dispute |
| Relief Granted | Final outcome and directions |
| PDF File | Local filename |
| Source URL | Original download URL |

## Setup

### Prerequisites
- Python 3.10+
- An [Anthropic API key](https://console.anthropic.com/) (for Claude AI)

### Install dependencies

```bash
pip install -r requirements.txt
```

### Set your API key (optional — the script will prompt if not set)

```bash
# Windows PowerShell
$env:ANTHROPIC_API_KEY = "sk-ant-..."

# Linux / macOS
export ANTHROPIC_API_KEY="sk-ant-..."
```

## Usage

```bash
python up_judgements_scraper.py
```

The script is **resumable** — if interrupted, re-run it and it picks up where it left off (progress is saved in `progress.json`).

## Cost

Uses Claude Haiku (vision) for OCR. Estimated cost for all ~900 judgements: **$3–6 USD**.

## Output

- `UP_Judgements/` — folder with all downloaded PDFs
- `UP_Judgements_Index.xlsx` — formatted Excel spreadsheet with extracted data

## Legal Context

These are final orders passed in revision petitions filed under the **Uttar Pradesh Urban Planning and Development Act**, published by the UP state government on their official portal.
