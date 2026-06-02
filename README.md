# UP Revision Petition Orders — Precedent Intelligence Tool

Turns an inaccessible corpus of scanned government orders into a structured, searchable precedent database — built by a litigator who needed it.

## Why this exists

Revision petition orders passed under the **Uttar Pradesh Urban Planning & Development Act** are public records, published on the [Invest UP portal](https://invest.up.gov.in/gos/). But they are scanned-image PDFs with no index, no search, and no way to find how the revisional authority has decided comparable disputes.

I hit this wall directly while acting in a writ matter against GNIDA: the precedent I needed to understand how this authority rules — and what relief it grants — was effectively locked inside hundreds of unsearchable PDFs. So I built a tool to unlock it.

## What it does

1. **Scrapes** every page of orders from `invest.up.gov.in/gos/` (page count detected dynamically).
2. **Downloads** each order PDF locally (skips files already fetched).
3. **Reads** each scanned document with **Claude vision** — pages are rasterised and sent to the model for OCR plus structured extraction.
4. **Builds a searchable index** (`UP_Judgements_Index.xlsx`) with the fields a litigator actually searches on:

| Field | What it captures |
| --- | --- |
| Case Name | Full party names |
| Date of Judgement | Date the order was passed |
| Author of Order | Signing officer and designation |
| Lawyers Who Appeared | Counsel for both sides |
| Brief Facts | 3–5 sentence summary of the dispute |
| Relief Granted | Final outcome and directions |
| Source URL | Original order on the portal |

The result: a corpus of ~900 scanned orders becomes a spreadsheet you can filter, sort, and search in minutes — the difference between *guessing* how an authority decides and *knowing*.

## Stack

- **Python**
- `requests` + `BeautifulSoup4` — scraping the paginated portal
- `PyMuPDF` (fitz) — rasterising scanned PDF pages to PNG (200 DPI)
- **Anthropic API — Claude Haiku (vision)** — OCR + structured field extraction
- `openpyxl` — formatted Excel output
- `progress.json` checkpoint — fully resumable across runs

Built AI-assisted (Claude).

## Setup

**Prerequisites:** Python 3.10+, Git, an [Anthropic API key](https://console.anthropic.com/).

```bash
git clone https://github.com/abhishekdagar-07/up-judgements-scraper.git
cd up-judgements-scraper
pip install -r requirements.txt
export ANTHROPIC_API_KEY="sk-ant-..."   # or the script will prompt
python up_judgements_scraper.py
```

The run is **resumable** — if interrupted, re-run and it picks up from `progress.json`.

## Cost

Claude Haiku (vision) for OCR. Estimated **$3–6 USD** for the full ~900-order corpus.

## Design notes

- **Vision, not text extraction** — the source PDFs are scans, so plain text parsing fails; pages are rendered to images and read by the model.
- **200 DPI** — tuned high enough for reliable OCR on degraded government scans, low enough to keep payloads and cost down.
- **Resumable by default** — a long batch job over a flaky government portal will get interrupted; the checkpoint makes that a non-event.
- **Fields chosen for litigation, not generic metadata** — "relief granted" and "counsel who appeared" are what actually drive whether an order is useful precedent.

## Generalising

The architecture — paginated scrape → download → rasterise → vision extraction → structured index — is portal-agnostic. Pointing it at another court or authority's order repository is a matter of swapping the listing selectors and the extraction schema.

## Legal context

These are final orders in revision petitions under the UP Urban Planning & Development Act, published by the State Government on its official portal.
