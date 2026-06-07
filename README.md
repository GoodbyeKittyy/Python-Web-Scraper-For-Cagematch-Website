# CageMatch Wrestling Scraper — Improved Python Edition

An improved tool to collect top-rated wrestling matches from [cagematch.net](https://www.cagematch.net) and export them to a formatted Excel workbook.

The original project was an Excel VBA macro. This Python rewrite modernises the approach with a full GUI, stealth browser automation, smarter scraping logic, and richer output. You can access the previous project here: https://github.com/GoodbyeKittyy/Excel-VBA-Web-Scraper-For-Cagematch-Website

See [What's improved over the VBA version](#whats-improved-over-the-vba-version) below.

---

## Setup

```bash
pip install playwright playwright-stealth beautifulsoup4 openpyxl
playwright install chromium
```

> The original VBA version required only Excel. This version requires Python 3.8+ and the packages above.

**Run GUI:**
```bash
python scraper.py
```

**Run CLI (headless, no GUI):**
```bash
python scraper.py --cli --output out.xlsx --min-rating 8.0
```

---

## How It Works

### Step 1 — Fetch Promotions

1. Launch the app and go to **Step 1 · Promotions**
2. Optionally change **Min Rating** (default: 8.00)
3. Click **🔄 Fetch Promotions**

<img width="2556" height="1351" alt="image" src="https://github.com/user-attachments/assets/570cd70d-f346-4219-94b8-a7177702230e" />


The app navigates to `cagematch.net/?id=8&view=promotions&sortby=colRating&sorttype=DESC` and pages through results, stopping once ratings drop below your threshold.




Each promotion appears in a checklist:

| Column | Description |
|--------|-------------|
| ☑ | Checkbox — tick to include this promotion |
| Name | Promotion name |
| Location | City / Country |
| Years | Active years |
| Rating | Average community rating |

After loading, a confirmation dialog shows how many promotions were found and how many are selected, before the browser proceeds to scrape match data.

<img width="2558" height="1352" alt="image" src="https://github.com/user-attachments/assets/ed7f4e35-06e8-4226-bd76-80454e347776" />
---

### Step 2 — Filters & Export

Go to **Step 2 · Filters & Export** and configure:

| Filter | Format | Example |
|--------|--------|---------|
| Date From | `DD.MM.YYYY` | `01.01.2000` |
| Date To | `DD.MM.YYYY` | `31.12.2023` |
| Min WON | Decimal stars | `4.5` (= 4.5★ or more) |
| Min Match Rating | Decimal | `8.00` |

<img width="2558" height="1352" alt="image" src="https://github.com/user-attachments/assets/7e1354a2-6190-430b-99ea-15bf8dcaec29" />


Click **⚡ Scrape Selected & Export to Excel**.

The scraper makes **two passes** per promotion:
- **Pass 1** — sorted by Cagematch community rating (collects matches meeting `Min Match Rating`)
- **Pass 2** — sorted by WON/Meltzer stars (collects additional matches meeting `Min WON` that weren't already captured in Pass 1)

This dual-pass approach means you won't miss highly-rated Meltzer matches that have a lower community score, or vice versa.

---

### Output Excel Workbook
<img width="2558" height="1352" alt="image" src="https://github.com/user-attachments/assets/f58403da-3856-4d77-b99d-e3661af993f1" />

- **One sheet per promotion** (named after the promotion, truncated to 31 characters per Excel's limit)
- Columns: `Date` · `Match Fixture` (hyperlinked to the match page) · `WON Rating` · `Rating`
- Rating cells are color-coded:
  - 🟢 Green (`00B050`): 9.0+
  - 🟡 Light green (`92D050`): 8.5–9.0
  - 🟡 Yellow (`FFFF00`): 8.0–8.5
- Auto-filter on all columns
- Summary row with match count and average rating

---

### Log Tab

A live log panel shows real-time progress: every URL fetched, page number, qualifying match count, and any errors or retries.
<img width="2558" height="470" alt="image" src="https://github.com/user-attachments/assets/6f9ad6e4-93c4-4952-b282-838969a0fa6e" />

---

## Notes

- The scraper uses a **3.5-second delay** between requests (up from 1.5s in older versions) to be respectful to the server.
- Match pages are sorted by rating DESC (`sortby=colRating&sorttype=DESC`) so the scraper can stop early once ratings drop below the threshold.
- WON star ratings may be blank for many matches — cagematch only shows Meltzer ratings when they've been entered by the community.
- Sheet names are truncated to 31 characters (Excel limit).
- Security/CAPTCHA pages are handled automatically where possible; if manual intervention is needed, the browser window stays open for you to solve it.

---

## What's Improved Over the VBA Version

The original VBA macro (`ExtractWebData` in Excel) was a solid starting point, but had several limitations this rewrite addresses:

### 1. Browser Automation Instead of Raw HTTP Requests
The VBA version used `XMLHTTP60` — a simple HTTP client that sent raw GET requests. This works for basic pages but fails on sites that require JavaScript rendering, set cookies dynamically, or challenge bots with Cloudflare/DDoS-Guard pages.

The Python version uses **Playwright** with a real Chromium browser and the `playwright-stealth` library, which masks automation signals (e.g. `navigator.webdriver`). It mimics a real user with a realistic viewport, user-agent, locale, and randomised scrolling and pausing. CAPTCHA and security challenge pages are detected and handled automatically.

### 2. Serialised Single-Thread Browser Worker
The VBA version ran all requests synchronously on the main Excel thread, which caused UI freezes and could crash if requests were made from multiple contexts.

The Python version runs all browser operations on a **single dedicated worker thread** with a job queue (`queue.Queue`). The GUI stays responsive during scraping, and there are no cross-thread Playwright crashes.

### 3. Correct Column Parsing
The VBA version used hardcoded child indices (`TR.Children(1)`, etc.) based on an assumed table structure, which broke silently if CageMatch changed their layout.

The Python version explicitly documents and maps the six-column matchguide table (`# | Date | Fixture | WON | Rating | Votes`) and uses named lookups, making it easier to maintain if the site structure changes.

### 4. Reliable WON Star Extraction
The VBA version didn't attempt to extract WON/Meltzer star ratings at all.

The Python version parses WON ratings with a two-step strategy: it tries BeautifulSoup first, then falls back to a raw HTML regex scan, because BeautifulSoup strips the `★` and `*` characters from `<span class="starRating">` in some encodings. The `_won_to_float()` function handles all known formats: `****1/2`, `***3/4`, `★★★½`, plain decimals, etc.

### 5. Dual-Pass Scraping (Rating + WON)
The VBA version collected matches from a single sorted pass and applied one rating threshold, with a hard `Exit Sub` the moment any match fell below it — meaning it could miss top-rated Meltzer matches that sat lower in the community-rating sort.

The Python version runs two independent sorted passes per promotion (by community rating, then by WON stars) and deduplicates results, so you get the full picture regardless of which rating system you prioritise.

### 6. Multi-Promotion Support with Selection UI
The VBA version scraped a single hardcoded URL from cell `B1` of the worksheet.

The Python version fetches **all qualifying promotions** from CageMatch's promotions index, presents them in a scrollable checklist, and lets you tick/untick which ones to scrape before proceeding.

### 7. Retry Logic and Error Handling
The VBA version had no error handling — a single failed HTTP request would silently stop the loop or crash the macro.

The Python version retries each request up to **3 times** with exponential backoff, logs every failure, and continues to the next promotion rather than aborting the entire run.

### 8. Richer Excel Output
The VBA version wrote plain values and applied basic conditional formatting (black/gold/blue fills).

The Python version produces a more polished workbook: styled headers with a dark blue fill and white bold font, alternating row shading, thin cell borders, hyperlinked fixture names, per-sheet summary rows, and auto-filters on every column.

### 9. CLI Mode
The VBA version required Excel to run.

The Python version includes a `--cli` flag for headless, scriptable use — useful for scheduled runs or CI pipelines.
