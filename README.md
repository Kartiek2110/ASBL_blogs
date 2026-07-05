# ASBL Blog Link Scraper

Crawls every post on [`https://asbl.in/blog/`](https://asbl.in/blog/) across all
paginated listing pages (1..43 by default), opens each blog, extracts every
`<a style="color: #0000ff;" href="...">…</a>` link in the article body, and
splits them into:

- **Internal links** – any link whose host is `asbl.in` (e.g.
  `https://asbl.in/blog/how-to-maintain-the-perfect-balcony-garden/`)
- **External links** – everything else (e.g.
  `https://www.coohom.com/in/article/2bhk-interior-design-ideas-maximize-space-style`)

Results are written to an Excel workbook (`asbl_blogs.xlsx` by default). Each blog
gets at least one row; posts with multiple styled links get extra rows so every
URL is a clickable hyperlink.

## Output columns

| Column | Description |
| --- | --- |
| Blog Name | Post title (from `<h1 class="entry-title">` / `og:title`) |
| Blog Link | Canonical post URL |
| Blog Date | Published date (`YYYY-MM-DD` when available, otherwise raw listing date) |
| Internal Link | Clickable hyperlink (one per row when a post has several) |
| External Link | Clickable hyperlink (one per row when a post has several) |
| # Internal | Count of internal styled links found |
| # External | Count of external styled links found |

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run

```bash
# Default: pages 1..43, 8 parallel workers, output -> asbl_blogs.xlsx
python scrape_asbl_blogs.py

# Custom range / output
python scrape_asbl_blogs.py --start 1 --end 43 --workers 10 --out asbl_blogs.xlsx

# Debug: only the first 5 posts of page 1
python scrape_asbl_blogs.py --start 1 --end 1 --limit 5 --out preview.xlsx

# Rebuild clickable links from an existing export (no re-scrape)
python scrape_asbl_blogs.py --from-excel asbl_blogs.xlsx --out asbl_blogs.xlsx
```

## How it works

1. Walks `https://asbl.in/blog/page/{n}/` for `n = start..end` and collects
   post URLs from `.insight-content` cards (title + date scraped from the
   listing too).
2. Fetches every post concurrently (`ThreadPoolExecutor`) with retry/backoff on
   5xx / 429.
3. For each post, extracts the title (`<h1 class="entry-title">` →
   `og:title` → `<title>`), the date (`article:published_time` →
   `<time datetime>` → `datePublished` JSON-LD → `.posted-on`), and every
   anchor matching `style="color: #0000ff;"` in the post body.
4. Classifies each anchor by hostname (`asbl.in` ⇒ internal, else external).
5. Writes everything to a styled `.xlsx` with frozen header row and auto filter.

## Latest run

- 43 listing pages → **553 unique blog posts** scraped, 0 failures.
- 1,369 internal styled links / 2,036 external styled links collected.
