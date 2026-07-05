"""
Find all ASBL blog posts that link to (or mention) Lakeside in internal links.

Scans every blog listed in internal_links.xlsx plus the full sitemap catalog,
visits each post, extracts internal links from the article body, and records
every Lakeside-related link with where it was found.

Run:
    python find_lakeside_links.py
    python find_lakeside_links.py --in "/path/to/internal_links.xlsx" --out lakeside_links_report.xlsx
"""

from __future__ import annotations

import argparse
import html
import re
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from typing import List, Optional, Set, Tuple
from urllib.parse import unquote, urlparse

import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from tqdm import tqdm

from scrape_asbl_blogs import _clean_text, fetch, make_session

SEARCH_TERM = "lakeside"
INTERNAL_HOST = "asbl.in"

HEADERS = [
    "Source Blog Title",
    "Source Blog URL",
    "Source Blog Date",
    "Lakeside Link Text",
    "Lakeside Link URL",
    "Match In",
    "Link Type",
    "Notes",
]


@dataclass
class LakesideHit:
    source_title: str
    source_url: str
    source_date: str
    link_text: str
    link_url: str
    match_in: str
    link_type: str
    notes: str


def _is_internal(url: str) -> bool:
    try:
        host = urlparse(url).netloc.lower()
    except ValueError:
        return False
        print(host)
    return host == INTERNAL_HOST or host.endswith("." + INTERNAL_HOST)


def _extract_title(soup: BeautifulSoup) -> str:
    h1 = soup.find("h1", class_="entry-title") or soup.find("h1")
    if h1:
        return _clean_text(h1.get_text())
    meta = soup.find("meta", property="og:title")
    if meta and meta.get("content"):
        return _clean_text(meta["content"])
    return ""


def _extract_date(soup: BeautifulSoup) -> str:
    meta = soup.find("meta", property="article:published_time")
    if meta and meta.get("content"):
        return meta["content"][:10]
    t = soup.find("time")
    if t and t.get("datetime"):
        return t["datetime"][:10]
    return ""


def _contains_lakeside(value: str) -> bool:
    return SEARCH_TERM in unquote(value or "").lower()


def _match_location(link_text: str, link_url: str) -> str:
    parts: List[str] = []
    if _contains_lakeside(link_url):
        parts.append("Link URL")
    if _contains_lakeside(link_text):
        parts.append("Link Text")
    return " + ".join(parts) if parts else "Page body"


def _scan_html(
    page_url: str,
    page_html: str,
    source_date: str = "",
) -> List[LakesideHit]:
    soup = BeautifulSoup(page_html, "lxml")
    title = _extract_title(soup) or page_url
    date = source_date or _extract_date(soup)

    content = soup.select_one(".entry-content") or soup.select_one(".post_content")
    if content is None:
        return []

    hits: List[LakesideHit] = []
    seen: Set[Tuple[str, str]] = set()

    # Blog itself is about Lakeside.
    if _contains_lakeside(title) or _contains_lakeside(page_url):
        key = ("__self__", page_url)
        if key not in seen:
            seen.add(key)
            hits.append(
                LakesideHit(
                    source_title=title,
                    source_url=page_url,
                    source_date=date,
                    link_text="(This blog is about Lakeside)",
                    link_url=page_url,
                    match_in="Blog title/URL",
                    link_type="Lakeside blog",
                    notes="Blog title or slug references Lakeside",
                )
            )

    for anchor in content.find_all("a", href=True):
        href = html.unescape(anchor["href"].strip())
        text = _clean_text(anchor.get_text(" ", strip=True))
        if not _contains_lakeside(href) and not _contains_lakeside(text):
            continue

        key = (text, href)
        if key in seen:
            continue
        seen.add(key)

        internal = _is_internal(href)
        hits.append(
            LakesideHit(
                source_title=title,
                source_url=page_url,
                source_date=date,
                link_text=text or href,
                link_url=href,
                match_in=_match_location(text, href),
                link_type="Internal" if internal else "External",
                notes="Found in article body",
            )
        )

    # Plain-text mention in body (no lakeside anchor).
    body_text = _clean_text(content.get_text(" ", strip=True)).lower()
    if _contains_lakeside(body_text) and not any(h.link_url for h in hits if h.link_type != "Lakeside blog"):
        if not hits:
            hits.append(
                LakesideHit(
                    source_title=title,
                    source_url=page_url,
                    source_date=date,
                    link_text="(Text mention only — no clickable Lakeside link)",
                    link_url="",
                    match_in="Article body text",
                    link_type="Text mention",
                    notes="Lakeside mentioned in blog body without a dedicated link",
                )
            )

    return hits


def scan_blog_url(session, url: str, source_date: str = "") -> Tuple[Optional[str], List[LakesideHit]]:
    page_html = fetch(session, url, retries=4, timeout=35)
    if not page_html:
        return None, []
    if SEARCH_TERM not in page_html.lower():
        return url, []
    return url, _scan_html(url, page_html, source_date=source_date)


def collect_urls(input_path: str, catalog_path: str) -> pd.DataFrame:
    """Return unique blog URLs with optional dates from input + full catalog."""
    rows = []

    input_df = pd.read_excel(input_path)
    for _, row in input_df.iterrows():
        blog_url = str(row.get("Blog Link", "") or "").strip()
        if blog_url.startswith("http"):
            rows.append(
                {
                    "url": blog_url,
                    "date": str(row.get("Blog Date", "") or "")[:10],
                    "from": "internal_links.xlsx",
                }
            )

    try:
        catalog_df = pd.read_excel(catalog_path, sheet_name="All Posts")
        for _, row in catalog_df.iterrows():
            blog_url = str(row.get("URL", "") or "").strip()
            if blog_url.startswith("http"):
                rows.append(
                    {
                        "url": blog_url,
                        "date": str(row.get("Published Date", "") or "")[:10],
                        "from": "sitemap catalog",
                    }
                )
    except Exception:
        pass

    frame = pd.DataFrame(rows)
    frame = frame.drop_duplicates(subset=["url"], keep="first")
    return frame


def write_report(hits: List[LakesideHit], scanned: int, failed: List[str], path: str) -> None:
    wb = Workbook()

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E78")
    wrap = Alignment(wrap_text=True, vertical="top")
    link_font = Font(color="0563C1", underline="single")

    ws = wb.active
    ws.title = "Lakeside Links"
    ws.append(HEADERS)
    for col_idx in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    unique_blogs = {h.source_url for h in hits}
    for hit in hits:
        ws.append(
            [
                hit.source_title,
                hit.source_url,
                hit.source_date,
                hit.link_text,
                hit.link_url,
                hit.match_in,
                hit.link_type,
                hit.notes,
            ]
        )
        row_idx = ws.max_row
        for col_idx in range(1, len(HEADERS) + 1):
            ws.cell(row=row_idx, column=col_idx).alignment = wrap
        url_cell = ws.cell(row=row_idx, column=2)
        if url_cell.value:
            url_cell.hyperlink = str(url_cell.value)
            url_cell.font = link_font
        link_cell = ws.cell(row=row_idx, column=5)
        if link_cell.value and str(link_cell.value).startswith("http"):
            link_cell.hyperlink = str(link_cell.value)
            link_cell.font = link_font

    widths = {
        "Source Blog Title": 52,
        "Source Blog URL": 62,
        "Source Blog Date": 14,
        "Lakeside Link Text": 40,
        "Lakeside Link URL": 62,
        "Match In": 22,
        "Link Type": 14,
        "Notes": 36,
    }
    for col_idx, header in enumerate(HEADERS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = widths.get(header, 20)
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions

    ws_summary = wb.create_sheet("Summary")
    ws_summary.append(["Metric", "Value"])
    for col_idx in (1, 2):
        cell = ws_summary.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
    summary = [
        ("Search term", SEARCH_TERM),
        ("Blogs scanned", scanned),
        ("Blogs with Lakeside links/mentions", len(unique_blogs)),
        ("Total Lakeside link rows", len(hits)),
        ("Failed to fetch", len(failed)),
    ]
    for idx, (metric, value) in enumerate(summary, start=2):
        ws_summary.cell(row=idx, column=1, value=metric)
        ws_summary.cell(row=idx, column=2, value=value)
    ws_summary.column_dimensions["A"].width = 34
    ws_summary.column_dimensions["B"].width = 20

    if failed:
        ws_failed = wb.create_sheet("Failed URLs")
        ws_failed.append(["URL"])
        ws_failed.cell(1, 1).font = header_font
        ws_failed.cell(1, 1).fill = header_fill
        for url in failed:
            ws_failed.append([url])
        ws_failed.column_dimensions["A"].width = 70

    wb.save(path)


def main() -> int:
    parser = argparse.ArgumentParser(description="Find blogs linking to Lakeside.")
    parser.add_argument(
        "--in",
        dest="input_path",
        default="/Users/kartiekeybhardwaj/Downloads/rewindDesgin/internal_links.xlsx",
    )
    parser.add_argument(
        "--catalog",
        default="asbl_post_sitemap_clusters.xlsx",
        help="Full blog catalog for comprehensive scan.",
    )
    parser.add_argument(
        "--out",
        default="lakeside_links_report.xlsx",
    )
    parser.add_argument("--workers", type=int, default=6)
    args = parser.parse_args()

    url_frame = collect_urls(args.input_path, args.catalog)
    print(f"Scanning {len(url_frame)} unique blog URLs for '{SEARCH_TERM}'...", file=sys.stderr)

    hits: List[LakesideHit] = []
    failed: List[str] = []

    def _job(row) -> Tuple[str, Optional[str], List[LakesideHit]]:
        session = make_session()
        url = row["url"]
        try:
            status, result = scan_blog_url(session, url, source_date=row.get("date", ""))
            return url, status, result
        except Exception:
            return url, None, []

    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        futures = [pool.submit(_job, row) for _, row in url_frame.iterrows()]
        for fut in tqdm(as_completed(futures), total=len(futures), desc="Scanning blogs"):
            url, status, result = fut.result()
            if status is None:
                failed.append(url)
            if result:
                hits.extend(result)

    # Stable sort: source blog, then link url
    hits.sort(key=lambda h: (h.source_title.lower(), h.link_url.lower()))

    write_report(hits, scanned=len(url_frame), failed=failed, path=args.out)

    unique_blogs = len({h.source_url for h in hits})
    print(
        f"Done. {unique_blogs} blogs with Lakeside links/mentions, "
        f"{len(hits)} total rows -> {args.out}",
        file=sys.stderr,
    )
    if failed:
        print(f"{len(failed)} URLs failed to fetch (see Failed URLs sheet).", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
