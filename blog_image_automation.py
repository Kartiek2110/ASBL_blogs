#!/usr/bin/env python3
"""
ASBL Blog Image Automation
===========================
For each H2/H3 heading section in a WordPress blog post:
  1. Reads the heading + content below it
  2. Searches Pinterest for a relevant image (new tab)
  3. If Pinterest yields nothing → opens ChatGPT in a new tab, pastes the
     content, and waits for it to generate an image
  4. Downloads the image and uploads to WP Media Library
  5. Inserts the image block in WordPress (Gutenberg) right after the section

NO OpenAI API key needed — everything runs through the browser.

Usage:
    export WP_USER="your_wp_username"
    export WP_APP_PASS="xxxx xxxx xxxx xxxx xxxx xxxx"

    python blog_image_automation.py

Requirements:
    pip install playwright httpx
    playwright install chromium
"""

from __future__ import annotations

import asyncio
import base64
import json
import os
import re
import sys
from dataclasses import dataclass
from typing import Optional
from urllib.parse import quote_plus

import httpx

# ─── CONFIG ───────────────────────────────────────────────────────────────────
WP_SITE_URL = "https://asbl.in/blog"
WP_POST_ID  = 9829
WP_USER     = os.environ.get("WP_USER", "")
WP_APP_PASS = os.environ.get("WP_APP_PASS", "")

HEADLESS       = False   # keep False so you can see what's happening
USE_PINTEREST  = True    # try Pinterest first
SECTION_SKIP   = 0       # skip N sections from top
MAX_SECTIONS   = 999     # cap how many sections to process


# ─── DATA ─────────────────────────────────────────────────────────────────────

@dataclass
class BlogSection:
    heading: str
    heading_tag: str
    content_text: str
    block_index: int
    image_url: Optional[str] = None
    image_alt: Optional[str] = None
    image_source: str = ""   # "pinterest" | "chatgpt" | "skipped"


# ─── WORDPRESS REST API ──────────────────────────────────────────────────────

def _wp_auth() -> dict:
    tok = base64.b64encode(f"{WP_USER}:{WP_APP_PASS}".encode()).decode()
    return {"Authorization": f"Basic {tok}"}


def wp_get_post(pid: int) -> dict:
    r = httpx.get(
        f"{WP_SITE_URL}/wp-json/wp/v2/posts/{pid}?context=edit",
        headers=_wp_auth(), timeout=30,
    )
    r.raise_for_status()
    return r.json()


def wp_update_post(pid: int, content: str) -> dict:
    r = httpx.post(
        f"{WP_SITE_URL}/wp-json/wp/v2/posts/{pid}",
        headers={**_wp_auth(), "Content-Type": "application/json"},
        content=json.dumps({"content": content}),
        timeout=60,
    )
    r.raise_for_status()
    return r.json()


def wp_upload_image(data: bytes, filename: str, mime: str = "image/png") -> dict:
    r = httpx.post(
        f"{WP_SITE_URL}/wp-json/wp/v2/media",
        headers={
            **_wp_auth(),
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Content-Type": mime,
        },
        content=data, timeout=120,
    )
    r.raise_for_status()
    return r.json()


# ─── HTML / GUTENBERG HELPERS ─────────────────────────────────────────────────

def plain_text(html: str) -> str:
    t = re.sub(r"<[^>]+>", " ", html)
    t = re.sub(r"&[a-z]+;", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def extract_sections(content: str) -> list[BlogSection]:
    heading_re = re.compile(
        r'<!-- wp:heading(?:\s+\{[^}]*\})?\s*-->(.*?)<!-- /wp:heading -->',
        re.DOTALL,
    )
    matches = list(heading_re.finditer(content))
    if not matches:
        return []

    sections: list[BlogSection] = []
    for idx, m in enumerate(matches):
        inner = m.group(1)
        tag_m = re.search(r'<(h[1-6])[^>]*>', inner, re.IGNORECASE)
        if not tag_m:
            continue
        tag = tag_m.group(1).lower()
        if tag not in ("h2", "h3"):
            continue
        heading = plain_text(inner)
        if not heading:
            continue

        c_start = m.end()
        c_end = matches[idx + 1].start() if idx + 1 < len(matches) else len(content)
        body = plain_text(content[c_start:c_end])

        sections.append(BlogSection(
            heading=heading, heading_tag=tag,
            content_text=body[:1500], block_index=idx,
        ))
    return sections


def make_image_block(url: str, alt: str, media_id: int | None = None) -> str:
    attrs: dict = {"sizeSlug": "large", "linkDestination": "none"}
    if media_id:
        attrs["id"] = media_id
    aj = json.dumps(attrs, separators=(",", ":"))
    id_a = f' id="attachment_{media_id}"' if media_id else ""
    mid = media_id or 0
    return (
        f'\n<!-- wp:image {aj} -->\n'
        f'<figure class="wp-block-image size-large"{id_a}>'
        f'<img src="{url}" alt="{alt}" class="wp-image-{mid}"/>'
        f'</figure>\n<!-- /wp:image -->\n'
    )


def inject_images(content: str, sections: list[BlogSection]) -> str:
    heading_re = re.compile(
        r'(<!-- wp:heading(?:\s+\{[^}]*\})?\s*-->.*?<!-- /wp:heading -->)',
        re.DOTALL,
    )
    matches = list(heading_re.finditer(content))
    insertions: list[tuple[int, str]] = []

    for sec in sections:
        if not sec.image_url:
            continue
        idx = sec.block_index
        if idx >= len(matches):
            continue
        pos = matches[idx + 1].start() if idx + 1 < len(matches) else len(content)
        insertions.append((pos, make_image_block(sec.image_url, sec.image_alt or sec.heading)))

    insertions.sort(key=lambda x: x[0], reverse=True)
    for pos, blk in insertions:
        content = content[:pos] + blk + content[pos:]
    return content


# ─── IMAGE: DOWNLOAD & UPLOAD TO WP ──────────────────────────────────────────

def download_and_upload(image_url: str, name_hint: str) -> tuple[str | None, int | None]:
    try:
        print(f"  [Upload] Downloading: {image_url[:80]}...")
        r = httpx.get(image_url, follow_redirects=True, timeout=60)
        r.raise_for_status()
        ct = r.headers.get("content-type", "image/jpeg").split(";")[0].strip()
        ext = {"image/jpeg": "jpg", "image/png": "png", "image/webp": "webp"}.get(ct, "jpg")
        fname = re.sub(r"[^a-z0-9_-]", "-", name_hint.lower())[:50] + f".{ext}"
        print(f"  [Upload] Uploading to WP as {fname!r}...")
        media = wp_upload_image(r.content, fname, ct)
        wp_url = media.get("source_url") or media.get("guid", {}).get("rendered", "")
        return wp_url, media.get("id")
    except Exception as e:
        print(f"  [Upload] Failed: {e}")
        return None, None


# ─── PINTEREST (NEW TAB) ─────────────────────────────────────────────────────

async def search_pinterest(context, query: str) -> Optional[str]:
    """Open Pinterest in a NEW tab, search, grab the first good image, close tab."""
    print(f"  [Pinterest] Searching: {query!r}")
    page = await context.new_page()
    try:
        url = f"https://www.pinterest.com/search/pins/?q={quote_plus(query)}&rs=typed"
        await page.goto(url, wait_until="domcontentloaded", timeout=30_000)
        await page.wait_for_timeout(4000)

        # Dismiss cookie banner
        for sel in [
            'button:has-text("Accept")', 'button:has-text("Allow all")',
            '[data-test-id="cookie-accept"]',
        ]:
            try:
                btn = await page.query_selector(sel)
                if btn:
                    await btn.click()
                    await page.wait_for_timeout(1000)
                    break
            except Exception:
                pass
        await page.wait_for_timeout(3000)

        # Grab 736-width pinimg images (good quality)
        urls: list[str] = await page.evaluate("""
            () => Array.from(document.querySelectorAll('img[src]'))
                .map(i => i.src)
                .filter(s =>
                    s.includes('pinimg.com') &&
                    !s.includes('avatar') &&
                    !s.includes('75x75') &&
                    !s.includes('60x60') &&
                    s.includes('/736x/')
                )
        """)
        if urls:
            print(f"  [Pinterest] Found {len(urls)} images. Using first.")
            return urls[0]

        # Fallback: any pinimg
        urls2: list[str] = await page.evaluate("""
            () => Array.from(document.querySelectorAll('img[src*="pinimg.com"]'))
                .map(i => i.src)
                .filter(s => !s.includes('75x75') && !s.includes('60x60'))
        """)
        if urls2:
            return urls2[0]

        print("  [Pinterest] No suitable images found.")
        return None
    except Exception as e:
        print(f"  [Pinterest] Error: {e}")
        return None
    finally:
        await page.close()


# ─── CHATGPT IMAGE GENERATION (NEW TAB) ──────────────────────────────────────

async def generate_image_via_chatgpt(context, section: BlogSection) -> Optional[str]:
    """
    Open ChatGPT in a NEW browser tab, paste the section content as a prompt
    asking for an image, wait for the generated image, grab its URL, close tab.
    """
    print(f"  [ChatGPT] Opening new tab for: {section.heading!r}")
    page = await context.new_page()
    try:
        await page.goto("https://chatgpt.com/", wait_until="domcontentloaded", timeout=30_000)
        await page.wait_for_timeout(4000)

        # If redirected to login / auth page, wait for user to log in manually
        current = page.url.lower()
        if "login" in current or "auth" in current:
            print("  [ChatGPT] Login required — please log in manually in the browser window...")
            print("            (waiting up to 2 minutes)")
            try:
                await page.wait_for_url("**/chatgpt.com/**", timeout=120_000)
                await page.wait_for_timeout(3000)
            except Exception:
                print("  [ChatGPT] Timed out waiting for login.")
                return None

        # Decide infographic vs generic image
        data_kws = [
            "cost", "rent", "price", "salary", "budget", "comparison",
            "average", "monthly", "expense", "fee", "rate", "lakh",
        ]
        is_data = any(kw in section.content_text.lower() for kw in data_kws)
        style = "infographic with clean data visualization" if is_data else "photorealistic illustration"

        prompt = (
            f"Generate a {style} image for a blog section.\n\n"
            f"Section Title: {section.heading}\n\n"
            f"Section Content:\n{section.content_text[:800]}\n\n"
            f"Make the image visually engaging, professional, relevant to "
            f"the content. Do NOT include any text/words in the image."
        )

        # Find and fill the prompt input
        # ChatGPT uses a contenteditable div (ProseMirror) or a textarea
        input_sel = '#prompt-textarea, [contenteditable="true"], textarea'
        try:
            box = await page.wait_for_selector(input_sel, timeout=15_000)
        except Exception:
            print("  [ChatGPT] Could not find prompt input box.")
            return None

        await box.click()
        await page.wait_for_timeout(500)

        # Use keyboard to type (handles contenteditable better than fill)
        await box.fill("")
        await page.wait_for_timeout(200)

        # Paste content via clipboard to handle large text reliably
        await page.evaluate(
            """(text) => {
                const el = document.querySelector('#prompt-textarea') ||
                           document.querySelector('[contenteditable="true"]') ||
                           document.querySelector('textarea');
                if (el.tagName === 'TEXTAREA' || el.tagName === 'INPUT') {
                    el.value = text;
                    el.dispatchEvent(new Event('input', {bubbles: true}));
                } else {
                    el.innerText = text;
                    el.dispatchEvent(new Event('input', {bubbles: true}));
                }
            }""",
            prompt,
        )
        await page.wait_for_timeout(1000)

        # Click send button or press Enter
        send_btn = await page.query_selector('[data-testid="send-button"], button[aria-label="Send prompt"]')
        if send_btn:
            await send_btn.click()
        else:
            await page.keyboard.press("Enter")

        print("  [ChatGPT] Prompt sent. Waiting for image generation (up to 120s)...")

        # Poll for an image in the assistant response
        image_url = None
        for attempt in range(40):  # 40 * 3s = 120s max
            await page.wait_for_timeout(3000)

            # Look for generated images in assistant messages
            image_url = await page.evaluate("""
                () => {
                    // Look for images in assistant message containers
                    const containers = document.querySelectorAll(
                        '[data-message-author-role="assistant"], .agent-turn, .markdown'
                    );
                    for (const c of containers) {
                        const imgs = c.querySelectorAll('img');
                        for (const img of imgs) {
                            const src = img.src || '';
                            if (src && (
                                src.includes('oaidalleapi') ||
                                src.includes('openai') ||
                                src.includes('dall-e') ||
                                src.includes('blob:') ||
                                (src.startsWith('https://') && src.includes('image'))
                            )) {
                                return src;
                            }
                        }
                    }
                    // Also check all img tags globally for DALL-E output
                    const allImgs = document.querySelectorAll('img');
                    for (const img of allImgs) {
                        const src = img.src || '';
                        if (src.includes('oaidalleapi') || src.includes('dall-e')) {
                            return src;
                        }
                    }
                    return null;
                }
            """)

            if image_url:
                print(f"  [ChatGPT] Image found after ~{(attempt+1)*3}s: {image_url[:80]}...")
                break

            # Check if ChatGPT is still generating (look for stop button)
            is_generating = await page.evaluate("""
                () => !!document.querySelector(
                    'button[aria-label="Stop generating"], [data-testid="stop-button"]'
                )
            """)
            if attempt > 10 and not is_generating and not image_url:
                # ChatGPT stopped but no image — might have refused or given text
                print("  [ChatGPT] Response complete but no image detected.")
                break

        if not image_url:
            # Last-resort: grab any large image that appeared
            image_url = await page.evaluate("""
                () => {
                    const imgs = Array.from(document.querySelectorAll('img'));
                    for (const img of imgs) {
                        if (img.naturalWidth > 256 && img.src &&
                            !img.src.includes('avatar') && !img.src.includes('logo')) {
                            return img.src;
                        }
                    }
                    return null;
                }
            """)
            if image_url:
                print(f"  [ChatGPT] Fallback image: {image_url[:80]}...")

        return image_url

    except Exception as e:
        print(f"  [ChatGPT] Error: {e}")
        return None
    finally:
        await page.close()


# ─── MAIN ─────────────────────────────────────────────────────────────────────

async def main():
    print("=" * 60)
    print("ASBL Blog Image Automation")
    print("  Pinterest → ChatGPT (browser) → WordPress")
    print("=" * 60)

    if not WP_USER or not WP_APP_PASS:
        print(
            "\n[ERROR] Set WP_USER and WP_APP_PASS env vars.\n"
            "  WP Admin → Users → Edit User → Application Passwords → Add New\n"
        )
        sys.exit(1)

    # ── 1. Fetch post ─────────────────────────────────────────────────────
    print("\n[Step 1] Fetching post from WordPress...")
    try:
        post = wp_get_post(WP_POST_ID)
    except httpx.HTTPStatusError as e:
        print(f"[ERROR] {e.response.status_code}: {e.response.text[:200]}")
        sys.exit(1)

    raw = post.get("content", {}).get("raw", "") or post.get("content", {}).get("rendered", "")
    if not raw:
        print("[ERROR] Post content is empty.")
        sys.exit(1)

    title = post.get("title", {}).get("rendered", "?")
    print(f"  Title: {title!r}  |  Length: {len(raw)} chars")

    # ── 2. Extract sections ───────────────────────────────────────────────
    print("\n[Step 2] Extracting heading sections...")
    all_sections = extract_sections(raw)
    print(f"  Found {len(all_sections)} H2/H3 sections:")
    for i, s in enumerate(all_sections):
        print(f"    [{i+1}] {s.heading_tag.upper()}: {s.heading!r}")

    if not all_sections:
        print("[ERROR] No sections found.")
        sys.exit(1)

    sections = all_sections[SECTION_SKIP:SECTION_SKIP + MAX_SECTIONS]
    print(f"\n  Will process {len(sections)} sections.")

    # ── 3. Open browser → find/generate images ───────────────────────────
    from playwright.async_api import async_playwright

    async with async_playwright() as pw:
        try:
            browser = await pw.chromium.launch(channel="chrome", headless=HEADLESS, args=["--start-maximized"])
        except Exception:
            print("  [Browser] Chrome not found, using bundled Chromium.")
            browser = await pw.chromium.launch(headless=HEADLESS, args=["--start-maximized"])

        ctx = await browser.new_context(viewport={"width": 1440, "height": 900})

        print(f"\n[Step 3] Finding/generating images ({len(sections)} sections)...")

        for i, sec in enumerate(sections):
            print(f"\n{'─'*50}")
            print(f"  [{i+1}/{len(sections)}] {sec.heading_tag.upper()}: {sec.heading!r}")
            print(f"  Content preview: {sec.content_text[:120]}...")

            img_url: str | None = None

            # 3a. Try Pinterest (new tab)
            if USE_PINTEREST:
                query = f"{sec.heading} Hyderabad India infographic"
                img_url = await search_pinterest(ctx, query)
                if img_url:
                    sec.image_source = "pinterest"

            # 3b. Fall back to ChatGPT (new tab)
            if not img_url:
                print("  [Fallback] Pinterest failed → opening ChatGPT...")
                img_url = await generate_image_via_chatgpt(ctx, sec)
                if img_url:
                    sec.image_source = "chatgpt"

            if not img_url:
                print(f"  [SKIP] No image for this section.")
                sec.image_source = "skipped"
                continue

            # 3c. Download & upload to WordPress media library
            wp_url, mid = download_and_upload(img_url, sec.heading)
            if wp_url:
                sec.image_url = wp_url
                sec.image_alt = sec.heading
                print(f"  [OK] Uploaded → {wp_url[:80]}")
            else:
                sec.image_url = img_url
                sec.image_alt = sec.heading
                print(f"  [WARN] Upload failed, using original URL.")

        await browser.close()

    # ── 4. Inject image blocks ────────────────────────────────────────────
    print(f"\n[Step 4] Injecting image blocks...")
    updated = inject_images(raw, sections)
    added = sum(1 for s in sections if s.image_url)
    print(f"  Images to add: {added}")

    if added == 0:
        print("[DONE] No images to add.")
        return

    # ── 5. Save to WordPress ──────────────────────────────────────────────
    print(f"\n[Step 5] Saving updated post to WordPress...")
    try:
        result = wp_update_post(WP_POST_ID, updated)
        print(f"  [OK] Post saved. Link: {result.get('link', 'N/A')}")
    except httpx.HTTPStatusError as e:
        print(f"  [ERROR] {e.response.status_code}: {e.response.text[:300]}")
        sys.exit(1)

    # ── Summary ───────────────────────────────────────────────────────────
    print(f"\n{'='*60}")
    print("DONE! Summary:")
    for s in sections:
        tag = f"[{s.image_source.upper()}]" if s.image_url else "[SKIPPED]"
        url_str = f" {s.image_url[:60]}..." if s.image_url else ""
        print(f"  {s.heading_tag.upper()}: {s.heading!r}")
        print(f"       → {tag}{url_str}")
    print(f"{'='*60}")
    print(f"\nView:  {WP_SITE_URL}/?p={WP_POST_ID}")
    print(f"Edit:  {WP_SITE_URL}/wp-admin/post.php?post={WP_POST_ID}&action=edit")


if __name__ == "__main__":
    asyncio.run(main())
