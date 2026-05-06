#!/usr/bin/env python3
"""
FB Rental Group Scraper
Launches a headed browser, waits for manual Facebook login, then scrapes
posts from configured group(s) within the given time window.
Output: raw_posts.json
"""

import asyncio
import json
import argparse
import sys
from datetime import datetime, timezone, timedelta
from pathlib import Path


def load_config():
    config_path = Path("config.json")
    if not config_path.exists():
        print("ERROR: config.json not found. Run /fb-rentals setup first.")
        sys.exit(1)
    with open(config_path, encoding="utf-8") as f:
        return json.load(f)


async def expand_see_more(article):
    """Click 'See more' buttons inside a post to expand full text."""
    try:
        # Look for any clickable element containing "See more"
        buttons = await article.query_selector_all(
            'div[role="button"], span[role="button"]'
        )
        for btn in buttons:
            try:
                text = (await btn.inner_text()).strip()
                if text.lower() in ("see more", "see more…", "see more..."):
                    await btn.click()
                    await asyncio.sleep(0.4)
                    break
            except Exception:
                continue
    except Exception:
        pass


async def get_post_url(article):
    """Extract the permalink URL of a post."""
    try:
        for selector in [
            'a[href*="/permalink/"]',
            'a[href*="/posts/"]',
            'a[href*="story_fbid"]',
        ]:
            links = await article.query_selector_all(selector)
            for link in links:
                href = await link.get_attribute("href")
                if not href:
                    continue
                if any(k in href for k in ("/permalink/", "/posts/", "story_fbid")):
                    if not href.startswith("http"):
                        href = "https://www.facebook.com" + href
                    # Strip tracking query params, keep the path
                    if "?" in href:
                        href = href.split("?")[0]
                    return href
    except Exception:
        pass
    return None


async def get_post_time(article):
    """Extract ISO timestamp from a post's <time> element."""
    try:
        time_elem = await article.query_selector("time[datetime]")
        if time_elem:
            dt_str = await time_elem.get_attribute("datetime")
            if dt_str:
                return datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
    except Exception:
        pass
    return None


async def get_author(article):
    """Extract post author name."""
    try:
        for selector in ["h2 a", "h3 a", "strong a", 'a[aria-label][role="link"]']:
            elems = await article.query_selector_all(selector)
            for elem in elems:
                name = (await elem.inner_text()).strip()
                if name and 2 <= len(name) <= 80:
                    return name
    except Exception:
        pass
    return "Unknown"


async def scrape_group(page, group_url, cutoff_time):
    """Scrape one Facebook group page. Returns list of raw post dicts."""
    posts = []
    seen_urls = set()

    nav_url = group_url.rstrip("/") + "?sorting_setting=CHRONOLOGICAL"
    print(f"\n  Navigating to: {nav_url}")
    try:
        await page.goto(nav_url, wait_until="domcontentloaded", timeout=30000)
    except Exception as e:
        print(f"  Warning: page load timed out or errored ({e}), continuing anyway.")
    await asyncio.sleep(4)

    stale_scrolls = 0
    max_stale = 6
    prev_seen_count = 0

    while stale_scrolls < max_stale:
        articles = await page.query_selector_all('div[role="article"]')

        new_this_round = 0
        oldest_in_round = None

        for article in articles:
            post_url = await get_post_url(article)
            if not post_url or post_url in seen_urls:
                continue

            post_time = await get_post_time(article)
            if not post_time:
                continue

            seen_urls.add(post_url)
            new_this_round += 1

            if oldest_in_round is None or post_time < oldest_in_round:
                oldest_in_round = post_time

            if post_time < cutoff_time:
                # Too old — mark that we found the boundary, but keep scrolling
                # a bit more to be sure
                continue

            # Expand collapsed text
            await expand_see_more(article)
            await asyncio.sleep(0.2)

            author = await get_author(article)
            raw_text = await article.inner_text()

            post = {
                "url": post_url,
                "author": author,
                "timestamp_iso": post_time.isoformat(),
                "raw_text": raw_text[:6000],
            }
            posts.append(post)
            print(
                f"  [+] {post_time.strftime('%m-%d %H:%M')} | {author[:25]:<25} | {post_url[-50:]}"
            )

        # Decide whether to keep scrolling
        if new_this_round == 0:
            stale_scrolls += 1
        elif oldest_in_round and oldest_in_round < cutoff_time:
            # We have posts older than cutoff — scroll a couple more times then stop
            stale_scrolls += 1
        else:
            stale_scrolls = 0

        # Scroll down
        await page.evaluate("window.scrollBy(0, 1500)")
        await asyncio.sleep(3.5)

    print(f"  Done: {len(posts)} posts within window from this group.")
    return posts


async def main(hours: int):
    config = load_config()
    group_urls = config.get("group_urls", [])
    if not group_urls:
        print("ERROR: No group_urls found in config.json")
        sys.exit(1)

    cutoff_time = datetime.now(timezone.utc) - timedelta(hours=hours)
    print(f"\n{'='*60}")
    print(f"  FB Rentals Scraper")
    print(f"  Window : last {hours} hours (since {cutoff_time.strftime('%Y-%m-%d %H:%M UTC')})")
    print(f"  Groups : {len(group_urls)}")
    print(f"{'='*60}")

    from playwright.async_api import async_playwright

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            slow_mo=80,
            args=["--disable-blink-features=AutomationControlled"],
        )
        context = await browser.new_context(
            viewport={"width": 1366, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
        )
        page = await context.new_page()

        # Navigate to Facebook for manual login
        await page.goto("https://www.facebook.com", wait_until="domcontentloaded")
        await asyncio.sleep(2)

        print("\n" + "="*60)
        print("  ACTION REQUIRED:")
        print("  1. Log in to Facebook in the browser window that just opened.")
        print("  2. Once you can see your feed, come back to this terminal.")
        print("="*60)
        input("  Press Enter when you are logged in... ")
        print()

        all_posts = []
        for i, url in enumerate(group_urls, 1):
            print(f"\nGroup {i}/{len(group_urls)}: {url}")
            try:
                posts = await scrape_group(page, url, cutoff_time)
                all_posts.extend(posts)
            except Exception as e:
                print(f"  ERROR scraping group: {e}")

        await browser.close()

    # Deduplicate by URL (in case same post appeared across groups)
    seen = set()
    unique_posts = []
    for p in all_posts:
        if p["url"] not in seen:
            seen.add(p["url"])
            unique_posts.append(p)

    output_path = Path("raw_posts.json")
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(unique_posts, f, ensure_ascii=False, indent=2, default=str)

    print(f"\n{'='*60}")
    print(f"  Scraping complete.")
    print(f"  Total unique posts found: {len(unique_posts)}")
    print(f"  Saved to: {output_path.resolve()}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Scrape FB rental group posts")
    parser.add_argument(
        "--hours", type=int, default=24, help="Hours to look back (default: 24)"
    )
    args = parser.parse_args()
    asyncio.run(main(args.hours))
