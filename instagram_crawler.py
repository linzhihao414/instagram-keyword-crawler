"""
Instagram Keyword Crawler v2
Search hashtags -> collect posts, likes, comments, author, caption, date -> export Excel
Edit keywords.txt to change search terms (hashtags without #).
"""
import asyncio
import re
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from playwright.async_api import async_playwright

# ============ Config ============
KEYWORDS_FILE = Path(__file__).parent / "keywords.txt"
MAX_RESULTS = 100
OUTPUT_DIR = Path(__file__).parent / "导出结果"
HEADLESS = False
# ================================


def load_keywords():
    if KEYWORDS_FILE.exists():
        text = KEYWORDS_FILE.read_text(encoding="utf-8").strip()
        return [k.strip() for k in re.split(r"[,，\n]", text) if k.strip()]
    return ["leatherbag"]


def fmt_num(s):
    s = s.replace(",", "").strip()
    m = re.match(r"([\d.]+)\s*([MK]?)", s, re.I)
    if not m:
        return 0
    v = float(m.group(1))
    u = m.group(2).upper()
    if u == "M":
        return int(v * 1_000_000)
    if u == "K":
        return int(v * 1_000)
    return int(v)


def parse_og_desc(desc):
    """Parse og:description like:
    '41 likes, 3 comments - zaafcollection on June 4, 2025: "caption..."'
    Returns (likes, comments, author, date, caption)
    """
    desc = desc.strip()
    likes = comments = author = date = caption = ""

    # Match likes/comments
    m = re.match(r"([\d.,]+[KM]?)\s*likes,\s*([\d.,]+[KM]?)\s*comments", desc)
    if m:
        likes = m.group(1)
        comments = m.group(2)
        rest = desc[m.end():]
    else:
        m2 = re.match(r"([\d.,]+[KM]?)\s*likes", desc)
        if m2:
            likes = m2.group(1)
            rest = desc[m2.end():]
        else:
            rest = desc

    # Match author and date: " - username on June 4, 2025: "
    m = re.match(r"\s*-\s*([\w.]+)\s+on\s+(.+?):\s*", rest)
    if m:
        author = m.group(1)
        date = m.group(2).strip()
        caption = rest[m.end():].strip()
    else:
        caption = rest.strip()

    return fmt_num(likes), fmt_num(comments), author, date, caption


async def wait_for_login(page):
    print("  If login page appears, please login in the browser...", flush=True)
    for _ in range(90):
        await asyncio.sleep(2)
        url = page.url
        if "instagram.com" in url and "accounts/login" not in url and "facebook.com" not in url:
            return True
    return True


async def crawl_hashtag(page, tag):
    """Collect posts from hashtag page"""
    url = f"https://www.instagram.com/explore/tags/{tag}/"
    await page.goto(url, wait_until="domcontentloaded", timeout=30000)
    await asyncio.sleep(8)

    if "accounts/login" in page.url:
        await wait_for_login(page)
        await page.goto(url, wait_until="domcontentloaded", timeout=30000)
        await asyncio.sleep(8)

    # Scroll to load posts
    for _ in range(20):
        await page.evaluate("window.scrollBy(0, 800)")
        await asyncio.sleep(1)

    links = await page.evaluate("""
        () => [...document.querySelectorAll('a[href*="/p/"], a[href*="/reel/"]')].map(a => a.href.split('?')[0])
    """)
    links = list(dict.fromkeys(links))[:MAX_RESULTS]
    print(f"  Found {len(links)} posts", flush=True)

    results = []
    for i, link in enumerate(links, 1):
        try:
            await page.goto(link, wait_until="domcontentloaded", timeout=20000)
            await asyncio.sleep(4)

            meta = await page.evaluate("""() => {
                const get = (sel) => {
                    const el = document.querySelector(sel);
                    return el ? el.getAttribute('content') || '' : '';
                };
                return {
                    title: get('meta[property="og:title"]'),
                    desc: get('meta[property="og:description"]'),
                    image: get('meta[property="og:image"]'),
                };
            }""")

            desc = meta["desc"]
            likes, comments, author, date, caption = parse_og_desc(desc)

            results.append({
                "keyword": tag,
                "author": author,
                "caption": caption[:500],
                "likes": likes,
                "comments": comments,
                "post_date": date,
                "url": page.url,
                "collect_time": datetime.now().strftime("%Y-%m-%d %H:%M"),
            })

            if i % 10 == 0:
                print(f"    [{i}/{len(links)}] @{author} | {likes} likes", flush=True)

            await asyncio.sleep(0.8)
        except Exception as e:
            print(f"    Skip: {e}", flush=True)

    return results


def save_to_excel(all_data, keyword):
    OUTPUT_DIR.mkdir(exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Posts"

    headers = ["Keyword", "Author", "Caption", "Likes", "Comments",
               "Post Date", "URL", "Collected At"]
    ws.append(headers)
    for r in all_data:
        ws.append([r["keyword"], r["author"], r["caption"], r["likes"], r["comments"],
                    r["post_date"], r["url"], r["collect_time"]])

    for i, w in enumerate([15, 25, 60, 10, 10, 20, 45, 18], 1):
        ws.column_dimensions[chr(64 + i)].width = w

    fname = OUTPUT_DIR / f"Instagram_{keyword}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    wb.save(str(fname))
    print(f"  Saved: {fname}", flush=True)


async def main():
    keywords = load_keywords()
    print("=" * 50, flush=True)
    print("  Instagram Keyword Crawler v2", flush=True)
    print("=" * 50, flush=True)
    print(f"Keywords: {', '.join(keywords)}", flush=True)
    print("(Hashtags - no # needed)", flush=True)

    async with async_playwright() as p:
        # Persistent context: login once, cookies saved in .browser_user
        user_dir = Path(__file__).parent / ".browser_user"
        ctx = await p.chromium.launch_persistent_context(
            user_data_dir=str(user_dir),
            headless=HEADLESS,
            channel="chrome",
            locale="en-US",
            viewport={"width": 1280, "height": 900},
        )
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()

        print("\nOpening Instagram... First run: please login. Later runs: no login needed.", flush=True)
        await page.goto("https://www.instagram.com", wait_until="domcontentloaded")
        await asyncio.sleep(6)

        total = 0
        for kw in keywords:
            print(f"\n=== #{kw} ===", flush=True)
            try:
                results = await crawl_hashtag(page, kw)
                save_to_excel(results, kw)
                total += len(results)
            except Exception as e:
                print(f"  Error: {e}", flush=True)

        print(f"\nDone! Total {total} posts.", flush=True)
        print(f"Output: {OUTPUT_DIR.resolve()}", flush=True)
        await ctx.close()
        input("Press Enter to exit...")


if __name__ == "__main__":
    asyncio.run(main())
