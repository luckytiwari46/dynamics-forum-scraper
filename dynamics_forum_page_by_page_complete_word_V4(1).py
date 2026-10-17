import os
import re
import time
import hashlib
from pathlib import Path
from urllib.parse import urljoin, urlparse, parse_qs, urlencode, urlunparse

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


BASE = Path(__file__).resolve().parent

DEFAULT_START_URL = "https://community.dynamics.com/forums/thread/"

OUTPUT_DIR = BASE / "dynamics_forum_page_word_files_V5"
PROFILE_DIR = BASE / "browser_profile_forum_word_V5"
MAX_TOPICS_PER_PAGE = 100
PAGE_WAIT = 1000
LOGIN_WAIT = 90
TOPIC_RETRIES = 4

BLOCKED_DOMAINS = [
    "google-analytics.com",
    "googletagmanager.com",
    "clarity.ms",
    "bing.com",
    "doubleclick.net",
    "adnxs.com",
    "scorecardresearch.com",
    "browser.pipe.aria.microsoft.com",
    "vortex.data.microsoft.com",
]


def clean(s):
    return re.sub(r"\s+", " ", s or "").strip()


def safe_name(s, max_len=100):
    s = clean(s)
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", s)
    s = re.sub(r"_+", "_", s).strip(" ._")
    return (s[:max_len] or "untitled")


def archive_page_number(url):
    try:
        return int(parse_qs(urlparse(url).query, keep_blank_values=True).get("page", ["1"])[0])
    except (TypeError, ValueError):
        return 1


def normalize_archive_url(url):
    """Normalize user-entered archive URL."""
    url = (url or "").strip()
    if not url:
        raise ValueError("No URL was entered.")
    if not url.lower().startswith(("https://", "http://")):
        url = "https://" + url
    parsed = urlparse(url)
    if not parsed.netloc:
        raise ValueError("That does not look like a valid URL.")
    if "community.dynamics.com" not in parsed.netloc.lower():
        print("WARNING: This URL is not on community.dynamics.com. The extractor may not match its page layout.")
    return url


def build_page_url(base_url, page_no):
    """Construct or replace the page query parameter for the given page number while preserving all other parameters and order."""
    base_url = (base_url or "").strip()
    fragment = ""
    if "#" in base_url:
        base_url, fragment = base_url.split("#", 1)
        fragment = "#" + fragment

    if re.search(r'([?&]page=)[^&]*', base_url, re.I):
        new_url = re.sub(r'([?&]page=)[^&]*', rf'\g<1>{page_no}', base_url, flags=re.I)
    else:
        separator = "&" if "?" in base_url else "?"
        new_url = f"{base_url}{separator}page={page_no}"

    return new_url + fragment


def absolute(href, base):
    if not href:
        return ""
    return urljoin(base, href)


def is_topic_url(url):
    u = (url or "").lower()
    return (
        "community.dynamics.com/forums/thread/details" in u
        or "community.dynamics.com/forums/thread/?" in u
    ) and ("threadid=" in u or "/details/" in u)


def canonical_topic_url(url):
    """Return the public thread URL only; never follow reply/login action links."""
    if not url:
        return ""
    try:
        p = urlparse(url)
        q = parse_qs(p.query, keep_blank_values=True)
        if "threadid" not in q:
            return url.split("#")[0]
        # Keep only the thread id. In particular, drop action=reply, which
        # was causing the scraper to revisit Microsoft sign-in/reply routes.
        tid = q["threadid"][0]
        return urlunparse((p.scheme, p.netloc, "/forums/thread/details/", "", urlencode({"threadid": tid}), ""))
    except Exception:
        return url.split("#")[0]


def is_login_page(page):
    """Detect a real authentication interstitial, not the normal forum header.

    Dynamics Community pages themselves contain a visible 'Sign in' link/header,
    even when the thread is completely public. We therefore require either an
    actual Microsoft auth URL or login-form controls/strong authentication markers.
    """
    try:
        u = (page.url or "").lower()
        if any(x in u for x in [
            "login.microsoftonline.com",
            "login.live.com",
            "account.live.com",
            "/oauth/",
            "oauth2/",
            "authorize.srf",
        ]):
            return True

        return page.evaluate("""() => {
            const authControls = document.querySelectorAll(
                "input[type='password'], " +
                "input[name='passwd'], " +
                "input[name='loginfmt'], " +
                "input[type='email'][autocomplete='username'], " +
                "input[autocomplete='current-password'], " +
                "form[action*='login.microsoftonline'], " +
                "form[action*='oauth']"
            ).length;
            if (authControls > 0) return true;

            const body = (document.body ? document.body.innerText : '').toLowerCase();
            const strongAuth = [
                "enter your password",
                "enter password",
                "use your password",
                "pick an account",
                "we need to verify your identity",
                "sign in to your microsoft account"
            ];
            return strongAuth.some(p => body.includes(p));
        }""")
    except Exception:
        return False


def wait_until_topic_ready(page, expected_url):
    """Wait for the requested thread, but never treat the sign-in shell as content."""
    expected_tid = ""
    try:
        expected_tid = parse_qs(urlparse(expected_url).query).get("threadid", [""])[0].lower()
    except Exception:
        pass
    deadline = time.time() + LOGIN_WAIT
    warned = False
    while time.time() < deadline:
        try:
            curr_url = (page.url or "").lower()
            if expected_tid and expected_tid not in curr_url:
                pass
            elif not is_login_page(page):
                ready = page.evaluate("""() => {
                    const hasThread = document.querySelector(".msd-forum-thread-container, [class*='forum-post'], [class*='thread-post']");
                    const body = (document.body ? document.body.innerText : '').replace(/\\s+/g, ' ').trim();
                    return !!hasThread || (body.length > 250 && (body.includes("Thread content") || body.length > 500));
                }""")
                if ready:
                    return True
            if is_login_page(page) and not warned:
                print("      [LOGIN DETECTED] Waiting for the public thread/session to recover...")
                warned = True
        except Exception:
            pass
        page.wait_for_timeout(250)
    return False


def wait_for_forum(page):
    """Quickly ensure the page DOM and forum container elements are loaded."""
    try:
        page.wait_for_load_state("domcontentloaded", timeout=30000)
    except Exception:
        pass
    try:
        for _ in range(25):
            ready = page.evaluate("""() => {
                return !!document.querySelector(".msd-forum-thread-container, [class*='forum-post'], [class*='thread-post'], .dc-thread-pane-item, article");
            }""")
            if ready:
                break
            page.wait_for_timeout(100)
    except Exception:
        pass


def scroll_full(page):
    """Fast, smooth scroll to trigger lazy loading of content and images."""
    try:
        page.evaluate("""async () => {
            let lastHeight = document.body.scrollHeight;
            for (let i = 0; i < 8; i++) {
                window.scrollBy(0, Math.max(800, window.innerHeight * 1.5));
                await new Promise(r => setTimeout(r, 60));
                let h = document.body.scrollHeight;
                if (window.scrollY + window.innerHeight >= h && h === lastHeight) break;
                lastHeight = h;
            }
            window.scrollTo(0, 0);
        }""")
    except Exception:
        pass


def click_load_more_replies(page):
    """Quickly check and click 'load more' reply buttons via in-browser evaluation."""
    for _ in range(10):
        try:
            clicked = page.evaluate("""() => {
                const patterns = /(load|show|view)\\s+(more|all)|more\\s+repl(?:y|ies)|show\\s+\\d+\\s+more|view\\s+\\d+\\s+more/i;
                const buttons = document.querySelectorAll("button, a, [role='button']");
                for (const el of buttons) {
                    if (!el.offsetParent && el.offsetWidth === 0 && el.offsetHeight === 0) continue;
                    const txt = (el.innerText || '').replace(/\\s+/g, ' ').trim();
                    const aria = (el.getAttribute('aria-label') || '').replace(/\\s+/g, ' ').trim();
                    const label = txt + ' ' + aria;
                    if (patterns.test(label)) {
                        el.scrollIntoView({block: 'center'});
                        el.click();
                        return true;
                    }
                }
                return false;
            }""")
            if not clicked:
                break
            page.wait_for_timeout(500)
        except Exception:
            break


def collect_page_cards(page, page_no):
    """Collect one canonical card per thread. Ignore reply/action links."""
    js_collect = """
    (base) => {
        function clean(s) { return (s || '').replace(/\\s+/g, ' ').trim(); }
        const selectors = [
            "a[href*='/forums/thread/details']",
            "a[href*='/forums/thread/?threadid=']",
            "a[href*='threadid=']"
        ];
        const elements = document.querySelectorAll(selectors.join(', '));
        const results = [];
        const seen = new Set();

        for (const a of elements) {
            let rawHref = a.getAttribute('href') || '';
            if (!rawHref) continue;
            let absUrl;
            try { absUrl = new URL(rawHref, base).href; } catch(e) { continue; }
            
            let u = absUrl.toLowerCase();
            let isTopic = (u.includes("community.dynamics.com/forums/thread/details") ||
                           u.includes("community.dynamics.com/forums/thread/?")) &&
                          (u.includes("threadid=") || u.includes("/details/"));
            if (!isTopic) continue;

            let txt = clean(a.innerText) || clean(a.getAttribute('title'));
            if (/^(reply|replies|reply to this post)$/i.test(txt)) continue;

            let tid = '';
            try {
                const urlObj = new URL(absUrl);
                tid = (urlObj.searchParams.get("threadid") || "").toLowerCase();
                if (!tid && urlObj.pathname.includes("/details/")) {
                    const parts = urlObj.pathname.split("/details/");
                    if (parts[1]) tid = parts[1].split('/')[0].split('?')[0].toLowerCase();
                }
            } catch(e) {}

            if (!tid || seen.has(tid)) continue;
            if (!txt || txt.length < 3) continue;

            let card = a.closest('.dc-thread-pane-item, article, [class*="card"], [class*="thread"], [class*="topic"], li');
            let cardText = card ? clean(card.innerText) : '';
            let replies = "";
            let m = cardText.match(/(?<!\\d)([\\d,]+)\\s+Replies\\b/i);
            if (m) replies = m[1].replace(/,/g, '');

            seen.add(tid);
            results.push({
                title: txt,
                rawUrl: absUrl,
                threadid: tid,
                replies: replies
            });
            if (results.length >= 100) break;
        }
        return results;
    }
    """
    try:
        raw_cards = page.evaluate(js_collect, page.url)
    except Exception:
        raw_cards = []

    cards = []
    for c in raw_cards:
        canonical_url = canonical_topic_url(c["rawUrl"])
        if is_topic_url(canonical_url):
            cards.append({
                "page": page_no,
                "title": c["title"],
                "url": canonical_url,
                "replies": c["replies"],
                "threadid": c["threadid"]
            })
            if len(cards) >= MAX_TOPICS_PER_PAGE:
                break
    return cards


def topic_title(page):
    try:
        t = page.evaluate("""() => {
            const h = document.querySelector("h1, main h1, [role='main'] h1, .thread-post-title");
            return h ? h.innerText.replace(/\\s+/g, ' ').trim() : document.title;
        }""")
        if t:
            return clean(t)
    except Exception:
        pass
    return clean(page.title())


def get_post_candidates(page):
    """Fast, accurate in-browser post candidate detector matching original layout rules."""
    js = """
    () => {
        function clean(s) { return (s || '').replace(/\\s+/g, ' ').trim(); }
        const selectors = [
            "article",
            "[role='article']",
            "[class*='forum-post']",
            "[class*='thread-post']",
            "[class*='post-container']",
            "[class*='reply-container']",
            "[class*='message-container']",
            "[class*='discussion-post']",
            "[class*='comment-container']"
        ];
        function looksLikePost(el) {
            const meta = ((el.className || '') + ' ' + (el.id || '') + ' ' + (el.getAttribute('data-testid') || '') + ' ' + (el.getAttribute('role') || '')).toLowerCase();
            const txt = clean(el.innerText);
            if (txt.length < 20) return false;
            const strong = /(post|reply|comment|message|discussion)/.test(meta);
            const action = /\\b(reply|like|report|share)\\b/i.test(txt);
            const authorish = /\\b(posted|posted on|last replied|replied|asked by|answered by)\\b/i.test(txt);
            return strong && (action || authorish || txt.length > 80);
        }
        let raw = [];
        const seen = new Set();
        for (const sel of selectors) {
            const els = document.querySelectorAll(sel);
            for (let i = 0; i < els.length; i++) {
                const el = els[i];
                if (!looksLikePost(el)) continue;
                if (!el.dataset.forumScrapeId) el.dataset.forumScrapeId = Math.random().toString(36).slice(2);
                const key = el.dataset.forumScrapeId;
                if (seen.has(key)) continue;
                seen.add(key);
                raw.push(el);
            }
        }
        const finalPosts = [];
        for (const el of raw) {
            const txt = clean(el.innerText);
            if (txt.length < 30) continue;
            let duplicate = false;
            for (const other of raw) {
                if (other === el) continue;
                const ot = clean(other.innerText);
                if (ot.length >= txt.length && txt === ot) { duplicate = true; break; }
            }
            if (!duplicate) finalPosts.push(el);
        }
        finalPosts.sort((a, b) => {
            const boxA = a.getBoundingClientRect();
            const boxB = b.getBoundingClientRect();
            return (boxA ? boxA.top : 1e9) - (boxB ? boxB.top : 1e9);
        });
        for (let i = 0; i < finalPosts.length; i++) {
            finalPosts[i].dataset.scrapePostIndex = String(i + 1);
        }
        return finalPosts.length;
    }
    """
    try:
        count = page.evaluate(js)
        return [page.locator(f"[data-scrape-post-index='{i}']").first for i in range(1, count + 1)]
    except Exception:
        return []


def fallback_post_containers(page):
    """If the site's class names change, identify likely top-level content blocks."""
    js = """
    () => {
      const out = [];
      const all = [...document.querySelectorAll('main, [role="main"], body *')];
      for (const e of all) {
        const t = (e.innerText || '').trim();
        if (t.length < 80 || t.length > 30000) continue;
        const c = ((e.className || '') + ' ' + (e.id || '')).toLowerCase();
        if (!/(post|reply|comment|message|discussion)/.test(c)) continue;
        out.push(e);
      }
      return out.length;
    }
    """
    try:
        count = page.evaluate(js)
        if count:
            return get_post_candidates(page)
    except Exception:
        pass
    return []


def extract_blocks_from_post(post, page):
    """
    Extract text, links, tables, images and code in DOM order.
    Images are tagged with unique data-scrape-img-id for instant element lookup.
    """
    js = """
    (root) => {
      function clean(s) { return (s || '').replace(/\\s+/g,' ').trim(); }
      const blocks = [];
      let imgCounter = 0;

      function walk(node) {
        if (!node) return;

        if (node.nodeType === Node.TEXT_NODE) {
          const t = clean(node.nodeValue);
          if (t) blocks.push({type:'text', text:t});
          return;
        }

        if (node.nodeType !== Node.ELEMENT_NODE) return;
        const tag = node.tagName.toLowerCase();

        if (['script','style','noscript','svg'].includes(tag)) return;

        if (tag === 'img') {
          const src = node.currentSrc || node.src ||
            node.getAttribute('data-src') ||
            node.getAttribute('data-original') ||
            node.getAttribute('data-lazy-src') || '';
          if (src) {
            if (!node.dataset.scrapeImgId) {
                node.dataset.scrapeImgId = 'img_' + (++imgCounter) + '_' + Math.random().toString(36).slice(2);
            }
            blocks.push({
              type:'image',
              src,
              alt: clean(node.getAttribute('alt')),
              w: node.naturalWidth || 0,
              h: node.naturalHeight || 0,
              img_id: node.dataset.scrapeImgId
            });
          }
          return;
        }

        if (tag === 'table') {
          const rows = [];
          for (const tr of node.querySelectorAll('tr')) {
            const cells = [...tr.querySelectorAll(':scope > th, :scope > td')];
            rows.push(cells.map(x => clean(x.innerText)));
          }
          if (rows.length) blocks.push({type:'table', rows});
          return;
        }

        if (tag === 'pre' || tag === 'code') {
          const t = node.innerText.trim();
          if (t) blocks.push({type:'code', text:t});
          return;
        }

        if (tag === 'a') {
          const href = node.href || node.getAttribute('href') || '';
          const t = clean(node.innerText);
          if (t && href) {
            blocks.push({type:'link', text:t, href});
            return;
          }
        }

        // Treat paragraphs, list items and headings as text blocks.
        if (/^(p|li|h2|h3|h4|blockquote|div)$/.test(tag)) {
          const hasBlockChildren = [...node.children].some(ch =>
            /^(p|div|ul|ol|table|pre|blockquote|img|h2|h3|h4)$/.test(ch.tagName.toLowerCase())
          );
          if (!hasBlockChildren) {
            const t = clean(node.innerText);
            if (t) blocks.push({type:'text', text:t});
            return;
          }
        }

        for (const child of node.childNodes) walk(child);
      }

      walk(root);

      // Remove consecutive duplicate text caused by wrapper elements.
      const out = [];
      let last = '';
      for (const b of blocks) {
        if (b.type === 'text') {
          if (b.text === last) continue;
          last = b.text;
        } else {
          last = '';
        }
        out.push(b);
      }
      return out;
    }
    """
    try:
        return post.evaluate(js)
    except Exception:
        return []


def post_header(post, index, is_question):
    try:
        raw_text = post.inner_text(timeout=1500)
        txt = clean(raw_text)
    except Exception:
        raw_text = ""
        txt = ""

    lines = [clean(x) for x in (raw_text.splitlines() if txt else []) if clean(x)]
    author = ""
    date = ""

    for line in lines[:15]:
        if re.search(r"\b(posted|replied|asked|answered)\b", line, re.I):
            date = line
            break

    for line in lines[:10]:
        if len(line) <= 80 and not re.search(
            r"\b(reply|like|share|report|posted|replied|asked|answered)\b",
            line, re.I
        ):
            author = line
            break

    return author, date


def find_image_locator(post, src):
    try:
        matches = post.locator("img")
        for i in range(matches.count()):
            img = matches.nth(i)
            current = img.evaluate(
                """e => e.currentSrc || e.src ||
                e.getAttribute('data-src') ||
                e.getAttribute('data-original') ||
                e.getAttribute('data-lazy-src') || ''"""
            )
            if current and (current in src or src in current or absolute(current, post.page.url) == src):
                return img
    except Exception:
        pass
    return None


def save_rendered_image(page, post, src, out_dir, number, img_id=None):
    ext = ".png"
    path = out_dir / f"image_{number:03d}{ext}"

    img = None
    if img_id:
        try:
            cand = post.locator(f"[data-scrape-img-id='{img_id}']").first
            if cand.count() > 0:
                img = cand
        except Exception:
            pass

    if not img:
        img = find_image_locator(post, src)

    if not img:
        return None

    try:
        img.scroll_into_view_if_needed(timeout=2500)
        page.wait_for_timeout(80)

        # Screenshot is deliberately used rather than only HTTP download:
        # it preserves what the user sees at that position in the thread.
        img.screenshot(path=str(path), type="png")

        if path.exists() and path.stat().st_size > 500:
            return path
    except Exception:
        pass

    return None


def add_link(paragraph, text, url):
    run = paragraph.add_run(text or url)
    try:
        run.font.color.rgb = None
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("r:id"), paragraph.part.relate_to(
            url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            is_external=True
        ))
        r = run._r
        hyperlink.append(r)
        paragraph._p.append(hyperlink)
    except Exception:
        pass


def write_post(doc, page, post, image_dir, post_index, is_question):
    label = "QUESTION" if is_question else f"REPLY {post_index-1}"
    doc.add_heading(label, level=2)
    author, date = post_header(post, post_index, is_question)
    if author or date:
        p = doc.add_paragraph()
        p.add_run("Author: ").bold = True
        p.add_run(author or "Unknown")
        if date:
            p.add_run("    ")
            r = p.add_run(date)
            r.italic = True

    blocks = extract_blocks_from_post(post, page)
    image_count = 0
    text_seen = set()

    for block in blocks:
        typ = block.get("type")
        if typ == "image":
            src = absolute(block.get("src"), page.url)
            if not src:
                continue
            image_count += 1
            path = save_rendered_image(page, post, src, image_dir, image_count, img_id=block.get("img_id"))
            if path:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                try:
                    p.add_run().add_picture(str(path), width=Inches(6.2))
                except Exception:
                    p.add_paragraph(f"[Image captured but could not be embedded: {src}]")
            else:
                p = doc.add_paragraph(f"[IMAGE COULD NOT BE CAPTURED] {src}")
                p.runs[0].italic = True
            alt = clean(block.get("alt"))
            if alt:
                cap = doc.add_paragraph(alt)
                cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if cap.runs:
                    cap.runs[0].italic = True
            continue
        if typ == "table":
            rows = block.get("rows") or []
            if rows:
                cols = max(len(r) for r in rows)
                table = doc.add_table(rows=len(rows), cols=cols)
                table.style = "Table Grid"
                for r, row in enumerate(rows):
                    for c, value in enumerate(row):
                        table.cell(r, c).text = value
                doc.add_paragraph()
            continue
        if typ == "code":
            p = doc.add_paragraph()
            r = p.add_run(block.get("text", ""))
            r.font.name = "Consolas"
            r.font.size = Pt(9)
            continue
        if typ == "link":
            p = doc.add_paragraph()
            add_link(p, block.get("text", ""), block.get("href", ""))
            continue
        if typ == "text":
            txt = clean(block.get("text"))
            if not txt or txt.lower() in {"reply", "like", "share", "report", "copy link", "copy link link copied!", "subscribe", "edit", "delete"}:
                continue
            key = txt.lower()
            if key in text_seen:
                continue
            text_seen.add(key)
            doc.add_paragraph(txt)
    return image_count


def extract_topic_into_page_doc(page, topic, topic_no, page_image_dir, page_doc):
    topic["url"] = canonical_topic_url(topic["url"])
    ready = False
    expected_tid = parse_qs(urlparse(topic["url"]).query).get("threadid", [""])[0].lower()
    for attempt in range(TOPIC_RETRIES):
        try:
            current_tid = parse_qs(urlparse(page.url).query).get("threadid", [""])[0].lower()
            if current_tid != expected_tid or is_login_page(page):
                page.goto(topic["url"], wait_until="commit", timeout=60000)
                wait_for_forum(page)

            if wait_until_topic_ready(page, topic["url"]):
                ready = True
                break

            if is_login_page(page):
                print(f"      Login shell still present; retrying topic ({attempt+1}/{TOPIC_RETRIES})")
                page.wait_for_timeout(2000)
        except Exception as e:
            print(f"      Navigation retry {attempt+1}/{TOPIC_RETRIES}: {e}")
            page.wait_for_timeout(1500)
    if not ready:
        raise RuntimeError("Thread did not become available without a login page; content was NOT scraped to prevent corrupt Word output.")

    click_load_more_replies(page)
    scroll_full(page)

    title = topic_title(page)
    topic_folder = page_image_dir / f"{topic_no:02d}_{safe_name(title)}"
    topic_folder.mkdir(parents=True, exist_ok=True)

    if topic_no > 1:
        page_doc.add_page_break()
    page_doc.add_heading(title, level=1)
    p = page_doc.add_paragraph()
    p.add_run("Topic URL: ").bold = True
    add_link(p, topic["url"], topic["url"])
    if topic.get("replies"):
        p = page_doc.add_paragraph()
        p.add_run("Replies shown on archive: ").bold = True
        p.add_run(str(topic["replies"]))

    posts = get_post_candidates(page)
    if not posts:
        posts = fallback_post_containers(page)

    if not posts:
        page_doc.add_heading("Thread content", level=2)
        try:
            main = page.locator("main").first
            body = clean(main.inner_text(timeout=3000)) if main.count() else clean(page.locator("body").inner_text(timeout=3000))
        except Exception:
            body = ""
        page_doc.add_paragraph(body or "[No thread content could be extracted.]")
        return 0, 0

    total_images = 0
    for i, post in enumerate(posts, 1):
        try:
            total_images += write_post(page_doc, page, post, topic_folder, i, i == 1)
        except Exception as e:
            page_doc.add_heading("EXTRACTION ERROR", level=2)
            page_doc.add_paragraph(str(e))
        page_doc.add_paragraph()
    return len(posts), total_images


def save_page_doc(doc, page_no):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    filename = OUTPUT_DIR / f"Page_{page_no:04d}.docx"
    doc.save(filename)
    return filename


def make_page_document(page_no, archive_url, topics):
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10.5)
    doc.add_heading(f"Dynamics 365 Community - Forum Page {page_no}", level=0)
    doc.add_paragraph(f"Archive page: {archive_url}")
    doc.add_paragraph(f"Cards discovered on this page: {len(topics)}")
    doc.add_paragraph("Complete extracted questions, replies, and inline captured images.")
    return doc


def process_archive_page(archive_page, archive_url, page_no):
    print("\n" + "="*92)
    print(f"ARCHIVE PAGE {page_no}")
    print(archive_url)
    print("="*92)
    try:
        archive_page.goto(archive_url, wait_until="commit", timeout=60000)
        wait_for_forum(archive_page)
        scroll_full(archive_page)
    except Exception as e:
        print(f"[STOP] Could not load the URL for page {page_no}: {e}")
        return False

    topics = collect_page_cards(archive_page, page_no)
    print(f"ARTICLES DISCOVERED ON THIS PAGE: {len(topics)}")
    if not topics:
        print(f"No article links were found on page {page_no}. Check that this archive page exists.")
        return False

    page_image_dir = OUTPUT_DIR / f"Page_{page_no:04d}_Images"
    page_image_dir.mkdir(parents=True, exist_ok=True)
    page_doc = make_page_document(page_no, archive_url, topics)
    completed = posts_total = images_total = 0

    for topic_no, topic in enumerate(topics, 1):
        print(f"\n[PAGE {page_no} | ARTICLE {topic_no}/{len(topics)}] {topic['title']}")
        topic_page = None
        try:
            # Open each article directly in a clean separate tab
            topic_page = archive_page.context.new_page()
            topic_page.goto(topic["url"], wait_until="commit", timeout=60000)

            pc, ic = extract_topic_into_page_doc(topic_page, topic, topic_no, page_image_dir, page_doc)
            completed += 1
            posts_total += pc
            images_total += ic
            fn = save_page_doc(page_doc, page_no)
            print(f"      COMPLETE | posts={pc} | images={ic}")
            print(f"      CHECKPOINT SAVED: {fn}")

        except Exception as e:
            print(f"      FAILED SAFELY: {e}")
            page_doc.add_page_break()
            page_doc.add_heading(f"FAILED TO EXTRACT: {topic['title']}", level=1)
            page_doc.add_paragraph(f"Topic URL: {topic['url']}")
            page_doc.add_paragraph(f"Error: {e}")
            save_page_doc(page_doc, page_no)
        finally:
            if topic_page is not None:
                try:
                    topic_page.close()
                except Exception:
                    pass

    final = save_page_doc(page_doc, page_no)
    print("\n" + "-"*92)
    print(f"PAGE {page_no} FINISHED | articles={completed}/{len(topics)} | posts={posts_total} | images={images_total}")
    print(f"WORD FILE: {final}")
    print("-" * 92)
    return True


def main():
    print("="*92)
    print("DYNAMICS 365 COMMUNITY - PAGE-BY-PAGE WORD EXPORTER V5")
    print("="*92)
    print("Enter the archive/listing page URL and specify the starting and ending page numbers.")
    print()

    while True:
        try:
            archive_url = normalize_archive_url(input("Paste the archive page URL and press Enter: "))
            break
        except ValueError as e:
            print(f"Invalid URL: {e}")

    detected_page = archive_page_number(archive_url)

    while True:
        start_input = input(f"Enter starting page number (default {detected_page}): ").strip()
        if not start_input:
            start_page = detected_page
            break
        try:
            start_page = int(start_input)
            if start_page < 1:
                print("Starting page number must be at least 1.")
                continue
            break
        except ValueError:
            print("Please enter a valid number for the starting page.")

    while True:
        end_input = input(f"Enter ending page number (default {start_page}): ").strip()
        if not end_input:
            end_page = start_page
            break
        try:
            end_page = int(end_input)
            if end_page < start_page:
                print(f"Ending page number must be greater than or equal to starting page ({start_page}).")
                continue
            break
        except ValueError:
            print("Please enter a valid number for the ending page.")

    total_pages = end_page - start_page + 1
    if total_pages == 1:
        print(f"\nSelected archive page number: {start_page}")
        print(f"Selected URL: {build_page_url(archive_url, start_page)}\n")
    else:
        print(f"\nSelected archive page range: {start_page} to {end_page} ({total_pages} pages)")
        print(f"Base URL: {archive_url}\n")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            str(PROFILE_DIR), headless=False, viewport={"width": 1440, "height": 900},
            args=["--disable-blink-features=AutomationControlled"])

        # Accelerate network loading by filtering out heavy third-party tracking & advertising beacons
        context.route(
            "**/*",
            lambda route: route.abort() if any(
                d in route.request.url.lower() for d in BLOCKED_DOMAINS
            ) else route.continue_()
        )

        archive_page = context.pages[0] if context.pages else context.new_page()

        try:
            for current_page_no in range(start_page, end_page + 1):
                current_url = build_page_url(archive_url, current_page_no)
                process_archive_page(archive_page, current_url, current_page_no)

                # Ensure any extra tabs opened during topic processing are closed
                for extra_page in context.pages[1:]:
                    try:
                        extra_page.close()
                    except Exception:
                        pass

                if current_page_no < end_page:
                    time.sleep(1)
        finally:
            try:
                context.close()
            except Exception:
                pass

    print("\n" + "="*92)
    print("PROCESSING FINISHED FOR ALL REQUESTED PAGES")
    print(f"OUTPUT FOLDER: {OUTPUT_DIR}")
    print("="*92)


if __name__ == "__main__":
    main()
