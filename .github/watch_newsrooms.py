#!/usr/bin/env python3
"""
Company Newsroom Watcher.
Tracks press releases published on corporate domains.
Enforces:
  1. Universal Date Sanitization (YYYY-MM-DD).
  2. Deep Article Publish-Date Extraction (ignoring sitemap CMS mass re-edits).
  3. Career/Job/Furniture Filtering.
  4. Full-Text Body Paragraph Extraction.
  5. Strict Time-Windowing with Latest-Release Fallback Status.
"""

import csv
import gzip
import html as html_mod
import json
import os
import re
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import email.utils
from datetime import datetime, timezone
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed

WORKSPACE = os.environ.get("WORKSPACE") or os.path.normpath(os.path.join(os.path.dirname(__file__), "..", ".."))

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

try:
    from robust_fetcher import extract_slug_date, extract_dateline_date, is_article_stale_redesign, decode_authoritative_link
except ImportError:
    # Local fallback
    def extract_slug_date(url): return ""
    def extract_dateline_date(text): return ""
    def is_article_stale_redesign(pub_date, url, text, cutoff): return False, ""
    def decode_authoritative_link(url): return url

TIMEOUT   = float(os.environ.get("NEWS_TIMEOUT", "25"))
WORKERS   = int(os.environ.get("NEWS_WORKERS", "6"))
MAX_SUBS  = int(os.environ.get("NEWS_MAX_SUBSITEMAPS", "20"))
STATE_CAP = int(os.environ.get("NEWS_STATE_CAP", "6000"))
ROW_CAP   = int(os.environ.get("NEWS_ROW_CAP", "1200"))
UA        = os.environ.get("NEWS_UA",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")
SEC_UA    = os.environ.get("SEC_UA", "")

MONTHS_MAP = {
    'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6,
    'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12,
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
}

def parse_any_date(date_str):
    """Sanitizes any date format into standard YYYY-MM-DD."""
    if not date_str:
        return ""
    date_str = str(date_str).strip()
    
    # RFC 2822 e.g. 'Tue, 04 Aug 2026 20:01:47 +0000'
    try:
        parsed_tuple = email.utils.parsedate_tz(date_str)
        if parsed_tuple:
            dt = datetime.fromtimestamp(email.utils.mktime_tz(parsed_tuple), tz=timezone.utc)
            return dt.strftime("%Y-%m-%d")
    except Exception:
        pass

    # ISO 8601 e.g. '2026-08-25T14:30:00Z' or '2026-08-25'
    iso_match = re.search(r'(\d{4})-(\d{2})-(\d{2})', date_str)
    if iso_match:
        return f"{iso_match.group(1)}-{iso_match.group(2)}-{iso_match.group(3)}"
        
    # 'Month DD, YYYY'
    m = re.search(r'([a-zA-Z]{3,9})\s+(\d{1,2}),?\s+(\d{4})', date_str)
    if m:
        mon_str = m.group(1).lower()
        if mon_str in MONTHS_MAP:
            return f"{int(m.group(3)):04d}-{MONTHS_MAP[mon_str]:02d}-{int(m.group(2)):02d}"
            
    # 'DD Month YYYY'
    m2 = re.search(r'(\d{1,2})\s+([a-zA-Z]{3,9})\s+(\d{4})', date_str)
    if m2:
        mon_str = m2.group(2).lower()
        if mon_str in MONTHS_MAP:
            return f"{int(m2.group(3)):04d}-{MONTHS_MAP[mon_str]:02d}-{int(m2.group(1)):02d}"
            
    return ""

NEWSY = re.compile(r"/(news|press|media|announcement|newsroom|press-release|media-release|"
                   r"press-releases|media-releases|stories|articles)", re.I)
JUNK  = re.compile(r"/(careers?|jobs?|search|login|privacy|cookie|sitemap|tag|category|author|"
                   r"page/\d+|feed|vacanc\w*|internships?|working-at-|talent|employment)(/|$)", re.I)
FEED_PATHS = ["/feed/", "/rss.xml", "/rss", "/news-rss.xml", "/feed.xml", "/atom.xml",
              "/news/rss.xml", "/news/feed/", "/media/feed/", "/rss/news-releases.xml"]

LOCALE_SEG = re.compile(r"(?<=/)([a-z]{2})[-_]([a-z]{2})(?=/)", re.I)

IR_JUNK = re.compile(
    r"/(analysts?|analyst-coverage|board|board-?members?|management|leadership|"
    r"governance|committees?|officers?|directors?|executives?|team|"
    r"stock-?(?:quote|chart|info|price)|historical-?data|contacts?|"
    r"media-?contacts?|media-?kits?|media-?resources?|email-?alerts?|"
    r"fact-?sheets?|faqs?|sec-?filings?|filings?|financials?|"
    r"annual-?reports?|events?(?:-and-presentations?)?|presentations?|"
    r"assets?|resources?|overview|profile|subscribe)(?:/|$|\.)", re.I)

NOISE = re.compile(r"/(?:[a-z]+-)?stories(?:/|$)|"
                   r"/(?:perspectives?|employee-spotlights?|spotlights?|"
                   r"life-at-[\w-]+|blogs?|podcasts?|insights?|our-people|"
                   r"culture|community)(?:/|$)", re.I)

SLUG_JUNK = re.compile(r"(fact-?sheet|media-?kit|press-?kit|corporate-?(?:brochure|deck)|"
                       r"brand-?guide|image-?library|photo-?library|logo|modern-?slavery|"
                       r"slavery-?act|subscribe-?to-?news|email-?alerts?)", re.I)

TITLE_JUNK = re.compile(
    r"(now hiring|we(?:'re| are) hiring|\bis hiring\b|\bhiring\b|"
    r"job (?:opening|posting|description|alert|vacanc)\w*|open positions?|"
    r"\bjobs? (?:in|at|near)\b|careers? (?:at|page|site|opportunit\w*)|"
    r"apply now|vacanc(?:y|ies)|internships?|recruit(?:ing|ment|er)\b|"
    r"modern slavery|slavery act statement|subscribe to news)", re.I)

PUBLISHER_JUNK = re.compile(r"(careers?|jobs?|recruit\w*|talent)\b", re.I)

SECTION_SLUG = re.compile(
    r"^(news|press|media|newsroom|press-?releases?|news-?releases?|"
    r"news-and-media|media-and-news|news-?(?:and-)?events?|media-cent(?:re|er)|"
    r"announcements?|stories|updates?|insights?|publications?|"
    r"press-?room|media-?room|latest-?news)$", re.I)

GENERIC_TITLE = re.compile(
    r"^(news|newsroom|news (?:&|and) (?:media|events)|press|press releases?|"
    r"media|media (?:centre|center|room|relations)|announcements?|"
    r"events?(?: (?:&|and) presentations?)?|presentations?|publications?|"
    r"investors?|investor relations|home|overview|contact(?: us)?|about(?: us)?)$",
    re.I)

def generic_title(title):
    parts = re.split(r"\s+[-–—|]\s+", title.strip())
    head = parts[0].strip(" :\u00b7")
    if len(parts) > 1 and PUBLISHER_JUNK.search(parts[-1]):
        return True
    return bool(head) and bool(GENERIC_TITLE.match(head))

def foreign_locale(url):
    for a, b in LOCALE_SEG.findall(url):
        if "en" not in (a.lower(), b.lower()):
            return True
    return False

def keep(url, title=""):
    if JUNK.search(url) or IR_JUNK.search(url) or NOISE.search(url):
        return False
    if foreign_locale(url):
        return False
    slug = urllib.parse.urlparse(url).path.rstrip("/").rsplit("/", 1)[-1]
    slug = re.sub(r"\.(html?|php|aspx|pdf)$", "", slug)
    if SLUG_JUNK.search(slug) or SECTION_SLUG.match(slug):
        return False
    if title and (TITLE_JUNK.search(title) or generic_title(title)):
        return False
    return True

def looks_like_release(url):
    slug = urllib.parse.urlparse(url).path.rstrip("/").split("/")[-1]
    slug = re.sub(r"\.(html?|php|aspx|pdf)$", "", slug)
    return slug.count("-") >= 2 or bool(re.search(r"\d", slug))

IR_PREFIXES = ["ir.", "investors.", "investor.", "news.", "newsroom.", "media."]

def domain_variants(domain):
    d = re.sub(r"^www\.", "", domain)
    out = [domain]
    for pre in IR_PREFIXES:
        if not d.startswith(pre):
            out.append(pre + d)
    brand = d.split(".")[0]
    if brand and len(brand) > 2:
        out.append(f"{brand}.gcs-web.com")
    seen = set()
    return [x for x in out if not (x in seen or seen.add(x))]

ALT_UA = "curl/8.5.0"

def fetch(url, timeout=None, ua=None):
    ua = ua or (SEC_UA if (SEC_UA and "sec.gov" in url) else UA)
    req = urllib.request.Request(url, headers={
        "User-Agent": ua,
        "Accept": "application/xml,text/xml,application/rss+xml,text/html;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
    })
    try:
        with urllib.request.urlopen(req, timeout=timeout or TIMEOUT) as r:
            b = r.read(8 * 1024 * 1024)
            if b[:2] == b"\x1f\x8b":
                b = gzip.decompress(b)
            return r.getcode(), b, r.geturl()
    except urllib.error.HTTPError as e:
        if e.code in (403, 406, 429) and ua != ALT_UA:
            return fetch(url, timeout, ALT_UA)
        return e.code, b"", url
    except Exception:
        return 0, b"", url

def _text(el):
    return " ".join((el.text or "").split()) if el is not None else ""

def parse_feed(body, base):
    try:
        root = ET.fromstring(body.decode("utf-8", "replace").lstrip("﻿ \r\n\t"))
    except ET.ParseError:
        return []
    def local(t): return t.rsplit("}", 1)[-1].lower()
    items = []
    for el in root.iter():
        if local(el.tag) in ("item", "entry"):
            title = link = raw_date = desc = ""
            for c in el:
                n = local(c.tag)
                if n == "title" and not title:
                    title = _text(c)
                elif n == "link" and not link:
                    link = _text(c) or c.get("href", "")
                elif n in ("pubdate", "published", "updated", "date") and not raw_date:
                    raw_date = _text(c)[:40]
                elif n in ("description", "summary", "content") and not desc:
                    desc = re.sub(r"<[^>]+>", " ", _text(c))
                    desc = " ".join(desc.split())[:300]
            if link:
                norm_date = parse_any_date(raw_date)
                items.append({
                    "title": title,
                    "url": urllib.parse.urljoin(base, link),
                    "date": norm_date,
                    "excerpt": desc,
                    "full_text": desc,
                    "date_source": "feed"
                })
    return items

def try_feeds(domain):
    for host in domain_variants(domain):
        base = f"https://{host}"
        code, body, _ = fetch(base, timeout=10)
        if code == 0 and host != domain:
            continue
        cands = []
        if body:
            html = body.decode("utf-8", "replace")
            for m in re.finditer(r'<link[^>]+type=["\']application/(?:rss|atom)\+xml["\'][^>]*>', html, re.I):
                href = re.search(r'href=["\']([^"\']+)', m.group(0))
                if href:
                    cands.append(urllib.parse.urljoin(base, href.group(1)))
        cands += [base + p for p in FEED_PATHS]
        seen = set()
        for u in cands:
            if u in seen:
                continue
            seen.add(u)
            c, b, _ = fetch(u)
            if c == 200 and b:
                items = [i for i in parse_feed(b, u) if keep(i["url"], i["title"])]
                if len(items) >= 3:
                    return u, items
    return None, []

def sitemap_urls(domain):
    roots = []
    c, b, _ = fetch(f"https://{domain}/robots.txt")
    if b:
        for hit in re.findall(r"(?i)^\s*sitemap:\s*(\S+)", b.decode("utf-8", "replace"), re.M):
            roots.append(urllib.parse.urljoin(f"https://{domain}/", hit.strip()))
    roots += [f"https://{domain}/sitemap_index.xml", f"https://{domain}/sitemap.xml",
              f"https://{domain}/sitemap-index.xml", f"https://{domain}/news-sitemap.xml"]
    out, seen = [], set()
    for r in roots:
        if r not in seen:
            seen.add(r)
            out.append(r)
    return out

def read_sitemap(url, depth=0):
    c, b, _ = fetch(url)
    if c != 200 or not b:
        return []
    text = b.decode("utf-8", "replace")
    if "<sitemapindex" in text[:2000].lower() and depth < 2:
        subs = re.findall(r"<loc>\s*([^<\s]+)\s*</loc>", text)
        subs.sort(key=lambda u: (0 if NEWSY.search(u) else 1, u))
        rows = []
        for s in subs[:MAX_SUBS]:
            rows += read_sitemap(s, depth + 1)
        return rows
    rows = []
    for block in re.findall(r"<url>(.*?)</url>", text, re.S):
        loc = re.search(r"<loc>\s*([^<\s]+)", block)
        lm  = re.search(r"<lastmod>\s*([^<\s]+)", block)
        if loc:
            clean_date = parse_any_date(lm.group(1)) if lm else ""
            rows.append((loc.group(1), clean_date))
    return rows

def try_sitemaps(domain, path_filter=""):
    pat = re.compile(re.escape(path_filter), re.I) if path_filter else NEWSY
    found, used, deadline = {}, [], time.time() + float(os.environ.get("NEWS_SITEMAP_BUDGET", "70"))
    roots = []
    for host in domain_variants(domain):
        roots += sitemap_urls(host)
    for sm in roots:
        if time.time() > deadline:
            break
        rows = read_sitemap(sm)
        hits = [(u, d) for u, d in rows
                if pat.search(u) and keep(u) and looks_like_release(u)]
        if not hits and rows:
            host = re.sub(r"^https?://", "", sm).split("/")[0]
            if any(host.startswith(p) for p in IR_PREFIXES) or "gcs-web.com" in host:
                hits = [(u, d) for u, d in rows if keep(u) and looks_like_release(u)]
        if hits:
            used.append(sm)
            for u, d in hits:
                if u not in found or d > found[u]:
                    found[u] = d
        if len(found) > 1500:
            break
    best = list(found.items())
    used = used[0] if used else ""
    items = [{"title": slug_title(u), "url": u, "date": d, "date_source": "sitemap_hint", "excerpt": "", "full_text": ""} for u, d in best]
    items.sort(key=lambda x: x["date"], reverse=True)
    return used, items

def slug_title(url):
    slug = urllib.parse.urlparse(url).path.rstrip("/").split("/")[-1]
    slug = re.sub(r"\.(html?|php|aspx)$", "", slug)
    slug = re.sub(r"[-_]+", " ", slug).strip()
    return slug[:1].upper() + slug[1:] if slug else url

def try_html(domain, newsroom_path):
    if not newsroom_path:
        return "", []
    url = urllib.parse.urljoin(f"https://{domain}", newsroom_path)
    c, b, final = fetch(url)
    if c != 200 or not b:
        return "", []
    html = b.decode("utf-8", "replace")
    items, seen = [], set()
    for m in re.finditer(r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', html, re.S | re.I):
        href, label = m.group(1), re.sub(r"<[^>]+>", " ", m.group(2))
        label = " ".join(label.split())
        full = urllib.parse.urljoin(final, href)
        if domain not in full or not NEWSY.search(full) or not keep(full, label):
            continue
        if len(label) < 25 or full in seen:
            continue
        seen.add(full)
        items.append({"title": label[:300], "url": full, "date": "", "date_source": "html", "excerpt": "", "full_text": ""})
    return url, items

def try_gnews(domain, days="90d"):
    q = urllib.parse.quote(f"site:{domain} when:{days}")
    url = f"https://news.google.com/rss/search?q={q}&hl=en-US&gl=US&ceid=US:en"
    c, b, _ = fetch(url)
    if c != 200 or not b:
        return "", []
    items = [i for i in parse_feed(b, url) if keep(i["url"], i["title"])]
    return url, items

def extract_full_text(html):
    """Extract clean readable paragraphs from article HTML, stripping boilerplate."""
    html_clean = re.sub(r'<(script|style|header|footer|nav|noscript)[^>]*>.*?</\1>', ' ', html, flags=re.S|re.I)
    paragraphs = re.findall(r'<p[^>]*>(.*?)</p>', html_clean, flags=re.S|re.I)
    cleaned_paras = []
    for p in paragraphs:
        text = re.sub(r'<[^>]+>', ' ', p)
        text = ' '.join(text.split())
        if len(text) > 45 and not re.search(r'(cookie|privacy policy|terms of use|all rights reserved|subscribe to newsletter|javascript is disabled)', text, re.I):
            cleaned_paras.append(text)
    return "\n\n".join(cleaned_paras)[:3500]

def enrich(item):
    """Deep scrape article page to extract true published date, real title, excerpt, and full text."""
    c, b, _ = fetch(item["url"], timeout=12)
    if c != 200 or not b:
        return
    html = b.decode("utf-8", "replace")
    
    # 1. Real Published Date extraction from page HTML
    real_date = ""
    # Check JSON-LD schema
    json_ld = re.findall(r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>', html, re.S | re.I)
    for j in json_ld:
        try:
            d_match = re.search(r'["\']datePublished["\']\s*:\s*["\']([^"\']+)["\']', j)
            if d_match:
                real_date = parse_any_date(d_match.group(1))
                if real_date:
                    break
        except Exception:
            pass
            
    if not real_date:
        # Check meta article:published_time (NOT article:modified_time)
        meta_d = re.search(r'<meta[^>]+property=["\']article:published_time["\'][^>]+content=["\']([^"\']+)', html, re.I)
        if not meta_d:
            meta_d = re.search(r'<meta[^>]+name=["\'](?:date|pubdate|publish-date|DC\.date\.issued)["\'][^>]+content=["\']([^"\']+)', html, re.I)
        if meta_d:
            real_date = parse_any_date(meta_d.group(1))

    if not real_date:
        # Check <time> tag
        time_tag = re.search(r'<time[^>]+datetime=["\']([^"\']+)', html, re.I)
        if time_tag:
            real_date = parse_any_date(time_tag.group(1))
            
    if not real_date:
        # Check visible lead dateline (e.g. "INDIANAPOLIS, Jan. 14, 2021 /PRNewswire/ --")
        dateline_d = extract_dateline_date(html[:4000])
        if dateline_d:
            real_date = dateline_d

    # URL Slug date verification (e.g. /2021/04/15/ or /2022-press-release)
    # The URL slug date is immutable and overrides sitemap lastmod timestamps from CMS redesigns
    slug_d = extract_slug_date(item["url"])
    if slug_d:
        if not real_date or (real_date > slug_d and slug_d[:4] < "2026"):
            real_date = slug_d

    if real_date:
        item["date"] = real_date
        item["date_source"] = "article"
    elif item.get("date_source") == "sitemap_hint":
        # If we could not find an article publish date, clear the sitemap lastmod hint so it's not assumed as today
        item["date_source"] = "sitemap_unverified"

    # 2. Title extraction
    title_m = re.search(r'<meta[^>]+property=["\']og:title["\'][^>]+content=["\']([^"\']+)', html, re.I)
    if not title_m:
        title_m = re.search(r'<title[^>]*>(.*?)</title>', html, re.I)
    if title_m:
        raw_t = re.sub(r'<[^>]+>', '', title_m.group(1))
        item["title"] = html_mod.unescape(" ".join(raw_t.split()))[:300]

    # 3. Excerpt extraction
    desc_m = re.search(r'<meta[^>]+property=["\']og:description["\'][^>]+content=["\']([^"\']+)', html, re.I)
    if not desc_m:
        desc_m = re.search(r'<meta[^>]+name=["\']description["\'][^>]+content=["\']([^"\']+)', html, re.I)
    if desc_m:
        item["excerpt"] = html_mod.unescape(" ".join(desc_m.group(1).split()))[:400]

    # 4. Full Text Content extraction
    full_text = extract_full_text(html)
    full_text_clean = html_mod.unescape(full_text) if full_text else ""
    item["full_text"] = full_text_clean or item.get("excerpt", "")
    if not item["excerpt"] and full_text_clean:
        item["excerpt"] = full_text_clean[:250] + "..."

def check_company(row):
    comp, dom = row[0].strip(), row[1].strip()
    path = row[2].strip() if len(row) > 2 else ""
    
    # 1. Native RSS Feed
    u, items = try_feeds(dom)
    if items:
        return {"company": comp, "domain": dom, "method": "feed", "source": u, "items": items}
        
    # 2. XML Sitemap
    u, items = try_sitemaps(dom, path)
    if items:
        return {"company": comp, "domain": dom, "method": "sitemap", "source": u, "items": items}
        
    # 3. HTML Scraping
    if path:
        u, items = try_html(dom, path)
        if items:
            return {"company": comp, "domain": dom, "method": "listing", "source": u, "items": items}
            
    # 4. Google News fallback
    u, items = try_gnews(dom)
    if items:
        return {"company": comp, "domain": dom, "method": "gnews", "source": u, "items": items}
        
    return {"company": comp, "domain": dom, "method": "none", "source": "", "items": []}

def main():
    import argparse
    parser = argparse.ArgumentParser(description="Company Newsroom Watcher")
    parser.add_argument("input_file", nargs="?", default="companies_sample.tsv")
    parser.add_argument("--days", type=int, default=int(os.environ.get("NEWS_DAYS_BACK", "3")), help="Days back to monitor (default: 3)")
    parser.add_argument("--hours", type=int, default=0, help="Hours back to monitor")
    parser.add_argument("--seed", action="store_true", help="Record releases without reporting new")
    args = parser.parse_args()

    days_back = args.days
    if args.hours > 0:
        cutoff_seconds = args.hours * 3600
        time_desc = f"{args.hours} hours"
    else:
        cutoff_seconds = days_back * 86400
        time_desc = f"{days_back} days"

    cutoff_date = time.strftime("%Y-%m-%d", time.gmtime(time.time() - cutoff_seconds))

    input_path = args.input_file if os.path.isabs(args.input_file) else os.path.join(WORKSPACE, args.input_file)
    if not os.path.exists(input_path):
        print(f"Error: companies file '{input_path}' not found.")
        sys.exit(1)

    with open(input_path, encoding="utf-8") as f:
        reader = csv.reader(f, delimiter="\t")
        companies = [r for r in reader if r and not r[0].startswith("#") and r[0] != "Company"]

    state_path = os.path.join(WORKSPACE, "state", "seen.json")
    os.makedirs(os.path.dirname(state_path), exist_ok=True)
    seen = json.load(open(state_path, encoding="utf-8")) if os.path.exists(state_path) else {}

    print("\n" + "="*78)
    print(f"🏢 Company Newsroom Watcher — Monitoring {len(companies)} companies")
    print(f"⏱️  Time Window: Last {time_desc} (since {cutoff_date}) · Workers: {WORKERS}")
    print("="*78 + "\n", flush=True)

    results = []
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futs = {pool.submit(check_company, c): c for c in companies}
        for f in as_completed(futs):
            try:
                r = f.result()
            except Exception as e:
                c = futs[f]
                r = {"company": c[0], "domain": c[1], "method": "error", "source": str(e)[:80], "items": []}
            results.append(r)
            print(f"  {r['company'][:28]:<28} {r['method']:<8} {len(r['items']):>4} items found   {r['source'][:65]}", flush=True)

    out_dir = os.environ.get("NEWS_OUT_DIR", os.path.join(WORKSPACE, "results"))
    os.makedirs(out_dir, exist_ok=True)

    HEADER = [
        "Company", "Domain", "Method", "Date (YYYY-MM-DD)", "Date Source",
        "Title / Headline", "Snippet / Excerpt", "Full Text Content",
        "Authentic URL", "Status"
    ]

    SEED = args.seed or os.environ.get("NEWS_SEED", "").lower() in ("1", "true", "yes")

    # Collect top candidate items per company to deep-enrich their real published dates & full text
    cands_to_enrich = []
    for r in results:
        # Pre-normalize any feed/sitemap dates
        for i in r["items"]:
            i["date"] = parse_any_date(i.get("date", ""))
            i.setdefault("excerpt", "")
            i.setdefault("full_text", "")
            i.setdefault("date_source", "feed" if r["method"] == "feed" else "sitemap_hint")
            
        # Top 15 newest items per company to enrich for real publish date
        cands_to_enrich += [i for i in r["items"] if i.get("date_source") != "feed"][:15]

    if not SEED and cands_to_enrich:
        to_enrich = cands_to_enrich[:300]
        print(f"\n[ENRICHMENT] Deep-scraping {len(to_enrich)} candidate articles for real publish dates & full text...", flush=True)
        with ThreadPoolExecutor(max_workers=WORKERS) as pool:
            list(pool.map(enrich, to_enrich))

    # Record discovered working feeds
    disc_feeds = [[r["company"], r["domain"], r["source"]] for r in results if r.get("method") == "feed" and r.get("source")]
    if disc_feeds:
        with open(os.path.join(out_dir, "discovered_feeds.tsv"), "w", encoding="utf-8") as f:
            f.write("Company\tDomain\tDiscovered_Feed_URL\n")
            for c, d, u in disc_feeds:
                f.write(f"{c}\t{d}\t{u}\n")

    rows = []
    total_in_window = 0

    for r in sorted(results, key=lambda x: x["company"].lower()):
        known = set(seen.get(r["domain"], []))
        
        # Strict Date Filter: Must be valid ISO date and >= cutoff_date
        # Only trust "article" or "feed" dates; do NOT trust unverified sitemap lastmod hints from mass CMS rebuilds
        in_window = [
            i for i in r["items"]
            if i.get("date") and i["date"] >= cutoff_date and i.get("date_source") in ("article", "feed", "gnews") and keep(i["url"], i["title"])
        ]
        
        if in_window:
            for i in in_window[:ROW_CAP]:
                total_in_window += 1
                status = "NEW (In Window)" if i["url"] not in known else "seen (in-window)"
                if SEED:
                    status = "seeded"
                rows.append([
                    r["company"], r["domain"], r["method"], i["date"],
                    i.get("date_source", ""), i["title"], i.get("excerpt", ""),
                    i.get("full_text", ""), i["url"], status
                ])
        else:
            # If no releases in window, find the single latest valid item for that company
            valid_items = [i for i in r["items"] if keep(i["url"], i["title"])]
            if valid_items:
                latest = valid_items[0]
                l_date = latest.get("date") or "Previous"
                l_title = latest.get("title", "Recent Release")
                rows.append([
                    r["company"], r["domain"], r["method"], l_date,
                    latest.get("date_source", ""),
                    f"No releases in last {time_desc} (Latest: '{l_title}')",
                    latest.get("excerpt", ""), latest.get("full_text", ""),
                    latest.get("url", ""), "NO_NEW_IN_WINDOW"
                ])
            else:
                rows.append([
                    r["company"], r["domain"], r["method"], "", "",
                    f"No press releases found on domain in last {time_desc}",
                    "", "", "", "NO_RELEASES_FOUND"
                ])

        # Update persistent state
        merged = list(dict.fromkeys([i["url"] for i in r["items"]] + sorted(known)))
        seen[r["domain"]] = merged[:STATE_CAP]

    with open(os.path.join(out_dir, "newsroom.tsv"), "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter="\t")
        writer.writerow(HEADER)
        writer.writerows(rows)

    json.dump(seen, open(state_path, "w", encoding="utf-8"), indent=2)

    new_in_win = sum(1 for r in rows if "NEW" in r[9])
    covered = sum(1 for r in results if r["items"])
    print("\n" + "=" * 78)
    print(f"✅ Newsroom Watch Complete — {covered}/{len(results)} companies covered")
    print(f"📊 Active releases in last {time_desc}: {total_in_window} ({new_in_win} new since last run)")
    print(f"📁 Results saved to: {os.path.join(out_dir, 'newsroom.tsv')}")
    print("=" * 78 + "\n")

    for r in rows[:30]:
        print(f"[{r[0]:<22}] {r[3]:<10} {r[9]:<18} {r[5][:50]}")

if __name__ == "__main__":
    main()
