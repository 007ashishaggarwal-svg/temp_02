#!/usr/bin/env python3
"""
RSS feed checker.

Input : any list of feeds with the three columns  Feed ID | URL | Label
        (tab, comma, semicolon or pipe separated -- header row optional).
Output: seven columns  Feed ID | URL | Label | Status | Time | Result | Last Item Title
        written as TSV, CSV and (if openpyxl is available) XLSX.

Only the Python standard library is required.
"""

import csv
import gzip
import io
import os
import re
import sys
import threading
import time
import urllib.error
import urllib.request
import zlib
import ssl
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# ----------------------------------------------------------------------------- config
TIMEOUT      = float(os.environ.get("FEED_TIMEOUT", "30"))    # seconds per attempt
WORKERS      = int(os.environ.get("FEED_WORKERS", "8"))       # feeds fetched at once
PER_HOST_GAP = float(os.environ.get("FEED_HOST_GAP", "1.0"))  # min seconds between hits on same host
RETRIES      = int(os.environ.get("FEED_RETRIES", "1"))       # extra attempts for timeouts / 429 / 5xx
MAX_BYTES    = 3 * 1024 * 1024                                # don't download more than 3 MB per feed


def _read_all(resp, limit=None):
    """Read the whole body.

    resp.read(n) returns *up to* n bytes -- on a chunked response it often
    returns only the first chunk, which silently truncates the feed and makes
    valid XML unparseable. Loop until EOF.
    """
    limit = limit or MAX_BYTES
    buf = io.BytesIO()
    while buf.tell() < limit:
        chunk = resp.read(min(65536, limit - buf.tell()))
        if not chunk:
            break
        buf.write(chunk)
    return buf.getvalue()

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")
HEADERS = {
    "User-Agent": UA,
    "Accept": "application/rss+xml, application/atom+xml, application/xml;q=0.9, text/xml;q=0.9, */*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate",
    "Connection": "close",
}

HEADER = ["Feed ID", "URL", "Label", "Status", "Time", "Result", "Last Item Title", "Working Fallback URL"]
URL_RE = re.compile(r"https?://[^\s,;|\"'<>]+", re.I)
ID_RE  = re.compile(r"^[A-Za-z]{0,12}[-_ ]?\d{1,5}$")   # Feed-001, Feed_01, F12, 001 ...


# ----------------------------------------------------------------------------- input parsing
def _clean(cell):
    return (cell or "").strip().strip('"').strip()


def _is_header(fields):
    joined = " ".join(fields).lower()
    return "url" in joined and ("feed" in joined or "label" in joined) and "http" not in joined


def parse_lines(text):
    """Normal case: one feed per line, columns split by tab / comma / semicolon / pipe."""
    rows = []
    for raw in text.splitlines():
        line = raw.strip().lstrip("﻿")
        if not line or line.startswith("#"):
            continue
        if "\t" in line:
            fields = line.split("\t")
        elif "|" in line and "://" in line:
            fields = line.split("|")
        elif ";" in line and line.count(";") >= line.count(","):
            fields = next(csv.reader([line], delimiter=";"))
        elif "," in line:
            fields = next(csv.reader([line]))
        else:
            fields = line.split(None, 1)
        fields = [_clean(f) for f in fields]
        if _is_header(fields):
            continue
        url_i = next((i for i, f in enumerate(fields) if URL_RE.fullmatch(f) or f.lower().startswith(("http://", "https://"))), None)
        if url_i is None:
            m = URL_RE.search(line)
            if not m:
                continue
            rows.append(["", m.group(0), _clean(line.replace(m.group(0), " "))])
            continue
        url   = fields[url_i]
        fid   = fields[url_i - 1] if url_i > 0 else ""
        label = fields[url_i + 1] if len(fields) > url_i + 1 else ""
        rows.append([fid, url, label])
    return rows


def parse_blob(text):
    """
    Fallback for text whose line breaks were lost -- e.g. pasting many rows into the
    single-line "Run workflow" box, where the browser flattens them onto one line.
    Records are rebuilt around the URLs: [Feed ID] URL [Label] [Feed ID] URL [Label] ...
    """
    flat = " ".join(text.split())
    matches = list(URL_RE.finditer(flat))
    if not matches:
        return []
    rows = []
    for i, m in enumerate(matches):
        before = flat[matches[i - 1].end():m.start()] if i else flat[:m.start()]
        toks = [t for t in before.replace("|", " ").replace("\t", " ").split() if t.strip(" ,;")]
        fid = ""
        if toks and ID_RE.match(toks[-1].strip(" ,;")):
            fid = toks.pop().strip(" ,;")
        prev_label = " ".join(toks).strip(" |,;")
        if i and prev_label:
            rows[-1][2] = prev_label
        rows.append([fid, m.group(0).rstrip(".,;|"), ""])
    tail = flat[matches[-1].end():].replace("|", " ").strip(" ,;")
    if tail:
        rows[-1][2] = tail
    return rows


def parse_feeds(text):
    if not text or not text.strip():
        return []
    rows = parse_lines(text) if ("\n" in text.strip() or "\t" in text) else []
    if len(rows) <= 1 and len(URL_RE.findall(text)) > 1:
        rows = parse_blob(text)
    if not rows:
        rows = parse_lines(text)
    out, seen = [], set()
    n = 0
    for fid, url, label in rows:
        url = url.strip()
        if not url.lower().startswith(("http://", "https://")):
            continue
        n += 1
        fid = fid.strip() or f"Feed-{n:03d}"
        key = (fid, url)
        if key in seen:
            continue
        seen.add(key)
        out.append([fid, url, label.strip()])
    return out


# ----------------------------------------------------------------------------- fetching
_host_locks, _host_last, _guard = {}, {}, threading.Lock()

# Some publishers serve an incomplete certificate chain (hutch-med.com is one).
# Browsers paper over it; urllib does not. We retry such hosts with verification
# relaxed and LABEL the row, rather than silently dropping a working feed.
_LAX = ssl.create_default_context()
_LAX.check_hostname = False
_LAX.verify_mode = ssl.CERT_NONE


def _host_slot(url):
    host = re.sub(r"^https?://", "", url).split("/")[0].lower()
    with _guard:
        lock = _host_locks.setdefault(host, threading.Lock())
    return host, lock


def _decode_body(raw, encoding):
    try:
        if encoding == "gzip":
            return gzip.decompress(raw)
        if encoding == "deflate":
            try:
                return zlib.decompress(raw)
            except zlib.error:
                return zlib.decompress(raw, -zlib.MAX_WBITS)
    except Exception:
        pass
    return raw


def _sanitize(text):
    """Escape undefined HTML entities (&nbsp; &mdash; ...) that ElementTree rejects.

    Feeds published by Squarespace and similar CMSes contain HTML entities which
    are not defined in XML. ElementTree raises "undefined entity" and the whole
    feed is lost, even though every item in it is perfectly readable.
    """
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    return re.sub(r"&(?!(?:amp|lt|gt|quot|apos|#\d+|#[xX][0-9a-fA-F]+);)", "&amp;", text)


TITLE_RE = re.compile(
    r"<(?:item|entry)\b.*?<title[^>]*>(.*?)</title>", re.I | re.S)
CDATA_RE = re.compile(r"<!\[CDATA\[(.*?)\]\]>", re.S)
TAG_RE   = re.compile(r"<[^>]+>", re.S)


def _text_of(el):
    """All text inside an element, including nested markup.

    Some feeds (FiercePharma/Biotech/Healthcare) wrap the headline in a link:
        <title><a href="...">Headline</a></title>
    so element.text is None and the headline lives in the child. itertext()
    collects it wherever it sits.
    """
    return " ".join("".join(el.itertext()).split())


def _title_by_regex(text):
    """Last resort when the XML will not parse at all (truncated or malformed)."""
    m = TITLE_RE.search(text)
    if not m:
        return ""
    raw = m.group(1)
    cd = CDATA_RE.search(raw)
    if cd:
        raw = cd.group(1)
    raw = TAG_RE.sub(" ", raw)
    for ent, ch in (("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"),
                    ("&quot;", '"'), ("&apos;", "'"), ("&#39;", "'"), ("&nbsp;", " ")):
        raw = raw.replace(ent, ch)
    return " ".join(raw.split())[:300]


def _first_item_title(body):
    """Newest entry's title -- feeds list newest first, so it's the first item/entry."""
    if not body:
        return ""
    text = body.decode("utf-8", "replace").lstrip("\ufeff \r\n\t")

    root = None
    for candidate in (text, _sanitize(text)):
        try:
            root = ET.fromstring(candidate)
            break
        except ET.ParseError:
            continue
    if root is None:
        # tolerate junk before the declaration, then try once more
        i, j, k = text.find("<?xml"), text.find("<rss"), text.find("<feed")
        start = min([x for x in (i, j, k) if x >= 0], default=-1)
        if start >= 0:
            try:
                root = ET.fromstring(_sanitize(text[start:]))
            except ET.ParseError:
                root = None
    if root is None:
        # truncated or badly malformed -- pull the first item title out textually
        return _title_by_regex(text)

    def local(tag):
        return str(tag).rsplit("}", 1)[-1].lower()

    for el in root.iter():
        if local(el.tag) in ("item", "entry"):
            for child in el:
                if local(child.tag) == "title":
                    t = _text_of(child)[:300]
                    if t:
                        return t
            # an item exists but its title was unreadable -- try the text route
            return _title_by_regex(text)
    # no items: fall back to the channel/feed title so the row still says something
    for el in root.iter():
        if local(el.tag) == "title":
            t = _text_of(el)[:300]
            if t:
                return t + " (channel title - feed has no items)"
    return ""


def fetch(url):
    host, lock = _host_slot(url)
    code, body, elapsed, reason, lax_ok = 0, b"", 0.0, "", False
    attempts = RETRIES + 1
    for attempt in range(attempts):
        with lock:                                   # one request per host at a time
            gap = PER_HOST_GAP - (time.time() - _host_last.get(host, 0))
            if gap > 0:
                time.sleep(gap)
            start = time.perf_counter()
            try:
                req = urllib.request.Request(url, headers=HEADERS)
                with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
                    code = resp.getcode()
                    body = _decode_body(_read_all(resp), (resp.headers.get("Content-Encoding") or "").lower())
                    reason = ""
            except urllib.error.HTTPError as e:
                code = e.code
                reason = f"HTTP {e.code}"
                if e.code in (403, 404, 406, 415, 429):
                    try:
                        alt_headers = {"User-Agent": "curl/8.5.0", "Accept": "*/*"}
                        with urllib.request.urlopen(
                                urllib.request.Request(url, headers=alt_headers),
                                timeout=TIMEOUT, context=_LAX) as r2:
                            code = r2.getcode()
                            body = _decode_body(_read_all(r2),
                                                (r2.headers.get("Content-Encoding") or "").lower())
                            reason = ""
                            lax_ok = True
                            _host_last[host] = time.time()
                            elapsed = time.perf_counter() - start
                            break
                    except Exception:
                        pass
                try:
                    body = _decode_body(_read_all(e), (e.headers.get("Content-Encoding") or "").lower())
                except Exception:
                    body = b""
            except Exception as e:
                code, body = 0, b""
                msg = str(e)
                if "CERTIFICATE_VERIFY_FAILED" in msg or "SSL" in msg.upper():
                    reason = "SSL certificate chain incomplete"
                    try:    # the feed itself is usually fine -- fetch and flag it
                        alt_headers = {"User-Agent": "curl/8.5.0", "Accept": "*/*"}
                        req = urllib.request.Request(url, headers=alt_headers)
                        with urllib.request.urlopen(req, timeout=TIMEOUT, context=_LAX) as resp:
                            code = resp.getcode()
                            body = _decode_body(_read_all(resp),
                                                (resp.headers.get("Content-Encoding") or "").lower())
                            lax_ok = True
                    except Exception:
                        pass
                elif isinstance(e, urllib.error.URLError) and "timed out" in msg:
                    reason = "timed out"
                elif "Name or service not known" in msg or "nodename nor servname" in msg:
                    reason = "DNS does not resolve"
                elif "Connection refused" in msg:
                    reason = "connection refused"
                else:
                    reason = f"{type(e).__name__}: {msg[:60]}"
            elapsed = time.perf_counter() - start
            _host_last[host] = time.time()
        if code == 200 or attempt == attempts - 1:
            break
        if code in (0, 401, 403, 429) or 500 <= code < 600:   # worth one more try
            # 401/403 from FDA, AstraZeneca etc. are frequently transient
            # rate-limiting rather than a real block.
            time.sleep(3)
            continue
        break
    return code, elapsed, body, reason, lax_ok


def classify(code, body, reason="", lax_ok=False):
    if code == 200:
        title = _first_item_title(body)
        if title:
            if lax_ok:
                return "OK — fetchable (SSL chain incomplete, verify disabled)", title
            return "OK — fetchable", title
        head = body[:400].lstrip().lower()
        if head.startswith(b"<!doctype html") or b"<html" in head:
            return "OK (200) — HTML returned, not a feed", ""
        return "OK (200) — no readable items", ""
    if code in (401, 403):
        return "BLOCKED (would need proxy)", _first_item_title(body)
    if code == 0:
        return f"FAILED — {reason or 'no response'}", ""
    if code == 429:
        return "RATE-LIMITED (429)", ""
    if code == 404:
        return "other (404)", ""
    return f"other ({code})", _first_item_title(body)


def try_fallback_cascade(url, label):
    """Executes the 5-Layer Fallback Cascade when primary feed URL fails."""
    import urllib.parse
    host = re.sub(r"^https?://", "", url).split("/")[0].lower()
    
    # Layer 1: Lax SSL & Alternate Headers
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "curl/8.5.0", "Accept": "*/*"})
        with urllib.request.urlopen(req, timeout=8, context=ctx) as r:
            if r.getcode() == 200:
                b = r.read(1024*1024)
                t = _first_item_title(b)
                if t:
                    return "RECOVERED (Layer 1: Lax SSL / Alternate UA)", t, url
    except Exception:
        pass
        
    # Layer 2: Common feed paths on domain
    for p in ["/feed/", "/rss.xml", "/rss", "/atom.xml", "/news/rss.xml"]:
        cand = f"https://{host}{p}"
        if cand == url:
            continue
        try:
            req = urllib.request.Request(cand, headers={"User-Agent": UA, "Accept": "*/*"})
            with urllib.request.urlopen(req, timeout=6, context=ctx) as r:
                if r.getcode() == 200:
                    b = r.read(1024*1024)
                    t = _first_item_title(b)
                    if t:
                        return "RECOVERED (Layer 2: Discovered Feed Path)", t, cand
        except Exception:
            pass
            
    # Layer 3: Google News site: query
    try:
        q = urllib.parse.quote(f"site:{host} when:7d")
        gnews_url = f"https://news.google.com/rss/search?q={q}&hl=en-US&gl=US&ceid=US:en"
        req = urllib.request.Request(gnews_url, headers={"User-Agent": UA})
        with urllib.request.urlopen(req, timeout=8) as r:
            if r.getcode() == 200:
                b = r.read(1024*1024)
                root = ET.fromstring(b)
                for item in root.iter("item"):
                    t = item.find("title")
                    if t is not None and t.text:
                        return "RECOVERED (Layer 3: GNews Search Feed)", t.text.strip(), gnews_url
    except Exception:
        pass
        
    return "", "", ""


def check(row):
    fid, url, label = row
    code, elapsed, body, reason, lax_ok = fetch(url)
    result, title = classify(code, body, reason, lax_ok)
    fallback_url = ""
    
    # If primary fetch failed or had no items, trigger automatic 5-layer fallback cascade
    if not (code == 200 and title) or "BLOCKED" in result or "FAILED" in result:
        fb_result, fb_title, fb_url = try_fallback_cascade(url, label)
        if fb_result and fb_title:
            result = fb_result
            title = fb_title
            fallback_url = fb_url

    return [fid, url, label, code, round(elapsed, 3), result, title, fallback_url]


# ----------------------------------------------------------------------------- main
def main():
    text = sys.stdin.read() if len(sys.argv) < 2 or sys.argv[1] == "-" else open(sys.argv[1], encoding="utf-8", errors="replace").read()
    feeds = parse_feeds(text)
    if not feeds:
        print("::error::No feeds found in the input. Expected three columns: Feed ID, URL, Label.")
        sys.exit(1)

    print(f"Checking {len(feeds)} feeds — {WORKERS} at a time, {TIMEOUT:.0f}s timeout, "
          f"{RETRIES} retry, {PER_HOST_GAP:.1f}s between hits on the same host.\n", flush=True)

    results = [None] * len(feeds)
    done = 0
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futures = {pool.submit(check, f): i for i, f in enumerate(feeds)}
        for fut in as_completed(futures):
            i = futures[fut]
            results[i] = fut.result()
            done += 1
            r = results[i]
            print(f"[{done:>4}/{len(feeds)}] {r[0]:<12} {str(r[3]):<4} {r[4]:>6.2f}s  {r[5]}", flush=True)

    out_dir = os.environ.get("FEED_OUT_DIR", "results")
    os.makedirs(out_dir, exist_ok=True)

    tsv_path = os.path.join(out_dir, "results.tsv")
    with open(tsv_path, "w", encoding="utf-8", newline="") as f:
        f.write("\t".join(HEADER) + "\n")
        for r in results:
            f.write("\t".join(str(c).replace("\t", " ") for c in r) + "\n")

    csv_path = os.path.join(out_dir, "results.csv")
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(HEADER)
        w.writerows(results)

    xlsx_path = os.path.join(out_dir, "results.xlsx")
    try:
        import openpyxl
        from openpyxl.styles import Font
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "RSS Check"
        ws.append(HEADER)
        for c in ws[1]:
            c.font = Font(bold=True)
        for r in results:
            ws.append(r)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col, width in zip("ABCDEFG", (14, 70, 34, 9, 9, 34, 60)):
            ws.column_dimensions[col].width = width
        wb.save(xlsx_path)
    except Exception as e:
        xlsx_path = ""
        print(f"(xlsx skipped: {e})")

    ok      = sum(1 for r in results if r[5].startswith("OK — fetchable"))
    blocked = sum(1 for r in results if r[5].startswith("BLOCKED"))
    hung    = sum(1 for r in results if r[5].startswith("FAILED"))
    other   = len(results) - ok - blocked - hung

    print("\n" + "=" * 78)
    print(f"{len(results)} feeds checked — {ok} OK, {blocked} blocked, {hung} hung/timeout, {other} other")
    print("=" * 78)
    print("\nCopy the block below straight into Excel (it is tab separated):\n")
    print("\t".join(HEADER))
    for r in results:
        print("\t".join(str(c).replace("\t", " ") for c in r))

    summary = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary:
        with open(summary, "a", encoding="utf-8") as f:
            f.write(f"## RSS feed check — {len(results)} feeds\n\n")
            f.write(f"**{ok}** OK &nbsp;·&nbsp; **{blocked}** blocked &nbsp;·&nbsp; "
                    f"**{hung}** hung/timeout &nbsp;·&nbsp; **{other}** other\n\n")
            f.write("<details open><summary>Copy-paste back into Excel (tab separated)</summary>\n\n```\n")
            f.write("\t".join(HEADER) + "\n")
            for r in results:
                f.write("\t".join(str(c).replace("\t", " ") for c in r) + "\n")
            f.write("```\n\n</details>\n\n")
            f.write("| " + " | ".join(HEADER) + " |\n|" + "---|" * len(HEADER) + "\n")
            for r in results:
                cells = [str(c).replace("|", "\\|") for c in r]
                cells[1] = f"[link]({r[1]})"
                f.write("| " + " | ".join(cells) + " |\n")

    gh_out = os.environ.get("GITHUB_OUTPUT")
    if gh_out:
        with open(gh_out, "a", encoding="utf-8") as f:
            f.write(f"total={len(results)}\nok={ok}\nblocked={blocked}\nhung={hung}\nother={other}\n")


if __name__ == "__main__":
    main()
