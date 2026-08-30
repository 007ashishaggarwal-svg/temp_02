#!/usr/bin/env python3
"""
Indication feed generator.

Given a list of indications, produce WORKING feed URLs that cover each one —
so a newly added therapy area is covered immediately, without waiting for any
publisher to create a feed for it.

Input : Indication | Synonyms (comma separated) | ClinicalTrials.gov condition
Output: rows ready to paste into the SOURCES sheet:
        URL | Type | Label | Projects | Status | Last Status | Last Item Title | Health | Priority

Every generated URL is fetched and parsed before it is emitted. Anything that
does not return real items is reported, not silently included.

Only the Python standard library is required.
"""

import csv
import gzip
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

TIMEOUT = float(os.environ.get("IND_TIMEOUT", "30"))
WORKERS = int(os.environ.get("IND_WORKERS", "4"))
OUT_DIR = os.environ.get("IND_OUT_DIR", "results")
# ClinicalTrials.gov and PubMed refuse anonymous datacentre traffic; both ask for
# a contact address in the User-Agent. Set repo variable CONTACT_UA to
# "Your Name your@email.com" to enable those two feed families.
CONTACT_UA = os.environ.get("CONTACT_UA", "").strip()

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")

HEADER = ["URL", "Type", "Label", "Projects (ALL or comma-separated)", "Status",
          "Last Status", "Last Item Title", "Health", "Priority"]


def fetch(url, ua=None):
    req = urllib.request.Request(url, headers={
        "User-Agent": ua or UA,
        "Accept": "application/rss+xml, application/atom+xml, application/xml;q=0.9, */*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
    })
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
            b = r.read(4 * 1024 * 1024)
            if b[:2] == b"\x1f\x8b":
                b = gzip.decompress(b)
            return r.getcode(), b
    except urllib.error.HTTPError as e:
        return e.code, b""
    except Exception:
        return 0, b""


def feed_items(body):
    """(count, first title) — only trusts a real parse."""
    if not body:
        return 0, ""
    text = body.decode("utf-8", "replace").lstrip("﻿ \r\n\t")
    text = re.sub(r"&(?!(?:amp|lt|gt|quot|apos|#\d+|#[xX][0-9a-fA-F]+);)", "&amp;", text)
    try:
        root = ET.fromstring(text)
    except ET.ParseError:
        return 0, ""

    def local(t):
        return str(t).rsplit("}", 1)[-1].lower()

    n, first = 0, ""
    for el in root.iter():
        if local(el.tag) in ("item", "entry"):
            n += 1
            if not first:
                for c in el:
                    if local(c.tag) == "title":
                        first = " ".join("".join(c.itertext()).split())[:200]
                        break
    return n, first


def q(s):
    return urllib.parse.quote(s, safe="")


# Every Google News query is ANDed with this, so results stay inside pharma.
# Without it, "MASH" returns high-school football and "IBD" returns Investor's
# Business Daily — both observed in a real run.
PHARMA_CONTEXT = ("(drug OR therapy OR treatment OR patients OR clinical OR "
                  "trial OR FDA OR biotech OR pharmaceutical)")

# Project names that are ordinary English words. These are kept as the LABEL but
# never used as a search term — the synonyms carry the meaning instead.
# "Safety" is the clearest case: even with the pharma context it returns the Meta
# child-safety trial, because "safety" and "trial" co-occur there.
GENERIC_NAMES = {
    "safety", "financing", "regulatory", "companies", "general", "news",
    "other", "deals and m&a", "deals", "m&a", "finance", "business",
}


def candidates(indication, synonyms, ctgov, exclude=""):
    """Every feed family worth generating for one indication."""
    syns = [t.strip() for t in synonyms.split(",") if t.strip()]
    if indication.strip().lower() in GENERIC_NAMES:
        terms = syns                       # the name alone is too generic to search
    else:
        terms = [indication] + syns
    if not terms:
        terms = [indication]
    core = "(" + " OR ".join(f'"{t}"' for t in terms[:6]) + ")"
    # Google News supports negative terms. Needed where one indication's name is
    # a substring of another's: "Small Cell Lung Cancer" matches every NSCLC
    # story until "-non-small" is added.
    neg = "".join(f' -"{t.strip()}"' for t in exclude.split(",") if t.strip())
    core = core + neg
    broad = f"{core} {PHARMA_CONTEXT}"
    out = [
        ("Google News — broad",
         f"https://news.google.com/rss/search?q={q(broad)}&hl=en-US&gl=US&ceid=US:en",
         f"GN-{indication}", "HIGH", None),
        ("Google News — clinical/regulatory",
         f"https://news.google.com/rss/search?q="
         f"{q(f'{core} (trial OR phase OR FDA OR approval OR topline OR data)')}"
         f"&hl=en-US&gl=US&ceid=US:en",
         f"GN-{indication}-Clinical", "MEDIUM", None),
    ]
    if ctgov.strip():
        # NOTE: adding "&count=" to this URL makes ClinicalTrials.gov return
        # 400 Bad Request. Without it the feed works and returns 13-60 trials.
        # Check your existing SOURCES rows for that parameter.
        out.append((
            "ClinicalTrials.gov",
            f"https://clinicaltrials.gov/api/rss?cond={q(ctgov.strip())}",
            f"CTgov-{indication}", "MEDIUM", None))
    return out


# PubMed is deliberately NOT generated here. Its /rss/search/ endpoint requires a
# feed ID created by hand in the PubMed UI and returns 500 to a constructed URL.
# Literature is already covered by the journal feeds in feeds.tsv, and your
# gatherer reaches PubMed through E-utilities separately.


def check(job):
    indication, family, url, label, priority, ua = job
    code, body = fetch(url, ua)
    n, first = feed_items(body)
    ok = code == 200 and n > 0
    return {"indication": indication, "family": family, "url": url, "label": label,
            "priority": priority, "status": code, "items": n, "first": first, "ok": ok}


def parse_input(text):
    rows = []
    for raw in text.splitlines():
        line = raw.strip().lstrip("﻿")
        if not line or line.startswith("#"):
            continue
        parts = [p.strip() for p in (line.split("\t") if "\t" in line
                 else next(csv.reader([line])))]
        while len(parts) < 4:
            parts.append("")
        if parts[0].lower() in ("indication", "project", "project name"):
            continue
        rows.append((parts[0], parts[1], parts[2], parts[3]))
    return rows


def main():
    text = (open(sys.argv[1], encoding="utf-8", errors="replace").read()
            if len(sys.argv) > 1 and sys.argv[1] != "-" else sys.stdin.read())
    inds = parse_input(text)
    if not inds:
        print("::error::No indications found. Expected: Indication, Synonyms, "
              "ClinicalTrials.gov condition, Exclude terms")
        sys.exit(1)

    jobs = []
    for indication, syn, ctgov, excl in inds:
        for family, url, label, priority, ua in candidates(indication, syn, ctgov, excl):
            jobs.append((indication, family, url, label, priority, ua))

    print(f"Generating and verifying {len(jobs)} feeds for {len(inds)} indications, "
          f"{WORKERS} at a time\n", flush=True)

    results = []
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        for i, r in enumerate(pool.map(check, jobs), 1):
            results.append(r)
            mark = "OK  " if r["ok"] else "FAIL"
            print(f"[{i:>3}/{len(jobs)}] {mark} {str(r['status']):>3} "
                  f"{r['items']:>4} items  {r['label'][:36]:<38} {r['family']}", flush=True)
            time.sleep(0.15)

    os.makedirs(OUT_DIR, exist_ok=True)
    good = [r for r in results if r["ok"]]
    bad = [r for r in results if not r["ok"]]

    path = os.path.join(OUT_DIR, "indication_sources.tsv")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\t".join(HEADER) + "\n")
        for r in good:
            f.write("\t".join([r["url"], "RSS", r["label"], r["indication"], "Active",
                               "", r["first"].replace("\t", " "), "", r["priority"]]) + "\n")

    print("\n" + "=" * 78)
    print(f"{len(good)}/{len(results)} feeds verified working, covering "
          f"{len({r['indication'] for r in good})}/{len(inds)} indications")
    print("=" * 78)
    if bad:
        print("\nNot working (not written to the output file):")
        for r in bad:
            hint = ""
            if "clinicaltrials.gov" in r["url"]:
                hint = "  <- try again; ClinicalTrials.gov rate-limits bursts"
            print(f"  {str(r['status']):>3}  {r['label'][:40]:<42} {r['family']}{hint}")

    print(f"\nPaste the block below into your SOURCES sheet (tab separated):\n")
    print("\t".join(HEADER))
    for r in good:
        print("\t".join([r["url"], "RSS", r["label"], r["indication"], "Active",
                         "", r["first"].replace("\t", " "), "", r["priority"]]))

    summary = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary:
        with open(summary, "a", encoding="utf-8") as f:
            f.write(f"## Indication feeds — {len(good)} working\n\n")
            f.write(f"Covering {len({r['indication'] for r in good})} of {len(inds)} "
                    f"indications.\n\n")
            f.write("<details open><summary>Paste into SOURCES (tab separated)"
                    "</summary>\n\n```\n")
            f.write("\t".join(HEADER) + "\n")
            for r in good:
                f.write("\t".join([r["url"], "RSS", r["label"], r["indication"],
                                   "Active", "", r["first"].replace("\t", " "),
                                   "", r["priority"]]) + "\n")
            f.write("```\n\n</details>\n")


if __name__ == "__main__":
    main()
