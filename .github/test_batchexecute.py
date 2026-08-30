#!/usr/bin/env python3
import urllib.request
import urllib.parse
import json
import ssl
import re
import sys

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

def decode_google_news_url(gurl: str) -> str:
    """Decode a Google News RSS article URL to its authentic destination URL."""
    if not gurl or "news.google.com" not in gurl:
        return gurl
    
    # Extract token
    m = re.search(r"articles/([^/?]+)", gurl)
    if not m:
        return gurl
    token = m.group(1)

    url = "https://news.google.com/_/DotsSplashUi/data/batchexecute"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
        "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8"
    }

    payload_inner = json.dumps(["garturlreq", [["en-US", "US", ["FINANCE_TOP_INDICES", "WEB_TEST_1_0_0"], None, None, 1, 1, "US:en", None, 180, None, None, None, None, None, 0, None, None, [1608, 435], 1], token]])
    payload_outer = json.dumps([[["Fbv4je", payload_inner, None, "generic"]]])
    post_data = urllib.parse.urlencode({"f.req": payload_outer}).encode("utf-8")

    try:
        req = urllib.request.Request(url, data=post_data, headers=headers)
        with urllib.request.urlopen(req, timeout=6, context=ctx) as resp:
            text = resp.read().decode("utf-8", "ignore")
            # Parse batchexecute envelope
            for line in text.splitlines():
                if "http" in line and "google.com" not in line:
                    urls = re.findall(r'https?://[^\s"\'\\]+', line)
                    for u in urls:
                        if not u.startswith("https://news.google") and not u.startswith("https://www.google"):
                            return u.replace(r"\u0026", "&").replace(r"\u003d", "=")
    except Exception as e:
        pass

    return gurl

if __name__ == "__main__":
    test_urls = [
        "https://news.google.com/rss/articles/CBMi4gFBVV95cUxQVFI5c2NtaEFIbHBYOG80aklkLThZSU9lcnFHZTVBMjV3ZUpIMnZ5eUZYb1hlLTBsc1NMNDRoR2xtbVB0YnppUkhkRXNIeWRCcUF3VVRrdTJLU2pnVVNuMUVaUWlyQ0Vsc0NLcWpyRnZ2N3o4SVdCUWIxdDg3c29YRjBPdGF0LW5xWUY5WHc1ZWVrSGxNeVlxVTZyMzVFWXhTVUFRRWtieEZQSXpzQ21NTy1MS0tKOEx1YmhpYkZsVFRrenRPTGpadHJqYmdoZVVJVEFZOC1SOHBGQThaRUNfd3NR0gHnAUFVX3lxTE5QRmpSV0JsLWtRV1BrT1VwMDgwWEhsWFVHbVFTMWVjMXQwTEFyTVZfd2ZvUS1ZaVIzUWJaaXRBOEx3SUVNR2s5SHpjajdqWjFNSllRTUNZbDNFRUs5LWItTUJYRE1QZmROUm9PSkNLR2FiZzZ4c2tDcmpkcVZ5VGtHd0ZGVF9LakUxZTNWbzNBVVFyUE53X05SdUliTFBrNFQ2bmN3aVZEQ25uWWlqUkswam82RklUSnM0WklXV3ByOXBKVDc0RTkzek0zc2M3ZHJFTGRaamRKb2hwNUNUSHFOQlJfdFZqbw?oc=5",
        "https://news.google.com/rss/articles/CBMi0AFBVV95cUxOVEFUUnJmYWs1bHFuUHdpTWdEWUo1cjA5bGZVOHQxVF81eldoMUdmTW9xdDBMX01tSk85VktmQ19OUlcxc3VvMHhKaDVWR0RlczVsYThhZmpDaHpGdWF6WUhSWW1IS2J6SXNJbjdOUFhYS0MtTDlWMGhac1czSmlyekhtTDk2WWg5N180elU4UU9iZnZNdWp4cG1PcGdLWWdxWjVLQ01uTEpxTlB1R083VjRWMGJxeXRaUmlMOFV6ZENSc1gxNXF3UmItWUZfN0FJ?oc=5"
    ]
    for u in test_urls:
        print("Original:", u[:50], "...")
        resolved = decode_google_news_url(u)
        print("Decoded Authentic Link:", resolved)
        print("-" * 70)
