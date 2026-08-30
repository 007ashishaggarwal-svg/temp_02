#!/usr/bin/env python3
import urllib.request
import ssl
import re
import sys
import base64

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"

def resolve_google_news_url(gurl: str, timeout: int = 5) -> str:
    """Resolve Google News redirect wrapper to authentic authoritative URL."""
    if not gurl or "news.google.com" not in gurl:
        return gurl

    # Attempt 1: Try direct HTTP fetch to follow 301/302 redirects or parse HTML destination
    req = urllib.request.Request(gurl, headers={"User-Agent": UA, "Accept": "text/html,*/*"})
    try:
        with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
            final_url = resp.geturl()
            if "news.google.com" not in final_url:
                return final_url
            
            # Read first 128KB of HTML body to locate target link
            body = resp.read(128 * 1024).decode("utf-8", "ignore")
            
            # Pattern 1: data-n-a-u attribute (Google News standard)
            m = re.search(r'data-n-a-u=["\'](https?://[^"\']+)["\']', body)
            if m:
                return m.group(1)
            
            # Pattern 2: Canonical link
            m_can = re.search(r'<link\s+rel=["\']canonical["\']\s+href=["\'](https?://(?!news\.google)[^"\']+)["\']', body)
            if m_can:
                return m_can.group(1)
            
            # Pattern 3: Top-level outbound anchor
            m_a = re.search(r'<a\s+[^>]*href=["\'](https?://(?!news\.google|accounts\.google|support\.google|www\.google)[^"\']+)["\']', body)
            if m_a:
                return m_a.group(1)
            
            # Pattern 4: JS window.location redirect
            m_js = re.search(r'window\.location\.replace\(["\'](https?://[^"\']+)["\']\)', body)
            if m_js:
                return m_js.group(1)
                
    except Exception as e:
        pass

    return gurl

if __name__ == "__main__":
    test_urls = [
        "https://news.google.com/rss/articles/CBMi4gFBVV95cUxQVFI5c2NtaEFIbHBYOG80aklkLThZSU9lcnFHZTVBMjV3ZUpIMnZ5eUZYb1hlLTBsc1NMNDRoR2xtbVB0YnppUkhkRXNIeWRCcUF3VVRrdTJLU2pnVVNuMUVaUWlyQ0Vsc0NLcWpyRnZ2N3o4SVdCUWIxdDg3c29YRjBPdGF0LW5xWUY5WHc1ZWVrSGxNeVlxVTZyMzVFWXhTVUFRRWtieEZQSXpzQ21NTy1MS0tKOEx1YmhpYkZsVFRrenRPTGpadHJqYmdoZVVJVEFZOC1SOHBGQThaRUNfd3NR0gHnAUFVX3lxTE5QRmpSV0JsLWtRV1BrT1VwMDgwWEhsWFVHbVFTMWVjMXQwTEFyTVZfd2ZvUS1ZaVIzUWJaaXRBOEx3SUVNR2s5SHpjajdqWjFNSllRTUNZbDNFRUs5LWItTUJYRE1QZmROUm9PSkNLR2FiZzZ4c2tDcmpkcVZ5VGtHd0ZGVF9LakUxZTNWbzNBVVFyUE53X05SdUliTFBrNFQ2bmN3aVZEQ25uWWlqUkswam82RklUSnM0WklXV3ByOXBKVDc0RTkzek0zc2M3ZHJFTGRaamRKb2hwNUNUSHFOQlJfdFZqbw?oc=5",
        "https://news.google.com/rss/articles/CBMi0AFBVV95cUxOVEFUUnJmYWs1bHFuUHdpTWdEWUo1cjA5bGZVOHQxVF81eldoMUdmTW9xdDBMX01tSk85VktmQ19OUlcxc3VvMHhKaDVWR0RlczVsYThhZmpDaHpGdWF6WUhSWW1IS2J6SXNJbjdOUFhYS0MtTDlWMGhac1czSmlyekhtTDk2WWg5N180elU4UU9iZnZNdWp4cG1PcGdLWWdxWjVLQ01uTEpxTlB1R083VjRWMGJxeXRaUmlMOFV6ZENSc1gxNXF3UmItWUZfN0FJ?oc=5"
    ]
    for u in test_urls:
        print("Input:", u[:60], "...")
        resolved = resolve_google_news_url(u)
        print("Resolved:", resolved)
        print("-" * 60)
