"""
Automated maintenance and testing utility.
"""
import os
import sys
import json
import requests

def main():
    endpoint = os.environ.get("ENDPOINT_URL", "").strip()
    token = os.environ.get("RUNNER_TOKEN", "").strip()

    if not endpoint or not token:
        print("[ERROR] Missing required runner configuration.")
        sys.exit(1)

    # Normalize target discovery and ingest URLs
    if endpoint.endswith("/relay-ingest"):
        discovery_url = endpoint
        ingest_url = endpoint
    elif endpoint.endswith("/scheduler"):
        discovery_url = f"{endpoint}?action=relay-targets"
        ingest_url = f"{endpoint}?action=relay-ingest"
    else:
        base = endpoint.rstrip("/")
        discovery_url = f"{base}/api/public/relay-ingest"
        ingest_url = f"{base}/api/public/relay-ingest"

    auth_headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    # 1. Discover targets
    try:
        res = requests.get(discovery_url, headers=auth_headers, timeout=15)
        if res.status_code != 200:
            print(f"[ERROR] Target discovery failed with status {res.status_code}.")
            sys.exit(1)
        data = res.json()
        targets = data.get("targets", [])
    except Exception:
        print("[ERROR] Failed to discover targets.")
        sys.exit(1)

    print(f"[INFO] Discovered {len(targets)} targets.")
    if not targets:
        print("[INFO] No targets pending rescue.")
        return

    # 2. Fetch and ingest target payloads
    fetch_headers = {
        "User-Agent": "Mozilla/5.0 (compatible; Feeder/1.0; +https://feeder.co)",
        "Accept": "application/rss+xml, application/atom+xml, application/xml, text/xml;q=0.9, */*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
    }

    total_synced = 0
    total_added = 0
    failed_count = 0

    for target in targets:
        feed_id = target.get("id")
        feed_url = target.get("feed_url")
        if not feed_id or not feed_url:
            continue

        try:
            r = requests.get(feed_url, headers=fetch_headers, timeout=20)
            if r.status_code == 200 and r.text and ("<" in r.text):
                # Post individual feed ingest to avoid worker request timeouts
                post_res = requests.post(
                    ingest_url,
                    headers=auth_headers,
                    json={"relay_results": [{"feed_id": feed_id, "xml": r.text}]},
                    timeout=30
                )
                if post_res.status_code == 200:
                    res_data = post_res.json()
                    total_synced += res_data.get("processed", 0)
                    total_added += res_data.get("added", 0)
                else:
                    failed_count += 1
            else:
                failed_count += 1
        except Exception:
            failed_count += 1

    print(f"[INFO] Batch sync complete: {total_synced} targets processed, {total_added} entries added, {failed_count} errors.")

if __name__ == "__main__":
    main()
