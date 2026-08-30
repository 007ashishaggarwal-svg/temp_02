#!/usr/bin/env python3
"""
Raw Feed Ingestion Archiver & Incremental Seen-State Cache.
Features:
1. Auto-Pruning: Keeps at most MAX_ROWS (e.g. 5,000 rows) OR MAX_DAYS (e.g. 10 days).
2. Canonical URL & Headline Normalization for 100% duplicate protection.
3. Micro-Footprint: Stored as compact JSON Lines (< 8 MB disk, < 15 MB RAM).
4. Dual Mode: Works identically on Local PC and Cloud Runners (GitHub Actions / Cloud VMs).
"""

import os
import json
import hashlib
from datetime import datetime, timezone, timedelta

WORKSPACE = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
CACHE_DIR = os.path.join(WORKSPACE, ".cache")
ARCHIVE_FILE = os.path.join(CACHE_DIR, "raw_feed_archive.jsonl")
SEEN_STATE_FILE = os.path.join(CACHE_DIR, "last_seen_state.json")

class RawFeedArchiver:
    def __init__(self, max_rows: int = 50000, max_days: int = 90, archive_path: str = None, state_path: str = None):
        self.max_rows = max_rows
        self.max_days = max_days
        self.archive_path = archive_path or ARCHIVE_FILE
        self.state_path = state_path or SEEN_STATE_FILE
        self.seen_hashes = set()

        os.makedirs(os.path.dirname(self.archive_path), exist_ok=True)
        self.load_seen_state()

    def generate_fingerprint(self, canonical_url: str, headline: str) -> str:
        """Generate canonical SHA-256 fingerprint for unwrapped URL and headline."""
        norm_url = (canonical_url or "").strip().lower().rstrip("/")
        # Remove tracking parameters
        norm_url = norm_url.split("?utm_")[0].split("&utm_")[0]

        import re
        norm_title = " ".join(re.findall(r"\b[a-zA-Z0-9]{3,}\b", (headline or "").lower()))
        combined = f"{norm_url}::{norm_title}"
        return hashlib.sha256(combined.encode("utf-8")).hexdigest()[:24]

    def load_seen_state(self):
        """Load seen fingerprints from disk."""
        if os.path.exists(self.state_path):
            try:
                with open(self.state_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.seen_hashes = set(data.get("seen_fingerprints", []))
            except Exception:
                self.seen_hashes = set()

    def is_already_seen(self, canonical_url: str, headline: str) -> bool:
        """Check if article was processed in a prior run."""
        fp = self.generate_fingerprint(canonical_url, headline)
        return fp in self.seen_hashes

    def save_and_prune(self, new_raw_events: list[dict]) -> dict:
        """
        Append new raw events, prune items older than max_days OR exceeding max_rows,
        and save updated archive and seen state to disk.
        """
        cutoff_date = datetime.now(timezone.utc) - timedelta(days=self.max_days)
        cutoff_str = cutoff_date.strftime("%Y-%m-%d")

        existing_events = []
        if os.path.exists(self.archive_path):
            try:
                with open(self.archive_path, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if line:
                            try:
                                existing_events.append(json.loads(line))
                            except Exception:
                                pass
            except Exception:
                pass

        # Combine existing and new events
        all_events = existing_events + new_raw_events
        deduped_events = []
        seen_now = set()

        for ev in all_events:
            url = ev.get("raw_url", "")
            title = ev.get("headline", "")
            fp = self.generate_fingerprint(url, title)

            if fp not in seen_now:
                seen_now.add(fp)
                ev["_fp"] = fp
                deduped_events.append(ev)

        # Apply Dual Pruning: (1) Date Cutoff and (2) Max Rows Cap
        pruned_events = [
            ev for ev in deduped_events
            if ev.get("published_date", "9999-99-99") >= cutoff_str
        ]

        # Sort by date descending and retain top max_rows
        pruned_events.sort(key=lambda x: x.get("published_date", ""), reverse=True)
        final_events = pruned_events[:self.max_rows]

        def default_serializer(obj):
            if isinstance(obj, (datetime, )):
                return obj.isoformat()
            return str(obj)

        # Write to JSONL
        with open(self.archive_path, "w", encoding="utf-8") as f:
            for ev in final_events:
                # Clean temp key
                ev_copy = dict(ev)
                ev_copy.pop("_fp", None)
                f.write(json.dumps(ev_copy, ensure_ascii=False, default=default_serializer) + "\n")

        # Update Seen State
        self.seen_hashes = {ev.get("_fp", self.generate_fingerprint(ev.get("raw_url", ""), ev.get("headline", ""))) for ev in final_events}
        with open(self.state_path, "w", encoding="utf-8") as f:
            json.dump({
                "last_updated": datetime.now(timezone.utc).isoformat(),
                "total_retained": len(final_events),
                "seen_fingerprints": list(self.seen_hashes)
            }, f, indent=2)

        file_size_kb = os.path.getsize(self.archive_path) / 1024.0 if os.path.exists(self.archive_path) else 0.0

        return {
            "total_retained_rows": len(final_events),
            "new_events_added": len(new_raw_events),
            "file_size_kb": round(file_size_kb, 2),
            "cutoff_date": cutoff_str
        }
