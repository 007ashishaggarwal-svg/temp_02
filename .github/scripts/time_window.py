#!/usr/bin/env python3
"""
Universal Time-Windowing & Guardrail Engine for RSSFeedChecker.
Supports:
1. Relative Windows: '24h', '48h', '72h', '3d', '7d', 'last 24 hours'
2. Absolute Start Dates: 'From 24 Aug 2026 01:00 AM IST onwards', '2026-08-24 01:00'
3. Bounded Date Ranges: '24 Aug 2026 01:00 AM IST to 25 Aug 2026 01:00 AM IST'
4. Timezone Normalization: IST, UTC, EST, EDT, PST, PDT, BST, CET, etc.
5. 7-Day Guardrail: Standard users cannot query beyond 7 days (168h); requires --admin.
"""

import os
import sys
import re
from datetime import datetime, timezone, timedelta

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Standard Timezone Offset Mapping
TZ_OFFSETS = {
    "UTC": 0, "GMT": 0, "Z": 0,
    "IST": 5.5,    # Indian Standard Time (+05:30)
    "EST": -5,     # Eastern Standard Time (-05:00)
    "EDT": -4,     # Eastern Daylight Time (-04:00)
    "CST": -6,     # Central Standard Time (-06:00)
    "CDT": -5,     # Central Daylight Time (-05:00)
    "MST": -7,     # Mountain Standard Time (-07:00)
    "MDT": -6,     # Mountain Daylight Time (-06:00)
    "PST": -8,     # Pacific Standard Time (-08:00)
    "PDT": -7,     # Pacific Daylight Time (-07:00)
    "BST": 1,      # British Summer Time (+01:00)
    "CET": 1,      # Central European Time (+01:00)
    "CEST": 2,     # Central European Summer Time (+02:00)
    "JST": 9,      # Japan Standard Time (+09:00)
    "KST": 9,      # Korea Standard Time (+09:00)
    "AEST": 10,    # Australian Eastern Standard Time (+10:00)
}

MAX_STANDARD_WINDOW_DAYS = 7.0  # Max 7 days for standard non-admin queries


class TimeWindow:
    def __init__(self, start_utc: datetime, end_utc: datetime, is_admin: bool = False, raw_input: str = ""):
        self.start_utc = start_utc
        self.end_utc = end_utc
        self.is_admin = is_admin
        self.raw_input = raw_input
        self.duration_hours = (end_utc - start_utc).total_seconds() / 3600.0
        self.duration_days = self.duration_hours / 24.0

    def is_in_window(self, dt: datetime) -> bool:
        """Check if a datetime falls within [start_utc, end_utc]."""
        if not dt:
            return False
        # Normalize naive dt to UTC if needed
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        else:
            dt = dt.astimezone(timezone.utc)
        return self.start_utc <= dt <= self.end_utc

    def format_summary(self) -> str:
        """Human-readable window summary."""
        ist_tz = timezone(timedelta(hours=5, minutes=30))
        s_ist = self.start_utc.astimezone(ist_tz).strftime("%Y-%m-%d %I:%M %p IST")
        e_ist = self.end_utc.astimezone(ist_tz).strftime("%Y-%m-%d %I:%M %p IST")
        s_utc = self.start_utc.strftime("%Y-%m-%d %H:%M UTC")
        e_utc = self.end_utc.strftime("%Y-%m-%d %H:%M UTC")
        admin_tag = " [ADMIN OVERRIDE ACTIVE]" if self.is_admin else ""
        return (f"Window: {s_ist} → {e_ist} ({self.duration_hours:.1f}h / {self.duration_days:.2f}d){admin_tag}\n"
                f"        (UTC: {s_utc} → {e_utc})")


def parse_tz_offset(text: str) -> tuple[float, str]:
    """Find and strip timezone name/offset from text. Returns (offset_hours, cleaned_text)."""
    text_clean = text.strip()
    for tz_name, offset in TZ_OFFSETS.items():
        pattern = r"\b" + tz_name + r"\b"
        if re.search(pattern, text_clean, re.I):
            text_clean = re.sub(pattern, "", text_clean, flags=re.I).strip()
            return offset, text_clean
    return 0.0, text_clean


def parse_datetime_flexible(text: str, default_tz_offset: float = 5.5) -> datetime:
    """
    Parse date/time strings flexibly:
    - '24 Aug 2026 01:00 AM IST'
    - '2026-08-24 01:00:00'
    - '2026-08-24'
    - '24-08-2026 01:00'
    """
    text = text.strip()
    offset, text_no_tz = parse_tz_offset(text)
    if offset == 0.0 and not any(z in text.upper() for z in ["UTC", "GMT", "Z"]):
        # Default to IST (+5.5) if not specified, matching user context
        offset = default_tz_offset

    tz = timezone(timedelta(hours=offset))

    # Clean punctuation and prefixes
    cleaned = re.sub(r"^(?:from|on|between|since|after)\s+", "", text_no_tz, flags=re.I).strip()
    cleaned = re.sub(r"\s+onwards$", "", cleaned, flags=re.I).strip()

    formats = [
        "%d %b %Y %I:%M %p",       # 24 Aug 2026 01:00 AM
        "%d %B %Y %I:%M %p",       # 24 August 2026 01:00 AM
        "%d %b %Y %H:%M",          # 24 Aug 2026 13:00
        "%d %B %Y %H:%M",          # 24 August 2026 13:00
        "%d %b %Y",                # 24 Aug 2026
        "%d %B %Y",                # 24 August 2026
        "%Y-%m-%d %H:%M:%S",       # 2026-08-24 01:00:00
        "%Y-%m-%d %H:%M",          # 2026-08-24 01:00
        "%Y-%m-%d %I:%M %p",       # 2026-08-24 01:00 AM
        "%Y-%m-%d",                # 2026-08-24
        "%Y/%m/%d %H:%M",          # 2026/08/24 01:00
        "%Y/%m/%d",                # 2026/08/24
        "%d-%m-%Y %H:%M",          # 24-08-2026 01:00
        "%d-%m-%Y",                # 24-08-2026
    ]

    for fmt in formats:
        try:
            dt = datetime.strptime(cleaned, fmt)
            # Localize to parsed timezone and convert to UTC
            dt_aware = dt.replace(tzinfo=tz)
            return dt_aware.astimezone(timezone.utc)
        except ValueError:
            continue

    raise ValueError(f"Unable to parse date/time string: '{text}'")


def create_time_window(query_str: str = "72h", is_admin: bool = False, now_utc: datetime = None) -> TimeWindow:
    """
    Construct a TimeWindow from user specification with 7-day non-admin guardrail.
    
    Examples of query_str:
    - '24h', '48h', '72h', '3d', '7d', 'last 24 hours'
    - 'From 24 Aug 2026 01:00 AM IST onwards'
    - '24 Aug 2026 01:00 AM IST to 25 Aug 2026 01:00 AM IST'
    - '2026-08-24 01:00 to 2026-08-25 01:00'
    """
    if now_utc is None:
        now_utc = datetime.now(timezone.utc)

    q = (query_str or "72h").strip()
    is_admin = is_admin or os.getenv("ADMIN_OVERRIDE", "").lower() in ("true", "1", "yes")

    # 1. Relative Hours / Days: '24h', '48h', '72h', '3d', '7d', 'last 48 hours'
    m_rel = re.match(r"^(?:last\s+)?(\d+(?:\.\d+)?)\s*(h|hr|hrs|hours?|d|days?)$", q, re.I)
    if m_rel:
        val = float(m_rel.group(1))
        unit = m_rel.group(2).lower()
        hours = val if unit.startswith("h") else val * 24.0

        # Guardrail check
        if not is_admin and hours > (MAX_STANDARD_WINDOW_DAYS * 24.0):
            print(f"⚠️  GUARDRAIL NOTICE: Requested window ({hours:.1f}h / {hours/24:.1f}d) exceeds 7-day standard limit.")
            print(f"   Clamping to maximum standard window: 7 days (168 hours). [Use --admin for historical queries]")
            hours = MAX_STANDARD_WINDOW_DAYS * 24.0

        start_utc = now_utc - timedelta(hours=hours)
        end_utc = now_utc
        return TimeWindow(start_utc, end_utc, is_admin=is_admin, raw_input=q)

    # 2. Bounded Range: 'START to END' / 'between START and END'
    range_split = re.split(r"\s+(?:to|until|through|and)\s+", q, flags=re.I)
    if len(range_split) == 2 and not q.lower().startswith("from") and not q.lower().endswith("onwards"):
        start_raw, end_raw = range_split[0], range_split[1]
        start_utc = parse_datetime_flexible(start_raw)
        end_utc = parse_datetime_flexible(end_raw)

        if start_utc > end_utc:
            start_utc, end_utc = end_utc, start_utc  # auto-correct inverted order

        # Guardrail check
        duration_days = (end_utc - start_utc).total_seconds() / 86400.0
        if not is_admin and duration_days > MAX_STANDARD_WINDOW_DAYS:
            print(f"⚠️  GUARDRAIL NOTICE: Requested range ({duration_days:.1f} days) exceeds 7-day standard limit.")
            print(f"   Clamping end date to 7 days from start date. [Use --admin for historical queries]")
            end_utc = start_utc + timedelta(days=MAX_STANDARD_WINDOW_DAYS)

        return TimeWindow(start_utc, end_utc, is_admin=is_admin, raw_input=q)

    # 3. Open-ended Start: 'From 24 Aug 2026 01:00 AM IST onwards' / 'since 2026-08-24'
    start_utc = parse_datetime_flexible(q)
    end_utc = now_utc

    if start_utc > end_utc:
        # If start is in the future, set window to [now, start]
        start_utc, end_utc = end_utc, start_utc

    duration_days = (end_utc - start_utc).total_seconds() / 86400.0
    if not is_admin and duration_days > MAX_STANDARD_WINDOW_DAYS:
        print(f"⚠️  GUARDRAIL NOTICE: Start date is {duration_days:.1f} days in the past (exceeds 7-day limit).")
        print(f"   Clamping to maximum standard window: 7 days ago. [Use --admin for historical queries]")
        start_utc = end_utc - timedelta(days=MAX_STANDARD_WINDOW_DAYS)

    return TimeWindow(start_utc, end_utc, is_admin=is_admin, raw_input=q)


def test_time_window():
    print("=== Testing Universal TimeWindow Engine ===\n")
    
    # Mock fixed current time for tests: 25 Aug 2026 23:45 IST
    mock_now_utc = datetime(2026, 8, 25, 18, 15, tzinfo=timezone.utc)

    test_inputs = [
        "24h",
        "72h",
        "last 48 hours",
        "7d",
        "14d",  # Should be clamped for non-admin
        "From 24 Aug 2026 01:00 AM IST onwards",
        "24 Aug 2026 01:00 AM IST to 25 Aug 2026 01:00 AM IST",
        "2026-08-20 00:00 to 2026-08-25 18:00",
        "10 Aug 2026 to 25 Aug 2026",  # Exceeds 7 days
    ]

    for inp in test_inputs:
        print(f"Input: \"{inp}\"")
        w = create_time_window(inp, is_admin=False, now_utc=mock_now_utc)
        print(f"  Result: {w.format_summary()}")
        print("-" * 65)

    print("\n=== Testing Admin Override (14 days) ===")
    w_admin = create_time_window("14d", is_admin=True, now_utc=mock_now_utc)
    print(f"  Admin Result: {w_admin.format_summary()}")


if __name__ == "__main__":
    test_time_window()
