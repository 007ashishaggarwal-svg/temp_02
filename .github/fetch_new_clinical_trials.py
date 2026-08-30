#!/usr/bin/env python3
"""
ClinicalTrials.gov New Trial Registration Fetcher (Simplified Pillar 3).
Focuses strictly on absorbing [NEWLY REGISTERED TRIALS] within the configured time window.
(Complex protocol delta / version diffing is segregated for separate dedicated reporting).
"""

import os
import sys
import json
import time
import urllib.request
import urllib.parse
import ssl
from datetime import datetime, timezone

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from time_window import create_time_window, TimeWindow

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
_CTX = ssl.create_default_context()
_CTX.check_hostname = False
_CTX.verify_mode = ssl.CERT_NONE


def fetch_new_trials_for_condition(condition: str, time_window: TimeWindow, max_studies: int = 50) -> list[dict]:
    """
    Query ClinicalTrials.gov API v2 for newly registered studies for a condition/theme,
    filtering strictly to studies whose first submission / first post date falls inside time_window.
    """
    cond_clean = condition.strip()
    if not cond_clean or cond_clean.lower() in ("none", "n/a"):
        return []

    # Query CT.gov sorted by first post date descending
    q_cond = urllib.parse.quote(cond_clean)
    url = (f"https://clinicaltrials.gov/api/v2/studies?query.cond={q_cond}"
           f"&sort=StudyFirstPostDate:desc&pageSize={min(max_studies, 100)}&countTotal=true")

    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=12, context=_CTX) as resp:
            data = json.loads(resp.read().decode("utf-8", "ignore"))
    except Exception as e:
        print(f"  [CT.gov Error] Condition '{cond_clean}': {e}")
        return []

    studies = data.get("studies", [])
    new_trials = []

    for study in studies:
        proto = study.get("protocolSection", {})
        id_mod = proto.get("identificationModule", {})
        status_mod = proto.get("statusModule", {})
        sponsor_mod = proto.get("sponsorCollaboratorsModule", {})
        design_mod = proto.get("designModule", {})
        desc_mod = proto.get("descriptionModule", {})

        nct_id = id_mod.get("nctId", "")
        brief_title = id_mod.get("briefTitle", "")
        official_title = id_mod.get("officialTitle", "")
        lead_sponsor = sponsor_mod.get("leadSponsor", {}).get("name", "Unknown Sponsor")
        phases = design_mod.get("phases", ["Phase Not Specified"])
        phase_str = "/".join(phases)
        overall_status = status_mod.get("overallStatus", "Unknown Status")
        brief_summary = desc_mod.get("briefSummary", "")

        # Dates
        first_post_raw = status_mod.get("studyFirstPostDateStruct", {}).get("date", "") # YYYY-MM-DD
        first_submit_raw = status_mod.get("studyFirstSubmitDate", "")                   # YYYY-MM-DD

        date_str = first_post_raw or first_submit_raw
        if not date_str:
            continue

        try:
            # Parse date YYYY-MM-DD as UTC
            dt = datetime.strptime(date_str, "%Y-%m-%d").replace(tzinfo=timezone.utc)
        except Exception:
            continue

        # Evaluate against TimeWindow
        if time_window.is_in_window(dt):
            new_trials.append({
                "nct_id": nct_id,
                "title": brief_title or official_title,
                "lead_sponsor": lead_sponsor,
                "condition": cond_clean,
                "phase": phase_str,
                "status": overall_status,
                "first_post_date": date_str,
                "summary": brief_summary[:300],
                "url": f"https://clinicaltrials.gov/study/{nct_id}",
                "event_type": "[NEW TRIAL REGISTERED]",
                "discovery_method": "Pillar 3: ClinicalTrials.gov API",
                "extraction_vector": "CT.gov API v2 (/studies)",
                "match_reason": f"Indication Watch: {cond_clean}",
            })

    return new_trials


def run_new_trials_scan(window_query="72h", is_admin=False):
    tw = create_time_window(window_query, is_admin=is_admin)
    print(f"=== ClinicalTrials.gov New Registration Scanner ===")
    print(tw.format_summary())
    print()

    # High-priority indication list from Tab 05
    indications = [
        "Obesity", "Rheumatoid Arthritis", "Breast Cancer", "NSCLC",
        "Multiple Myeloma", "Alzheimers", "Parkinson", "Schizophrenia",
        "Inflammatory Bowel Disease", "Lupus", "Small Cell Lung Cancer",
        "Macular Degeneration", "MASH"
    ]

    total_discovered = 0
    all_new_trials = []

    for ind in indications:
        trials = fetch_new_trials_for_condition(ind, tw)
        total_discovered += len(trials)
        all_new_trials.extend(trials)
        if trials:
            print(f"  ✨ Found {len(trials)} NEW registered trials for '{ind}':")
            for t in trials:
                print(f"     • {t['nct_id']} [{t['phase']}] ({t['first_post_date']}): {t['title'][:90]}")
                print(f"       Sponsor: {t['lead_sponsor']} | Link: {t['url']}")
        else:
            print(f"  • '{ind}': 0 new trials registered in window.")

    print(f"\nTotal newly registered trials discovered: {total_discovered}")
    return all_new_trials


if __name__ == "__main__":
    w_arg = sys.argv[1] if len(sys.argv) > 1 else "72h"
    admin_arg = "--admin" in sys.argv
    run_new_trials_scan(w_arg, is_admin=admin_arg)
