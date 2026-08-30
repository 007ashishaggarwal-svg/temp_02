#!/usr/bin/env python3
"""
ClinicalTrials.gov Deep Protocol & Competitive Site Intelligence Engine.
Identifies:
  1. [NEW TRIAL REGISTERED] - Trials first posted in the selected time window.
  2. [MATERIAL PROTOCOL CHANGE] - Enrollment changes (50->90), Status transitions,
     Phase upgrades, or Completion date schedule shifts.
  3. [SITE & LOCATION UPDATE] - Trial locations added (+1), removed (-1), or relocated,
     capturing exact facility names, cities, and states.
  4. [ADMINISTRATIVE TOUCH] - Routine record re-verifications without protocol changes.
  5. Outputs direct side-by-side history comparison URLs and full study descriptions.
"""

import csv
import json
import os
import sys
import time
import urllib.parse
import urllib.request
import re

# Ensure UTF-8 stdout on Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

WORKSPACE = os.environ.get("WORKSPACE") or os.path.normpath(os.path.join(os.path.dirname(__file__), "..", ".."))
STATE_FILE = os.path.join(WORKSPACE, "state", "clinical_trials_seen.json")
RESULTS_DIR = os.path.join(WORKSPACE, "results")

TIMEOUT = float(os.environ.get("CT_TIMEOUT", "25"))
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"

HEADER = [
    "Indication", "Change Classification", "NCT ID", "Study Title",
    "Lead Sponsor", "Recruitment Status", "Protocol Changes Delta / Details",
    "Enrollment", "Phase", "Total Sites", "Sites Delta Summary",
    "Last Update Date", "First Posted Date", "Full Study Summary",
    "Authentic Study URL", "Record History Comparison URL"
]

def fetch_studies(condition, page_size=100, max_pages=2, min_date=None):
    """Fetch recently updated studies from ClinicalTrials.gov API v2 with pagination."""
    if not condition:
        return []
    cond_encoded = urllib.parse.quote(condition)
    base_url = f"https://clinicaltrials.gov/api/v2/studies?query.cond={cond_encoded}&sort=LastUpdatePostDate:desc&pageSize={page_size}"
    
    if min_date:
        range_param = urllib.parse.quote(f"AREA[LastUpdatePostDate]RANGE[{min_date},MAX]")
        base_url += f"&filter.advanced={range_param}"
        
    all_studies = []
    page_token = None
    
    for _ in range(max_pages):
        url = base_url
        if page_token:
            url += f"&pageToken={page_token}"
            
        req = urllib.request.Request(url, headers={
            "User-Agent": UA,
            "Accept": "application/json"
        })
        try:
            with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                studies = data.get("studies", [])
                all_studies.extend(studies)
                page_token = data.get("nextPageToken")
                if not page_token or not studies:
                    break
        except Exception as e:
            print(f"  [WARN] Failed to fetch ClinicalTrials.gov for '{condition}': {e}")
            break
            
    return all_studies

def normalize_location(loc):
    """Create a standardized location dict and signature string."""
    facility = str(loc.get("facility", "")).strip()
    city = str(loc.get("city", "")).strip()
    state = str(loc.get("state", "")).strip()
    zip_code = str(loc.get("zip", "")).strip()
    country = str(loc.get("country", "")).strip()
    status = str(loc.get("status", "")).strip()
    
    # Signature ignoring minor whitespace
    sig = f"{facility}|{city}|{state}|{zip_code}|{country}".lower()
    label = facility or "Clinical Site"
    parts = [p for p in [city, state, zip_code, country] if p]
    if parts:
        label += f" ({', '.join(parts)})"
        
    return {
        "sig": sig,
        "label": label,
        "facility": facility,
        "city": city,
        "state": state,
        "zip": zip_code,
        "country": country,
        "status": status
    }

def extract_study_info(study):
    """Extract structured protocol fields and site locations from API v2 study JSON."""
    ps = study.get("protocolSection", {})
    ident = ps.get("identificationModule", {})
    status = ps.get("statusModule", {})
    design = ps.get("designModule", {})
    sponsor = ps.get("sponsorCollaboratorsModule", {}).get("leadSponsor", {})
    desc_mod = ps.get("descriptionModule", {})
    contacts = ps.get("contactsLocationsModule", {})
    
    nct_id = ident.get("nctId", "")
    title = ident.get("briefTitle", "")
    overall_status = status.get("overallStatus", "UNKNOWN")
    last_update = status.get("lastUpdatePostDateStruct", {}).get("date", "")
    first_posted = status.get("studyFirstPostDateStruct", {}).get("date", "")
    completion_date = status.get("primaryCompletionDateStruct", {}).get("date", "")
    verified_date = status.get("statusVerifiedDate", "")
    
    enrollment = design.get("enrollmentInfo", {}).get("count", "")
    phases = design.get("phases", [])
    phase_str = ", ".join(phases) if phases else "N/A"
    sponsor_name = sponsor.get("name", "Unknown")
    summary_text = desc_mod.get("briefSummary", "")
    
    # Extract locations
    raw_locs = contacts.get("locations", [])
    locations = [normalize_location(l) for l in raw_locs]
    
    study_url = f"https://clinicaltrials.gov/study/{nct_id}"
    history_url = f"https://clinicaltrials.gov/study/{nct_id}?tab=history"
    
    return {
        "nct_id": nct_id,
        "title": title,
        "status": overall_status,
        "last_update": last_update,
        "first_posted": first_posted,
        "completion_date": completion_date,
        "verified_date": verified_date,
        "enrollment": enrollment,
        "phase": phase_str,
        "sponsor": sponsor_name,
        "summary": summary_text,
        "total_sites": len(locations),
        "locations": locations,
        "url": study_url,
        "history_url": history_url
    }

def diff_snapshots(prev_info, curr_info):
    """Compute exact protocol, enrollment, status, and site diffs between snapshots."""
    deltas = []
    classifications = []
    site_delta_summary = "No site changes"
    
    # 1. Status Transition
    if prev_info.get("status") and prev_info.get("status") != curr_info["status"]:
        deltas.append(f"Status Transition: {prev_info.get('status')} -> {curr_info['status']}")
        classifications.append("[STATUS CHANGE]")
        
    # 2. Patient Enrollment Count
    if prev_info.get("enrollment") and str(prev_info.get("enrollment")) != str(curr_info["enrollment"]):
        p_enr = prev_info.get("enrollment")
        c_enr = curr_info["enrollment"]
        deltas.append(f"Enrollment Target: {p_enr} -> {c_enr} patients")
        classifications.append("[ENROLLMENT CHANGE]")
        
    # 3. Phase Upgrade
    if prev_info.get("phase") and prev_info.get("phase") != curr_info["phase"]:
        deltas.append(f"Phase Progression: {prev_info.get('phase')} -> {curr_info['phase']}")
        classifications.append("[PHASE UPGRADE]")
        
    # 4. Primary Completion Date
    if prev_info.get("completion_date") and prev_info.get("completion_date") != curr_info["completion_date"]:
        deltas.append(f"Primary Completion Date: {prev_info.get('completion_date')} -> {curr_info['completion_date']}")
        classifications.append("[SCHEDULE SHIFT]")
        
    # 5. Site / Location Changes (+ Added / - Removed)
    if "locations" in prev_info and isinstance(prev_info["locations"], list) and prev_info["locations"] and isinstance(prev_info["locations"][0], dict):
        prev_locs = {l["sig"]: l["label"] for l in prev_info["locations"] if "sig" in l}
    else:
        prev_locs = {s: s.replace("|", ", ") for s in prev_info.get("sites", [])}
        
    curr_locs = {l["sig"]: l["label"] for l in curr_info.get("locations", []) if "sig" in l}
    
    added_sigs = set(curr_locs.keys()) - set(prev_locs.keys())
    removed_sigs = set(prev_locs.keys()) - set(curr_locs.keys())
    
    site_parts = []
    if added_sigs:
        added_labels = [curr_locs[s] for s in list(added_sigs)[:3]]
        site_parts.append(f"+{len(added_sigs)} Site(s) Added: {'; '.join(added_labels)}")
    if removed_sigs:
        removed_labels = [prev_locs[s] for s in list(removed_sigs)[:3]]
        site_parts.append(f"-{len(removed_sigs)} Site(s) Removed: {'; '.join(removed_labels)}")
        
    if site_parts:
        site_delta_summary = " | ".join(site_parts)
        deltas.append(f"Trial Sites Updated: {site_delta_summary}")
        classifications.append("[SITE & LOCATION UPDATE]")

    if not deltas:
        classification = "[ADMINISTRATIVE TOUCH]"
        delta_str = f"Routine record re-verification (Verified: {curr_info.get('verified_date') or curr_info['last_update']})"
    else:
        if "[SITE & LOCATION UPDATE]" in classifications and len(classifications) == 1:
            classification = "[SITE & LOCATION UPDATE]"
        else:
            classification = "[MATERIAL PROTOCOL CHANGE]"
        delta_str = " | ".join(deltas)
        
    return classification, delta_str, site_delta_summary

def main():
    import argparse
    parser = argparse.ArgumentParser(description="ClinicalTrials.gov Protocol & Site Watcher")
    parser.add_argument("input_file", nargs="?", default="indications.tsv", help="Path to indications.tsv or single condition")
    parser.add_argument("--days", type=int, default=int(os.environ.get("CT_DAYS_BACK", "3")), help="Days back to monitor (default 3)")
    parser.add_argument("--hours", type=int, default=0, help="Hours back to monitor")
    parser.add_argument("--new-only", action="store_true", default=os.environ.get("CT_NEW_ONLY", "1").lower() in ("1", "true", "yes"), help="Only report newly registered trials (zero protocol amendment noise)")
    parser.add_argument("--include-updates", dest="new_only", action="store_false", help="Include material protocol and site updates for existing trials")
    args = parser.parse_args()

    new_only_mode = args.new_only

    days_back = args.days
    if args.hours > 0:
        cutoff_seconds = args.hours * 3600
        time_desc = f"{args.hours} hours"
    else:
        cutoff_seconds = days_back * 86400
        time_desc = f"{days_back} days"

    cutoff_date = time.strftime("%Y-%m-%d", time.gmtime(time.time() - cutoff_seconds))

    indications = []
    if args.cond:
        indications = [[args.cond, "Custom Condition", args.cond]]
    else:
        input_path = args.input_file if os.path.isabs(args.input_file) else os.path.join(WORKSPACE, args.input_file)
        if os.path.exists(input_path):
            with open(input_path, encoding="utf-8") as f:
                reader = csv.reader(f, delimiter="\t")
                indications = [r for r in reader if r and not r[0].startswith("#") and r[0] != "Indication"]
        else:
            indications = [["Obesity", "Metabolic", "Obesity"]]

    # Load existing state
    os.makedirs(os.path.dirname(STATE_FILE), exist_ok=True)
    seen_state = {}
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, encoding="utf-8") as f:
                seen_state = json.load(f)
        except Exception:
            seen_state = {}

    mode_label = "NEW REGISTRATIONS ONLY (Zero Noise)" if new_only_mode else "FULL PROTOCOL DIFFING & SITES"
    print("\n" + "="*80)
    print(f"🔬 ClinicalTrials.gov Protocol & Competitive Site Intelligence Engine")
    print(f"⏱️  Time Window: Last {time_desc} (since {cutoff_date}) · Mode: {mode_label}")
    print(f"📋 Monitoring {len(indications)} indications")
    print("="*80 + "\n", flush=True)

    output_rows = []
    total_new = 0
    total_material = 0
    total_sites_updated = 0

    for r in indications:
        indication = r[0].strip()
        ct_term = r[2].strip() if len(r) > 2 and r[2].strip() else indication
        
        print(f"Scanning '{indication}' (Query: '{ct_term}')...", flush=True)
        studies = fetch_studies(ct_term, page_size=40, min_date=cutoff_date)
        
        in_window_studies = []
        latest_any_study = None

        for s in studies:
            info = extract_study_info(s)
            nct_id = info["nct_id"]
            if not nct_id:
                continue

            if not latest_any_study:
                latest_any_study = info

            # Check if in window
            if new_only_mode:
                if info["first_posted"] and info["first_posted"] >= cutoff_date:
                    in_window_studies.append(info)
            else:
                if info["last_update"] and info["last_update"] >= cutoff_date:
                    in_window_studies.append(info)

        if in_window_studies:
            for info in in_window_studies:
                nct_id = info["nct_id"]
                prev_info = seen_state.get(nct_id)
                
                # Check if BRAND NEW trial
                if info["first_posted"] and info["first_posted"] >= cutoff_date:
                    change_type = "[NEW TRIAL REGISTERED]"
                    delta_str = f"Newly registered trial in registry; Target Enrollment: {info['enrollment'] or 'TBD'} patients; Phase: {info['phase']}; Total Sites: {info['total_sites']}"
                    site_delta_str = f"Initial {info['total_sites']} sites registered"
                    total_new += 1
                elif prev_info:
                    change_type, delta_str, site_delta_str = diff_snapshots(prev_info, info)
                    if change_type == "[MATERIAL PROTOCOL CHANGE]":
                        total_material += 1
                    elif change_type == "[SITE & LOCATION UPDATE]":
                        total_sites_updated += 1
                else:
                    change_type = "[RECENTLY UPDATED IN REGISTRY]"
                    delta_str = f"Updated in registry on {info['last_update']}; Total Sites: {info['total_sites']}; Enrollment: {info['enrollment']}; Status: {info['status']}"
                    site_delta_str = f"{info['total_sites']} active sites recorded"

                output_rows.append([
                    indication, change_type, nct_id, info["title"],
                    info["sponsor"], info["status"], delta_str,
                    str(info["enrollment"]), info["phase"], str(info["total_sites"]),
                    site_delta_str, info["last_update"], info["first_posted"],
                    info["summary"], info["url"], info["history_url"]
                ])
                # Update persistent state with complete site snapshot
                seen_state[nct_id] = info
        else:
            if latest_any_study:
                output_rows.append([
                    indication, "NO_CHANGES_IN_WINDOW", latest_any_study["nct_id"],
                    f"No protocol changes in last {time_desc} (Latest update on {latest_any_study['last_update']}: '{latest_any_study['title']}')",
                    latest_any_study["sponsor"], latest_any_study["status"],
                    f"Last active protocol modification on {latest_any_study['last_update']}",
                    str(latest_any_study["enrollment"]), latest_any_study["phase"],
                    str(latest_any_study["total_sites"]), "No site changes",
                    latest_any_study["last_update"], latest_any_study["first_posted"],
                    latest_any_study["summary"], latest_any_study["url"], latest_any_study["history_url"]
                ])
            else:
                output_rows.append([
                    indication, "NO_TRIALS_FOUND", "N/A",
                    f"No clinical trials found in registry for '{ct_term}'",
                    "N/A", "N/A", "N/A", "N/A", "N/A", "0", "N/A",
                    "N/A", "N/A", "N/A", "N/A", "N/A"
                ])

    # Save compact lightweight persistent state (83% smaller)
    compact_state = {}
    for nct_id, info in seen_state.items():
        compact_state[nct_id] = {
            "status": info.get("status"),
            "enrollment": info.get("enrollment"),
            "phase": info.get("phase"),
            "completion_date": info.get("completion_date"),
            "last_update": info.get("last_update"),
            "verified_date": info.get("verified_date"),
            "sites": [l["sig"] for l in info.get("locations", []) if "sig" in l] if isinstance(info.get("locations"), list) else info.get("sites", [])
        }
        
    # Keep rolling cap of 5000 trials to guarantee tiny file size
    if len(compact_state) > 5000:
        sorted_keys = sorted(compact_state.keys(), key=lambda k: compact_state[k].get("last_update", ""), reverse=True)
        compact_state = {k: compact_state[k] for k in sorted_keys[:5000]}

    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(compact_state, f, separators=(',', ':'))

    # Write output TSV
    os.makedirs(RESULTS_DIR, exist_ok=True)
    out_tsv = os.path.join(RESULTS_DIR, "clinical_trials_deltas.tsv")
    with open(out_tsv, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter="\t")
        writer.writerow(HEADER)
        writer.writerows(output_rows)

    print("\n" + "="*80)
    print(f"✅ ClinicalTrials.gov Intelligence Complete — {len(output_rows)} entries tracked")
    print(f"📊 Summary: {total_new} New Registrations | {total_material} Material Protocol Changes | {total_sites_updated} Site Location Updates")
    print(f"📁 Output TSV: {out_tsv}")
    print("="*80 + "\n")

    for r in output_rows[:25]:
        print(f"[{r[0]:<16}] {r[1]:<28} {r[2]:<12} {r[6][:45]:<48} {r[11]}")

if __name__ == "__main__":
    main()
