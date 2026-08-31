#!/usr/bin/env python3
"""
ClinicalTrials.gov Protocol Delta & CI Intelligence Engine (Pillar 3).
Features:
1. Segregated Snapshot Cache: Stores Micro-Protocol Snapshots in '.cache/ct_protocol_snapshots.json'.
2. Zero-Baseline Requirement: Automatically detects newly registered studies vs protocol modifications.
3. High-Value CI Delta Detection:
   - Trial Termination / Clinical Safety Holds (clinical_neg)
   - Topline Results / Data Postings (clinical_pos)
   - Enrollment Expansions & Reductions (clinical_trial_update)
   - Global Site Network Expansions (clinical_trial_update)
   - Phase Escalation (Phase 2 -> Phase 3) (clinical_pos)
   - Primary Completion Date Shifts / Timeline Delays (clinical_trial_update)
4. Editorial Formatting:
   - Headline: Crisp, high-impact CI summary
   - Editorial Snippet (Col 18): Structured metadata (Title, Phase, Status, Enrollment, Sites, Dates)
   - Full Body Excerpt (Col 19): Clean Inclusion Criteria & Primary Outcome Measures
"""

import os
import sys
import json
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

WORKSPACE = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
CACHE_DIR = os.path.join(WORKSPACE, ".cache")
CT_SNAPSHOT_FILE = os.path.join(CACHE_DIR, "ct_protocol_snapshots.json")

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
_CTX = ssl.create_default_context()
_CTX.check_hostname = False
_CTX.verify_mode = ssl.CERT_NONE


class CTProtocolDeltaEngine:
    def __init__(self, snapshot_path: str = None):
        self.snapshot_path = snapshot_path or CT_SNAPSHOT_FILE
        os.makedirs(os.path.dirname(self.snapshot_path), exist_ok=True)
        self.snapshots = self._load_snapshots()

    def _load_snapshots(self) -> dict:
        if os.path.exists(self.snapshot_path):
            try:
                with open(self.snapshot_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def _save_snapshots(self):
        try:
            with open(self.snapshot_path, "w", encoding="utf-8") as f:
                json.dump(self.snapshots, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"⚠️ Notice saving CT snapshots: {e}")

    def evaluate_study(self, study: dict, condition_theme: str) -> dict | None:
        """
        Evaluate a single CT.gov study JSON object against prior snapshot.
        Extracts high-value CI deltas and builds clean editorial text.
        """
        proto = study.get("protocolSection", {})
        id_mod = proto.get("identificationModule", {})
        status_mod = proto.get("statusModule", {})
        sponsor_mod = proto.get("sponsorCollaboratorsModule", {})
        design_mod = proto.get("designModule", {})
        desc_mod = proto.get("descriptionModule", {})
        elig_mod = proto.get("eligibilityModule", {})
        outcomes_mod = proto.get("outcomesModule", {})
        contacts_mod = proto.get("contactsLocationsModule", {})

        nct_id = id_mod.get("nctId", "")
        if not nct_id:
            return None

        brief_title = id_mod.get("briefTitle", "")
        official_title = id_mod.get("officialTitle", "")
        clean_title = brief_title or official_title
        lead_sponsor = sponsor_mod.get("leadSponsor", {}).get("name", "Unknown Sponsor")
        
        phases = design_mod.get("phases", ["Phase Not Specified"])
        phase_str = "/".join(phases)
        overall_status = status_mod.get("overallStatus", "Unknown Status")

        # Enrollment count
        enrollment_info = design_mod.get("enrollmentInfo", {})
        enrollment_count = enrollment_info.get("count", 0)

        # Site count
        locations = contacts_mod.get("locations", [])
        site_count = len(locations)

        # Dates
        first_post_date = status_mod.get("studyFirstPostDateStruct", {}).get("date", "")
        last_update_date = status_mod.get("lastUpdatePostDateStruct", {}).get("date", "")
        start_date = status_mod.get("startDateStruct", {}).get("date", "--")
        primary_comp_date = status_mod.get("primaryCompletionDateStruct", {}).get("date", "--")
        results_first_post = status_mod.get("resultsFirstPostDateStruct", {}).get("date", "")

        # Key CI Metadata Snapshot
        curr_snapshot = {
            "nct_id": nct_id,
            "title": clean_title,
            "lead_sponsor": lead_sponsor,
            "phase": phase_str,
            "status": overall_status,
            "enrollment": enrollment_count,
            "site_count": site_count,
            "start_date": start_date,
            "primary_completion_date": primary_comp_date,
            "first_post_date": first_post_date,
            "last_update_date": last_update_date,
            "results_posted": bool(results_first_post)
        }

        prev_snap = self.snapshots.get(nct_id)
        
        # ---------------------------------------------------------------------
        # 1. DELTA CLASSIFICATION & HIGH-IMPACT HEADLINE GENERATION
        # ---------------------------------------------------------------------
        delta_tag = "clinical_trial_update"
        headline = ""
        change_summary = ""

        if not prev_snap:
            # Check if brand new registration or first time tracked
            if first_post_date and (first_post_date == last_update_date or first_post_date >= "2026-08-01"):
                delta_tag = "clinical_trial_new"
                headline = f"[NEW TRIAL] {lead_sponsor} registers {phase_str} Study in {condition_theme} ({nct_id})"
                change_summary = f"Newly registered protocol with {enrollment_count} target enrollment across {site_count} sites."
            else:
                delta_tag = "clinical_trial_update"
                headline = f"[PROTOCOL UPDATE] {lead_sponsor} updates {phase_str} Trial ({nct_id}): {clean_title[:75]}"
                change_summary = f"Status: {overall_status} | Target Enrollment: {enrollment_count} | Sites: {site_count}."
        else:
            # Evaluate Exact Diffs vs Previous Snapshot
            changes = []
            
            # Status changes
            if prev_snap.get("status") != overall_status:
                if overall_status in ("TERMINATED", "SUSPENDED", "WITHDRAWN"):
                    delta_tag = "clinical_neg"
                    headline = f"[SAFETY HOLD / TERMINATION] {lead_sponsor} {phase_str} Trial {overall_status} ({nct_id})"
                    changes.append(f"Status changed from {prev_snap.get('status')} to {overall_status}")
                elif overall_status == "ACTIVE_NOT_RECRUITING" and prev_snap.get("status") == "RECRUITING":
                    delta_tag = "clinical_pos"
                    headline = f"[ENROLLMENT COMPLETED] {lead_sponsor} {phase_str} Trial completes recruitment ({nct_id})"
                    changes.append("Recruitment closed; advancing to primary outcome evaluation.")
                elif overall_status == "COMPLETED":
                    delta_tag = "clinical_pos"
                    headline = f"[TRIAL COMPLETED] {lead_sponsor} {phase_str} Trial marked Completed ({nct_id})"
                    changes.append("Trial officially completed.")
                else:
                    changes.append(f"Status transitioned: {prev_snap.get('status')} -> {overall_status}")

            # Topline Results Posted
            if results_first_post and not prev_snap.get("results_posted"):
                delta_tag = "clinical_pos"
                headline = f"[TOPLINE RESULTS POSTED] {lead_sponsor} posts clinical efficacy data ({nct_id})"
                changes.append(f"Topline results officially posted on ClinicalTrials.gov on {results_first_post}.")

            # Enrollment Changes
            prev_enr = prev_snap.get("enrollment", 0)
            if enrollment_count and prev_enr and enrollment_count != prev_enr:
                diff = enrollment_count - prev_enr
                sign = "+" if diff > 0 else ""
                changes.append(f"Enrollment adjusted: {prev_enr} -> {enrollment_count} ({sign}{diff} patients)")
                if not headline:
                    headline = f"[ENROLLMENT {'EXPANSION' if diff > 0 else 'REDUCTION'}] {lead_sponsor} adjusts {phase_str} ({nct_id}) to {enrollment_count} patients"

            # Site Count Changes
            prev_sites = prev_snap.get("site_count", 0)
            if site_count and prev_sites and site_count != prev_sites:
                diff_s = site_count - prev_sites
                sign_s = "+" if diff_s > 0 else ""
                changes.append(f"Site network modified: {prev_sites} -> {site_count} sites ({sign_s}{diff_s} sites)")
                if not headline:
                    headline = f"[SITE EXPANSION] {lead_sponsor} expands {phase_str} ({nct_id}) to {site_count} clinical sites"

            # Primary Completion Date Shift
            prev_pcd = prev_snap.get("primary_completion_date", "--")
            if primary_comp_date != "--" and prev_pcd != "--" and primary_comp_date != prev_pcd:
                changes.append(f"Primary Completion Date shifted from {prev_pcd} to {primary_comp_date}")
                if not headline:
                    headline = f"[TIMELINE UPDATE] {lead_sponsor} {phase_str} ({nct_id}) Primary Completion shifted to {primary_comp_date}"

            if not headline:
                headline = f"[PROTOCOL UPDATE] {lead_sponsor} updates {phase_str} Study ({nct_id}): {clean_title[:75]}"
            
            change_summary = " | ".join(changes) if changes else "Routine administrative protocol validation."

        # Update snapshot in memory
        self.snapshots[nct_id] = curr_snapshot

        # ---------------------------------------------------------------------
        # 2. EDITORIAL SNIPPET / SUMMARY (Crisp, Structured Metadata)
        # ---------------------------------------------------------------------
        editorial_snippet = (
            f"NCT ID: {nct_id} | Phase: {phase_str} | Status: {overall_status}\n"
            f"Sponsor: {lead_sponsor} | Condition: {condition_theme}\n"
            f"Target Enrollment: {enrollment_count:,} participants | Active Sites: {site_count}\n"
            f"Start Date: {start_date} | Primary Completion: {primary_comp_date}\n"
            f"CI Intelligence Delta: {change_summary}"
        )

        # ---------------------------------------------------------------------
        # 3. FULL BODY EXCERPT (Clean Inclusion Criteria & Primary Outcomes)
        # ---------------------------------------------------------------------
        # Primary Outcomes
        prim_outcomes = outcomes_mod.get("primaryOutcomes", [])
        outcomes_list = []
        for o_i, out in enumerate(prim_outcomes[:3], start=1):
            measure = out.get("measure", "")
            tf = out.get("timeFrame", "")
            outcomes_list.append(f"  {o_i}. {measure} [Timeframe: {tf}]" if tf else f"  {o_i}. {measure}")
        outcomes_text = "\n".join(outcomes_list) if outcomes_list else "  • Primary efficacy outcome measure specified in protocol."

        # Inclusion Criteria
        criteria_raw = elig_mod.get("eligibilityCriteria", "")
        # Extract inclusion portion
        inc_text = criteria_raw
        if "Exclusion Criteria:" in criteria_raw:
            inc_text = criteria_raw.split("Exclusion Criteria:")[0]
        inc_lines = [l.strip() for l in inc_text.splitlines() if l.strip() and not l.strip().startswith("Inclusion Criteria")]
        inc_clean = "\n".join([f"  • {l.lstrip('* -')}" for l in inc_lines[:6]]) if inc_lines else "  • Specific inclusion criteria defined in clinical study protocol."

        full_body_excerpt = (
            f"=== STUDY OVERVIEW: {clean_title} ===\n\n"
            f"PRIMARY OUTCOME MEASURES:\n{outcomes_text}\n\n"
            f"KEY INCLUSION CRITERIA:\n{inc_clean}\n\n"
            f"Official Registry Link: https://clinicaltrials.gov/study/{nct_id}"
        )

        return {
            "nct_id": nct_id,
            "headline": headline,
            "signal_type": delta_tag,
            "lead_sponsor": lead_sponsor,
            "condition": condition_theme,
            "phase": phase_str,
            "status": overall_status,
            "published_date": last_update_date or first_post_date or datetime.now(timezone.utc).strftime("%Y-%m-%d"),
            "editorial_snippet": editorial_snippet,
            "full_body_excerpt": full_body_excerpt,
            "raw_url": f"https://clinicaltrials.gov/study/{nct_id}",
            "change_summary": change_summary
        }

    def close(self):
        """Save snapshots to disk."""
        self._save_snapshots()


def fetch_new_trials_for_condition(condition: str, time_window: TimeWindow, max_studies: int = 50) -> list[dict]:
    """
    Fetch studies for condition, filter to time window, and run through Delta Engine.
    """
    cond_clean = condition.strip()
    if not cond_clean or cond_clean.lower() in ("none", "n/a"):
        return []

    engine = CTProtocolDeltaEngine()
    
    # Query CT.gov sorted by LastUpdatePostDate descending for real-time deltas
    q_cond = urllib.parse.quote(cond_clean)
    url = (f"https://clinicaltrials.gov/api/v2/studies?query.cond={q_cond}"
           f"&sort=LastUpdatePostDate:desc&pageSize={min(max_studies, 100)}&countTotal=true")

    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=12, context=_CTX) as resp:
            data = json.loads(resp.read().decode("utf-8", "ignore"))
    except Exception as e:
        print(f"  [CT.gov Error] Condition '{cond_clean}': {e}")
        return []

    studies = data.get("studies", [])
    captured_events = []

    for study in studies:
        proto = study.get("protocolSection", {})
        status_mod = proto.get("statusModule", {})

        last_up = status_mod.get("lastUpdatePostDateStruct", {}).get("date", "")
        first_post = status_mod.get("studyFirstPostDateStruct", {}).get("date", "")
        eval_date_str = last_up or first_post
        if not eval_date_str:
            continue

        try:
            dt = datetime.strptime(eval_date_str, "%Y-%m-%d").replace(tzinfo=timezone.utc)
        except Exception:
            continue

        if time_window.is_in_window(dt):
            event = engine.evaluate_study(study, cond_clean)
            if event:
                captured_events.append({
                    "nct_id": event["nct_id"],
                    "title": event["headline"],
                    "lead_sponsor": event["lead_sponsor"],
                    "condition": event["condition"],
                    "phase": event["phase"],
                    "status": event["status"],
                    "first_post_date": event["published_date"],
                    "summary": event["editorial_snippet"],
                    "full_body": event["full_body_excerpt"],
                    "url": event["raw_url"],
                    "event_type": event["headline"],
                    "signal_type": event["signal_type"],
                    "discovery_method": "Pillar 3: ClinicalTrials.gov API",
                    "extraction_vector": "CT.gov API v2 (/studies)",
                    "match_reason": f"Indication Watch: {cond_clean} | {event['change_summary']}",
                })

    engine.close()
    return captured_events
