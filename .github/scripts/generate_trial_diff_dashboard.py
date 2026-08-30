#!/usr/bin/env python3
"""
Generates an interactive, visual ClinicalTrials.gov Protocol & Site Diff Dashboard (HTML).
Renders:
  - 🟢 Green highlighted cards for added sites & expanded enrollment targets.
  - 🔴 Pink strikethrough cards for removed investigator sites & discontinued locations.
  - 🟣 Protocol metric chips for Phase progression & Status transitions.
  - Direct 1-click launch buttons to ClinicalTrials.gov Side-by-Side Record History.
"""

import csv
import html
import os
import sys
import time

WORKSPACE = r"c:\Users\User\Desktop\App\All_Apps\RSSFeedChecker"
TSV_PATH = os.path.join(WORKSPACE, "results", "clinical_trials_deltas.tsv")
OUTPUT_HTML = os.path.join(WORKSPACE, "results", "clinical_trials_diff_dashboard.html")

def generate_dashboard():
    if not os.path.exists(TSV_PATH):
        print(f"[ERR] File not found: {TSV_PATH}")
        return

    with open(TSV_PATH, encoding="utf-8") as f:
        reader = list(csv.reader(f, delimiter="\t"))

    if not reader:
        print("[WARN] Empty TSV file.")
        return

    header = reader[0]
    rows = reader[1:]

    # Categorize rows
    new_trials = [r for r in rows if r[1] == "[NEW TRIAL REGISTERED]"]
    site_updates = [r for r in rows if r[1] == "[SITE & LOCATION UPDATE]"]
    material_changes = [r for r in rows if r[1] == "[MATERIAL PROTOCOL CHANGE]"]
    admin_touches = [r for r in rows if r[1] == "[ADMINISTRATIVE TOUCH]"]

    html_cards = []

    for r in rows:
        if len(r) < 16:
            continue
        indication = html.escape(r[0])
        change_type = html.escape(r[1])
        nct_id = html.escape(r[2])
        title = html.escape(r[3])
        sponsor = html.escape(r[4])
        status = html.escape(r[5])
        details = html.escape(r[6])
        enrollment = html.escape(r[7])
        phase = html.escape(r[8])
        total_sites = html.escape(r[9])
        sites_delta = html.escape(r[10])
        last_update = html.escape(r[11])
        first_posted = html.escape(r[12])
        summary = html.escape(r[13])
        study_url = html.escape(r[14])
        history_url = html.escape(r[15])

        # Change badge style
        if change_type == "[NEW TRIAL REGISTERED]":
            badge_class = "badge-new"
            badge_icon = "✨"
        elif change_type == "[SITE & LOCATION UPDATE]":
            badge_class = "badge-site"
            badge_icon = "📍"
        elif change_type == "[MATERIAL PROTOCOL CHANGE]":
            badge_class = "badge-material"
            badge_icon = "⚡"
        elif change_type == "[ADMINISTRATIVE TOUCH]":
            badge_class = "badge-admin"
            badge_icon = "📋"
        else:
            badge_class = "badge-none"
            badge_icon = "ℹ️"

        # Format visual site diffs
        site_diff_html = ""
        if "+" in sites_delta or "-" in sites_delta:
            parts = sites_delta.split(" | ")
            diff_chips = []
            for p in parts:
                if p.startswith("+"):
                    diff_chips.append(f'<div class="diff-add"><strong>+ Added:</strong> {p[1:].strip()}</div>')
                elif p.startswith("-"):
                    diff_chips.append(f'<div class="diff-remove"><strong>− Removed:</strong> {p[1:].strip()}</div>')
                else:
                    diff_chips.append(f'<div class="diff-neutral">{p}</div>')
            site_diff_html = "".join(diff_chips)
        else:
            site_diff_html = f'<div class="diff-neutral">{sites_delta}</div>'

        card = f"""
        <div class="trial-card" data-change="{change_type}">
            <div class="card-header">
                <div class="header-left">
                    <span class="badge {badge_class}">{badge_icon} {change_type}</span>
                    <span class="indication-tag">{indication}</span>
                    <span class="nct-badge"><a href="{study_url}" target="_blank">{nct_id}</a></span>
                </div>
                <div class="header-right">
                    <span class="date-chip">📅 Updated: {last_update}</span>
                </div>
            </div>
            
            <h3 class="trial-title">{title}</h3>
            
            <div class="meta-row">
                <div class="meta-item"><strong>Sponsor:</strong> {sponsor}</div>
                <div class="meta-item"><strong>Status:</strong> <span class="status-pill">{status}</span></div>
                <div class="meta-item"><strong>Phase:</strong> {phase}</div>
                <div class="meta-item"><strong>Enrollment:</strong> {enrollment} patients</div>
                <div class="meta-item"><strong>Active Sites:</strong> {total_sites}</div>
            </div>

            <div class="diff-container">
                <div class="diff-title">🔍 Protocol & Site Intelligence Delta:</div>
                <div class="diff-body">
                    <div class="detail-text">{details}</div>
                    {site_diff_html}
                </div>
            </div>

            <div class="card-footer">
                <a href="{history_url}" target="_blank" class="btn btn-compare">
                    <span>⚡ Compare Versions Side-by-Side on ClinicalTrials.gov</span>
                </a>
                <a href="{study_url}" target="_blank" class="btn btn-view">
                    <span>📖 Full Study Record</span>
                </a>
            </div>
        </div>
        """
        html_cards.append(card)

    cards_joined = "\n".join(html_cards)

    full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ClinicalTrials.gov Protocol & Competitive Site Intelligence Dashboard</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {{
            --bg: #0d1117;
            --surface: #161b22;
            --border: #30363d;
            --text-main: #f0f6fc;
            --text-muted: #8b949e;
            --accent-blue: #58a6ff;
            --accent-green: #3fb950;
            --accent-pink: #f85149;
            --accent-purple: #bc8cff;
            --accent-amber: #d29922;
            --green-bg: rgba(63, 185, 80, 0.15);
            --pink-bg: rgba(248, 81, 73, 0.15);
        }}
        * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif; }}
        body {{ background-color: var(--bg); color: var(--text-main); padding: 32px 24px; }}
        .container {{ max-width: 1200px; margin: 0 auto; }}
        
        header {{ margin-bottom: 28px; padding-bottom: 20px; border-bottom: 1px solid var(--border); }}
        h1 {{ font-size: 26px; font-weight: 700; color: #ffffff; margin-bottom: 8px; display: flex; align-items: center; gap: 12px; }}
        p.subtitle {{ color: var(--text-muted); font-size: 14px; }}
        
        .stats-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 28px; }}
        .stat-card {{ background: var(--surface); border: 1px solid var(--border); border-radius: 8px; padding: 16px; }}
        .stat-val {{ font-size: 24px; font-weight: 700; margin-bottom: 4px; }}
        .stat-label {{ font-size: 12px; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; }}
        
        .trial-card {{ background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 20px; margin-bottom: 16px; transition: transform 0.15s ease, border-color 0.15s ease; }}
        .trial-card:hover {{ border-color: var(--accent-blue); transform: translateY(-2px); }}
        
        .card-header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; flex-wrap: wrap; gap: 8px; }}
        .header-left {{ display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }}
        
        .badge {{ font-size: 11px; font-weight: 600; padding: 4px 10px; border-radius: 20px; text-transform: uppercase; }}
        .badge-new {{ background: rgba(63, 185, 80, 0.2); color: #56d364; border: 1px solid rgba(63, 185, 80, 0.4); }}
        .badge-site {{ background: rgba(188, 140, 255, 0.2); color: #d2a8ff; border: 1px solid rgba(188, 140, 255, 0.4); }}
        .badge-material {{ background: rgba(210, 153, 34, 0.2); color: #e3b341; border: 1px solid rgba(210, 153, 34, 0.4); }}
        .badge-admin {{ background: rgba(139, 148, 158, 0.2); color: #8b949e; border: 1px solid rgba(139, 148, 158, 0.4); }}
        .badge-none {{ background: rgba(88, 166, 255, 0.1); color: #58a6ff; }}
        
        .indication-tag {{ background: #21262d; color: #c9d1d9; font-size: 11px; font-weight: 600; padding: 4px 8px; border-radius: 6px; border: 1px solid var(--border); }}
        .nct-badge a {{ color: var(--accent-blue); font-weight: 700; text-decoration: none; font-size: 13px; font-family: monospace; }}
        .nct-badge a:hover {{ text-decoration: underline; }}
        .date-chip {{ font-size: 12px; color: var(--text-muted); }}
        
        .trial-title {{ font-size: 16px; font-weight: 600; color: #ffffff; line-height: 1.4; margin-bottom: 12px; }}
        
        .meta-row {{ display: flex; flex-wrap: wrap; gap: 16px; font-size: 13px; color: #c9d1d9; margin-bottom: 14px; background: #0d1117; padding: 10px 14px; border-radius: 6px; border: 1px solid #21262d; }}
        .status-pill {{ background: #1f6feb22; color: #79c0ff; padding: 2px 6px; border-radius: 4px; font-weight: 500; }}
        
        .diff-container {{ background: #0d1117; border-left: 3px solid var(--accent-purple); border-radius: 0 6px 6px 0; padding: 12px 14px; margin-bottom: 16px; }}
        .diff-title {{ font-size: 12px; font-weight: 600; color: var(--text-muted); margin-bottom: 6px; text-transform: uppercase; }}
        .detail-text {{ font-size: 13px; color: #f0f6fc; margin-bottom: 8px; font-weight: 500; }}
        
        .diff-add {{ background: var(--green-bg); color: #7ee787; padding: 6px 10px; border-radius: 4px; margin-bottom: 4px; font-size: 12px; border-left: 2px solid var(--accent-green); }}
        .diff-remove {{ background: var(--pink-bg); color: #ff7b72; padding: 6px 10px; border-radius: 4px; margin-bottom: 4px; font-size: 12px; text-decoration: line-through; border-left: 2px solid var(--accent-pink); }}
        .diff-neutral {{ font-size: 12px; color: var(--text-muted); padding: 4px 0; }}
        
        .card-footer {{ display: flex; gap: 10px; flex-wrap: wrap; }}
        .btn {{ display: inline-flex; align-items: center; gap: 6px; padding: 7px 14px; border-radius: 6px; font-size: 12px; font-weight: 600; text-decoration: none; transition: all 0.15s ease; }}
        .btn-compare {{ background: #238636; color: #ffffff; }}
        .btn-compare:hover {{ background: #2ea043; }}
        .btn-view {{ background: #21262d; color: #c9d1d9; border: 1px solid var(--border); }}
        .btn-view:hover {{ background: #30363d; color: #ffffff; }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>🔬 ClinicalTrials.gov Protocol & Site Intelligence Dashboard</h1>
            <p class="subtitle">Generated on {time.strftime("%Y-%m-%d %H:%M:%S UTC", time.gmtime())} · Automated Delta & Site Tracking Engine</p>
        </header>

        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-val" style="color: #56d364;">{len(new_trials)}</div>
                <div class="stat-label">New Trials Registered</div>
            </div>
            <div class="stat-card">
                <div class="stat-val" style="color: #d2a8ff;">{len(site_updates)}</div>
                <div class="stat-label">Site & Location Updates</div>
            </div>
            <div class="stat-card">
                <div class="stat-val" style="color: #e3b341;">{len(material_changes)}</div>
                <div class="stat-label">Material Protocol Changes</div>
            </div>
            <div class="stat-card">
                <div class="stat-val" style="color: #8b949e;">{len(admin_touches)}</div>
                <div class="stat-label">Administrative Touches</div>
            </div>
        </div>

        <div class="trials-list">
            {cards_joined}
        </div>
    </div>
</body>
</html>
"""

    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(full_html)

    print(f"\n[OK] Interactive Visual Dashboard created at: {OUTPUT_HTML}")

if __name__ == "__main__":
    generate_dashboard()
