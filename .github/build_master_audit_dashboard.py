#!/usr/bin/env python3
"""
Compiles all 10 Deep Audit artifacts into a single, unified, interactive, beautifully styled HTML Dashboard.
Uses Python markdown library with tag sanitization and robust DOM event handling.
File output: RSSFeedChecker_Master_Audit_Report.html
"""

import os
import sys
import re
import markdown

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

WORKSPACE = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", ".."))
HTML_OUT = os.path.join(WORKSPACE, "RSSFeedChecker_Master_Audit_Report.html")

def read_file(name):
    p = os.path.join(WORKSPACE, name)
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return f.read()
    return f"*(File {name} not found)*"

def sanitize_and_convert(md_text: str) -> str:
    """Sanitize raw executable HTML tags and convert Markdown to clean HTML."""
    t = md_text
    # Escape script/style/iframe tags so browser does not execute them
    t = re.sub(r"<(script|style|noscript|iframe|object|embed)\b", r"&lt;\1", t, flags=re.I)
    t = re.sub(r"</(script|style|noscript|iframe|object|embed)>", r"&lt;/\1&gt;", t, flags=re.I)
    # Convert absolute Windows file paths to clean relative links for local browsing
    t = re.sub(r"file:///[Cc]:/Users/User/Desktop/App/All_Apps/RSSFeedChecker/", "", t)
    t = re.sub(r"file:///c:/Users/User/Desktop/App/All_Apps/RSSFeedChecker/", "", t)
    
    # Convert using python markdown with tables & code fences
    h = markdown.markdown(t, extensions=['tables', 'fenced_code'])
    
    # Beautify status badges in tables
    h = re.sub(r"\b(PASS WITH FINDINGS)\b", r'<span class="status-warn">\1</span>', h)
    h = re.sub(r"\b(PASS)\b", r'<span class="status-pass">\1</span>', h)
    h = re.sub(r"\b(REWORK REQUIRED|REJECTED|BLOCKED|CRITICAL|HIGH|RED)\b", r'<span class="status-risk">\1</span>', h)
    return h

# Load all 10 documents
doc_morning = read_file("Morning_Decision_Sheet.md")
doc_deep_audit = read_file("Deep_Audit_Report.md")
doc_feed_audit = read_file("Unified_Feed_Audit.md")
doc_findings = read_file("Findings_Register.md")
doc_checklist = read_file("Production_Readiness_Checklist.md")
doc_arch = read_file("Target_Architecture_Design.md")
doc_regression = read_file("Regression_Test_Catalogue.md")
doc_evidence = read_file("Evidence_Ledger.md")
doc_batch_gate = read_file("Batch_Gate_Table.md")
doc_adversarial = read_file("Final_Adversarial_Review.md")

html_morning = sanitize_and_convert(doc_morning)
html_deep_audit = sanitize_and_convert(doc_deep_audit)
html_feed_audit = sanitize_and_convert(doc_feed_audit)
html_findings = sanitize_and_convert(doc_findings)
html_checklist = sanitize_and_convert(doc_checklist)
html_arch = sanitize_and_convert(doc_arch)
html_regression = sanitize_and_convert(doc_regression)
html_evidence = sanitize_and_convert(doc_evidence)
html_batch_gate = sanitize_and_convert(doc_batch_gate)
html_adversarial = sanitize_and_convert(doc_adversarial)

html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>RSSFeedChecker — Master Autonomous Deep Audit & Executive Dossier</title>
    <style>
        :root {{
            --primary: #1B365D;
            --primary-light: #2C5282;
            --accent: #0066CC;
            --bg: #F4F6F9;
            --card-bg: #FFFFFF;
            --text-main: #1A202C;
            --text-muted: #718096;
            --border: #E2E8F0;
            --success-bg: #DEF7EC;
            --success-text: #03543F;
            --warning-bg: #FEF08A;
            --warning-text: #713F12;
            --danger-bg: #FDE8E8;
            --danger-text: #9B1C1C;
        }}

        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            background-color: var(--bg);
            color: var(--text-main);
            display: flex;
            height: 100vh;
            overflow: hidden;
        }}

        /* Sidebar Navigation */
        .sidebar {{
            width: 320px;
            background-color: var(--primary);
            color: #FFFFFF;
            display: flex;
            flex-direction: column;
            flex-shrink: 0;
            box-shadow: 2px 0 10px rgba(0,0,0,0.15);
            z-index: 10;
        }}

        .sidebar-header {{
            padding: 24px 20px;
            border-bottom: 1px solid rgba(255,255,255,0.12);
            background: linear-gradient(180deg, #1B365D 0%, #15294A 100%);
        }}

        .sidebar-header h1 {{
            font-size: 18px;
            font-weight: 700;
            letter-spacing: -0.5px;
            display: flex;
            align-items: center;
            gap: 8px;
            color: #FFFFFF;
        }}

        .sidebar-header p {{
            font-size: 12px;
            color: #CBD5E0;
            margin-top: 6px;
            line-height: 1.4;
        }}

        .nav-list {{
            list-style: none;
            overflow-y: auto;
            padding: 12px 10px;
            flex: 1;
        }}

        .nav-btn {{
            width: 100%;
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 12px 14px;
            margin-bottom: 5px;
            border-radius: 6px;
            border: none;
            background: transparent;
            cursor: pointer;
            font-size: 13px;
            font-weight: 500;
            color: #E2E8F0;
            text-align: left;
            transition: all 0.15s ease;
            outline: none;
        }}

        .nav-btn:hover {{
            background-color: rgba(255,255,255,0.12);
            color: #FFFFFF;
        }}

        .nav-btn.active {{
            background-color: var(--accent);
            color: #FFFFFF;
            font-weight: 600;
            box-shadow: 0 2px 6px rgba(0, 102, 204, 0.4);
        }}

        .sidebar-footer {{
            padding: 14px 20px;
            border-top: 1px solid rgba(255,255,255,0.12);
            font-size: 11px;
            color: #A0AEC0;
            text-align: center;
            background: #15294A;
        }}

        /* Main Content Viewport */
        .main-viewport {{
            flex: 1;
            overflow-y: auto;
            padding: 32px 40px;
            scroll-behavior: smooth;
        }}

        .content-container {{
            max-width: 1150px;
            margin: 0 auto;
        }}

        .tab-pane {{
            display: none;
        }}

        .tab-pane.active {{
            display: block;
            animation: fadeIn 0.2s ease-in-out;
        }}

        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(4px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}

        /* Top Executive Banner */
        .exec-banner {{
            background: linear-gradient(135deg, #1B365D 0%, #2C5282 100%);
            color: white;
            padding: 24px 28px;
            border-radius: 12px;
            margin-bottom: 24px;
            box-shadow: 0 4px 14px rgba(27, 54, 93, 0.15);
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 16px;
        }}

        .exec-banner h2 {{
            font-size: 20px;
            margin-bottom: 4px;
            color: #FFFFFF;
        }}

        .exec-banner p {{
            font-size: 13px;
            color: #E2E8F0;
        }}

        .banner-badges {{
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
        }}

        .badge-pill {{
            background: rgba(255,255,255,0.15);
            padding: 6px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            border: 1px solid rgba(255,255,255,0.25);
            color: #FFFFFF;
        }}

        /* Card Container */
        .card {{
            background: var(--card-bg);
            border-radius: 10px;
            border: 1px solid var(--border);
            padding: 30px;
            margin-bottom: 24px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.03);
        }}

        /* Markdown Typography */
        h1 {{ font-size: 22px; margin-bottom: 16px; color: var(--primary); font-weight: 700; }}
        h2 {{ font-size: 17px; margin: 24px 0 12px 0; color: var(--primary-light); border-bottom: 2px solid var(--bg); padding-bottom: 6px; font-weight: 700; }}
        h3 {{ font-size: 14.5px; margin: 16px 0 8px 0; color: var(--text-main); font-weight: 600; }}
        p {{ font-size: 13.5px; line-height: 1.6; margin-bottom: 12px; color: #2D3748; }}
        ul, ol {{ margin-left: 24px; margin-bottom: 14px; font-size: 13.5px; line-height: 1.6; color: #2D3748; }}
        li {{ margin-bottom: 4px; }}
        strong {{ color: #1A202C; }}

        blockquote {{
            border-left: 4px solid var(--accent);
            background: #EBF8FF;
            padding: 12px 16px;
            margin: 14px 0;
            border-radius: 0 6px 6px 0;
            font-size: 13.5px;
            color: #2B6CB0;
        }}

        /* Tables */
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 16px 0 24px 0;
            font-size: 12.5px;
            background: white;
            border-radius: 8px;
            overflow: hidden;
            border: 1px solid var(--border);
        }}

        th {{
            background-color: #1B365D;
            color: #FFFFFF;
            font-weight: 600;
            text-align: left;
            padding: 10px 12px;
            font-size: 12px;
        }}

        td {{
            padding: 9px 12px;
            border-bottom: 1px solid var(--border);
            color: #2D3748;
            vertical-align: top;
        }}

        tr:nth-child(even) td {{
            background-color: #F8FAFC;
        }}

        tr:hover td {{
            background-color: #EDF2F7;
        }}

        /* Code Blocks & Badges */
        pre {{
            background: #1A202C;
            color: #E2E8F0;
            padding: 14px 16px;
            border-radius: 8px;
            font-family: "Consolas", "Courier New", monospace;
            font-size: 12px;
            overflow-x: auto;
            margin: 14px 0;
            line-height: 1.5;
        }}

        code {{
            font-family: "Consolas", "Courier New", monospace;
            background: #EDF2F7;
            color: #C53030;
            padding: 2px 5px;
            border-radius: 4px;
            font-size: 12px;
        }}

        pre code {{
            background: transparent;
            color: #E2E8F0;
            padding: 0;
        }}

        hr {{
            border: 0;
            height: 1px;
            background: var(--border);
            margin: 24px 0;
        }}

        a {{
            color: var(--accent);
            text-decoration: none;
            font-weight: 500;
        }}

        a:hover {{
            text-decoration: underline;
        }}

        /* Status Pills */
        .status-pass {{ background: var(--success-bg); color: var(--success-text); padding: 3px 8px; border-radius: 12px; font-weight: 600; font-size: 11px; display: inline-block; }}
        .status-warn {{ background: var(--warning-bg); color: var(--warning-text); padding: 3px 8px; border-radius: 12px; font-weight: 600; font-size: 11px; display: inline-block; }}
        .status-risk {{ background: var(--danger-bg); color: var(--danger-text); padding: 3px 8px; border-radius: 12px; font-weight: 600; font-size: 11px; display: inline-block; }}

        /* Print optimization */
        @media print {{
            .sidebar {{ display: none; }}
            .main-viewport {{ padding: 0; overflow: visible; }}
            .tab-pane {{ display: block !important; page-break-after: always; }}
        }}
    </style>
</head>
<body>

    <!-- Left Navigation Sidebar -->
    <div class="sidebar">
        <div class="sidebar-header">
            <h1>🔬 RSSFeedChecker</h1>
            <p>Master Autonomous Deep Audit &amp; Intelligence Dossier</p>
        </div>
        <div class="nav-list" id="sidebar-nav">
            <button class="nav-btn active" data-target="tab-morning">☕ 1. Morning Decision Brief</button>
            <button class="nav-btn" data-target="tab-deep-audit">📑 2. Comprehensive Audit Report</button>
            <button class="nav-btn" data-target="tab-feed-audit">🔬 3. Unified Feed Forensic Audit</button>
            <button class="nav-btn" data-target="tab-findings">🚨 4. Findings &amp; Risk Register</button>
            <button class="nav-btn" data-target="tab-checklist">🚦 5. Production Readiness Scorecard</button>
            <button class="nav-btn" data-target="tab-arch">🏛️ 6. Target Architecture (A/B/C)</button>
            <button class="nav-btn" data-target="tab-regression">🧪 7. Regression Test Catalogue</button>
            <button class="nav-btn" data-target="tab-evidence">📊 8. Evidence &amp; Verification Ledger</button>
            <button class="nav-btn" data-target="tab-batch-gate">🚪 9. Batch-by-Batch Gate Table</button>
            <button class="nav-btn" data-target="tab-adversarial">🛡️ 10. Adversarial Red-Team Review</button>
        </div>
        <div class="sidebar-footer">
            <span>Verified against 2,542 Rows across 8 Excel Sheets<br>&bull; Master Guide &amp; Data v2.0 &bull;</span>
        </div>
    </div>

    <!-- Main Viewport -->
    <div class="main-viewport" id="main-viewport">
        <div class="content-container">

            <!-- Executive Quick Banner -->
            <div class="exec-banner">
                <div>
                    <h2>Biopharma Intelligence Suite — E2E Forensic Audit</h2>
                    <p>Independent multi-role assessment across 459 Feeds, 616 Companies, 18 Indications, and ClinicalTrials.gov API v2</p>
                </div>
                <div class="banner-badges">
                    <span class="badge-pill">📊 1,349 Live Unique Items</span>
                    <span class="badge-pill">🛡️ 0 JS / HTML Leaks</span>
                    <span class="badge-pill">🎯 6.1% High-Priority Signal</span>
                    <span class="badge-pill">⏱️ ~4.5m Full Catalog Run</span>
                </div>
            </div>

            <!-- TAB 1: MORNING DECISION BRIEF -->
            <div id="tab-morning" class="tab-pane active">
                <div class="card">
                    {html_morning}
                </div>
            </div>

            <!-- TAB 2: COMPREHENSIVE AUDIT REPORT -->
            <div id="tab-deep-audit" class="tab-pane">
                <div class="card">
                    {html_deep_audit}
                </div>
            </div>

            <!-- TAB 3: UNIFIED FEED FORENSIC AUDIT -->
            <div id="tab-feed-audit" class="tab-pane">
                <div class="card">
                    {html_feed_audit}
                </div>
            </div>

            <!-- TAB 4: FINDINGS & RISK REGISTER -->
            <div id="tab-findings" class="tab-pane">
                <div class="card">
                    {html_findings}
                </div>
            </div>

            <!-- TAB 5: PRODUCTION READINESS SCORECARD -->
            <div id="tab-checklist" class="tab-pane">
                <div class="card">
                    {html_checklist}
                </div>
            </div>

            <!-- TAB 6: TARGET ARCHITECTURE (A/B/C) -->
            <div id="tab-arch" class="tab-pane">
                <div class="card">
                    {html_arch}
                </div>
            </div>

            <!-- TAB 7: REGRESSION TEST CATALOGUE -->
            <div id="tab-regression" class="tab-pane">
                <div class="card">
                    {html_regression}
                </div>
            </div>

            <!-- TAB 8: EVIDENCE & VERIFICATION LEDGER -->
            <div id="tab-evidence" class="tab-pane">
                <div class="card">
                    {html_evidence}
                </div>
            </div>

            <!-- TAB 9: BATCH GATE TABLE -->
            <div id="tab-batch-gate" class="tab-pane">
                <div class="card">
                    {html_batch_gate}
                </div>
            </div>

            <!-- TAB 10: ADVERSARIAL RED-TEAM REVIEW -->
            <div id="tab-adversarial" class="tab-pane">
                <div class="card">
                    {html_adversarial}
                </div>
            </div>

        </div>
    </div>

    <script>
        document.addEventListener('DOMContentLoaded', function() {{
            const navButtons = document.querySelectorAll('.nav-btn');
            const tabPanes = document.querySelectorAll('.tab-pane');
            const mainViewport = document.getElementById('main-viewport');

            navButtons.forEach(function(btn) {{
                btn.addEventListener('click', function(e) {{
                    e.preventDefault();
                    const targetId = this.getAttribute('data-target');

                    // Remove active from all buttons
                    navButtons.forEach(b => b.classList.remove('active'));
                    // Hide all tab panes
                    tabPanes.forEach(pane => pane.classList.remove('active'));

                    // Activate clicked button and target pane
                    this.classList.add('active');
                    const targetPane = document.getElementById(targetId);
                    if (targetPane) {{
                        targetPane.classList.add('active');
                    }}

                    // Scroll to top
                    mainViewport.scrollTop = 0;
                }});
            }});
        }});
    </script>
</body>
</html>
"""

with open(HTML_OUT, "w", encoding="utf-8") as f:
    f.write(html_template)

print(f"✨ Master Interactive Deep Audit Dossier generated successfully at:\n   {HTML_OUT}")
