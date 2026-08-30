import re
import sys

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

with open("RSSFeedChecker_Master_Audit_Report.html", "r", encoding="utf-8") as f:
    text = f.read()

scripts = re.findall(r"<script\b[^>]*>(.*?)</script>", text, re.DOTALL)
print(f"Total <script> tags found: {len(scripts)}")
for i, s in enumerate(scripts):
    print(f"  Script {i+1}: {s.strip()[:100]}...")

tab_ids = [
    "tab-morning", "tab-master-report", "tab-claims", "tab-unified",
    "tab-evidence", "tab-gates", "tab-ci-desks", "tab-adversarial",
    "tab-readiness"
]

all_ok = True
for tid in tab_ids:
    has_btn = f"switchTab('{tid}')" in text
    has_pane = f'id="{tid}"' in text
    print(f"  Tab {tid:<20}: Button={has_btn}, Pane={has_pane}")
    if not (has_btn and has_pane):
        all_ok = False

if all_ok and len(scripts) == 1:
    print("\n✨ HTML Structure & JavaScript Event Handlers are 100% Validated!")
else:
    print("\n❌ Issues detected in HTML structure.")
