import os
import sys
import pypdf
import re

sys.stdout.reconfigure(encoding='utf-8')

folder = r"c:\Users\User\Desktop\App\All_Apps\RSSFeedChecker\ct.gov"
f1 = os.path.join(folder, "Record History _ ver. 1_ 2018-06-21 _ NCT03574597 _ ClinicalTrials.gov.pdf")
f2 = os.path.join(folder, "Record History _ ver. 201_ 2024-08-06 _ NCT03574597 _ ClinicalTrials.gov.pdf")

r1 = pypdf.PdfReader(f1)
r2 = pypdf.PdfReader(f2)

text1_full = "\n".join([p.extract_text() for p in r1.pages])
text2_full = "\n".join([p.extract_text() for p in r2.pages])

# Find the "Viewing V1" and "Viewing V201" sections - the actual study record content
v1_start = text1_full.find("Viewing V1")
v201_start = text2_full.find("Viewing V201")

text1 = text1_full[v1_start:] if v1_start != -1 else text1_full
text2 = text2_full[v201_start:] if v201_start != -1 else text2_full

# ---- EXTRACT KEY FIELDS FROM BOTH ----

def find_value(text, label):
    """Find the value after a label on the same or next line."""
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if label.lower() in line.lower():
            # Check if value is on same line after the label
            after = line.split(label, 1)[-1].strip() if label in line else ""
            if after:
                return after
            # Check next line
            if i + 1 < len(lines):
                return lines[i+1].strip()
    return "NOT FOUND"

fields = [
    "Overall Status",
    "Study Start",
    "Primary Completion",
    "Study Completion",
    "Enrollment",
    "First Posted",
    "Record Verification",
]

print("=" * 90)
print("FIELD-BY-FIELD COMPARISON: NCT03574597 (SELECT Trial - Semaglutide CV Outcomes)")
print("Novo Nordisk A/S | Phase 3 | Overweight/Obesity")
print("=" * 90)

for field in fields:
    v1_val = find_value(text1, field)
    v201_val = find_value(text2, field)
    changed = "⚠️ CHANGED" if v1_val != v201_val else "  (same)"
    print(f"\n{field}:")
    print(f"  V1  (2018-06-21): {v1_val}")
    print(f"  V201(2024-08-06): {v201_val}")
    print(f"  Status: {changed}")

# ---- ELIGIBILITY ----
print("\n" + "=" * 90)
print("ELIGIBILITY CRITERIA COMPARISON")
print("=" * 90)

def extract_section(text, start_label, end_labels):
    idx = text.find(start_label)
    if idx == -1:
        return "NOT FOUND"
    end = len(text)
    for el in end_labels:
        ei = text.find(el, idx + len(start_label))
        if ei != -1 and ei < end:
            end = ei
    return text[idx:end].strip()

elig1 = extract_section(text1, "Eligibility Criteria", ["Contacts and Locations", "Sex/Gender", "IPD Sharing"])
elig2 = extract_section(text2, "Eligibility Criteria", ["Contacts and Locations", "Sex/Gender", "IPD Sharing"])

# Count inclusion/exclusion items
inc1 = len(re.findall(r"Inclusion Criteria|●", elig1))
inc2 = len(re.findall(r"Inclusion Criteria|●", elig2))

print(f"V1 eligibility text length:  {len(elig1)} chars")
print(f"V201 eligibility text length: {len(elig2)} chars")

# ---- OUTCOME MEASURES COMPARISON ----
print("\n" + "=" * 90)
print("OUTCOME MEASURES COMPARISON")
print("=" * 90)

om1 = extract_section(text1, "Outcome Measures", ["Eligibility Criteria", "Arms and Interventions"])
om2 = extract_section(text2, "Outcome Measures", ["Eligibility Criteria", "Arms and Interventions"])

# Count primary and secondary
pri1 = om1.count("Primary Outcome")
pri2 = om2.count("Primary Outcome")
sec1 = len(re.findall(r"^\d+\.", om1, re.M))
sec2 = len(re.findall(r"^\d+\.", om2, re.M))

print(f"V1 Outcome Measures text: {len(om1)} chars")
print(f"V201 Outcome Measures text: {len(om2)} chars")
print(f"V1 numbered endpoints: {sec1}")
print(f"V201 numbered endpoints: {sec2}")

# ---- V1 TIME FRAMES ----
print("\n--- V1 Time Frames ---")
for m in re.finditer(r"\[Time Frame:.*?\]", om1):
    print(f"  {m.group(0)}")

print("\n--- V201 Time Frames ---")
for m in re.finditer(r"\[Time Frame:.*?\]", om2):
    print(f"  {m.group(0)}")

# ---- RESULTS CHECK ----
print("\n" + "=" * 90)
print("STUDY RESULTS")
print("=" * 90)

has_results_v1 = "Study Results" in text1 and "Participant Flow" in text1
has_results_v201 = "Study Results" in text2 and "Participant Flow" in text2

print(f"V1 has Study Results:   {has_results_v1}")
print(f"V201 has Study Results: {has_results_v201}")

if has_results_v201:
    # Extract key numbers from participant flow
    pf = extract_section(text2, "Participant Flow", ["Baseline Characteristics"])
    print(f"\nParticipant Flow snippet:")
    print(pf[:500])
    
    # Extract baseline characteristics snippet
    bc = extract_section(text2, "Baseline Characteristics", ["Outcome Measures"])
    print(f"\nBaseline Characteristics snippet:")
    print(bc[:500])

    # Adverse events
    ae = extract_section(text2, "Adverse Events", ["More Information"])
    print(f"\nAdverse Events text length: {len(ae)} chars")
    print("AE snippet:")
    print(ae[:500])

# ---- COUNTRIES ----
print("\n" + "=" * 90)
print("COUNTRY COUNT")
print("=" * 90)

countries1 = set()
countries2 = set()
for line in text1.splitlines():
    m = re.match(r".*,\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)$", line.strip())
    if m and len(m.group(1)) > 3:
        countries1.add(m.group(1))

for line in text2.splitlines():
    m = re.match(r".*,\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)$", line.strip())
    if m and len(m.group(1)) > 3:
        countries2.add(m.group(1))

print(f"V1 unique countries detected: {len(countries1)}")
if countries1:
    print(f"  {sorted(countries1)[:20]}")
print(f"V201 unique countries detected: {len(countries2)}")
if countries2:
    print(f"  {sorted(countries2)[:20]}")

# ---- VERSION HISTORY LOG - what categories changed ----
print("\n" + "=" * 90)
print("VERSION HISTORY: CATEGORIES OF CHANGES ACROSS 201 VERSIONS")
print("=" * 90)

# Extract version history table from the first part of the PDF
history_text = text2_full[:text2_full.find("Viewing V201") if text2_full.find("Viewing V201") != -1 else len(text2_full)]

change_categories = {}
for m in re.finditer(r"\d+\s+\d{4}-\d{2}-\d{2}\s+(.*?)(?=\d+\s+\d{4}-\d{2}-\d{2}|\Z)", history_text, re.S):
    changes = m.group(1).strip()
    for cat in ["Study Status", "Contacts/Locations", "Eligibility", "Recruitment Status", 
                 "Study Design", "Outcome Measures", "References", "Document Section",
                 "Participant Flow", "Baseline Characteristics", "Outcome Measures (Results)",
                 "Adverse Events", "More Information"]:
        if cat in changes:
            change_categories[cat] = change_categories.get(cat, 0) + 1

print("How many versions touched each category:")
for cat, count in sorted(change_categories.items(), key=lambda x: -x[1]):
    print(f"  {cat:<35} {count} times")
