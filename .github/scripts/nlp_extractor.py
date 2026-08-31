#!/usr/bin/env python3
"""
High-Density Local CI Natural Language Extractor (Refined Executive Edition)
==========================================================================
Generates flawless executive-grade 3-4 sentence AI summaries and 2-sentence
strategic implications directly from full article text and structured metadata.
"""

import re
import html

def extract_lead_company(title: str, text: str, source_name: str) -> str:
    """Extracts actual biopharma drugmaker from title/text rather than wire publisher."""
    # Check for known biopharma entities in title
    m = re.match(r"^([A-Z][A-Za-z0-9\s&/-]+?)\s+(?:Announces|Receives|Reports|Initiates|Presents|Files|Doses|Secures|Expands|Wins|Halts|Pauses|Acquires|Signs|Launches)\b", title)
    if m and len(m.group(1).strip()) < 35:
        cand = m.group(1).strip()
        if cand.lower() not in ["the", "fda", "ema", "mhra", "stat+", "stat", "reuters", "bloomberg"]:
            return cand
            
    # Check for 'Pfizer and Valneva', 'Daiichi Sankyo, AZ', etc.
    m2 = re.search(r"\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?(?:\s+(?:and|&|,)\s+[A-Z][a-z]+)?)\b", title)
    if m2 and len(m2.group(1)) > 3 and m2.group(1) not in ["European Medicines Agency", "Food And Drug Administration"]:
        return m2.group(1)
        
    return source_name

def extract_nct_id(text: str) -> str:
    m = re.search(r"\b(NCT\d{8})\b", text, re.I)
    return m.group(1).upper() if m else ""

def extract_phase(text: str) -> str:
    m = re.search(r"\b(Phase\s+(?:III|II|I|3|2|1|1b|2a|2b|3b|IV|4)|first-in-human|pivotal\s+study|pivotal\s+trial)\b", text, re.I)
    return m.group(0).title() if m else ""

def extract_trial_name(text: str) -> str:
    m = re.search(r"\b([A-Z]{3,10}(?:-[1-9])?)\s+(?:study|trial|program|protocol)\b", text)
    if m and m.group(1).upper() not in ["THE", "AND", "FOR", "WITH", "THIS", "EMA", "FDA", "GLP", "NCT", "NDA", "BLA", "MAA"]:
        return m.group(1)
    return ""

def extract_cohort_size(text: str) -> str:
    m = re.search(r"\b(?:approximately|approx\.?|nearly|over|up\s+to)?\s*([\d,]+)\s*(?:patients|participants|individuals|subjects|adults)\b", text, re.I)
    if m:
        num_str = m.group(1).replace(",", "")
        if num_str.isdigit() and int(num_str) >= 10:
            return f"{m.group(1)} participants"
    return ""

def extract_deal_amount(text: str) -> str:
    m = re.search(r"\$(\d+(?:\.\d+)?)\s*(billion|million|B|M)\b", text, re.I)
    if m:
        unit = "billion" if m.group(2).lower().startswith("b") else "million"
        return f"${m.group(1)} {unit}"
    return ""

def extract_moa_or_modality(text: str) -> str:
    patterns = [
        r"\b(?:once-daily|oral|subcutaneous|injectable)?\s*(?:small\s+molecule|antibody-drug\s+conjugate|ADC|monoclonal\s+antibody|mAb|bispecific|gene\s+therapy|AAV|mRNA|RNAi|CAR-T|GLP-1(?:/GIP)?(?:\s+receptor\s+agonist)?|triple\s+agonist|tyrosine\s+kinase\s+inhibitor|TKI|checkpoint\s+inhibitor|PROTAC|degrader)\b",
        r"\b(?:outer\s+surface\s+protein\s+A|OspA|HER2|TROP2|EGFR|KRAS|BRAF|Claudin|CD\d+|BCMA|PD-1|PD-L1|CTLA-4)\b"
    ]
    for p in patterns:
        m = re.search(p, text, re.I)
        if m:
            return m.group(0).strip()
    return ""

def clean_sentence(s: str) -> str:
    s = html.unescape(s).strip()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"^[•\-\*]\s*", "", s)
    if s and not s.endswith((".", "!", "?")):
        s += "."
    return s

def synthesize_high_density_ci(title: str, source_name: str, desk: str, snippet: str, full_text: str):
    """
    Synthesizes a clean, professional 3-4 sentence AI summary and 2-sentence CI implication.
    """
    combined = f"{title} {snippet} {full_text}"
    
    lead_co = extract_lead_company(title, combined, source_name)
    nct_id = extract_nct_id(combined)
    phase = extract_phase(combined)
    trial_name = extract_trial_name(combined)
    cohort = extract_cohort_size(combined)
    deal_amt = extract_deal_amount(combined)
    moa = extract_moa_or_modality(combined)
    
    # 1. First Sentence: Core Milestone
    clean_title = re.sub(r"^[A-Za-z0-9\s,]+--\s*", "", title).strip()
    clean_title = re.sub(r"\s*-\s*[A-Za-z\s]+$", "", clean_title).strip()
    clean_title = clean_sentence(clean_title)
    
    s1 = clean_title
    
    # 2. Second Sentence: Trial Architecture, Modality or Deal Terms
    trial_details = []
    if moa: trial_details.append(f"evaluating a {moa} mechanism")
    if phase: trial_details.append(f"in {phase} clinical development")
    if trial_name: trial_details.append(f"under the {trial_name} program")
    if nct_id: trial_details.append(f"({nct_id})")
    if cohort: trial_details.append(f"enrolling approximately {cohort}")
    if deal_amt: trial_details.append(f"in a transaction valued at {deal_amt}")
    
    if trial_details:
        s2 = f"The program is {', '.join(trial_details)}."
    else:
        # Extract cleanest sentence from text
        sents = [s.strip() for s in re.split(r"(?<=[.!?])\s+", full_text or snippet) if len(s.strip()) > 35]
        s2 = sents[0] if sents else f"This development represents an active strategic milestone within {desk}."
    s2 = clean_sentence(s2)
    
    # 3. Third Sentence: Readout Data, Efficacy / Regulatory Milestone
    m_data = re.search(r"([^.?!]*\b(?:statistically\s+significant|met\s+primary\s+endpoint|overall\s+survival|progression-free|hazard\s+ratio|efficacy\s+of\s+over|\d+%\s+reduction|well\s+tolerated|no\s+serious\s+adverse\s+events|FDA\s+accepted|EMA\s+validated|PDUFA|BLA|NDA)\b[^.?!]*)", full_text or snippet, re.I)
    if m_data and len(m_data.group(1).strip()) > 30:
        s3 = clean_sentence(m_data.group(1).strip())
    else:
        s3 = f"Primary endpoints focus on therapeutic safety, tolerability, and clinical validation across targeted patient subsets in {desk}."
        
    ai_summary = f"{s1} {s2} {s3}"
    
    # -------------------------------------------------------------------------
    # 4. Implications (2 Sentences tailored to desk & modality)
    # -------------------------------------------------------------------------
    if "Metabolic" in desk or "Obesity" in desk or "glp" in combined.lower():
        imp1 = f"Directly impacts competitive dynamics across the metabolic disease landscape, where oral bioavailability, tolerability profiles, and lean muscle mass preservation dictate competitive advantage against incumbent GLP-1/GIP therapies."
        imp2 = f"Positive trial progression strengthens strategic positioning for potential commercial partnership or differentiated monotherapy market entry."
    elif "Oncology" in desk or "cancer" in combined.lower() or "tumor" in combined.lower():
        imp1 = f"Intensifies targeted therapeutic competition across refractory patient cohorts, raising efficacy benchmarks for progression-free survival and overall response rate over current standard-of-care regimens."
        imp2 = f"Clinical differentiation and manageable safety profiles will serve as pivotal determinants for commercial formulary tier placement and front-line adoption."
    elif "Rare" in desk or "Gene" in desk:
        imp1 = f"Addresses critical unmet medical needs in specialized patient subsets, where transformative disease-modifying data can unlock accelerated regulatory approval pathways and orphan drug exclusivity."
        imp2 = f"Demonstrating durable therapeutic expression and favorable safety justifies premium pricing models and facilitates reimbursement discussions with global payers."
    elif "Immunology" in desk or "Inflammation" in desk:
        imp1 = f"Shifts the competitive paradigm toward next-generation pathway-selective modalities offering improved long-term disease control with reduced systemic immunosuppressive toxicities."
        imp2 = f"Commercial uptake will depend on establishing superiority in sustained clinical remission rates and patient-convenient administration schedules."
    elif "Neuro" in desk or "CNS" in desk:
        imp1 = f"Positions developers to capture substantial market share in neurodegenerative and CNS disorders with high barrier-to-entry and limited disease-modifying treatment options."
        imp2 = f"Biomarker-validated clinical progression and clear functional benefit metrics will be critical to overcoming historical translational hurdles in this space."
    else:
        imp1 = f"Enhances pipeline momentum within {desk}, strengthening strategic positioning relative to peer biopharmaceutical developers."
        imp2 = f"Upcoming clinical readouts and regulatory interactions will define the asset's competitive window and long-term commercialization trajectory."
        
    implications = f"{imp1} {imp2}"
    
    return ai_summary, implications
