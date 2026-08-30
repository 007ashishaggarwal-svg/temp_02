#!/usr/bin/env python3
"""
AI Synthesis & Biopharma CI Strategic Implications Engine
=========================================================
1. Factual Summary (< 70 words):
   - Integrates AI REST API (if service key is configured in environment).
   - High-precision NLP synthesis engine that purges journalistic bylines, boilerplate,
     and datelines, synthesizing clean, factual, professional executive summaries (< 70 words).

2. Competitive Implications (< 70 words):
   - Top-tier Biopharma CI Analyst insight (Umer Raffat / Geoffrey Porges / Jacob Plieth style).
   - Evidence-led, commercially sharp, clinically/mechanistically aware.
   - Interprets: What changed, why it matters, and strategic implications (< 70 words).
"""

import os
import re
import html
import json
import urllib.request
import urllib.parse
from datetime import datetime

# Read from masked env var (CI_SVC_B in runner) with fallback for local dev
_API_KEY = os.environ.get("SVC_B") or os.environ.get("GEMINI_API_KEY", "")
GEMINI_API_KEY = _API_KEY  # internal alias kept for compatibility


def clean_text_input(t: str) -> str:
    if not t:
        return ""
    txt = html.unescape(str(t)).strip()
    if "&lt;" in txt or "&gt;" in txt or "&amp;" in txt:
        txt = html.unescape(txt)
    txt = re.sub(r'<(?:figure|script|style|svg)[^>]*>.*?</(?:figure|script|style|svg)>', ' ', txt, flags=re.DOTALL | re.IGNORECASE)
    txt = re.sub(r'<[^>]+>', ' ', txt)
    
    # Strip common journalistic boilerplate, bylines, and dates
    txt = re.sub(r'Written by [^|.\n]+(?:\|\s*[A-Za-z]+\.?\s*\d{1,2},?\s*\d{4})?', '', txt, flags=re.IGNORECASE)
    txt = re.sub(r'By\s+[A-Za-z\s,.\-–—]+(?:\|\s*[A-Za-z]+\.?\s*\d{1,2},?\s*\d{4})?', '', txt, flags=re.IGNORECASE)
    txt = re.sub(r'Photo (?:credit|by):?[^\n.]+', '', txt, flags=re.IGNORECASE)
    txt = re.sub(r'Published:?\s*[A-Za-z]+\s+\d{1,2},\s+\d{4}', '', txt, flags=re.IGNORECASE)
    txt = re.sub(r'Source:\s*\[?[^\]\n]+\]?', '', txt, flags=re.IGNORECASE)
    txt = re.sub(r'\[headline-only\]', '', txt, flags=re.IGNORECASE)
    txt = re.sub(r'\bGN-[A-Za-z0-9_\-]+\b', '', txt)
    txt = re.sub(r'\bRSS-[A-Za-z0-9_\-]+\b', '', txt)
    
    txt = re.sub(r'\s+', ' ', txt).strip()
    return txt


def clamp_words(text: str, max_words: int = 68) -> str:
    """Ensure output does not exceed word limit while maintaining complete sentences."""
    words = text.strip().split()
    if len(words) <= max_words:
        return text.strip()
    
    trimmed = " ".join(words[:max_words])
    last_period = max(trimmed.rfind("."), trimmed.rfind("!"), trimmed.rfind("?"))
    if last_period > len(trimmed) * 0.65:
        return trimmed[:last_period + 1].strip()
    return trimmed.rstrip(",;:- ") + "."


def call_gemini_api(prompt: str, timeout: int = 5) -> str:
    """Calls Gemini REST API directly if GEMINI_API_KEY is configured in environment."""
    if not GEMINI_API_KEY:
        return ""
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GEMINI_API_KEY}"
        headers = {"Content-Type": "application/json"}
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": 0.2, "maxOutputTokens": 150}
        }
        req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), headers=headers)
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            return data["candidates"][0]["content"]["parts"][0]["text"].strip()
    except Exception:
        return ""


def generate_factual_ai_summary(
    headline: str,
    url: str,
    snippet: str,
    full_text: str = "",
    source_name: str = "",
    desk: str = "",
    signal_type: str = ""
) -> str:
    """
    Generates a direct, fact-first executive summary (< 70 words).
    Starts directly with the biopharma fact, sponsor, or drug asset.
    """
    h_clean = clean_text_input(headline)
    # Strip trailing publisher suffix from headline if present (e.g. - Yahoo Finance)
    h_clean = re.sub(r'\s+-\s+[A-Za-z0-9\s\.\!&–—\']+$', '', h_clean).strip()
    
    s_clean = clean_text_input(snippet)
    f_clean = clean_text_input(full_text)
    lead_source = source_name or "Trade Press"
    if "GN-" in lead_source or "RSS-" in lead_source:
        lead_source = "Biopharma Newsroom"
    desk_clean = desk.replace(" Desk", "").strip() if desk else "Biopharma"

    # Try Gemini API if key available
    if GEMINI_API_KEY:
        prompt = (
            "Write a single, factual, high-quality biopharma news summary (< 70 words) for the following development. "
            "Start directly with the company/sponsor, drug asset, or key milestone. "
            "Never start with 'Source reported that' or system IDs. "
            "Include company, asset, phase/mechanism, and clinical/economic metrics if present.\n\n"
            f"Headline: {h_clean}\n"
            f"Source: {lead_source}\n"
            f"Snippet: {s_clean[:350]}\n"
            f"Full Text Excerpt: {f_clean[:500]}"
        )
        gemini_res = call_gemini_api(prompt)
        if gemini_res and len(gemini_res.split()) > 10:
            return clamp_words(gemini_res, max_words=68)

    # Advanced Rule-Based NLP Synthesis
    comb = f"{h_clean} {s_clean} {f_clean}"

    # Extract Key Clinical / Regulatory entities
    m_phase = re.search(r'\b(Phase\s+[1234]|Phase\s+I{1,3}|Phase\s+IV|Pivotal\s+Trial|PDUFA|NDA|BLA|IND|EIR|CRL|FDA\s+Approval|FDA\s+Cleared|CHMP\s+Opinion|Breakthrough\s+Therapy|Fast\s+Track|Orphan\s+Drug|CE\s+Mark)\b', comb, re.I)
    phase_str = m_phase.group(1).title() if m_phase else ""

    m_num = re.search(r'(\b\d+(\.\d+)?%\s*(?:weight\s+loss|reduction|response\s+rate|orr|pfs|os|efficacy|mace)\b|\bp\s*[<:=]\s*0\.\d+\b|\$\d+(\.\d+)?\s*(?:billion|million|B|M|/month)\b)', comb, re.I)
    num_str = m_num.group(1) if m_num else ""

    # Parse clean sentences from snippet or full text
    sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', s_clean) if s.strip() and len(s.strip()) > 20]
    
    clean_sentences = []
    for s in sentences:
        s_low = s.lower()
        if any(bad in s_low for bad in ["written by", "photo credit", "all rights reserved", "subscribe to", "read more", "celebrated our", "gn-"]):
            continue
        if s_clean.lower().count(s_low[:30]) > 1:
            continue
        if h_clean.lower() in s_low and len(s_low) <= len(h_clean) + 10:
            continue
        clean_sentences.append(s)

    if clean_sentences and len(clean_sentences[0].split()) >= 10:
        first_s = clean_sentences[0]
        if h_clean.lower() in first_s.lower():
            raw_summary = first_s
        else:
            raw_summary = f"{h_clean}. {first_s}"
    else:
        # Build direct fact-first biopharma sentence
        s_parts = [f"{h_clean}."]
        if phase_str and phase_str.lower() not in h_clean.lower():
            s_parts.append(f"The program is advancing in {phase_str} clinical evaluation.")
        elif num_str and num_str.lower() not in h_clean.lower():
            s_parts.append(f"The asset demonstrated notable therapeutic impact ({num_str}).")
        else:
            s_parts.append(f"This represents a key developmental milestone within {desk_clean}.")
        raw_summary = " ".join(s_parts)

    raw_summary = re.sub(r'\s+', ' ', raw_summary).strip()
    return clamp_words(raw_summary, max_words=68)


def generate_ci_competitive_implications(
    headline: str,
    url: str,
    snippet: str,
    full_text: str = "",
    source_name: str = "",
    desk: str = "",
    signal_type: str = "",
    matched_keywords: str = ""
) -> str:
    """
    Generates a top-tier Biopharma CI Analyst insight (< 70 words).
    Interprets: What changed, why it matters, and strategic implications for payers, market access, or rival pipelines.
    """
    h_clean = clean_text_input(headline)
    s_clean = clean_text_input(snippet)
    f_clean = clean_text_input(full_text)
    comb = f"{h_clean} {s_clean} {f_clean}".lower()
    desk_name = desk.replace(" Desk", "").strip() if desk else "Biopharma"

    # Try Gemini API if key available
    if GEMINI_API_KEY:
        prompt = (
            "Using current evidence, write a ≤68-word competitive strategic insight. "
            "Think like a top-tier biopharma CI analyst (e.g., Umer Raffat, Geoffrey Porges, Jacob Plieth): "
            "evidence-led, commercially sharp, and clinically/mechanistically aware.\n"
            "Do not summarize—interpret the competitive significance: what changed, why it matters, and strategic implications.\n\n"
            f"Headline: {h_clean}\n"
            f"Desk: {desk_name}\n"
            f"Signal: {signal_type}\n"
            f"Details: {s_clean[:400]}"
        )
        gemini_imp = call_gemini_api(prompt)
        if gemini_imp and len(gemini_imp.split()) > 10:
            return clamp_words(gemini_imp, max_words=68)

    # Category 1: Health Economics / HEOR / Payer Coverage / Pricing
    if any(w in comb for w in ["heor", "cost offset", "medicare", "medicaid", "hospitalization", "pricing", "mfn", "reimbursement", "payer", "ira", "negotiated"]):
        if any(w in comb for w in ["zepbound", "wegovy", "glp-1", "hospital", "cost offset"]):
            implication = (
                "Provides critical actuarial proof to justify Medicare, Medicaid, and commercial employer coverage. "
                "Quantifying acute care cost offsets directly dismantles payer pushback over high upfront drug costs by demonstrating immediate downstream medical expenditure savings."
            )
        elif any(w in comb for w in ["mfn", "white house", "reference price", "pricing reform"]):
            implication = (
                "Signals expanding federal pressure on drug pricing beyond top Medicare-negotiated drugs into mid-tier biopharma. "
                "Broad international reference pricing could compress operating margins and alter launch pricing strategies for emerging specialty entrants."
            )
        else:
            implication = (
                "Strengthens manufacturer negotiating leverage during formulary reviews by linking clinical efficacy to direct healthcare cost savings. "
                "Demonstrating quantifiable economic offsets reduces payer restrictions and accelerates commercial tier placement."
            )

    # Category 2: Metabolic, Obesity, GLP-1/GIP, Diabetes, Cardiometabolic
    elif "Metabolic" in desk_name or "Obesity" in desk_name or any(w in comb for w in ["obesity", "glp-1", "tirzepatide", "semaglutide", "incretin", "weight loss", "cagrisema", "dd18"]):
        if any(w in comb for w in ["sensor", "wearable", "glucose", "ketone", "dka", "device"]):
            implication = (
                "Significantly de-risks metabolic and diabetes management, particularly for patients on intensive combination therapies at risk of metabolic complications. "
                "Represents a major technology upgrade for integrated continuous cardiometabolic monitoring."
            )
        elif any(w in comb for w in ["phase 3", "positive", "weight loss", "efficacy", "oral"]):
            implication = (
                "Directly challenges incumbent GLP-1 leaders on efficacy magnitude, tolerability profile, or oral dosing convenience. "
                "Tightens the commercial differentiation window for next-generation incretins and raises the clinical bar for maintenance regimens."
            )
        elif any(w in comb for w in ["deal", "acquire", "license", "partner", "buyout"]):
            implication = (
                "Accelerates clinical pipeline access as large-cap drugmakers seek diversified non-incretin and oral assets ahead of patent expirations. "
                "Increases premium valuations for clinical-stage metabolic biotechs."
            )
        else:
            implication = (
                "Reinforces competitive intensity across metabolic indications, where differentiation on lean mass preservation, dosing intervals, and cardiovascular benefits dictates long-term market leadership."
            )

    # Category 3: Cardiovascular / Inflammation / Vaccines / Epidemiology
    elif any(w in comb for w in ["cardiovascular", "mace", "stroke", "infarction", "shingrix", "vaccine", "heart failure"]):
        implication = (
            "Highlights the systemic vascular benefits of suppressing chronic inflammation in aging and cardiometabolic populations. "
            "Adds weight to the paradigm that mitigating inflammatory triggers provides additive cardiovascular risk reduction alongside standard metabolic therapies."
        )

    # Category 4: Oncology / Myeloma / Lung Cancer / ADCs
    elif "Oncology" in desk_name or "Myeloma" in desk_name or any(w in comb for w in ["oncology", "cancer", "nsclc", "sclc", "myeloma", "adc", "car-t"]):
        if any(w in comb for w in ["approval", "cleared", "bla", "nda", "fda"]):
            implication = (
                "Reshapes standard-of-care lines in refractory patient subsets, intensifying competitive pressure on incumbent targeted therapies. "
                "Commercial traction will hinge on real-world progression-free survival and favorable toxicity management."
            )
        elif any(w in comb for w in ["fail", "missed", "hold", "halt", "crl"]):
            implication = (
                "Creates an immediate commercial opening for competing targeted and ADC programs in the same treatment line, prompting strategic reallocation of pipeline capital toward derisked combination assets."
            )
        else:
            implication = (
                "Highlights accelerating competition across next-generation targeted modalities. "
                "Establishing superior overall survival and manageable safety profiles will determine first-line market access against established standards."
            )

    # Category 5: Neuroscience / CNS / ALS / Alzheimer's / Rare Diseases
    elif "Neuroscience" in desk_name or "CNS" in desk_name or any(w in comb for w in ["neuro", "cns", "als", "alzheimer", "parkinson", "rare"]):
        implication = (
            "Marks a critical clinical inflection in a high-unmet-need CNS indication where disease-modifying therapies command premium pricing. "
            "Commercial adoption depends heavily on biomarker validation, patient identification infrastructure, and long-term safety monitoring."
        )

    # Category 6: Corporate Strategy & M&A / Partnerships
    elif "Corporate" in desk_name or any(w in comb for w in ["acquire", "deal", "licensing", "restructuring", "layoff", "pipeline"]):
        implication = (
            "Reflects strategic portfolio prioritization and capital redeployment toward high-conviction clinical assets. "
            "Signals potential follow-on licensing deals or asset divestitures as large-cap players optimize therapeutic focus."
        )

    # General Biopharma CI Insight
    else:
        implication = (
            f"Provides actionable intelligence on pipeline momentum and regulatory milestones for {source_name or 'monitored entities'}. "
            "Near-term catalyst readouts will define competitive standing and market positioning within this indication."
        )

    return clamp_words(implication, max_words=68)
