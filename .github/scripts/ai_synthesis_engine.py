#!/usr/bin/env python3
"""
ENTERPRISE MULTI-PROVIDER AI SYNTHESIS & DUAL-QUOTA POOL ENGINE
==============================================================
Supported Providers:
1. Google Gemini Flash (Dual/Multi-Key Pool with Round-Robin & 429 Failover)
2. Cloudflare Workers AI (Llama 3.3 70B / Llama 3.1 8B Edge Inference)
3. High-Density Local CI Natural Language Extractor (100% Deterministic)

Features:
- Dual/Multi-Key Round-Robin for Gemini (doubles/triples RPM and RPD quotas)
- Zero-Cost Edge Inference via Cloudflare Workers AI (10,000 Neurons/day free)
- Automatic Cascading Failover (Primary -> Secondary -> Tertiary)
- Granular Provenance Tagging per row
"""

import os
import sys
import re
import json
import time
import urllib.request
import itertools

sys.path.insert(0, os.path.dirname(__file__))
from robust_fetcher import strip_wire_datelines, clean_snippet
from nlp_extractor import synthesize_high_density_ci

# -----------------------------------------------------------------------------
# 1. CREDENTIALS & KEY POOLS
# -----------------------------------------------------------------------------
# Read keys from environment variables or local Secrets fallback
def get_secret_keys():
    gemini_keys = []
    # Check env vars (CI_SVC series)
    for k in ["CI_SVC_B", "CI_SVC_C", "GEMINI_API_KEY"]:
        v = os.environ.get(k, "").strip()
        if v and v not in gemini_keys:
            gemini_keys.append(v)
            
    # Check Cloudflare env vars (CI_SVC series)
    cf_account_id = os.environ.get("CI_SVC_D", "").strip() or os.environ.get("CLOUDFLARE_ACCOUNT_ID", "").strip()
    cf_api_token = os.environ.get("CI_SVC_E", "").strip() or os.environ.get("CLOUDFLARE_API_TOKEN", "").strip()

    # Check Secrets.txt file if present locally
    sec_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "Secrets.txt"))
    
    if os.path.exists(sec_path):
        try:
            with open(sec_path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()
                # Find Gemini keys
                m_keys = re.findall(r"(?:AQ\.[A-Za-z0-9_-]{40,}|AIzaSy[A-Za-z0-9_-]{33})", content)
                for k in m_keys:
                    if k not in gemini_keys:
                        gemini_keys.append(k)
                        
                # Find Cloudflare credentials if recorded
                m_cf_acc = re.search(r"CLOUDFLARE_ACCOUNT_ID\s*[:=]\s*([a-f0-9]{32})", content, re.I)
                if m_cf_acc and not cf_account_id:
                    cf_account_id = m_cf_acc.group(1).strip()
                    
                m_cf_tok = re.search(r"CLOUDFLARE_API_TOKEN\s*[:=]\s*([A-Za-z0-9_-]{35,})", content, re.I)
                if m_cf_tok and not cf_api_token:
                    cf_api_token = m_cf_tok.group(1).strip()
        except Exception:
            pass
            
    # Return discovered keys from environment / Secrets.txt
    return gemini_keys, cf_account_id, cf_api_token

# Initialize Key Pools
GEMINI_KEYS, CF_ACCOUNT_ID, CF_API_TOKEN = get_secret_keys()
_gemini_key_cycle = itertools.cycle(GEMINI_KEYS) if GEMINI_KEYS else None

# -----------------------------------------------------------------------------
# 2. PROMPT BUILDER
# -----------------------------------------------------------------------------
def build_grounded_prompt(title: str, company: str, desk: str, snippet: str, full_text: str) -> str:
    context = full_text if (full_text and len(full_text) > 150) else (snippet if snippet else title)
    context = strip_wire_datelines(context)[:3200]
    return f"""You are a Senior Biopharmaceutical Competitive Intelligence (CI) Analyst.
Synthesize this pharmaceutical/clinical event into exactly two structured sections based STRICTLY on the facts provided below. Do not hallucinate or repeat generic templates.

Headline: {title}
Company: {company}
Desk: {desk}
Source Text:
\"\"\"
{context}
\"\"\"

Output format must be strictly valid JSON:
{{
  "ai_summary": "A comprehensive 3-to-4 sentence summary detailing (1) the core milestone, drug candidate (INN), MOA/target, and sponsor; (2) clinical trial phase/NCT ID or deal terms; (3) quantitative efficacy/safety data and next regulatory milestone.",
  "implications": "A precise 2-sentence CI analysis detailing (1) impact on standard-of-care benchmarks and competitor positioning in {desk}; (2) commercial market access or dosing differentiation."
}}
Return ONLY valid JSON."""

# -----------------------------------------------------------------------------
# 3. ENGINE 1: DUAL-KEY ROUND-ROBIN GEMINI FLASH
# -----------------------------------------------------------------------------
def generate_gemini_flash(title: str, company: str, desk: str, snippet: str, full_text: str):
    """Calls Google Gemini Flash using multi-key round-robin with automatic failover."""
    global GEMINI_KEYS
    if not GEMINI_KEYS:
        return None, None, "No Gemini API Keys configured"
        
    prompt = build_grounded_prompt(title, company, desk, snippet, full_text)
    
    # Try all available keys before failing
    for key_idx, api_key in enumerate(GEMINI_KEYS, start=1):
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {
                "temperature": 0.1,
                "maxOutputTokens": 350
            }
        }
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"}, method="POST")
        
        try:
            with urllib.request.urlopen(req, timeout=12) as resp:
                res = json.loads(resp.read().decode("utf-8"))
                raw_text = res["candidates"][0]["content"]["parts"][0]["text"]
                m_json = re.search(r"\{.*\}", raw_text, re.DOTALL)
                if m_json:
                    parsed = json.loads(m_json.group(0))
                    ai_s = parsed.get("ai_summary", "")
                    imp = parsed.get("implications", "")
                    if len(ai_s) > 100 and len(imp) > 40:
                        prov_tag = f"⚡ Google Gemini 2.5 Flash (Key #{key_idx})" if len(GEMINI_KEYS) > 1 else "⚡ Google Gemini 2.5 Flash"
                        return ai_s, imp, prov_tag
        except urllib.error.HTTPError as e:
            if e.code == 429:
                # Rate limited on this key, immediately try next key in pool
                continue
        except Exception:
            continue
            
    return None, None, "Gemini Quota Exceeded / Unavailable"

# -----------------------------------------------------------------------------
# 4. ENGINE 2: CLOUDFLARE WORKERS AI (LLAMA 3.3 70B / LLAMA 3.1 8B EDGE)
# -----------------------------------------------------------------------------
def generate_cloudflare_ai(title: str, company: str, desk: str, snippet: str, full_text: str):
    """Calls Cloudflare Workers AI Edge Inference API ($0.00 / 10,000 Free Neurons/day)."""
    _, account_id, api_token = get_secret_keys()
    if not account_id or not api_token:
        return None, None, "Cloudflare Account ID / API Token not configured"
        
    prompt = build_grounded_prompt(title, company, desk, snippet, full_text)
    
    # Use Llama 3.3 70B or Llama 3.1 8B Instruct model
    model = "@cf/meta/llama-3.3-70b-instruct"
    url = f"https://api.cloudflare.com/client/v4/accounts/{account_id}/ai/run/{model}"
    
    payload = {
        "messages": [
            {"role": "system", "content": "You are a Senior Biopharmaceutical CI Analyst. Output strictly valid JSON with keys ai_summary and implications."},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.1,
        "max_tokens": 350
    }
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        headers={
            "Authorization": f"Bearer {api_token}",
            "Content-Type": "application/json"
        },
        method="POST"
    )
    
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            res = json.loads(resp.read().decode("utf-8"))
            if res.get("success"):
                raw_text = res.get("result", {}).get("response", "")
                m_json = re.search(r"\{.*\}", raw_text, re.DOTALL)
                if m_json:
                    parsed = json.loads(m_json.group(0))
                    ai_s = parsed.get("ai_summary", "")
                    imp = parsed.get("implications", "")
                    if len(ai_s) > 100 and len(imp) > 40:
                        return ai_s, imp, "🧠 Cloudflare Workers AI (Llama 3.3 70B Edge)"
    except Exception as e:
        pass
        
    return None, None, "Cloudflare Workers AI Unavailable"

# -----------------------------------------------------------------------------
# 5. ENGINE 3: HIGH-DENSITY LOCAL DETERMINISTIC EXTRACTOR (100% RELIABLE)
# -----------------------------------------------------------------------------
def generate_local_extractor(title: str, company: str, desk: str, snippet: str, full_text: str):
    """Executes high-density local clinical and strategic NLP extraction."""
    ai_s, imp = synthesize_high_density_ci(title, company, desk, snippet, full_text)
    return ai_s, imp, "🔬 Local High-Density CI Extractor"

# -----------------------------------------------------------------------------
# 6. UNIFIED CASCADE ROUTER
# -----------------------------------------------------------------------------
def synthesize_event(title: str, company: str, desk: str, snippet: str, full_text: str,
                     primary_mode: str = "Option B", secondary_mode: str = "Option A",
                     priority_tier: str = "Tier 1"):
    """
    Synthesizes an event according to the user's priority cascade configuration.
    """
    clean_s = strip_wire_datelines(clean_snippet(snippet))
    clean_f = strip_wire_datelines(full_text) if full_text else clean_s
    
    # Define engine dispatcher
    def run_engine(mode_str):
        if "Option B" in mode_str:
            return generate_gemini_flash(title, company, desk, clean_s, clean_f)
        elif "Option C" in mode_str:
            return generate_cloudflare_ai(title, company, desk, clean_s, clean_f)
        else:
            return generate_local_extractor(title, company, desk, clean_s, clean_f)
            
    # 1. Attempt Primary Engine
    ai_txt, imp_txt, prov = run_engine(primary_mode)
    if ai_txt and imp_txt:
        return ai_txt, imp_txt, prov
        
    # 2. Attempt Secondary Failover Engine
    if secondary_mode and "None" not in secondary_mode:
        ai_txt, imp_txt, prov = run_engine(secondary_mode)
        if ai_txt and imp_txt:
            return ai_txt, imp_txt, f"{prov} (Failover)"
            
    # 3. Ultimate Fallback to Local High-Density Extractor
    ai_txt, imp_txt, prov = generate_local_extractor(title, company, desk, clean_s, clean_f)
    return ai_txt, imp_txt, f"{prov} (Fallback)"
