#!/usr/bin/env python3
import re
import html
import sys

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

def clean_snippet(raw_text: str) -> str:
    """Clean RSS snippet: unescape entities, strip tags, and remove boilerplate."""
    if not raw_text:
        return ""
    t = raw_text
    # Unescape HTML entities multiple times (e.g. &lt;a href...&gt;)
    for _ in range(3):
        t = html.unescape(t)
    # Strip HTML tags
    t = re.sub(r"<[^>]+>", " ", t)
    # Strip publisher suffixes
    suffixes = [
        r"\s*[-–—|:]\s*STAT(?:\s*News)?\s*$",
        r"\s*[-–—|:]\s*Endpoints(?:\s*News)?\s*$",
        r"\s*[-–—|:]\s*FiercePharma\s*$",
        r"\s*[-–—|:]\s*FierceBiotech\s*$",
        r"\s*[-–—|:]\s*BioSpace\s*$",
        r"\s*[-–—|:]\s*PR\s*Newswire\s*$",
        r"\s*[-–—|:]\s*Business\s*Wire\s*$",
        r"\s*[-–—|:]\s*GlobeNewswire\s*$",
        r"\s*[-–—|:]\s*Reuters\s*$",
        r"\s*[-–—|:]\s*Bloomberg\s*$",
    ]
    for suf in suffixes:
        t = re.sub(suf, "", t, flags=re.I)
    # Collapse whitespace
    t = " ".join(t.split()).strip()
    return t


def clean_article_html(raw_html: str) -> str:
    """
    Extract genuine English article paragraphs from raw HTML.
    Strips ALL <script>, <style>, <noscript>, <svg>, <nav>, <footer>, and JavaScript code.
    """
    if not raw_html:
        return ""
    
    # 1. Remove all scripts, styles, noscripts, iframes, and HTML comments
    h = re.sub(r"<(script|style|noscript|iframe|svg|nav|footer|header|aside)\b[^>]*>.*?</\1>", " ", raw_html, flags=re.I | re.S)
    h = re.sub(r"<!--.*?-->", " ", h, flags=re.S)

    # 2. Extract <p> paragraphs
    p_matches = re.findall(r"<p\b[^>]*>(.*?)</p>", h, re.I | re.S)
    clean_paragraphs = []

    # Filter out JavaScript tokens, minified code, or garbage
    js_indicators = [
        "function(", "var ", "const ", "let ", "===", "!==", "firstChild",
        "getAttribute", "setAttribute", "document.", "window.", "jQuery", "$.",
        "addEventListener", "return ", "typeof ", "prototype", ".min.js", "cookie",
        "{", "}", "();", ");", "/*", "*/", "eval("
    ]

    for raw_p in p_matches:
        p_clean = clean_snippet(raw_p)
        if len(p_clean) < 45 or len(p_clean) > 1200:
            continue
        
        # Check if paragraph contains code/JS artifacts
        if any(js_token in p_clean for js_token in js_indicators):
            continue
        
        # Must have English letters and sentence structure
        words = p_clean.split()
        if len(words) < 7:
            continue

        clean_paragraphs.append(p_clean)

    if clean_paragraphs:
        return " ".join(clean_paragraphs[:3])[:800]
    return ""

if __name__ == "__main__":
    # Test 1: Kalkine media raw Google News RSS snippet
    raw_google_desc = '&lt;a href="https://news.google.com/rss/articles/CBMijwJBVV95cUxPUElEemszX1c0MEwyZWl4QTVKUlktNUxWdnF0Z204RTR1TEw2a2ptbUQ2MUxSVHowaFdYMXVQejZFSk9UWDVYaWJRWVI2eVpjb1dzQUlCd3lpUlBUUmVfQ0dBbWVlVEJnSjhwMDV6RUdFeWFCakVOb1FKNkxJRkNVXzhsTmFPN24tU0FmX3lwTm1iNUZEYndWXzVtMkRTZk5lNlRRQk1PR2E3SldBOGFJYUpoNzlOOUMtQk5DaEhsTWgtdkxJNkthZVd3MmNUcVduYmpKdlVRNnBWQzYzb2o5" target="_blank"&gt;PureTech Health\'s Celea Therapeutics to Unveil SURPASS-IPF Phase 3 Trial Design and Deupirfenidone Interaction Data at ERS Congress 2026&lt;/a&gt;&amp;nbsp;&amp;nbsp;&lt;font color="#6f6f6f"&gt;Kalkine Media&lt;/font&gt;'
    print("Cleaned Snippet:")
    print(clean_snippet(raw_google_desc))
    print("-" * 60)

    # Test 2: Manila Times garbage JavaScript HTML
    raw_manila_html = '<html><head><script>","#"===e.firstChild.getAttribute("href")})||fe("type|href|height|width",function(e,t,n){if(!n)return e.getAttribute(t,"type"===t.toLowerCase()?1:2)}),d.attributes&&ce(function(e){return e.innerHTML=" ",e.firstChild.setAttribute("value",""),""===e.firstChild.getAttribute("value")})||fe("value",function(e,t,n){if(!n&&"input"===e.nodeName.toLowerCase())return e.defaultValue}),ce(function(e){return null==e.getAttribute("disabled")})||fe(R,function(e,t,n){var r;if(!n)return!0===e[t]?t.toLowerCase():(r=e.getAttributeNode(t))&&r.specified?r.value:null}),se}(C);S.find=d,S.expr=d.selectors,S.expr[":"]</script></head><body><p>Endovia Health Sciences, formerly Splash Beverage Group, announced today that it has achieved its first major regulatory milestone with the U.S. Food and Drug Administration for its CannEpil veterinary development program.</p><p>The company said the FDA confirmed the regulatory pathway for CannEpil, a proprietary formulation designed to treat canine epilepsy.</p></body></html>'
    print("Cleaned Body Paragraphs:")
    print(clean_article_html(raw_manila_html))
    print("-" * 60)
