import re
import urllib.request
import ssl

with open('PFE_AT.bib', 'r', encoding='utf-8', errors='ignore') as f:
    content = f.read()

urls = re.findall(r'url\s*=\s*\{(.+?)\}', content)

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

for u in urls:
    if 'legifrance' in u:
        print(f"[SKIP] Legifrance URL (anti-bot): {u}")
        continue
    
    # Clean up TeX escaped chars for testing
    clean_u = u.replace('\\%', '%')
    
    req = urllib.request.Request(clean_u, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
    try:
        resp = urllib.request.urlopen(req, context=ctx, timeout=10)
        print(f"[OK] {resp.status} - {u}")
    except Exception as e:
        print(f"[ERROR] {u} -> {e}")
