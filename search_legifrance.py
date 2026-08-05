import urllib.request
import urllib.parse
import re

queries = [
    "LEGITEXT000006068236",
    "LEGIARTI000006842596"
]

for q in queries:
    url = 'https://html.duckduckgo.com/html/?q=' + urllib.parse.quote(q)
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
    try:
        html = urllib.request.urlopen(req).read().decode('utf-8')
        snippets = re.findall(r'<a class="result__snippet[^>]+>(.*?)</a>', html, re.IGNORECASE|re.DOTALL)
        print(f"Results for '{q}':")
        for s in snippets[:3]:
            # remove bold tags
            clean_s = re.sub(r'<[^>]+>', '', s).strip()
            print("  -", clean_s)
    except Exception as e:
        print(f"Error for '{q}': {e}")
