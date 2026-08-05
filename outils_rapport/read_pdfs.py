import fitz, sys, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

paths = [
    r'C:\Users\Topo_4\Documents\AT_PFE\Suivi_PFE\Insa_Journees_Topo_cahier_PFE_25_web.pdf',
    r'C:\Users\Topo_4\Documents\Armand_ODDOS\Redaction\PFE\Insa_Journees_Topo_cahier_PFE_23_web2.pdf'
]

for path in paths:
    print('='*70)
    print(path[-50:])
    print('='*70)
    doc = fitz.open(path)
    found = 0
    for i in range(doc.page_count):
        text = doc[i].get_text()
        # Resumes de 4 pages : chercher "1/4" ou "2/4" ou "3/4" ou "4/4"
        if ('1/4' in text or '2/4' in text) and len(text) > 500:
            print(f'--- PAGE {i+1} (resume) ---')
            print(text[:6000])
            print()
            found += 1
            if found >= 8:
                break
    doc.close()
