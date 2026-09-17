with open('generate_final_180s_deck.py', 'r', encoding='utf-8') as f:
    text = f.read()

# Remplacer les sauts de lignes accidentels
text = text.replace('p1.text = "• " + c_hd + " :\n"', 'p1.text = "• " + c_hd + " :"')
text = text.replace('po1.text = "• " + ohd + " :\n"', 'po1.text = "• " + ohd + " :"')
text = text.replace('pgl_h.text = "GLiNER : APPROCHE BI-ENCODEUR ZERO-SHOT (RETENU)\n"', 'pgl_h.text = "GLiNER : Approche bi-encodeur zero-shot (retenu)"')
text = text.replace('plm_h.text = "LayoutLMv3 : APPROCHE MULTIMODALE 2D (ÉCARTÉ)\n"', 'plm_h.text = "LayoutLMv3 : Approche multimodale 2D (écarté)"')

with open('generate_final_180s_deck.py', 'w', encoding='utf-8') as f:
    f.write(text)

print("Nettoyage syntaxique termine !")
