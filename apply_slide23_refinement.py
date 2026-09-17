# -*- coding: utf-8 -*-
"""
Update Slide 23 in generate_final_180s_deck.py:
- Showcases the real crop of « Lachapelle /s/ AUBENAS ».
- Justifies why OCR/HTR + local database is chosen instead of an all-purpose national AI.
- Cites neighboring homonymous communes: Lachapelle-Graillouse (07), Lachapelle-sous-Chanéac (07), La Chapelle-en-Vercors (26).
- Replaces pompous jargon ('Désambiguïsation toponymique contextuelle') with direct, concrete phrasing.
- Inserts punchy, keyword-based speaker notes for Adrien's oral defense.
"""

with open('generate_final_180s_deck.py', 'r', encoding='utf-8') as f:
    text = f.read()

s23_start = text.find('# Slide 23 (4.4)')
s23_end = text.find('# ==================== CHAPITRE 5')

new_slide23 = '''# Slide 23 (4.4) : Justification de la chaîne hybride et arbitrage VLM local
print("Génération Slide 23 (Exemple concret Lachapelle-sous-Aubenas et justification hybride)...")
s23, c23 = init_standard_slide(3, 3, 23, "4.4 Arbitrage visuel local : Pourquoi une chaîne hybride plutôt qu'une IA globale ?")

# Colonne gauche (5.48 in) : Cas réel et risque d'une IA globale
left_x = 1.05
col_w = 5.48

# Carte 1 (Haut gauche) : Mention réelle manuscrite
b_c1 = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left_x), Inches(1.55), Inches(col_w), Inches(1.85))
b_c1.adjustments[0] = 0.03
b_c1.fill.solid()
b_c1.fill.fore_color.rgb = C_BG_SLIDE
b_c1.line.color.rgb = C_LINE
b_c1.line.width = Pt(0.75)

# Accent gauche
b_acc1 = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left_x), Inches(1.55), Inches(0.08), Inches(1.85))
b_acc1.adjustments[0] = 0.20
b_acc1.fill.solid()
b_acc1.fill.fore_color.rgb = C_NAVY
b_acc1.line.fill.background()

bx_c1_hdr = s23.shapes.add_textbox(Inches(left_x + 0.18), Inches(1.62), Inches(col_w - 0.28), Inches(0.30))
tf_c1_hdr = bx_c1_hdr.text_frame
tf_c1_hdr.margin_left = tf_c1_hdr.margin_top = tf_c1_hdr.margin_right = tf_c1_hdr.margin_bottom = 0
p1_h = tf_c1_hdr.paragraphs[0]
p1_h.text = "CAS RÉEL D'ARCHIVE : L'ABRÉVIATION MANUSCRITE"
r1_h = p1_h.runs[0]
r1_h.font.name = "Arial"
r1_h.font.size = Pt(10.0)
r1_h.font.bold = True
r1_h.font.color.rgb = C_NAVY

# Insertion image crop
if os.path.exists("img/crop_lachapelle_tight.png"):
    s23.shapes.add_picture("img/crop_lachapelle_tight.png", Inches(left_x + 0.18), Inches(1.95), width=Inches(2.45), height=Inches(1.25))

bx_c1_txt = s23.shapes.add_textbox(Inches(left_x + 2.75), Inches(1.95), Inches(col_w - 2.85), Inches(1.35))
tf_c1_txt = bx_c1_txt.text_frame
tf_c1_txt.word_wrap = True
tf_c1_txt.margin_left = tf_c1_txt.margin_top = tf_c1_txt.margin_right = tf_c1_txt.margin_bottom = 0

p_m1 = tf_c1_txt.paragraphs[0]
p_m1.text = "Mention : « Lachapelle /s/ AUBENAS »"
r_m1 = p_m1.runs[0]
r_m1.font.name = "Arial"
r_m1.font.size = Pt(9.0)
r_m1.font.bold = True
r_m1.font.color.rgb = C_NAVY

p_m2 = tf_c1_txt.add_paragraph()
p_m2.space_before = Pt(3)
r_m2 = p_m2.add_run()
r_m2.text = "• « /s/ » = abréviation locale classique pour « sous ».\\n• Écriture cursive inclinée du géomètre.\\n• Échec OCR conventionnel (lit « Aubenas » seul ou produit du bruit)."
r_m2.font.name = "Arial"
r_m2.font.size = Pt(8.0)
r_m2.font.color.rgb = C_DARK

# Carte 2 (Bas gauche) : Le piège d'une IA globale
b_c2 = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left_x), Inches(3.50), Inches(col_w), Inches(2.25))
b_c2.adjustments[0] = 0.03
b_c2.fill.solid()
b_c2.fill.fore_color.rgb = C_BG_SLIDE
b_c2.line.color.rgb = C_WARN
b_c2.line.width = Pt(1.0)

b_acc2 = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left_x), Inches(3.50), Inches(0.08), Inches(2.25))
b_acc2.adjustments[0] = 0.20
b_acc2.fill.solid()
b_acc2.fill.fore_color.rgb = C_WARN
b_acc2.line.fill.background()

bx_c2 = s23.shapes.add_textbox(Inches(left_x + 0.18), Inches(3.58), Inches(col_w - 0.28), Inches(2.10))
tf_c2 = bx_c2.text_frame
tf_c2.word_wrap = True
tf_c2.margin_left = tf_c2.margin_top = tf_c2.margin_right = tf_c2.margin_bottom = 0

p2_h = tf_c2.paragraphs[0]
p2_h.text = "LE PIÈGE D'UNE IA GLOBALE (BASE DE DONNÉES NATIONALE)"
r2_h = p2_h.runs[0]
r2_h.font.name = "Arial"
r2_h.font.size = Pt(10.0)
r2_h.font.bold = True
r2_h.font.color.rgb = C_WARN

items_c2 = [
    ("Confusion entre communes homonymes",
     "Une IA généraliste entraînée sur toute la France hésite parmi de multiples communes aux toponymes très proches :\\n"
     "  → Lachapelle-Graillouse (07310)\\n"
     "  → Lachapelle-sous-Chanéac (07310)\\n"
     "  → La Chapelle-en-Vercors (26420)\\n"
     "  → Autres « La Chapelle » hors secteur (sur-Erdre, etc.)"),
    ("Coût de calcul CPU prohibitif",
     "5 à 8 secondes par document avec un VLM sur processeur standard (~30x plus lent que l'OCR), paralysant le traitement des 23 600 dossiers.")
]
for ih, ib in items_c2:
    p_i = tf_c2.add_paragraph()
    p_i.space_before = Pt(4)
    r1 = p_i.add_run()
    r1.text = "• " + ih + " :\\n"
    r1.font.name = "Arial"
    r1.font.size = Pt(8.5)
    r1.font.bold = True
    r1.font.color.rgb = C_DARK
    r2 = p_i.add_run()
    r2.text = ib
    r2.font.name = "Arial"
    r2.font.size = Pt(8.0)
    r2.font.color.rgb = C_MUTED

# Colonne droite (5.48 in) : La solution retenue
right_x = 6.80

# Carte 3 (Haut droite) : Pourquoi coupler OCR/HTR et référentiel local
b_c3 = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(right_x), Inches(1.55), Inches(col_w), Inches(2.70))
b_c3.adjustments[0] = 0.03
b_c3.fill.solid()
b_c3.fill.fore_color.rgb = C_BG_SLIDE
b_c3.line.color.rgb = C_ACCENT
b_c3.line.width = Pt(1.0)

b_acc3 = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(right_x), Inches(1.55), Inches(0.08), Inches(2.70))
b_acc3.adjustments[0] = 0.20
b_acc3.fill.solid()
b_acc3.fill.fore_color.rgb = C_ACCENT
b_acc3.line.fill.background()

bx_c3 = s23.shapes.add_textbox(Inches(right_x + 0.18), Inches(1.62), Inches(col_w - 0.28), Inches(2.55))
tf_c3 = bx_c3.text_frame
tf_c3.word_wrap = True
tf_c3.margin_left = tf_c3.margin_top = tf_c3.margin_right = tf_c3.margin_bottom = 0

p3_h = tf_c3.paragraphs[0]
p3_h.text = "LA SOLUTION RETENUE : CHAÎNE HYBRIDE SÉCURISÉE"
r3_h = p3_h.runs[0]
r3_h.font.name = "Arial"
r3_h.font.size = Pt(10.0)
r3_h.font.bold = True
r3_h.font.color.rgb = C_ACCENT

items_c3 = [
    ("1. OCR / HTR déterministes et ultra-rapides (0.2 s)",
     "Transcription fidèle des caractères sans invention de texte, préservant l'intégrité de l'archive."),
    ("2. Verrouillage par le référentiel des 335 communes d'Ardèche",
     "Rapprochement automatique par distance de Levenshtein strictement borné au secteur du cabinet. Élimination garantie de toute hallucination hors département."),
    ("3. Arbitre VLM local (Ollama) réservé aux seuls cas ambigus",
     "Convoqué uniquement si le score de confiance est inférieur à 0.65 (< 15% des cas), sa lecture restant vérifiée par la base cadastrale.")
]
for ih, ib in items_c3:
    p_i = tf_c3.add_paragraph()
    p_i.space_before = Pt(4)
    r1 = p_i.add_run()
    r1.text = "• " + ih + " :\\n"
    r1.font.name = "Arial"
    r1.font.size = Pt(8.5)
    r1.font.bold = True
    r1.font.color.rgb = C_NAVY
    r2 = p_i.add_run()
    r2.text = ib
    r2.font.name = "Arial"
    r2.font.size = Pt(8.0)
    r2.font.color.rgb = C_DARK

# Carte 4 (Bas droite) : Résultat certifié
b_c4 = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(right_x), Inches(4.35), Inches(col_w), Inches(1.40))
b_c4.adjustments[0] = 0.03
b_c4.fill.solid()
b_c4.fill.fore_color.rgb = RGBColor(240, 253, 244)  # #F0FDF4 Vert très clair
b_c4.line.color.rgb = C_SUCCESS
b_c4.line.width = Pt(1.2)

bx_c4 = s23.shapes.add_textbox(Inches(right_x + 0.18), Inches(4.45), Inches(col_w - 0.28), Inches(1.20))
tf_c4 = bx_c4.text_frame
tf_c4.word_wrap = True
tf_c4.margin_left = tf_c4.margin_top = tf_c4.margin_right = tf_c4.margin_bottom = 0

p4_1 = tf_c4.paragraphs[0]
p4_1.text = "RÉSULTAT APRÈS ARBITRAGE LOCAL ET VÉRIFICATION :"
r4_1 = p4_1.runs[0]
r4_1.font.name = "Arial"
r4_1.font.size = Pt(8.5)
r4_1.font.bold = True
r4_1.font.color.rgb = RGBColor(22, 101, 52)  # Vert foncé

p4_2 = tf_c4.add_paragraph()
p4_2.space_before = Pt(2)
r4_2 = p4_2.add_run()
r4_2.text = "« La Chapelle-sous-Aubenas »"
r4_2.font.name = "Arial"
r4_2.font.size = Pt(13.0)
r4_2.font.bold = True
r4_2.font.color.rgb = RGBColor(21, 128, 61)

p4_3 = tf_c4.add_paragraph()
p4_3.space_before = Pt(2)
r4_3 = p4_3.add_run()
r4_3.text = "Code INSEE officiel : 07058  |  Zéro hallucination  |  100% local (Ollama)"
r4_3.font.name = "Arial"
r4_3.font.size = Pt(8.0)
r4_3.font.bold = True
r4_3.font.color.rgb = C_NAVY

# Bandeau inférieur officiel
b_syn23 = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.90), Inches(11.23), Inches(0.80))
b_syn23.adjustments[0] = 0.03
b_syn23.fill.solid()
b_syn23.fill.fore_color.rgb = C_CALLOUT_BG
b_syn23.line.color.rgb = C_ACCENT
b_syn23.line.width = Pt(1.0)

bx_s23 = s23.shapes.add_textbox(Inches(1.20), Inches(5.95), Inches(10.93), Inches(0.70))
tf_s23 = bx_s23.text_frame
tf_s23.word_wrap = True
tf_s23.margin_left = tf_s23.margin_top = tf_s23.margin_right = tf_s23.margin_bottom = 0

ps23_1 = tf_s23.paragraphs[0]
ps23_1.text = "PRINCIPE FONDAMENTAL DE LA DÉMARCHE INGÉNIEUR :"
rs23_1 = ps23_1.runs[0]
rs23_1.font.name = "Arial"
rs23_1.font.size = Pt(9.5)
rs23_1.font.bold = True
rs23_1.font.color.rgb = C_ACCENT

ps23_2 = tf_s23.add_paragraph()
ps23_2.space_before = Pt(2)
rs23_2 = ps23_2.add_run()
rs23_2.text = "L'IA générative n'est pas un substitut universel. L'association de moteurs OCR/HTR rapides et d'un référentiel cadastral local garantit la sécurité juridique et la performance sur grand volume."
rs23_2.font.name = "Arial"
rs23_2.font.size = Pt(8.5)
rs23_2.font.color.rgb = C_DARK

# NOTES DU PRÉSENTATEUR (Mots-clés pour l'oral)
s23_notes = s23.notes_slide.notes_text_frame
s23_notes.text = (
    "MOTS-CLÉS À DIRE À L'ORAL (Slide 23) :\\n"
    "• Anticiper la question du jury : Pourquoi ne pas tout envoyer à une IA générative (VLM) ?\\n"
    "• 1. Risque d'erreur toponymique : Une IA globale a une base nationale. Sur l'abréviation « Lachapelle /s/ AUBENAS », "
    "elle peut hésiter avec Lachapelle-Graillouse (07), Lachapelle-sous-Chanéac (07) ou La Chapelle-en-Vercors (26), voire inventer une commune.\\n"
    "• 2. Vitesse et ressource : L'inférence VLM prend 5 à 8 secondes sur CPU, impensable sur 23 600 dossiers. OCR et TrOCR prennent 0.2 s.\\n"
    "• 3. Notre solution : Des moteurs OCR/HTR rapides qui lisent sans inventer + filtrage immédiat par distance de Levenshtein sur les 335 communes d'Ardèche.\\n"
    "• 4. Rôle de l'IA (VLM) : Elle intervient uniquement en arbitre de secours sur les cas ambigus (score < 0.65), et sa lecture reste verrouillée par la base locale. Résultat : zéro hallucination et conformité cadastrale absolue."
)

'''

text = text[:s23_start] + new_slide23 + text[s23_end:]

with open('generate_final_180s_deck.py', 'w', encoding='utf-8') as f:
    f.write(text)

print("Slide 23 updated and notes added successfully!")
