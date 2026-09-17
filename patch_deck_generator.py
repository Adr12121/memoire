import re

# Lire le fichier source
with open("generate_final_180s_deck.py", "r", encoding="utf-8") as f:
    content = f.read()

# 1. Mise à jour de CHAPTERS_DATA
old_chapters = """CHAPTERS_DATA = [
    {
        "short": "1. Introduction",
        "full": "Introduction et Contexte du Projet",
        "badge": "01",
        "subtitle": "Structure d'accueil GEO-SIAPP, cadre réglementaire et problématique foncière",
        "subs": ["1.1 Structure GEO-SIAPP", "1.2 Cadre réglementaire", "1.3 Contexte & Antériorité", "1.4 Problématique foncière"]
    },
    {
        "short": "2. Données & Métier",
        "full": "Analyse Métier, de la Donnée et Enjeux Documentaires",
        "badge": "02",
        "subtitle": "Distinction cadastre et actes, spécifications API Géofoncier et filiation parcellaire",
        "subs": ["2.1 Cadre juridique", "2.2 Spécifications API", "2.3 Typologie des archives", "2.4 Filiation parcellaire"]
    },
    {
        "short": "3. État de l'Art",
        "full": "État de l'Art des Technologies d'Analyse Documentaire",
        "badge": "03",
        "subtitle": "Reconnaissance OCR et HTR, extraction d'entités, modèles Vision-Langage et synthèse",
        "subs": ["3.1 OCR imprimé", "3.2 HTR manuscrit", "3.3 Extraction NER", "3.4 Arbitrage VLM", "3.5 Synthèse comparative"]
    },
    {
        "short": "4. Architecture",
        "full": "Architecture Logicielle et Pipeline de Traitement",
        "badge": "04",
        "subtitle": "Chaîne modulaire, segmentation spatiale YOLOv8, extraction textuelle et arbitrage local",
        "subs": ["4.1 Architecture logicielle", "4.2 Segmentation YOLO", "4.3 Extraction GLiNER", "4.4 Arbitrage VLM"]
    },
    {
        "short": "5. Fiabilisation",
        "full": "Fiabilisation des Données et Interface Opérateur",
        "badge": "05",
        "subtitle": "Validation contradictoire humaine Streamlit, règles métier et contrôle spatial",
        "subs": ["5.1 Interface Streamlit", "5.2 Règles métier", "5.3 Contrôle spatial"]
    },
    {
        "short": "6. Géofoncier",
        "full": "Intégration Géofoncier et Déploiement Opérationnel",
        "badge": "06",
        "subtitle": "Démonstration en conditions réelles, protocole d'injection et expérimentation Prades",
        "subs": ["6.1 Démonstration vidéo", "6.2 Protocole API", "6.3 Validation Prades"]
    },
    {
        "short": "7. Limites",
        "full": "Analyse des Limites et Perspectives d'Évolution",
        "badge": "07",
        "subtitle": "Diagnostic des modes de défaillance, altérations physiques et feuille de route technique",
        "subs": ["7.1 Limites techniques", "7.2 Perspectives d'évolution"]
    },
    {
        "short": "8. Conclusion",
        "full": "Conclusion Générale et Bilan du Projet",
        "badge": "08",
        "subtitle": "Bilan opérationnel pour GEO-SIAPP, souveraineté et compétences d'ingénieur INSA",
        "subs": ["8.1 Bilan du projet", "8.2 Remerciements & Échanges"]
    }
]"""

new_chapters = """CHAPTERS_DATA = [
    {
        "short": "1. Introduction",
        "full": "Introduction et contexte du projet",
        "badge": "1",
        "subtitle": "Structure d'accueil GEO-SIAPP, cadre réglementaire et problématique foncière",
        "subs": ["1.1 Structure GEO-SIAPP", "1.2 Cadre réglementaire", "1.3 Contexte & Antériorité", "1.4 Problématique foncière"]
    },
    {
        "short": "2. Données & Métier",
        "full": "Analyse métier et enjeux documentaires",
        "badge": "2",
        "subtitle": "Distinction cadastre et actes, spécifications API Géofoncier et filiation parcellaire",
        "subs": ["2.1 Cadre juridique", "2.2 Spécifications API", "2.3 Typologie des archives", "2.4 Filiation parcellaire"]
    },
    {
        "short": "3. État de l'Art",
        "full": "État de l'art des technologies d'analyse",
        "badge": "3",
        "subtitle": "Reconnaissance OCR et HTR, extraction d'entités, modèles Vision-Langage et synthèse",
        "subs": ["3.1 OCR imprimé", "3.2 HTR manuscrit", "3.3 Extraction NER", "3.4 Arbitrage VLM", "3.5 Synthèse comparative"]
    },
    {
        "short": "4. Architecture",
        "full": "Architecture logicielle et chaîne de traitement",
        "badge": "4",
        "subtitle": "Chaîne modulaire, segmentation spatiale YOLOv8, extraction textuelle et arbitrage local",
        "subs": ["4.1 Architecture logicielle", "4.2 Segmentation YOLO", "4.3 Extraction GLiNER", "4.4 Arbitrage VLM"]
    },
    {
        "short": "5. Fiabilisation",
        "full": "Fiabilisation des données et interface opérateur",
        "badge": "5",
        "subtitle": "Validation contradictoire humaine Streamlit, règles métier et contrôle spatial",
        "subs": ["5.1 Interface Streamlit", "5.2 Règles métier", "5.3 Contrôle spatial"]
    },
    {
        "short": "6. Géofoncier",
        "full": "Intégration Géofoncier et expérimentation",
        "badge": "6",
        "subtitle": "Démonstration en conditions réelles, protocole d'injection et expérimentation Prades",
        "subs": ["6.1 Démonstration vidéo", "6.2 Protocole API", "6.3 Validation Prades"]
    },
    {
        "short": "7. Limites",
        "full": "Analyse des limites et perspectives d'évolution",
        "badge": "7",
        "subtitle": "Diagnostic des modes de défaillance, altérations physiques et feuille de route technique",
        "subs": ["7.1 Limites techniques", "7.2 Perspectives d'évolution"]
    },
    {
        "short": "8. Conclusion",
        "full": "Conclusion générale et bilan du projet",
        "badge": "8",
        "subtitle": "Bilan opérationnel pour GEO-SIAPP, traitement local et compétences d'ingénieur INSA",
        "subs": ["8.1 Bilan du projet", "8.2 Remerciements & Échanges"]
    }
]"""

assert old_chapters in content, "old_chapters non trouvé !"
content = content.replace(old_chapters, new_chapters)

# 2. Mise à jour de create_transition_slide
old_trans = """    # Badge Chapitre centré
    badge_w = 2.00
    badge_h = 0.38
    badge_x = 0.70 + (11.93 - badge_w) / 2.0
    badge_y = 3.00
    
    b_badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(badge_x), Inches(badge_y), Inches(badge_w), Inches(badge_h))
    b_badge.adjustments[0] = 0.15
    b_badge.fill.solid()
    b_badge.fill.fore_color.rgb = C_NAVY
    b_badge.line.fill.background()
    tf_b = b_badge.text_frame
    tf_b.margin_left = tf_b.margin_top = tf_b.margin_right = tf_b.margin_bottom = 0
    pb = tf_b.paragraphs[0]
    pb.text = f"CHAPITRE {CHAPTERS_DATA[chap_idx]['badge']}"
    pb.alignment = PP_ALIGN.CENTER
    rb = pb.runs[0]
    rb.font.name = "Arial"
    rb.font.size = Pt(11.0)
    rb.font.bold = True
    rb.font.color.rgb = C_WHITE"""

new_trans = """    # Badge Chapitre centré
    badge_w = 2.20
    badge_h = 0.38
    badge_x = 0.70 + (11.93 - badge_w) / 2.0
    badge_y = 3.00
    
    b_badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(badge_x), Inches(badge_y), Inches(badge_w), Inches(badge_h))
    b_badge.adjustments[0] = 0.15
    b_badge.fill.solid()
    b_badge.fill.fore_color.rgb = C_NAVY
    b_badge.line.fill.background()
    tf_b = b_badge.text_frame
    tf_b.margin_left = tf_b.margin_top = tf_b.margin_right = tf_b.margin_bottom = 0
    pb = tf_b.paragraphs[0]
    pb.text = f"Chapitre {chap_idx + 1}"
    pb.alignment = PP_ALIGN.CENTER
    rb = pb.runs[0]
    rb.font.name = "Arial"
    rb.font.size = Pt(11.5)
    rb.font.bold = True
    rb.font.color.rgb = C_WHITE"""

assert old_trans in content, "old_trans non trouvé !"
content = content.replace(old_trans, new_trans)

# 3. Refonte Slide 2 (Sommaire) : Nouvelle disposition sans majuscules superflues, badges Chapitre 1 à 8
old_s2 = """# ==================== SLIDE 2 : SOMMAIRE ÉPURÉ ====================
print("Génération Slide 2 (Sommaire)...")
s2, c2 = init_standard_slide(-1, -1, 2, "Plan de la présentation : Sommaire")

cw2 = (11.93 - 0.50 - 0.40) / 2.0
ch2 = (5.80 - 0.50 - 3 * 0.18) / 4.0
sx2 = 0.70 + 0.25
sy2 = 1.35 + 0.25

for idx, chap in enumerate(CHAPTERS_DATA):
    col = idx // 4
    row = idx % 4
    x = sx2 + col * (cw2 + 0.40)
    y = sy2 + row * (ch2 + 0.18)
    
    b_card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(cw2), Inches(ch2))
    b_card.adjustments[0] = 0.04
    b_card.fill.solid()
    b_card.fill.fore_color.rgb = C_BG_SLIDE
    b_card.line.color.rgb = C_BORDER_TAB
    b_card.line.width = Pt(0.75)
    
    badge_w2 = 0.80
    badge_h2 = ch2 - 0.24
    badge_shp = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x + 0.12), Inches(y + 0.12), Inches(badge_w2), Inches(badge_h2))
    badge_shp.adjustments[0] = 0.08
    badge_shp.fill.solid()
    badge_shp.fill.fore_color.rgb = C_NAVY
    badge_shp.line.fill.background()
    tf_b = badge_shp.text_frame
    tf_b.margin_left = tf_b.margin_top = tf_b.margin_right = tf_b.margin_bottom = 0
    p_b = tf_b.paragraphs[0]
    p_b.alignment = PP_ALIGN.CENTER
    p_b.text = chap['badge']
    r_b = p_b.runs[0]
    r_b.font.name = "Arial"
    r_b.font.size = Pt(14.0)
    r_b.font.bold = True
    r_b.font.color.rgb = C_WHITE
    
    bx_ht = s2.shapes.add_textbox(Inches(x + 1.10), Inches(y + 0.15), Inches(cw2 - 1.25), Inches(ch2 - 0.30))
    tf_ht = bx_ht.text_frame
    tf_ht.word_wrap = True
    tf_ht.margin_left = tf_ht.margin_top = tf_ht.margin_right = tf_ht.margin_bottom = 0
    p_ht = tf_ht.paragraphs[0]
    p_ht.text = chap["full"]
    r_ht = p_ht.runs[0]
    r_ht.font.name = "Arial"
    r_ht.font.size = Pt(12.5)
    r_ht.font.bold = True
    r_ht.font.color.rgb = C_NAVY"""

new_s2 = """# ==================== SLIDE 2 : SOMMAIRE ÉPURÉ ====================
print("Génération Slide 2 (Sommaire élégant et aéré)...")
s2, c2 = init_standard_slide(-1, -1, 2, "Plan de la présentation : Sommaire")

# Nouvelle disposition : 2 colonnes équilibrées de 4 fiches structurées
cw2 = 5.50
ch2 = 1.10
cg2_y = 0.18
left_x = 1.05
right_x = 6.78
start_y = 1.60

for idx, chap in enumerate(CHAPTERS_DATA):
    col = idx // 4
    row = idx % 4
    cx = left_x if col == 0 else right_x
    cy = start_y + row * (ch2 + cg2_y)
    
    b_card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(cy), Inches(cw2), Inches(ch2))
    b_card.adjustments[0] = 0.03
    b_card.fill.solid()
    b_card.fill.fore_color.rgb = C_BG_SLIDE
    b_card.line.color.rgb = C_LINE
    b_card.line.width = Pt(0.75)
    
    # Pilule élégante Chapitre X
    pill_w = 1.30
    pill_h = 0.32
    pill_shp = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx + 0.18), Inches(cy + 0.16), Inches(pill_w), Inches(pill_h))
    pill_shp.adjustments[0] = 0.18
    pill_shp.fill.solid()
    pill_shp.fill.fore_color.rgb = C_NAVY
    pill_shp.line.fill.background()
    tf_p = pill_shp.text_frame
    tf_p.margin_left = tf_p.margin_top = tf_p.margin_right = tf_p.margin_bottom = 0
    p_p = tf_p.paragraphs[0]
    p_p.alignment = PP_ALIGN.CENTER
    p_p.text = f"Chapitre {idx + 1}"
    r_p = p_p.runs[0]
    r_p.font.name = "Arial"
    r_p.font.size = Pt(8.5)
    r_p.font.bold = True
    r_p.font.color.rgb = C_WHITE
    
    # Titre du chapitre en minuscules soignées (Sentence Case)
    bx_tit = s2.shapes.add_textbox(Inches(cx + 1.60), Inches(cy + 0.12), Inches(cw2 - 1.75), Inches(0.38))
    tf_tit = bx_tit.text_frame
    tf_tit.word_wrap = True
    tf_tit.margin_left = tf_tit.margin_top = tf_tit.margin_right = tf_tit.margin_bottom = 0
    p_tit = tf_tit.paragraphs[0]
    p_tit.text = chap["full"]
    r_tit = p_tit.runs[0]
    r_tit.font.name = "Arial"
    r_tit.font.size = Pt(11.5)
    r_tit.font.bold = True
    r_tit.font.color.rgb = C_NAVY
    
    # Ligne descriptive sobre
    bx_sub = s2.shapes.add_textbox(Inches(cx + 0.18), Inches(cy + 0.58), Inches(cw2 - 0.36), Inches(0.44))
    tf_sub = bx_sub.text_frame
    tf_sub.word_wrap = True
    tf_sub.margin_left = tf_sub.margin_top = tf_sub.margin_right = tf_sub.margin_bottom = 0
    p_sub = tf_sub.paragraphs[0]
    p_sub.text = chap["subtitle"]
    r_sub = p_sub.runs[0]
    r_sub.font.name = "Arial"
    r_sub.font.size = Pt(8.0)
    r_sub.font.color.rgb = C_MUTED"""

assert old_s2 in content, "old_s2 non trouvé !"
content = content.replace(old_s2, new_s2)

# 4. Slide 3 : Transition Chapitre 1
content = content.replace(
    'title="01 - Introduction et Contexte du Projet",',
    'title="Chapitre 1 : Introduction et contexte du projet",'
)

# 5. Slide 6 : Colonnes sans ALL-CAPS
old_s6_cols = """cols_s6 = [
    ("CADRE JURIDIQUE STRICT",
     "« Bornage sur bornage ne vaut »",
     C_NAVY,
     [
         ("Article 646 Code civil", "Obligation légale de rechercher systématiquement les actes antérieurs."),
         ("Continuité des limites", "Rétablissement obligatoire de la ligne séparative d'origine."),
         ("Opposabilité légale", "Tout nouvel acte dressé sans antériorité est juridiquement caduc."),
         ("Sécurisation du client", "Garantie absolue contre les contestations et litiges de voisinage.")
     ]),
    ("RÉALITÉ OPÉRATIONNELLE",
     "Recherche physique manuelle",
     C_WARN,
     [
         ("Cartons d'Aubenas", "Fouille manuelle dans des centaines de boîtes d'archives physiques."),
         ("Temps perdu", "15 à 30 minutes requises par dossier, sans certitude de résultat immédiat."),
         ("Fragilité matérielle", "Documents de 1959 à 2007 usés par les manipulations répétées."),
         ("Frein de production", "Retard direct sur le démarrage des nouveaux chantiers de terrain.")
     ]),
    ("ENJEU DE L'ENTREPRISE",
     "23 600 dossiers dormants",
     C_ACCENT,
     [
         ("Volume massif", "23 600 dossiers papier non indexés au siège d'Aubenas."),
         ("Mémoire technique", "Transmission du savoir issu de 5 cabinets de géomètres prédécesseurs."),
         ("Échanges confraternels", "Demandes régulières de confrères nécessitant des recherches rapides."),
         ("Objectif du PFE", "Extraction automatique et publication de pastilles sur Géofoncier.")
     ])
]"""

new_s6_cols = """cols_s6 = [
    ("Cadre juridique strict",
     "« Bornage sur bornage ne vaut »",
     C_NAVY,
     [
         ("Article 646 Code civil", "Obligation légale de rechercher systématiquement les actes antérieurs."),
         ("Continuité des limites", "Rétablissement obligatoire de la ligne séparative d'origine."),
         ("Opposabilité légale", "Tout nouvel acte dressé sans antériorité est juridiquement caduc."),
         ("Sécurisation du client", "Garantie absolue contre les contestations et litiges de voisinage.")
     ]),
    ("Réalité opérationnelle",
     "Recherche physique manuelle",
     C_WARN,
     [
         ("Cartons d'Aubenas", "Fouille manuelle dans des centaines de boîtes d'archives physiques."),
         ("Temps perdu", "15 à 30 minutes requises par dossier, sans certitude de résultat immédiat."),
         ("Fragilité matérielle", "Documents de 1959 à 2007 usés par les manipulations répétées."),
         ("Frein de production", "Retard direct sur le démarrage des nouveaux chantiers de terrain.")
     ]),
    ("Enjeu pour le cabinet SIAPP",
     "23 600 dossiers dormants",
     C_ACCENT,
     [
         ("Volume massif", "23 600 dossiers papier non indexés au siège d'Aubenas."),
         ("Mémoire technique", "Transmission du savoir issu de 5 cabinets de géomètres prédécesseurs."),
         ("Échanges confraternels", "Demandes régulières de confrères nécessitant des recherches rapides."),
         ("Objectif du PFE", "Extraction automatique et publication de pastilles sur Géofoncier.")
     ])
]"""

assert old_s6_cols in content, "old_s6_cols non trouvé !"
content = content.replace(old_s6_cols, new_s6_cols)

# 6. Slide 7 (1.4) : Problématique SANS annonce prématurée du pipeline (reprise de la phrase du mémoire)
old_s7 = """# Slide 7 (1.4) : Problématique foncière et démarche méthodologique
s7, c7 = init_standard_slide(0, 3, 7, "1.4 Problématique foncière et démarche méthodologique")
add_two_column_content(
    s7,
    items=[
        ("Hétérogénéité des fonds (1959–2007)",
         "Supports très variés (calques, registres manuscrits, tapuscrits), encres dégradées, pliures d'usure et scans de faible contraste."),
        ("Échec des moteurs OCR classiques",
         "Sur l'écriture manuscrite des registres, les OCR standards dépassent 30% de taux d'erreur de caractères (CER)."),
        ("Souveraineté et secret professionnel",
         "Données foncières nominatives assujetties au secret professionnel : exécution 100% locale sur PC Windows de bureau."),
        ("Problématique centrale du PFE",
         "Comment concevoir une chaîne logicielle automatisée capable d'extraire avec fiabilité les données d'archives variées en local ?")
    ],
    img_path="img/pipeline_6_etapes_clean.png",
    img_caption="Démarche méthodologique en 6 étapes retenue pour le traitement des archives",
    callout_title="DÉMARCHE OPÉRATIONNELLE EN 6 ÉTAPES",
    callout_text="1. Classification  →  2. Découpage YOLOv8  →  3. Lecture OCR/HTR  →  4. Cohérence GLiNER  →  5. Contrôle Streamlit  →  6. Export Géofoncier."
)"""

new_s7 = """# Slide 7 (1.4) : Problématique et contraintes techniques (Phrase exacte du mémoire)
print("Génération Slide 7 (Problématique pure et contraintes sans annonce prématurée)...")
s7, c7 = init_standard_slide(0, 3, 7, "1.4 Problématique et contraintes techniques du projet")

# Colonne gauche : 3 contraintes majeures
left_x7 = 1.05
left_w7 = 5.50
start_y7 = 1.60

constraints_s7 = [
    ("Hétérogénéité des archives physiques (1959–2007)",
     "Près de 50 ans d'histoire couvrant des supports très variés : plans sur calque, registres manuscrits, formulaires dactylographiés, encres pâlies et pliures d'usure."),
    ("Limites des moteurs de lecture conventionnels",
     "Échec des logiciels de reconnaissance de texte sur l'écriture cursive des géomètres, les abréviations locales et les documents penchés (taux d'erreur inexploitable)."),
    ("Confidentialité des données et exécution locale",
     "Les actes fonciers comportent des données personnelles assujetties au secret professionnel : obligation stricte d'exécuter l'outil localement sur station CPU, sans cloud.")
]

gap7 = 0.12
card_h7 = (4.00 - 2 * gap7) / 3.0
for ci, (c_hd, c_bd) in enumerate(constraints_s7):
    cy = start_y7 + ci * (card_h7 + gap7)
    b_c = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left_x7), Inches(cy), Inches(left_w7), Inches(card_h7))
    b_c.adjustments[0] = 0.03
    b_c.fill.solid()
    b_c.fill.fore_color.rgb = C_BG_SLIDE
    b_c.line.color.rgb = C_LINE
    b_c.line.width = Pt(0.75)
    
    bx_c = s7.shapes.add_textbox(Inches(left_x7 + 0.16), Inches(cy + 0.08), Inches(left_w7 - 0.32), Inches(card_h7 - 0.16))
    tf_c = bx_c.text_frame
    tf_c.word_wrap = True
    tf_c.margin_left = tf_c.margin_top = tf_c.margin_right = tf_c.margin_bottom = 0
    
    p1 = tf_c.paragraphs[0]
    p1.text = "• " + c_hd + " :\n"
    r1 = p1.runs[0]
    r1.font.name = "Arial"
    r1.font.size = Pt(10.0)
    r1.font.bold = True
    r1.font.color.rgb = C_NAVY
    
    p2 = tf_c.add_paragraph()
    p2.space_before = Pt(2)
    p2.text = c_bd
    r2 = p2.runs[0]
    r2.font.name = "Arial"
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = C_DARK

# Colonne droite : Registre manuscrit authentique illustrant la difficulté réelle
frame_r7 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.78), Inches(start_y7), Inches(5.50), Inches(4.00))
frame_r7.adjustments[0] = 0.03
frame_r7.fill.solid()
frame_r7.fill.fore_color.rgb = C_BG_SLIDE
frame_r7.line.color.rgb = C_LINE
frame_r7.line.width = Pt(0.75)
add_fitted_picture(s7, "img/registre_manuscrit.jpg", 6.93, start_y7 + 0.10, 5.20, 3.80, caption="Fonds d'archives physiques du cabinet : écriture manuscrite et abréviations (Mémoire, Figure 1.5)")

# Grand bandeau inférieur mettant en valeur la phrase exacte du mémoire
b_prob7 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.80), Inches(11.23), Inches(0.85))
b_prob7.adjustments[0] = 0.03
b_prob7.fill.solid()
b_prob7.fill.fore_color.rgb = C_CALLOUT_BG
b_prob7.line.color.rgb = C_ACCENT
b_prob7.line.width = Pt(1.0)

bx_prob = s7.shapes.add_textbox(Inches(1.20), Inches(5.86), Inches(10.93), Inches(0.72))
tf_prob = bx_prob.text_frame
tf_prob.word_wrap = True
tf_prob.margin_left = tf_prob.margin_top = tf_prob.margin_right = tf_prob.margin_bottom = 0

pp1 = tf_prob.paragraphs[0]
pp1.text = "PROBLÉMATIQUE CENTRALE DU MÉMOIRE :"
rp1 = pp1.runs[0]
rp1.font.name = "Arial"
rp1.font.size = Pt(9.5)
rp1.font.bold = True
rp1.font.color.rgb = C_ACCENT

pp2 = tf_prob.add_paragraph()
pp2.space_before = Pt(3)
pp2.text = "« Dans quelle mesure est-il possible de concevoir un outil logiciel capable d'extraire avec fiabilité les données d'archives foncières très variées, tout en fonctionnant localement sur un ordinateur de bureau ? »"
rp2 = pp2.runs[0]
rp2.font.name = "Arial"
rp2.font.size = Pt(10.5)
rp2.font.bold = True
rp2.font.color.rgb = C_NAVY"""

assert old_s7 in content, "old_s7 non trouvé !"
content = content.replace(old_s7, new_s7)

# 7. Slide 8 : Transition Chapitre 2
content = content.replace(
    'title="02 - Analyse Métier, de la Donnée et Enjeux Documentaires",',
    'title="Chapitre 2 : Analyse métier et enjeux documentaires",'
)

# 8. Slide 10 (2.2) : Insertion de la figure du mémoire arborescence_versement_geofoncier.png
old_s10_img = """frame_r10 = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.78), Inches(1.60), Inches(5.50), Inches(5.05))
frame_r10.adjustments[0] = 0.03
frame_r10.fill.solid()
frame_r10.fill.fore_color.rgb = C_BG_SLIDE
frame_r10.line.color.rgb = C_LINE
frame_r10.line.width = Pt(0.75)
add_fitted_picture(s10, "img/Chap2_pastilles_enhanced.jpg", 6.93, 1.75, 5.20, 4.75, caption="Capture d'écran du portail GéofoncierEXPERT avec pastilles géolocalisées (Mémoire, Figure 2.1)")"""

new_s10_img = """frame_r10 = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.78), Inches(1.60), Inches(5.50), Inches(5.05))
frame_r10.adjustments[0] = 0.03
frame_r10.fill.solid()
frame_r10.fill.fore_color.rgb = C_BG_SLIDE
frame_r10.line.color.rgb = C_LINE
frame_r10.line.width = Pt(0.75)
add_fitted_picture(s10, "img/arborescence_versement_geofoncier.png", 6.93, 1.75, 5.20, 4.75, caption="Structure JSON obligatoire à renseigner pour le versement au RFU (Mémoire, Chapitre 2)")"""

assert old_s10_img in content, "old_s10_img non trouvé !"
content = content.replace(old_s10_img, new_s10_img)

# 9. Slide 13 : Transition Chapitre 3
content = content.replace(
    'title="03 - État de l\'Art des Technologies d\'Analyse Documentaire",',
    'title="Chapitre 3 : État de l\'art des technologies d\'analyse documentaire",'
)

# 10. Variation de mise en page dans l'état de l'art (Slides 14-17)
old_s14_to_s17 = """# Slide 14 (3.1) : OCR imprimé
s14, c14 = init_standard_slide(2, 0, 14, "3.1 Reconnaissance de texte imprimé : Les moteurs OCR classiques")
add_two_column_content(
    s14,
    items=[
        ("Principe d'analyse par caractères isolés",
         "Détection préalable des blocs réguliers, puis segmentation et classification typographique caractère par caractère."),
        ("Tesseract v5 (Smith, 2007) : Léger sur CPU",
         "Moteur historique très économe en ressources. Efficace sur textes imprimés propres, inopérant sur écriture manuscrite."),
        ("EasyOCR (JaidedAI) : Détection DBNet + CRNN",
         "Architecture combinant détection spatiale CRAFT et extracteur récurrent CRNN. Vitesse élevée (< 0.5s par cartouche).")
    ],
    img_path="img/baek_ocr_pipeline_crop.png",
    img_caption="Chaîne de traitement standard de l'OCR : détection spatiale puis transcription par zone",
    callout_title="ÉVALUATION DU TEXTE IMPRIMÉ",
    callout_text="Efficace sur les cartouches de formulaires Cerfa récents. En revanche, le taux d'erreur dépasse 30% sur écriture manuscrite."
)

# Slide 15 (3.2) : HTR manuscrit
s15, c15 = init_standard_slide(2, 1, 15, "3.2 Reconnaissance d'écriture manuscrite : L'approche HTR")
add_two_column_content(
    s15,
    items=[
        ("Taxonomie des systèmes (Alkendi, 2024)",
         "Distinction entre scans statiques hors ligne et saisie en ligne. L'écriture cursive impose une analyse séquentielle continue."),
        ("TrOCR (Li et al., Microsoft Research, 2021)",
         "Vision Transformer bout-en-bout couplant un encodeur d'image (DeiT) et un décodeur linguistique autoregressif (RoBERTa)."),
        ("PyLaia et corpus spécialisés HTR-United",
         "Modèles combinant CNN et BLSTM entraînés sur des registres d'archives historiques françaises pour déchiffrer les graphies anciennes.")
    ],
    img_path="img/fig_alkendi_p3_1.jpeg",
    img_caption="Taxonomie des systèmes de reconnaissance de texte : distinction imprimé / manuscrit (Alkendi et al., 2024)",
    callout_title="ÉVALUATION DU TEXTE MANUSCRIT",
    callout_text="TrOCR atteint un taux d'erreur de caractères (CER) inférieur à 6% sur l'écriture manuscrite, sous réserve d'un découpage précis des lignes."
)

# Slide 16 (3.3) : NER
s16, c16 = init_standard_slide(2, 2, 16, "3.3 Extraction d'entités nommées (NER) : GLiNER vs LayoutLM")
add_two_column_content(
    s16,
    items=[
        ("GLiNER (Zaratiana et al., 2023)",
         "Modèle bi-encodeur BERT compact (340M paramètres). Extraction zero-shot d'entités arbitraires sans réentraînement préalable."),
        ("LayoutLMv3 (Xu et al., Microsoft Research, 2020)",
         "Modèle Document AI multimodal intégrant texte, coordonnées 2D et image. Conçu pour les formulaires imprimés réguliers."),
        ("Comportement face aux textes dégradés",
         "LayoutLM chute à 88% dès que l'OCR contient des fautes. À l'inverse, GLiNER tolère le texte bruité grâce au contexte global.")
    ],
    img_path="img/gliner_prompt_render.png",
    img_caption="Fonctionnement zero-shot de GLiNER associant texte extrait et entités cibles au runtime",
    callout_title="ÉVALUATION DES MODÈLES D'EXTRACTION",
    callout_text="GLiNER permet d'extraire dynamiquement les 5 entités clés (commune, date, nature, contenance, référence) avec score de confiance."
)

# Slide 17 (3.4) : VLM
s17, c17 = init_standard_slide(2, 3, 17, "3.4 Modèles Vision-Langage (VLM) : Arbitrage multimodal local")
add_two_column_content(
    s17,
    items=[
        ("Paradigme Vision-Langage (LLaVA, MiniCPM)",
         "Couplage d'un encodeur visuel (CLIP ViT) et d'un LLM autoregressif. L'image est encodée comme une séquence de tokens visuels."),
        ("Raisonnement visuel et compréhension de scène",
         "Interprétation directe du document sans découpage : capacité à relier une mention manuscrite avec sa position sur la feuille."),
        ("Compromis latence et contrainte de souveraineté",
         "Temps d'inférence de ~5s à 8s par vue sur CPU de bureau. Ce coût impose de réserver le VLM en tant qu'arbitre de second niveau.")
    ],
    img_path="img/llava_arch_render.png",
    img_caption="Architecture des modèles Vision-Langage : projection des patchs visuels dans le LLM (Liu et al., 2023)",
    callout_title="SOUVERAINETÉ ET SECRET PROFESSIONNEL",
    callout_text="Exécution 100% locale via le moteur Ollama : aucune donnée nominative d'archive n'est transmise vers des serveurs externes."
)"""

new_s14_to_s17 = """# Slide 14 (3.1) : OCR imprimé (Disposition inversée : Image à gauche, fiches techniques à droite)
print("Génération Slide 14 (OCR imprimé - Image à gauche)...")
s14, c14 = init_standard_slide(2, 0, 14, "3.1 Reconnaissance de texte imprimé : Les moteurs OCR classiques")

# Image à gauche
frame_l14 = s14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(1.60), Inches(5.30), Inches(5.05))
frame_l14.adjustments[0] = 0.03
frame_l14.fill.solid()
frame_l14.fill.fore_color.rgb = C_BG_SLIDE
frame_l14.line.color.rgb = C_LINE
frame_l14.line.width = Pt(0.75)
add_fitted_picture(s14, "img/baek_ocr_pipeline_crop.png", 1.20, 1.75, 5.00, 4.75, caption="Chaîne standard de l'OCR imprimé : détection puis transcription (Baek et al., 2019)")

# Fiches techniques à droite
rx14 = 6.58
rw14 = 5.70
items_ocr = [
    ("Principe d'analyse par caractères isolés",
     "Segmentation géométrique des blocs de texte, puis classification typographique caractère par caractère."),
    ("Tesseract v5 (Smith, 2007) : Léger sur CPU",
     "Moteur open-source très économe en mémoire vive. Performant sur imprimés propres, mais inadapté au manuscrit."),
    ("EasyOCR (JaidedAI) : Détection CRAFT + CRNN",
     "Réseau neuronal combinant détection spatiale et extracteur récurrent, rapide pour les cartouches normés.")
]
start_y14 = 1.60
gap14 = 0.12
card_h14 = 1.15
for oi, (ohd, obd) in enumerate(items_ocr):
    oy = start_y14 + oi * (card_h14 + gap14)
    b_o = s14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(rx14), Inches(oy), Inches(rw14), Inches(card_h14))
    b_o.adjustments[0] = 0.03
    b_o.fill.solid()
    b_o.fill.fore_color.rgb = C_BG_SLIDE
    b_o.line.color.rgb = C_LINE
    b_o.line.width = Pt(0.75)
    
    bx_o = s14.shapes.add_textbox(Inches(rx14 + 0.16), Inches(oy + 0.08), Inches(rw14 - 0.32), Inches(card_h14 - 0.16))
    tf_o = bx_o.text_frame
    tf_o.word_wrap = True
    tf_o.margin_left = tf_o.margin_top = tf_o.margin_right = tf_o.margin_bottom = 0
    
    po1 = tf_o.paragraphs[0]
    po1.text = "• " + ohd + " :\n"
    ro1 = po1.runs[0]
    ro1.font.name = "Arial"
    ro1.font.size = Pt(10.0)
    ro1.font.bold = True
    ro1.font.color.rgb = C_NAVY
    
    po2 = tf_o.add_paragraph()
    po2.space_before = Pt(2)
    po2.text = obd
    ro2 = po2.runs[0]
    ro2.font.name = "Arial"
    ro2.font.size = Pt(8.5)
    ro2.font.color.rgb = C_DARK

# Callout bas droite
b_eval14 = s14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(rx14), Inches(5.45), Inches(rw14), Inches(1.20))
b_eval14.adjustments[0] = 0.03
b_eval14.fill.solid()
b_eval14.fill.fore_color.rgb = C_CALLOUT_BG
b_eval14.line.color.rgb = C_ACCENT
b_eval14.line.width = Pt(1.0)

bx_ev14 = s14.shapes.add_textbox(Inches(rx14 + 0.16), Inches(5.53), Inches(rw14 - 0.32), Inches(1.04))
tf_ev14 = bx_ev14.text_frame
tf_ev14.word_wrap = True
tf_ev14.margin_left = tf_ev14.margin_top = tf_ev14.margin_right = tf_ev14.margin_bottom = 0

pe1 = tf_ev14.paragraphs[0]
pe1.text = "ÉVALUATION OPÉRATIONNELLE DU TEXTE IMPRIMÉ :"
re1 = pe1.runs[0]
re1.font.name = "Arial"
re1.font.size = Pt(9.0)
re1.font.bold = True
re1.font.color.rgb = C_ACCENT

pe2 = tf_ev14.add_paragraph()
pe2.space_before = Pt(2)
pe2.text = "Très efficace sur les formulaires Cerfa récents. En revanche, le taux d'erreur dépasse 30% dès que le texte est manuscrit ou altéré."
re2 = pe2.runs[0]
re2.font.name = "Arial"
re2.font.size = Pt(8.5)
re2.font.color.rgb = C_DARK

# Slide 15 (3.2) : HTR manuscrit (Fiches à gauche, image taxonomie à droite)
print("Génération Slide 15 (HTR manuscrit)...")
s15, c15 = init_standard_slide(2, 1, 15, "3.2 Reconnaissance d'écriture manuscrite : L'approche HTR")
add_two_column_content(
    s15,
    items=[
        ("Taxonomie des systèmes (Alkendi, 2024)",
         "L'écriture cursive manuscrite impose une analyse séquentielle continue plutôt qu'une séparation par caractères."),
        ("TrOCR (Microsoft Research, 2021)",
         "Vision Transformer combinant un encodeur d'image (DeiT) et un décodeur linguistique autorégressif (RoBERTa)."),
        ("PyLaia & corpus spécialisés HTR-United",
         "Réseau hybride CNN-BLSTM pré-entraîné sur des registres d'archives historiques françaises pour lire les graphies anciennes.")
    ],
    img_path="img/fig_alkendi_p3_1.jpeg",
    img_caption="Classification des systèmes : distinction imprimé (OCR) et manuscrit (HTR) (Alkendi et al., 2024)",
    callout_title="ÉVALUATION DU TEXTE MANUSCRIT",
    callout_text="TrOCR atteint un taux d'erreur de caractères (CER) inférieur à 6% sur le manuscrit, à condition d'isoler les lignes au préalable."
)

# Slide 16 (3.3) : NER (Disposition comparative 2 grandes colonnes : GLiNER vs LayoutLMv3)
print("Génération Slide 16 (NER - Comparatif GLiNER vs LayoutLM)...")
s16, c16 = init_standard_slide(2, 2, 16, "3.3 Extraction d'entités nommées (NER) : GLiNER vs LayoutLM")

cw16 = 5.50
# Colonne 1 : GLiNER
b_gl = s16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(1.60), Inches(cw16), Inches(4.05))
b_gl.adjustments[0] = 0.03
b_gl.fill.solid()
b_gl.fill.fore_color.rgb = C_BG_SLIDE
b_gl.line.color.rgb = C_SUCCESS
b_gl.line.width = Pt(1.2)

bx_gl = s16.shapes.add_textbox(Inches(1.20), Inches(1.72), Inches(cw16 - 0.30), Inches(3.80))
tf_gl = bx_gl.text_frame
tf_gl.word_wrap = True
tf_gl.margin_left = tf_gl.margin_top = tf_gl.margin_right = tf_gl.margin_bottom = 0

pgl_h = tf_gl.paragraphs[0]
pgl_h.text = "GLiNER : APPROCHE BI-ENCODEUR ZERO-SHOT (RETENU)\n"
rgl_h = pgl_h.runs[0]
rgl_h.font.name = "Arial"
rgl_h.font.size = Pt(10.5)
rgl_h.font.bold = True
rgl_h.font.color.rgb = C_SUCCESS

gl_items = [
    ("Architecture compacte", "Modèle bi-encodeur DeBERTa (340M paramètres) exécutable rapidement sur CPU."),
    ("Extraction sans réentraînement", "Extraction zero-shot d'entités arbitraires (commune, date, numéro, parcelles)."),
    ("Tolérance au bruit OCR", "Compréhension contextuelle robuste permettant de surmonter les coquilles de lecture."),
    ("Précision mesurée", "F1 = 0.88 sur le corpus de test du cabinet, avec score de confiance par entité.")
]
for gh, gb in gl_items:
    p_g = tf_gl.add_paragraph()
    p_g.space_before = Pt(5)
    r1 = p_g.add_run()
    r1.text = "✓ " + gh + " : "
    r1.font.name = "Arial"
    r1.font.size = Pt(9.0)
    r1.font.bold = True
    r1.font.color.rgb = C_NAVY
    r2 = p_g.add_run()
    r2.text = gb
    r2.font.name = "Arial"
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = C_DARK

# Colonne 2 : LayoutLMv3
b_lm = s16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.78), Inches(1.60), Inches(cw16), Inches(4.05))
b_lm.adjustments[0] = 0.03
b_lm.fill.solid()
b_lm.fill.fore_color.rgb = C_BG_SLIDE
b_lm.line.color.rgb = C_BORDER_TAB
b_lm.line.width = Pt(0.75)

bx_lm = s16.shapes.add_textbox(Inches(6.93), Inches(1.72), Inches(cw16 - 0.30), Inches(3.80))
tf_lm = bx_lm.text_frame
tf_lm.word_wrap = True
tf_lm.margin_left = tf_lm.margin_top = tf_lm.margin_right = tf_lm.margin_bottom = 0

plm_h = tf_lm.paragraphs[0]
plm_h.text = "LayoutLMv3 : APPROCHE MULTIMODALE 2D (ÉCARTÉ)\n"
rlm_h = plm_h.runs[0]
rlm_h.font.name = "Arial"
rlm_h.font.size = Pt(10.5)
rlm_h.font.bold = True
rlm_h.font.color.rgb = C_MUTED

lm_items = [
    ("Modèle multimodal lourd", "Combine signal visuel, coordonnées spatiales 2D et contenu textuel (~8 Go VRAM)."),
    ("Optimisé formulaires rigides", "Très adapté aux imprimés stricts, mais inopérant sur les registres en colonnes libres."),
    ("Sensibilité extrême aux fautes", "Chute brutale des performances dès que l'OCR initial contient des erreurs de lecture."),
    ("Dépendance matérielle", "Nécessite des ressources GPU indisponibles sur les postes standards du cabinet.")
]
for lh, lb in lm_items:
    p_l = tf_lm.add_paragraph()
    p_l.space_before = Pt(5)
    r1 = p_l.add_run()
    r1.text = "✗ " + lh + " : "
    r1.font.name = "Arial"
    r1.font.size = Pt(9.0)
    r1.font.bold = True
    r1.font.color.rgb = C_MUTED
    r2 = p_l.add_run()
    r2.text = lb
    r2.font.name = "Arial"
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = C_DARK

# Grand bandeau bas
b_syn16 = s16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.80), Inches(11.23), Inches(0.85))
b_syn16.adjustments[0] = 0.03
b_syn16.fill.solid()
b_syn16.fill.fore_color.rgb = C_CALLOUT_BG
b_syn16.line.color.rgb = C_ACCENT
b_syn16.line.width = Pt(1.0)

bx_s16 = s16.shapes.add_textbox(Inches(1.20), Inches(5.88), Inches(10.93), Inches(0.70))
tf_s16 = bx_s16.text_frame
tf_s16.word_wrap = True
tf_s16.margin_left = tf_s16.margin_top = tf_s16.margin_right = tf_s16.margin_bottom = 0
ps16_1 = tf_s16.paragraphs[0]
ps16_1.text = "CONCLUSION COMPARATIVE DE L'EXTRACTION D'ENTITÉS :"
rs16_1 = ps16_1.runs[0]
rs16_1.font.name = "Arial"
rs16_1.font.size = Pt(9.0)
rs16_1.font.bold = True
rs16_1.font.color.rgb = C_ACCENT

ps16_2 = tf_s16.add_paragraph()
ps16_2.space_before = Pt(2)
ps16_2.text = "GLiNER est retenu pour son autonomie d'extraction sans réentraînement, sa faible empreinte mémoire sur CPU et sa robustesse éprouvée face aux imperfections de l'OCR."
rs16_2 = ps16_2.runs[0]
rs16_2.font.name = "Arial"
rs16_2.font.size = Pt(8.5)
rs16_2.font.color.rgb = C_NAVY

# Slide 17 (3.4) : VLM (Cartes à gauche, schéma architecture à droite)
print("Génération Slide 17 (VLM local)...")
s17, c17 = init_standard_slide(2, 3, 17, "3.4 Modèles Vision-Langage (VLM) : Arbitrage multimodal local")
add_two_column_content(
    s17,
    items=[
        ("Paradigme Vision-Langage (LLaVA / MiniCPM)",
         "Couplage d'un encodeur visuel (CLIP ViT) et d'un LLM autorégressif. L'image est traitée directement comme une séquence de tokens."),
        ("Raisonnement visuel et compréhension de scène",
         "Interprétation conjointe du document : capacité à relier une mention manuscrite avec sa position spatiale sur la page."),
        ("Compromis latence et exécution locale",
         "Temps d'inférence de ~5s à 8s sur processeur CPU. Ce coût impose de réserver le VLM en tant qu'arbitre de second niveau.")
    ],
    img_path="img/llava_arch_render.png",
    img_caption="Architecture des modèles Vision-Langage : projection des patchs visuels dans le LLM (Liu et al., 2023)",
    callout_title="CONFIDENTIALITÉ ET TRAITEMENT LOCAL",
    callout_text="Exécution 100% locale via le moteur Ollama : aucune donnée nominative d'archive n'est transmise vers des serveurs externes."
)"""

assert old_s14_to_s17 in content, "old_s14_to_s17 non trouvé !"
content = content.replace(old_s14_to_s17, new_s14_to_s17)

# 11. Slide 18 (3.5) : Synthèse comparative EXACTEMENT CONFORME AU TABLEAU 3.1 DU MÉMOIRE
old_s18 = """# Slide 18 (3.5) : Synthèse comparative et Choix retenus
print("Génération Slide 18 (Synthèse comparative officielle)...")
s18, c18 = init_standard_slide(2, 4, 18, "3.5 Synthèse comparative et justification des choix retenus")

headers18 = ["Modèle", "Méthode", "Ressources", "Points forts", "Limites", "Décision"]
rows18 = [
    ["Tesseract v5", "Analyse par caractères", "CPU", "Rapide et économe en mémoire", "Très faible sur manuscrit", "ÉCARTÉ"],
    ["TrOCR / PyLaia", "Vision Transformer", "~4 Go GPU", "Forte précision manuscrit (CER < 6%)", "Exige découpage précis", "RETENU"],
    ["YOLOv8", "Détection d'objets CIoU", "< 2 Go GPU", "Détection très rapide (< 0.2s) et robuste", "Isole sans lire le texte", "RETENU"],
    ["LayoutLMv3", "Multimodal 2D texte+boîtes", "~8 Go GPU", "Analyse conjointe forme et mots", "Sensible aux fautes OCR", "ÉCARTÉ"],
    ["GLiNER", "Bi-encodeur BERT", "~2 Go GPU", "Zero-shot F1 91.5% sans entraînement", "Dépend de la qualité texte", "RETENU"],
    ["LLaVA / Ollama", "Vision-Langage local", "> 8 Go GPU", "Raisonnement visuel direct en local", "Latence CPU (~5s)", "RETENU (Secours)"],
    ["Cloud APIs", "Vision distante (Google/AWS)", "Cloud externe", "Haute précision sur documents types", "Violation secret pro / RGPD", "ÉCARTÉ"]
]

cw18 = [1.70, 2.25, 1.40, 2.75, 2.13, 1.00]
rh18 = 0.49
tx18 = 1.05
ty18 = 1.60

hx = tx18
for hi, htext in enumerate(headers18):
    h_box = s18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(hx), Inches(ty18), Inches(cw18[hi]), Inches(rh18))
    h_box.adjustments[0] = 0.08
    h_box.fill.solid()
    h_box.fill.fore_color.rgb = C_NAVY
    h_box.line.fill.background()
    
    bx = s18.shapes.add_textbox(Inches(hx + 0.04), Inches(ty18 + 0.08), Inches(cw18[hi] - 0.08), Inches(rh18 - 0.16))
    tf = bx.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = htext
    p.alignment = PP_ALIGN.CENTER
    r = p.runs[0]
    r.font.name = "Arial"
    r.font.size = Pt(9.5)
    r.font.bold = True
    r.font.color.rgb = C_WHITE
    hx += cw18[hi]

for ri, row in enumerate(rows18):
    ry = ty18 + (ri + 1) * rh18
    bg_r = C_CARD_BG if ri % 2 == 0 else C_BG_SLIDE
    rx = tx18
    for ci, cell in enumerate(row):
        c_box = s18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(rx), Inches(ry), Inches(cw18[ci]), Inches(rh18))
        c_box.adjustments[0] = 0.08
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = bg_r
        c_box.line.color.rgb = C_LINE
        c_box.line.width = Pt(0.75)
        
        bx = s18.shapes.add_textbox(Inches(rx + 0.06), Inches(ry + 0.06), Inches(cw18[ci] - 0.12), Inches(rh18 - 0.12))
        tf = bx.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = cell
        p.alignment = PP_ALIGN.CENTER if ci in (0, 2, 5) else PP_ALIGN.LEFT
        r = p.runs[0]
        r.font.name = "Arial"
        r.font.size = Pt(8.5)
        
        if "RETENU" in cell:
            r.font.bold = True
            r.font.color.rgb = C_SUCCESS
        elif "ÉCARTÉ" in cell:
            r.font.bold = True
            r.font.color.rgb = C_RED
        elif ci == 0:
            r.font.bold = True
            r.font.color.rgb = C_NAVY
        else:
            r.font.color.rgb = C_DARK
            
        rx += cw18[ci]

sy18 = 5.82
s_box18 = s18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(tx18), Inches(sy18), Inches(sum(cw18)), Inches(0.83))
s_box18.adjustments[0] = 0.03
s_box18.fill.solid()
s_box18.fill.fore_color.rgb = C_CALLOUT_BG
s_box18.line.color.rgb = C_ACCENT
s_box18.line.width = Pt(1.0)

bx_syn18 = s18.shapes.add_textbox(Inches(tx18 + 0.15), Inches(sy18 + 0.10), Inches(sum(cw18) - 0.30), Inches(0.63))
tf_syn18 = bx_syn18.text_frame
tf_syn18.word_wrap = True
p_syn18 = tf_syn18.paragraphs[0]
p_syn18.text = "ARCHITECTURE HYBRIDE SOUVERAINE :"
p_syn18.alignment = PP_ALIGN.CENTER
r_syn18 = p_syn18.runs[0]
r_syn18.font.name = "Arial"
r_syn18.font.size = Pt(9.5)
r_syn18.font.bold = True
r_syn18.font.color.rgb = C_ACCENT

p_syn18_2 = tf_syn18.add_paragraph()
p_syn18_2.space_before = Pt(2)
p_syn18_2.text = "Combinaison optimale : YOLOv8 (découpage) + TrOCR / EasyOCR (transcription) + GLiNER (NER) + LLaVA (arbitre local)."
p_syn18_2.alignment = PP_ALIGN.CENTER
r_syn18_2 = p_syn18_2.runs[0]
r_syn18_2.font.name = "Arial"
r_syn18_2.font.size = Pt(8.5)
r_syn18_2.font.bold = True
r_syn18_2.font.color.rgb = C_NAVY"""

new_s18 = """# Slide 18 (3.5) : Synthèse comparative (Tableau 3.1 exact du mémoire)
print("Génération Slide 18 (Tableau 3.1 exact du mémoire)...")
s18, c18 = init_standard_slide(2, 4, 18, "3.5 Synthèse comparative des technologies d'analyse de document")

# Les 5 colonnes exactes du Tableau 3.1 du mémoire
headers18 = ["Modèle", "Méthode utilisée", "Ressources", "Points forts", "Limites principales"]
# Les 8 lignes exactes du mémoire (Tableau 3.1)
rows18 = [
    ["Tesseract v5", "Analyse par caractères", "CPU", "Rapide et économe en mémoire.", "Précision très faible sur le manuscrit."],
    ["PyLaia / TrOCR", "Réseau convolutif et Transformer", "~4 Go GPU", "Forte précision sur le manuscrit.", "Nécessite d'isoler les lignes au préalable."],
    ["YOLOv8", "Détection d'objets spatiale", "< 2 Go GPU", "Détection visuelle très rapide et robuste.", "Isole les zones sans lire le texte."],
    ["LayoutLMv3", "Modèle multimodal vision-texte", "~8 Go GPU", "Analyse conjointe de la forme et des mots.", "Sensible aux erreurs de lecture initiale."],
    ["GLiNER", "Comparaison de représentations", "~2 Go GPU", "Identification d'entités sans réentraînement.", "Dépend de la propreté du texte transmis."],
    ["Donut", "Décodage visuel direct", "~4 Go GPU", "Extraction sans OCR externe.", "Limité aux types de fichiers connus."],
    ["Qwen-VL", "VLM à résolution dynamique", "~8 Go GPU", "Lecture précise des détails et petits textes.", "Temps de calcul élevé sur grand format."],
    ["LLaVA / Ollama", "Modèle Vision-Langage local", "> 8 Go GPU", "Raisonnement visuel direct en local.", "Calcul lent et risque d'hallucination."]
]

cw18 = [1.80, 2.50, 1.30, 2.80, 2.83]
rh18 = 0.43
tx18 = 1.05
ty18 = 1.55

hx = tx18
for hi, htext in enumerate(headers18):
    h_box = s18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(hx), Inches(ty18), Inches(cw18[hi]), Inches(0.38))
    h_box.adjustments[0] = 0.08
    h_box.fill.solid()
    h_box.fill.fore_color.rgb = C_NAVY
    h_box.line.fill.background()
    
    bx = s18.shapes.add_textbox(Inches(hx + 0.04), Inches(ty18 + 0.04), Inches(cw18[hi] - 0.08), Inches(0.30))
    tf = bx.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = htext
    p.alignment = PP_ALIGN.CENTER
    r = p.runs[0]
    r.font.name = "Arial"
    r.font.size = Pt(9.0)
    r.font.bold = True
    r.font.color.rgb = C_WHITE
    hx += cw18[hi]

for ri, row in enumerate(rows18):
    ry = ty18 + 0.38 + ri * rh18
    bg_r = C_CARD_BG if ri % 2 == 0 else C_BG_SLIDE
    rx = tx18
    for ci, cell in enumerate(row):
        c_box = s18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(rx), Inches(ry), Inches(cw18[ci]), Inches(rh18))
        c_box.adjustments[0] = 0.08
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = bg_r
        c_box.line.color.rgb = C_LINE
        c_box.line.width = Pt(0.75)
        
        bx = s18.shapes.add_textbox(Inches(rx + 0.06), Inches(ry + 0.04), Inches(cw18[ci] - 0.12), Inches(rh18 - 0.08))
        tf = bx.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = cell
        p.alignment = PP_ALIGN.CENTER if ci in (0, 2) else PP_ALIGN.LEFT
        r = p.runs[0]
        r.font.name = "Arial"
        r.font.size = Pt(8.0)
        
        if ci == 0:
            r.font.bold = True
            r.font.color.rgb = C_NAVY
        else:
            r.font.color.rgb = C_DARK
            
        rx += cw18[ci]

sy18 = 5.45
s_box18 = s18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(tx18), Inches(sy18), Inches(sum(cw18)), Inches(1.15))
s_box18.adjustments[0] = 0.03
s_box18.fill.solid()
s_box18.fill.fore_color.rgb = C_CALLOUT_BG
s_box18.line.color.rgb = C_ACCENT
s_box18.line.width = Pt(1.0)

bx_syn18 = s18.shapes.add_textbox(Inches(tx18 + 0.15), Inches(sy18 + 0.10), Inches(sum(cw18) - 0.30), Inches(0.95))
tf_syn18 = bx_syn18.text_frame
tf_syn18.word_wrap = True
p_syn18 = tf_syn18.paragraphs[0]
p_syn18.text = "CHOIX D'ARCHITECTURE RETENUS DANS LE MÉMOIRE :"
p_syn18.alignment = PP_ALIGN.CENTER
r_syn18 = p_syn18.runs[0]
r_syn18.font.name = "Arial"
r_syn18.font.size = Pt(9.5)
r_syn18.font.bold = True
r_syn18.font.color.rgb = C_ACCENT

p_syn18_2 = tf_syn18.add_paragraph()
p_syn18_2.space_before = Pt(3)
p_syn18_2.text = "Couplage optimal : YOLOv8 (découpage spatial) + PyLaia / EasyOCR (transcription ciblée) + GLiNER (extraction d'entités) + LLaVA (arbitre local de second niveau sur seuil de confiance)."
p_syn18_2.alignment = PP_ALIGN.CENTER
r_syn18_2 = p_syn18_2.runs[0]
r_syn18_2.font.name = "Arial"
r_syn18_2.font.size = Pt(8.5)
r_syn18_2.font.bold = True
r_syn18_2.font.color.rgb = C_NAVY"""

assert old_s18 in content, "old_s18 non trouvé !"
content = content.replace(old_s18, new_s18)

# 12. Slide 19 : Transition Chapitre 4
content = content.replace(
    'title="04 - Architecture Logicielle et Pipeline de Traitement",',
    'title="Chapitre 4 : Architecture logicielle et chaîne de traitement",'
)

# 13. Slide 20 (4.1) : Figure 4.1 du mémoire (pipeline_6_etapes_clean.png) + rappel protection des données
old_s20 = """# Slide 20 (4.1) : Architecture globale
s20, c20 = init_standard_slide(3, 0, 20, "4.1 Architecture logicielle : Schéma de la chaîne de traitement")

frame_arch = s20.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(1.55), Inches(11.23), Inches(4.15))
frame_arch.adjustments[0] = 0.03
frame_arch.fill.solid()
frame_arch.fill.fore_color.rgb = C_BG_SLIDE
frame_arch.line.color.rgb = C_LINE
frame_arch.line.width = Pt(0.75)

add_fitted_picture(s20, "img/architecture_pfe_finale_fixed.png", 1.20, 1.65, 10.93, 3.95, caption="Schéma général de la chaîne de traitement développée (Mémoire, Chapitre 4)")

b_flux20 = s20.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.82), Inches(11.23), Inches(0.83))
b_flux20.adjustments[0] = 0.03
b_flux20.fill.solid()
b_flux20.fill.fore_color.rgb = C_CALLOUT_BG
b_flux20.line.color.rgb = C_ACCENT
b_flux20.line.width = Pt(1.0)

bx_flt = s20.shapes.add_textbox(Inches(1.20), Inches(5.88), Inches(10.93), Inches(0.25))
tf_flt = bx_flt.text_frame
p_flt = tf_flt.paragraphs[0]
p_flt.text = "FLUX OPÉRATIONNEL DE BOUT EN BOUT :"
r_flt = p_flt.runs[0]
r_flt.font.name = "Arial"
r_flt.font.size = Pt(9.0)
r_flt.font.bold = True
r_flt.font.color.rgb = C_ACCENT

bx_flb = s20.shapes.add_textbox(Inches(1.20), Inches(6.12), Inches(10.93), Inches(0.50))
tf_flb = bx_flb.text_frame
tf_flb.word_wrap = True
p_flb = tf_flb.paragraphs[0]
p_flb.text = "1. Images d'archives brutes  →  2. Découpage spatial (YOLOv8)  →  3. Transcription hybride (TrOCR / EasyOCR)  →  4. Extraction d'entités (GLiNER)  →  5. Contrôle contradictoire (Streamlit)  →  6. Injection API Géofoncier (RFU)."
r_flb = p_flb.runs[0]
r_flb.font.name = "Arial"
r_flb.font.size = Pt(9.0)
r_flb.font.bold = True
r_flb.font.color.rgb = C_NAVY"""

new_s20 = """# Slide 20 (4.1) : Architecture globale : Figure 4.1 du mémoire + Protection des données
print("Génération Slide 20 (Figure 4.1 pipeline 6 étapes + protection des données)...")
s20, c20 = init_standard_slide(3, 0, 20, "4.1 Architecture logicielle : Schéma de la chaîne de traitement")

# Cadre supérieur image : Figure 4.1 authentique du mémoire
frame_arch = s20.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(1.55), Inches(11.23), Inches(3.40))
frame_arch.adjustments[0] = 0.03
frame_arch.fill.solid()
frame_arch.fill.fore_color.rgb = C_BG_SLIDE
frame_arch.line.color.rgb = C_LINE
frame_arch.line.width = Pt(0.75)

add_fitted_picture(s20, "img/pipeline_6_etapes_clean.png", 1.20, 1.62, 10.93, 3.25, caption="Figure 4.1 : Enchaînement opérationnel de la chaîne de traitement en 6 étapes (Mémoire, Chapitre 4)")

# Encadré protection des données (Secret professionnel / RGPD / Local)
b_prot20 = s20.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.05), Inches(11.23), Inches(0.70))
b_prot20.adjustments[0] = 0.03
b_prot20.fill.solid()
b_prot20.fill.fore_color.rgb = C_BG_SLIDE
b_prot20.line.color.rgb = C_SUCCESS
b_prot20.line.width = Pt(1.0)

bx_prot = s20.shapes.add_textbox(Inches(1.20), Inches(5.10), Inches(10.93), Inches(0.60))
tf_prot = bx_prot.text_frame
tf_prot.word_wrap = True
tf_prot.margin_left = tf_prot.margin_top = tf_prot.margin_right = tf_prot.margin_bottom = 0

pp1 = tf_prot.paragraphs[0]
pp1.text = "PROTECTION DES DONNÉES ET RESPECT DU SECRET PROFESSIONNEL :"
rp1 = pp1.runs[0]
rp1.font.name = "Arial"
rp1.font.size = Pt(8.5)
rp1.font.bold = True
rp1.font.color.rgb = C_SUCCESS

pp2 = tf_prot.add_paragraph()
pp2.space_before = Pt(2)
pp2.text = "Exécution 100% locale sur l'unité centrale (CPU) d'un ordinateur de bureau du cabinet. Aucune donnée d'archive ni nom de client n'est transmis vers un serveur externe, assurant la conformité RGPD."
rp2 = pp2.runs[0]
rp2.font.name = "Arial"
rp2.font.size = Pt(8.0)
rp2.font.color.rgb = C_DARK

# Bandeau bas : Flux opérationnel maintenu
b_flux20 = s20.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.85), Inches(11.23), Inches(0.80))
b_flux20.adjustments[0] = 0.03
b_flux20.fill.solid()
b_flux20.fill.fore_color.rgb = C_CALLOUT_BG
b_flux20.line.color.rgb = C_ACCENT
b_flux20.line.width = Pt(1.0)

bx_flt = s20.shapes.add_textbox(Inches(1.20), Inches(5.90), Inches(10.93), Inches(0.24))
tf_flt = bx_flt.text_frame
p_flt = tf_flt.paragraphs[0]
p_flt.text = "FLUX OPÉRATIONNEL DE BOUT EN BOUT :"
r_flt = p_flt.runs[0]
r_flt.font.name = "Arial"
r_flt.font.size = Pt(8.5)
r_flt.font.bold = True
r_flt.font.color.rgb = C_ACCENT

bx_flb = s20.shapes.add_textbox(Inches(1.20), Inches(6.12), Inches(10.93), Inches(0.48))
tf_flb = bx_flb.text_frame
tf_flb.word_wrap = True
p_flb = tf_flb.paragraphs[0]
p_flb.text = "1. Images d'archives brutes  →  2. Découpage spatial (YOLOv8)  →  3. Transcription hybride (TrOCR / EasyOCR)  →  4. Extraction d'entités (GLiNER)  →  5. Contrôle contradictoire (Streamlit)  →  6. Injection API Géofoncier (RFU)."
r_flb = p_flb.runs[0]
r_flb.font.name = "Arial"
r_flb.font.size = Pt(8.5)
r_flb.font.bold = True
r_flb.font.color.rgb = C_NAVY"""

assert old_s20 in content, "old_s20 non trouvé !"
content = content.replace(old_s20, new_s20)

# 14. Slide 23 (4.4) : Zoom sur Lachapelle-sous-Aubenas et vocabulaire rigoureux (Confidentialité et traitement local)
old_s23 = """# Slide 23 (4.4) : Arbitrage VLM
s23, c23 = init_standard_slide(3, 3, 23, "4.4 Arbitrage visuel multimodal par modèle Vision-Langage")
add_two_column_content(
    s23,
    items=[
        ("Déclenchement conditionnel sur seuil",
         "Le modèle VLM local n'est interrogé que si le score de confiance GLiNER est inférieur à 0.65, préservant la fluidité globale."),
        ("Désambiguïsation toponymique complexe",
         "Capacité à interpréter les mentions manuscrites locales complexes et les abréviations de communes ardéchoises (ex. « La Chap./A. »)."),
        ("Garantie absolue de confidentialité locale",
         "Exécution intégrale via Ollama sur l'ordinateur du cabinet, interdisant tout transfert de données vers un serveur externe.")
    ],
    img_path="img/vlm_arbitrage_chapelle_exact.png",
    img_caption="Exemple d'arbitrage visuel local : résolution de « La Chap./A. » en « La Chapelle-sous-Aubenas »",
    callout_title="SOUVERAINETÉ ET DÉONTOLOGIE",
    callout_text="L'arbitre VLM local résout les cas litigieux tout en assurant le respect absolu du secret professionnel et du RGPD."
)"""

new_s23 = """# Slide 23 (4.4) : Arbitrage VLM (Zoom très net sur Lachapelle-sous-Aubenas et vocabulaire rigoureux)
print("Génération Slide 23 (Arbitrage VLM avec zoom haute résolution Lachapelle)...")
s23, c23 = init_standard_slide(3, 3, 23, "4.4 Arbitrage visuel multimodal par modèle Vision-Langage")
add_two_column_content(
    s23,
    items=[
        ("Déclenchement conditionnel sur seuil",
         "L'arbitre VLM local n'est sollicité que si le score de confiance GLiNER est inférieur à 0.65, préservant la fluidité globale."),
        ("Désambiguïsation toponymique contextuelle",
         "Capacité à interpréter les mentions manuscrites difficiles et abréviations ardéchoises (ex. « Lachapelle /s/ AUBENAS »)."),
        ("Exécution 100% locale sur station de travail",
         "Déploiement direct via Ollama sur l'ordinateur du cabinet, garantissant l'absence de fuite et le respect du secret professionnel.")
    ],
    img_path="img/vlm_arbitrage_chapelle_zoom.png",
    img_caption="Arbitrage visuel local sur toponyme manuscrit complexe : résolution de « Lachapelle /s/ AUBENAS »",
    callout_title="CONFIDENTIALITÉ ET TRAITEMENT LOCAL",
    callout_text="L'arbitre VLM local résout les cas complexes tout en assurant le respect strict du secret professionnel et du RGPD."
)"""

assert old_s23 in content, "old_s23 non trouvé !"
content = content.replace(old_s23, new_s23)

# 15. Slide 24 : Transition Chapitre 5
content = content.replace(
    'title="05 - Fiabilisation des Données et Interface Opérateur",',
    'title="Chapitre 5 : Fiabilisation des données et interface opérateur",'
)

# 16. Slide 26 : Remplacer le titre ALL-CAPS du callout
content = content.replace(
    'ps26_1.text = "SÉCURISATION DU PROCESSUS PAR 17 RÈGLES MÉTIER :"',
    'ps26_1.text = "Sécurisation par 17 règles de cohérence métier :"'
)

# 17. Slide 28 : Transition Chapitre 6
content = content.replace(
    'title="06 - Intégration Géofoncier et Déploiement Opérationnel",',
    'title="Chapitre 6 : Intégration Géofoncier et expérimentation",'
)

# 18. Slide 30 : Callout sans ALL-CAPS
content = content.replace(
    'ps30_1.text = "CONFIRMATION OFFICIELLE DU VERSEMENT :"',
    'ps30_1.text = "Validation officielle du versement sur Géofoncier :"'
)

# 19. Slide 31 : Titre sans ALL-CAPS
content = content.replace(
    'p_dh.text = "BILAN DE L\'EXPÉRIMENTATION OPÉRATIONNELLE :\\n"',
    'p_dh.text = "Bilan de l\'expérimentation sur Prades :\\n"'
)
content = content.replace(
    'pf_h.text = "SCORES F1 PAR CHAMP OBLIGATOIRE (LOT DE PRADES) :"',
    'pf_h.text = "Scores F1 par champ obligatoire (lot de Prades) :"'
)

# 20. Slide 32 : Transition Chapitre 7
content = content.replace(
    'title="07 - Analyse des Limites et Perspectives d\'Évolution",',
    'title="Chapitre 7 : Analyse des limites et perspectives d\'évolution",'
)

# 21. Slide 34 : Callout sans ALL-CAPS
content = content.replace(
    'p_c34.text = "VALORISATION PATRIMONIALE : Un potentiel concret de dématérialisation d\'environ 40 000 dossiers répartis sur l\'ensemble du réseau SIAPP."',
    'p_c34.text = "Valorisation patrimoniale : Un potentiel concret de dématérialisation d\'environ 40 000 dossiers répartis sur l\'ensemble du réseau SIAPP."'
)

# 22. Slide 35 : Transition Chapitre 8
content = content.replace(
    'title="08 - Conclusion Générale et Bilan du Projet",',
    'title="Chapitre 8 : Conclusion générale et bilan du projet",'
)
content = content.replace(
    'subtitle="Bilan opérationnel pour GEO-SIAPP, souveraineté et compétences d\'ingénieur INSA"',
    'subtitle="Bilan opérationnel pour GEO-SIAPP, traitement local et compétences d\'ingénieur INSA"'
)

# 23. Slide 36 : Remplacer les buzzwords "souveraineté" et majuscules
content = content.replace(
    '("100% SOUVERAIN", "EXÉCUTION LOCALE CPU / GPU", C_ACCENT),',
    '("100% LOCAL", "STATION DE TRAVAIL CPU / GPU", C_ACCENT),'
)
content = content.replace(
    'p_bot36.text = "SYNTHÈSE : La transformation concrète d\'un passif d\'archives papier en un actif numérique opposable, géolocalisé et souverain."',
    'p_bot36.text = "Synthèse : La transformation concrète d\'un fonds d\'archives papier en données numériques opposables, géolocalisées et sécurisées."'
)
content = content.replace(
    'ppl_h.text = "APPORTS MAJEURS POUR LE CABINET GEO-SIAPP\\n"',
    'ppl_h.text = "Apports opérationnels pour le cabinet GEO-SIAPP\\n"'
)
content = content.replace(
    'ppr_h.text = "COMPÉTENCES ET RETOUR D\'EXPÉRIENCE INGÉNIEUR\\n"',
    'ppr_h.text = "Compétences d\'ingénieur et retour d\'expérience\\n"'
)

# Sauvegarde
with open("generate_final_180s_deck.py", "w", encoding="utf-8") as f:
    f.write(content)

print("Mise à jour réussie de generate_final_180s_deck.py !")
