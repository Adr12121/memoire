# -*- coding: utf-8 -*-
"""
Script to apply the requested refinements to generate_final_180s_deck.py:
1. Sommaire (Slide 2): ONLY chapter titles, no subtitles, no descriptions!
2. Header Ribbon (frise): Remove excess capitals in CHAPTERS_DATA (short & subs).
3. Slide 26 (5.2 Règles métier): Exploit full slide space with 2x2 grid of the 4 authentic rule families from thesis (17 rules).
4. Slide 34 (7.2 Perspectives): Align with 07_Limites_et_Perspectives.tex (FastAPI/HTML, WFS DGFiP & Plugin QGIS, Transposabilité, Apprentissage actif).
5. Slide 36 (8.1 Bilan PFE): Align with 08_Conclusion.tex (Charte OGE 2026, 4 principes, gain x6, sécurité et pérennité, posture ingénieur).
"""

with open('generate_final_180s_deck.py', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Update CHAPTERS_DATA: lowercase excess capitals in short titles and subs
old_chapters = '''CHAPTERS_DATA = [
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
]'''

new_chapters = '''CHAPTERS_DATA = [
    {
        "short": "1. Introduction",
        "full": "Introduction et contexte du projet",
        "badge": "1",
        "subtitle": "Structure d'accueil GEO-SIAPP, cadre réglementaire et problématique foncière",
        "subs": ["1.1 Structure GEO-SIAPP", "1.2 Cadre réglementaire", "1.3 Contexte & antériorité", "1.4 Problématique foncière"]
    },
    {
        "short": "2. Données & métier",
        "full": "Analyse métier et enjeux documentaires",
        "badge": "2",
        "subtitle": "Distinction cadastre et actes, spécifications API Géofoncier et filiation parcellaire",
        "subs": ["2.1 Cadre juridique", "2.2 Spécifications API", "2.3 Typologie des archives", "2.4 Filiation parcellaire"]
    },
    {
        "short": "3. État de l'art",
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
        "subs": ["8.1 Bilan du projet", "8.2 Remerciements & échanges"]
    }
]'''

if old_chapters in code:
    code = code.replace(old_chapters, new_chapters)
    print("1. CHAPTERS_DATA updated successfully.")
else:
    print("WARNING: old_chapters not found exactly.")

# 2. Update Slide 2 (Sommaire): strictly chapter titles, no subtitles, no descriptions
slide2_start = code.find('# ==================== SLIDE 2')
slide2_end = code.find('# ==================== CHAPITRE 1')

new_slide2 = '''# ==================== SLIDE 2 : SOMMAIRE ÉPURÉ ====================
print("Génération Slide 2 (Présentation sobre des 8 titres sans descriptif)...")
s2, c2 = init_standard_slide(-1, -1, 2, "Plan de la présentation : Sommaire")

# 2 colonnes équilibrées de 4 fiches élégantes présentant UNIQUEMENT les titres
cw2 = 5.50
ch2 = 1.08
cg2_y = 0.20
left_x = 1.05
right_x = 6.78
start_y = 1.62

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
    
    # Pilule Chapitre X
    pill_w = 1.35
    pill_h = 0.38
    pill_y = cy + (ch2 - pill_h) / 2.0
    pill_shp = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx + 0.22), Inches(pill_y), Inches(pill_w), Inches(pill_h))
    pill_shp.adjustments[0] = 0.20
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
    r_p.font.size = Pt(9.0)
    r_p.font.bold = True
    r_p.font.color.rgb = C_WHITE
    
    # Titre officiel du chapitre centré verticalement, sans aucun texte descriptif
    tit_x = cx + 1.75
    tit_w = cw2 - 1.95
    tit_h = 0.58
    tit_y = cy + (ch2 - tit_h) / 2.0
    bx_tit = s2.shapes.add_textbox(Inches(tit_x), Inches(tit_y), Inches(tit_w), Inches(tit_h))
    tf_tit = bx_tit.text_frame
    tf_tit.word_wrap = True
    tf_tit.margin_left = tf_tit.margin_top = tf_tit.margin_right = tf_tit.margin_bottom = 0
    p_tit = tf_tit.paragraphs[0]
    p_tit.text = chap["full"]
    r_tit = p_tit.runs[0]
    r_tit.font.name = "Arial"
    r_tit.font.size = Pt(13.0)
    r_tit.font.bold = True
    r_tit.font.color.rgb = C_NAVY

'''

if slide2_start != -1 and slide2_end != -1:
    code = code[:slide2_start] + new_slide2 + code[slide2_end:]
    print("2. Slide 2 updated successfully.")
else:
    print("WARNING: Slide 2 boundaries not found.")

# 3. Update Slide 26 (5.2 Règles métier): full size 2x2 grid of the 4 rule families (17 rules)
slide26_start = code.find('# Slide 26 (5.2) : Règles métier')
slide26_end = code.find('# Slide 27 (5.3) : Contrôle spatial Folium')

new_slide26 = '''# Slide 26 (5.2) : Règles métier
print("Génération Slide 26 (Moteur complet des 17 règles métier en grille 2x2 optimisée)...")
s26, c26 = init_standard_slide(4, 1, 26, "5.2 Moteur de 17 règles de cohérence métier et répertoires")

# 4 familles de règles issues du mémoire (section 5.3, 17 règles au total)
rule_groups = [
    {
        "tag": "GROUPE 1 • 5 RÈGLES",
        "title": "Contrôles cadastraux et toponymiques",
        "color": C_NAVY,
        "rules": [
            ("Rapprochement toponymique INSEE", "Levenshtein sur 335 communes 07 / 363 Drôme."),
            ("Validité de la section cadastrale", "Contrôle de l'existence de la section sur la commune."),
            ("Nettoyage des numéros de parcelles", "Élimination des parasites de tampons ou tirets."),
            ("Exclusion parcelle nulle & rattachement", "Rejet du numéro 0 et présence minimale d'1 parcelle.")
        ]
    },
    {
        "tag": "GROUPE 2 • 4 RÈGLES",
        "title": "Contrôles temporels et dates",
        "color": C_ACCENT,
        "rules": [
            ("Plage d'activité historique", "Cohérence avec le fonds documentaire (1950 à nos jours)."),
            ("Période d'exercice du géomètre", "Concordance date d'acte et exercice du signataire."),
            ("Validation calendaire réelle", "Format strict ISO 8601 et années bissextiles."),
            ("Antériorité chronologique", "Date d'enregistrement postérieure à la date du plan.")
        ]
    },
    {
        "tag": "GROUPE 3 • 4 RÈGLES",
        "title": "Contrôles des personnes et intervenants",
        "color": RGBColor(180, 83, 9),  # #B45309 Ambre professionnel
        "rules": [
            ("Annuaire des géomètres du cabinet", "Vérification du signataire dans la liste officielle."),
            ("Alerte identité parties", "Signalement si anciens propriétaires = acquéreurs."),
            ("Filtrage mentions isolées", "Suppression des termes « consorts », « veuve », « Idem »."),
            ("Donneur d'ordre & riverains", "Identification du client et chaînage des limitrophes.")
        ]
    },
    {
        "tag": "GROUPE 4 • 4 RÈGLES",
        "title": "Contrôles de dossier et anti-doublon",
        "color": C_SUCCESS,
        "rules": [
            ("Syntaxe de référence dossier", "Validation du masque propre à chaque fonds de géomètre."),
            ("Unicité stricte au répertoire", "Contrôle d'absence de doublon dans le registre général."),
            ("Détection conflits spatiaux", "Alerte si plusieurs actes sur même parcelle même année."),
            ("Intégrité du document source", "Vérification de présence et lisibilité du fichier PDF.")
        ]
    }
]

# Grille 2x2 équilibrée exploitant 100% de la largeur et de la hauteur utile
grid_w = 5.48
grid_h = 2.05
col_xs = [1.05, 6.80]
row_ys = [1.55, 3.72]

for gi, gdata in enumerate(rule_groups):
    gx = col_xs[gi % 2]
    gy = row_ys[gi // 2]
    
    b_g = s26.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(gx), Inches(gy), Inches(grid_w), Inches(grid_h))
    b_g.adjustments[0] = 0.03
    b_g.fill.solid()
    b_g.fill.fore_color.rgb = C_BG_SLIDE
    b_g.line.color.rgb = C_LINE
    b_g.line.width = Pt(0.75)
    
    # Accent couleur sur bordure gauche
    b_acc = s26.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(gx), Inches(gy), Inches(0.08), Inches(grid_h))
    b_acc.adjustments[0] = 0.20
    b_acc.fill.solid()
    b_acc.fill.fore_color.rgb = gdata["color"]
    b_acc.line.fill.background()
    
    bx_g = s26.shapes.add_textbox(Inches(gx + 0.18), Inches(gy + 0.08), Inches(grid_w - 0.28), Inches(grid_h - 0.16))
    tf_g = bx_g.text_frame
    tf_g.word_wrap = True
    tf_g.margin_left = tf_g.margin_top = tf_g.margin_right = tf_g.margin_bottom = 0
    
    # En-tête : Tag + Titre
    p_hdr = tf_g.paragraphs[0]
    r_tag = p_hdr.add_run()
    r_tag.text = gdata["tag"] + "  |  "
    r_tag.font.name = "Arial"
    r_tag.font.size = Pt(8.0)
    r_tag.font.bold = True
    r_tag.font.color.rgb = gdata["color"]
    
    r_tit = p_hdr.add_run()
    r_tit.text = gdata["title"]
    r_tit.font.name = "Arial"
    r_tit.font.size = Pt(10.5)
    r_tit.font.bold = True
    r_tit.font.color.rgb = C_NAVY
    
    # 4 règles détaillées par groupe
    for rh, rb in gdata["rules"]:
        p_r = tf_g.add_paragraph()
        p_r.space_before = Pt(2.5)
        rr_h = p_r.add_run()
        rr_h.text = "• " + rh + " : "
        rr_h.font.name = "Arial"
        rr_h.font.size = Pt(8.5)
        rr_h.font.bold = True
        rr_h.font.color.rgb = C_DARK
        
        rr_b = p_r.add_run()
        rr_b.text = rb
        rr_b.font.name = "Arial"
        rr_b.font.size = Pt(8.0)
        rr_b.font.color.rgb = C_MUTED

# Bandeau de synthèse officiel en bas
b_syn26 = s26.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.90), Inches(11.23), Inches(0.80))
b_syn26.adjustments[0] = 0.03
b_syn26.fill.solid()
b_syn26.fill.fore_color.rgb = C_CALLOUT_BG
b_syn26.line.color.rgb = C_ACCENT
b_syn26.line.width = Pt(1.0)

bx_s26 = s26.shapes.add_textbox(Inches(1.20), Inches(5.95), Inches(10.93), Inches(0.70))
tf_s26 = bx_s26.text_frame
tf_s26.word_wrap = True
tf_s26.margin_left = tf_s26.margin_top = tf_s26.margin_right = tf_s26.margin_bottom = 0
ps26_1 = tf_s26.paragraphs[0]
ps26_1.text = "SÉCURISATION GLOBALE PAR LE MOTEUR DES 17 RÈGLES MÉTIER :"
rs26_1 = ps26_1.runs[0]
rs26_1.font.name = "Arial"
rs26_1.font.size = Pt(9.5)
rs26_1.font.bold = True
rs26_1.font.color.rgb = C_ACCENT

ps26_2 = tf_s26.add_paragraph()
ps26_2.space_before = Pt(2)
rs26_2 = ps26_2.add_run()
rs26_2.text = "L'ensemble des 17 règles s'exécute automatiquement avant tout versement. Tout dossier en défaut reçoit une pastille d'avertissement et est soumis à l'arbitrage immédiat de l'opérateur en un clic."
rs26_2.font.name = "Arial"
rs26_2.font.size = Pt(8.5)
rs26_2.font.color.rgb = C_DARK

'''

if slide26_start != -1 and slide26_end != -1:
    code = code[:slide26_start] + new_slide26 + code[slide26_end:]
    print("3. Slide 26 updated successfully.")
else:
    print("WARNING: Slide 26 boundaries not found.")

# 4. Update Slide 34 (7.2 Perspectives): align with 07_Limites_et_Perspectives.tex
slide34_start = code.find('# Slide 34 (7.2) : Perspectives d\'évolution')
slide34_end = code.find('# ==================== CHAPITRE 8')

new_slide34 = '''# Slide 34 (7.2) : Perspectives d'évolution
print("Génération Slide 34 (Perspectives réelles du mémoire : FastAPI/HTML, WFS DGFiP, Transposabilité)...")
s34, c34 = init_standard_slide(6, 1, 34, "7.2 Perspectives d'évolution et feuille de route technique")

perspectives_data = [
    ("ÉVOLUTION ERGONOMIQUE", "MIGRATION VERS FASTAPI ET FRONTEND HTML SANS RECHARGEMENT",
     "Suppression du rechargement complet de page entre documents pour fluidifier le traitement des grands lots.\n"
     "Sauvegarde automatique continue des saisies de l'opérateur, éliminant tout risque de perte de corrections en cours.",
     C_NAVY),
    ("INTÉGRATION SIG DIRECTE", "CONNEXION AUX SERVICES WFS DGFIP ET PLUGIN QGIS MÉTIER",
     "Calcul automatique de la géométrie de parcelle via les flux WFS officiels, évitant le placement manuel du marqueur.\n"
     "Développement d'un plugin QGIS pour offrir aux géomètres l'accès direct aux archives depuis leur outil de production.",
     C_ACCENT),
    ("TRANSPOSABILITÉ CABINETS", "DÉCOUPLAGE MODULAIRE ET ADAPTATION À D'AUTRES FONDS",
     "Architecture scindée en étapes indépendantes : simple fichier de configuration pour cartouches et communes.\n"
     "Réentraînement ciblé de YOLOv8 sur les nouvelles mises en page sans redéveloppement logiciel complet.",
     C_SUCCESS)
]

y_p = 1.60
p_height = 1.25
gap_p = 0.15

for pi, (p_tag, p_title, p_body, p_color) in enumerate(perspectives_data):
    py = y_p + pi * (p_height + gap_p)
    
    b_persp = s34.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(py), Inches(11.23), Inches(p_height))
    b_persp.adjustments[0] = 0.03
    b_persp.fill.solid()
    b_persp.fill.fore_color.rgb = C_BG_SLIDE
    b_persp.line.color.rgb = p_color
    b_persp.line.width = Pt(1.2)
    
    b_badge = s34.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.20), Inches(py + 0.12), Inches(2.30), Inches(0.30))
    b_badge.adjustments[0] = 0.10
    b_badge.fill.solid()
    b_badge.fill.fore_color.rgb = p_color
    b_badge.line.fill.background()
    tf_pb = b_badge.text_frame
    tf_pb.margin_left = tf_pb.margin_top = tf_pb.margin_right = tf_pb.margin_bottom = 0
    ppb = tf_pb.paragraphs[0]
    ppb.text = p_tag
    ppb.alignment = PP_ALIGN.CENTER
    rpb = ppb.runs[0]
    rpb.font.name = "Arial"
    rpb.font.size = Pt(8.0)
    rpb.font.bold = True
    rpb.font.color.rgb = C_WHITE
    
    bx_ptit = s34.shapes.add_textbox(Inches(3.65), Inches(py + 0.10), Inches(8.40), Inches(0.32))
    tf_ptit = bx_ptit.text_frame
    tf_ptit.word_wrap = True
    tf_ptit.margin_left = tf_ptit.margin_top = tf_ptit.margin_right = tf_ptit.margin_bottom = 0
    ppt = tf_ptit.paragraphs[0]
    ppt.text = p_title
    rpt = ppt.runs[0]
    rpt.font.name = "Arial"
    rpt.font.size = Pt(10.5)
    rpt.font.bold = True
    rpt.font.color.rgb = C_NAVY
    
    bx_pbod = s34.shapes.add_textbox(Inches(1.20), Inches(py + 0.48), Inches(10.85), Inches(p_height - 0.52))
    tf_pbod = bx_pbod.text_frame
    tf_pbod.word_wrap = True
    lines_pb = p_body.split("\\n")
    for li_idx, line_txt in enumerate(lines_pb):
        p_line = tf_pbod.paragraphs[0] if li_idx == 0 else tf_pbod.add_paragraph()
        if li_idx > 0:
            p_line.space_before = Pt(2)
        r_line = p_line.add_run()
        r_line.text = line_txt
        r_line.font.name = "Arial"
        r_line.font.size = Pt(8.5)
        r_line.font.color.rgb = C_DARK if li_idx == 0 else C_MUTED

b_call34 = s34.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.85), Inches(11.23), Inches(0.80))
b_call34.adjustments[0] = 0.03
b_call34.fill.solid()
b_call34.fill.fore_color.rgb = C_CALLOUT_BG
b_call34.line.color.rgb = C_ACCENT
b_call34.line.width = Pt(1.0)

bx_c34 = s34.shapes.add_textbox(Inches(1.20), Inches(5.95), Inches(10.93), Inches(0.60))
tf_c34 = bx_c34.text_frame
p_c34 = tf_c34.paragraphs[0]
p_c34.alignment = PP_ALIGN.CENTER
p_c34.text = "Boucle d'apprentissage actif : Réinjection continue des corrections de l'opérateur pour enrichir les jeux d'entraînement locaux sans aucune fuite de données."
r_c34 = p_c34.runs[0]
r_c34.font.name = "Arial"
r_c34.font.size = Pt(9.5)
r_c34.font.bold = True
r_c34.font.color.rgb = C_NAVY

'''

if slide34_start != -1 and slide34_end != -1:
    code = code[:slide34_start] + new_slide34 + code[slide34_end:]
    print("4. Slide 34 updated successfully.")
else:
    print("WARNING: Slide 34 boundaries not found.")

# 5. Update Slide 36 (8.1 Bilan général du PFE): align with 08_Conclusion.tex
slide36_start = code.find('# Slide 36 (8.1) : Bilan général')
slide36_end = code.find('# Slide 37 (8.2) : Remerciements')

new_slide36 = '''# Slide 36 (8.1) : Bilan général
print("Génération Slide 36 (Bilan rigoureux conforme au chapitre 8 du mémoire)...")
s36, c36 = init_standard_slide(7, 0, 36, "8.1 Bilan général du Projet de Fin d'Études")

kpis_top36 = [
    ("23 600", "DOSSIERS VALORISABLES AU SIÈGE", C_NAVY),
    ("100% LOCAL", "RESPECT DU SECRET PROFESSIONNEL", C_ACCENT),
    ("FACTEUR 6", "GAIN DE PRODUCTIVITÉ CONSTATÉ", C_SUCCESS),
    ("0 €", "COÛT D'INFRASTRUCTURE LOGICIELLE", RGBColor(180, 83, 9))
]

kw36 = (11.23 - 3 * 0.15) / 4.0
kh36 = 0.95
y_top36 = 1.60

for ki, (kval, klbl, kcol) in enumerate(kpis_top36):
    kx = 1.05 + ki * (kw36 + 0.15)
    b_kpi = s36.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(kx), Inches(y_top36), Inches(kw36), Inches(kh36))
    b_kpi.adjustments[0] = 0.05
    b_kpi.fill.solid()
    b_kpi.fill.fore_color.rgb = C_BG_SLIDE
    b_kpi.line.color.rgb = kcol
    b_kpi.line.width = Pt(1.2)
    
    bx_k = s36.shapes.add_textbox(Inches(kx + 0.06), Inches(y_top36 + 0.10), Inches(kw36 - 0.12), Inches(kh36 - 0.20))
    tf_k = bx_k.text_frame
    tf_k.word_wrap = True
    tf_k.margin_left = tf_k.margin_top = tf_k.margin_right = tf_k.margin_bottom = 0
    
    pk1 = tf_k.paragraphs[0]
    pk1.text = kval
    pk1.alignment = PP_ALIGN.CENTER
    rk1 = pk1.runs[0]
    rk1.font.name = "Arial"
    rk1.font.size = Pt(14.0)
    rk1.font.bold = True
    rk1.font.color.rgb = kcol
    
    pk2 = tf_k.add_paragraph()
    pk2.space_before = Pt(2)
    pk2.text = klbl
    pk2.alignment = PP_ALIGN.CENTER
    rk2 = pk2.runs[0]
    rk2.font.name = "Arial"
    rk2.font.size = Pt(7.5)
    rk2.font.bold = True
    rk2.font.color.rgb = C_DARK

pw36 = (11.23 - 0.25) / 2.0
py36 = 2.70
ph36 = 3.00

p_left36 = s36.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(py36), Inches(pw36), Inches(ph36))
p_left36.adjustments[0] = 0.03
p_left36.fill.solid()
p_left36.fill.fore_color.rgb = C_BG_SLIDE
p_left36.line.color.rgb = C_LINE
p_left36.line.width = Pt(0.75)

bx_pl = s36.shapes.add_textbox(Inches(1.20), Inches(py36 + 0.12), Inches(pw36 - 0.30), Inches(ph36 - 0.24))
tf_pl = bx_pl.text_frame
tf_pl.word_wrap = True
tf_pl.margin_left = tf_pl.margin_top = tf_pl.margin_right = tf_pl.margin_bottom = 0

ppl_h = tf_pl.paragraphs[0]
ppl_h.text = "Apports opérationnels pour le cabinet GEO-SIAPP\\n"
rpl_h = ppl_h.runs[0]
rpl_h.font.name = "Arial"
rpl_h.font.size = Pt(11.0)
rpl_h.font.bold = True
rpl_h.font.color.rgb = C_NAVY

items_pl = [
    ("Sécurisation juridique des antériorités", "Recherche systématique garantie et respect absolu de la règle « Bornage sur bornage ne vaut »."),
    ("Gain de productivité constaté", "Temps de traitement ramené de 25 minutes de fouille physique à moins de 3 minutes par dossier."),
    ("Valorisation de 50 ans d'archives privées", "Transformation d'un fonds papier dormant en un actif cartographique directement exploitable sur le terrain."),
    ("Deux modes d'exploitation pérennes", "Mode interactif unitaire pour les affaires en cours et mode batch nocturne pour le traitement de masse.")
]
for ih, ib in items_pl:
    p_i = tf_pl.add_paragraph()
    p_i.space_before = Pt(4)
    r1 = p_i.add_run()
    r1.text = "✓ " + ih + " : "
    r1.font.name = "Arial"
    r1.font.size = Pt(8.5)
    r1.font.bold = True
    r1.font.color.rgb = C_ACCENT
    r2 = p_i.add_run()
    r2.text = ib
    r2.font.name = "Arial"
    r2.font.size = Pt(8.0)
    r2.font.color.rgb = C_DARK

p_right36 = s36.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05 + pw36 + 0.25), Inches(py36), Inches(pw36), Inches(ph36))
p_right36.adjustments[0] = 0.03
p_right36.fill.solid()
p_right36.fill.fore_color.rgb = C_BG_SLIDE
p_right36.line.color.rgb = C_LINE
p_right36.line.width = Pt(0.75)

bx_pr = s36.shapes.add_textbox(Inches(1.20 + pw36 + 0.25), Inches(py36 + 0.12), Inches(pw36 - 0.30), Inches(ph36 - 0.24))
tf_pr = bx_pr.text_frame
tf_pr.word_wrap = True
tf_pr.margin_left = tf_pr.margin_top = tf_pr.margin_right = tf_pr.margin_bottom = 0

ppr_h = tf_pr.paragraphs[0]
ppr_h.text = "Rigueur déontologique et posture d'ingénieur INSA\\n"
rpr_h = ppr_h.runs[0]
rpr_h.font.name = "Arial"
rpr_h.font.size = Pt(11.0)
rpr_h.font.bold = True
rpr_h.font.color.rgb = C_NAVY

items_pr = [
    ("Conformité à la Charte IA de l'OGE (2026)", "Respect strict des 4 principes directeurs : transparence, protection des données, traçabilité et responsabilité."),
    ("Responsabilité juridique entière conservée", "L'algorithme assiste l'opérateur par pré-remplissage ; l'humain valide chaque acte contradictoirement."),
    ("Zéro dépendance cloud et exécution locale", "Traitement direct sur les stations existantes sans abonnement tiers ni diffusion externe des plans."),
    ("Cadrage d'un projet industriel complet", "De l'analyse sur archives physiques dégradées jusqu'à l'injection certifiée sur le portail national Géofoncier.")
]
for ih, ib in items_pr:
    p_i = tf_pr.add_paragraph()
    p_i.space_before = Pt(4)
    r1 = p_i.add_run()
    r1.text = "✓ " + ih + " : "
    r1.font.name = "Arial"
    r1.font.size = Pt(8.5)
    r1.font.bold = True
    r1.font.color.rgb = C_SUCCESS
    r2 = p_i.add_run()
    r2.text = ib
    r2.font.name = "Arial"
    r2.font.size = Pt(8.0)
    r2.font.color.rgb = C_DARK

b_bot36 = s36.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.85), Inches(11.23), Inches(0.80))
b_bot36.adjustments[0] = 0.03
b_bot36.fill.solid()
b_bot36.fill.fore_color.rgb = C_CALLOUT_BG
b_bot36.line.color.rgb = C_ACCENT
b_bot36.line.width = Pt(1.0)

bx_bot36 = s36.shapes.add_textbox(Inches(1.20), Inches(5.92), Inches(10.93), Inches(0.66))
tf_bot36 = bx_bot36.text_frame
tf_bot36.word_wrap = True
tf_bot36.margin_left = tf_bot36.margin_top = tf_bot36.margin_right = tf_bot36.margin_bottom = 0
p_bot36 = tf_bot36.paragraphs[0]
p_bot36.alignment = PP_ALIGN.CENTER
p_bot36.text = "SYNTHÈSE DU PROJET DE FIN D'ÉTUDES :"
r_bot36 = p_bot36.runs[0]
r_bot36.font.name = "Arial"
r_bot36.font.size = Pt(9.5)
r_bot36.font.bold = True
r_bot36.font.color.rgb = C_ACCENT

p_bot36_2 = tf_bot36.add_paragraph()
p_bot36_2.space_before = Pt(2)
p_bot36_2.alignment = PP_ALIGN.CENTER
p_bot36_2.text = "Démonstration qu'un cabinet de géomètre-expert peut, avec des outils ouverts et locaux, moderniser ses processus tout en garantissant la sécurité et la pérennité de l'information foncière."
r_bot36_2 = p_bot36_2.runs[0]
r_bot36_2.font.name = "Arial"
r_bot36_2.font.size = Pt(8.5)
r_bot36_2.font.color.rgb = C_NAVY

'''

if slide36_start != -1 and slide36_end != -1:
    code = code[:slide36_start] + new_slide36 + code[slide36_end:]
    print("5. Slide 36 updated successfully.")
else:
    print("WARNING: Slide 36 boundaries not found.")

with open('generate_final_180s_deck.py', 'w', encoding='utf-8') as f:
    f.write(code)

print("\nAll updates applied successfully to generate_final_180s_deck.py!")
