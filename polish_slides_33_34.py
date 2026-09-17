# -*- coding: utf-8 -*-
with open('generate_final_180s_deck.py', 'r', encoding='utf-8') as f:
    text = f.read()

# Polish Slide 33
s33_start = text.find('# Slide 33 (7.1) : Limites techniques')
s33_end = text.find('# Slide 34 (7.2) : Perspectives d\'évolution')

new_slide33 = '''# Slide 33 (7.1) : Limites techniques
print("Génération Slide 33 (Limites techniques et diagnostic physique)...")
s33, c33 = init_standard_slide(6, 0, 33, "7.1 Limites techniques et analyse des modes de défaillance")

left_x33 = 1.05
left_w33 = 5.50

limites_diagnostic = [
    ("Supports calques & encres altérées",
     "Transparence recto/verso des calques créant des écritures superposées. Déformation des pages près de la reliure des registres.",
     C_WARN),
    ("Tampons administratifs recouvrants",
     "Tampons de réception rouge/bleu posés sur les numéros de parcelles. Confusions récurrentes sur les chiffres visuellement proches (3/8, 0/6, 1/7).",
     C_RED),
    ("Disparité d'écriture & Répertoire B",
     "Écriture manuscrite libre sans colonnes prédéfinies. Taux de réussite inférieur de 15 à 20% par rapport au Géomètre A, exigeant une reprise manuelle.",
     C_ACCENT)
]

y_lim = 1.60
lim_h = 1.25
gap_lim = 0.14

for li, (l_head, l_body, l_col) in enumerate(limites_diagnostic):
    ly = y_lim + li * (lim_h + gap_lim)
    b_lim = s33.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left_x33), Inches(ly), Inches(left_w33), Inches(lim_h))
    b_lim.adjustments[0] = 0.03
    b_lim.fill.solid()
    b_lim.fill.fore_color.rgb = C_BG_SLIDE
    b_lim.line.color.rgb = l_col
    b_lim.line.width = Pt(1.0)
    
    bx_l = s33.shapes.add_textbox(Inches(left_x33 + 0.18), Inches(ly + 0.12), Inches(left_w33 - 0.36), Inches(lim_h - 0.24))
    tf_l = bx_l.text_frame
    tf_l.word_wrap = True
    tf_l.margin_left = tf_l.margin_top = tf_l.margin_right = tf_l.margin_bottom = 0
    
    pl1 = tf_l.paragraphs[0]
    pl1.text = "• " + l_head + " :"
    rl1 = pl1.runs[0]
    rl1.font.name = "Arial"
    rl1.font.size = Pt(10.5)
    rl1.font.bold = True
    rl1.font.color.rgb = C_NAVY
    
    pl2 = tf_l.add_paragraph()
    pl2.space_before = Pt(4)
    rl2 = pl2.add_run()
    rl2.text = l_body
    rl2.font.name = "Arial"
    rl2.font.size = Pt(8.5)
    rl2.font.color.rgb = C_DARK

b_cpu = s33.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left_x33), Inches(5.82), Inches(left_w33), Inches(0.85))
b_cpu.adjustments[0] = 0.03
b_cpu.fill.solid()
b_cpu.fill.fore_color.rgb = C_CALLOUT_BG
b_cpu.line.color.rgb = C_ACCENT
b_cpu.line.width = Pt(1.0)

bx_cpu = s33.shapes.add_textbox(Inches(left_x33 + 0.18), Inches(5.88), Inches(left_w33 - 0.36), Inches(0.72))
tf_cpu = bx_cpu.text_frame
tf_cpu.word_wrap = True
tf_cpu.margin_left = tf_cpu.margin_top = tf_cpu.margin_right = tf_cpu.margin_bottom = 0
pc1 = tf_cpu.paragraphs[0]
pc1.text = "CONTRAINTE MATÉRIELLE CPU :"
rc1 = pc1.runs[0]
rc1.font.name = "Arial"
rc1.font.size = Pt(9.0)
rc1.font.bold = True
rc1.font.color.rgb = C_ACCENT

pc2 = tf_cpu.add_paragraph()
pc2.space_before = Pt(2)
rc2 = pc2.add_run()
rc2.text = "L'absence de GPU dédié allonge le temps d'inférence VLM (~5s à 8s sur CPU), justifiant de réserver l'arbitre aux seules ambiguïtés résiduelles."
rc2.font.name = "Arial"
rc2.font.size = Pt(8.0)
rc2.font.color.rgb = C_DARK

rx33 = 6.78
rw33 = 5.50

frame_deg = s33.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(rx33), Inches(1.60), Inches(rw33), Inches(3.20))
frame_deg.adjustments[0] = 0.03
frame_deg.fill.solid()
frame_deg.fill.fore_color.rgb = C_BG_SLIDE
frame_deg.line.color.rgb = C_LINE
frame_deg.line.width = Pt(0.75)
add_fitted_picture(s33, "img/crop_serret.png", rx33 + 0.10, 1.70, rw33 - 0.20, 3.00, caption="Extrait d'archive dégradée : encres pâles et ratures manuelles (Mémoire)")

b_val33 = s33.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(rx33), Inches(4.95), Inches(rw33), Inches(1.72))
b_val33.adjustments[0] = 0.03
b_val33.fill.solid()
b_val33.fill.fore_color.rgb = C_BG_SLIDE
b_val33.line.color.rgb = C_LINE
b_val33.line.width = Pt(0.75)

bx_vt = s33.shapes.add_textbox(Inches(rx33 + 0.18), Inches(5.05), Inches(rw33 - 0.36), Inches(1.52))
tf_vt = bx_vt.text_frame
tf_vt.word_wrap = True
tf_vt.margin_left = tf_vt.margin_top = tf_vt.margin_right = tf_vt.margin_bottom = 0

pvt_h = tf_vt.paragraphs[0]
pvt_h.text = "RÉPARTITION DES RETOUCHES OPÉRATEUR :"
rvt_h = pvt_h.runs[0]
rvt_h.font.name = "Arial"
rvt_h.font.size = Pt(9.5)
rvt_h.font.bold = True
rvt_h.font.color.rgb = C_NAVY

retouches_stats = [
    ("82% Validés sans retouche", "Pré-remplissage conforme en 1 clic (DMPC Cerfa).", C_SUCCESS),
    ("15% Retouche d'un champ", "Ajustement d'un chiffre sous tampon ou date pâlie.", C_WARN),
    ("3% Saisie intégrale", "Documents calques très altérés ou déchirés.", C_RED)
]
for r_stat, r_desc, r_col in retouches_stats:
    p_r = tf_vt.add_paragraph()
    p_r.space_before = Pt(3)
    rr1 = p_r.add_run()
    rr1.text = "• " + r_stat + " : "
    rr1.font.name = "Arial"
    rr1.font.size = Pt(8.5)
    rr1.font.bold = True
    rr1.font.color.rgb = r_col
    rr2 = p_r.add_run()
    rr2.text = r_desc
    rr2.font.name = "Arial"
    rr2.font.size = Pt(8.0)
    rr2.font.color.rgb = C_MUTED

'''

text = text[:s33_start] + new_slide33 + text[s33_end:]

# Polish Slide 34
s34_start = text.find('# Slide 34 (7.2) : Perspectives d\'évolution')
s34_end = text.find('# ==================== CHAPITRE 8')

new_slide34 = '''# Slide 34 (7.2) : Perspectives d'évolution
print("Génération Slide 34 (Perspectives réelles du mémoire en typographie soignée)...")
s34, c34 = init_standard_slide(6, 1, 34, "7.2 Perspectives d'évolution et feuille de route technique")

perspectives_data = [
    ("Évolution ergonomique", "Migration vers FastAPI et frontend HTML sans rechargement",
     [
         "Suppression du rechargement complet de page entre documents pour fluidifier le traitement des grands lots.",
         "Sauvegarde automatique continue des saisies de l'opérateur, éliminant tout risque de perte de corrections en cours."
     ],
     C_NAVY),
    ("Intégration SIG", "Connexion aux services WFS DGFiP et plugin QGIS métier",
     [
         "Calcul automatique de la géométrie de parcelle via les flux WFS officiels, évitant le placement manuel du marqueur.",
         "Développement d'un plugin QGIS pour offrir aux géomètres l'accès direct aux archives depuis leur outil de production."
     ],
     C_ACCENT),
    ("Transposabilité cabinets", "Découplage modulaire et adaptation à d'autres fonds",
     [
         "Architecture scindée en étapes indépendantes : simple fichier de configuration pour cartouches et communes.",
         "Réentraînement ciblé de YOLOv8 sur les nouvelles mises en page sans redéveloppement logiciel complet."
     ],
     C_SUCCESS)
]

y_p = 1.60
p_height = 1.25
gap_p = 0.15

for pi, (p_tag, p_title, p_lines, p_color) in enumerate(perspectives_data):
    py = y_p + pi * (p_height + gap_p)
    
    b_persp = s34.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(py), Inches(11.23), Inches(p_height))
    b_persp.adjustments[0] = 0.03
    b_persp.fill.solid()
    b_persp.fill.fore_color.rgb = C_BG_SLIDE
    b_persp.line.color.rgb = p_color
    b_persp.line.width = Pt(1.2)
    
    b_badge = s34.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.20), Inches(py + 0.12), Inches(2.10), Inches(0.30))
    b_badge.adjustments[0] = 0.12
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
    rpb.font.size = Pt(8.5)
    rpb.font.bold = True
    rpb.font.color.rgb = C_WHITE
    
    bx_ptit = s34.shapes.add_textbox(Inches(3.45), Inches(py + 0.10), Inches(8.60), Inches(0.32))
    tf_ptit = bx_ptit.text_frame
    tf_ptit.word_wrap = True
    tf_ptit.margin_left = tf_ptit.margin_top = tf_ptit.margin_right = tf_ptit.margin_bottom = 0
    ppt = tf_ptit.paragraphs[0]
    ppt.text = p_title
    rpt = ppt.runs[0]
    rpt.font.name = "Arial"
    rpt.font.size = Pt(11.0)
    rpt.font.bold = True
    rpt.font.color.rgb = C_NAVY
    
    bx_pbod = s34.shapes.add_textbox(Inches(1.20), Inches(py + 0.48), Inches(10.85), Inches(p_height - 0.52))
    tf_pbod = bx_pbod.text_frame
    tf_pbod.word_wrap = True
    for li_idx, line_txt in enumerate(p_lines):
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

bx_c34 = s34.shapes.add_textbox(Inches(1.20), Inches(5.92), Inches(10.93), Inches(0.66))
tf_c34 = bx_c34.text_frame
tf_c34.word_wrap = True
tf_c34.margin_left = tf_c34.margin_top = tf_c34.margin_right = tf_c34.margin_bottom = 0
p_c34 = tf_c34.paragraphs[0]
p_c34.alignment = PP_ALIGN.CENTER
p_c34.text = "BOUCLE D'APPRENTISSAGE ACTIF EN LOCAL (Figure 7.2 du mémoire) :"
r_c34 = p_c34.runs[0]
r_c34.font.name = "Arial"
r_c34.font.size = Pt(9.5)
r_c34.font.bold = True
r_c34.font.color.rgb = C_ACCENT

p_c34_2 = tf_c34.add_paragraph()
p_c34_2.space_before = Pt(2)
p_c34_2.alignment = PP_ALIGN.CENTER
p_c34_2.text = "Réinjection continue des corrections de l'opérateur pour enrichir les jeux d'entraînement locaux sans aucune fuite de données vers l'extérieur."
r_c34_2 = p_c34_2.runs[0]
r_c34_2.font.name = "Arial"
r_c34_2.font.size = Pt(8.5)
r_c34_2.font.color.rgb = C_NAVY

'''

text = text[:s34_start] + new_slide34 + text[s34_end:]

with open('generate_final_180s_deck.py', 'w', encoding='utf-8') as f:
    f.write(text)

print("Slides 33 and 34 polished successfully!")
