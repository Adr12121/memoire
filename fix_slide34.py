# -*- coding: utf-8 -*-
with open('generate_final_180s_deck.py', 'r', encoding='utf-8') as f:
    text = f.read()

s34_start = text.find('# Slide 34 (7.2) : Perspectives d\'évolution')
s34_end = text.find('# ==================== CHAPITRE 8')

clean_slide34 = '''# Slide 34 (7.2) : Perspectives d'évolution
print("Génération Slide 34 (Perspectives réelles du mémoire : FastAPI/HTML, WFS DGFiP, Transposabilité)...")
s34, c34 = init_standard_slide(6, 1, 34, "7.2 Perspectives d'évolution et feuille de route technique")

perspectives_data = [
    ("ÉVOLUTION ERGONOMIQUE", "MIGRATION VERS FASTAPI ET FRONTEND HTML SANS RECHARGEMENT",
     [
         "Suppression du rechargement complet de page entre documents pour fluidifier le traitement des grands lots.",
         "Sauvegarde automatique continue des saisies de l'opérateur, éliminant tout risque de perte de corrections en cours."
     ],
     C_NAVY),
    ("INTÉGRATION SIG DIRECTE", "CONNEXION AUX SERVICES WFS DGFIP ET PLUGIN QGIS MÉTIER",
     [
         "Calcul automatique de la géométrie de parcelle via les flux WFS officiels, évitant le placement manuel du marqueur.",
         "Développement d'un plugin QGIS pour offrir aux géomètres l'accès direct aux archives depuis leur outil de production."
     ],
     C_ACCENT),
    ("TRANSPOSABILITÉ CABINETS", "DÉCOUPLAGE MODULAIRE ET ADAPTATION À D'AUTRES FONDS",
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

text = text[:s34_start] + clean_slide34 + text[s34_end:]
with open('generate_final_180s_deck.py', 'w', encoding='utf-8') as f:
    f.write(text)

print("Slide 34 fixed successfully!")
