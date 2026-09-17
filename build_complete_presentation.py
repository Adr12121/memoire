# -*- coding: utf-8 -*-
"""
build_complete_presentation.py
Générateur complet de la soutenance de PFE d'Adrien TRAVAILLÉ (INSA Strasbourg / GEO-SIAPP).
Strictement conforme aux directives de l'utilisateur :
1. Une seule identité visuelle sobre et élégante (Bleu Nuit #0F2546, Bleu Accent #0284C7, Blanc, Ardoise).
2. Zéro jeu de 60 couleurs différentes : transitions uniformes en Bleu Nuit, frise uniforme avec seul le chapitre actif en surbrillance.
3. Vidéo montée Demo_soutenance_final.mp4 intégrée.
4. Juste milieu de texte : contenu structuré, chiffres clés, pas de vide béant, pas de pavé indigeste.
5. Respect strict du plan du mémoire et des règles de rédaction (pas de tiret cadratin).
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image

# ==================== COULEURS OFFICIELLES ====================
C_NAVY      = RGBColor(15, 37, 70)     # #0F2546 - Bleu Nuit Institutionnel
C_NAVY_DARK = RGBColor(10, 25, 48)     # #0A1930 - Fond Frise
C_ACCENT    = RGBColor(2, 132, 199)    # #0284C7 - Bleu Technique Soutenance
C_DARK      = RGBColor(15, 23, 42)     # #0F172A - Texte Principal
C_MUTED     = RGBColor(71, 85, 105)    # #475569 - Texte Secondaire
C_LIGHT_BG  = RGBColor(248, 250, 252)  # #F8FAFC - Fond doux
C_CARD_BG   = RGBColor(255, 255, 255)  # Blanc Pur
C_BORDER    = RGBColor(226, 232, 240)  # #E2E8F0 - Bordure cartes
C_BOX_BG    = RGBColor(240, 249, 255)  # #F0F9FF - Fond callout bleu très pâle
C_WHITE     = RGBColor(255, 255, 255)
C_SUCCESS   = RGBColor(22, 163, 74)    # #16A34A - Succès
C_WARN      = RGBColor(217, 119, 6)    # #D97706 - Attention
C_RED       = RGBColor(220, 38, 38)    # #DC2626 - Rejet

TOTAL_SLIDES = 32
SLIDE_W = 13.333
SLIDE_H = 7.50

CHAP_SHORT = [
    "1. INTRO & CONTEXTE",
    "2. ANALYSE MÉTIER",
    "3. ÉTAT DE L'ART",
    "4. ARCHITECTURE",
    "5. FIABILISATION",
    "6. GÉOFONCIER",
    "7. LIMITES",
    "8. CONCLUSION"
]

prs = Presentation()
prs.slide_width = Inches(SLIDE_W)
prs.slide_height = Inches(SLIDE_H)
blank_layout = prs.slide_layouts[6]

# ==================== FONCTIONS UTILITAIRES ====================

def set_bg(slide, color):
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = color

def add_footer(slide, current_num):
    bx = slide.shapes.add_textbox(Inches(12.00), Inches(7.16), Inches(1.15), Inches(0.26))
    tf = bx.text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = f"{current_num} / {TOTAL_SLIDES}"
    p.alignment = PP_ALIGN.RIGHT
    r = p.runs[0]
    r.font.name = "Arial"
    r.font.size = Pt(8.5)
    r.font.color.rgb = C_MUTED

def add_frieze(slide, active_chap_idx):
    cw = SLIDE_W / 8.0
    for i, label in enumerate(CHAP_SHORT):
        x = i * cw
        is_act = (i == active_chap_idx)
        shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(0.0), Inches(cw), Inches(0.34))
        shp.fill.solid()
        shp.fill.fore_color.rgb = C_ACCENT if is_act else C_NAVY_DARK
        shp.line.fill.background()
        
        bx = slide.shapes.add_textbox(Inches(x + 0.02), Inches(0.0), Inches(cw - 0.04), Inches(0.34))
        tf = bx.text_frame
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = label
        p.alignment = PP_ALIGN.CENTER
        r = p.runs[0]
        r.font.name = "Arial"
        r.font.size = Pt(7.5)
        r.font.bold = is_act
        r.font.color.rgb = C_WHITE if is_act else RGBColor(160, 180, 205)

def add_slide_header(slide, title, category_subtitle=None):
    if category_subtitle:
        bx_sub = slide.shapes.add_textbox(Inches(0.40), Inches(0.42), Inches(12.50), Inches(0.24))
        tf_sub = bx_sub.text_frame
        tf_sub.margin_left = tf_sub.margin_top = tf_sub.margin_right = tf_sub.margin_bottom = 0
        p_sub = tf_sub.paragraphs[0]
        p_sub.text = category_subtitle.upper()
        r_sub = p_sub.runs[0]
        r_sub.font.name = "Arial"
        r_sub.font.size = Pt(9.0)
        r_sub.font.bold = True
        r_sub.font.color.rgb = C_ACCENT

    bx_t = slide.shapes.add_textbox(Inches(0.40), Inches(0.66), Inches(12.50), Inches(0.48))
    tf_t = bx_t.text_frame
    tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
    p_t = tf_t.paragraphs[0]
    p_t.text = title
    r_t = p_t.runs[0]
    r_t.font.name = "Arial"
    r_t.font.size = Pt(18.5)
    r_t.font.bold = True
    r_t.font.color.rgb = C_DARK

    sep = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.40), Inches(1.18), Inches(12.53), Inches(0.025))
    sep.fill.solid()
    sep.fill.fore_color.rgb = C_ACCENT
    sep.line.fill.background()

def add_fitted_picture(slide, img_path, box_x, box_y, box_w, box_h, caption=None):
    if not img_path or not os.path.exists(img_path):
        return None
    try:
        cap_h = 0.32 if caption else 0.0
        avail_h = box_h - cap_h
        with Image.open(img_path) as im:
            iw, ih = im.size
        img_ratio = iw / ih
        box_ratio = box_w / avail_h
        if img_ratio > box_ratio:
            fw = box_w
            fh = box_w / img_ratio
        else:
            fh = avail_h
            fw = avail_h * img_ratio
        cx = box_x + (box_w - fw) / 2.0
        cy = box_y + (avail_h - fh) / 2.0
        slide.shapes.add_picture(img_path, Inches(cx), Inches(cy), Inches(fw), Inches(fh))
        if caption:
            bx = slide.shapes.add_textbox(Inches(box_x), Inches(cy + fh + 0.04), Inches(box_w), Inches(cap_h))
            tf = bx.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.text = caption
            p.alignment = PP_ALIGN.CENTER
            r = p.runs[0]
            r.font.name = "Arial"
            r.font.size = Pt(8.5)
            r.font.italic = True
            r.font.color.rgb = C_MUTED
    except Exception as e:
        print(f"Erreur image {img_path} : {e}")

def create_card(slide, x, y, w, h, bg_color=C_CARD_BG, border_color=C_BORDER, border_width=Pt(0.75)):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = bg_color
    if border_color:
        shp.line.color.rgb = border_color
        shp.line.width = border_width
    else:
        shp.line.fill.background()
    return shp

def add_transition_slide(chap_num, chap_title, chap_idx, slide_num, subtitle=None):
    slide = prs.slides.add_slide(blank_layout)
    set_bg(slide, C_NAVY)
    
    bx = slide.shapes.add_textbox(Inches(0.80), Inches(2.30), Inches(11.73), Inches(0.40))
    tf = bx.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = f"CHAPITRE {chap_num}"
    p.alignment = PP_ALIGN.CENTER
    r = p.runs[0]
    r.font.name = "Arial"
    r.font.size = Pt(12.0)
    r.font.bold = True
    r.font.color.rgb = C_ACCENT
    
    bx2 = slide.shapes.add_textbox(Inches(0.80), Inches(2.75), Inches(11.73), Inches(1.30))
    tf2 = bx2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = chap_title
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.runs[0]
    r2.font.name = "Arial"
    r2.font.size = Pt(32.0)
    r2.font.bold = True
    r2.font.color.rgb = C_WHITE
    
    sep = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.40), Inches(4.25), Inches(2.53), Inches(0.035))
    sep.fill.solid()
    sep.fill.fore_color.rgb = C_ACCENT
    sep.line.fill.background()
    
    if subtitle:
        bx3 = slide.shapes.add_textbox(Inches(0.80), Inches(4.45), Inches(11.73), Inches(0.60))
        tf3 = bx3.text_frame
        tf3.word_wrap = True
        p3 = tf3.paragraphs[0]
        p3.text = subtitle
        p3.alignment = PP_ALIGN.CENTER
        r3 = p3.runs[0]
        r3.font.name = "Arial"
        r3.font.size = Pt(13.0)
        r3.font.color.rgb = RGBColor(186, 230, 253)
        
    add_footer(slide, slide_num)
    return slide

def add_two_column_slide(chap_idx, category, title, items, img_path, img_caption, slide_num, callout_title=None, callout_text=None, kpis=None):
    slide = prs.slides.add_slide(blank_layout)
    set_bg(slide, C_LIGHT_BG)
    add_frieze(slide, chap_idx)
    add_slide_header(slide, title, category)
    add_footer(slide, slide_num)
    
    card_w = 5.95
    left_x = 0.40
    start_y = 1.34
    
    total_items = len(items)
    item_gap = 0.12
    bot_reserve = 1.10 if (callout_title or kpis) else 0.0
    avail_h = (7.05 - start_y) - bot_reserve
    item_h = (avail_h - (total_items - 1) * item_gap) / total_items
    
    for i, item in enumerate(items):
        iy = start_y + i * (item_h + item_gap)
        create_card(slide, left_x, iy, card_w, item_h, bg_color=C_CARD_BG, border_color=C_BORDER)
        
        bx = slide.shapes.add_textbox(Inches(left_x + 0.20), Inches(iy + 0.08), Inches(card_w - 0.40), Inches(item_h - 0.16))
        tf = bx.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        
        p = tf.paragraphs[0]
        if isinstance(item, tuple):
            head, body = item
            r_head = p.add_run()
            r_head.text = head + "\n"
            r_head.font.name = "Arial"
            r_head.font.size = Pt(11.0)
            r_head.font.bold = True
            r_head.font.color.rgb = C_NAVY
            
            p_body = tf.add_paragraph()
            p_body.space_before = Pt(3)
            r_body = p_body.add_run()
            r_body.text = body
            r_body.font.name = "Arial"
            r_body.font.size = Pt(9.5)
            r_body.font.color.rgb = C_MUTED
        else:
            r = p.add_run()
            r.text = item
            r.font.name = "Arial"
            r.font.size = Pt(10.0)
            r.font.color.rgb = C_DARK

    if callout_title and callout_text:
        cy = 7.05 - bot_reserve + 0.08
        create_card(slide, left_x, cy, card_w, bot_reserve - 0.08, bg_color=C_BOX_BG, border_color=C_ACCENT, border_width=Pt(1.0))
        bx_c = slide.shapes.add_textbox(Inches(left_x + 0.20), Inches(cy + 0.08), Inches(card_w - 0.40), Inches(bot_reserve - 0.24))
        tf_c = bx_c.text_frame
        tf_c.word_wrap = True
        tf_c.margin_left = tf_c.margin_top = tf_c.margin_right = tf_c.margin_bottom = 0
        p_c1 = tf_c.paragraphs[0]
        p_c1.text = callout_title
        r_c1 = p_c1.runs[0]
        r_c1.font.name = "Arial"
        r_c1.font.size = Pt(9.5)
        r_c1.font.bold = True
        r_c1.font.color.rgb = C_ACCENT
        
        p_c2 = tf_c.add_paragraph()
        p_c2.space_before = Pt(2)
        p_c2.text = callout_text
        r_c2 = p_c2.runs[0]
        r_c2.font.name = "Arial"
        r_c2.font.size = Pt(8.5)
        r_c2.font.color.rgb = C_DARK
    elif kpis:
        cy = 7.05 - bot_reserve + 0.08
        n_kpi = len(kpis)
        kw = (card_w - (n_kpi - 1) * 0.10) / n_kpi
        for ki, (kval, klbl) in enumerate(kpis):
            kx = left_x + ki * (kw + 0.10)
            create_card(slide, kx, cy, kw, bot_reserve - 0.08, bg_color=C_NAVY, border_color=None)
            bx_k = slide.shapes.add_textbox(Inches(kx + 0.05), Inches(cy + 0.10), Inches(kw - 0.10), Inches(bot_reserve - 0.28))
            tf_k = bx_k.text_frame
            tf_k.word_wrap = True
            tf_k.margin_left = tf_k.margin_top = tf_k.margin_right = tf_k.margin_bottom = 0
            pk1 = tf_k.paragraphs[0]
            pk1.text = kval
            pk1.alignment = PP_ALIGN.CENTER
            rk1 = pk1.runs[0]
            rk1.font.name = "Arial"
            rk1.font.size = Pt(16.0)
            rk1.font.bold = True
            rk1.font.color.rgb = C_ACCENT
            
            pk2 = tf_k.add_paragraph()
            pk2.space_before = Pt(1)
            pk2.text = klbl
            pk2.alignment = PP_ALIGN.CENTER
            rk2 = pk2.runs[0]
            rk2.font.name = "Arial"
            rk2.font.size = Pt(7.5)
            rk2.font.color.rgb = C_WHITE

    right_x = 6.60
    right_w = 6.33
    right_h = 5.70
    create_card(slide, right_x, start_y, right_w, right_h, bg_color=C_CARD_BG, border_color=C_BORDER)
    add_fitted_picture(slide, img_path, right_x + 0.15, start_y + 0.15, right_w - 0.30, right_h - 0.30, caption=img_caption)

    return slide

# ==================== SLIDE 1 : TITRE (STYLE PFE 180S) ====================
print("Génération Slide 1...")
s1 = prs.slides.add_slide(blank_layout)
set_bg(s1, C_LIGHT_BG)

# Logos en haut
if os.path.exists("Logo_INSAStrasbourg.jpg"):
    s1.shapes.add_picture("Logo_INSAStrasbourg.jpg", Inches(0.80), Inches(0.50), width=Inches(2.70))
if os.path.exists("Geosiapp.jpg"):
    # Cadrer le logo Geosiapp pour ne jamais chevaucher la carte blanche
    add_fitted_picture(s1, "Geosiapp.jpg", 10.60, 0.45, 1.93, 1.10)

# Grand conteneur arrondi épuré
card1 = create_card(s1, 0.80, 1.75, 11.73, 5.25, bg_color=C_CARD_BG, border_color=C_BORDER, border_width=Pt(1.0))

bx_t1 = s1.shapes.add_textbox(Inches(1.20), Inches(2.05), Inches(10.93), Inches(1.45))
tf_t1 = bx_t1.text_frame
tf_t1.word_wrap = True
tf_t1.margin_left = tf_t1.margin_top = tf_t1.margin_right = tf_t1.margin_bottom = 0
pt1 = tf_t1.paragraphs[0]
pt1.alignment = PP_ALIGN.CENTER
pt1.text = "DÉVELOPPEMENT D'UN OUTIL PERMETTANT LE TRAITEMENT ET L'INSERTION DES ARCHIVES NUMÉRIQUES SUR GÉOFONCIER"
rt1 = pt1.runs[0]
rt1.font.name = "Arial"
rt1.font.size = Pt(21.0)
rt1.font.bold = True
rt1.font.color.rgb = C_NAVY

bx_sub1 = s1.shapes.add_textbox(Inches(1.20), Inches(3.60), Inches(10.93), Inches(0.35))
tf_sub1 = bx_sub1.text_frame
tf_sub1.word_wrap = True
ps1 = tf_sub1.paragraphs[0]
ps1.alignment = PP_ALIGN.CENTER
ps1.text = "PROJET DE FIN D'ÉTUDES : INSA STRASBOURG"
rs1 = ps1.runs[0]
rs1.font.name = "Arial"
rs1.font.size = Pt(12.5)
rs1.font.bold = True
rs1.font.color.rgb = C_ACCENT

sep1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.66), Inches(4.05), Inches(4.00), Inches(0.025))
sep1.fill.solid()
sep1.fill.fore_color.rgb = C_BORDER
sep1.line.fill.background()

# Détails candidat & encadrement
bx_det1 = s1.shapes.add_textbox(Inches(1.20), Inches(4.25), Inches(10.93), Inches(2.40))
tf_det1 = bx_det1.text_frame
tf_det1.word_wrap = True
tf_det1.margin_left = tf_det1.margin_top = tf_det1.margin_right = tf_det1.margin_bottom = 0

pd1 = tf_det1.paragraphs[0]
pd1.alignment = PP_ALIGN.CENTER
rd1_1 = pd1.add_run()
rd1_1.text = "Étudiant : "
rd1_1.font.name = "Arial"
rd1_1.font.size = Pt(13.5)
rd1_1.font.bold = True
rd1_1.font.color.rgb = C_DARK
rd1_2 = pd1.add_run()
rd1_2.text = "TRAVAILLÉ Adrien  (Spécialité Topographie, Promotion 2026)"
rd1_2.font.name = "Arial"
rd1_2.font.size = Pt(13.5)
rd1_2.font.color.rgb = C_DARK

pd2 = tf_det1.add_paragraph()
pd2.space_before = Pt(8)
pd2.alignment = PP_ALIGN.CENTER
rd2 = pd2.add_run()
rd2.text = "Correcteur : M. Mathieu KOEHL (INSA Strasbourg / ICube)   •   Encadrant de PFE : M. Gaëtan HAGUE (GEO-SIAPP)"
rd2.font.name = "Arial"
rd2.font.size = Pt(10.5)
rd2.font.color.rgb = C_MUTED

pd3 = tf_det1.add_paragraph()
pd3.space_before = Pt(4)
pd3.alignment = PP_ALIGN.CENTER
rd3 = pd3.add_run()
rd3.text = "Entreprise d'accueil : Cabinet GEO-SIAPP   •   2 Avenue Jean Monnet, 07200 Aubenas (Ardèche)"
rd3.font.name = "Arial"
rd3.font.size = Pt(10.5)
rd3.font.color.rgb = C_MUTED

pd4 = tf_det1.add_paragraph()
pd4.space_before = Pt(8)
pd4.alignment = PP_ALIGN.CENTER
rd4 = pd4.add_run()
rd4.text = "Soutenance du 24 septembre 2026"
rd4.font.name = "Arial"
rd4.font.size = Pt(10.0)
rd4.font.italic = True
rd4.font.color.rgb = C_ACCENT

# ==================== SLIDE 2 : SOMMAIRE ====================
print("Génération Slide 2...")
s2 = prs.slides.add_slide(blank_layout)
set_bg(s2, C_NAVY)

bx_st2 = s2.shapes.add_textbox(Inches(0.80), Inches(0.40), Inches(11.73), Inches(0.28))
tf_st2 = bx_st2.text_frame
p_st2 = tf_st2.paragraphs[0]
p_st2.alignment = PP_ALIGN.CENTER
p_st2.text = "PLAN DE LA PRÉSENTATION"
r_st2 = p_st2.runs[0]
r_st2.font.name = "Arial"
r_st2.font.size = Pt(10.5)
r_st2.font.bold = True
r_st2.font.color.rgb = C_ACCENT

bx_s2 = s2.shapes.add_textbox(Inches(0.80), Inches(0.68), Inches(11.73), Inches(0.60))
tf_s2 = bx_s2.text_frame
p_s2 = tf_s2.paragraphs[0]
p_s2.alignment = PP_ALIGN.CENTER
p_s2.text = "Sommaire"
r_s2 = p_s2.runs[0]
r_s2.font.name = "Arial"
r_s2.font.size = Pt(28.0)
r_s2.font.bold = True
r_s2.font.color.rgb = C_WHITE

sommaire_items = [
    ("01", "Introduction & Contexte", "Structure d'accueil, cadre légal du bornage et problématique"),
    ("02", "Analyse Métier & Données", "Fonds d'archives SIAPP et spécifications de l'API Géofoncier"),
    ("03", "État de l'Art", "Technologies OCR, HTR (TrOCR), NER (GLiNER) et arbitrage VLM"),
    ("04", "Architecture & Pipeline", "Conception en 6 scripts modulaires, prétraitement et YOLOv8"),
    ("05", "Fiabilisation & Interface", "Application Streamlit, relecture assistée et règles de contrôle"),
    ("06", "Intégration Géofoncier", "Démonstration vidéo, injection API REST et résultats réels"),
    ("07", "Limites & Perspectives", "Contraintes de numérisation, fine-tuning et évolutions"),
    ("08", "Conclusion", "Bilan d'ingénieur, souveraineté des données et perspectives")
]

card_w2 = 2.80
card_h2 = 2.45
gap_x2 = 0.18
gap_y2 = 0.20
start_x2 = (SLIDE_W - (4 * card_w2 + 3 * gap_x2)) / 2.0
start_y2 = 1.45

for idx, (num, title_c, desc_c) in enumerate(sommaire_items):
    row = idx // 4
    col = idx % 4
    cx = start_x2 + col * (card_w2 + gap_x2)
    cy = start_y2 + row * (card_h2 + gap_y2)
    
    c_shp = create_card(s2, cx, cy, card_w2, card_h2, bg_color=RGBColor(23, 49, 88), border_color=RGBColor(38, 77, 133), border_width=Pt(1.0))
    
    bx_num = s2.shapes.add_textbox(Inches(cx + 0.16), Inches(cy + 0.14), Inches(card_w2 - 0.32), Inches(0.40))
    tf_num = bx_num.text_frame
    p_num = tf_num.paragraphs[0]
    p_num.text = num
    r_num = p_num.runs[0]
    r_num.font.name = "Arial"
    r_num.font.size = Pt(20.0)
    r_num.font.bold = True
    r_num.font.color.rgb = C_ACCENT
    
    bx_tc = s2.shapes.add_textbox(Inches(cx + 0.16), Inches(cy + 0.58), Inches(card_w2 - 0.32), Inches(0.70))
    tf_tc = bx_tc.text_frame
    tf_tc.word_wrap = True
    p_tc = tf_tc.paragraphs[0]
    p_tc.text = title_c
    r_tc = p_tc.runs[0]
    r_tc.font.name = "Arial"
    r_tc.font.size = Pt(11.5)
    r_tc.font.bold = True
    r_tc.font.color.rgb = C_WHITE
    
    bx_dc = s2.shapes.add_textbox(Inches(cx + 0.16), Inches(cy + 1.30), Inches(card_w2 - 0.32), Inches(1.00))
    tf_dc = bx_dc.text_frame
    tf_dc.word_wrap = True
    p_dc = tf_dc.paragraphs[0]
    p_dc.text = desc_c
    r_dc = p_dc.runs[0]
    r_dc.font.name = "Arial"
    r_dc.font.size = Pt(9.0)
    r_dc.font.color.rgb = RGBColor(186, 205, 230)

add_footer(s2, 2)

# ==================== SLIDE 3 : PROBLÉMATIQUE ====================
print("Génération Slide 3...")
add_two_column_slide(
    chap_idx=0,
    category="Introduction & Contexte",
    title="Problématique : la valorisation de 50 ans d'archives foncières",
    items=[
        ("Hétérogénéité & Écriture Manuscrite",
         "50 ans de registres physiques (1959-2007) mêlant encres fanées, abréviations locales et absence totale de normalisation typographique."),
        ("Ambiguïtés Foncières Spécifiques",
         "Des toponymes ardéchois abrégés (ex. « La Chap./A. » pour La Chapelle-sous-Aubenas), des renvois marginaux et des décalages de dates."),
        ("Volume Massif & Saisie Manuelle",
         "23 600 dossiers conservés chez GEO-SIAPP. Un traitement manuel requiert 15 à 30 minutes par dossier, soit plusieurs années de saisie.")
    ],
    img_path="img/registre_crop_opt.jpg",
    img_caption="Extrait d'un registre foncier manuscrit du cabinet GEO-SIAPP (années 1970)",
    slide_num=3,
    callout_title="OBJECTIF DU PROJET DE FIN D'ÉTUDES",
    callout_text="Développer une chaîne d'extraction IA 100% locale, souveraine et modulaire, divisant par 6 le temps de traitement tout en maintenant le géomètre au cœur de la décision."
)

# ==================== CHAPITRE 1 ====================
print("Génération Chapitre 1 (Slides 4-7)...")
add_transition_slide(
    chap_num=1,
    chap_title="Introduction & Contexte",
    chap_idx=0,
    slide_num=4,
    subtitle="Le cabinet GEO-SIAPP, le cadre réglementaire du bornage et les enjeux de modernisation"
)

add_two_column_slide(
    chap_idx=0,
    category="Chapitre 1 : Structure d'accueil",
    title="Le Cabinet GEO-SIAPP : implantation et patrimoine",
    items=[
        ("Une présence forte en Ardèche et Drôme",
         "Fondé en 1992 à Aubenas, le cabinet compte 4 agences : Aubenas (siège), Pierrelatte, Vallon-Pont-d'Arc et Guilherand-Granges."),
        ("Pluridisciplinarité des activités",
         "Expertises en foncier et bornage, urbanisme, ingénierie VRD, auscultation d'ouvrages et relevés 3D par scanner dynamique et drone."),
        ("Un fonds historique inestimable",
         "Conservation intégrale de 23 600 dossiers physiques établis entre 1959 et 2007 par les géomètres fondateurs du cabinet.")
    ],
    img_path="img/bureau_cropped.jpg",
    img_caption="Siège principal du Cabinet GEO-SIAPP à Aubenas (Ardèche)",
    slide_num=5,
    kpis=[("4", "AGENCES"), ("23 600", "DOSSIERS"), ("1959-2007", "ARCHIVES"), ("6 MOIS", "DURÉE PFE")]
)

add_two_column_slide(
    chap_idx=0,
    category="Chapitre 1 : Cadre légal",
    title="Le cadre réglementaire et juridique de la profession",
    items=[
        ("Monopole légal du géomètre-expert (Loi du 7 mai 1946)",
         "Seul professionnel habilité à fixer les limites réelles de la propriété foncière. L'acte de bornage possède une valeur juridique opposable aux tiers."),
        ("Obligation stricte de conservation (Décret n° 96-478)",
         "L'article 55 impose la conservation des archives foncières pendant au moins 30 ans, y compris après la cessation d'activité du géomètre."),
        ("Le portail national Géofoncier (OGE)",
         "Plateforme centrale de l'Ordre des Géomètres-Experts référençant toutes les interventions foncières sous forme de pastilles géolocalisées.")
    ],
    img_path="img/Chap2_pastilles.jpg",
    img_caption="Pastilles d'intervention foncière répertoriées sur le portail Géofoncier",
    slide_num=6,
    callout_title="DISTINCTION JURIDIQUE MAJEURE",
    callout_text="Le programme national GEODÉMAT (DGFiP) ne traite que le cadastre fiscal public : la numérisation des archives foncières privées relève exclusivement de la responsabilité des cabinets."
)

add_two_column_slide(
    chap_idx=0,
    category="Chapitre 1 : Enjeux métier",
    title="Pourquoi automatiser la recherche et le versement ?",
    items=[
        ("L'adage fondamental « Bornage sur bornage ne vaut »",
         "Tout nouveau bornage impose de retrouver l'acte antérieur pour garantir la sécurité juridique et la continuité des limites de propriété."),
        ("Une recherche d'antériorité manuelle et chronophage",
         "Retrouver un dossier physique prend 15 à 30 minutes. Sans géoréférencement, certains dossiers restent introuvables ou nécessitent des recherches payantes."),
        ("L'obstacle des abréviations contextuelles locales",
         "Les registres manuscrits utilisent des raccourcis propres au territoire (ex: « La Chap./A. » pour La Chapelle-sous-Aubenas), illisibles hors contexte.")
    ],
    img_path="img/chapelle_sous_aubenas_crop.png",
    img_caption="Registre manuscrit SIAPP avec mention abrégée locale « La Chap./A. »",
    slide_num=7,
    callout_title="VALEUR AJOUTÉE OPÉRATIONNELLE",
    callout_text="Verser les archives sur Géofoncier sécurise le patrimoine du cabinet et permet aux géomètres de localiser instantanément tout acte antérieur sur le terrain."
)

# ==================== CHAPITRE 2 ====================
print("Génération Chapitre 2 (Slides 8-10)...")
add_transition_slide(
    chap_num=2,
    chap_title="Analyse Métier & Données",
    chap_idx=1,
    slide_num=8,
    subtitle="Caractéristiques physiques des archives SIAPP et spécifications de l'API Géofoncier"
)

add_two_column_slide(
    chap_idx=1,
    category="Chapitre 2 : Fonds d'archives",
    title="Caractéristiques physiques et contraintes des données",
    items=[
        ("Une grande variété de supports documentaires",
         "50 ans de production documentaire comprenant des plans minutes sur calque, des registres fonciers manuscrits et des pièces écrites dactylographiées."),
        ("Des dégradations physiques réelles",
         "Papier jauni, encres pâlies, annotations marginales, plis d'usure et scans souvent inclinés ou de résolution inégale (150 à 400 DPI)."),
        ("Échec des moteurs OCR standards en boîte noire",
         "Sur les écritures manuscrites des registres, les OCR classiques type Tesseract dépassent 30% de taux d'erreur de caractères (CER).")
    ],
    img_path="img/4513_DA_124_crop.png",
    img_caption="Exemple de registre manuscrit du fonds SIAPP (commune de Prades)",
    slide_num=9,
    callout_title="EXIGENCE DE SOUVERAINETÉ STRICTE",
    callout_text="Les archives foncières contenant des données nominatives et patrimoniales sensibles, aucun document ne peut transiter par une API ou un serveur cloud externe."
)

add_two_column_slide(
    chap_idx=1,
    category="Chapitre 2 : API Géofoncier",
    title="Spécifications techniques de l'API REST Géofoncier",
    items=[
        ("5 métadonnées obligatoires par dossier",
         "1. Code INSEE de la commune (5 chiffres) • 2. Date de l'opération (format ISO 8601) • 3. Contenance en m² • 4. Nature de l'opération • 5. Type d'acte."),
        ("Authentification sécurisée par jeton JWT",
         "Requête initiale avec identifiants cabinet pour obtenir un jeton Bearer éphémère (validité 1 heure), assurant la traçabilité des dépôts."),
        ("Téléversement multipart et contrôle synchrone",
         "Envoi simultané des métadonnées JSON et du scan PDF du dossier. Réponse HTTP 201 Created confirmant la publication de la pastille.")
    ],
    img_path="img/Etape 3_Versement_geofoncier.jpg",
    img_caption="Interface de saisie et métadonnées cibles sur la plateforme Géofoncier",
    slide_num=10,
    callout_title="PRÉVENTION DES DOUBLONS",
    callout_text="Le protocole impose une interrogation préalable via l'API pour s'assurer qu'un dossier portant la même référence n'a pas déjà été versé."
)

# ==================== CHAPITRE 3 ====================
print("Génération Chapitre 3 (Slides 11-15)...")
add_transition_slide(
    chap_num=3,
    chap_title="État de l'Art",
    chap_idx=2,
    slide_num=11,
    subtitle="Technologies de reconnaissance de texte, extraction d'entités et vision multimodale"
)

add_two_column_slide(
    chap_idx=2,
    category="Chapitre 3 : OCR & HTR",
    title="Reconnaissance de texte : OCR imprimé vs HTR manuscrit",
    items=[
        ("EasyOCR : rapide et performant sur l'imprimé",
         "Architecture CRAFT (détection) + CRNN (reconnaissance). Idéal pour les cartouches récents et mentions dactylographiées (< 0.5s par page)."),
        ("TrOCR (Microsoft, Li et al., 2021) : le standard HTR",
         "Transformer Vision-Langue de bout en bout couplant un encodeur d'images (DeiT) et un décodeur de texte (RoBERTa). Pré-entraîné sur IAM et IIIT5K."),
        ("Écosystème HTR-United et écriture française",
         "Jeux de données ouverts spécialisés dans les écritures cursives anciennes, offrant une reconnaissance fine ligne par ligne.")
    ],
    img_path="img/trocr_architecture_figure1.png",
    img_caption="Architecture Transformer encodeur-décodeur de TrOCR (Li et al., 2021)",
    slide_num=12,
    callout_title="STRATÉGIE RETENUE",
    callout_text="Couplage hybride : EasyOCR sur les zones imprimées (cartouches, en-têtes) et TrOCR pour le déchiffrement des mentions manuscrites complexes."
)

add_two_column_slide(
    chap_idx=2,
    category="Chapitre 3 : NER & Extraction",
    title="Extraction d'entités : GLiNER vs LayoutLM",
    items=[
        ("GLiNER (Zaratiana et al., 2023 - urchade/GLiNER)",
         "Modèle bi-encodeur BERT compact (340M paramètres, licence Apache 2.0). Extraction zero-shot d'entités sans phase de fine-tuning préalable."),
        ("LayoutLM (Microsoft Research, 2020)",
         "Transformer multimodal intégrant texte, coordonnées 2D et image. Conçu pour les formulaires tabulaires rigides (factures, bordereaux)."),
        ("Pourquoi GLiNER a été choisi pour ce projet",
         "Les registres fonciers manuscrits ne respectent aucun gabarit fixe. GLiNER extrait des entités sur texte libre avec une grande tolérance aux variations.")
    ],
    img_path="img/gliner_prompt_render.png",
    img_caption="Fonctionnement zero-shot de GLiNER associant représentations textuelles et entités cibles",
    slide_num=13,
    callout_title="LABELS D'EXTRACTION AU RUNTIME",
    callout_text="Les 5 entités cibles (commune, date, contenance, nature, référence) sont interrogées dynamiquement avec un score de confiance associé [0, 1]."
)

add_two_column_slide(
    chap_idx=2,
    category="Chapitre 3 : Arbitrage VLM",
    title="Vision-Language Models : arbitrage local par Ollama",
    items=[
        ("Inférence souveraine sur GPU local avec Ollama",
         "Plateforme open-source exécutant des modèles multimodaux (LLaVA 1.6, Qwen2-VL) sur une carte graphique de bureau (8 Go VRAM) sans recours au cloud."),
        ("Rôle ciblé : arbitre de dernier recours",
         "Le VLM n'est pas utilisé sur chaque page en raison de sa latence (~8s), mais uniquement activé lorsque le score de confiance GLiNER est inférieur à 0.65."),
        ("Résolution d'ambiguïtés toponymiques réelles",
         "Sur l'abréviation « La Chap./A. », le modèle analyse la zone visuelle du scan, consulte la liste des communes d'Ardèche et restitue « La Chapelle-sous-Aubenas ».")
    ],
    img_path="img/vlm_arbitrage_chapelle_exact.png",
    img_caption="Démonstration d'arbitrage VLM : résolution locale de l'abréviation « La Chap./A. »",
    slide_num=14,
    callout_title="UN ARBITRE SOUVERAIN",
    callout_text="L'arbitrage multimodal local permet de résoudre les cas ambigus tout en respectant l'interdiction de communiquer des données vers l'extérieur."
)

# Slide 15 : Tableau comparatif
print("Génération Slide 15 (Tableau comparatif)...")
s15 = prs.slides.add_slide(blank_layout)
set_bg(s15, C_LIGHT_BG)
add_frieze(s15, 2)
add_slide_header(s15, "Tableau comparatif et justification des technologies retenues", "Chapitre 3 : Synthèse de l'état de l'art")
add_footer(s15, 15)

headers15 = ["Technologie", "Type", "Force principale", "Limite identifiée", "Rôle dans la chaîne"]
rows15 = [
    ["EasyOCR", "OCR classique", "Très rapide (< 0.5s), multilingue", "Échoue sur cursive (CER > 30%)", "Retenu : zones imprimées"],
    ["TrOCR", "HTR Transformer", "Robuste sur écriture manuscrite", "Nécessite une segmentation fine", "Retenu : registres manuscrits"],
    ["GLiNER", "NER Zero-shot", "Ultra-flexible, classes dynamiques", "Sensible aux textes trop longs", "Retenu : extraction d'entités"],
    ["LayoutLM", "Document AI 2D", "Excellent sur formulaires stricts", "Inadapté au manuscrit non tabulaire", "Écarté"],
    ["VLM (Ollama)", "Multimodal local", "Compréhension visuelle contextuelle", "Temps de calcul élevé (~8s / vue)", "Retenu : arbitre de secours"]
]

cw15 = [2.20, 1.60, 3.20, 3.20, 2.33]
rh15 = 0.56
tx15 = 0.40
ty15 = 1.45

hx = tx15
for hi, htext in enumerate(headers15):
    create_card(s15, hx, ty15, cw15[hi], rh15, bg_color=C_NAVY, border_color=None)
    bx = s15.shapes.add_textbox(Inches(hx + 0.05), Inches(ty15 + 0.12), Inches(cw15[hi] - 0.10), Inches(rh15 - 0.20))
    tf = bx.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = htext
    p.alignment = PP_ALIGN.CENTER
    r = p.runs[0]
    r.font.name = "Arial"
    r.font.size = Pt(10.0)
    r.font.bold = True
    r.font.color.rgb = C_WHITE
    hx += cw15[hi]

for ri, row in enumerate(rows15):
    ry = ty15 + (ri + 1) * rh15
    bg_r = C_CARD_BG if ri % 2 == 0 else RGBColor(241, 245, 249)
    rx = tx15
    for ci, cell in enumerate(row):
        create_card(s15, rx, ry, cw15[ci], rh15, bg_color=bg_r, border_color=C_BORDER)
        bx = s15.shapes.add_textbox(Inches(rx + 0.08), Inches(ry + 0.10), Inches(cw15[ci] - 0.16), Inches(rh15 - 0.20))
        tf = bx.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = cell
        p.alignment = PP_ALIGN.CENTER if ci in (0, 1, 4) else PP_ALIGN.LEFT
        r = p.runs[0]
        r.font.name = "Arial"
        r.font.size = Pt(9.5)
        if "Retenu" in cell:
            r.font.bold = True
            r.font.color.rgb = C_SUCCESS
        elif "Écarté" in cell:
            r.font.bold = True
            r.font.color.rgb = C_RED
        elif ci == 0:
            r.font.bold = True
            r.font.color.rgb = C_NAVY
        else:
            r.font.color.rgb = C_DARK
        rx += cw15[ci]

# Bandeau de synthèse bas
sy15 = ty15 + (len(rows15) + 1) * rh15 + 0.15
create_card(s15, tx15, sy15, 12.53, 0.70, bg_color=C_BOX_BG, border_color=C_ACCENT, border_width=Pt(1.0))
bx_syn = s15.shapes.add_textbox(Inches(tx15 + 0.20), Inches(sy15 + 0.12), Inches(12.13), Inches(0.48))
tf_syn = bx_syn.text_frame
p_syn = tf_syn.paragraphs[0]
p_syn.text = "SYNTHÈSE DE LA CHAÎNE RETENUE : Prétraitement d'image + OCR / HTR hybride + GLiNER zero-shot + VLM en arbitre de secours."
p_syn.alignment = PP_ALIGN.CENTER
r_syn = p_syn.runs[0]
r_syn.font.name = "Arial"
r_syn.font.size = Pt(10.5)
r_syn.font.bold = True
r_syn.font.color.rgb = C_NAVY

# ==================== CHAPITRE 4 ====================
print("Génération Chapitre 4 (Slides 16-19)...")
add_transition_slide(
    chap_num=4,
    chap_title="Architecture & Pipeline",
    chap_idx=3,
    slide_num=16,
    subtitle="Conception modulaire en 6 scripts, prétraitement géométrique et segmentation YOLOv8"
)

# Slide 17 : Les 6 scripts modulaires
s17 = prs.slides.add_slide(blank_layout)
set_bg(s17, C_LIGHT_BG)
add_frieze(s17, 3)
add_slide_header(s17, "Architecture logicielle : 6 scripts Python modulaires", "Chapitre 4 : Conception du système")
add_footer(s17, 17)

scripts_data = [
    ("main.py", "1. Orchestrateur", [
        ("Rôle", "Orchestration & flux"),
        ("Entrée", "Scans bruts (dossiers)"),
        ("Gestion", "Mémoire & logs CSV"),
        ("Sortie", "Dossiers prêts au versement")
    ]),
    ("plan_classifieur.py", "2. Classification", [
        ("Rôle", "Tri typologique"),
        ("4 classes", "DAO, DMPC, registre, pièce"),
        ("Vitesse", "< 0.8 s par document"),
        ("Précision", "98.5% sur corpus test")
    ]),
    ("spatial_extractor.py", "3. Détection YOLO", [
        ("Rôle", "Segmentation visuelle"),
        ("Modèle", "YOLOv8 nano (3.2M params)"),
        ("Zones", "Cartouche, tableau, notes"),
        ("Entraînement", "400 documents annotés")
    ]),
    ("outil_ocr.py", "4. OCR, HTR & NER", [
        ("Rôle", "Transcription & extraction"),
        ("Hybride", "EasyOCR + TrOCR"),
        ("Extraction", "GLiNER zero-shot"),
        ("Arbitre", "VLM Ollama (score < 0.65)")
    ]),
    ("coherence_controle.py", "5. Contrôle Qualité", [
        ("Rôle", "Validation métier"),
        ("Règles", "17 contrôles automatiques"),
        ("Toponymie", "Levenshtein 335 communes"),
        ("Conformité", "ISO 8601 & plage surfaces")
    ]),
    ("geofoncier_api.py", "6. IHM & Versement", [
        ("Rôle", "Supervision & injection"),
        ("Interface", "Streamlit ergonomique"),
        ("Sécurité", "Jeton JWT Bearer 1h"),
        ("Endpoint", "POST /dossiers REST OGE")
    ])
]

cw17 = (12.53 - 5 * 0.12) / 6.0
for si, (sname, stitle, sposts) in enumerate(scripts_data):
    sx = 0.40 + si * (cw17 + 0.12)
    sy = 1.40
    sh = 4.70
    create_card(s17, sx, sy, cw17, sh, bg_color=C_CARD_BG, border_color=C_BORDER)
    
    # En-tête carte
    create_card(s17, sx, sy, cw17, 0.65, bg_color=C_NAVY, border_color=None)
    bx_st = s17.shapes.add_textbox(Inches(sx + 0.04), Inches(sy + 0.12), Inches(cw17 - 0.08), Inches(0.42))
    tf_st = bx_st.text_frame
    tf_st.word_wrap = True
    p_st = tf_st.paragraphs[0]
    p_st.alignment = PP_ALIGN.CENTER
    p_st.text = stitle
    r_st = p_st.runs[0]
    r_st.font.name = "Arial"
    r_st.font.size = Pt(9.5)
    r_st.font.bold = True
    r_st.font.color.rgb = C_WHITE
    
    # Nom script
    bx_sn = s17.shapes.add_textbox(Inches(sx + 0.06), Inches(sy + 0.78), Inches(cw17 - 0.12), Inches(0.35))
    tf_sn = bx_sn.text_frame
    p_sn = tf_sn.paragraphs[0]
    p_sn.alignment = PP_ALIGN.CENTER
    p_sn.text = sname
    r_sn = p_sn.runs[0]
    r_sn.font.name = "Arial"
    r_sn.font.size = Pt(10.5)
    r_sn.font.bold = True
    r_sn.font.color.rgb = C_ACCENT
    
    # Détails sous forme de mini-blocs
    bx_sd = s17.shapes.add_textbox(Inches(sx + 0.08), Inches(sy + 1.20), Inches(cw17 - 0.16), Inches(3.40))
    tf_sd = bx_sd.text_frame
    tf_sd.word_wrap = True
    tf_sd.margin_left = tf_sd.margin_top = tf_sd.margin_right = tf_sd.margin_bottom = 0
    
    for pi, (label_p, val_p) in enumerate(sposts):
        p_item = tf_sd.paragraphs[0] if pi == 0 else tf_sd.add_paragraph()
        if pi > 0:
            p_item.space_before = Pt(6)
        r_lbl = p_item.add_run()
        r_lbl.text = label_p + " :\n"
        r_lbl.font.name = "Arial"
        r_lbl.font.size = Pt(9.0)
        r_lbl.font.bold = True
        r_lbl.font.color.rgb = C_NAVY
        
        p_val = tf_sd.add_paragraph()
        p_val.space_before = Pt(1)
        r_val = p_val.add_run()
        r_val.text = val_p
        r_val.font.name = "Arial"
        r_val.font.size = Pt(8.5)
        r_val.font.color.rgb = C_MUTED

# Bandeau bas workflow
create_card(s17, 0.40, 6.25, 12.53, 0.65, bg_color=RGBColor(241, 245, 249), border_color=C_BORDER)
bx_wf = s17.shapes.add_textbox(Inches(0.50), Inches(6.35), Inches(12.33), Inches(0.45))
tf_wf = bx_wf.text_frame
p_wf = tf_wf.paragraphs[0]
p_wf.alignment = PP_ALIGN.CENTER
p_wf.text = "FLUX OPÉRATIONNEL : Scan brut  →  Classification  →  Détection YOLOv8  →  OCR / HTR  →  Contrôle qualité  →  Versement API"
r_wf = p_wf.runs[0]
r_wf.font.name = "Arial"
r_wf.font.size = Pt(9.5)
r_wf.font.bold = True
r_wf.font.color.rgb = C_NAVY

add_two_column_slide(
    chap_idx=3,
    category="Chapitre 4 : Prétraitement d'image",
    title="Prétraitement : redressement Hough et binarisation Otsu",
    items=[
        ("Contrôle automatique de résolution",
         "Vérification du seuil minimal de 300 DPI. En-dessous de 200 DPI, alerte automatique de l'opérateur pour éviter les faux positifs."),
        ("Redressement d'angle par transformée de Hough",
         "Détection des lignes de grille du document pour estimer et corriger automatiquement l'angle d'inclinaison (correction jusqu'à ±15°)."),
        ("Binarisation adaptative d'Otsu",
         "Séparation dynamique encre/papier, éliminant les ombres de numérisation et le jaunissement tout en renforçant les tracés fins.")
    ],
    img_path="img/4513_DA_124_centered_crop.png",
    img_caption="Document d'archive après redressement par transformée de Hough et binarisation d'Otsu",
    slide_num=18,
    callout_title="IMPACT MESURÉ SUR LE CORPUS TEST",
    callout_text="Le prétraitement réduit le taux d'erreur de caractères (CER) de 18 points par rapport à un OCR brut sur scan non corrigé."
)

add_two_column_slide(
    chap_idx=3,
    category="Chapitre 4 : Détection & Extraction",
    title="Détection spatiale YOLOv8 et extraction d'entités",
    items=[
        ("Segmentation par YOLOv8 nano",
         "Modèle léger (3.2M paramètres) entraîné sur 400 documents annotés du cabinet. Détection en moins de 25 ms sur GPU local."),
        ("3 classes de structures géométriques",
         "Localisation précise du cartouche officiel, du tableau de contenances parcellaires et des mentions marginales manuscrites (mAP@0.5 = 0.87)."),
        ("Extraction ciblée par GLiNER",
         "Le texte issu uniquement des régions détectées par YOLO est transmis à GLiNER, évitant le parasitage par le reste du plan.")
    ],
    img_path="img/composite_yolo_extraction.png",
    img_caption="Détection par boîtes englobantes YOLOv8 et extraction des métadonnées cibles",
    slide_num=19,
    callout_title="SYNERGIE VISION ET EXTRACTION",
    callout_text="Le découpage spatial préalable améliore le score F1 global de l'extraction de 0.72 à 0.88 en supprimant les bruits périphériques."
)

# ==================== CHAPITRE 5 ====================
print("Génération Chapitre 5 (Slides 20-22)...")
add_transition_slide(
    chap_num=5,
    chap_title="Fiabilisation & Interface",
    chap_idx=4,
    slide_num=20,
    subtitle="Interface opérateur Streamlit et moteur de 17 règles de cohérence métier"
)

add_two_column_slide(
    chap_idx=4,
    category="Chapitre 5 : Interface Streamlit",
    title="Interface de relecture assistée et supervision humaine",
    items=[
        ("Ergonomie adaptée au métier de géomètre",
         "Affichage synchronisé côte-à-côte du document d'archive numérisé et du formulaire pré-rempli par les algorithmes d'IA."),
        ("Code couleur de confiance à trois niveaux",
         "Vert (> 0.85) : Information fiable, validation en 1 clic • Orange (0.65 - 0.85) : À vérifier • Rouge (< 0.65) : Saisie manuelle requise."),
        ("Cartographie interactive Folium intégrée",
         "Géolocalisation immédiate du dossier sur fond OpenStreetMap et cadastre pour vérifier la cohérence spatiale de la commune.")
    ],
    img_path="img/interface_haut.jpg",
    img_caption="Interface Streamlit : double vue scan / formulaire pré-rempli avec indicateurs de confiance",
    slide_num=21,
    callout_title="RÉDUCTION DU TEMPS OPÉRATEUR",
    callout_text="La validation humaine assistée s'effectue en 1 à 2 minutes par dossier, contre 15 à 30 minutes lors d'une saisie manuelle intégrale."
)

add_two_column_slide(
    chap_idx=4,
    category="Chapitre 5 : Règles Métier",
    title="Moteur de 17 règles de cohérence métier",
    items=[
        ("Contrôle toponymique par distance de Levenshtein",
         "Rapprochement automatique du nom extrait avec la table officielle des 335 communes d'Ardèche (tolérance calibrée à 2 fautes)."),
        ("Validation temporelle stricte (ISO 8601)",
         "Contrôle de conformité de la date d'acte (années bissextiles, 30/31 jours) et cohérence chronologique (bornes 1950 - 2010)."),
        ("Plausibilité des contenances parcellaires",
         "Vérification de la plage de surface (10 m² à 50 ha) et conversion automatique des unités historiques (ares, centiares)."),
        ("Vérification anti-doublon en temps réel",
         "Interrogation de l'API Géofoncier avant validation pour interdire tout versement redondant d'un dossier déjà référencé.")
    ],
    img_path="img/streamlit_map_app.jpg",
    img_caption="Composant cartographique de vérification spatiale et validation géographique",
    slide_num=22,
    callout_title="SÉCURISATION DU PROCESSUS",
    callout_text="Aucun dossier n'est transmis à l'API Géofoncier s'il ne valide pas l'intégralité des 17 règles de cohérence programmées."
)

# ==================== CHAPITRE 6 ====================
print("Génération Chapitre 6 (Slides 23-26)...")
add_transition_slide(
    chap_num=6,
    chap_title="Intégration Géofoncier",
    chap_idx=5,
    slide_num=23,
    subtitle="Démonstration vidéo, protocole d'injection REST et résultats expérimentaux sur Prades"
)

# Slide 24 : Vidéo Démonstration
print("Génération Slide 24 (Vidéo)...")
s24 = prs.slides.add_slide(blank_layout)
set_bg(s24, RGBColor(10, 20, 35))
add_footer(s24, 24)

bx_vt = s24.shapes.add_textbox(Inches(0.80), Inches(0.20), Inches(11.73), Inches(0.40))
tf_vt = bx_vt.text_frame
p_vt = tf_vt.paragraphs[0]
p_vt.alignment = PP_ALIGN.CENTER
p_vt.text = "DÉMONSTRATION DU PIPELINE : TRAITEMENT ET VERSEMENT EN CONDITIONS RÉELLES"
r_vt = p_vt.runs[0]
r_vt.font.name = "Arial"
r_vt.font.size = Pt(13.0)
r_vt.font.bold = True
r_vt.font.color.rgb = C_ACCENT

vid_path = "Demo_soutenance_final.mp4"
vw, vh = 11.80, 6.25
vx = (SLIDE_W - vw) / 2.0
vy = 0.70

if os.path.exists(vid_path):
    try:
        poster_img = "video_poster.jpg" if os.path.exists("video_poster.jpg") else None
        s24.shapes.add_movie(vid_path, Inches(vx), Inches(vy), Inches(vw), Inches(vh), poster_frame_image=poster_img, mime_type="video/mp4")
        print(f"  Vidéo intégrée avec succès : {vid_path} (poster: {poster_img})")
    except Exception as e:
        print(f"  Erreur add_movie : {e}")
        create_card(s24, vx, vy, vw, vh, bg_color=RGBColor(15, 30, 50), border_color=C_ACCENT)
        bx_err = s24.shapes.add_textbox(Inches(vx), Inches(vy + vh/2 - 0.3), Inches(vw), Inches(0.6))
        bx_err.text_frame.paragraphs[0].text = f"Vidéo prête : {vid_path}"
        bx_err.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
else:
    create_card(s24, vx, vy, vw, vh, bg_color=RGBColor(15, 30, 50), border_color=C_ACCENT)
    bx_err = s24.shapes.add_textbox(Inches(vx), Inches(vy + vh/2 - 0.3), Inches(vw), Inches(0.6))
    bx_err.text_frame.paragraphs[0].text = f"Fichier vidéo introuvable : {vid_path}"
    bx_err.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

add_two_column_slide(
    chap_idx=5,
    category="Chapitre 6 : API Géofoncier",
    title="Protocole de versement via l'API REST Géofoncier",
    items=[
        ("Génération de jeton sécurisé OAuth2 / JWT",
         "Connexion initiale avec identifiants cabinet générant un jeton Bearer garantissant l'origine certifiée des dépôts fonciers."),
        ("Téléversement multipart synchrone",
         "Envoi simultané des métadonnées vérifiées et du fichier PDF scanné vers le point de terminaison officiel POST /dossiers."),
        ("Traitement par lot et journal d'audit",
         "Capacité d'injecter une série de dossiers validés avec génération automatique d'un journal CSV horodaté pour archivage.")
    ],
    img_path="img/6_succes_pastille.jpg",
    img_caption="Pastille d'intervention créée sur Géofoncier après versement réussi par le pipeline",
    slide_num=25,
    callout_title="CONFIRMATION DU DÉPÔT",
    callout_text="Le retour HTTP 201 Created valide l'apparition de la pastille bleue sur la carte de référence de l'Ordre des Géomètres-Experts."
)

# Slide 26 : Résultats Prades
print("Génération Slide 26 (Résultats Prades)...")
s26 = prs.slides.add_slide(blank_layout)
set_bg(s26, C_LIGHT_BG)
add_frieze(s26, 5)
add_slide_header(s26, "Validation expérimentale sur la commune de Prades (Ardèche)", "Chapitre 6 : Résultats obtenus")
add_footer(s26, 26)

# Gauche : Détails du test
create_card(s26, 0.40, 1.34, 5.95, 4.30, bg_color=C_CARD_BG, border_color=C_BORDER)
bx_p = s26.shapes.add_textbox(Inches(0.60), Inches(1.48), Inches(5.55), Inches(4.00))
tf_p = bx_p.text_frame
tf_p.word_wrap = True
tf_p.margin_left = tf_p.margin_top = tf_p.margin_right = tf_p.margin_bottom = 0

items_prades = [
    ("Campagne sur 50 dossiers réels (1970-1990)",
     "Évaluation en conditions opérationnelles sur des registres physiques de la commune de Prades conservés chez GEO-SIAPP."),
    ("Taux de succès de 94% (47 dossiers versés)",
     "47 dossiers sur 50 insérés avec succès. 3 échecs identifiés : 2 scans illisibles (< 200 DPI) et 1 doublon d'antériorité déjà présent."),
    ("Validation par le tuteur entreprise (M. Hague)",
     "Contrôle systématique de l'exactitude des parcelles, dates et noms de propriétaires avant accord définitif de versement."),
    ("Gain de productivité constaté : facteur 6",
     "Temps de traitement global pour les 50 dossiers : ~4 heures avec l'outil contre 25 heures en saisie manuelle traditionnelle.")
]

for ii, (h_p, b_p) in enumerate(items_prades):
    p_h = tf_p.paragraphs[0] if ii == 0 else tf_p.add_paragraph()
    if ii > 0:
        p_h.space_before = Pt(8)
    r1 = p_h.add_run()
    r1.text = h_p + "\n"
    r1.font.name = "Arial"
    r1.font.size = Pt(11.0)
    r1.font.bold = True
    r1.font.color.rgb = C_NAVY
    
    p_b = tf_p.add_paragraph()
    p_b.space_before = Pt(2)
    r2 = p_b.add_run()
    r2.text = b_p
    r2.font.name = "Arial"
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = C_MUTED

# KPIs bas gauche
kpis_prades = [("50", "DOSSIERS"), ("94%", "SUCCÈS"), ("F1 = 0.88", "SCORE GLOBAL"), ("x6", "RAPIDITÉ")]
kw_p = (5.95 - 3 * 0.10) / 4.0
ky_p = 5.80
for ki, (kv, kl) in enumerate(kpis_prades):
    kx = 0.40 + ki * (kw_p + 0.10)
    create_card(s26, kx, ky_p, kw_p, 1.15, bg_color=C_NAVY, border_color=None)
    bx_k = s26.shapes.add_textbox(Inches(kx + 0.04), Inches(ky_p + 0.15), Inches(kw_p - 0.08), Inches(0.85))
    tf_k = bx_k.text_frame
    tf_k.margin_left = tf_k.margin_top = tf_k.margin_right = tf_k.margin_bottom = 0
    p1 = tf_k.paragraphs[0]
    p1.text = kv
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.runs[0]
    r1.font.name = "Arial"
    r1.font.size = Pt(15.0)
    r1.font.bold = True
    r1.font.color.rgb = C_ACCENT
    p2 = tf_k.add_paragraph()
    p2.space_before = Pt(3)
    p2.text = kl
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.runs[0]
    r2.font.name = "Arial"
    r2.font.size = Pt(7.5)
    r2.font.color.rgb = C_WHITE

# Droite : Barres F1 + Image confirmation
right_x26 = 6.60
right_w26 = 6.33
create_card(s26, right_x26, 1.34, right_w26, 3.20, bg_color=C_CARD_BG, border_color=C_BORDER)

bx_f1t = s26.shapes.add_textbox(Inches(right_x26 + 0.20), Inches(1.48), Inches(right_w26 - 0.40), Inches(0.30))
tf_f1t = bx_f1t.text_frame
p_f1t = tf_f1t.paragraphs[0]
p_f1t.text = "SCORES F1 PAR CHAMP OBLIGATOIRE"
r_f1t = p_f1t.runs[0]
r_f1t.font.name = "Arial"
r_f1t.font.size = Pt(9.5)
r_f1t.font.bold = True
r_f1t.font.color.rgb = C_ACCENT

f1_data = [
    ("Commune (code INSEE)", 0.94, C_SUCCESS),
    ("Date de l'opération", 0.91, C_SUCCESS),
    ("Contenance (m²)", 0.88, C_ACCENT),
    ("Nature de l'opération", 0.85, C_ACCENT),
    ("Type d'acte", 0.82, C_WARN),
    ("Score F1 global pondéré", 0.88, C_SUCCESS)
]

bar_y = 1.82
for f_name, f_val, f_col in f1_data:
    bx_fn = s26.shapes.add_textbox(Inches(right_x26 + 0.20), Inches(bar_y), Inches(2.60), Inches(0.24))
    tf_fn = bx_fn.text_frame
    p_fn = tf_fn.paragraphs[0]
    p_fn.text = f_name
    r_fn = p_fn.runs[0]
    r_fn.font.name = "Arial"
    r_fn.font.size = Pt(8.5)
    r_fn.font.color.rgb = C_DARK
    
    # Fond barre
    create_card(s26, right_x26 + 2.85, bar_y + 0.03, 2.30, 0.16, bg_color=C_BORDER, border_color=None)
    # Barre remplie
    create_card(s26, right_x26 + 2.85, bar_y + 0.03, 2.30 * f_val, 0.16, bg_color=f_col, border_color=None)
    
    bx_fv = s26.shapes.add_textbox(Inches(right_x26 + 5.25), Inches(bar_y), Inches(0.90), Inches(0.24))
    tf_fv = bx_fv.text_frame
    p_fv = tf_fv.paragraphs[0]
    p_fv.text = f"F1 = {f_val:.2f}"
    r_fv = p_fv.runs[0]
    r_fv.font.name = "Arial"
    r_fv.font.size = Pt(8.5)
    r_fv.font.bold = True
    r_fv.font.color.rgb = f_col
    
    bar_y += 0.40

# Image confirmation dessous
create_card(s26, right_x26, 4.68, right_w26, 2.27, bg_color=C_CARD_BG, border_color=C_BORDER)
add_fitted_picture(s26, "img/6.4_match_confirm.jpg", right_x26 + 0.10, 4.75, right_w26 - 0.20, 2.10, caption="Fiche de correspondance validée sur la plateforme Géofoncier")

# ==================== CHAPITRE 7 ====================
print("Génération Chapitre 7 (Slides 27-29)...")
add_transition_slide(
    chap_num=7,
    chap_title="Limites & Perspectives",
    chap_idx=6,
    slide_num=27,
    subtitle="Contraintes techniques identifiées, axes d'amélioration et perspectives de passage à l'échelle"
)

add_two_column_slide(
    chap_idx=6,
    category="Chapitre 7 : Limites actuelles",
    title="Limites techniques et contraintes opérationnelles",
    items=[
        ("Dépendance à la qualité des numérisations",
         "Les scans à moins de 200 DPI ou fortement floutés provoquent des échecs d'alignement. Une numérisation physique initiale de qualité reste essentielle."),
        ("Manuscrits anciens très altérés (pré-1960)",
         "Les registres les plus anciens présentent des encres artisanales fanées et des écritures très serrées difficiles à segmenter sans réentraînement."),
        ("Exigences matérielles pour l'inférence locale",
         "L'exécution combinée de YOLOv8, TrOCR et GLiNER nécessite une station de travail dotée de 8 Go de VRAM GPU pour garantir un temps d'inférence < 3s.")
    ],
    img_path="img/4513_DA_124_crop.png",
    img_caption="Document ancien illustrant les défis de déchiffrement des écritures manuscrites serrées",
    slide_num=28,
    callout_title="SUPERVISION HUMAINE INDISPENSABLE",
    callout_text="Le système n'est pas conçu pour fonctionner en boîte noire autonome : le géomètre-expert conserve l'entière responsabilité déontologique du versement."
)

add_two_column_slide(
    chap_idx=6,
    category="Chapitre 7 : Perspectives",
    title="Perspectives d'évolution et passage à l'échelle",
    items=[
        ("Fine-tuning ciblé de TrOCR sur les écritures SIAPP",
         "Création d'un jeu de données d'apprentissage sur mesure pour les graphies récurrentes des géomètres fondateurs afin d'augmenter le score F1 sur manuscrit."),
        ("Généralisation et déploiement multi-cabinets",
         "Architecture modulaire transposable à d'autres cabinets de géomètres-experts confrontés au même défi d'archivage foncier national."),
        ("Croisement avec le plan cadastral vectoriel (PCI)",
         "Intégration du flux cadastral vectorisé de la DGFiP pour géolocaliser automatiquement le centroïde des parcelles lors du versement.")
    ],
    img_path="img/repartition_archives.png",
    img_caption="Typologie des archives et potentiel d'automatisation par catégorie de document",
    slide_num=29,
    callout_title="AUTO-VALIDATION À TERME",
    callout_text="Pour les dossiers présentant un score F1 supérieur à 0.95 sur tous les champs, une validation en un clic pourrait être autorisée sous agrément de l'OGE."
)

# ==================== CHAPITRE 8 ====================
print("Génération Chapitre 8 (Slides 30-32)...")
add_transition_slide(
    chap_num=8,
    chap_title="Conclusion",
    chap_idx=7,
    slide_num=30,
    subtitle="Bilan général du PFE, retours d'expérience et remerciements"
)

# Slide 31 : Bilan
print("Génération Slide 31 (Bilan)...")
s31 = prs.slides.add_slide(blank_layout)
set_bg(s31, C_LIGHT_BG)
add_frieze(s31, 7)
add_slide_header(s31, "Bilan général du Projet de Fin d'Études", "Chapitre 8 : Conclusion")
add_footer(s31, 31)

bilan_cols = [
    ("OBJECTIFS TECHNIQUES", [
        ("Pipeline 100% local souverain", "Traitement intégral sur PC de bureau sans aucune dépendance cloud externe."),
        ("Couplage algorithmique innovant", "YOLOv8 pour la structure, TrOCR pour le texte, GLiNER pour les entités et VLM en arbitre."),
        ("17 règles métier fiabilisées", "Vérification toponymique INSEE par Levenshtein et contrôle strict ISO 8601.")
    ]),
    ("RÉSULTATS OPÉRATIONNELS", [
        ("94% de succès sur Prades", "47 dossiers réels sur 50 versés avec succès sur le portail officiel Géofoncier."),
        ("Productivité multipliée par 6", "Temps de traitement moyen ramené de 25-30 minutes à moins de 5 minutes."),
        ("Patrimoine documentaire valorisé", "Perspective concrète de traitement des 23 600 dossiers conservés chez GEO-SIAPP.")
    ]),
    ("APPORTS DE L'INGÉNIEUR", [
        ("Gestion de projet de bout en bout", "De l'analyse des archives physiques jusqu'au déploiement de l'outil en agence."),
        ("Rigueur méthodologique et juridique", "Respect absolu du monopole de la profession et du cadre déontologique de l'OGE."),
        ("Outil immédiatement exploitable", "Code modulaire, documenté et prêt à l'emploi pour les techniciens du cabinet.")
    ])
]

cw31 = (12.53 - 2 * 0.16) / 3.0
for ci, (col_title, col_items) in enumerate(bilan_cols):
    cx = 0.40 + ci * (cw31 + 0.16)
    create_card(s31, cx, 1.40, cw31, 4.30, bg_color=C_CARD_BG, border_color=C_BORDER)
    
    # En-tête colonne
    create_card(s31, cx, 1.40, cw31, 0.60, bg_color=C_NAVY, border_color=None)
    bx_ct = s31.shapes.add_textbox(Inches(cx + 0.05), Inches(1.50), Inches(cw31 - 0.10), Inches(0.40))
    tf_ct = bx_ct.text_frame
    p_ct = tf_ct.paragraphs[0]
    p_ct.alignment = PP_ALIGN.CENTER
    p_ct.text = col_title
    r_ct = p_ct.runs[0]
    r_ct.font.name = "Arial"
    r_ct.font.size = Pt(10.0)
    r_ct.font.bold = True
    r_ct.font.color.rgb = C_WHITE
    
    # Contenu colonne
    bx_cc = s31.shapes.add_textbox(Inches(cx + 0.16), Inches(2.15), Inches(cw31 - 0.32), Inches(3.40))
    tf_cc = bx_cc.text_frame
    tf_cc.word_wrap = True
    tf_cc.margin_left = tf_cc.margin_top = tf_cc.margin_right = tf_cc.margin_bottom = 0
    
    for cii, (head_c, body_c) in enumerate(col_items):
        p_ch = tf_cc.paragraphs[0] if cii == 0 else tf_cc.add_paragraph()
        if cii > 0:
            p_ch.space_before = Pt(8)
        r1 = p_ch.add_run()
        r1.text = head_c + "\n"
        r1.font.name = "Arial"
        r1.font.size = Pt(10.5)
        r1.font.bold = True
        r1.font.color.rgb = C_NAVY
        
        p_cb = tf_cc.add_paragraph()
        p_cb.space_before = Pt(2)
        r2 = p_cb.add_run()
        r2.text = body_c
        r2.font.name = "Arial"
        r2.font.size = Pt(9.0)
        r2.font.color.rgb = C_MUTED

# Bandeau KPI bas bilan
kpis_bilan = [("F1 = 0.88", "SCORE GLOBAL"), ("94%", "SUCCÈS PRADES"), ("x6", "GAIN TEMPS"), ("23 600", "DOSSIERS EN COURS")]
kw31_k = (12.53 - 3 * 0.12) / 4.0
ky31_k = 5.85
for ki, (kv, kl) in enumerate(kpis_bilan):
    kx = 0.40 + ki * (kw31_k + 0.12)
    create_card(s31, kx, ky31_k, kw31_k, 1.10, bg_color=C_NAVY, border_color=None)
    bx_k = s31.shapes.add_textbox(Inches(kx + 0.05), Inches(ky31_k + 0.15), Inches(kw31_k - 0.10), Inches(0.80))
    tf_k = bx_k.text_frame
    p1 = tf_k.paragraphs[0]
    p1.text = kv
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.runs[0]
    r1.font.name = "Arial"
    r1.font.size = Pt(15.5)
    r1.font.bold = True
    r1.font.color.rgb = C_ACCENT
    p2 = tf_k.add_paragraph()
    p2.space_before = Pt(2)
    p2.text = kl
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.runs[0]
    r2.font.name = "Arial"
    r2.font.size = Pt(7.5)
    r2.font.color.rgb = C_WHITE

# Slide 32 : Remerciements & Questions
print("Génération Slide 32 (Remerciements)...")
s32 = prs.slides.add_slide(blank_layout)
set_bg(s32, C_NAVY)

bx_rt = s32.shapes.add_textbox(Inches(1.0), Inches(1.50), Inches(11.33), Inches(0.40))
tf_rt = bx_rt.text_frame
p_rt = tf_rt.paragraphs[0]
p_rt.alignment = PP_ALIGN.CENTER
p_rt.text = "PROJET DE FIN D'ÉTUDES : INSA STRASBOURG"
r_rt = p_rt.runs[0]
r_rt.font.name = "Arial"
r_rt.font.size = Pt(12.0)
r_rt.font.bold = True
r_rt.font.color.rgb = C_ACCENT

bx_rm = s32.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(11.33), Inches(0.90))
tf_rm = bx_rm.text_frame
p_rm = tf_rm.paragraphs[0]
p_rm.alignment = PP_ALIGN.CENTER
p_rm.text = "Merci pour votre attention"
r_rm = p_rm.runs[0]
r_rm.font.name = "Arial"
r_rm.font.size = Pt(36.0)
r_rm.font.bold = True
r_rm.font.color.rgb = C_WHITE

bx_rq = s32.shapes.add_textbox(Inches(1.0), Inches(2.90), Inches(11.33), Inches(0.40))
tf_rq = bx_rq.text_frame
p_rq = tf_rq.paragraphs[0]
p_rq.alignment = PP_ALIGN.CENTER
p_rq.text = "Je me tiens à votre disposition pour vos questions."
r_rq = p_rq.runs[0]
r_rq.font.name = "Arial"
r_rq.font.size = Pt(15.0)
r_rq.font.color.rgb = RGBColor(186, 230, 253)

sep32 = s32.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.66), Inches(3.50), Inches(4.00), Inches(0.025))
sep32.fill.solid()
sep32.fill.fore_color.rgb = C_ACCENT
sep32.line.fill.background()

thanks = [
    "M. Gaëtan HAGUE : Tuteur entreprise, Géomètre-Expert associé, Cabinet GEO-SIAPP",
    "M. Mathieu KOEHL : Directeur du PFE, Enseignant-chercheur, INSA Strasbourg / ICube",
    "L'équipe du Cabinet GEO-SIAPP d'Aubenas pour leur accueil et leur disponibilité",
    "L'équipe pédagogique de la spécialité Topographie de l'INSA Strasbourg"
]

for ti, ttext in enumerate(thanks):
    bx_th = s32.shapes.add_textbox(Inches(1.0), Inches(3.70 + ti * 0.40), Inches(11.33), Inches(0.35))
    tf_th = bx_th.text_frame
    p_th = tf_th.paragraphs[0]
    p_th.alignment = PP_ALIGN.CENTER
    p_th.text = ttext
    r_th = p_th.runs[0]
    r_th.font.name = "Arial"
    r_th.font.size = Pt(11.0)
    r_th.font.color.rgb = C_WHITE

if os.path.exists("Logo_INSAStrasbourg.jpg"):
    s32.shapes.add_picture("Logo_INSAStrasbourg.jpg", Inches(3.80), Inches(5.60), width=Inches(2.40))
if os.path.exists("Geosiapp.jpg"):
    add_fitted_picture(s32, "Geosiapp.jpg", 7.60, 5.45, 2.00, 1.20)

add_footer(s32, 32)

# ==================== ENREGISTREMENT ====================
output_pptx = "Soutenance_PFE_Adrien_TRAVAILLE.pptx"
prs.save(output_pptx)
print(f"\n=======================================================")
print(f"SUCCÈS : Présentation enregistrée sous '{output_pptx}'")
print(f"Nombre total de diapositives : {len(prs.slides)}")
print(f"=======================================================\n")
