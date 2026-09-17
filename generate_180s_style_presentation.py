# -*- coding: utf-8 -*-
"""
generate_180s_style_presentation.py
Générateur de la présentation de soutenance de PFE d'Adrien TRAVAILLÉ.
Strictement fidèle à la direction artistique de 'Presentation_PFE_180s.pptx' :
1. Fond clair #F8FAFC et grand conteneur arrondi blanc #FFFFFF.
2. En-tête permanent à double niveau inspiré du 180s :
   - Logos INSA Strasbourg (gauche) et GEO-SIAPP (droite).
   - Niveau 1 : Pilules arrondies des 8 grands chapitres (chapitre actif en Bleu Nuit #102C57).
   - Niveau 2 : Ruban des sous-parties du chapitre actif (sous-partie active en Bleu Accent #0284C7).
3. Sommaire au même style épuré (pas de fond bleu uniforme).
4. Respect strict de la logique du mémoire (chapitres 1 à 8 et leurs sous-sections).
5. Vidéo Demo_soutenance_final.mp4 intégrée avec poster frame (intouchée).
6. Respect des directives de rédaction (aucun tiret cadratin).
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image

# ==================== CONSTANTES GRAPHIQUES ====================
SLIDE_W = 13.333
SLIDE_H = 7.50
TOTAL_SLIDES = 27

C_NAVY       = RGBColor(16, 44, 87)       # #102C57 - Bleu Nuit Institutionnel
C_ACCENT     = RGBColor(2, 132, 199)      # #0284C7 - Bleu Accent Soutenance
C_DARK       = RGBColor(15, 23, 42)       # #0F172A - Texte Principal
C_MUTED      = RGBColor(100, 116, 139)    # #64748B - Texte Secondaire
C_BG_SLIDE   = RGBColor(248, 250, 252)    # #F8FAFC - Fond général diapositive
C_CARD_BG    = RGBColor(255, 255, 255)    # #FFFFFF - Grand conteneur blanc
C_BORDER_TAB = RGBColor(203, 213, 225)    # #CBD5E1 - Bordure onglet inactif
C_BG_TAB     = RGBColor(241, 245, 249)    # #F1F5F9 - Fond onglet inactif
C_LINE       = RGBColor(226, 232, 240)    # #E2E8F0 - Ligne séparatrice
C_CALLOUT_BG = RGBColor(240, 249, 255)    # #F0F9FF - Fond encadré bleu clair
C_WHITE      = RGBColor(255, 255, 255)
C_SUCCESS    = RGBColor(22, 163, 74)      # #16A34A - Succès
C_WARN       = RGBColor(217, 119, 6)      # #D97706 - Attention
C_RED        = RGBColor(220, 38, 38)      # #DC2626 - Rejet

# Définition des 8 chapitres et de leurs sous-parties réelles
CHAPTERS_DATA = [
    {
        "short": "1. Intro",
        "full": "Introduction & Contexte",
        "subs": ["1.1 Structure GEO-SIAPP", "1.2 Cadre réglementaire", "1.3 Contexte & Enjeux", "1.4 Problématique"]
    },
    {
        "short": "2. Métier",
        "full": "Analyse Métier & Données",
        "subs": ["2.1 Cadre juridique", "2.2 Spécifications API", "2.3 Typologie des archives"]
    },
    {
        "short": "3. État Art",
        "full": "État de l'Art",
        "subs": ["3.1 OCR vs HTR", "3.2 NER : GLiNER", "3.3 Arbitrage VLM", "3.4 Synthèse comparative"]
    },
    {
        "short": "4. Archi",
        "full": "Architecture & Pipeline",
        "subs": ["4.1 Chaîne des 6 scripts", "4.2 Prétraitement d'image", "4.3 Détection YOLOv8", "4.4 Extraction & NER"]
    },
    {
        "short": "5. Fiab.",
        "full": "Fiabilisation & Interface",
        "subs": ["5.1 Application Streamlit", "5.2 Règles métier", "5.3 Contrôle cartographique"]
    },
    {
        "short": "6. Géofoncier",
        "full": "Intégration Géofoncier",
        "subs": ["6.1 Démonstration vidéo", "6.2 Protocole de versement", "6.3 Résultats Prades"]
    },
    {
        "short": "7. Limites",
        "full": "Limites & Perspectives",
        "subs": ["7.1 Limites actuelles", "7.2 Perspectives d'évolution"]
    },
    {
        "short": "8. Concl.",
        "full": "Conclusion",
        "subs": ["8.1 Bilan du projet", "8.2 Remerciements & Échanges"]
    }
]

prs = Presentation()
prs.slide_width = Inches(SLIDE_W)
prs.slide_height = Inches(SLIDE_H)
blank_layout = prs.slide_layouts[6]

# ==================== FONCTIONS DE MISE EN PAGE ====================

def set_slide_bg(slide):
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = C_BG_SLIDE

def add_header_180s(slide, active_chap_idx, active_sub_idx):
    """
    Reproduit fidèlement l'en-tête de Presentation_PFE_180s.pptx avec :
    - Logo INSA à gauche
    - Logo GEO-SIAPP à droite
    - Ligne 1 : Les 8 pilules des chapitres
    - Ligne 2 : Le ruban des sous-parties du chapitre actif
    - Ligne séparatrice inférieure
    """
    # Bande supérieure
    if os.path.exists("Logo_INSAStrasbourg.jpg"):
        slide.shapes.add_picture("Logo_INSAStrasbourg.jpg", Inches(0.50), Inches(0.12), height=Inches(0.48))
    if os.path.exists("Geosiapp.jpg"):
        slide.shapes.add_picture("Geosiapp.jpg", Inches(11.95), Inches(0.12), height=Inches(0.48))

    cx_start = 2.40
    cx_end = 11.80
    avail_w = cx_end - cx_start

    # Niveau 1 : Les 8 grands chapitres
    tab_w = (avail_w - 7 * 0.08) / 8.0
    for i, chap in enumerate(CHAPTERS_DATA):
        tx = cx_start + i * (tab_w + 0.08)
        is_act = (i == active_chap_idx)
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(tx), Inches(0.08), Inches(tab_w), Inches(0.32))
        pill.fill.solid()
        pill.fill.fore_color.rgb = C_NAVY if is_act else C_BG_TAB
        pill.line.color.rgb = C_NAVY if is_act else C_BORDER_TAB
        pill.line.width = Pt(0.5)
        
        tf = pill.text_frame
        tf.word_wrap = False
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = ("● " if is_act else "") + chap["short"]
        p.alignment = PP_ALIGN.CENTER
        r = p.runs[0]
        r.font.name = "Arial"
        r.font.size = Pt(7.5)
        r.font.bold = is_act
        r.font.color.rgb = C_WHITE if is_act else C_MUTED

    # Niveau 2 : Ruban des sous-parties pour le chapitre actif
    if active_chap_idx >= 0 and active_chap_idx < len(CHAPTERS_DATA):
        subs = CHAPTERS_DATA[active_chap_idx]["subs"]
        n_subs = len(subs)
        sub_tab_w = (avail_w - (n_subs - 1) * 0.08) / n_subs
        for j, s_label in enumerate(subs):
            sx = cx_start + j * (sub_tab_w + 0.08)
            is_sub_act = (j == active_sub_idx)
            spill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(sx), Inches(0.45), Inches(sub_tab_w), Inches(0.24))
            spill.fill.solid()
            spill.fill.fore_color.rgb = C_ACCENT if is_sub_act else C_WHITE
            spill.line.color.rgb = C_ACCENT if is_sub_act else C_LINE
            spill.line.width = Pt(0.5)
            
            stf = spill.text_frame
            stf.word_wrap = False
            stf.margin_left = stf.margin_top = stf.margin_right = stf.margin_bottom = 0
            sp_p = stf.paragraphs[0]
            sp_p.text = s_label
            sp_p.alignment = PP_ALIGN.CENTER
            sr = sp_p.runs[0]
            sr.font.name = "Arial"
            sr.font.size = Pt(7.5)
            sr.font.bold = is_sub_act
            sr.font.color.rgb = C_WHITE if is_sub_act else C_DARK

    # Ligne séparatrice
    sep = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.50), Inches(0.76), Inches(12.33), Inches(0.015))
    sep.fill.solid()
    sep.fill.fore_color.rgb = C_LINE
    sep.line.fill.background()

def create_container_card(slide):
    """Grand conteneur blanc arrondi officiel (y=1.35 à 7.15)."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.70), Inches(1.35), Inches(11.93), Inches(5.80))
    card.fill.solid()
    card.fill.fore_color.rgb = C_CARD_BG
    card.line.color.rgb = C_BORDER_TAB
    card.line.width = Pt(0.75)
    return card

def add_slide_title(slide, title_text):
    """Titre de la diapositive placé immédiatement sous l'en-tête (y=0.84)."""
    tb = slide.shapes.add_textbox(Inches(0.70), Inches(0.84), Inches(11.93), Inches(0.44))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = title_text
    r = p.runs[0]
    r.font.name = "Arial"
    r.font.size = Pt(17.5)
    r.font.bold = True
    r.font.color.rgb = C_NAVY

def add_footer_180s(slide, num):
    """Numéro de diapositive discret en bas à droite."""
    bx = slide.shapes.add_textbox(Inches(11.50), Inches(7.20), Inches(1.13), Inches(0.24))
    tf = bx.text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = f"{num} / {TOTAL_SLIDES}"
    p.alignment = PP_ALIGN.RIGHT
    r = p.runs[0]
    r.font.name = "Arial"
    r.font.size = Pt(8.5)
    r.font.color.rgb = C_MUTED

def add_fitted_picture(slide, img_path, box_x, box_y, box_w, box_h, caption=None):
    """Insère et centre une image sans aucune déformation d'aspect."""
    if not img_path or not os.path.exists(img_path):
        return None
    try:
        cap_h = 0.30 if caption else 0.0
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
        print(f"Erreur add_fitted_picture ({img_path}) : {e}")

def init_standard_slide(chap_idx, sub_idx, slide_num, slide_title):
    """Initialise une diapositive complète au format 180s."""
    slide = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide)
    add_header_180s(slide, chap_idx, sub_idx)
    add_slide_title(slide, slide_title)
    card = create_container_card(slide)
    add_footer_180s(slide, slide_num)
    return slide, card

def add_two_column_content(slide, items, img_path, img_caption, callout_title=None, callout_text=None, kpis=None):
    """Remplit le grand conteneur blanc avec une colonne texte à gauche et une image à droite."""
    card_x = 0.70
    card_y = 1.35
    card_w = 11.93
    card_h = 5.80
    
    # Colonne gauche (largeur 5.75)
    left_x = card_x + 0.25
    left_w = 5.65
    start_y = card_y + 0.25
    
    total_items = len(items)
    gap_y = 0.12
    bot_reserve = 1.10 if (callout_title or kpis) else 0.0
    avail_h = (card_h - 0.50) - bot_reserve
    item_h = (avail_h - (total_items - 1) * gap_y) / total_items
    
    for i, item in enumerate(items):
        iy = start_y + i * (item_h + gap_y)
        # Petite carte blanche interne avec bordure subtile
        b_shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left_x), Inches(iy), Inches(left_w), Inches(item_h))
        b_shp.fill.solid()
        b_shp.fill.fore_color.rgb = C_BG_SLIDE
        b_shp.line.color.rgb = C_LINE
        b_shp.line.width = Pt(0.75)
        
        bx = slide.shapes.add_textbox(Inches(left_x + 0.18), Inches(iy + 0.08), Inches(left_w - 0.36), Inches(item_h - 0.16))
        tf = bx.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        
        p = tf.paragraphs[0]
        if isinstance(item, tuple):
            head, body = item
            r1 = p.add_run()
            r1.text = head + "\n"
            r1.font.name = "Arial"
            r1.font.size = Pt(11.0)
            r1.font.bold = True
            r1.font.color.rgb = C_NAVY
            
            p2 = tf.add_paragraph()
            p2.space_before = Pt(2)
            r2 = p2.add_run()
            r2.text = body
            r2.font.name = "Arial"
            r2.font.size = Pt(9.5)
            r2.font.color.rgb = C_MUTED
        else:
            r = p.add_run()
            r.text = item
            r.font.name = "Arial"
            r.font.size = Pt(10.0)
            r.font.color.rgb = C_DARK

    if callout_title and callout_text:
        cy = card_y + card_h - bot_reserve - 0.15
        c_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left_x), Inches(cy), Inches(left_w), Inches(bot_reserve))
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = C_CALLOUT_BG
        c_box.line.color.rgb = C_ACCENT
        c_box.line.width = Pt(1.0)
        
        bx_c = slide.shapes.add_textbox(Inches(left_x + 0.18), Inches(cy + 0.08), Inches(left_w - 0.36), Inches(bot_reserve - 0.16))
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
        r_c2 = p_c2.add_run()
        r_c2.text = callout_text
        r_c2.font.name = "Arial"
        r_c2.font.size = Pt(8.5)
        r_c2.font.color.rgb = C_DARK
    elif kpis:
        cy = card_y + card_h - bot_reserve - 0.15
        n_kpi = len(kpis)
        kw = (left_w - (n_kpi - 1) * 0.10) / n_kpi
        for ki, (kv, kl) in enumerate(kpis):
            kx = left_x + ki * (kw + 0.10)
            k_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(kx), Inches(cy), Inches(kw), Inches(bot_reserve))
            k_box.fill.solid()
            k_box.fill.fore_color.rgb = C_NAVY
            k_box.line.fill.background()
            
            bx_k = slide.shapes.add_textbox(Inches(kx + 0.04), Inches(cy + 0.12), Inches(kw - 0.08), Inches(bot_reserve - 0.24))
            tf_k = bx_k.text_frame
            tf_k.margin_left = tf_k.margin_top = tf_k.margin_right = tf_k.margin_bottom = 0
            pk1 = tf_k.paragraphs[0]
            pk1.text = kv
            pk1.alignment = PP_ALIGN.CENTER
            rk1 = pk1.runs[0]
            rk1.font.name = "Arial"
            rk1.font.size = Pt(15.5)
            rk1.font.bold = True
            rk1.font.color.rgb = C_ACCENT
            
            pk2 = tf_k.add_paragraph()
            pk2.space_before = Pt(2)
            pk2.text = kl
            pk2.alignment = PP_ALIGN.CENTER
            rk2 = pk2.runs[0]
            rk2.font.name = "Arial"
            rk2.font.size = Pt(7.5)
            rk2.font.color.rgb = C_WHITE

    # Colonne droite (largeur 5.75)
    right_x = card_x + card_w - 5.75 - 0.25
    right_w = 5.75
    right_h = card_h - 0.50
    right_y = card_y + 0.25
    
    # Cadre discret pour l'image
    frame_r = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(right_x), Inches(right_y), Inches(right_w), Inches(right_h))
    frame_r.fill.solid()
    frame_r.fill.fore_color.rgb = C_BG_SLIDE
    frame_r.line.color.rgb = C_LINE
    frame_r.line.width = Pt(0.75)
    
    add_fitted_picture(slide, img_path, right_x + 0.15, right_y + 0.15, right_w - 0.30, right_h - 0.30, caption=img_caption)

# ==================== SLIDE 1 : TITRE (STYLE PFE 180S) ====================
print("Génération Slide 1 (Titre 180s)...")
s1 = prs.slides.add_slide(blank_layout)
set_slide_bg(s1)

if os.path.exists("Logo_INSAStrasbourg.jpg"):
    s1.shapes.add_picture("Logo_INSAStrasbourg.jpg", Inches(0.80), Inches(0.50), width=Inches(2.70))
if os.path.exists("Geosiapp.jpg"):
    add_fitted_picture(s1, "Geosiapp.jpg", 10.60, 0.45, 1.93, 1.10)

card1 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.80), Inches(1.75), Inches(11.73), Inches(5.25))
card1.fill.solid()
card1.fill.fore_color.rgb = C_WHITE
card1.line.color.rgb = C_BORDER_TAB
card1.line.width = Pt(1.0)

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
sep1.fill.fore_color.rgb = C_LINE
sep1.line.fill.background()

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

# ==================== SLIDE 2 : SOMMAIRE (STYLE 180S) ====================
print("Génération Slide 2 (Sommaire 180s)...")
s2, c2 = init_standard_slide(-1, -1, 2, "Plan de la présentation : Sommaire")

# Grille des 8 chapitres dans le grand conteneur blanc
cw2 = (11.93 - 0.50 - 3 * 0.15) / 4.0
ch2 = (5.80 - 0.50 - 0.15) / 2.0
sx2 = 0.70 + 0.25
sy2 = 1.35 + 0.25

for idx, chap in enumerate(CHAPTERS_DATA):
    row = idx // 4
    col = idx % 4
    x = sx2 + col * (cw2 + 0.15)
    y = sy2 + row * (ch2 + 0.15)
    
    b_card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(cw2), Inches(ch2))
    b_card.fill.solid()
    b_card.fill.fore_color.rgb = C_BG_SLIDE
    b_card.line.color.rgb = C_LINE
    b_card.line.width = Pt(0.75)
    
    # En-tête bleu nuit
    h_box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(cw2), Inches(0.55))
    h_box.fill.solid()
    h_box.fill.fore_color.rgb = C_NAVY
    h_box.line.fill.background()
    
    bx_ht = s2.shapes.add_textbox(Inches(x + 0.08), Inches(y + 0.08), Inches(cw2 - 0.16), Inches(0.40))
    tf_ht = bx_ht.text_frame
    tf_ht.word_wrap = True
    p_ht = tf_ht.paragraphs[0]
    p_ht.text = chap["full"]
    r_ht = p_ht.runs[0]
    r_ht.font.name = "Arial"
    r_ht.font.size = Pt(9.5)
    r_ht.font.bold = True
    r_ht.font.color.rgb = C_WHITE
    
    # Sous-parties
    bx_sub = s2.shapes.add_textbox(Inches(x + 0.12), Inches(y + 0.65), Inches(cw2 - 0.24), Inches(ch2 - 0.75))
    tf_sub = bx_sub.text_frame
    tf_sub.word_wrap = True
    tf_sub.margin_left = tf_sub.margin_top = tf_sub.margin_right = tf_sub.margin_bottom = 0
    
    for si, s_item in enumerate(chap["subs"]):
        p_si = tf_sub.paragraphs[0] if si == 0 else tf_sub.add_paragraph()
        if si > 0:
            p_si.space_before = Pt(3)
        r_si = p_si.add_run()
        r_si.text = "• " + s_item
        r_si.font.name = "Arial"
        r_si.font.size = Pt(8.5)
        r_si.font.color.rgb = C_DARK

# ==================== CHAPITRE 1 : INTRODUCTION & CONTEXTE ====================
print("Génération Chapitre 1 (Slides 3-6)...")

# Slide 3 (1.1)
s3, c3 = init_standard_slide(0, 0, 3, "1.1 Présentation de la structure d'accueil : Le Cabinet GEO-SIAPP")
add_two_column_content(
    s3,
    items=[
        ("Une implantation forte en Ardèche et Drôme",
         "Fondé en 1992 à Aubenas, le cabinet GEO-SIAPP regroupe 4 agences : Aubenas (siège), Pierrelatte, Vallon-Pont-d'Arc et Guilherand-Granges."),
        ("Pluridisciplinarité des activités",
         "Missions foncières (bornage, divisions), urbanisme réglementaire, ingénierie VRD et topographie de précision (scanner laser 3D et drone)."),
        ("Un fonds historique de 23 600 dossiers",
         "Conservation intégrale de 50 ans d'archives physiques (1959-2007) issues du cabinet et des études de géomètres rachetées au fil des années.")
    ],
    img_path="img/bureau_cropped.jpg",
    img_caption="Siège principal du Cabinet GEO-SIAPP à Aubenas (Ardèche)",
    kpis=[("4", "AGENCES"), ("23 600", "DOSSIERS"), ("1959-2007", "ARCHIVES"), ("6 MOIS", "STAGE PFE")]
)

# Slide 4 (1.2)
s4, c4 = init_standard_slide(0, 1, 4, "1.2 Le cadre réglementaire et professionnel du bornage")
add_two_column_content(
    s4,
    items=[
        ("Monopole légal du géomètre-expert (Loi du 7 mai 1946)",
         "Seul professionnel habilité à fixer les limites réelles de la propriété foncière. L'acte de bornage possède une valeur juridique opposable aux tiers."),
        ("Obligation de conservation trentenaire (Décret n° 96-478)",
         "L'article 55 impose la conservation des archives foncières pendant au moins 30 ans, y compris après la cessation d'activité du professionnel."),
        ("Le portail national Géofoncier (OGE)",
         "Plateforme centrale de l'Ordre des Géomètres-Experts référençant l'ensemble des interventions foncières sous forme de pastilles géolocalisées.")
    ],
    img_path="img/Chap2_pastilles.jpg",
    img_caption="Pastilles d'intervention foncière visualisées sur le portail Géofoncier",
    callout_title="DISTINCTION JURIDIQUE FONDAMENTALE",
    callout_text="Le programme national GEODÉMAT (DGFiP) ne traite que le cadastre public : la numérisation des archives foncières privées relève exclusivement des cabinets."
)

# Slide 5 (1.3)
s5, c5 = init_standard_slide(0, 2, 5, "1.3 Contexte du projet : La recherche d'antériorité obligatoire")
add_two_column_content(
    s5,
    items=[
        ("L'adage fondamental « Bornage sur bornage ne vaut »",
         "Conformément à l'article 646 du Code civil, toute nouvelle opération impose de retrouver l'acte antérieur pour garantir la continuité juridique des limites."),
        ("Une recherche physique chronophage et coûteuse",
         "Retrouver un vieux dossier prend 15 à 30 minutes au sein des archives physiques, sans garantie de classement, entraînant des pertes de productivité."),
        ("Des dossiers physiques parfois inaccessibles",
         "L'absence d'indexation numérique rend certains dossiers introuvables lors des chantiers urgents, certains cabinets facturant même cette recherche.")
    ],
    img_path="img/registre_crop_opt.jpg",
    img_caption="Extrait d'un registre foncier manuscrit conservé au cabinet GEO-SIAPP",
    callout_title="OBJECTIF OPÉRATIONNEL",
    callout_text="Verser les archives sur Géofoncier permet aux géomètres de localiser instantanément tout acte antérieur sur le terrain via la carte interactive."
)

# Slide 6 (1.4)
s6, c6 = init_standard_slide(0, 3, 6, "1.4 Problématique foncière et contraintes techniques")
add_two_column_content(
    s6,
    items=[
        ("Hétérogénéité documentaire sur 50 ans (1959-2007)",
         "Supports très variés (calques, registres manuscrits, tapuscrits), encres dégradées, pliures d'usure et scans inclinés ou de faible résolution."),
        ("Ambiguïtés toponymiques contextuelles locales",
         "Abréviations manuscrites incompréhensibles hors contexte ardéchois (ex. « La Chap./A. » pour La Chapelle-sous-Aubenas)."),
        ("Échec des logiciels d'OCR standards",
         "Sur l'écriture manuscrite des registres, les OCR classiques dépassent 30% de taux d'erreur de caractères (CER).")
    ],
    img_path="img/chapelle_sous_aubenas_crop.png",
    img_caption="Registre SIAPP illustrant l'abréviation locale « La Chap./A. »",
    callout_title="CONTRAINTE DE SOUVERAINETÉ STRICTE",
    callout_text="Les données foncières étant personnelles et confidentielles, l'outil doit obligatoirement fonctionner à 100% en local sur un PC de bureau, sans aucun cloud."
)

# ==================== CHAPITRE 2 : ANALYSE MÉTIER & DONNÉES ====================
print("Génération Chapitre 2 (Slides 7-9)...")

# Slide 7 (2.1)
s7, c7 = init_standard_slide(1, 0, 7, "2.1 Cadastre fiscal et propriété foncière : distinction juridique")
add_two_column_content(
    s7,
    items=[
        ("Le cadastre : un rôle purement fiscal",
         "Créé en 1807, le plan cadastral (PCI DGFiP) ne fournit qu'une présomption fiscale. Il ne constitue ni un titre de propriété ni une délimitation juridique."),
        ("L'imprécision du plan cadastral en zone rurale",
         "Issu d'anciens plans napoléoniens vectorisés, le cadastre présente souvent des décalages métriques avec la réalité du terrain."),
        ("L'acte de bornage : la vérité juridique du terrain",
         "Seul le procès-verbal dressé par le géomètre-expert et signé par les riverains fixe définitivement les limites réelles, matérialisées par des bornes.")
    ],
    img_path="img/r4p_fig2_tight.png",
    img_caption="Plan d'arpentage et délimitation contradictoire des propriétés",
    callout_title="IMPORTANCE DES ARCHIVES PRIVÉES",
    callout_text="En cas de litige, la justice s'appuie en priorité sur les PV de bornage anciens conservés par les cabinets, le cadastre n'étant que subsidiaire."
)

# Slide 8 (2.2)
s8, c8 = init_standard_slide(1, 1, 8, "2.2 Le versement sur Géofoncier : Spécifications de l'API REST")
add_two_column_content(
    s8,
    items=[
        ("Authentification sécurisée par jeton JWT Bearer",
         "Connexion via identifiants officiels du géomètre-expert avec délivrance d'un jeton éphémère (validité 24h) garantissant la traçabilité des dépôts."),
        ("Les 5 métadonnées obligatoires par dossier",
         "1. Code INSEE (5 chiffres) • 2. Date d'acte (ISO 8601 YYYY-MM-DD) • 3. Contenance en m² • 4. Nature de l'opération • 5. Type d'acte OGE."),
        ("Gestion des réponses et validation synchrone",
         "Retour HTTP 201 Created confirmant la publication de la pastille. En cas d'erreur de format, retour HTTP 400 ou 422 avec motif détaillé.")
    ],
    img_path="img/Etape 3_Versement_geofoncier.jpg",
    img_caption="Interface de versement et métadonnées cibles sur Géofoncier",
    callout_title="GÉOLOCALISATION LAMBERT-93",
    callout_text="L'API exige également le rattachement géographique du dossier (centroïde ou emprise polygonale) en coordonnées officielles Lambert-93 (EPSG:2154)."
)

# Slide 9 (2.3)
s9, c9 = init_standard_slide(1, 2, 9, "2.3 Typologie des actes fonciers et caractéristiques des fonds SIAPP")
add_two_column_content(
    s9,
    items=[
        ("Les DMPC (Documents Modificatifs du Parcellaire Cadastral)",
         "Formulaires Cerfa standardisés accompagnant les divisions. Cartouche normé facilitant le repérage géométrique des informations clés."),
        ("Les Procès-Verbaux de bornage",
         "Actes sous seing privé rédigés en texte libre manuscrit. Grande variabilité de mise en page selon les géomètres, signatures et croquis annexés."),
        ("Les registres d'affaires centralisateurs",
         "Cahiers récapitulant l'ensemble des dossiers au fil des décennies. Chaque ligne contient le numéro d'affaire, la date, la commune et le client.")
    ],
    img_path="img/4513_DA_124_crop.png",
    img_caption="Registre d'affaires manuscrit SIAPP (commune de Prades)",
    callout_title="DOCUMENTS HORS PÉRIMÈTRE",
    callout_text="Les pièces de travail internes (croquis de terrain bruts, carnets de mesures) ne sont pas destinées à être versées et restent exclues du pipeline."
)

# ==================== CHAPITRE 3 : ÉTAT DE L'ART ====================
print("Génération Chapitre 3 (Slides 10-13)...")

# Slide 10 (3.1)
s10, c10 = init_standard_slide(2, 0, 10, "3.1 Reconnaissance de texte : OCR imprimé vs HTR manuscrit")
add_two_column_content(
    s10,
    items=[
        ("EasyOCR : adapté aux cartouches imprimés récents",
         "Architecture CRAFT (détection) + CRNN (reconnaissance). Rapide (< 0.5s par page), efficace sur texte dactylographié mais limité sur cursive."),
        ("TrOCR (Microsoft Research, Li et al., 2021) : le standard HTR",
         "Transformer Vision-Langue bout-en-bout couplant un encodeur visuel (DeiT) et un décodeur linguistique (RoBERTa). Pré-entraîné sur IAM et IIIT5K."),
        ("L'écosystème ouvert HTR-United",
         "Dépôt collaboratif de modèles et de jeux de données spécialisés dans les écritures historiques françaises avec segmentation fine ligne par ligne.")
    ],
    img_path="img/trocr_architecture_figure1.png",
    img_caption="Architecture Transformer encodeur-décodeur de TrOCR (Li et al., 2021)",
    callout_title="STRATÉGIE HYBRIDE RETENUE",
    callout_text="Utilisation conjointe d'EasyOCR pour les cartouches imprimés et de TrOCR pour le déchiffrement des écritures manuscrites des registres."
)

# Slide 11 (3.2)
s11, c11 = init_standard_slide(2, 1, 11, "3.2 Reconnaissance d'Entités Nommées (NER) : GLiNER vs LayoutLM")
add_two_column_content(
    s11,
    items=[
        ("GLiNER (Zaratiana et al., 2023 - urchade/GLiNER)",
         "Modèle bi-encodeur BERT compact (340M paramètres, licence Apache 2.0). Extraction zero-shot d'entités sans entraînement préalable requis."),
        ("LayoutLM (Microsoft Research, 2020)",
         "Modèle Document AI multimodal intégrant texte, coordonnées 2D et image. Conçu spécifiquement pour les formulaires réguliers stricts."),
        ("Pourquoi GLiNER est supérieur sur notre corpus",
         "Les registres fonciers manuscrits ne respectent aucune grille tabulaire rigide. GLiNER excelle sur le texte libre et tolère les variations de syntaxe.")
    ],
    img_path="img/gliner_prompt_render.png",
    img_caption="Fonctionnement zero-shot de GLiNER associant texte et entités cibles au runtime",
    callout_title="EXTRACTION DYNAMIQUE DES 5 ENTITÉS",
    callout_text="Les labels cibles (commune, date, contenance, nature, référence) sont interrogés dynamiquement avec un score de confiance calibré [0, 1]."
)

# Slide 12 (3.3)
s12, c12 = init_standard_slide(2, 2, 12, "3.3 Modèles Vision-Langage (VLM) : Arbitrage local par Ollama")
add_two_column_content(
    s12,
    items=[
        ("Inférence souveraine sur GPU local via Ollama",
         "Exécution de modèles multimodaux (LLaVA 1.6, Qwen2-VL) sur une carte graphique de bureau (8 Go VRAM) sans aucune dépendance au cloud."),
        ("Rôle ciblé : arbitre de dernier recours",
         "Le VLM n'est pas utilisé systématiquement en raison de sa latence (~8s), mais activé uniquement quand la confiance GLiNER est inférieure à 0.65."),
        ("Résolution locale des ambiguïtés toponymiques",
         "Sur l'abréviation « La Chap./A. », le modèle analyse visuellement l'en-tête du registre, consulte la liste des communes et restitue « La Chapelle-sous-Aubenas ».")
    ],
    img_path="img/vlm_arbitrage_chapelle_exact.png",
    img_caption="Démonstration d'arbitrage VLM : désambiguïsation locale de « La Chap./A. »",
    callout_title="SOUVERAINETÉ ABSOLUE",
    callout_text="L'arbitrage visuel local résout les cas complexes tout en garantissant qu'aucune donnée foncière ne quitte les locaux du cabinet."
)

# Slide 13 (3.4) : Tableau comparatif
print("Génération Slide 13 (Synthèse comparative)...")
s13, c13 = init_standard_slide(2, 3, 13, "3.4 Synthèse comparative et justification des choix technologiques")

headers13 = ["Technologie", "Type", "Force principale", "Limite identifiée", "Décision dans le pipeline"]
rows13 = [
    ["EasyOCR", "OCR classique", "Très rapide (< 0.5s), multilingue", "Échoue sur cursive (CER > 30%)", "Retenu : cartouches imprimés"],
    ["TrOCR", "HTR Transformer", "Robuste sur écriture manuscrite", "Nécessite une segmentation fine", "Retenu : registres manuscrits"],
    ["GLiNER", "NER Zero-shot", "Ultra-flexible, classes dynamiques", "Sensible aux textes trop longs", "Retenu : extraction d'entités"],
    ["LayoutLM", "Document AI 2D", "Excellent sur formulaires stricts", "Inadapté au manuscrit non tabulaire", "Écarté"],
    ["VLM (Ollama)", "Multimodal local", "Compréhension visuelle contextuelle", "Temps de calcul élevé (~8s / vue)", "Retenu : arbitre de secours"]
]

cw13 = [1.90, 1.60, 2.90, 2.90, 2.13]
rh13 = 0.55
tx13 = 0.95
ty13 = 1.60

hx = tx13
for hi, htext in enumerate(headers13):
    h_box = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(hx), Inches(ty13), Inches(cw13[hi]), Inches(rh13))
    h_box.fill.solid()
    h_box.fill.fore_color.rgb = C_NAVY
    h_box.line.fill.background()
    
    bx = s13.shapes.add_textbox(Inches(hx + 0.05), Inches(ty13 + 0.12), Inches(cw13[hi] - 0.10), Inches(rh13 - 0.20))
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
    hx += cw13[hi]

for ri, row in enumerate(rows13):
    ry = ty13 + (ri + 1) * rh13
    bg_r = C_CARD_BG if ri % 2 == 0 else C_BG_SLIDE
    rx = tx13
    for ci, cell in enumerate(row):
        c_box = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(rx), Inches(ry), Inches(cw13[ci]), Inches(rh13))
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = bg_r
        c_box.line.color.rgb = C_LINE
        c_box.line.width = Pt(0.75)
        
        bx = s13.shapes.add_textbox(Inches(rx + 0.08), Inches(ry + 0.10), Inches(cw13[ci] - 0.16), Inches(rh13 - 0.20))
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
        rx += cw13[ci]

# Bandeau synthèse
sy13 = ty13 + (len(rows13) + 1) * rh13 + 0.20
s_box = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(tx13), Inches(sy13), Inches(sum(cw13)), Inches(0.65))
s_box.fill.solid()
s_box.fill.fore_color.rgb = C_CALLOUT_BG
s_box.line.color.rgb = C_ACCENT
s_box.line.width = Pt(1.0)
bx_syn = s13.shapes.add_textbox(Inches(tx13 + 0.20), Inches(sy13 + 0.12), Inches(sum(cw13) - 0.40), Inches(0.40))
tf_syn = bx_syn.text_frame
p_syn = tf_syn.paragraphs[0]
p_syn.text = "SYNTHÈSE DE LA CHAÎNE : Prétraitement d'image + OCR / HTR hybride + GLiNER zero-shot + Arbitre VLM local."
p_syn.alignment = PP_ALIGN.CENTER
r_syn = p_syn.runs[0]
r_syn.font.name = "Arial"
r_syn.font.size = Pt(10.5)
r_syn.font.bold = True
r_syn.font.color.rgb = C_NAVY

# ==================== CHAPITRE 4 : ARCHITECTURE & PIPELINE ====================
print("Génération Chapitre 4 (Slides 14-17)...")

# Slide 14 (4.1) : 6 scripts modulaires
s14, c14 = init_standard_slide(3, 0, 14, "4.1 Architecture logicielle : 6 scripts Python modulaires")
scripts_info = [
    ("main.py", "1. Orchestrateur", [
        ("Rôle", "Orchestration & flux"),
        ("Entrée", "Scans bruts du fonds"),
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
        ("Sécurité", "Jeton JWT Bearer 24h"),
        ("Endpoint", "POST /dossiers REST OGE")
    ])
]

cw14 = (11.93 - 0.40 - 5 * 0.10) / 6.0
for si, (sname, stitle, sposts) in enumerate(scripts_info):
    sx = 0.70 + 0.20 + si * (cw14 + 0.10)
    sy = 1.60
    sh = 4.30
    
    c_card = s14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(sx), Inches(sy), Inches(cw14), Inches(sh))
    c_card.fill.solid()
    c_card.fill.fore_color.rgb = C_BG_SLIDE
    c_card.line.color.rgb = C_LINE
    c_card.line.width = Pt(0.75)
    
    h_box = s14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(sx), Inches(sy), Inches(cw14), Inches(0.55))
    h_box.fill.solid()
    h_box.fill.fore_color.rgb = C_NAVY
    h_box.line.fill.background()
    
    bx_st = s14.shapes.add_textbox(Inches(sx + 0.04), Inches(sy + 0.08), Inches(cw14 - 0.08), Inches(0.40))
    tf_st = bx_st.text_frame
    tf_st.word_wrap = True
    p_st = tf_st.paragraphs[0]
    p_st.alignment = PP_ALIGN.CENTER
    p_st.text = stitle
    r_st = p_st.runs[0]
    r_st.font.name = "Arial"
    r_st.font.size = Pt(8.5)
    r_st.font.bold = True
    r_st.font.color.rgb = C_WHITE
    
    bx_sn = s14.shapes.add_textbox(Inches(sx + 0.04), Inches(sy + 0.65), Inches(cw14 - 0.08), Inches(0.32))
    tf_sn = bx_sn.text_frame
    p_sn = tf_sn.paragraphs[0]
    p_sn.alignment = PP_ALIGN.CENTER
    p_sn.text = sname
    r_sn = p_sn.runs[0]
    r_sn.font.name = "Arial"
    r_sn.font.size = Pt(9.5)
    r_sn.font.bold = True
    r_sn.font.color.rgb = C_ACCENT
    
    bx_sd = s14.shapes.add_textbox(Inches(sx + 0.06), Inches(sy + 1.05), Inches(cw14 - 0.12), Inches(3.10))
    tf_sd = bx_sd.text_frame
    tf_sd.word_wrap = True
    tf_sd.margin_left = tf_sd.margin_top = tf_sd.margin_right = tf_sd.margin_bottom = 0
    
    for pi, (label_p, val_p) in enumerate(sposts):
        p_item = tf_sd.paragraphs[0] if pi == 0 else tf_sd.add_paragraph()
        if pi > 0:
            p_item.space_before = Pt(5)
        r_lbl = p_item.add_run()
        r_lbl.text = label_p + " :\n"
        r_lbl.font.name = "Arial"
        r_lbl.font.size = Pt(8.5)
        r_lbl.font.bold = True
        r_lbl.font.color.rgb = C_NAVY
        
        p_val = tf_sd.add_paragraph()
        p_val.space_before = Pt(1)
        r_val = p_val.add_run()
        r_val.text = val_p
        r_val.font.name = "Arial"
        r_val.font.size = Pt(8.0)
        r_val.font.color.rgb = C_MUTED

# Bandeau bas workflow
wf_box = s14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.90), Inches(6.05), Inches(11.53), Inches(0.60))
wf_box.fill.solid()
wf_box.fill.fore_color.rgb = C_CALLOUT_BG
wf_box.line.color.rgb = C_ACCENT
wf_box.line.width = Pt(1.0)
bx_wf = s14.shapes.add_textbox(Inches(1.00), Inches(6.15), Inches(11.33), Inches(0.40))
tf_wf = bx_wf.text_frame
p_wf = tf_wf.paragraphs[0]
p_wf.alignment = PP_ALIGN.CENTER
p_wf.text = "FLUX OPÉRATIONNEL : Scan brut  →  Classification  →  Détection YOLOv8  →  OCR / HTR  →  Contrôle qualité  →  Versement API"
r_wf = p_wf.runs[0]
r_wf.font.name = "Arial"
r_wf.font.size = Pt(9.5)
r_wf.font.bold = True
r_wf.font.color.rgb = C_NAVY

# Slide 15 (4.2)
s15, c15 = init_standard_slide(3, 1, 15, "4.2 Prétraitement d'image : Redressement Hough et binarisation Otsu")
add_two_column_content(
    s15,
    items=[
        ("Contrôle automatique de conformité résolution",
         "Vérification du seuil minimal de 300 DPI. En-dessous de 200 DPI, alerte automatique de l'opérateur pour prévenir les échecs de reconnaissance."),
        ("Redressement d'orientation par transformée de Hough",
         "Détection des lignes de grille du tableau pour estimer et corriger automatiquement l'angle d'inclinaison du scan (correction jusqu'à ±15°)."),
        ("Binarisation adaptative d'Otsu",
         "Séparation dynamique fond / encre, supprimant le jaunissement du papier et les ombres de numérisation tout en renforçant les traits fins.")
    ],
    img_path="img/4513_DA_124_centered_crop.png",
    img_caption="Document après redressement Hough et binarisation d'Otsu",
    callout_title="IMPACT MESURÉ SUR CORPUS TEST",
    callout_text="Le prétraitement réduit le taux d'erreur de caractères (CER) de 18 points par rapport à un OCR appliqué sur scan brut non corrigé."
)

# Slide 16 (4.3)
s16, c16 = init_standard_slide(3, 2, 16, "4.3 Détection spatiale par YOLOv8 nano")
add_two_column_content(
    s16,
    items=[
        ("YOLOv8 nano : légèreté et performance",
         "Modèle ultra-léger (3.2M paramètres) entraîné sur 400 documents annotés du fonds SIAPP. Inférence en moins de 25 ms sur GPU de bureau."),
        ("3 classes d'éléments structurants détectées",
         "Localisation précise du cartouche officiel, du tableau de contenances parcellaires et des mentions marginales manuscrites (mAP@0.5 = 0.87)."),
        ("Isolation des zones utiles du document",
         "Permet de ne traiter que les régions contenant les 5 métadonnées obligatoires et d'éliminer les millions de pixels parasites du plan.")
    ],
    img_path="img/composite_yolo_extraction.png",
    img_caption="Détection par boîtes englobantes YOLOv8 des zones du document",
    callout_title="SEGMENTATION GÉOMÉTRIQUE STRICTE",
    callout_text="La détection spatiale garantit que les moteurs OCR et NER ne s'exécutent que sur les zones pertinentes sans bruit textuel environnant."
)

# Slide 17 (4.4)
s17, c17 = init_standard_slide(3, 3, 17, "4.4 Extraction ciblée OCR et reconnaissance d'entités GLiNER")
add_two_column_content(
    s17,
    items=[
        ("Transmission ciblée des boîtes englobantes",
         "Le texte brut issu des régions découpées par YOLOv8 est transmis directement à GLiNER, évitant les surcharges de contexte."),
        ("Gain substantiel sur le score F1 d'extraction",
         "Le ciblage spatial préalable permet de faire passer le score F1 global de l'extraction d'entités de 0.72 à 0.88 sur le corpus de test."),
        ("Déclenchement sélectif de l'arbitre VLM",
         "Seules les entités dont le score de confiance résiduel est inférieur à 0.65 sont transmises à Ollama pour arbitrage visuel local.")
    ],
    img_path="img/plan_yolo_framed.png",
    img_caption="Plan foncier découpé spatialement avant extraction sémantique",
    callout_title="SYNERGIE VISION ET EXTRACTION",
    callout_text="La combinaison d'un découpage spatial par vision (YOLO) et d'un bi-encodeur linguistique (GLiNER) assure la robustesse du pipeline."
)

# ==================== CHAPITRE 5 : FIABILISATION & INTERFACE ====================
print("Génération Chapitre 5 (Slides 18-20)...")

# Slide 18 (5.1)
s18, c18 = init_standard_slide(4, 0, 18, "5.1 Interface de relecture assistée Streamlit")
add_two_column_content(
    s18,
    items=[
        ("Ergonomie optimisée pour l'opérateur",
         "Double affichage synchronisé : document original scanné à droite et formulaire pré-rempli par l'IA à gauche pour un contrôle immédiat."),
        ("Indicateurs de confiance à trois niveaux",
         "Code couleur intuitif : Vert (> 0.85) validé en 1 clic • Orange (0.65 - 0.85) à vérifier visuellement • Rouge (< 0.65) saisie manuelle requise."),
        ("Gain de temps opérationnel constaté",
         "Le temps moyen de relecture et de validation par un technicien passe de 15-30 minutes à seulement 1 à 2 minutes par dossier.")
    ],
    img_path="img/interface_haut.jpg",
    img_caption="Interface Streamlit : double vue scan / formulaire avec jauges de confiance",
    callout_title="LE GÉOMÈTRE GARDE LE CONTRÔLE",
    callout_text="L'outil ne remplace pas l'humain : il pré-remplit les champs et laisse la décision finale de signature au technicien du cabinet."
)

# Slide 19 (5.2)
s19, c19 = init_standard_slide(4, 1, 19, "5.2 Moteur de 17 règles de cohérence métier")
add_two_column_content(
    s19,
    items=[
        ("Correction toponymique par distance de Levenshtein",
         "Rapprochement automatique du nom extrait avec la table officielle des 335 communes d'Ardèche (tolérance calibrée à 2 fautes de frappe)."),
        ("Contrôle de validité temporelle stricte (ISO 8601)",
         "Vérification du format de date, des années bissextiles et de la cohérence historique de l'acte foncier (bornes 1950 - 2010)."),
        ("Plausibilité des contenances parcellaires",
         "Filtrage des surfaces hors plage réaliste (10 m² à 50 ha) et conversion automatique des unités historiques (ares, centiares)."),
        ("Détection préventive des doublons fonciers",
         "Interrogation de l'API Géofoncier avant validation pour interdire tout ré-enregistrement d'un dossier portant la même référence.")
    ],
    img_path="img/zoom_bandeau_coherence.jpg",
    img_caption="Bandeau d'alerte et contrôles de cohérence intégrés à l'interface",
    callout_title="SÉCURISATION DU PROCESSUS",
    callout_text="Aucun dossier n'est envoyé à l'API Géofoncier sans avoir franchi avec succès l'intégralité des 17 règles de cohérence."
)

# Slide 20 (5.3)
s20, c20 = init_standard_slide(4, 2, 20, "5.3 Contrôle cartographique interactif (Folium)")
add_two_column_content(
    s20,
    items=[
        ("Géolocalisation automatique sur carte OpenStreetMap",
         "Positionnement instantané de la pastille sur fond cadastral et vue satellite pour valider la cohérence spatiale de la commune."),
        ("Interaction cartographique dynamique",
         "L'opérateur peut cliquer sur la pastille pour afficher la fiche récapitulative : référence dossier, date d'intervention et surface."),
        ("Vérification visuelle de proximité foncière",
         "Permet de constater immédiatement la présence d'autres interventions du cabinet dans le voisinage de la parcelle concernée.")
    ],
    img_path="img/streamlit_map_app.jpg",
    img_caption="Module cartographique Folium intégré à l'application Streamlit",
    callout_title="VALIDATION SPATIALE COMPLÈTE",
    callout_text="L'intégration cartographique garantit que les coordonnées Lambert-93 calculées correspondent parfaitement à la commune déclarée."
)

# ==================== CHAPITRE 6 : INTÉGRATION GÉOFONCIER ====================
print("Génération Chapitre 6 (Slides 21-23)...")

# Slide 21 (6.1) : Vidéo Démonstration
print("Génération Slide 21 (Vidéo 180s)...")
s21, c21 = init_standard_slide(5, 0, 21, "6.1 Démonstration du pipeline en conditions réelles")

# Intégration de la vidéo à l'intérieur du conteneur blanc
vid_path = "Demo_soutenance_final.mp4"
poster_img = "video_poster.jpg" if os.path.exists("video_poster.jpg") else None
vw, vh = 11.40, 5.30
vx = 0.70 + (11.93 - vw) / 2.0
vy = 1.35 + (5.80 - vh) / 2.0

if os.path.exists(vid_path):
    try:
        s21.shapes.add_movie(vid_path, Inches(vx), Inches(vy), Inches(vw), Inches(vh), poster_frame_image=poster_img, mime_type="video/mp4")
        print(f"  Vidéo intégrée dans le conteneur blanc : {vid_path}")
    except Exception as e:
        print(f"  Erreur add_movie : {e}")
        add_fitted_picture(s21, poster_img, vx, vy, vw, vh, caption=f"Vidéo prête pour lecture : {vid_path}")
else:
    add_fitted_picture(s21, poster_img, vx, vy, vw, vh, caption="Fichier vidéo non trouvé")

# Slide 22 (6.2)
s22, c22 = init_standard_slide(5, 1, 22, "6.2 Protocole de versement via l'API REST Géofoncier")
add_two_column_content(
    s22,
    items=[
        ("Authentification sécurisée par token JWT",
         "Génération d'un token Bearer éphémère à partir des identifiants certifiés du cabinet GEO-SIAPP pour authentifier chaque requête."),
        ("Téléversement multipart synchrone",
         "Envoi simultané des métadonnées JSON vérifiées et du scan PDF du dossier vers le point d'entrée officiel POST /dossiers de l'OGE."),
        ("Traitement par lot et journalisation CSV",
         "Possibilité d'injecter une série de dossiers validés avec création d'un journal d'audit horodaté garantissant la traçabilité.")
    ],
    img_path="img/6_succes_pastille.jpg",
    img_caption="Pastille d'intervention créée sur Géofoncier suite au versement réussi",
    callout_title="CONFIRMATION DU VERSEMENT",
    callout_text="Le code HTTP 201 Created confirme la publication instantanée de la pastille bleue sur la carte nationale de Géofoncier."
)

# Slide 23 (6.3) : Résultats Prades
print("Génération Slide 23 (Résultats Prades)...")
s23, c23 = init_standard_slide(5, 2, 23, "6.3 Validation expérimentale sur la commune de Prades (Ardèche)")

# Gauche : Détails du banc d'essai
b_left23 = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(1.60), Inches(5.65), Inches(4.30))
b_left23.fill.solid()
b_left23.fill.fore_color.rgb = C_BG_SLIDE
b_left23.line.color.rgb = C_LINE
b_left23.line.width = Pt(0.75)

bx_p23 = s23.shapes.add_textbox(Inches(1.15), Inches(1.75), Inches(5.25), Inches(4.00))
tf_p23 = bx_p23.text_frame
tf_p23.word_wrap = True
tf_p23.margin_left = tf_p23.margin_top = tf_p23.margin_right = tf_p23.margin_bottom = 0

items_prades_txt = [
    ("Campagne sur 50 dossiers réels (1970-1990)",
     "Évaluation en conditions opérationnelles sur des registres physiques de la commune de Prades conservés chez GEO-SIAPP."),
    ("Taux de succès de 94% (47 dossiers versés)",
     "47 dossiers sur 50 insérés avec succès. 3 échecs identifiés : 2 scans illisibles (< 200 DPI) et 1 doublon d'antériorité déjà présent."),
    ("Validation par le tuteur entreprise (M. Hague)",
     "Contrôle systématique de l'exactitude des parcelles, dates et noms de propriétaires avant accord définitif de versement."),
    ("Gain de productivité constaté : facteur 6",
     "Temps de traitement global pour les 50 dossiers : ~4 heures avec l'outil contre 25 heures en saisie manuelle traditionnelle.")
]

for ii, (h_p, b_p) in enumerate(items_prades_txt):
    p_h = tf_p23.paragraphs[0] if ii == 0 else tf_p23.add_paragraph()
    if ii > 0:
        p_h.space_before = Pt(8)
    r1 = p_h.add_run()
    r1.text = h_p + "\n"
    r1.font.name = "Arial"
    r1.font.size = Pt(11.0)
    r1.font.bold = True
    r1.font.color.rgb = C_NAVY
    
    p_b = tf_p23.add_paragraph()
    p_b.space_before = Pt(2)
    r2 = p_b.add_run()
    r2.text = b_p
    r2.font.name = "Arial"
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = C_MUTED

# KPIs bas gauche
kpis_prades23 = [("50", "DOSSIERS"), ("94%", "SUCCÈS"), ("F1 = 0.88", "GLOBAL"), ("x6", "RAPIDITÉ")]
kw_p23 = (5.65 - 3 * 0.10) / 4.0
ky_p23 = 6.05
for ki, (kv, kl) in enumerate(kpis_prades23):
    kx = 0.95 + ki * (kw_p23 + 0.10)
    k_shp = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(kx), Inches(ky_p23), Inches(kw_p23), Inches(0.95))
    k_shp.fill.solid()
    k_shp.fill.fore_color.rgb = C_NAVY
    k_shp.line.fill.background()
    
    bx_k = s23.shapes.add_textbox(Inches(kx + 0.04), Inches(ky_p23 + 0.12), Inches(kw_p23 - 0.08), Inches(0.70))
    tf_k = bx_k.text_frame
    tf_k.margin_left = tf_k.margin_top = tf_k.margin_right = tf_k.margin_bottom = 0
    p1 = tf_k.paragraphs[0]
    p1.text = kv
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.runs[0]
    r1.font.name = "Arial"
    r1.font.size = Pt(14.0)
    r1.font.bold = True
    r1.font.color.rgb = C_ACCENT
    p2 = tf_k.add_paragraph()
    p2.space_before = Pt(2)
    p2.text = kl
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.runs[0]
    r2.font.name = "Arial"
    r2.font.size = Pt(7.0)
    r2.font.color.rgb = C_WHITE

# Droite : Barres F1 + Image confirmation
rx23 = 6.85
rw23 = 5.50
b_right23 = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(rx23), Inches(1.60), Inches(rw23), Inches(2.95))
b_right23.fill.solid()
b_right23.fill.fore_color.rgb = C_BG_SLIDE
b_right23.line.color.rgb = C_LINE
b_right23.line.width = Pt(0.75)

bx_f1t = s23.shapes.add_textbox(Inches(rx23 + 0.20), Inches(1.72), Inches(rw23 - 0.40), Inches(0.30))
tf_f1t = bx_f1t.text_frame
p_f1t = tf_f1t.paragraphs[0]
p_f1t.text = "SCORES F1 PAR CHAMP OBLIGATOIRE"
r_f1t = p_f1t.runs[0]
r_f1t.font.name = "Arial"
r_f1t.font.size = Pt(9.5)
r_f1t.font.bold = True
r_f1t.font.color.rgb = C_ACCENT

f1_data23 = [
    ("Commune (code INSEE)", 0.94, C_SUCCESS),
    ("Date de l'opération", 0.91, C_SUCCESS),
    ("Contenance (m²)", 0.88, C_ACCENT),
    ("Nature de l'opération", 0.85, C_ACCENT),
    ("Type d'acte", 0.82, C_WARN),
    ("Score F1 global pondéré", 0.88, C_SUCCESS)
]

bar_y = 2.05
for f_name, f_val, f_col in f1_data23:
    bx_fn = s23.shapes.add_textbox(Inches(rx23 + 0.20), Inches(bar_y), Inches(2.30), Inches(0.24))
    tf_fn = bx_fn.text_frame
    p_fn = tf_fn.paragraphs[0]
    p_fn.text = f_name
    r_fn = p_fn.runs[0]
    r_fn.font.name = "Arial"
    r_fn.font.size = Pt(8.5)
    r_fn.font.color.rgb = C_DARK
    
    # Fond barre
    b_bg = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(rx23 + 2.55), Inches(bar_y + 0.03), Inches(2.00), Inches(0.15))
    b_bg.fill.solid()
    b_bg.fill.fore_color.rgb = C_BORDER_TAB
    b_bg.line.fill.background()
    
    # Remplissage barre
    b_fg = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(rx23 + 2.55), Inches(bar_y + 0.03), Inches(2.00 * f_val), Inches(0.15))
    b_fg.fill.solid()
    b_fg.fill.fore_color.rgb = f_col
    b_fg.line.fill.background()
    
    bx_fv = s23.shapes.add_textbox(Inches(rx23 + 4.65), Inches(bar_y), Inches(0.80), Inches(0.24))
    tf_fv = bx_fv.text_frame
    p_fv = tf_fv.paragraphs[0]
    p_fv.text = f"F1 = {f_val:.2f}"
    r_fv = p_fv.runs[0]
    r_fv.font.name = "Arial"
    r_fv.font.size = Pt(8.5)
    r_fv.font.bold = True
    r_fv.font.color.rgb = f_col
    
    bar_y += 0.38

# Image confirmation dessous
b_img23 = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(rx23), Inches(4.70), Inches(rw23), Inches(2.30))
b_img23.fill.solid()
b_img23.fill.fore_color.rgb = C_BG_SLIDE
b_img23.line.color.rgb = C_LINE
b_img23.line.width = Pt(0.75)
add_fitted_picture(s23, "img/6.4_match_confirm.jpg", rx23 + 0.10, 4.75, rw23 - 0.20, 2.15, caption="Fiche de correspondance validée sur Géofoncier")

# ==================== CHAPITRE 7 : LIMITES & PERSPECTIVES ====================
print("Génération Chapitre 7 (Slides 24-25)...")

# Slide 24 (7.1)
s24, c24 = init_standard_slide(6, 0, 24, "7.1 Limites techniques et contraintes opérationnelles")
add_two_column_content(
    s24,
    items=[
        ("Dépendance à la qualité de numérisation",
         "Les documents scannés à moins de 200 DPI ou fortement floutés provoquent des échecs d'alignement. Une numérisation physique soignée reste indispensable."),
        ("Écritures cursives très anciennes (pré-1960)",
         "Certaines encres artisanales fanées et graphies très serrées restent hors de portée du modèle TrOCR sans réentraînement ciblé sur le fonds."),
        ("Exigences matérielles locales pour l'inférence",
         "L'exécution combinée de YOLOv8, TrOCR et GLiNER requiert une station de travail dotée d'au moins 8 Go de VRAM GPU pour un temps d'inférence < 3s.")
    ],
    img_path="img/4513_DA_124_crop.png",
    img_caption="Document ancien illustrant les défis de déchiffrement des cursives très serrées",
    callout_title="RESPONSABILITÉ DU PROFESSIONNEL",
    callout_text="Le système est un copilote d'aide à la saisie : le géomètre-expert conserve l'entière responsabilité déontologique de la validation de chaque acte."
)

# Slide 25 (7.2)
s25, c25 = init_standard_slide(6, 1, 25, "7.2 Perspectives d'évolution et passage à l'échelle")
add_two_column_content(
    s25,
    items=[
        ("Fine-tuning de TrOCR sur les écritures SIAPP",
         "Création d'un jeu de données d'apprentissage dédié aux écritures spécifiques des géomètres fondateurs afin d'augmenter le score F1 sur manuscrit."),
        ("Généralisation et déploiement multi-cabinets",
         "Architecture modulaire et conteneurisable, facilement transposable à d'autres structures de l'Ordre des Géomètres-Experts."),
        ("Croisement avec le plan cadastral vectoriel (PCI)",
         "Intégration future du cadastre vectoriel de la DGFiP pour localiser automatiquement le centroïde parcellaire lors du versement.")
    ],
    img_path="img/repartition_archives.png",
    img_caption="Typologie des archives et potentiel d'automatisation documentaire",
    callout_title="AUTO-VALIDATION DIRECTE À TERME",
    callout_text="Pour les dossiers présentant un score F1 supérieur à 0.95 sur tous les champs, une validation en un clic pourrait être autorisée sous agrément OGE."
)

# ==================== CHAPITRE 8 : CONCLUSION ====================
print("Génération Chapitre 8 (Slides 26-27)...")

# Slide 26 (8.1) : Bilan
s26, c26 = init_standard_slide(7, 0, 26, "8.1 Bilan général du Projet de Fin d'Études")

bilan_data = [
    ("OBJECTIFS TECHNIQUES", [
        ("Chaîne 100% locale souveraine", "Traitement intégral sur PC de bureau sans aucun recours au cloud."),
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

cw26 = (11.93 - 0.40 - 2 * 0.15) / 3.0
for ci, (col_title, col_items) in enumerate(bilan_data):
    cx = 0.70 + 0.20 + ci * (cw26 + 0.15)
    cy = 1.60
    ch = 4.30
    
    b_col = s26.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(cy), Inches(cw26), Inches(ch))
    b_col.fill.solid()
    b_col.fill.fore_color.rgb = C_BG_SLIDE
    b_col.line.color.rgb = C_LINE
    b_col.line.width = Pt(0.75)
    
    h_col = s26.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(cy), Inches(cw26), Inches(0.55))
    h_col.fill.solid()
    h_col.fill.fore_color.rgb = C_NAVY
    h_col.line.fill.background()
    
    bx_ct = s26.shapes.add_textbox(Inches(cx + 0.05), Inches(cy + 0.08), Inches(cw26 - 0.10), Inches(0.40))
    tf_ct = bx_ct.text_frame
    p_ct = tf_ct.paragraphs[0]
    p_ct.alignment = PP_ALIGN.CENTER
    p_ct.text = col_title
    r_ct = p_ct.runs[0]
    r_ct.font.name = "Arial"
    r_ct.font.size = Pt(10.0)
    r_ct.font.bold = True
    r_ct.font.color.rgb = C_WHITE
    
    bx_cc = s26.shapes.add_textbox(Inches(cx + 0.15), Inches(cy + 0.70), Inches(cw26 - 0.30), Inches(3.45))
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
kpis_b26 = [("F1 = 0.88", "SCORE GLOBAL"), ("94%", "SUCCÈS PRADES"), ("x6", "GAIN TEMPS"), ("23 600", "DOSSIERS EN COURS")]
kw26_k = (11.53 - 3 * 0.12) / 4.0
for ki, (kv, kl) in enumerate(kpis_b26):
    kx = 0.90 + ki * (kw26_k + 0.12)
    k_box = s26.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(kx), Inches(6.05), Inches(kw26_k), Inches(0.95))
    k_box.fill.solid()
    k_box.fill.fore_color.rgb = C_NAVY
    k_box.line.fill.background()
    
    bx_k = s26.shapes.add_textbox(Inches(kx + 0.05), Inches(6.15), Inches(kw26_k - 0.10), Inches(0.75))
    tf_k = bx_k.text_frame
    p1 = tf_k.paragraphs[0]
    p1.text = kv
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.runs[0]
    r1.font.name = "Arial"
    r1.font.size = Pt(15.0)
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

# Slide 27 (8.2) : Remerciements
print("Génération Slide 27 (Remerciements 180s)...")
s27, c27 = init_standard_slide(7, 1, 27, "8.2 Remerciements et ouverture vers les échanges")

# Carte centrale élégante
bx_rm = s27.shapes.add_textbox(Inches(1.50), Inches(1.80), Inches(10.33), Inches(0.70))
tf_rm = bx_rm.text_frame
p_rm = tf_rm.paragraphs[0]
p_rm.alignment = PP_ALIGN.CENTER
p_rm.text = "Merci pour votre attention"
r_rm = p_rm.runs[0]
r_rm.font.name = "Arial"
r_rm.font.size = Pt(32.0)
r_rm.font.bold = True
r_rm.font.color.rgb = C_NAVY

bx_rq = s27.shapes.add_textbox(Inches(1.50), Inches(2.55), Inches(10.33), Inches(0.40))
tf_rq = bx_rq.text_frame
p_rq = tf_rq.paragraphs[0]
p_rq.alignment = PP_ALIGN.CENTER
p_rq.text = "Je me tiens à votre disposition pour vos questions."
r_rq = p_rq.runs[0]
r_rq.font.name = "Arial"
r_rq.font.size = Pt(14.0)
r_rq.font.color.rgb = C_ACCENT

sep27 = s27.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.66), Inches(3.15), Inches(4.00), Inches(0.025))
sep27.fill.solid()
sep27.fill.fore_color.rgb = C_LINE
sep27.line.fill.background()

thanks_lines = [
    "M. Gaëtan HAGUE : Tuteur entreprise, Géomètre-Expert associé, Cabinet GEO-SIAPP",
    "M. Mathieu KOEHL : Directeur du PFE, Enseignant-chercheur, INSA Strasbourg / ICube",
    "L'équipe du Cabinet GEO-SIAPP d'Aubenas pour leur accueil et leur accompagnement technique",
    "L'équipe pédagogique de la spécialité Topographie de l'INSA Strasbourg"
]

for ti, ttext in enumerate(thanks_lines):
    bx_th = s27.shapes.add_textbox(Inches(1.50), Inches(3.45 + ti * 0.45), Inches(10.33), Inches(0.35))
    tf_th = bx_th.text_frame
    p_th = tf_th.paragraphs[0]
    p_th.alignment = PP_ALIGN.CENTER
    p_th.text = ttext
    r_th = p_th.runs[0]
    r_th.font.name = "Arial"
    r_th.font.size = Pt(11.0)
    r_th.font.color.rgb = C_DARK

# Logos centrés en bas de la carte
if os.path.exists("Logo_INSAStrasbourg.jpg"):
    s27.shapes.add_picture("Logo_INSAStrasbourg.jpg", Inches(3.80), Inches(5.55), width=Inches(2.50))
if os.path.exists("Geosiapp.jpg"):
    add_fitted_picture(s27, "Geosiapp.jpg", 7.60, 5.40, 2.00, 1.20)

# ==================== ENREGISTREMENT ====================
output_pptx = "Soutenance_PFE_Adrien_TRAVAILLE.pptx"
prs.save(output_pptx)
print(f"\n=======================================================")
print(f"SUCCÈS : Présentation 180s enregistrée sous '{output_pptx}'")
print(f"Nombre total de diapositives : {len(prs.slides)}")
print(f"=======================================================\n")
