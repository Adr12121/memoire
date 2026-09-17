# -*- coding: utf-8 -*-
"""
create_soutenance_sovereign.py
Générateur complet de la soutenance de PFE (27 slides)
- Design inspiré de la présentation PFE 180s (en-tête à onglets épuré, fond blanc, grand conteneur arrondi)
- Contenu fidèle au mémoire d'ingénieur (chapitres 1 à 8)
- Découpage sobre : 1 slide par idée avec visuel mis en valeur
- Pas de pavés de texte, pas de couleurs criardes, grand confort de lecture pour le jury
"""

import os
import shutil
from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

OUTPUT_PATH = "Soutenance_PFE_Adrien_TRAVAILLE.pptx"
BACKUP_PATH = "Soutenance_PFE_backup.pptx"

if os.path.exists(OUTPUT_PATH) and not os.path.exists(BACKUP_PATH):
    shutil.copy(OUTPUT_PATH, BACKUP_PATH)
    print(f"Backup créé : {BACKUP_PATH}")

prs = Presentation()
prs.slide_width = Cm(33.87)
prs.slide_height = Cm(19.05)
blank_layout = prs.slide_layouts[6] # layout 6 = blank

# Charte graphique sobre
C_NAVY   = RGBColor(16,  44,  87)   # #102C57 - Bleu Nuit Institutionnel
C_BLUE   = RGBColor(2,   132, 199)  # #0284C7 - Bleu accent soutenance
C_DARK   = RGBColor(30,  41,  59)   # #1E293B - Texte principal haute lisibilité
C_MUTED  = RGBColor(100, 116, 139)  # #64748B - Texte secondaire / légendes
C_BG_TAB = RGBColor(241, 245, 249)  # #F1F5F9 - Fond onglet inactif
C_CARD   = RGBColor(255, 255, 255)  # Fond blanc pur
C_BORDER = RGBColor(226, 232, 240)  # #E2E8F0 - Bordure très discrète
C_LINE   = RGBColor(203, 213, 225)  # #CBD5E1 - Ligne séparatrice
C_GREEN  = RGBColor(22,  163, 74)   # #16A34A - Succès / validation
C_WHITE  = RGBColor(255, 255, 255)

TABS = [
    "1. Contexte & Enjeux",
    "2. État de l'art",
    "3. Chaîne d'extraction IA",
    "4. Fiabilisation & Géofoncier",
    "5. Bilan & Perspectives"
]

def add_header_2tier(slide, active_main_idx, active_sub_idx, sub_list, slide_num, total_slides=29):
    """
    Crée un en-tête à double frise de navigation :
    - Rangée 1 : Logos et 5 grands chapitres du mémoire (l'actif en Bleu Nuit #102C57)
    - Rangée 2 : Ruban des sous-parties du chapitre actif (la sous-partie active en Bleu Accent #0284C7)
    """
    # Fond d'en-tête blanc épuré
    bg_hdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(0), Cm(0), Cm(33.87), Cm(2.00))
    bg_hdr.fill.solid()
    bg_hdr.fill.fore_color.rgb = C_WHITE
    bg_hdr.line.color.rgb = C_LINE
    bg_hdr.line.width = Pt(0.75)
    
    # Logo INSA (gauche)
    if os.path.exists("Logo_INSAStrasbourg.jpg"):
        slide.shapes.add_picture("Logo_INSAStrasbourg.jpg", Cm(1.2), Cm(0.14), width=Cm(3.8))
        
    # Logo GEO-SIAPP (droite)
    if os.path.exists("Geosiapp.jpg"):
        slide.shapes.add_picture("Geosiapp.jpg", Cm(31.2), Cm(0.14), height=Cm(0.85))
        
    # RANGÉE 1 : 5 Grands Chapitres
    tab_w = Cm(4.85)
    tab_h = Cm(0.80)
    tab_gap = Cm(0.20)
    start_x = Cm(5.4)
    y_pos = Cm(0.14)
    
    for i, tab_name in enumerate(TABS):
        tx = start_x + i * (tab_w + tab_gap)
        is_active = (i == active_main_idx)
        
        tab_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, tx, y_pos, tab_w, tab_h)
        tab_shape.fill.solid()
        tab_shape.fill.fore_color.rgb = C_NAVY if is_active else C_BG_TAB
        tab_shape.line.color.rgb = C_NAVY if is_active else C_BORDER
        tab_shape.line.width = Pt(0.5)
        
        tf = tab_shape.text_frame
        tf.word_wrap = False
        tf.margin_left = Cm(0.06)
        tf.margin_right = Cm(0.06)
        tf.margin_top = Cm(0.05)
        tf.margin_bottom = Cm(0.05)
        p = tf.paragraphs[0]
        p.text = ("● " if is_active else "") + tab_name
        p.alignment = PP_ALIGN.CENTER
        p.runs[0].font.name = "Arial"
        p.runs[0].font.size = Pt(9.5)
        p.runs[0].font.bold = is_active
        p.runs[0].font.color.rgb = C_WHITE if is_active else C_MUTED

    # RANGÉE 2 : Frise des Sous-parties du chapitre actif
    if sub_list and len(sub_list) > 0:
        n_sub = len(sub_list)
        sub_gap = Cm(0.20)
        max_sub_w = Cm(5.6)
        calc_sub_w = min(max_sub_w, (Cm(26.5) - (n_sub - 1) * sub_gap) / n_sub)
        total_sub_w = n_sub * calc_sub_w + (n_sub - 1) * sub_gap
        sub_start_x = (Cm(33.87) - total_sub_w) / 2
        sub_y = Cm(1.06)
        sub_h = Cm(0.74)
        
        for j, sub_name in enumerate(sub_list):
            is_sub_active = (j == active_sub_idx)
            sx = sub_start_x + j * (calc_sub_w + sub_gap)
            
            sub_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sx, sub_y, calc_sub_w, sub_h)
            sub_shape.fill.solid()
            sub_shape.fill.fore_color.rgb = C_BLUE if is_sub_active else RGBColor(248, 250, 252)
            sub_shape.line.color.rgb = C_BLUE if is_sub_active else C_BORDER
            sub_shape.line.width = Pt(0.75 if is_sub_active else 0.5)
            
            tf_s = sub_shape.text_frame
            tf_s.word_wrap = False
            tf_s.margin_left = Cm(0.06)
            tf_s.margin_right = Cm(0.06)
            tf_s.margin_top = Cm(0.05)
            tf_s.margin_bottom = Cm(0.05)
            p_s = tf_s.paragraphs[0]
            p_s.text = ("▸ " if is_sub_active else "") + sub_name
            p_s.alignment = PP_ALIGN.CENTER
            p_s.runs[0].font.name = "Arial"
            p_s.runs[0].font.size = Pt(8.5)
            p_s.runs[0].font.bold = is_sub_active
            p_s.runs[0].font.color.rgb = C_WHITE if is_sub_active else RGBColor(71, 85, 105)

    # Numéro de slide en bas à droite
    tx_num = slide.shapes.add_textbox(Cm(30.0), Cm(18.25), Cm(3.2), Cm(0.6))
    p_num = tx_num.text_frame.paragraphs[0]
    p_num.text = f"{slide_num} / {total_slides}"
    p_num.alignment = PP_ALIGN.RIGHT
    p_num.runs[0].font.name = "Arial"
    p_num.runs[0].font.size = Pt(10)
    p_num.runs[0].font.color.rgb = C_MUTED

def add_header(slide, active_tab_idx, slide_num, total_slides=29):
    """Compatibilité : redirige vers add_header_2tier avec sous-liste vide"""
    add_header_2tier(slide, active_tab_idx, -1, [], slide_num, total_slides)

def add_title(slide, title_text, category_subtitle=None):
    tx_box = slide.shapes.add_textbox(Cm(1.80), Cm(2.08), Cm(30.27), Cm(1.28))
    tf = tx_box.text_frame
    tf.word_wrap = True
    tf.margin_left = Cm(0)
    tf.margin_top = Cm(0)
    tf.margin_bottom = Cm(0)
    
    p = tf.paragraphs[0]
    p.text = title_text
    p.runs[0].font.name = "Arial"
    p.runs[0].font.size = Pt(19.0)
    p.runs[0].font.bold = True
    p.runs[0].font.color.rgb = C_NAVY
    
    if category_subtitle:
        p_sub = tf.add_paragraph()
        p_sub.space_before = Pt(2)
        p_sub.text = category_subtitle
        p_sub.runs[0].font.name = "Arial"
        p_sub.runs[0].font.size = Pt(11.0)
        p_sub.runs[0].font.bold = True
        p_sub.runs[0].font.color.rgb = C_BLUE

def create_card(slide):
    # Grand conteneur blanc arrondi calibré pour ne jamais chevaucher le titre
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(1.80), Cm(3.45), Cm(30.27), Cm(14.65))
    card.fill.solid()
    card.fill.fore_color.rgb = C_WHITE
    card.line.color.rgb = C_BORDER
    card.line.width = Pt(1.0)
    return card

def init_slide_layout(prs, active_main_idx, active_sub_idx, sub_list, slide_num, title, subtitle=None, total_slides=29):
    """
    Initialise une slide avec :
    1. Fond gris très clair #F8FAFC pour un contraste souverain
    2. Grand conteneur blanc arrondi #FFFFFF (y=3.45 à 18.10)
    3. Double frise de navigation (y=0 à 2.00)
    4. Titre et sous-titre (y=2.08 à 3.36) sans AUCUN chevauchement
    """
    slide = prs.slides.add_slide(blank_layout)
    
    # Fond général gris très doux
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(0), Cm(0), Cm(33.87), Cm(19.05))
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(248, 250, 252) # #F8FAFC
    bg.line.fill.background()
    
    # Grand conteneur blanc
    card = create_card(slide)
    
    # En-tête 2 niveaux
    add_header_2tier(slide, active_main_idx, active_sub_idx, sub_list, slide_num, total_slides)
    
    # Titre & sous-titre
    add_title(slide, title, subtitle)
    
    return slide, card


def add_bullets(slide, x, y, w, h, bullet_items):
    """bullet_items = list of (title_bold, text_normal) or string"""
    tx_box = slide.shapes.add_textbox(x, y, w, h)
    tf = tx_box.text_frame
    tf.word_wrap = True
    tf.margin_left = Cm(0.2)
    tf.margin_right = Cm(0.2)
    tf.margin_top = Cm(0.1)
    tf.margin_bottom = Cm(0.1)
    
    count = len(bullet_items)
    if count <= 3:
        font_sz = Pt(16.5)
        sp_after = Pt(16)
        sp_before = Pt(4)
    elif count == 4:
        font_sz = Pt(15.0)
        sp_after = Pt(10)
        sp_before = Pt(2)
    else:
        font_sz = Pt(14.0)
        sp_after = Pt(6)
        sp_before = Pt(2)
    
    for idx, item in enumerate(bullet_items):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.space_after = sp_after
        p.space_before = sp_before
        
        if isinstance(item, tuple):
            title_b, text_n = item
            r1 = p.add_run()
            r1.text = "• " + title_b + " : "
            r1.font.name = "Arial"
            r1.font.size = font_sz
            r1.font.bold = True
            r1.font.color.rgb = C_NAVY
            
            r2 = p.add_run()
            r2.text = text_n
            r2.font.name = "Arial"
            r2.font.size = font_sz
            r2.font.bold = False
            r2.font.color.rgb = C_DARK
        else:
            r = p.add_run()
            r.text = "• " + item
            r.font.name = "Arial"
            r.font.size = font_sz
            r.font.color.rgb = C_DARK

def add_fitted_picture(slide, img_path, box_x, box_y, box_w, box_h, caption_text=None):
    """
    Intelligently scales and centers an image within (box_x, box_y, box_w, box_h).
    Leaves space at the bottom for caption_text if provided, and places the caption
    right below the image, centered horizontally.
    """
    from PIL import Image
    if not os.path.exists(img_path):
        print(f"Warning: {img_path} not found.")
        return None
        
    cap_h = Cm(0.7) if caption_text else Cm(0)
    gap = Cm(0.2) if caption_text else Cm(0)
    avail_w = box_w
    avail_h = box_h - (cap_h + gap)
    
    with Image.open(img_path) as im:
        orig_w, orig_h = im.size
    
    img_ratio = orig_w / orig_h
    box_ratio = avail_w / avail_h
    
    if img_ratio > box_ratio:
        final_w = avail_w
        final_h = avail_w / img_ratio
    else:
        final_h = avail_h
        final_w = avail_h * img_ratio
        
    img_x = box_x + (box_w - final_w) / 2
    img_y = box_y + (avail_h - final_h) / 2
    
    pic = slide.shapes.add_picture(img_path, img_x, img_y, width=final_w, height=final_h)
    
    if caption_text:
        cap_y = img_y + final_h + gap
        tx = slide.shapes.add_textbox(box_x, cap_y, box_w, cap_h)
        tf = tx.text_frame
        tf.word_wrap = True
        tf.margin_top = Cm(0)
        tf.margin_bottom = Cm(0)
        tf.margin_left = Cm(0)
        tf.margin_right = Cm(0)
        p = tf.paragraphs[0]
        p.text = caption_text
        p.alignment = PP_ALIGN.CENTER
        p.runs[0].font.name = "Arial"
        p.runs[0].font.size = Pt(11)
        p.runs[0].font.color.rgb = C_MUTED
        p.runs[0].font.italic = True
        
    return pic

def add_card_box(slide, x, y, w, h, title, items, bg_color=C_BG_TAB, title_color=C_NAVY, title_size=15, item_size=12):
    """
    Creates a rounded rectangle card with a title and properly styled bullet items,
    ensuring NO text lines lose their font color.
    items: list of (bold_prefix, text) or string
    """
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = C_BORDER
    card.line.width = Pt(0.75)
    
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Cm(0.6)
    tf.margin_right = Cm(0.6)
    tf.margin_top = Cm(0.5)
    tf.margin_bottom = Cm(0.4)
    
    p = tf.paragraphs[0]
    p.text = title
    p.runs[0].font.name = "Arial"
    p.runs[0].font.size = Pt(title_size)
    p.runs[0].font.bold = True
    p.runs[0].font.color.rgb = title_color
    
    for item in items:
        p_item = tf.add_paragraph()
        p_item.space_before = Pt(4)
        if isinstance(item, tuple):
            prefix_b, text_n = item
            r1 = p_item.add_run()
            r1.text = prefix_b + " : "
            r1.font.name = "Arial"
            r1.font.size = Pt(item_size)
            r1.font.bold = True
            r1.font.color.rgb = C_BLUE
            
            r2 = p_item.add_run()
            r2.text = text_n
            r2.font.name = "Arial"
            r2.font.size = Pt(item_size)
            r2.font.bold = False
            r2.font.color.rgb = C_DARK
        else:
            r = p_item.add_run()
            r.text = item
            r.font.name = "Arial"
            r.font.size = Pt(item_size)
            r.font.color.rgb = C_DARK
            
    return card

def add_kpi_box(slide, x, y, w, h, number_str, label_str, subtext_str=None, num_color=C_BLUE, bg_color=RGBColor(248, 250, 252)):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = C_BORDER
    card.line.width = Pt(0.75)
    
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Cm(0.3)
    tf.margin_right = Cm(0.3)
    tf.margin_top = Cm(0.3)
    tf.margin_bottom = Cm(0.2)
    
    p0 = tf.paragraphs[0]
    p0.text = number_str
    p0.alignment = PP_ALIGN.CENTER
    p0.runs[0].font.name = "Arial"
    p0.runs[0].font.size = Pt(26)
    p0.runs[0].font.bold = True
    p0.runs[0].font.color.rgb = num_color
    
    p1 = tf.add_paragraph()
    p1.space_before = Pt(2)
    p1.text = label_str
    p1.alignment = PP_ALIGN.CENTER
    p1.runs[0].font.name = "Arial"
    p1.runs[0].font.size = Pt(11)
    p1.runs[0].font.bold = True
    p1.runs[0].font.color.rgb = C_NAVY
    
    if subtext_str:
        p2 = tf.add_paragraph()
        p2.space_before = Pt(2)
        p2.text = subtext_str
        p2.alignment = PP_ALIGN.CENTER
        p2.runs[0].font.name = "Arial"
        p2.runs[0].font.size = Pt(9.5)
        p2.runs[0].font.color.rgb = C_MUTED
    return card

def add_callout_box(slide, x, y, w, h, bold_title, text_body, bg_color=RGBColor(240, 249, 255), border_color=C_BLUE):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(1.0)
    
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Cm(0.5)
    tf.margin_right = Cm(0.5)
    tf.margin_top = Cm(0.35)
    tf.margin_bottom = Cm(0.3)
    
    p0 = tf.paragraphs[0]
    p0.text = bold_title
    p0.runs[0].font.name = "Arial"
    p0.runs[0].font.size = Pt(12)
    p0.runs[0].font.bold = True
    p0.runs[0].font.color.rgb = border_color
    
    p1 = tf.add_paragraph()
    p1.space_before = Pt(2)
    p1.text = text_body
    p1.runs[0].font.name = "Arial"
    p1.runs[0].font.size = Pt(10.5)
    p1.runs[0].font.color.rgb = C_DARK
    return card

print("Initialisation du générateur enrichi...")


