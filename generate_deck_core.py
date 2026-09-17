# -*- coding: utf-8 -*-
"""
Primitives graphiques communes — Soutenance PFE
Adrien TRAVAILLÉ — Diplôme d'Ingénieur Topographe INSA Strasbourg
Cabinet GEO-SIAPP (Aubenas, Ardèche) — Soutenance septembre 2026
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# =============================================================================
# CHARTE GRAPHIQUE PROFESSIONNELLE
# =============================================================================
C_NAVY          = RGBColor(15, 37, 70)       # #0F2546 — bleu nuit
C_NAVY_MID      = RGBColor(25, 60, 110)      # #193C6E — bleu mid
C_NAVY_LIGHT    = RGBColor(241, 245, 249)    # #F1F5F9 — fond slides contenu
C_ACCENT        = RGBColor(2, 132, 199)      # #0284C7 — bleu topo
C_ACCENT_LIGHT  = RGBColor(224, 242, 254)    # #E0F2FE — bleu ciel pastel
C_SUCCESS       = RGBColor(22, 163, 74)      # #16A34A — vert validation
C_SUCCESS_BG    = RGBColor(240, 253, 244)    # #F0FDF4 — vert pastel
C_WARN          = RGBColor(217, 119, 6)      # #D97706 — orange alerte
C_WARN_BG       = RGBColor(255, 251, 235)    # #FFFBEB — fond orange doux
C_TEXT_DARK     = RGBColor(15, 23, 42)       # #0F172A — titres
C_TEXT_MUTED    = RGBColor(71, 85, 105)      # #475569 — sous-titres
C_TEXT_BODY     = RGBColor(51, 65, 85)       # #334155 — texte courant
C_BORDER        = RGBColor(226, 232, 240)    # #E2E8F0 — bordures
C_WHITE         = RGBColor(255, 255, 255)    # #FFFFFF — blanc
C_CALLOUT_BG    = RGBColor(238, 242, 255)    # #EEF2FF — encadré bas

# 8 chapitres du mémoire (étiquettes courtes pour la frise)
CHAPTERS = [
    "1. INTRO",
    "2. MÉTIER",
    "3. ÉTAT ART",
    "4. ARCHI",
    "5. VALID.",
    "6. INTÉGR.",
    "7. LIMITES",
    "8. CONCL.",
]

# Couleurs d'accent par chapitre (pour les tuiles du sommaire et les transitions)
CHAPTER_COLORS = [
    RGBColor(15, 37, 70),    # 1 — bleu nuit
    RGBColor(2, 132, 199),   # 2 — bleu topo
    RGBColor(88, 28, 135),   # 3 — violet
    RGBColor(6, 95, 70),     # 4 — vert sombre
    RGBColor(154, 52, 18),   # 5 — orange brûlé
    RGBColor(15, 118, 110),  # 6 — teal
    RGBColor(109, 40, 217),  # 7 — violet clair
    RGBColor(15, 37, 70),    # 8 — bleu nuit (conclusion)
]

SUB_PARTS = {
    0: ["1.1 Structure d'accueil", "1.2 Cadre légal", "1.3 Contexte & Besoins"],
    1: ["2.1 Cadastre & Propriété", "2.2 API Géofoncier", "2.3 Fonds SIAPP", "2.4 Registres & Filiation"],
    2: ["3.1 OCR & HTR", "3.2 GLiNER & LayoutLM", "3.3 VLM & Tableaux", "3.4 Synthèse"],
    3: ["4.1 Architecture", "4.2 Prétraitement", "4.3 YOLOv8 & GLiNER", "4.4 Arbitrage VLM"],
    4: ["5.1 Streamlit", "5.2 Vue miroir & Folium", "5.3 Démonstration vidéo"],
    5: ["6.1 API REST & JWT", "6.2 Versement par lot", "6.3 Cas de Prades", "6.4 Éval. F1"],
    6: ["7.1 Limites physiques", "7.2 Perspectives"],
    7: ["8.1 Bilan", "8.2 Déontologie & OGE"],
    8: ["A. Bibliographie", "B. Dépôts open-source"],
    9: ["A. GLiNER", "B. YOLOv8", "C. Pipeline OCR", "D. VLM/Ollama"],
}


# =============================================================================
# FOND
# =============================================================================

def apply_background(slide):
    """Fond clair pour les slides de contenu."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = C_NAVY_LIGHT


def apply_dark_background(slide, color=None):
    """Fond sombre pour slides de titre, transition et clôture."""
    if color is None:
        color = C_NAVY
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


# =============================================================================
# FRISE DE NAVIGATION (2 NIVEAUX)
# =============================================================================

def add_navigation_bars(slide, active_chap_idx=0, active_sub_idx=0):
    """
    Frise supérieure sur 2 niveaux :
    Ligne 1 : 8 chapitres (ou bandeau spécial pour Références/Annexes)
    Ligne 2 : sous-parties du chapitre actif
    """
    # --- Cas spéciaux : Références (idx 8) et Annexes (idx 9) ---
    if active_chap_idx in (8, 9):
        label = ("RÉFÉRENCES BIBLIOGRAPHIQUES & SOURCES" if active_chap_idx == 8
                 else "ANNEXES TECHNIQUES — FICHES JURY")
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                     Inches(0.30), Inches(0.06),
                                     Inches(12.733), Inches(0.40))
        bar.fill.solid()
        bar.fill.fore_color.rgb = C_NAVY
        bar.line.fill.background()
        tf = bar.text_frame
        tf.word_wrap = False
        tf.margin_top = Inches(0.07)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = label
        p.font.size = Pt(9.0)
        p.font.bold = True
        p.font.color.rgb = C_WHITE

        # Sous-parties
        sub_list = SUB_PARTS.get(active_chap_idx, [])
        _draw_subparts(slide, sub_list, active_sub_idx, y_top=0.50)
        return

    # --- Cas standard : 8 chapitres ---
    n = len(CHAPTERS)
    total_w = 12.733
    gap = 0.04
    chap_w = (total_w - (n - 1) * gap) / n

    for i, chap_name in enumerate(CHAPTERS):
        cx = 0.30 + i * (chap_w + gap)
        is_active = (i == active_chap_idx)
        color = CHAPTER_COLORS[i] if is_active else C_WHITE
        rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                      Inches(cx), Inches(0.06),
                                      Inches(chap_w), Inches(0.40))
        rect.fill.solid()
        rect.fill.fore_color.rgb = color
        rect.line.color.rgb = CHAPTER_COLORS[i] if is_active else C_BORDER
        rect.line.width = Pt(1.5 if is_active else 0.6)
        tf = rect.text_frame
        tf.word_wrap = False
        tf.margin_top = Inches(0.09)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = chap_name
        p.font.size = Pt(7.0)
        p.font.bold = is_active
        p.font.color.rgb = C_WHITE if is_active else C_TEXT_MUTED

    # Sous-parties
    sub_list = SUB_PARTS.get(active_chap_idx, [])
    _draw_subparts(slide, sub_list, active_sub_idx, y_top=0.50)


def _draw_subparts(slide, sub_list, active_sub_idx, y_top=0.50):
    n_sub = len(sub_list)
    if n_sub == 0:
        return
    total_w = 12.733
    gap = 0.06
    sub_w = (total_w - (n_sub - 1) * gap) / n_sub
    for i, sub_name in enumerate(sub_list):
        sx = 0.30 + i * (sub_w + gap)
        is_active = (i == active_sub_idx)
        sub_rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                          Inches(sx), Inches(y_top),
                                          Inches(sub_w), Inches(0.28))
        sub_rect.fill.solid()
        sub_rect.fill.fore_color.rgb = C_ACCENT if is_active else C_WHITE
        sub_rect.line.color.rgb = C_ACCENT if is_active else C_BORDER
        sub_rect.line.width = Pt(1.2 if is_active else 0.6)
        stf = sub_rect.text_frame
        stf.word_wrap = False
        stf.margin_top = Inches(0.04)
        sp = stf.paragraphs[0]
        sp.alignment = PP_ALIGN.CENTER
        sp.text = sub_name
        sp.font.size = Pt(7.5)
        sp.font.bold = is_active
        sp.font.color.rgb = C_WHITE if is_active else C_TEXT_MUTED


# =============================================================================
# EN-TÊTE DE SLIDE (CATÉGORIE + TITRE + SOUS-TITRE OPTIONNEL)
# =============================================================================

def add_slide_header(slide, category_str, main_title, subtitle_str=""):
    """En-tête de slide : catégorie bleue, titre sombre, sous-titre grisé."""
    # Catégorie
    cat_box = slide.shapes.add_textbox(Inches(0.30), Inches(0.84), Inches(12.733), Inches(0.22))
    ctf = cat_box.text_frame
    ctf.word_wrap = True
    ctf.margin_left = ctf.margin_top = ctf.margin_right = ctf.margin_bottom = 0
    cp = ctf.paragraphs[0]
    cp.text = category_str.upper()
    cp.font.size = Pt(8.0)
    cp.font.bold = True
    cp.font.color.rgb = C_ACCENT

    # Titre principal
    title_box = slide.shapes.add_textbox(Inches(0.30), Inches(1.02), Inches(12.733), Inches(0.42))
    ttf = title_box.text_frame
    ttf.word_wrap = True
    ttf.margin_left = ttf.margin_top = ttf.margin_right = ttf.margin_bottom = 0
    tp = ttf.paragraphs[0]
    tp.text = main_title
    tp.font.size = Pt(17.0)
    tp.font.bold = True
    tp.font.color.rgb = C_TEXT_DARK

    # Sous-titre optionnel
    if subtitle_str:
        sub_box = slide.shapes.add_textbox(Inches(0.30), Inches(1.42), Inches(12.733), Inches(0.22))
        stf = sub_box.text_frame
        stf.word_wrap = True
        stf.margin_left = stf.margin_top = stf.margin_right = stf.margin_bottom = 0
        sp = stf.paragraphs[0]
        sp.text = subtitle_str
        sp.font.size = Pt(9.5)
        sp.font.color.rgb = C_TEXT_MUTED

    # Séparateur fin
    sep = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                 Inches(0.30), Inches(1.64), Inches(12.733), Inches(0.018))
    sep.fill.solid()
    sep.fill.fore_color.rgb = C_BORDER
    sep.line.fill.background()


# =============================================================================
# PIED DE PAGE MINIMALISTE
# =============================================================================

def add_footer(slide, current_slide, total_slides=35):
    ftr = slide.shapes.add_textbox(Inches(11.50), Inches(7.18), Inches(1.533), Inches(0.24))
    rtf = ftr.text_frame
    rtf.margin_left = rtf.margin_top = rtf.margin_right = rtf.margin_bottom = 0
    rp = rtf.paragraphs[0]
    rp.alignment = PP_ALIGN.RIGHT
    rp.text = f"{current_slide} / {total_slides}"
    rp.font.size = Pt(8.5)
    rp.font.bold = False
    rp.font.color.rgb = C_TEXT_MUTED


# =============================================================================
# SLIDE DE TRANSITION (FOND BLEU PLEIN, TITRE CHAPITRE EN GRAND)
# =============================================================================

def add_transition_slide(prs, chap_num, chap_title, chap_subtitle="", slide_num=0,
                         total_slides=35, chap_idx=0):
    """
    Slide de transition entre chapitres :
    fond bleu nuit plein, numéro de chapitre petit en haut,
    titre grand en blanc centré, sous-titre grisé.
    """
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    color = CHAPTER_COLORS[chap_idx % len(CHAPTER_COLORS)]
    apply_dark_background(slide, color)

    # Pastille numéro de chapitre
    pill_w = 1.60
    pill_h = 0.36
    pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Inches((13.333 - pill_w) / 2),
                                  Inches(2.10),
                                  Inches(pill_w), Inches(pill_h))
    pill.fill.solid()
    pill.fill.fore_color.rgb = C_WHITE
    pill.line.fill.background()
    ptf = pill.text_frame
    ptf.margin_top = Inches(0.06)
    pp = ptf.paragraphs[0]
    pp.alignment = PP_ALIGN.CENTER
    pp.text = f"CHAPITRE {chap_num}"
    pp.font.size = Pt(9.5)
    pp.font.bold = True
    pp.font.color.rgb = color

    # Titre
    tbox = slide.shapes.add_textbox(Inches(1.0), Inches(2.58), Inches(11.333), Inches(1.20))
    ttf = tbox.text_frame
    ttf.word_wrap = True
    ttf.margin_left = ttf.margin_top = ttf.margin_right = ttf.margin_bottom = 0
    tp = ttf.paragraphs[0]
    tp.text = chap_title
    tp.alignment = PP_ALIGN.CENTER
    tp.font.size = Pt(34.0)
    tp.font.bold = True
    tp.font.color.rgb = C_WHITE

    # Sous-titre
    if chap_subtitle:
        stbox = slide.shapes.add_textbox(Inches(1.5), Inches(3.88), Inches(10.333), Inches(0.50))
        stf = stbox.text_frame
        stf.word_wrap = True
        stf.margin_left = stf.margin_top = stf.margin_right = stf.margin_bottom = 0
        sp = stf.paragraphs[0]
        sp.text = chap_subtitle
        sp.alignment = PP_ALIGN.CENTER
        sp.font.size = Pt(13.0)
        sp.font.color.rgb = RGBColor(180, 210, 240)

    # Ligne décorative
    deco = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                  Inches(5.2), Inches(4.55), Inches(2.933), Inches(0.04))
    deco.fill.solid()
    deco.fill.fore_color.rgb = C_ACCENT
    deco.line.fill.background()

    # Numéro de slide
    if slide_num > 0:
        ftr = slide.shapes.add_textbox(Inches(11.50), Inches(7.18), Inches(1.533), Inches(0.24))
        rtf = ftr.text_frame
        rtf.margin_left = rtf.margin_top = rtf.margin_right = rtf.margin_bottom = 0
        rp = rtf.paragraphs[0]
        rp.alignment = PP_ALIGN.RIGHT
        rp.text = f"{slide_num} / {total_slides}"
        rp.font.size = Pt(8.5)
        rp.font.color.rgb = RGBColor(120, 150, 190)

    return slide


# =============================================================================
# CARTE DE CONTENU (BANDEAU + POINTS)
# =============================================================================

def draw_styled_card(slide, left, top, width, height, title, items,
                     banner_color=None, tag_str="", bottom_callout=""):
    if banner_color is None:
        banner_color = C_NAVY

    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = C_WHITE
    card.line.color.rgb = C_BORDER
    card.line.width = Pt(1.0)

    # Bandeau titre
    banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    Inches(left), Inches(top), Inches(width), Inches(0.36))
    banner.fill.solid()
    banner.fill.fore_color.rgb = banner_color
    banner.line.fill.background()
    btf = banner.text_frame
    btf.margin_left = Inches(0.14)
    btf.margin_top = Inches(0.06)
    bp = btf.paragraphs[0]
    bp.text = title
    bp.font.size = Pt(10.0)
    bp.font.bold = True
    bp.font.color.rgb = C_WHITE

    # Badge optionnel
    if tag_str:
        tag_w = max(1.0, min(2.0, len(tag_str) * 0.10 + 0.30))
        tag = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                     Inches(left + width - tag_w - 0.10),
                                     Inches(top + 0.05),
                                     Inches(tag_w), Inches(0.26))
        tag.fill.solid()
        tag.fill.fore_color.rgb = C_WHITE
        tag.line.fill.background()
        ttf2 = tag.text_frame
        ttf2.margin_top = Inches(0.03)
        ttf2.margin_left = ttf2.margin_right = ttf2.margin_bottom = 0
        tp2 = ttf2.paragraphs[0]
        tp2.alignment = PP_ALIGN.CENTER
        tp2.text = tag_str.upper()
        tp2.font.size = Pt(7.0)
        tp2.font.bold = True
        tp2.font.color.rgb = banner_color
        btf.margin_right = Inches(tag_w + 0.20)

    callout_h = 0.30 if bottom_callout else 0.0
    body_top = top + 0.40
    body_h = height - 0.40 - callout_h - 0.06

    body_box = slide.shapes.add_textbox(Inches(left + 0.12), Inches(body_top),
                                        Inches(width - 0.24), Inches(body_h))
    tf = body_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    is_compact = (height < 3.0)
    title_pt = Pt(9.0 if is_compact else 9.5)
    desc_pt = Pt(8.0 if is_compact else 8.5)
    space_pt = Pt(3 if is_compact else 5)

    first = True
    for item_title, item_desc in items:
        if first:
            p_t = tf.paragraphs[0]
            first = False
        else:
            p_t = tf.add_paragraph()
            p_t.space_before = space_pt
        p_t.text = f"\u2022  {item_title}"
        p_t.font.size = title_pt
        p_t.font.bold = True
        p_t.font.color.rgb = C_TEXT_DARK

        if item_desc:
            p_d = tf.add_paragraph()
            p_d.text = f"   {item_desc}"
            p_d.font.size = desc_pt
            p_d.font.color.rgb = C_TEXT_BODY
            p_d.space_before = Pt(1)

    if bottom_callout:
        call_box = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(left + 0.08), Inches(top + height - callout_h - 0.04),
            Inches(width - 0.16), Inches(callout_h))
        call_box.fill.solid()
        call_box.fill.fore_color.rgb = C_CALLOUT_BG
        call_box.line.color.rgb = RGBColor(199, 210, 254)
        call_box.line.width = Pt(0.8)
        tf_call = call_box.text_frame
        tf_call.word_wrap = True
        tf_call.margin_top = Inches(0.04)
        p_call = tf_call.paragraphs[0]
        p_call.alignment = PP_ALIGN.CENTER
        p_call.text = bottom_callout
        p_call.font.size = Pt(8.0)
        p_call.font.bold = True
        p_call.font.color.rgb = C_NAVY


# =============================================================================
# IMAGE CADRÉE SANS DÉFORMATION
# =============================================================================

def draw_framed_image(slide, left, top, width, height, image_path,
                      caption_title="", caption_sub=""):
    """
    Affiche une image dans un cadre sobre en conservant son ratio d'aspect.
    Jamais de déformation : on inscrit l'image dans l'espace et on centre.
    """
    from PIL import Image

    frame = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   Inches(left), Inches(top),
                                   Inches(width), Inches(height))
    frame.fill.solid()
    frame.fill.fore_color.rgb = C_WHITE
    frame.line.color.rgb = C_BORDER
    frame.line.width = Pt(1.0)

    header_h = 0.28 if caption_title else 0.0
    footer_h = 0.24 if caption_sub else 0.0
    avail_h = height - header_h - footer_h - 0.14
    avail_w = width - 0.16

    if caption_title:
        th = slide.shapes.add_textbox(Inches(left + 0.08), Inches(top + 0.04),
                                      Inches(width - 0.16), Inches(header_h))
        ttf = th.text_frame
        ttf.word_wrap = True
        ttf.margin_top = ttf.margin_bottom = ttf.margin_left = ttf.margin_right = 0
        tp = ttf.paragraphs[0]
        tp.alignment = PP_ALIGN.CENTER
        tp.text = caption_title.upper()
        tp.font.size = Pt(8.0)
        tp.font.bold = True
        tp.font.color.rgb = C_NAVY

    if os.path.exists(image_path):
        try:
            with Image.open(image_path) as pil_img:
                img_w_px, img_h_px = pil_img.size
            scale = min(avail_w / img_w_px, avail_h / img_h_px)
            final_w = img_w_px * scale
            final_h = img_h_px * scale
            img_left = left + 0.08 + (avail_w - final_w) / 2.0
            img_top = top + header_h + 0.06 + (avail_h - final_h) / 2.0
            slide.shapes.add_picture(image_path,
                                     Inches(img_left), Inches(img_top),
                                     Inches(final_w), Inches(final_h))
        except Exception as e:
            # Si PIL échoue, on laisse le cadre vide plutôt que de crasher
            pass

    if caption_sub:
        cap = slide.shapes.add_textbox(
            Inches(left + 0.08), Inches(top + height - footer_h - 0.02),
            Inches(width - 0.16), Inches(footer_h))
        ctf = cap.text_frame
        ctf.word_wrap = True
        ctf.margin_top = ctf.margin_bottom = ctf.margin_left = ctf.margin_right = 0
        cp = ctf.paragraphs[0]
        cp.alignment = PP_ALIGN.CENTER
        cp.text = caption_sub
        cp.font.size = Pt(7.5)
        cp.font.color.rgb = C_TEXT_MUTED


# =============================================================================
# CARTE KPI (VALEUR GRANDE + LABEL + SOUS-LABEL)
# =============================================================================

def draw_kpi_card(slide, left, top, width, height,
                  value_str, label_str, sub_str="",
                  val_color=None, bg_color=None):
    if val_color is None:
        val_color = C_NAVY
    if bg_color is None:
        bg_color = C_WHITE

    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                  Inches(left), Inches(top),
                                  Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = C_BORDER
    card.line.width = Pt(1.0)

    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_top = Inches(0.07)
    tf.margin_left = tf.margin_right = tf.margin_bottom = 0

    p_val = tf.paragraphs[0]
    p_val.alignment = PP_ALIGN.CENTER
    p_val.text = value_str
    p_val.font.size = Pt(18.0)
    p_val.font.bold = True
    p_val.font.color.rgb = val_color

    p_lbl = tf.add_paragraph()
    p_lbl.alignment = PP_ALIGN.CENTER
    p_lbl.text = label_str.upper()
    p_lbl.font.size = Pt(7.0)
    p_lbl.font.bold = True
    p_lbl.font.color.rgb = C_TEXT_DARK
    p_lbl.space_before = Pt(2)

    if sub_str:
        p_sub = tf.add_paragraph()
        p_sub.alignment = PP_ALIGN.CENTER
        p_sub.text = sub_str
        p_sub.font.size = Pt(6.5)
        p_sub.font.color.rgb = C_TEXT_MUTED
        p_sub.space_before = Pt(1)


# =============================================================================
# TEXTE CENTRÉ SIMPLE (POUR SLIDES DE TITRE ET TRANSITIONS)
# =============================================================================

def add_centered_text(slide, text, y, width=11.333, font_size=13, bold=False,
                      color=None, align=PP_ALIGN.CENTER):
    if color is None:
        color = C_WHITE
    left = (13.333 - width) / 2
    box = slide.shapes.add_textbox(Inches(left), Inches(y), Inches(width), Inches(0.60))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = align
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color


print("Modules de style et charte graphique chargés.")
