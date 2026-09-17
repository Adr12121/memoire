"""
Générateur de la présentation officielle de soutenance PFE : Adrien TRAVAILLÉ
Format 180s pur : Fond #F8FAFC, logos INSA / GEO-SIAPP, frise permanente à 2 niveaux avec noms complets,
grand conteneur blanc arrondi #FFFFFF avec bordure #CBD5E1 et angles modernes (ajustement 0.025).
Vidéo intégrée en conditions réelles : Demo_soutenance_final.mp4 (INTOUCHÉE).
Textes condensés en mots-clés percutants, sans paragraphes narratifs excessifs.
Figures rigoureusement conformes au plan du mémoire sans anticipation de cas réels.
"""

import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image

# ==================== CONSTANTES GRAPHIQUES ====================
SLIDE_W = 13.333
SLIDE_H = 7.50
TOTAL_SLIDES = 37

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

# Noms complets non tronqués dans la frise d'en-tête
CHAPTERS_DATA = [
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
    En-tête avec :
    - Logo INSA à gauche
    - Logo GEO-SIAPP à droite
    - Ligne 1 : Les 8 pilules des chapitres (noms complets)
    - Ligne 2 : Le ruban des sous-parties du chapitre actif
    - Ligne séparatrice inférieure
    """
    if os.path.exists("Logo_INSAStrasbourg.jpg"):
        slide.shapes.add_picture("Logo_INSAStrasbourg.jpg", Inches(0.50), Inches(0.12), height=Inches(0.48))
    if os.path.exists("Geosiapp.jpg"):
        slide.shapes.add_picture("Geosiapp.jpg", Inches(11.95), Inches(0.12), height=Inches(0.48))

    cx_start = 2.40
    cx_end = 11.80
    avail_w = cx_end - cx_start

    # Niveau 1 : Les 8 grands chapitres avec noms complets
    tab_w = (avail_w - 7 * 0.08) / 8.0
    for i, chap in enumerate(CHAPTERS_DATA):
        tx = cx_start + i * (tab_w + 0.08)
        is_act = (i == active_chap_idx)
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(tx), Inches(0.08), Inches(tab_w), Inches(0.32))
        pill.adjustments[0] = 0.20
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
        r.font.size = Pt(7.0)
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
            spill.adjustments[0] = 0.20
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
    """Grand conteneur blanc arrondi officiel (y=1.35 à 7.15) avec angles subtils (0.025) sans débordement."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.70), Inches(1.35), Inches(11.93), Inches(5.80))
    card.adjustments[0] = 0.025  # Arrondi fin et moderne : aucun débordement possible
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
        cap_h = 0.28 if caption else 0.0
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

def create_transition_slide(chap_idx, slide_num, title, subtitle):
    """
    Diapositive de transition SUPER SOBRE (passée en 0.5 seconde) :
    Juste le nom du chapitre en grand et une petite ligne avec de quoi on va parler.
    """
    slide = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide)
    add_header_180s(slide, chap_idx, -1)
    card = create_container_card(slide)
    add_footer_180s(slide, slide_num)
    
    # Badge Chapitre centré
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
    rb.font.color.rgb = C_WHITE
    
    # Titre du chapitre centré
    bx_t = slide.shapes.add_textbox(Inches(1.20), Inches(3.60), Inches(10.93), Inches(0.90))
    tf_t = bx_t.text_frame
    tf_t.word_wrap = True
    tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
    pt = tf_t.paragraphs[0]
    pt.alignment = PP_ALIGN.CENTER
    pt.text = title
    rt = pt.runs[0]
    rt.font.name = "Arial"
    rt.font.size = Pt(26.0)
    rt.font.bold = True
    rt.font.color.rgb = C_NAVY
    
    # Une petite ligne sobre avec de quoi on va parler
    bx_st = slide.shapes.add_textbox(Inches(1.20), Inches(4.65), Inches(10.93), Inches(0.50))
    tf_st = bx_st.text_frame
    tf_st.word_wrap = True
    tf_st.margin_left = tf_st.margin_top = tf_st.margin_right = tf_st.margin_bottom = 0
    pst = tf_st.paragraphs[0]
    pst.alignment = PP_ALIGN.CENTER
    pst.text = subtitle
    rst = pst.runs[0]
    rst.font.name = "Arial"
    rst.font.size = Pt(13.0)
    rst.font.color.rgb = C_ACCENT
    
    return slide


def set_formatted_notes(slide, notes_content):
    # Formate les notes avec prise en compte du gras (**texte**)
    import re
    notes_frame = slide.notes_slide.notes_text_frame
    notes_frame.clear()
    lines = notes_content.strip().split("\n")
    for li, line in enumerate(lines):
        p = notes_frame.paragraphs[0] if li == 0 else notes_frame.add_paragraph()
        parts = re.split(r'(\*\*.*?\*\*)', line)
        for part in parts:
            if not part:
                continue
            if part.startswith("**") and part.endswith("**"):
                r = p.add_run()
                r.text = part[2:-2]
                r.font.bold = True
            else:
                r = p.add_run()
                r.text = part

def init_standard_slide(chap_idx, sub_idx, slide_num, slide_title):
    """Initialise une diapositive standard complète au format 180s."""
    slide = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide)
    add_header_180s(slide, chap_idx, sub_idx)
    add_slide_title(slide, slide_title)
    card = create_container_card(slide)
    add_footer_180s(slide, slide_num)
    return slide, card

def add_two_column_content(slide, items, img_path, img_caption, callout_title=None, callout_text=None, kpis=None):
    """Remplissage standard 2 colonnes parfaitement ajusté à l'intérieur du conteneur sans débordement."""
    left_x = 1.05
    left_w = 5.50
    right_x = 6.78
    right_w = 5.50
    start_y = 1.60
    
    total_items = len(items)
    gap_y = 0.10
    bot_reserve = 0.85 if (callout_title or kpis) else 0.0
    # Hauteur disponible jusqu'à y=6.65 (marge de sécurité de 0.50" au-dessus du bord 7.15)
    avail_h = 5.05 - bot_reserve
    item_h = (avail_h - (total_items - 1) * gap_y) / total_items
    
    for i, item in enumerate(items):
        iy = start_y + i * (item_h + gap_y)
        b_shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left_x), Inches(iy), Inches(left_w), Inches(item_h))
        b_shp.adjustments[0] = 0.03
        b_shp.fill.solid()
        b_shp.fill.fore_color.rgb = C_BG_SLIDE
        b_shp.line.color.rgb = C_LINE
        b_shp.line.width = Pt(0.75)
        
        bx = slide.shapes.add_textbox(Inches(left_x + 0.16), Inches(iy + 0.06), Inches(left_w - 0.32), Inches(item_h - 0.12))
        tf = bx.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        
        p = tf.paragraphs[0]
        if isinstance(item, tuple):
            head, body = item
            r1 = p.add_run()
            r1.text = "• " + head + " :\n"
            r1.font.name = "Arial"
            r1.font.size = Pt(10.5)
            r1.font.bold = True
            r1.font.color.rgb = C_NAVY
            
            p2 = tf.add_paragraph()
            p2.space_before = Pt(2)
            r2 = p2.add_run()
            r2.text = body
            r2.font.name = "Arial"
            r2.font.size = Pt(9.0)
            r2.font.color.rgb = C_DARK
        else:
            r = p.add_run()
            r.text = item
            r.font.name = "Arial"
            r.font.size = Pt(9.5)
            r.font.color.rgb = C_DARK

    if callout_title and callout_text:
        cy = 5.80
        c_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left_x), Inches(cy), Inches(left_w), Inches(bot_reserve))
        c_box.adjustments[0] = 0.03
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = C_CALLOUT_BG
        c_box.line.color.rgb = C_ACCENT
        c_box.line.width = Pt(1.0)
        
        bx_c = slide.shapes.add_textbox(Inches(left_x + 0.16), Inches(cy + 0.08), Inches(left_w - 0.32), Inches(bot_reserve - 0.16))
        tf_c = bx_c.text_frame
        tf_c.word_wrap = True
        tf_c.margin_left = tf_c.margin_top = tf_c.margin_right = tf_c.margin_bottom = 0
        p_c1 = tf_c.paragraphs[0]
        p_c1.text = callout_title
        r_c1 = p_c1.runs[0]
        r_c1.font.name = "Arial"
        r_c1.font.size = Pt(9.0)
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
        cy = 5.80
        n_kpi = len(kpis)
        kw = (left_w - (n_kpi - 1) * 0.10) / n_kpi
        for ki, (kv, kl) in enumerate(kpis):
            kx = left_x + ki * (kw + 0.10)
            k_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(kx), Inches(cy), Inches(kw), Inches(bot_reserve))
            k_box.adjustments[0] = 0.04
            k_box.fill.solid()
            k_box.fill.fore_color.rgb = C_NAVY
            k_box.line.fill.background()
            
            bx_k = slide.shapes.add_textbox(Inches(kx + 0.04), Inches(cy + 0.10), Inches(kw - 0.08), Inches(bot_reserve - 0.20))
            tf_k = bx_k.text_frame
            tf_k.margin_left = tf_k.margin_top = tf_k.margin_right = tf_k.margin_bottom = 0
            pk1 = tf_k.paragraphs[0]
            pk1.text = kv
            pk1.alignment = PP_ALIGN.CENTER
            rk1 = pk1.runs[0]
            rk1.font.name = "Arial"
            rk1.font.size = Pt(15.0)
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

    # Colonne droite : cadre image
    right_h = 5.05
    frame_r = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(right_x), Inches(start_y), Inches(right_w), Inches(right_h))
    frame_r.adjustments[0] = 0.03
    frame_r.fill.solid()
    frame_r.fill.fore_color.rgb = C_BG_SLIDE
    frame_r.line.color.rgb = C_LINE
    frame_r.line.width = Pt(0.75)
    
    add_fitted_picture(slide, img_path, right_x + 0.15, start_y + 0.15, right_w - 0.30, right_h - 0.30, caption=img_caption)


# ==================== SLIDE 1 : TITRE ====================
print("Génération Slide 1 (Titre)...")
s1 = prs.slides.add_slide(blank_layout)
set_slide_bg(s1)

if os.path.exists("Logo_INSAStrasbourg.jpg"):
    s1.shapes.add_picture("Logo_INSAStrasbourg.jpg", Inches(0.80), Inches(0.48), width=Inches(2.70))
if os.path.exists("Geosiapp.jpg"):
    add_fitted_picture(s1, "Geosiapp.jpg", 10.60, 0.42, 1.93, 1.10)

card1 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.80), Inches(1.70), Inches(11.73), Inches(5.35))
card1.adjustments[0] = 0.025
card1.fill.solid()
card1.fill.fore_color.rgb = C_WHITE
card1.line.color.rgb = C_BORDER_TAB
card1.line.width = Pt(1.0)

bx_t1 = s1.shapes.add_textbox(Inches(1.20), Inches(2.20), Inches(10.93), Inches(1.50))
tf_t1 = bx_t1.text_frame
tf_t1.word_wrap = True
tf_t1.margin_left = tf_t1.margin_top = tf_t1.margin_right = tf_t1.margin_bottom = 0
pt1 = tf_t1.paragraphs[0]
pt1.alignment = PP_ALIGN.CENTER
pt1.text = "DÉVELOPPEMENT D'UN OUTIL PERMETTANT LE TRAITEMENT ET L'INSERTION DES ARCHIVES NUMÉRIQUES SUR GÉOFONCIER"
rt1 = pt1.runs[0]
rt1.font.name = "Arial"
rt1.font.size = Pt(22.0)
rt1.font.bold = True
rt1.font.color.rgb = C_NAVY

bx_sub1 = s1.shapes.add_textbox(Inches(1.20), Inches(3.90), Inches(10.93), Inches(0.40))
tf_sub1 = bx_sub1.text_frame
tf_sub1.word_wrap = True
ps1 = tf_sub1.paragraphs[0]
ps1.alignment = PP_ALIGN.CENTER
ps1.text = "SOUTENANCE DE PROJET DE FIN D'ÉTUDES : INSA STRASBOURG"
rs1 = ps1.runs[0]
rs1.font.name = "Arial"
rs1.font.size = Pt(13.0)
rs1.font.bold = True
rs1.font.color.rgb = C_ACCENT

bx_det1 = s1.shapes.add_textbox(Inches(1.20), Inches(4.50), Inches(10.93), Inches(2.20))
tf_det1 = bx_det1.text_frame
tf_det1.word_wrap = True

pd1 = tf_det1.paragraphs[0]
pd1.alignment = PP_ALIGN.CENTER
rd1_1 = pd1.add_run()
rd1_1.text = "Étudiant : "
rd1_1.font.name = "Arial"
rd1_1.font.size = Pt(14.0)
rd1_1.font.bold = True
rd1_1.font.color.rgb = C_DARK
rd1_2 = pd1.add_run()
rd1_2.text = "TRAVAILLÉ Adrien (Spécialité Topographie, Promotion 2026)"
rd1_2.font.name = "Arial"
rd1_2.font.size = Pt(14.0)
rd1_2.font.color.rgb = C_DARK

pd2 = tf_det1.add_paragraph()
pd2.space_before = Pt(8)
pd2.alignment = PP_ALIGN.CENTER
rd2_1 = pd2.add_run()
rd2_1.text = "Tuteur Entreprise : "
rd2_1.font.name = "Arial"
rd2_1.font.size = Pt(12.5)
rd2_1.font.bold = True
rd2_1.font.color.rgb = C_NAVY
rd2_2 = pd2.add_run()
rd2_2.text = "M. Gaëtan HAGUE, Géomètre-Expert associé, Cabinet GEO-SIAPP"
rd2_2.font.name = "Arial"
rd2_2.font.size = Pt(12.5)
rd2_2.font.color.rgb = C_DARK

pd3 = tf_det1.add_paragraph()
pd3.space_before = Pt(4)
pd3.alignment = PP_ALIGN.CENTER
rd3_1 = pd3.add_run()
rd3_1.text = "Directeur de PFE : "
rd3_1.font.name = "Arial"
rd3_1.font.size = Pt(12.5)
rd3_1.font.bold = True
rd3_1.font.color.rgb = C_NAVY
rd3_2 = pd3.add_run()
rd3_2.text = "M. Mathieu KOEHL, Enseignant-chercheur, INSA Strasbourg / ICube"
rd3_2.font.name = "Arial"
rd3_2.font.size = Pt(12.5)
rd3_2.font.color.rgb = C_DARK

pd4 = tf_det1.add_paragraph()
pd4.space_before = Pt(10)
pd4.alignment = PP_ALIGN.CENTER
rd4 = pd4.add_run()
rd4.text = "Soutenance du 24 septembre 2026"
rd4.font.name = "Arial"
rd4.font.size = Pt(11.0)
rd4.font.italic = True
rd4.font.color.rgb = C_ACCENT


# ==================== SLIDE 2 : SOMMAIRE ÉPURÉ ====================
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

# ==================== CHAPITRE 1 : INTRODUCTION & CONTEXTE ====================
print("Génération Chapitre 1 (Slides 3-7)...")

# Slide 3 : Transition Chapitre 1
create_transition_slide(
    chap_idx=0,
    slide_num=3,
    title="Chapitre 1 : Introduction et contexte du projet",
    subtitle="Structure d'accueil GEO-SIAPP, cadre réglementaire et problématique foncière"
)

# Slide 4 (1.1) : Présentation GEO-SIAPP
s4, c4 = init_standard_slide(0, 0, 4, "1.1 Présentation de la structure d'accueil : Le Cabinet GEO-SIAPP")

left_x4 = 1.05
left_w4 = 4.80

bx_info4 = s4.shapes.add_textbox(Inches(left_x4), Inches(1.60), Inches(left_w4), Inches(4.05))
tf_info4 = bx_info4.text_frame
tf_info4.word_wrap = True
tf_info4.margin_left = tf_info4.margin_top = tf_info4.margin_right = tf_info4.margin_bottom = 0

items_s4 = [
    ("Implantation territoriale solide",
     "Fondé en 1992 à Aubenas (siège), le cabinet s'articule autour de 4 agences : Aubenas, Pierrelatte, Vallon-Pont-d'Arc et Guilherand-Granges."),
    ("Pluridisciplinarité métier",
     "Missions foncières (bornage, divisions), urbanisme réglementaire, ingénierie VRD et topographie 3D (scanner laser, photogrammétrie drone)."),
    ("Patrimoine d'archives d'envergure",
     "Conservation intégrale de 50 ans d'archives physiques (1959–2007) issues du cabinet et des études de 5 géomètres prédécesseurs rachetées.")
]

for ii, (head_i, body_i) in enumerate(items_s4):
    p_h = tf_info4.paragraphs[0] if ii == 0 else tf_info4.add_paragraph()
    if ii > 0:
        p_h.space_before = Pt(14)
    r1 = p_h.add_run()
    r1.text = "• " + head_i + " :\n"
    r1.font.name = "Arial"
    r1.font.size = Pt(11.5)
    r1.font.bold = True
    r1.font.color.rgb = C_NAVY
    
    p_b = tf_info4.add_paragraph()
    p_b.space_before = Pt(3)
    r2 = p_b.add_run()
    r2.text = body_i
    r2.font.name = "Arial"
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = C_DARK

# Colonne droite : Photo agrandie des bureaux
img_x4 = 6.05
img_w4 = 6.23
img_h4 = 4.05
frame_r4 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(img_x4), Inches(1.60), Inches(img_w4), Inches(img_h4))
frame_r4.adjustments[0] = 0.03
frame_r4.fill.solid()
frame_r4.fill.fore_color.rgb = C_BG_SLIDE
frame_r4.line.color.rgb = C_LINE
frame_r4.line.width = Pt(0.75)
add_fitted_picture(s4, "img/bureau_cropped.jpg", img_x4 + 0.10, 1.70, img_w4 - 0.20, img_h4 - 0.40, caption="Implantation territoriale du cabinet GEO-SIAPP (Aubenas, Pierrelatte, Vallon, Guilherand)")

# Ligne du bas : 4 indicateurs clés métier (SANS durée PFE)
b_kpi_bar = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.82), Inches(11.23), Inches(0.83))
b_kpi_bar.adjustments[0] = 0.03
b_kpi_bar.fill.solid()
b_kpi_bar.fill.fore_color.rgb = C_BG_SLIDE
b_kpi_bar.line.color.rgb = C_LINE
b_kpi_bar.line.width = Pt(0.75)

kpis_s4 = [
    ("4", "Agences en Ardèche et Drôme"),
    ("5", "Géomètres-experts associés"),
    ("23 600", "Dossiers d'archives physiques"),
    ("1959–2007", "Fonds historique non indexé")
]
kw4 = 11.23 / 4.0
for ki, (kv, kl) in enumerate(kpis_s4):
    kx = 1.05 + ki * kw4
    bx_k = s4.shapes.add_textbox(Inches(kx), Inches(5.86), Inches(kw4), Inches(0.75))
    tf_k = bx_k.text_frame
    tf_k.margin_left = tf_k.margin_top = tf_k.margin_right = tf_k.margin_bottom = 0
    
    pk1 = tf_k.paragraphs[0]
    pk1.text = kv
    pk1.alignment = PP_ALIGN.CENTER
    rk1 = pk1.runs[0]
    rk1.font.name = "Arial"
    rk1.font.size = Pt(17.0)
    rk1.font.bold = True
    rk1.font.color.rgb = C_NAVY
    
    pk2 = tf_k.add_paragraph()
    pk2.space_before = Pt(2)
    pk2.text = kl
    pk2.alignment = PP_ALIGN.CENTER
    rk2 = pk2.runs[0]
    rk2.font.name = "Arial"
    rk2.font.size = Pt(8.0)
    rk2.font.color.rgb = C_MUTED

# Slide 5 (1.2) : Cadre réglementaire
s5, c5 = init_standard_slide(0, 1, 5, "1.2 Le cadre réglementaire et professionnel du bornage")

# Carte 1 (Gauche) : Cadre légal et obligation de conservation
b_jur = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(1.60), Inches(5.48), Inches(4.05))
b_jur.adjustments[0] = 0.03
b_jur.fill.solid()
b_jur.fill.fore_color.rgb = C_BG_SLIDE
b_jur.line.color.rgb = C_LINE
b_jur.line.width = Pt(0.75)

b_acc_jur = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(1.60), Inches(0.08), Inches(4.05))
b_acc_jur.adjustments[0] = 0.20
b_acc_jur.fill.solid()
b_acc_jur.fill.fore_color.rgb = C_NAVY
b_acc_jur.line.fill.background()

bx_jur = s5.shapes.add_textbox(Inches(1.25), Inches(1.75), Inches(5.15), Inches(3.75))
tf_jur = bx_jur.text_frame
tf_jur.word_wrap = True
tf_jur.margin_left = tf_jur.margin_top = tf_jur.margin_right = tf_jur.margin_bottom = 0

pj_h = tf_jur.paragraphs[0]
pj_h.text = "1. LE CADRE LÉGAL : MONOPOLE ET CONSERVATION OBLIGATOIRE"
rj_h = pj_h.runs[0]
rj_h.font.name = "Arial"
rj_h.font.size = Pt(10.5)
rj_h.font.bold = True
rj_h.font.color.rgb = C_NAVY

items_jur = [
    ("Monopole légal (Loi du 7 mai 1946)",
     "Le géomètre-expert est le seul professionnel habilité à fixer les limites réelles de propriété entre parcelles contiguës."),
    ("Droit au bornage (Article 646 du Code civil)",
     "Tout propriétaire peut contraindre son voisin au bornage contradictoire. L'acte signé lie définitivement propriétaires et successeurs."),
    ("Obligation de conservation (Décret 96-478, art. 55)",
     "Les archives de bornage doivent être rigoureusement conservées au moins 30 ans par le cabinet pour garantir la continuité des limites."),
    ("Lien direct avec le terrain",
     "L'archive ancienne n'est pas un simple document historique : elle constitue la pièce juridique obligatoire pour retravailler sur place.")
]
for pi, (hd, bd) in enumerate(items_jur):
    p_hd = tf_jur.add_paragraph()
    p_hd.space_before = Pt(6)
    r1 = p_hd.add_run()
    r1.text = "• " + hd + " :\n"
    r1.font.name = "Arial"
    r1.font.size = Pt(9.0)
    r1.font.bold = True
    r1.font.color.rgb = C_NAVY
    
    p_bd = tf_jur.add_paragraph()
    p_bd.space_before = Pt(1)
    r2 = p_bd.add_run()
    r2.text = bd
    r2.font.name = "Arial"
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = C_DARK

# Carte 2 (Droite) : Mise en valeur PRIORITAIRE : Le fossé GEODÉMAT vs Géofoncier
b_geo = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.80), Inches(1.60), Inches(5.48), Inches(4.05))
b_geo.adjustments[0] = 0.03
b_geo.fill.solid()
b_geo.fill.fore_color.rgb = RGBColor(240, 249, 255)  # Fond légèrement bleuté
b_geo.line.color.rgb = C_ACCENT
b_geo.line.width = Pt(1.2)

b_acc_geo = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.80), Inches(1.60), Inches(0.08), Inches(4.05))
b_acc_geo.adjustments[0] = 0.20
b_acc_geo.fill.solid()
b_acc_geo.fill.fore_color.rgb = C_ACCENT
b_acc_geo.line.fill.background()

bx_geo = s5.shapes.add_textbox(Inches(7.00), Inches(1.75), Inches(5.15), Inches(3.75))
tf_geo = bx_geo.text_frame
tf_geo.word_wrap = True
tf_geo.margin_left = tf_geo.margin_top = tf_geo.margin_right = tf_geo.margin_bottom = 0

pg_h = tf_geo.paragraphs[0]
pg_h.text = "2. LE FOSSÉ OPÉRATIONNEL : ÉTAT (GEODÉMAT) vs ARCHIVES PRIVÉES"
rg_h = pg_h.runs[0]
rg_h.font.name = "Arial"
rg_h.font.size = Pt(10.5)
rg_h.font.bold = True
rg_h.font.color.rgb = C_ACCENT

items_geo = [
    ("Le programme public GEODÉMAT (DGFiP)",
     "L'État numérise exclusivement le cadastre fiscal et les documents DMPC enregistrés. Les archives privées de bornage en sont totalement absentes."),
    ("La responsabilité exclusive du cabinet",
     "L'État ne viendra jamais numériser les cartons du cabinet. L'indexation de nos 23 600 dossiers d'Aubenas repose à 100% sur notre initiative."),
    ("Le portail national Géofoncier (OGE)",
     "Créé par l'Ordre en 2010 pour centraliser les interventions foncières sous forme de pastilles géolocalisées consultables par la profession."),
    ("Priorité absolue du projet PFE",
     "Combler ce fossé en automatisant le versement des archives historiques sur Géofoncier pour rendre opposables 50 ans de travaux.")
]
for pi, (hd, bd) in enumerate(items_geo):
    p_hd = tf_geo.add_paragraph()
    p_hd.space_before = Pt(6)
    r1 = p_hd.add_run()
    r1.text = "• " + hd + " :\n"
    r1.font.name = "Arial"
    r1.font.size = Pt(9.0)
    r1.font.bold = True
    r1.font.color.rgb = C_ACCENT
    
    p_bd = tf_geo.add_paragraph()
    p_bd.space_before = Pt(1)
    r2 = p_bd.add_run()
    r2.text = bd
    r2.font.name = "Arial"
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = C_DARK

# Callout bas de synthèse logique
b_syn5 = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.82), Inches(11.23), Inches(0.83))
b_syn5.adjustments[0] = 0.03
b_syn5.fill.solid()
b_syn5.fill.fore_color.rgb = C_CALLOUT_BG
b_syn5.line.color.rgb = C_ACCENT
b_syn5.line.width = Pt(1.0)

bx_s5 = s5.shapes.add_textbox(Inches(1.20), Inches(5.88), Inches(10.93), Inches(0.72))
tf_s5 = bx_s5.text_frame
tf_s5.word_wrap = True
tf_s5.margin_left = tf_s5.margin_top = tf_s5.margin_right = tf_s5.margin_bottom = 0
ps5_1 = tf_s5.paragraphs[0]
ps5_1.text = "ENJEU PRIORITAIRE DU PROJET :"
rs5_1 = ps5_1.runs[0]
rs5_1.font.name = "Arial"
rs5_1.font.size = Pt(9.5)
rs5_1.font.bold = True
rs5_1.font.color.rgb = C_ACCENT

ps5_2 = tf_s5.add_paragraph()
ps5_2.space_before = Pt(2)
rs5_2 = ps5_2.add_run()
rs5_2.text = "L'État gère le cadastre fiscal, mais l'opposabilité juridique repose sur les archives privées du cabinet. Sans outil d'indexation interne, 50 ans de limites restent invisibles pour la profession."
rs5_2.font.name = "Arial"
rs5_2.font.size = Pt(8.5)
rs5_2.font.color.rgb = C_DARK

# Slide 6 (1.3) : Contexte du projet : La recherche d'antériorité obligatoire
s6, c6 = init_standard_slide(0, 2, 6, "1.3 Contexte du projet : La recherche d'antériorité obligatoire")

# Colonne gauche (4.60 in) : Le principe juridique fondamental
b_principe = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(1.60), Inches(4.60), Inches(4.05))
b_principe.adjustments[0] = 0.03
b_principe.fill.solid()
b_principe.fill.fore_color.rgb = C_BG_SLIDE
b_principe.line.color.rgb = C_NAVY
b_principe.line.width = Pt(1.0)

b_acc_p = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(1.60), Inches(0.08), Inches(4.05))
b_acc_p.adjustments[0] = 0.20
b_acc_p.fill.solid()
b_acc_p.fill.fore_color.rgb = C_NAVY
b_acc_p.line.fill.background()

bx_p = s6.shapes.add_textbox(Inches(1.25), Inches(1.75), Inches(4.25), Inches(3.75))
tf_p = bx_p.text_frame
tf_p.word_wrap = True
tf_p.margin_left = tf_p.margin_top = tf_p.margin_right = tf_p.margin_bottom = 0

pp_tag = tf_p.paragraphs[0]
pp_tag.text = "RÈGLE D'OR DE LA PROFESSION"
rp_tag = pp_tag.runs[0]
rp_tag.font.name = "Arial"
rp_tag.font.size = Pt(9.0)
rp_tag.font.bold = True
rp_tag.font.color.rgb = C_NAVY

pp_tit = tf_p.add_paragraph()
pp_tit.space_before = Pt(3)
pp_tit.text = "« Bornage sur bornage ne vaut »"
rp_tit = pp_tit.runs[0]
rp_tit.font.name = "Arial"
rp_tit.font.size = Pt(13.0)
rp_tit.font.bold = True
rp_tit.font.color.rgb = C_NAVY

bullets_principe = [
    ("Article 646 du Code civil", "Obligation légale absolue de rechercher tout acte de bornage antérieur avant d'intervenir."),
    ("Continuité des limites", "Rétablissement obligatoire de la ligne séparative d'origine fixée par les prédécesseurs."),
    ("Sécurité juridique", "Tout nouvel acte dressé sans consultation de l'antériorité est juridiquement contestable."),
    ("Protection du client", "Garantie incontestable contre les litiges de voisinage et contestations ultérieures.")
]
for pi, (bh, bd) in enumerate(bullets_principe):
    p_b = tf_p.add_paragraph()
    p_b.space_before = Pt(6)
    r1 = p_b.add_run()
    r1.text = "• " + bh + " : "
    r1.font.name = "Arial"
    r1.font.size = Pt(8.5)
    r1.font.bold = True
    r1.font.color.rgb = C_NAVY
    r2 = p_b.add_run()
    r2.text = bd
    r2.font.name = "Arial"
    r2.font.size = Pt(8.0)
    r2.font.color.rgb = C_DARK

# Colonne droite (6.35 in) : Duel entre la réalité manuelle et la solution PFE
# Carte haut droite : Réalité manuelle (Alerte Orange)
b_man = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.93), Inches(1.60), Inches(6.35), Inches(1.95))
b_man.adjustments[0] = 0.03
b_man.fill.solid()
b_man.fill.fore_color.rgb = C_BG_SLIDE
b_man.line.color.rgb = C_WARN
b_man.line.width = Pt(1.0)

b_acc_man = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.93), Inches(1.60), Inches(0.08), Inches(1.95))
b_acc_man.adjustments[0] = 0.20
b_acc_man.fill.solid()
b_acc_man.fill.fore_color.rgb = C_WARN
b_acc_man.line.fill.background()

bx_man = s6.shapes.add_textbox(Inches(6.13), Inches(1.70), Inches(6.00), Inches(1.75))
tf_man = bx_man.text_frame
tf_man.word_wrap = True
tf_man.margin_left = tf_man.margin_top = tf_man.margin_right = tf_man.margin_bottom = 0

pm_h = tf_man.paragraphs[0]
pm_h.text = "LA RÉALITÉ MANUELLE ACTUELLE : UN FREIN OPÉRATIONNEL"
rm_h = pm_h.runs[0]
rm_h.font.name = "Arial"
rm_h.font.size = Pt(9.5)
rm_h.font.bold = True
rm_h.font.color.rgb = C_WARN

bullets_man = [
    ("Fouille physique chronophage", "15 à 30 minutes requises par dossier dans les cartons d'Aubenas."),
    ("Risque d'omission", "Plans anciens non trouvés faute d'indexation moderne ou parcelles renommées."),
    ("Fragilité des supports", "Documents de 1959 à 2007 (calques, tirages) usés par les manipulations.")
]
for bh, bd in bullets_man:
    p_b = tf_man.add_paragraph()
    p_b.space_before = Pt(3)
    r1 = p_b.add_run()
    r1.text = "✗ " + bh + " : "
    r1.font.name = "Arial"
    r1.font.size = Pt(8.5)
    r1.font.bold = True
    r1.font.color.rgb = C_DARK
    r2 = p_b.add_run()
    r2.text = bd
    r2.font.name = "Arial"
    r2.font.size = Pt(8.0)
    r2.font.color.rgb = C_MUTED

# Carte bas droite : Solution PFE (Succès Vert/Bleu)
b_sol = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.93), Inches(3.70), Inches(6.35), Inches(1.95))
b_sol.adjustments[0] = 0.03
b_sol.fill.solid()
b_sol.fill.fore_color.rgb = RGBColor(240, 253, 244)
b_sol.line.color.rgb = C_SUCCESS
b_sol.line.width = Pt(1.0)

b_acc_sol = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.93), Inches(3.70), Inches(0.08), Inches(1.95))
b_acc_sol.adjustments[0] = 0.20
b_acc_sol.fill.solid()
b_acc_sol.fill.fore_color.rgb = C_SUCCESS
b_acc_sol.line.fill.background()

bx_sol = s6.shapes.add_textbox(Inches(6.13), Inches(3.80), Inches(6.00), Inches(1.75))
tf_sol = bx_sol.text_frame
tf_sol.word_wrap = True
tf_sol.margin_left = tf_sol.margin_top = tf_sol.margin_right = tf_sol.margin_bottom = 0

ps_h = tf_sol.paragraphs[0]
ps_h.text = "LA SOLUTION DÉVELOPPÉE DANS CE PFE : INDEXATION AUTOMATIQUE"
rs_h = ps_h.runs[0]
rs_h.font.name = "Arial"
rs_h.font.size = Pt(9.5)
rs_h.font.bold = True
rs_h.font.color.rgb = C_SUCCESS

bullets_sol = [
    ("Extraction automatique", "Lecture des métadonnées (commune, date, numéro, parcelles) sans saisie."),
    ("Recherche en 1 clic", "Consultation instantanée des antériorités depuis la carte Géofoncier."),
    ("Préservation du patrimoine", "Numérisation pérenne des 23 600 dossiers sans manipulation physique.")
]
for bh, bd in bullets_sol:
    p_b = tf_sol.add_paragraph()
    p_b.space_before = Pt(3)
    r1 = p_b.add_run()
    r1.text = "✓ " + bh + " : "
    r1.font.name = "Arial"
    r1.font.size = Pt(8.5)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(21, 128, 61)
    r2 = p_b.add_run()
    r2.text = bd
    r2.font.name = "Arial"
    r2.font.size = Pt(8.0)
    r2.font.color.rgb = C_DARK

# Ligne de KPIs du bas (SANS 100% opposabilité)
b_kpi6 = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.82), Inches(11.23), Inches(0.83))
b_kpi6.adjustments[0] = 0.03
b_kpi6.fill.solid()
b_kpi6.fill.fore_color.rgb = C_BG_SLIDE
b_kpi6.line.color.rgb = C_LINE
b_kpi6.line.width = Pt(0.75)

kpis_s6 = [
    ("23 600", "Dossiers d'archives physiques"),
    ("15 à 30 min", "Par recherche manuelle"),
    ("< 1 min", "Recherche ciblée Géofoncier"),
    ("50 ans", "Fonds historique sécurisé")
]
kw6 = 11.23 / 4.0
for ki, (kv, kl) in enumerate(kpis_s6):
    kx = 1.05 + ki * kw6
    bx_k = s6.shapes.add_textbox(Inches(kx), Inches(5.86), Inches(kw6), Inches(0.75))
    tf_k = bx_k.text_frame
    tf_k.margin_left = tf_k.margin_top = tf_k.margin_right = tf_k.margin_bottom = 0
    
    pk1 = tf_k.paragraphs[0]
    pk1.text = kv
    pk1.alignment = PP_ALIGN.CENTER
    rk1 = pk1.runs[0]
    rk1.font.name = "Arial"
    rk1.font.size = Pt(16.0)
    rk1.font.bold = True
    rk1.font.color.rgb = C_NAVY if ki != 2 else C_SUCCESS
    
    pk2 = tf_k.add_paragraph()
    pk2.space_before = Pt(2)
    pk2.text = kl
    pk2.alignment = PP_ALIGN.CENTER
    rk2 = pk2.runs[0]
    rk2.font.name = "Arial"
    rk2.font.size = Pt(8.0)
    rk2.font.color.rgb = C_MUTED

# Slide 7 (1.4) : Problématique et contraintes techniques (Phrase exacte du mémoire)
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
    p1.text = "• " + c_hd + " :"
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
rp2.font.color.rgb = C_NAVY


# ==================== CHAPITRE 2 : ANALYSE MÉTIER & DONNÉES ====================
print("Génération Chapitre 2 (Slides 8-12)...")

# Slide 8 : Transition Chapitre 2
create_transition_slide(
    chap_idx=1,
    slide_num=8,
    title="Chapitre 2 : Analyse métier et enjeux documentaires",
    subtitle="Distinction cadastre et actes, spécifications API Géofoncier et filiation parcellaire"
)

# Slide 9 (2.1) : Cadastre fiscal et propriété foncière
s9, c9 = init_standard_slide(1, 0, 9, "2.1 Cadastre fiscal et propriété foncière : distinction juridique")

# Tableau comparatif direct : Duel juridique Cadastre vs PV de Bornage
headers_s9 = ["Critère d'analyse", "Cadastre fiscal (PCI DGFiP)", "Acte de bornage (GEO-SIAPP)"]
col_w_s9 = [2.60, 4.30, 4.33]

rows_s9 = [
    ("Objectif d'origine",
     "Instrument fiscal napoléonien (1807) créé pour la juste répartition de la contribution foncière.",
     "Délimitation contradictoire réelle de la propriété privée entre parcelles contiguës (Art. 646 C. civ.)."),
    ("Valeur juridique",
     "Simple présomption administrative. La jurisprudence constante rappelle qu'il ne vaut pas titre.",
     "Acte sous seing privé opposable aux tiers liant définitivement les propriétaires actuels et successeurs."),
    ("Précision terrain",
     "Plans anciens vectorisés, erreurs et décalages dépassant souvent plusieurs mètres en zone rurale.",
     "Mesures et repérage centimétriques exacts des bornes physiques scellées sur le terrain."),
    ("Devant le juge",
     "Écarté par les tribunaux dès lors qu'une limite réelle contradictoire est prouvée.",
     "Force probante prépondérante retenue en priorité par les juges lors de tout litige foncier.")
]

ty9 = 1.55
rh9 = 0.88

# En-tête du tableau
hx9 = 1.05
for hi, htext in enumerate(headers_s9):
    h_box = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(hx9), Inches(ty9), Inches(col_w_s9[hi]), Inches(0.40))
    h_box.adjustments[0] = 0.08
    h_box.fill.solid()
    h_box.fill.fore_color.rgb = C_NAVY if hi < 2 else C_ACCENT
    h_box.line.fill.background()
    
    bx = s9.shapes.add_textbox(Inches(hx9 + 0.05), Inches(ty9 + 0.05), Inches(col_w_s9[hi] - 0.10), Inches(0.30))
    tf = bx.text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = htext
    p.alignment = PP_ALIGN.CENTER
    r = p.runs[0]
    r.font.name = "Arial"
    r.font.size = Pt(9.5)
    r.font.bold = True
    r.font.color.rgb = C_WHITE
    hx9 += col_w_s9[hi]

# Lignes du tableau
for ri, (c_crit, c_cad, c_bor) in enumerate(rows_s9):
    ry9 = ty9 + 0.40 + ri * rh9
    bg_col = C_CARD_BG if ri % 2 == 0 else C_BG_SLIDE
    
    # Col 1 : Critère
    b1 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(ry9), Inches(col_w_s9[0]), Inches(rh9))
    b1.adjustments[0] = 0.06
    b1.fill.solid()
    b1.fill.fore_color.rgb = C_BG_TAB
    b1.line.color.rgb = C_LINE
    b1.line.width = Pt(0.75)
    bx1 = s9.shapes.add_textbox(Inches(1.15), Inches(ry9 + 0.12), Inches(col_w_s9[0] - 0.20), Inches(rh9 - 0.24))
    tf1 = bx1.text_frame
    tf1.word_wrap = True
    tf1.margin_left = tf1.margin_top = tf1.margin_right = tf1.margin_bottom = 0
    p1 = tf1.paragraphs[0]
    p1.text = c_crit
    r1 = p1.runs[0]
    r1.font.name = "Arial"
    r1.font.size = Pt(9.0)
    r1.font.bold = True
    r1.font.color.rgb = C_NAVY
    
    # Col 2 : Cadastre
    b2 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05 + col_w_s9[0]), Inches(ry9), Inches(col_w_s9[1]), Inches(rh9))
    b2.adjustments[0] = 0.06
    b2.fill.solid()
    b2.fill.fore_color.rgb = bg_col
    b2.line.color.rgb = C_LINE
    b2.line.width = Pt(0.75)
    bx2 = s9.shapes.add_textbox(Inches(1.15 + col_w_s9[0]), Inches(ry9 + 0.10), Inches(col_w_s9[1] - 0.20), Inches(rh9 - 0.20))
    tf2 = bx2.text_frame
    tf2.word_wrap = True
    tf2.margin_left = tf2.margin_top = tf2.margin_right = tf2.margin_bottom = 0
    p2 = tf2.paragraphs[0]
    p2.text = c_cad
    r2 = p2.runs[0]
    r2.font.name = "Arial"
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = C_DARK
    
    # Col 3 : Bornage
    b3 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05 + col_w_s9[0] + col_w_s9[1]), Inches(ry9), Inches(col_w_s9[2]), Inches(rh9))
    b3.adjustments[0] = 0.06
    b3.fill.solid()
    b3.fill.fore_color.rgb = RGBColor(240, 249, 255) if ri % 2 == 0 else C_WHITE
    b3.line.color.rgb = C_ACCENT
    b3.line.width = Pt(0.75)
    bx3 = s9.shapes.add_textbox(Inches(1.15 + col_w_s9[0] + col_w_s9[1]), Inches(ry9 + 0.10), Inches(col_w_s9[2] - 0.20), Inches(rh9 - 0.20))
    tf3 = bx3.text_frame
    tf3.word_wrap = True
    tf3.margin_left = tf3.margin_top = tf3.margin_right = tf3.margin_bottom = 0
    p3 = tf3.paragraphs[0]
    p3.text = c_bor
    r3 = p3.runs[0]
    r3.font.name = "Arial"
    r3.font.size = Pt(8.5)
    r3.font.color.rgb = C_NAVY

# Bandeau bas officiel
b_syn9 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.82), Inches(11.23), Inches(0.83))
b_syn9.adjustments[0] = 0.03
b_syn9.fill.solid()
b_syn9.fill.fore_color.rgb = C_CALLOUT_BG
b_syn9.line.color.rgb = C_ACCENT
b_syn9.line.width = Pt(1.0)

bx_s9 = s9.shapes.add_textbox(Inches(1.20), Inches(5.88), Inches(10.93), Inches(0.70))
tf_s9 = bx_s9.text_frame
tf_s9.word_wrap = True
tf_s9.margin_left = tf_s9.margin_top = tf_s9.margin_right = tf_s9.margin_bottom = 0
p_s9_1 = tf_s9.paragraphs[0]
p_s9_1.text = "CONCLUSION DUEL JURIDIQUE :"
r_s9_1 = p_s9_1.runs[0]
r_s9_1.font.name = "Arial"
r_s9_1.font.size = Pt(9.5)
r_s9_1.font.bold = True
r_s9_1.font.color.rgb = C_ACCENT

p_s9_2 = tf_s9.add_paragraph()
p_s9_2.space_before = Pt(2)
r_s9_2 = p_s9_2.add_run()
r_s9_2.text = "Le plan cadastral n'est qu'un outil fiscal indicatif. Seuls les procès-verbaux de bornage conservés par GEO-SIAPP fixent juridiquement la limite réelle de propriété."
r_s9_2.font.name = "Arial"
r_s9_2.font.size = Pt(8.5)
r_s9_2.font.color.rgb = C_DARK

# Slide 10 (2.2) : Versement Géofoncier : Spécifications de l'API REST
s10, c10 = init_standard_slide(1, 1, 10, "2.2 Le versement sur Géofoncier : Spécifications de l'API REST")

left_x10 = 1.05
left_w10 = 5.50

# Bloc 1 : Spécifications API REST
b_api = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left_x10), Inches(1.60), Inches(left_w10), Inches(1.65))
b_api.adjustments[0] = 0.03
b_api.fill.solid()
b_api.fill.fore_color.rgb = C_BG_SLIDE
b_api.line.color.rgb = C_LINE
b_api.line.width = Pt(0.75)

bx_api = s10.shapes.add_textbox(Inches(left_x10 + 0.16), Inches(1.68), Inches(left_w10 - 0.32), Inches(1.50))
tf_api = bx_api.text_frame
tf_api.word_wrap = True
tf_api.margin_left = tf_api.margin_top = tf_api.margin_right = tf_api.margin_bottom = 0

p_apih = tf_api.paragraphs[0]
p_apih.text = "SPÉCIFICATIONS SWAGGER ET API REST GÉOFONCIER"
r_apih = p_apih.runs[0]
r_apih.font.name = "Arial"
r_apih.font.size = Pt(10.5)
r_apih.font.bold = True
r_apih.font.color.rgb = C_NAVY

bullets_api = [
    ("Authentification sécurisée", "Token JWT Bearer renouvelable toutes les 24h via les identifiants du cabinet."),
    ("Endpoints dédiés", "Route /rfuoge pour les métadonnées de l'acte et /dossiersoge pour la pièce jointe."),
    ("Validation serveur", "Code HTTP 201 Created confirmant la publication de la pastille sur le portail.")
]
for bh, bd in bullets_api:
    p_b = tf_api.add_paragraph()
    p_b.space_before = Pt(3)
    r1 = p_b.add_run()
    r1.text = "• " + bh + " : "
    r1.font.name = "Arial"
    r1.font.size = Pt(8.5)
    r1.font.bold = True
    r1.font.color.rgb = C_ACCENT
    r2 = p_b.add_run()
    r2.text = bd
    r2.font.name = "Arial"
    r2.font.size = Pt(8.0)
    r2.font.color.rgb = C_DARK

# Bloc 2 : Structure JSON obligatoire avec mention explicite des données fictives
b_json = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left_x10), Inches(3.35), Inches(left_w10), Inches(2.35))
b_json.adjustments[0] = 0.03
b_json.fill.solid()
b_json.fill.fore_color.rgb = C_BG_SLIDE
b_json.line.color.rgb = C_LINE
b_json.line.width = Pt(0.75)

bx_json = s10.shapes.add_textbox(Inches(left_x10 + 0.16), Inches(3.42), Inches(left_w10 - 0.32), Inches(2.20))
tf_json = bx_json.text_frame
tf_json.word_wrap = True
tf_json.margin_left = tf_json.margin_top = tf_json.margin_right = tf_json.margin_bottom = 0

pj_h = tf_json.paragraphs[0]
pj_h.text = "STRUCTURE DU PAYLOAD JSON (/rfuoge) • "
rj_h = pj_h.runs[0]
rj_h.font.name = "Arial"
rj_h.font.size = Pt(10.0)
rj_h.font.bold = True
rj_h.font.color.rgb = C_NAVY

rj_badge = pj_h.add_run()
rj_badge.text = "[Données fictives et anonymisées]"
rj_badge.font.name = "Arial"
rj_badge.font.size = Pt(8.0)
rj_badge.font.bold = True
rj_badge.font.color.rgb = C_ACCENT

fields_json = [
    ("code_insee", "\"07019\"", "Code officiel COG de la commune fictive"),
    ("date_operation", "\"1984-05-12\"", "Date de l'acte (format ISO 8601 AAAA-MM-JJ)"),
    ("type_acte", "\"Bo\"", "Nature de l'opération (Bo = Bornage, DA = Division)"),
    ("reference_dossier", "\"GEO-1984-089\"", "Identifiant unique interne au cabinet"),
    ("geometry", "Point [801452.35, 6391204.18]", "Centroïde calculé en Lambert-93 (EPSG:2154)")
]
for f_name, f_val, f_desc in fields_json:
    p_f = tf_json.add_paragraph()
    p_f.space_before = Pt(2)
    r1 = p_f.add_run()
    r1.text = f"• {f_name} : "
    r1.font.name = "Arial"
    r1.font.size = Pt(8.5)
    r1.font.bold = True
    r1.font.color.rgb = C_DARK
    r2 = p_f.add_run()
    r2.text = f"{f_val} "
    r2.font.name = "Arial"
    r2.font.size = Pt(8.5)
    r2.font.bold = True
    r2.font.color.rgb = C_ACCENT
    r3 = p_f.add_run()
    r3.text = f"({f_desc})"
    r3.font.name = "Arial"
    r3.font.size = Pt(7.5)
    r3.font.color.rgb = C_MUTED

# Callout bas gauche
b_syn10 = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left_x10), Inches(5.82), Inches(left_w10), Inches(0.83))
b_syn10.adjustments[0] = 0.03
b_syn10.fill.solid()
b_syn10.fill.fore_color.rgb = C_CALLOUT_BG
b_syn10.line.color.rgb = C_ACCENT
b_syn10.line.width = Pt(1.0)

bx_s10 = s10.shapes.add_textbox(Inches(left_x10 + 0.16), Inches(5.88), Inches(left_w10 - 0.32), Inches(0.70))
tf_s10 = bx_s10.text_frame
tf_s10.word_wrap = True
tf_s10.margin_left = tf_s10.margin_top = tf_s10.margin_right = tf_s10.margin_bottom = 0
ps1 = tf_s10.paragraphs[0]
ps1.text = "OBJECTIF DE LA CHAÎNE D'EXTRACTION :"
rs1 = ps1.runs[0]
rs1.font.name = "Arial"
rs1.font.size = Pt(9.0)
rs1.font.bold = True
rs1.font.color.rgb = C_ACCENT

ps2 = tf_s10.add_paragraph()
ps2.space_before = Pt(2)
rs2 = ps2.add_run()
rs2.text = "Extraire automatiquement ces 5 champs cibles depuis les scans d'archives pour alimenter directement l'API Géofoncier sans saisie manuelle."
rs2.font.name = "Arial"
rs2.font.size = Pt(8.0)
rs2.font.color.rgb = C_DARK

# Colonne droite : Schéma arborescence
frame_r10 = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.80), Inches(1.60), Inches(5.48), Inches(5.05))
frame_r10.adjustments[0] = 0.03
frame_r10.fill.solid()
frame_r10.fill.fore_color.rgb = C_BG_SLIDE
frame_r10.line.color.rgb = C_LINE
frame_r10.line.width = Pt(0.75)
add_fitted_picture(s10, "img/arborescence_versement_geofoncier.png", 6.95, 1.75, 5.18, 4.75, caption="Arborescence des données obligatoires à renseigner pour le versement au RFU (Mémoire, Chapitre 2)")

# Slide 11 (2.3) : Typologie des archives
s11, c11 = init_standard_slide(1, 2, 11, "2.3 Typologie des archives foncières et caractéristiques des fonds SIAPP")

left_x11 = 1.05
left_w11 = 5.48

# 3 cartes typologiques travaillées avec badges colorés
types_archives = [
    ("IMPRIMÉ NORMÉ • DÉCRET 1955",
     "Formulaires Cerfa DMPC et DA",
     "Documents normés accompagnant les divisions foncières. Cartouches rigides, typographie dactylographiée ou imprimée, structure tabulaire géométrique.",
     C_NAVY),
    ("MANUSCRIT TABULAIRE • FONDS HISTORIQUES",
     "Registres d'affaires centralisateurs (Fonds A & B)",
     "Grands cahiers récapitulatifs tenus par les anciens géomètres. Écriture cursive inclinée, abréviations fréquentes (« id. », « /s/ »), colonnes libres.",
     C_ACCENT),
    ("ACTES LIBRES & CALQUES DIAZOÏQUES",
     "Procès-Verbaux de bornage & Croquis",
     "Actes sous seing privé en texte libre (dactylographiés ou manuscrits). Présence de calques translucides, tirages diazoïques et tampons administratifs.",
     RGBColor(180, 83, 9))
]

card_h11 = 1.25
gap11 = 0.12
for ti, (t_tag, t_title, t_desc, t_col) in enumerate(types_archives):
    cy = 1.60 + ti * (card_h11 + gap11)
    b_t = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left_x11), Inches(cy), Inches(left_w11), Inches(card_h11))
    b_t.adjustments[0] = 0.03
    b_t.fill.solid()
    b_t.fill.fore_color.rgb = C_BG_SLIDE
    b_t.line.color.rgb = t_col
    b_t.line.width = Pt(0.75)
    
    b_acc = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left_x11), Inches(cy), Inches(0.08), Inches(card_h11))
    b_acc.adjustments[0] = 0.20
    b_acc.fill.solid()
    b_acc.fill.fore_color.rgb = t_col
    b_acc.line.fill.background()
    
    bx_t = s11.shapes.add_textbox(Inches(left_x11 + 0.18), Inches(cy + 0.08), Inches(left_w11 - 0.28), Inches(card_h11 - 0.16))
    tf_t = bx_t.text_frame
    tf_t.word_wrap = True
    tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
    
    pt1 = tf_t.paragraphs[0]
    pt1.text = t_tag
    rt1 = pt1.runs[0]
    rt1.font.name = "Arial"
    rt1.font.size = Pt(8.0)
    rt1.font.bold = True
    rt1.font.color.rgb = t_col
    
    pt2 = tf_t.add_paragraph()
    pt2.space_before = Pt(2)
    pt2.text = t_title
    rt2 = pt2.runs[0]
    rt2.font.name = "Arial"
    rt2.font.size = Pt(10.5)
    rt2.font.bold = True
    rt2.font.color.rgb = C_NAVY
    
    pt3 = tf_t.add_paragraph()
    pt3.space_before = Pt(2)
    pt3.text = t_desc
    rt3 = pt3.runs[0]
    rt3.font.name = "Arial"
    rt3.font.size = Pt(8.0)
    rt3.font.color.rgb = C_DARK

# Colonne droite : Graphique de répartition avec cadre
frame_r11 = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.80), Inches(1.60), Inches(5.48), Inches(4.05))
frame_r11.adjustments[0] = 0.03
frame_r11.fill.solid()
frame_r11.fill.fore_color.rgb = C_BG_SLIDE
frame_r11.line.color.rgb = C_LINE
frame_r11.line.width = Pt(0.75)
add_fitted_picture(s11, "img/repartition_archives.png", 6.95, 1.70, 5.18, 3.85, caption="Répartition estimée des 23 600 dossiers du fonds SIAPP (5 fonds de géomètres prédécesseurs)")

# Callout bas de synthèse
b_syn11 = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.82), Inches(11.23), Inches(0.83))
b_syn11.adjustments[0] = 0.03
b_syn11.fill.solid()
b_syn11.fill.fore_color.rgb = C_CALLOUT_BG
b_syn11.line.color.rgb = C_ACCENT
b_syn11.line.width = Pt(1.0)

bx_s11 = s11.shapes.add_textbox(Inches(1.20), Inches(5.88), Inches(10.93), Inches(0.70))
tf_s11 = bx_s11.text_frame
tf_s11.word_wrap = True
tf_s11.margin_left = tf_s11.margin_top = tf_s11.margin_right = tf_s11.margin_bottom = 0
ps11_1 = tf_s11.paragraphs[0]
ps11_1.text = "IMPACT TECHNIQUE DE L'HÉTÉROGÉNÉITÉ DOCUMENTAIRE :"
rs11_1 = ps11_1.runs[0]
rs11_1.font.name = "Arial"
rs11_1.font.size = Pt(9.5)
rs11_1.font.bold = True
rs11_1.font.color.rgb = C_ACCENT

ps11_2 = tf_s11.add_paragraph()
ps11_2.space_before = Pt(2)
rs11_2 = ps11_2.add_run()
rs11_2.text = "L'hétérogénéité entre Cerfa rigides, registres cursifs et calques libres interdit un moteur OCR unique et impose une chaîne modulaire avec segmentation spatiale préalable."
rs11_2.font.name = "Arial"
rs11_2.font.size = Pt(8.5)
rs11_2.font.color.rgb = C_DARK

# Slide 12 (2.4) : La filiation parcellaire et traçabilité cadastrale
s12, c12 = init_standard_slide(1, 3, 12, "2.4 La filiation parcellaire et traçabilité cadastrale")

steps_12 = [
    ("1. IDENTIFIANT ACTUEL DU TERRAIN",
     [
         ("Parcelle d'intervention", "Numéro contemporain visible aujourd'hui (ex. A 115 ou A 116)."),
         ("Recherche Géofoncier", "L'opérateur clique sur la parcelle du cadastre actuel."),
         ("Point de rupture", "Cette parcelle n'existait pas lors de l'acte historique.")
     ],
     C_NAVY),
    ("2. DIVISION ET PARCELLE MÈRE DISPARUE",
     [
         ("Archive d'époque", "Rédigée sous la parcelle mère d'origine (ex. section A n° 14)."),
         ("Mutation cadastrale", "La parcelle A 14 est radiée de la matrice lors de la division."),
         ("Risque métier", "Sans lien historique, l'acte antérieur reste introuvable sur A 115.")
     ],
     C_WARN),
    ("3. RECONSTITUTION DE LA GÉNÉALOGIE",
     [
         ("Filiation cadastrale", "Traçage de la chaîne des divisions successives (A 14 → A 115 + A 116)."),
         ("Liaison logicielle", "Rattachement automatique de l'acte ancien aux parcelles filles."),
         ("Résultat Géofoncier", "L'antériorité apparaît immédiatement lors du clic sur A 115.")
     ],
     C_SUCCESS)
]

cw12 = 3.55
cg12 = 0.29
for si, (st_h, st_bullets, st_col) in enumerate(steps_12):
    cx = 1.05 + si * (cw12 + cg12)
    b_st = s12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(1.60), Inches(cw12), Inches(1.55))
    b_st.adjustments[0] = 0.04
    b_st.fill.solid()
    b_st.fill.fore_color.rgb = C_BG_SLIDE
    b_st.line.color.rgb = st_col
    b_st.line.width = Pt(1.0)
    
    bx_st = s12.shapes.add_textbox(Inches(cx + 0.14), Inches(1.68), Inches(cw12 - 0.28), Inches(1.38))
    tf_st = bx_st.text_frame
    tf_st.word_wrap = True
    tf_st.margin_left = tf_st.margin_top = tf_st.margin_right = tf_st.margin_bottom = 0
    
    p1 = tf_st.paragraphs[0]
    p1.text = st_h
    r1 = p1.runs[0]
    r1.font.name = "Arial"
    r1.font.size = Pt(9.5)
    r1.font.bold = True
    r1.font.color.rgb = st_col
    
    for b_k, b_v in st_bullets:
        p_sub = tf_st.add_paragraph()
        p_sub.space_before = Pt(3)
        r_k = p_sub.add_run()
        r_k.text = "• " + b_k + " : "
        r_k.font.name = "Arial"
        r_k.font.size = Pt(8.0)
        r_k.font.bold = True
        r_k.font.color.rgb = C_DARK
        r_v = p_sub.add_run()
        r_v.text = b_v
        r_v.font.name = "Arial"
        r_v.font.size = Pt(8.0)
        r_v.font.color.rgb = C_MUTED

# Diagramme centré grand format
frame_d12 = s12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(3.28), Inches(11.23), Inches(2.42))
frame_d12.adjustments[0] = 0.03
frame_d12.fill.solid()
frame_d12.fill.fore_color.rgb = C_BG_SLIDE
frame_d12.line.color.rgb = C_LINE
frame_d12.line.width = Pt(0.75)

add_fitted_picture(s12, "img/filiation_parcellaire.png", 1.25, 3.36, 10.83, 2.26, caption="Principe de filiation cadastrale lors d'un DMPC : remplacement de la parcelle mère A 14 par les parcelles filles A 115 et A 116")

# Bandeau bas réécrit proprement
b_syn12 = s12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.82), Inches(11.23), Inches(0.83))
b_syn12.adjustments[0] = 0.03
b_syn12.fill.solid()
b_syn12.fill.fore_color.rgb = C_CALLOUT_BG
b_syn12.line.color.rgb = C_ACCENT
b_syn12.line.width = Pt(1.0)

bx_syn12 = s12.shapes.add_textbox(Inches(1.20), Inches(5.88), Inches(10.93), Inches(0.70))
tf_syn12 = bx_syn12.text_frame
tf_syn12.word_wrap = True
tf_syn12.margin_left = tf_syn12.margin_top = tf_syn12.margin_right = tf_syn12.margin_bottom = 0
psyn1 = tf_syn12.paragraphs[0]
psyn1.text = "RÔLE CLÉ DANS L'OUTIL DÉVELOPPÉ :"
rsyn1 = psyn1.runs[0]
rsyn1.font.name = "Arial"
rsyn1.font.size = Pt(9.5)
rsyn1.font.bold = True
rsyn1.font.color.rgb = C_ACCENT

psyn2 = tf_syn12.add_paragraph()
psyn2.space_before = Pt(2)
rsyn2 = psyn2.add_run()
rsyn2.text = "Le traçage des divisions permet de rattacher automatiquement un acte ancien rédigé sous une ancienne parcelle mère aux parcelles filles actuelles, garantissant la découverte immédiate de l'antériorité."
rsyn2.font.name = "Arial"
rsyn2.font.size = Pt(8.5)
rsyn2.font.color.rgb = C_DARK

# ==================== CHAPITRE 3 : ÉTAT DE L'ART ====================
print("Génération Chapitre 3 (Slides 13-18)...")

# Slide 13 : Transition Chapitre 3
create_transition_slide(
    chap_idx=2,
    slide_num=13,
    title="Chapitre 3 : État de l'art des technologies d'analyse documentaire",
    subtitle="Reconnaissance OCR et HTR, extraction d'entités, modèles Vision-Langage et synthèse"
)

# Slide 14 (3.1) : OCR imprimé
s14, c14 = init_standard_slide(2, 0, 14, "3.1 Reconnaissance de texte imprimé : Les moteurs OCR classiques")

# 1. Figure Baek panoramique en haut (format large adapté, sans blanc parasite)
frame_top14 = s14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(1.55), Inches(11.23), Inches(1.65))
frame_top14.adjustments[0] = 0.03
frame_top14.fill.solid()
frame_top14.fill.fore_color.rgb = C_BG_SLIDE
frame_top14.line.color.rgb = C_LINE
frame_top14.line.width = Pt(0.75)
add_fitted_picture(s14, "img/baek_ocr_pipeline_tight.png", 1.15, 1.60, 11.03, 1.52, caption="Chaîne de reconnaissance de texte imprimé en quatre étapes (Baek et al., 2019)")

# 2. Explication détaillée des 4 étapes de la figure en 4 colonnes en dessous
steps_baek = [
    ("1. Transformation (TPS)",
     "Normalisation géométrique par spline (Thin Plate Spline) pour redresser le texte courbé ou incliné."),
    ("2. Extraction visuelle (CNN)",
     "Réseau convolutif (ResNet/VGG) extrayant les caractéristiques graphiques de chaque ligne de texte."),
    ("3. Séquence (BiLSTM)",
     "Réseau récurrent bidirectionnel modélisant les dépendances et le contexte entre les caractères."),
    ("4. Prédiction (CTC / Attn)",
     "Fonction d'alignement ou mécanisme d'attention décodant la séquence en caractères de texte exploitables.")
]

cw14 = 2.65
cg14 = 0.21
for bi, (b_hdr, b_txt) in enumerate(steps_baek):
    bx_pos = 1.05 + bi * (cw14 + cg14)
    b_box = s14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(bx_pos), Inches(3.35), Inches(cw14), Inches(2.32))
    b_box.adjustments[0] = 0.04
    b_box.fill.solid()
    b_box.fill.fore_color.rgb = C_BG_SLIDE
    b_box.line.color.rgb = C_LINE
    b_box.line.width = Pt(0.75)
    
    b_acc = s14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(bx_pos), Inches(3.35), Inches(cw14), Inches(0.06))
    b_acc.adjustments[0] = 0.20
    b_acc.fill.solid()
    b_acc.fill.fore_color.rgb = C_NAVY if bi < 2 else C_ACCENT
    b_acc.line.fill.background()
    
    bx = s14.shapes.add_textbox(Inches(bx_pos + 0.12), Inches(3.48), Inches(cw14 - 0.24), Inches(2.10))
    tf = bx.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    p1 = tf.paragraphs[0]
    p1.text = b_hdr
    r1 = p1.runs[0]
    r1.font.name = "Arial"
    r1.font.size = Pt(9.5)
    r1.font.bold = True
    r1.font.color.rgb = C_NAVY if bi < 2 else C_ACCENT
    
    p2 = tf.add_paragraph()
    p2.space_before = Pt(4)
    p2.text = b_txt
    r2 = p2.runs[0]
    r2.font.name = "Arial"
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = C_DARK

# 3. Callout bas de synthèse technique
b_eval14 = s14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.82), Inches(11.23), Inches(0.83))
b_eval14.adjustments[0] = 0.03
b_eval14.fill.solid()
b_eval14.fill.fore_color.rgb = C_CALLOUT_BG
b_eval14.line.color.rgb = C_ACCENT
b_eval14.line.width = Pt(1.0)

bx_ev14 = s14.shapes.add_textbox(Inches(1.20), Inches(5.88), Inches(10.93), Inches(0.70))
tf_ev14 = bx_ev14.text_frame
tf_ev14.word_wrap = True
tf_ev14.margin_left = tf_ev14.margin_top = tf_ev14.margin_right = tf_ev14.margin_bottom = 0

pe1 = tf_ev14.paragraphs[0]
pe1.text = "APPORT ET LIMITE MAJEURE DU TEXTE IMPRIMÉ :"
re1 = pe1.runs[0]
re1.font.name = "Arial"
re1.font.size = Pt(9.0)
re1.font.bold = True
re1.font.color.rgb = C_ACCENT

pe2 = tf_ev14.add_paragraph()
pe2.space_before = Pt(2)
pe2.text = "Tesseract v5 et EasyOCR (CRAFT) excellent sur les formulaires Cerfa dactylographiés récents. En revanche, le taux d'erreur dépasse 30% dès que le texte est manuscrit ou altéré."
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

# Slide 16 (3.3) : NER (Démonstration visuelle et comparatif)
s16, c16 = init_standard_slide(2, 2, 16, "3.3 Extraction d'entités nommées (NER) : GLiNER vs LayoutLM")

# 1. En haut : Démonstration CONCRÈTE du rôle du NER sur une phrase réelle
b_demo = s16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(1.55), Inches(11.23), Inches(1.50))
b_demo.adjustments[0] = 0.03
b_demo.fill.solid()
b_demo.fill.fore_color.rgb = C_BG_SLIDE
b_demo.line.color.rgb = C_LINE
b_demo.line.width = Pt(0.75)

bx_demo = s16.shapes.add_textbox(Inches(1.20), Inches(1.62), Inches(10.93), Inches(1.35))
tf_demo = bx_demo.text_frame
tf_demo.word_wrap = True
tf_demo.margin_left = tf_demo.margin_top = tf_demo.margin_right = tf_demo.margin_bottom = 0

pd_h = tf_demo.paragraphs[0]
pd_h.text = "RÔLE DU NER : EXTRACTION AUTOMATIQUE DES 5 ENTITÉS CIBLES DANS LE TEXTE BRUT TRANSCRIT"
rd_h = pd_h.runs[0]
rd_h.font.name = "Arial"
rd_h.font.size = Pt(9.5)
rd_h.font.bold = True
rd_h.font.color.rgb = C_NAVY

pd_ex = tf_demo.add_paragraph()
pd_ex.space_before = Pt(4)
rd_1 = pd_ex.add_run()
rd_1.text = "Texte brut issu de l'OCR : « Procès-verbal dressé le "
rd_1.font.size = Pt(9.0)
rd_1.font.color.rgb = C_DARK

# Tags colorés
rd_date = pd_ex.add_run()
rd_date.text = "[12 mai 1984 : DATE_ACTE] "
rd_date.font.size = Pt(9.0)
rd_date.font.bold = True
rd_date.font.color.rgb = C_ACCENT

rd_2 = pd_ex.add_run()
rd_2.text = "par le cabinet à "
rd_2.font.size = Pt(9.0)
rd_2.font.color.rgb = C_DARK

rd_com = pd_ex.add_run()
rd_com.text = "[Aubenas : COMMUNE] "
rd_com.font.size = Pt(9.0)
rd_com.font.bold = True
rd_com.font.color.rgb = C_SUCCESS

rd_3 = pd_ex.add_run()
rd_3.text = "pour la parcelle "
rd_3.font.size = Pt(9.0)
rd_3.font.color.rgb = C_DARK

rd_parc = pd_ex.add_run()
rd_parc.text = "[Section A n° 14 : PARCELLE] "
rd_parc.font.size = Pt(9.0)
rd_parc.font.bold = True
rd_parc.font.color.rgb = RGBColor(180, 83, 9)

rd_4 = pd_ex.add_run()
rd_4.text = "sous le dossier "
rd_4.font.size = Pt(9.0)
rd_4.font.color.rgb = C_DARK

rd_dos = pd_ex.add_run()
rd_dos.text = "[GEO-1984-089 : RÉFÉRENCE] »"
rd_dos.font.size = Pt(9.0)
rd_dos.font.bold = True
rd_dos.font.color.rgb = C_NAVY

# 2. En bas : Deux colonnes contrastées comparant les deux paradigmes
cw16 = 5.48
# Colonne 1 : GLiNER (Retenu)
b_gl = s16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(3.20), Inches(cw16), Inches(2.45))
b_gl.adjustments[0] = 0.03
b_gl.fill.solid()
b_gl.fill.fore_color.rgb = RGBColor(240, 253, 244)
b_gl.line.color.rgb = C_SUCCESS
b_gl.line.width = Pt(1.2)

bx_gl = s16.shapes.add_textbox(Inches(1.20), Inches(3.30), Inches(cw16 - 0.30), Inches(2.25))
tf_gl = bx_gl.text_frame
tf_gl.word_wrap = True
tf_gl.margin_left = tf_gl.margin_top = tf_gl.margin_right = tf_gl.margin_bottom = 0

pgl_h = tf_gl.paragraphs[0]
pgl_h.text = "✓ GLiNER : Modèle bi-encodeur zero-shot (RETENU)"
rgl_h = pgl_h.runs[0]
rgl_h.font.name = "Arial"
rgl_h.font.size = Pt(10.5)
rgl_h.font.bold = True
rgl_h.font.color.rgb = C_SUCCESS

gl_points = [
    ("Modèle compact", "Encodeur DeBERTa léger (~340M paramètres) exécutable très rapidement sur simple CPU."),
    ("Extraction zero-shot", "Identifie n'importe quelle entité arbitraire sans nécessiter des milliers d'exemples."),
    ("Tolérance au bruit OCR", "Compréhension contextuelle robuste capable de surmonter les coquilles de lecture."),
    ("Précision mesurée", "F1 = 0.88 sur les archives du cabinet avec attribution d'un score de confiance.")
]
for gh, gb in gl_points:
    p_g = tf_gl.add_paragraph()
    p_g.space_before = Pt(3)
    r1 = p_g.add_run()
    r1.text = "• " + gh + " : "
    r1.font.name = "Arial"
    r1.font.size = Pt(8.5)
    r1.font.bold = True
    r1.font.color.rgb = C_NAVY
    r2 = p_g.add_run()
    r2.text = gb
    r2.font.name = "Arial"
    r2.font.size = Pt(8.0)
    r2.font.color.rgb = C_DARK

# Colonne 2 : LayoutLMv3 (Écarté)
b_lm = s16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.80), Inches(3.20), Inches(cw16), Inches(2.45))
b_lm.adjustments[0] = 0.03
b_lm.fill.solid()
b_lm.fill.fore_color.rgb = C_BG_SLIDE
b_lm.line.color.rgb = C_BORDER_TAB
b_lm.line.width = Pt(0.75)

bx_lm = s16.shapes.add_textbox(Inches(6.95), Inches(3.30), Inches(cw16 - 0.30), Inches(2.25))
tf_lm = bx_lm.text_frame
tf_lm.word_wrap = True
tf_lm.margin_left = tf_lm.margin_top = tf_lm.margin_right = tf_lm.margin_bottom = 0

plm_h = tf_lm.paragraphs[0]
plm_h.text = "✗ LayoutLMv3 : Modèle multimodal 2D (ÉCARTÉ)"
rlm_h = plm_h.runs[0]
rlm_h.font.name = "Arial"
rlm_h.font.size = Pt(10.5)
rlm_h.font.bold = True
rlm_h.font.color.rgb = C_MUTED

lm_points = [
    ("Modèle multimodal lourd", "Combine image, boîtes 2D et texte, nécessitant ~8 Go de VRAM GPU dédiée."),
    ("Réentraînement supervisé", "Exige d'annoter à la main des centaines de formulaires spécifiques au cabinet."),
    ("Sensibilité extrême aux fautes", "Chute brutale de la précision dès que l'OCR initial commet des fautes de lecture."),
    ("Inopérant sur manuscrit", "Inadapté aux registres manuscrits et aux actes libres sans structure figée.")
]
for lh, lb in lm_points:
    p_l = tf_lm.add_paragraph()
    p_l.space_before = Pt(3)
    r1 = p_l.add_run()
    r1.text = "• " + lh + " : "
    r1.font.name = "Arial"
    r1.font.size = Pt(8.5)
    r1.font.bold = True
    r1.font.color.rgb = C_MUTED
    r2 = p_l.add_run()
    r2.text = lb
    r2.font.name = "Arial"
    r2.font.size = Pt(8.0)
    r2.font.color.rgb = C_DARK

# 3. Callout bas de synthèse
b_syn16 = s16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.82), Inches(11.23), Inches(0.83))
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
ps16_1.text = "JUSTIFICATION DU CHOIX DE GLiNER :"
rs16_1 = ps16_1.runs[0]
rs16_1.font.name = "Arial"
rs16_1.font.size = Pt(9.0)
rs16_1.font.bold = True
rs16_1.font.color.rgb = C_ACCENT

ps16_2 = tf_s16.add_paragraph()
ps16_2.space_before = Pt(2)
rs16_2 = ps16_2.add_run()
rs16_2.text = "GLiNER est retenu pour sa flexibilité zero-shot, son exécution ultra-rapide sur les postes de travail bureautiques sans GPU et sa résilience face aux coquilles de l'OCR."
rs16_2.font.name = "Arial"
rs16_2.font.size = Pt(8.5)
rs16_2.font.color.rgb = C_NAVY

# Slide 17 (3.4) : VLM (Cartes à gauche, schéma architecture à droite)
print("Génération Slide 17 (VLM local)...")
s17, c17 = init_standard_slide(2, 3, 17, "3.4 Modèles Vision-Langage (VLM) : Arbitrage multimodal local")
add_two_column_content(
    s17,
    items=[
        ("Modèles Vision-Langage (LLaVA / MiniCPM)",
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
)

# Slide 18 (3.5) : Synthèse comparative (Tableau 3.1 exact du mémoire avec sélections vérifiées)
s18, c18 = init_standard_slide(2, 4, 18, "3.5 Synthèse comparative des technologies d'analyse de document")

headers18 = ["Modèle", "Statut PFE", "Méthode utilisée", "Ressources", "Points forts", "Limites principales"]
# Sélections vérifiées rigoureusement avec la section 3.6 du mémoire
rows18 = [
    ("Tesseract v5", "Écarté", "Analyse par caractères", "CPU", "Rapide et économe.", "Précision très faible sur le manuscrit."),
    ("PyLaia / TrOCR", "RETENU (HTR)", "Réseau convolutif et Transformer", "~4 Go GPU", "Forte précision manuscrit (< 6% CER).", "Exige d'isoler les lignes au préalable."),
    ("YOLOv8", "RETENU (Vision)", "Détection d'objets spatiale", "< 2 Go GPU", "Découpage visuel ultra-rapide (0.15 s).", "Isole les zones sans lire le texte."),
    ("LayoutLMv3", "Écarté", "Modèle multimodal vision-texte", "~8 Go GPU", "Analyse conjointe forme et mots.", "Sensible aux fautes initiales de l'OCR."),
    ("GLiNER", "RETENU (NER)", "Comparaison de représentations", "~2 Go GPU", "Extraction zero-shot sans réentraînement.", "Dépend de la propreté du texte transmis."),
    ("Donut", "Écarté", "Décodage visuel direct", "~4 Go GPU", "Extraction sans OCR externe.", "Limité aux types de fichiers connus."),
    ("Qwen-VL", "Écarté", "VLM à résolution dynamique", "~8 Go GPU", "Lecture fine des petits détails.", "Temps de calcul élevé, API cloud externe."),
    ("LLaVA / MiniCPM", "RETENU (Arbitre)", "Modèle Vision-Langage local", "> 8 Go GPU", "Raisonnement visuel 100% local (Ollama).", "Inférence lente (4.8 s) et risque d'hallucination.")
]

cw18 = [1.60, 1.30, 2.30, 1.20, 2.40, 2.43]
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
    r.font.size = Pt(8.5)
    r.font.bold = True
    r.font.color.rgb = C_WHITE
    hx += cw18[hi]

for ri, (m_nom, m_stat, m_meth, m_ress, m_pts, m_lim) in enumerate(rows18):
    ry = ty18 + 0.38 + ri * rh18
    is_ret = "RETENU" in m_stat
    bg_r = RGBColor(240, 253, 244) if is_ret else (C_CARD_BG if ri % 2 == 0 else C_BG_SLIDE)
    
    row_vals = [m_nom, m_stat, m_meth, m_ress, m_pts, m_lim]
    rx = tx18
    for ci, cell in enumerate(row_vals):
        c_box = s18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(rx), Inches(ry), Inches(cw18[ci]), Inches(rh18))
        c_box.adjustments[0] = 0.08
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = bg_r
        c_box.line.color.rgb = C_SUCCESS if is_ret and ci == 1 else C_LINE
        c_box.line.width = Pt(1.0 if is_ret and ci == 1 else 0.5)
        
        bx = s18.shapes.add_textbox(Inches(rx + 0.04), Inches(ry + 0.03), Inches(cw18[ci] - 0.08), Inches(rh18 - 0.06))
        tf = bx.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = cell
        p.alignment = PP_ALIGN.CENTER if ci in (0, 1, 3) else PP_ALIGN.LEFT
        r = p.runs[0]
        r.font.name = "Arial"
        r.font.size = Pt(7.5)
        
        if ci == 0:
            r.font.bold = True
            r.font.color.rgb = C_NAVY
        elif ci == 1:
            r.font.bold = True
            r.font.color.rgb = C_SUCCESS if is_ret else C_MUTED
        else:
            r.font.color.rgb = C_DARK
            
        rx += cw18[ci]

# Grand bandeau bas : Les 4 méthodes retenues (conforme section 3.6 du mémoire)
sy18 = 5.45
s_box18 = s18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(tx18), Inches(sy18), Inches(sum(cw18)), Inches(1.15))
s_box18.adjustments[0] = 0.03
s_box18.fill.solid()
s_box18.fill.fore_color.rgb = C_CALLOUT_BG
s_box18.line.color.rgb = C_ACCENT
s_box18.line.width = Pt(1.0)

bx_syn18 = s18.shapes.add_textbox(Inches(tx18 + 0.15), Inches(sy18 + 0.08), Inches(sum(cw18) - 0.30), Inches(0.98))
tf_syn18 = bx_syn18.text_frame
tf_syn18.word_wrap = True
p_syn18 = tf_syn18.paragraphs[0]
p_syn18.text = "LES 4 MÉTHODES SÉLECTIONNÉES (SECTION 3.6 DU MÉMOIRE) :"
p_syn18.alignment = PP_ALIGN.CENTER
r_syn18 = p_syn18.runs[0]
r_syn18.font.name = "Arial"
r_syn18.font.size = Pt(9.5)
r_syn18.font.bold = True
r_syn18.font.color.rgb = C_ACCENT

p_syn18_2 = tf_syn18.add_paragraph()
p_syn18_2.space_before = Pt(3)
p_syn18_2.text = "1. Segmentation spatiale : YOLOv8  |  2. Extraction textuelle : EasyOCR (imprimé) + TrOCR / PyLaia (manuscrit)"
p_syn18_2.alignment = PP_ALIGN.CENTER
r_syn18_2 = p_syn18_2.runs[0]
r_syn18_2.font.name = "Arial"
r_syn18_2.font.size = Pt(8.5)
r_syn18_2.font.bold = True
r_syn18_2.font.color.rgb = C_NAVY

p_syn18_3 = tf_syn18.add_paragraph()
p_syn18_3.space_before = Pt(2)
p_syn18_3.text = "3. Repérage d'entités : GLiNER (zero-shot)  |  4. Arbitrage local de secours : LLaVA et MiniCPM-V via Ollama (100% local)."
p_syn18_3.alignment = PP_ALIGN.CENTER
r_syn18_3 = p_syn18_3.runs[0]
r_syn18_3.font.name = "Arial"
r_syn18_3.font.size = Pt(8.5)
r_syn18_3.font.bold = True
r_syn18_3.font.color.rgb = C_NAVY

# ==================== CHAPITRE 4 : ARCHITECTURE & PIPELINE ====================
print("Génération Chapitre 4 (Slides 19-23)...")

# Slide 19 : Transition Chapitre 4
create_transition_slide(
    chap_idx=3,
    slide_num=19,
    title="Chapitre 4 : Architecture logicielle et chaîne de traitement",
    subtitle="Chaîne modulaire, segmentation spatiale YOLOv8, extraction textuelle et arbitrage local"
)

# Slide 20 (4.1) : Architecture globale : Figure 4.1 du mémoire + Protection des données
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
r_flb.font.color.rgb = C_NAVY

# Slide 21 (4.2) : Segmentation YOLOv8
# FIGURE CONFORME : FIGURE 4.2 DU MÉMOIRE (img/r4p_fig2_tight.png), STRICTEMENT AUCUN CAS RÉEL DE PRADES EN AVANCE
print("Génération Slide 21 (Figure 4.2 authentique du mémoire r4p_fig2_tight)...")
s21, c21 = init_standard_slide(3, 1, 21, "4.2 Détection et segmentation spatiale par YOLOv8")
add_two_column_content(
    s21,
    items=[
        ("Segmentation spatiale ciblée",
         "Localisation directe des zones d'intérêt : cartouche Cerfa, tableau des contenances, mentions du géomètre et parcelles."),
        ("Suppression des encadrements doubles (NMS)",
         "L'algorithme Non-Maximum Suppression filtre les boîtes superposées pour ne conserver que la détection optimale."),
        ("Inférence locale ultra-rapide (< 0.2s)",
         "Le modèle pré-entraîné YOLOv8 nano traite les documents en quelques fractions de seconde sur simple CPU de bureau.")
    ],
    img_path="img/r4p_fig2_tight.png",
    img_caption="Détection des boîtes englobantes des cartouches et parcelles par YOLOv8 (Mémoire, Figure 4.2)",
    callout_title="RÔLE FONDAMENTAL DU DÉCOUPAGE",
    callout_text="La segmentation spatiale garantit que les moteurs de lecture ne reçoivent que des imagettes parfaitement cadrées sur les champs cibles."
)

# Slide 22 (4.3) : Extraction GLiNER
s22, c22 = init_standard_slide(3, 2, 22, "4.3 Transcription textuelle et extraction d'entités par GLiNER")
add_two_column_content(
    s22,
    items=[
        ("Transmission ciblée des imagettes",
         "Seules les régions utiles sont transmises aux moteurs de lecture, évitant les surcharges de contexte et les confusions visuelles."),
        ("Extraction zero-shot des 5 métadonnées",
         "GLiNER extrait dynamiquement : commune, date de l'acte, nature d'opération, référence interne et contenance avec indice de confiance."),
        ("Tolérance remarquable au bruit de lecture",
         "Grâce à ses représentations bidirectionnelles DeBERTa, GLiNER compense les fautes de frappe ou d'OCR sans réentraînement.")
    ],
    img_path="img/extraction_ocr_gliner.png",
    img_caption="Extraction et balisage sémantique des entités cibles par GLiNER",
    callout_title="PERFORMANCE D'EXTRACTION",
    callout_text="Le découpage spatial préalable permet de faire passer le score F1 global de l'extraction d'entités de 0.72 à 0.88 sur le corpus d'archives."
)

# Slide 23 (4.4) : Justification de la chaîne hybride et arbitrage VLM local
s23, c23 = init_standard_slide(3, 3, 23, "4.4 Arbitrage visuel local : Pourquoi une chaîne hybride plutôt qu'une IA globale ?")

# 1. En haut : Cas réel manuscrit et résultat certifié (Présentation épurée avec badges)
b_top_scan = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(1.55), Inches(5.48), Inches(1.40))
b_top_scan.adjustments[0] = 0.03
b_top_scan.fill.solid()
b_top_scan.fill.fore_color.rgb = C_BG_SLIDE
b_top_scan.line.color.rgb = C_LINE
b_top_scan.line.width = Pt(0.75)

if os.path.exists("img/crop_lachapelle_tight.png"):
    s23.shapes.add_picture("img/crop_lachapelle_tight.png", Inches(1.20), Inches(1.65), width=Inches(2.50), height=Inches(1.20))

bx_ts = s23.shapes.add_textbox(Inches(3.80), Inches(1.65), Inches(2.60), Inches(1.20))
tf_ts = bx_ts.text_frame
tf_ts.word_wrap = True
tf_ts.margin_left = tf_ts.margin_top = tf_ts.margin_right = tf_ts.margin_bottom = 0
pts1 = tf_ts.paragraphs[0]
pts1.text = "CAS RÉEL D'ARCHIVE :"
rts1 = pts1.runs[0]
rts1.font.name = "Arial"
rts1.font.size = Pt(8.5)
rts1.font.bold = True
rts1.font.color.rgb = C_NAVY
pts2 = tf_ts.add_paragraph()
pts2.space_before = Pt(2)
pts2.text = "« Lachapelle /s/ AUBENAS »"
rts2 = pts2.runs[0]
rts2.font.name = "Arial"
rts2.font.size = Pt(10.0)
rts2.font.bold = True
rts2.font.color.rgb = C_ACCENT
pts3 = tf_ts.add_paragraph()
pts3.space_before = Pt(2)
pts3.text = "« /s/ » = abréviation cursive locale pour « sous »."
rts3 = pts3.runs[0]
rts3.font.name = "Arial"
rts3.font.size = Pt(7.5)
rts3.font.color.rgb = C_MUTED

# Carte Résultat certifié en haut à droite
b_top_res = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.80), Inches(1.55), Inches(5.48), Inches(1.40))
b_top_res.adjustments[0] = 0.03
b_top_res.fill.solid()
b_top_res.fill.fore_color.rgb = RGBColor(240, 253, 244)
b_top_res.line.color.rgb = C_SUCCESS
b_top_res.line.width = Pt(1.2)

bx_tr = s23.shapes.add_textbox(Inches(7.00), Inches(1.68), Inches(5.10), Inches(1.15))
tf_tr = bx_tr.text_frame
tf_tr.word_wrap = True
tf_tr.margin_left = tf_tr.margin_top = tf_tr.margin_right = tf_tr.margin_bottom = 0
ptr1 = tf_tr.paragraphs[0]
ptr1.text = "RÉSULTAT APRÈS ARBITRAGE LOCAL SÉCURISÉ :"
rtr1 = ptr1.runs[0]
rtr1.font.name = "Arial"
rtr1.font.size = Pt(8.5)
rtr1.font.bold = True
rtr1.font.color.rgb = RGBColor(22, 101, 52)
ptr2 = tf_tr.add_paragraph()
ptr2.space_before = Pt(2)
ptr2.text = "🎯 « La Chapelle-sous-Aubenas »"
rtr2 = ptr2.runs[0]
rtr2.font.name = "Arial"
rtr2.font.size = Pt(13.0)
rtr2.font.bold = True
rtr2.font.color.rgb = RGBColor(21, 128, 61)
ptr3 = tf_tr.add_paragraph()
ptr3.space_before = Pt(2)
ptr3.text = "Code INSEE 07058 certifié  |  Zéro hallucination  |  100% local (Ollama)"
rtr3 = ptr3.runs[0]
rtr3.font.name = "Arial"
rtr3.font.size = Pt(8.0)
rtr3.font.bold = True
rtr3.font.color.rgb = C_NAVY

# 2. Au milieu : Duel direct par mots-clés percutants (allégé, zéro texte inutile)
# Colonne gauche : Risques de l'IA globale
b_d1 = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(3.08), Inches(5.48), Inches(2.60))
b_d1.adjustments[0] = 0.03
b_d1.fill.solid()
b_d1.fill.fore_color.rgb = C_BG_SLIDE
b_d1.line.color.rgb = C_WARN
b_d1.line.width = Pt(1.0)

bx_d1 = s23.shapes.add_textbox(Inches(1.25), Inches(3.18), Inches(5.10), Inches(2.40))
tf_d1 = bx_d1.text_frame
tf_d1.word_wrap = True
tf_d1.margin_left = tf_d1.margin_top = tf_d1.margin_right = tf_d1.margin_bottom = 0

pd1_h = tf_d1.paragraphs[0]
pd1_h.text = "PIÈGES D'UNE IA GÉNÉRATIVE GLOBALE (VLM SEUL)"
rd1_h = pd1_h.runs[0]
rd1_h.font.name = "Arial"
rd1_h.font.size = Pt(10.0)
rd1_h.font.bold = True
rd1_h.font.color.rgb = C_WARN

d1_items = [
    ("Base nationale confuse", "Hésite parmi les communes aux toponymes proches : Lachapelle-Graillouse (07), Lachapelle-sous-Chanéac (07), Vercors (26)..."),
    ("Lenteur prohibitive", "5 à 8 secondes par zone sur CPU standard (30x plus lent que l'OCR). Bloque le flux des 23 600 dossiers."),
    ("Risque d'hallucination", "Invention d'un toponyme crédible face à une graphie cursive dégradée.")
]
for dh, db in d1_items:
    p = tf_d1.add_paragraph()
    p.space_before = Pt(4)
    r1 = p.add_run()
    r1.text = "✗ " + dh + " : "
    r1.font.name = "Arial"
    r1.font.size = Pt(8.5)
    r1.font.bold = True
    r1.font.color.rgb = C_DARK
    r2 = p.add_run()
    r2.text = db
    r2.font.name = "Arial"
    r2.font.size = Pt(8.0)
    r2.font.color.rgb = C_MUTED

# Colonne droite : Chaîne hybride retenue
b_d2 = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.80), Inches(3.08), Inches(5.48), Inches(2.60))
b_d2.adjustments[0] = 0.03
b_d2.fill.solid()
b_d2.fill.fore_color.rgb = C_BG_SLIDE
b_d2.line.color.rgb = C_ACCENT
b_d2.line.width = Pt(1.0)

bx_d2 = s23.shapes.add_textbox(Inches(7.00), Inches(3.18), Inches(5.10), Inches(2.40))
tf_d2 = bx_d2.text_frame
tf_d2.word_wrap = True
tf_d2.margin_left = tf_d2.margin_top = tf_d2.margin_right = tf_d2.margin_bottom = 0

pd2_h = tf_d2.paragraphs[0]
pd2_h.text = "SOLUTION RETENUE : CHAÎNE HYBRIDE SÉCURISÉE"
rd2_h = pd2_h.runs[0]
rd2_h.font.name = "Arial"
rd2_h.font.size = Pt(10.0)
rd2_h.font.bold = True
rd2_h.font.color.rgb = C_ACCENT

d2_items = [
    ("OCR / HTR déterministes", "Transcription brute et rapide en 0.2 s sans invention de texte."),
    ("Verrouillage 335 communes 07", "Rapprochement Levenshtein strictement borné au secteur du cabinet. Élimination garantie de toute hallucination hors département."),
    ("Arbitre VLM local ciblé", "Convoqué uniquement si score de confiance < 0.65 (< 15% des cas), sa lecture restant verrouillée par la base cadastrale.")
]
for dh, db in d2_items:
    p = tf_d2.add_paragraph()
    p.space_before = Pt(4)
    r1 = p.add_run()
    r1.text = "✓ " + dh + " : "
    r1.font.name = "Arial"
    r1.font.size = Pt(8.5)
    r1.font.bold = True
    r1.font.color.rgb = C_NAVY
    r2 = p.add_run()
    r2.text = db
    r2.font.name = "Arial"
    r2.font.size = Pt(8.0)
    r2.font.color.rgb = C_DARK

# 3. Callout bas officiel
b_syn23 = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.82), Inches(11.23), Inches(0.83))
b_syn23.adjustments[0] = 0.03
b_syn23.fill.solid()
b_syn23.fill.fore_color.rgb = C_CALLOUT_BG
b_syn23.line.color.rgb = C_ACCENT
b_syn23.line.width = Pt(1.0)

bx_s23 = s23.shapes.add_textbox(Inches(1.20), Inches(5.88), Inches(10.93), Inches(0.70))
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
rs23_2.text = "L'association d'outils déterministes rapides et d'un référentiel cadastral local élimine tout risque d'hallucination toponymique tout en divisant par 30 le temps de traitement."
rs23_2.font.name = "Arial"
rs23_2.font.size = Pt(8.5)
rs23_2.font.color.rgb = C_DARK

# ==================== CHAPITRE 5 : FIABILISATION & INTERFACE ====================
print("Génération Chapitre 5 (Slides 24-27)...")

# Slide 24 : Transition Chapitre 5
create_transition_slide(
    chap_idx=4,
    slide_num=24,
    title="Chapitre 5 : Fiabilisation des données et interface opérateur",
    subtitle="Validation contradictoire humaine Streamlit, règles métier et contrôle spatial"
)

# Slide 25 (5.1) : Interface Streamlit (Mise en scène type logiciel / navigateur moderne)
s25, c25 = init_standard_slide(4, 0, 25, "5.1 Interface de relecture assistée Streamlit")

# Colonne gauche (6.40 in) : Conteneur moderne avec barre de titre "application web"
frame_browser = s25.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(1.55), Inches(6.40), Inches(4.15))
frame_browser.adjustments[0] = 0.02
frame_browser.fill.solid()
frame_browser.fill.fore_color.rgb = C_BG_SLIDE
frame_browser.line.color.rgb = C_LINE
frame_browser.line.width = Pt(0.75)

# Barre d'en-tête du navigateur mock
bar_top = s25.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(1.55), Inches(6.40), Inches(0.32))
bar_top.adjustments[0] = 0.08
bar_top.fill.solid()
bar_top.fill.fore_color.rgb = C_BG_TAB
bar_top.line.color.rgb = C_LINE
bar_top.line.width = Pt(0.75)

bx_url = s25.shapes.add_textbox(Inches(1.20), Inches(1.58), Inches(6.10), Inches(0.26))
tf_url = bx_url.text_frame
tf_url.margin_left = tf_url.margin_top = tf_url.margin_right = tf_url.margin_bottom = 0
pu = tf_url.paragraphs[0]
pu.text = "● ● ●   http://localhost:8501  •  Module de relecture et validation contradictoire GEO-SIAPP"
ru = pu.runs[0]
ru.font.name = "Arial"
ru.font.size = Pt(7.5)
ru.font.color.rgb = C_MUTED

# Image réelle Streamlit à l'intérieur
add_fitted_picture(s25, "img/interface_haut.jpg", 1.15, 1.92, 6.20, 3.65, caption="Interface de relecture contradictoire : double vue scan / formulaire avec jauges de confiance (Joyeuse)")

# Colonne droite (4.60 in) : 3 fonctionnalités clés sous forme de cartes d'action
features_ui = [
    ("Double affichage synchronisé",
     ["Scan original à gauche face au formulaire pré-rempli à droite pour contrôle oculaire immédiat."],
     C_NAVY),
    ("Jauges de confiance tricolores",
     ["🟢 Vert (> 0.85) : validé en 1 clic sans retouche",
      "🟠 Orange (0.65–0.85) : contrôle visuel suggéré",
      "🔴 Rouge (< 0.65) : reprise manuelle requise"],
     C_ACCENT),
    ("Validation et export en 1 clic",
     ["Enregistrement dans la base d'audit et génération directe du payload conforme à l'API Géofoncier."],
     C_SUCCESS)
]

card_h25 = 1.25
gap25 = 0.15
for fi, (f_title, f_lines, f_col) in enumerate(features_ui):
    fy = 1.55 + fi * (card_h25 + gap25)
    b_f = s25.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.68), Inches(fy), Inches(4.60), Inches(card_h25))
    b_f.adjustments[0] = 0.03
    b_f.fill.solid()
    b_f.fill.fore_color.rgb = C_BG_SLIDE
    b_f.line.color.rgb = f_col
    b_f.line.width = Pt(0.75)
    
    b_acc = s25.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.68), Inches(fy), Inches(0.08), Inches(card_h25))
    b_acc.adjustments[0] = 0.20
    b_acc.fill.solid()
    b_acc.fill.fore_color.rgb = f_col
    b_acc.line.fill.background()
    
    bx_f = s25.shapes.add_textbox(Inches(7.86), Inches(fy + 0.08), Inches(4.35), Inches(card_h25 - 0.16))
    tf_f = bx_f.text_frame
    tf_f.word_wrap = True
    tf_f.margin_left = tf_f.margin_top = tf_f.margin_right = tf_f.margin_bottom = 0
    
    pf1 = tf_f.paragraphs[0]
    pf1.text = f_title
    rf1 = pf1.runs[0]
    rf1.font.name = "Arial"
    rf1.font.size = Pt(10.0)
    rf1.font.bold = True
    rf1.font.color.rgb = f_col
    
    for l_idx, f_line in enumerate(f_lines):
        pf2 = tf_f.add_paragraph()
        pf2.space_before = Pt(2.5 if len(f_lines) > 1 else 3.5)
        rf2 = pf2.add_run()
        rf2.text = f_line
        rf2.font.name = "Arial"
        rf2.font.size = Pt(8.0)
        rf2.font.color.rgb = C_DARK

# Callout bas de synthèse opérationnelle
b_syn25 = s25.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.82), Inches(11.23), Inches(0.83))
b_syn25.adjustments[0] = 0.03
b_syn25.fill.solid()
b_syn25.fill.fore_color.rgb = C_CALLOUT_BG
b_syn25.line.color.rgb = C_ACCENT
b_syn25.line.width = Pt(1.0)

bx_s25 = s25.shapes.add_textbox(Inches(1.20), Inches(5.88), Inches(10.93), Inches(0.70))
tf_s25 = bx_s25.text_frame
tf_s25.word_wrap = True
tf_s25.margin_left = tf_s25.margin_top = tf_s25.margin_right = tf_s25.margin_bottom = 0
ps25_1 = tf_s25.paragraphs[0]
ps25_1.text = "LE GÉOMÈTRE-EXPERT CONSERVE LE CONTRÔLE DÉCISIONNEL ABSOLU :"
rs25_1 = ps25_1.runs[0]
rs25_1.font.name = "Arial"
rs25_1.font.size = Pt(9.5)
rs25_1.font.bold = True
rs25_1.font.color.rgb = C_ACCENT

ps25_2 = tf_s25.add_paragraph()
ps25_2.space_before = Pt(2)
rs25_2 = ps25_2.add_run()
rs25_2.text = "Principe de « l'humain dans la boucle » : l'algorithme pré-remplit les champs et détecte les anomalies, mais le technicien valide chaque dossier. Le temps de saisie est divisé par 10 (1 à 2 min par dossier)."
rs25_2.font.name = "Arial"
rs25_2.font.size = Pt(8.5)
rs25_2.font.color.rgb = C_DARK

# Slide 26 (5.2) : Règles métier (Version épurée, percutante par puces courtes)
s26, c26 = init_standard_slide(4, 1, 26, "5.2 Moteur de 17 règles de cohérence métier et répertoires")

rule_groups_clean = [
    {
        "tag": "GROUPE 1 • 5 RÈGLES",
        "title": "Cadastre & Toponymie",
        "color": C_NAVY,
        "rules": [
            ("Communes INSEE", "Rapprochement Levenshtein sur 335 communes 07."),
            ("Section cadastrale", "Existence avérée au PCI Vecteur contemporain."),
            ("Numéros de parcelles", "Nettoyage typographique et exclusion du zéro."),
            ("Rattachement minimal", "Présence obligatoire d'au moins 1 parcelle mère/fille.")
        ]
    },
    {
        "tag": "GROUPE 2 • 4 RÈGLES",
        "title": "Cohérence temporelle",
        "color": C_ACCENT,
        "rules": [
            ("Plage historique", "Cohérence avec le fonds documentaire (1950–2007)."),
            ("Exercice géomètre", "Concordance date d'acte et activité du signataire."),
            ("Format ISO 8601", "Dates réelles et gestion des années bissextiles."),
            ("Chronologie", "Date d'enregistrement postérieure à celle du plan.")
        ]
    },
    {
        "tag": "GROUPE 3 • 4 RÈGLES",
        "title": "Intervenants & Parties",
        "color": RGBColor(180, 83, 9),
        "rules": [
            ("Annuaire cabinet", "Validation du géomètre dans la liste certifiée."),
            ("Donneur d'ordre", "Identification du client demandeur et des riverains."),
            ("Nettoyage textuel", "Suppression des mentions « consorts », « veuve », « id. »."),
            ("Alerte identité", "Signalement si anciens propriétaires = acquéreurs.")
        ]
    },
    {
        "tag": "GROUPE 4 • 4 RÈGLES",
        "title": "Dossier & Anti-doublon",
        "color": C_SUCCESS,
        "rules": [
            ("Masque de référence", "Syntaxe spécifique selon le fonds racheté."),
            ("Unicité stricte", "Détection immédiate de doublon au registre."),
            ("Conflits spatiaux", "Alerte si plusieurs actes sur même parcelle même année."),
            ("Intégrité PDF", "Vérification de présence et lisibilité du scan source.")
        ]
    }
]

grid_w = 5.48
grid_h = 2.05
col_xs = [1.05, 6.80]
row_ys = [1.55, 3.72]

for gi, gdata in enumerate(rule_groups_clean):
    gx = col_xs[gi % 2]
    gy = row_ys[gi // 2]
    
    b_g = s26.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(gx), Inches(gy), Inches(grid_w), Inches(grid_h))
    b_g.adjustments[0] = 0.03
    b_g.fill.solid()
    b_g.fill.fore_color.rgb = C_BG_SLIDE
    b_g.line.color.rgb = C_LINE
    b_g.line.width = Pt(0.75)
    
    b_acc = s26.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(gx), Inches(gy), Inches(0.08), Inches(grid_h))
    b_acc.adjustments[0] = 0.20
    b_acc.fill.solid()
    b_acc.fill.fore_color.rgb = gdata["color"]
    b_acc.line.fill.background()
    
    bx_g = s26.shapes.add_textbox(Inches(gx + 0.18), Inches(gy + 0.08), Inches(grid_w - 0.28), Inches(grid_h - 0.16))
    tf_g = bx_g.text_frame
    tf_g.word_wrap = True
    tf_g.margin_left = tf_g.margin_top = tf_g.margin_right = tf_g.margin_bottom = 0
    
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
    r_tit.font.size = Pt(10.0)
    r_tit.font.bold = True
    r_tit.font.color.rgb = C_NAVY
    
    for rh, rb in gdata["rules"]:
        p_r = tf_g.add_paragraph()
        p_r.space_before = Pt(2.5)
        rr_h = p_r.add_run()
        rr_h.text = "✓ " + rh + " : "
        rr_h.font.name = "Arial"
        rr_h.font.size = Pt(8.0)
        rr_h.font.bold = True
        rr_h.font.color.rgb = C_DARK
        
        rr_b = p_r.add_run()
        rr_b.text = rb
        rr_b.font.name = "Arial"
        rr_b.font.size = Pt(8.0)
        rr_b.font.color.rgb = C_MUTED

# Bandeau de synthèse officiel
b_syn26 = s26.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.82), Inches(11.23), Inches(0.83))
b_syn26.adjustments[0] = 0.03
b_syn26.fill.solid()
b_syn26.fill.fore_color.rgb = C_CALLOUT_BG
b_syn26.line.color.rgb = C_ACCENT
b_syn26.line.width = Pt(1.0)

bx_s26 = s26.shapes.add_textbox(Inches(1.20), Inches(5.88), Inches(10.93), Inches(0.70))
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
rs26_2.text = "Exécution en moins de 0.1 s. Tout dossier non conforme reçoit une pastille d'alerte et est soumis à l'arbitrage immédiat du géomètre avant tout envoi."
rs26_2.font.name = "Arial"
rs26_2.font.size = Pt(8.5)
rs26_2.font.color.rgb = C_DARK

# Slide 27 (5.3) : Contrôle spatial Folium
s27, c27 = init_standard_slide(4, 2, 27, "5.3 Contrôle cartographique interactif (Folium)")
add_two_column_content(
    s27,
    items=[
        ("Géolocalisation automatique sur fond OpenStreetMap et IGN",
         "Positionnement instantané de la pastille sur fond cadastral pour valider la cohérence spatiale de la commune et de la section."),
        ("Superposition des filiations de parcelles en violet",
         "Affichage automatique des nouvelles parcelles filles issues des divisions successives pour garantir le bon calage de l'acte ancien."),
        ("Vérification visuelle de proximité foncière",
         "Visualisation immédiate des autres pastilles Géofoncier existantes à proximité pour interdire la création de tout doublon.")
    ],
    img_path="img/Etape_2_localisation_cartographie.jpg",
    img_caption="Carte interactive Folium : calage cadastral IGN, parcelles filles en violet et pastilles OGE",
    callout_title="VALIDATION SPATIALE COMPLÈTE",
    callout_text="L'intégration cartographique garantit que les coordonnées Lambert-93 calculées correspondent parfaitement à la commune déclarée."
)


# ==================== CHAPITRE 6 : INTÉGRATION GÉOFONCIER ====================
print("Génération Chapitre 6 (Slides 28-31)...")

# Slide 28 : Transition Chapitre 6
create_transition_slide(
    chap_idx=5,
    slide_num=28,
    title="Chapitre 6 : Intégration Géofoncier et expérimentation",
    subtitle="Démonstration en conditions réelles, protocole d'injection et expérimentation Prades"
)

# Slide 29 (6.1) : Vidéo Démonstration (INTOUCHÉE)
print("Génération Slide 29 (Vidéo Démonstration 180s - STRICTEMENT INTOUCHÉE)...")
s29, c29 = init_standard_slide(5, 0, 29, "6.1 Démonstration du pipeline en conditions réelles")

vid_path = "Demo_soutenance_final.mp4"
poster_img = "video_poster.jpg" if os.path.exists("video_poster.jpg") else None
vw, vh = 11.40, 5.30
vx = 0.70 + (11.93 - vw) / 2.0
vy = 1.35 + (5.80 - vh) / 2.0

if os.path.exists(vid_path):
    try:
        s29.shapes.add_movie(vid_path, Inches(vx), Inches(vy), Inches(vw), Inches(vh), poster_frame_image=poster_img, mime_type="video/mp4")
        print(f"  Vidéo intégrée dans le conteneur blanc : {vid_path}")
    except Exception as e:
        print(f"  Erreur add_movie : {e}")
        add_fitted_picture(s29, poster_img, vx, vy, vw, vh, caption=f"Vidéo prête pour lecture : {vid_path}")
else:
    add_fitted_picture(s29, poster_img, vx, vy, vw, vh, caption="Fichier vidéo non trouvé")

# Slide 30 (6.2) : Protocole API (Pipeline horizontal moderne en 4 étapes)
s30, c30 = init_standard_slide(5, 1, 30, "6.2 Protocole de versement via l'API REST Géofoncier")

api_pipeline = [
    ("1. AUTHENTIFICATION",
     "POST /token",
     C_NAVY,
     "Certificats OGE du cabinet",
     "Jeton JWT Bearer valide 24h"),
    ("2. GÉORÉFÉRENCEMENT",
     "EPSG:2154 (Lambert-93)",
     C_ACCENT,
     "Calcul automatique centroïde",
     "Coordonnées X / Y réglementaires"),
    ("3. INJECTION MÉTADONNÉES",
     "POST /rfuoge (JSON)",
     C_NAVY,
     "5 attributs cibles vérifiés",
     "Code retour HTTP 201 Created"),
    ("4. TÉLÉVERSEMENT PIÈCES",
     "POST /dossiersoge (PDF)",
     C_SUCCESS,
     "Upload multipart Plan + PV",
     "Pastille RFU active sur Géofoncier")
]

cw30 = 2.65
cg30 = 0.21
for pi, (p_step, p_route, p_col, p_line1, p_line2) in enumerate(api_pipeline):
    px = 1.05 + pi * (cw30 + cg30)
    b_p = s30.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(px), Inches(1.60), Inches(cw30), Inches(2.45))
    b_p.adjustments[0] = 0.04
    b_p.fill.solid()
    b_p.fill.fore_color.rgb = C_BG_SLIDE
    b_p.line.color.rgb = p_col
    b_p.line.width = Pt(1.0)
    
    b_acc = s30.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(px), Inches(1.60), Inches(cw30), Inches(0.08))
    b_acc.adjustments[0] = 0.20
    b_acc.fill.solid()
    b_acc.fill.fore_color.rgb = p_col
    b_acc.line.fill.background()
    
    bx_p = s30.shapes.add_textbox(Inches(px + 0.12), Inches(1.75), Inches(cw30 - 0.24), Inches(2.20))
    tf_p = bx_p.text_frame
    tf_p.word_wrap = True
    tf_p.margin_left = tf_p.margin_top = tf_p.margin_right = tf_p.margin_bottom = 0
    
    pp1 = tf_p.paragraphs[0]
    pp1.text = p_step
    rp1 = pp1.runs[0]
    rp1.font.name = "Arial"
    rp1.font.size = Pt(9.0)
    rp1.font.bold = True
    rp1.font.color.rgb = p_col
    
    pp2 = tf_p.add_paragraph()
    pp2.space_before = Pt(4)
    pp2.text = p_route
    rp2 = pp2.runs[0]
    rp2.font.name = "Arial"
    rp2.font.size = Pt(10.5)
    rp2.font.bold = True
    rp2.font.color.rgb = C_NAVY
    
    pp3 = tf_p.add_paragraph()
    pp3.space_before = Pt(8)
    pp3.text = "• " + p_line1
    rp3 = pp3.runs[0]
    rp3.font.name = "Arial"
    rp3.font.size = Pt(8.5)
    rp3.font.bold = True
    rp3.font.color.rgb = C_DARK
    
    pp4 = tf_p.add_paragraph()
    pp4.space_before = Pt(4)
    pp4.text = "• " + p_line2
    rp4 = pp4.runs[0]
    rp4.font.name = "Arial"
    rp4.font.size = Pt(8.0)
    rp4.font.color.rgb = C_MUTED
    
    # Flèche de liaison entre étapes
    if pi < 3:
        bx_arrow = s30.shapes.add_textbox(Inches(px + cw30), Inches(2.65), Inches(cg30), Inches(0.35))
        tfa = bx_arrow.text_frame
        tfa.margin_left = tfa.margin_top = tfa.margin_right = tfa.margin_bottom = 0
        pa = tfa.paragraphs[0]
        pa.text = "→"
        pa.alignment = PP_ALIGN.CENTER
        ra = pa.runs[0]
        ra.font.name = "Arial"
        ra.font.size = Pt(16.0)
        ra.font.bold = True
        ra.font.color.rgb = C_ACCENT

# Terminal mock de validation serveur au milieu bas
b_term = s30.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(4.20), Inches(11.23), Inches(1.48))
b_term.adjustments[0] = 0.03
b_term.fill.solid()
b_term.fill.fore_color.rgb = RGBColor(15, 23, 42)  # Fond sombre type console
b_term.line.color.rgb = C_NAVY
b_term.line.width = Pt(1.0)

bx_term = s30.shapes.add_textbox(Inches(1.25), Inches(4.30), Inches(10.83), Inches(1.30))
tf_term = bx_term.text_frame
tf_term.word_wrap = True
tf_term.margin_left = tf_term.margin_top = tf_term.margin_right = tf_term.margin_bottom = 0

pt1 = tf_term.paragraphs[0]
pt1.text = "RÉPONSE DU SERVEUR GÉOFONCIER (HTTP 201 CREATED) :"
rt1 = pt1.runs[0]
rt1.font.name = "Arial"
rt1.font.size = Pt(8.5)
rt1.font.bold = True
rt1.font.color.rgb = RGBColor(56, 189, 248)  # Bleu clair console

pt2 = tf_term.add_paragraph()
pt2.space_before = Pt(3)
pt2.text = '{ "status": 201, "rfu_id": "GEO-07019-1984-089", "geom": "POINT(801452.35 6391204.18)", "message": "Pastille publiée avec succès sur Géofoncier" }'
rt2 = pt2.runs[0]
rt2.font.name = "Courier New"
rt2.font.size = Pt(9.0)
rt2.font.bold = True
rt2.font.color.rgb = RGBColor(74, 222, 128)  # Vert terminal

pt3 = tf_term.add_paragraph()
pt3.space_before = Pt(3)
pt3.text = "Journalisation locale certifiée dans suivi_publications.csv • Traçabilité intégrale de l'acte et des pièces jointes."
rt3 = pt3.runs[0]
rt3.font.name = "Arial"
rt3.font.size = Pt(8.0)
rt3.font.color.rgb = RGBColor(148, 163, 184)

# Callout bas de synthèse
b_syn30 = s30.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(5.82), Inches(11.23), Inches(0.83))
b_syn30.adjustments[0] = 0.03
b_syn30.fill.solid()
b_syn30.fill.fore_color.rgb = C_CALLOUT_BG
b_syn30.line.color.rgb = C_ACCENT
b_syn30.line.width = Pt(1.0)

bx_s30 = s30.shapes.add_textbox(Inches(1.20), Inches(5.88), Inches(10.93), Inches(0.70))
tf_s30 = bx_s30.text_frame
tf_s30.word_wrap = True
tf_s30.margin_left = tf_s30.margin_top = tf_s30.margin_right = tf_s30.margin_bottom = 0
ps30_1 = tf_s30.paragraphs[0]
ps30_1.text = "VALIDATION OFFICIELLE DU VERSEMENT SUR GÉOFONCIER :"
rs30_1 = ps30_1.runs[0]
rs30_1.font.name = "Arial"
rs30_1.font.size = Pt(9.5)
rs30_1.font.bold = True
rs30_1.font.color.rgb = C_ACCENT

ps30_2 = tf_s30.add_paragraph()
ps30_2.space_before = Pt(2)
rs30_2 = ps30_2.add_run()
rs30_2.text = "Le code retour HTTP 201 Created atteste de la création instantanée de la pastille sur la carte nationale de Géofoncier, rendant l'antériorité consultable par toute la profession."
rs30_2.font.name = "Arial"
rs30_2.font.size = Pt(8.5)
rs30_2.font.color.rgb = C_DARK

# Slide 31 (6.3) : Validation Prades
print("Génération Slide 31 (Résultats Prades)...")
s31, c31 = init_standard_slide(5, 2, 31, "6.3 Validation expérimentale sur la commune de Prades (Ardèche)")

left_x31 = 1.05
left_w31 = 5.50

kpi_tiles = [
    ("50 Dossiers", "Lot d'archives réelles testé", C_NAVY),
    ("100% Succès", "Concordance cadastrale totale", C_SUCCESS),
    ("F1 = 0.94", "Score global pondéré", C_ACCENT),
    ("x6 Plus Rapide", "Temps divisé par six", C_WARN)
]

kw31 = (left_w31 - 0.15) / 2.0
kh31 = 0.95

for ti, (t_val, t_desc, t_col) in enumerate(kpi_tiles):
    row_t = ti // 2
    col_t = ti % 2
    tx = left_x31 + col_t * (kw31 + 0.15)
    ty = 1.60 + row_t * (kh31 + 0.12)
    
    b_tile = s31.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(tx), Inches(ty), Inches(kw31), Inches(kh31))
    b_tile.adjustments[0] = 0.05
    b_tile.fill.solid()
    b_tile.fill.fore_color.rgb = C_BG_SLIDE
    b_tile.line.color.rgb = t_col
    b_tile.line.width = Pt(1.2)
    
    bx_t = s31.shapes.add_textbox(Inches(tx + 0.08), Inches(ty + 0.10), Inches(kw31 - 0.16), Inches(kh31 - 0.20))
    tf_t = bx_t.text_frame
    tf_t.word_wrap = True
    tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
    pt1 = tf_t.paragraphs[0]
    pt1.text = t_val
    rt1 = pt1.runs[0]
    rt1.font.name = "Arial"
    rt1.font.size = Pt(14.0)
    rt1.font.bold = True
    rt1.font.color.rgb = t_col
    
    pt2 = tf_t.add_paragraph()
    pt2.space_before = Pt(2)
    rt2 = pt2.add_run()
    rt2.text = t_desc
    rt2.font.name = "Arial"
    rt2.font.size = Pt(8.0)
    rt2.font.color.rgb = C_MUTED

b_det31 = s31.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left_x31), Inches(3.85), Inches(left_w31), Inches(2.80))
b_det31.adjustments[0] = 0.03
b_det31.fill.solid()
b_det31.fill.fore_color.rgb = C_BG_SLIDE
b_det31.line.color.rgb = C_LINE
b_det31.line.width = Pt(0.75)

bx_det31 = s31.shapes.add_textbox(Inches(left_x31 + 0.16), Inches(3.95), Inches(left_w31 - 0.32), Inches(2.60))
tf_det31 = bx_det31.text_frame
tf_det31.word_wrap = True
tf_det31.margin_left = tf_det31.margin_top = tf_det31.margin_right = tf_det31.margin_bottom = 0

p_dh = tf_det31.paragraphs[0]
p_dh.text = "Bilan de l'expérimentation sur Prades :\n"
r_dh = p_dh.runs[0]
r_dh.font.name = "Arial"
r_dh.font.size = Pt(10.5)
r_dh.font.bold = True
r_dh.font.color.rgb = C_NAVY

bullets_prades = [
    ("47 dossiers versés sans retouche", "Les métadonnées ont été reconnues avec une certitude absolue et validées en 1 clic."),
    ("3 dossiers corrigés visuellement", "Anomalies mineures de détection dues à des plis marqués sur les calques anciens."),
    ("Zéro fausse pastille sur Géofoncier", "Le filtre des 17 règles métier a bloqué toute injection erronée vers le portail national.")
]
for bh, bd in bullets_prades:
    p_b1 = tf_det31.add_paragraph()
    p_b1.space_before = Pt(5)
    r1 = p_b1.add_run()
    r1.text = "• " + bh + " : "
    r1.font.name = "Arial"
    r1.font.size = Pt(9.0)
    r1.font.bold = True
    r1.font.color.rgb = C_ACCENT
    r2 = p_b1.add_run()
    r2.text = bd
    r2.font.name = "Arial"
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = C_DARK

rx31 = 6.78
rw31 = 5.50

frame_img31 = s31.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(rx31), Inches(1.60), Inches(rw31), Inches(3.35))
frame_img31.adjustments[0] = 0.03
frame_img31.fill.solid()
frame_img31.fill.fore_color.rgb = C_BG_SLIDE
frame_img31.line.color.rgb = C_LINE
frame_img31.line.width = Pt(0.75)
add_fitted_picture(s31, "img/fig_geofoncier_prades.png", rx31 + 0.10, 1.70, rw31 - 0.20, 3.15, caption="Visualisation des pastilles publiées sur la commune de Prades (GéofoncierEXPERT)")

b_f1_31 = s31.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(rx31), Inches(5.10), Inches(rw31), Inches(1.55))
b_f1_31.adjustments[0] = 0.03
b_f1_31.fill.solid()
b_f1_31.fill.fore_color.rgb = C_CALLOUT_BG
b_f1_31.line.color.rgb = C_ACCENT
b_f1_31.line.width = Pt(1.0)

bx_f1_31 = s31.shapes.add_textbox(Inches(rx31 + 0.16), Inches(5.18), Inches(rw31 - 0.32), Inches(1.40))
tf_f1_31 = bx_f1_31.text_frame
tf_f1_31.word_wrap = True
tf_f1_31.margin_left = tf_f1_31.margin_top = tf_f1_31.margin_right = tf_f1_31.margin_bottom = 0

pf_h = tf_f1_31.paragraphs[0]
pf_h.text = "Scores F1 par champ obligatoire (lot de Prades) :"
rf_h = pf_h.runs[0]
rf_h.font.name = "Arial"
rf_h.font.size = Pt(9.0)
rf_h.font.bold = True
rf_h.font.color.rgb = C_NAVY

f1_items = [
    ("Commune (code INSEE COG)", "F1 = 0.98"),
    ("Date de l'acte (ISO 8601)", "F1 = 0.95"),
    ("Filiation parcellaire", "F1 = 0.94"),
    ("Référence interne cabinet", "F1 = 0.92")
]
for f_lbl, f_sc in f1_items:
    pf_row = tf_f1_31.add_paragraph()
    pf_row.space_before = Pt(2)
    r1 = pf_row.add_run()
    r1.text = "✓ " + f_lbl + " : "
    r1.font.name = "Arial"
    r1.font.size = Pt(8.0)
    r1.font.bold = True
    r1.font.color.rgb = C_DARK
    r2 = pf_row.add_run()
    r2.text = f_sc
    r2.font.name = "Arial"
    r2.font.size = Pt(8.0)
    r2.font.bold = True
    r2.font.color.rgb = C_SUCCESS


# ==================== CHAPITRE 7 : LIMITES & PERSPECTIVES ====================
print("Génération Chapitre 7 (Slides 32-34)...")

# Slide 32 : Transition Chapitre 7
create_transition_slide(
    chap_idx=6,
    slide_num=32,
    title="Chapitre 7 : Analyse des limites et perspectives d'évolution",
    subtitle="Diagnostic des modes de défaillance, altérations physiques et feuille de route technique"
)

# Slide 33 (7.1) : Limites techniques
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

# Slide 34 (7.2) : Perspectives d'évolution
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

# ==================== CHAPITRE 8 : CONCLUSION ====================
print("Génération Chapitre 8 (Slides 35-37)...")

# Slide 35 : Transition Chapitre 8
create_transition_slide(
    chap_idx=7,
    slide_num=35,
    title="Chapitre 8 : Conclusion générale et bilan du projet",
    subtitle="Bilan opérationnel pour GEO-SIAPP, traitement local et compétences d'ingénieur INSA"
)

# Slide 36 (8.1) : Bilan général
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
ppl_h.text = "Apports opérationnels pour le cabinet GEO-SIAPP\n"
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
ppr_h.text = "Rigueur déontologique et posture d'ingénieur INSA\n"
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

# Slide 37 (8.2) : Remerciements
print("Génération Slide 37 (Remerciements)...")
s37, c37 = init_standard_slide(7, 1, 37, "8.2 Remerciements et ouverture vers les échanges")

bx_rm = s37.shapes.add_textbox(Inches(1.50), Inches(1.80), Inches(10.33), Inches(0.70))
tf_rm = bx_rm.text_frame
p_rm = tf_rm.paragraphs[0]
p_rm.alignment = PP_ALIGN.CENTER
p_rm.text = "Merci pour votre attention"
r_rm = p_rm.runs[0]
r_rm.font.name = "Arial"
r_rm.font.size = Pt(32.0)
r_rm.font.bold = True
r_rm.font.color.rgb = C_NAVY

bx_rq = s37.shapes.add_textbox(Inches(1.50), Inches(2.55), Inches(10.33), Inches(0.40))
tf_rq = bx_rq.text_frame
p_rq = tf_rq.paragraphs[0]
p_rq.alignment = PP_ALIGN.CENTER
p_rq.text = "Je me tiens à votre entière disposition pour répondre à l'ensemble de vos questions."
r_rq = p_rq.runs[0]
r_rq.font.name = "Arial"
r_rq.font.size = Pt(14.0)
r_rq.font.color.rgb = C_ACCENT

thanks_lines = [
    "M. Gaëtan HAGUE : Tuteur entreprise, Géomètre-Expert associé, Cabinet GEO-SIAPP",
    "M. Mathieu KOEHL : Directeur du PFE, Enseignant-chercheur, INSA Strasbourg / ICube",
    "L'équipe du Cabinet GEO-SIAPP d'Aubenas pour leur accueil et leur accompagnement technique",
    "L'équipe pédagogique de la spécialité Topographie de l'INSA Strasbourg"
]

for ti, ttext in enumerate(thanks_lines):
    bx_th = s37.shapes.add_textbox(Inches(1.50), Inches(3.45 + ti * 0.45), Inches(10.33), Inches(0.35))
    tf_th = bx_th.text_frame
    p_th = tf_th.paragraphs[0]
    p_th.alignment = PP_ALIGN.CENTER
    p_th.text = ttext
    r_th = p_th.runs[0]
    r_th.font.name = "Arial"
    r_th.font.size = Pt(11.0)
    r_th.font.color.rgb = C_DARK

if os.path.exists("Logo_INSAStrasbourg.jpg"):
    s37.shapes.add_picture("Logo_INSAStrasbourg.jpg", Inches(3.80), Inches(5.55), width=Inches(2.50))
if os.path.exists("Geosiapp.jpg"):
    add_fitted_picture(s37, "Geosiapp.jpg", 7.60, 5.40, 2.00, 1.20)


# ==================== ENREGISTREMENT ====================

# ==================== NOTES DU PRÉSENTATEUR POUR L'ORAL (37 DIAPOSITIVES) ====================
# Discours direct mot à mot avec timing de parole par diapositive (Total : 20 min)
SPEAKER_NOTES = {
    1: (
        "[Temps : 30 s | Chrono : 00:30]\n"
        "Bonjour à tous, membres du jury. Je vous présente aujourd'hui mon projet de fin d'études d'ingénieur topographe à l'INSA Strasbourg, intitulé : "
        "Développement d'un outil permettant le traitement et l'insertion des archives numériques sur Géofoncier. "
        "Ce travail a été réalisé au sein du cabinet GEO-SIAPP sous le tutorat de Monsieur Gaëtan Hague, géomètre-expert associé, et la direction de Monsieur Mathieu Koehl."
    ),
    2: (
        "[Temps : 30 s | Chrono : 01:00]\n"
        "Notre présentation s'articule en 8 chapitres. Nous partirons du cadre réglementaire du cabinet et de la problématique foncière, avant d'aborder l'état de l'art des technologies de lecture. "
        "Nous détaillerons ensuite l'architecture développée, le moteur de fiabilisation et l'expérimentation réelle sur le terrain, pour terminer par l'analyse des limites et le bilan pour le cabinet."
    ),
    3: (
        "[Temps : 5 s | Chrono : 01:05]\n"
        "Commençons par le premier chapitre avec la présentation de la structure d'accueil et du contexte métier."
    ),
    4: "[Temps : 30 s | Chrono : 01:35]\\n• **Cabinet GEO-SIAPP** : structure historique d'Ardèche et Drôme, **5 géomètres-experts associés**, 4 agences.\\n• **Siège d'Aubenas** : conservation de plus de **23 600 dossiers physiques** (1959–2007).\\n• **Enjeu majeur** : 50 ans d'actes d'arpentage et de bornage dormants en cartons à **rendre accessibles et exploitables**.",
    5: '[Temps : 40 s | Chrono : 02:15]\\n• **Monopole légal (Loi 1946)** : seul le géomètre-expert fixe les **limites réelles de propriété**.\\n• **Conservation 30 ans** (Décret 1996) : obligation stricte de conserver et réutiliser les archives.\\n• **Point clé à retenir** : le programme public **GEODÉMAT ne numérise que le cadastre fiscal**.\\n• **Responsabilité du cabinet** : nos 23 600 PV de bornage privés restent en cartons si **nous ne les versons pas nous-mêmes** sur Géofoncier.',
    6: "[Temps : 40 s | Chrono : 02:55]\\n• **Principe clé** : « **Bornage sur bornage ne vaut** » (Art. 646 C. civ.).\\n• **Obligation métier** : rechercher systématiquement les **actes antérieurs** avant tout piquetage.\\n• **Réalité actuelle** : fouille physique manuelle, **15 à 30 minutes par dossier**, risque d'omission.\\n• **Objectif PFE** : transformer ces 23 600 dossiers papier en **pastilles Géofoncier consultables en 1 clic**.",
    7: (
        "[Temps : 45 s | Chrono : 03:40]\n"
        "La problématique centrale du mémoire est la suivante : Dans quelle mesure est-il possible de concevoir un outil logiciel capable d'extraire avec fiabilité les données d'archives foncières très variées, tout en fonctionnant localement sur un ordinateur de bureau ? "
        "Trois contraintes fortes encadrent ce travail : l'hétérogénéité d'un fonds de 50 ans avec des encres pâlies et des calques, l'échec des logiciels OCR du marché sur l'écriture manuscrite des géomètres, "
        "et surtout le secret professionnel qui interdit formellement d'envoyer les plans clients vers le cloud. Tout doit tourner localement sur les PC du cabinet."
    ),
    8: (
        "[Temps : 5 s | Chrono : 03:45]\n"
        "Abordons le Chapitre 2 avec l'analyse métier et les spécifications des données que nous devons traiter."
    ),
    9: '[Temps : 35 s | Chrono : 04:20]\\n• **Distinction essentielle** : le cadastre est un outil **purement fiscal et indicatif**, pas un titre de propriété.\\n• **Seul le PV de bornage** contradictoire signé par les riverains a une **valeur juridique opposable**.\\n• **En justice** : le juge écarte le cadastre au profit des **plans de bornage du géomètre-expert**.\\n• **Notre mission** : indexer ces archives privées, **seule mémoire authentique des limites réelles**.',
    10: "[Temps : 35 s | Chrono : 04:55]\\n• **Spécifications API Géofoncier** : échange REST via jeton **JWT Bearer** sécurisé.\\n• **5 métadonnées cibles** : code INSEE, date ISO, référence, type d'acte et **centroïde Lambert-93**.\\n• *(Préciser à l'oral)* : les données affichées sur l'exemple sont **fictives et anonymisées**.\\n• **Validation** : retour **HTTP 201 Created** pour publication automatique de la pastille.",
    11: "[Temps : 35 s | Chrono : 05:30]\\n• **Trois familles d'archives** très contrastées dans les 23 600 dossiers d'Aubenas :\\n• **1. DMPC récents** : formulaires Cerfa normés, **texte imprimé et cartouches fixes**.\\n• **2. Registres historiques A et B** : cahiers manuscrits en colonnes, **écriture cursive et abréviations**.\\n• **3. Actes de bornage libres** : plans sur calque et textes sans structure figée.\\n• **Conséquence technique** : impossible d'utiliser un outil unique, **nécessité d'une chaîne modulaire**.",
    12: "[Temps : 35 s | Chrono : 06:05]\\n• **Dynamique cadastrale** : les parcelles sont divisées au fil du temps (parcelle mère vers parcelles filles).\\n• **Rupture de numérotation** : un acte de 1984 rédigé sous la parcelle **mère A 14** ne porte pas les numéros actuels **A 115 ou A 116**.\\n• **Moteur de filiation** : remonte l'arbre généalogique cadastral pour **rattacher automatiquement l'archive aux parcelles contemporaines**.\\n• **Bénéfice** : l'antériorité apparaît immédiatement dès que le géomètre clique sur le terrain actuel.",
    13: (
        "[Temps : 5 s | Chrono : 06:10]\n"
        "Passons au Chapitre 3 avec l'état de l'art des technologies d'analyse de documents."
    ),
    14: "[Temps : 35 s | Chrono : 06:45]\\n• **Chaîne OCR moderne (Baek et al., 2019)** articulée en 4 étapes :\\n• **1. Transformation géométrique** : redressement des lignes inclinées.\\n• **2. Extraction de formes (CNN)** et **3. Séquençage (BiLSTM)** pour comprendre l'ordre des lettres.\\n• **4. Prédiction (CTC/Attention)** pour transcrire les caractères.\\n• **Moteurs testés** : Tesseract v5 et EasyOCR (CRAFT) parfaits sur les Cerfa, mais **dépassent 30% d'erreur sur l'écriture manuscrite**.",
    15: (
        "[Temps : 35 s | Chrono : 07:20]\n"
        "Pour le manuscrit, l'approche HTR par Transformer change la donne. Le modèle TrOCR développé par Microsoft associe un encodeur d'image et un décodeur linguistique. "
        "Contrairement à l'OCR qui analyse des caractères isolés, TrOCR lit l'écriture cursive de manière séquentielle et continue. Il atteint un taux d'erreur de caractères inférieur à 6 %, à condition de lui fournir des lignes de texte correctement découpées."
    ),
    16: "[Temps : 35 s | Chrono : 07:55]\\n• **Rôle du NER** : repérer automatiquement les 5 entités clés dans le texte brut transcrit.\\n• **Exemple concret** : isoler la commune, la date, la parcelle et la référence de dossier.\\n• **GLiNER (retenu)** : modèle compact (340M), **zero-shot sans réentraînement**, tourne sur simple CPU, F1 = 0.88.\\n• **LayoutLMv3 (écarté)** : trop lourd (~8 Go GPU) et **s'effondre dès que l'OCR fait une faute de lecture**.",
    17: (
        "[Temps : 35 s | Chrono : 08:30]\n"
        "Les modèles Vision-Langage comme LLaVA ou MiniCPM permettent d'analyser conjointement l'image et le texte. Exécutés localement via Ollama, ils sont capables de comprendre des contextes visuels difficiles. "
        "Cependant, leur coût d'inférence sur CPU est élevé : 5 à 8 secondes par document. Nous avons donc fait le choix de ne pas les utiliser sur tout le flux, mais de les réserver comme arbitres de secours."
    ),
    18: "[Temps : 40 s | Chrono : 09:10]\\n• **Bilan de l'état de l'art** : aucun modèle ne résout seul le problème.\\n• **Les 4 sélections retenues** (conformes à la section 3.6 du mémoire) :\\n• **1. YOLOv8** pour la segmentation spatiale des zones.\\n• **2. EasyOCR et TrOCR** pour la transcription imprimée et manuscrite.\\n• **3. GLiNER** pour l'extraction zero-shot des entités.\\n• **4. LLaVA et MiniCPM-V (Ollama)** comme arbitres locaux de secours sur seuil de confiance.",
    19: (
        "[Temps : 5 s | Chrono : 09:15]\n"
        "Voici le Chapitre 4 détaillant l'architecture logicielle et la chaîne de traitement."
    ),
    20: (
        "[Temps : 50 s | Chrono : 10:05]\n"
        "Voici la chaîne de traitement complète articulée en 6 étapes indépendantes. Le document brut entre à gauche, il est orienté selon sa nature : DMPC, registre ou acte libre. "
        "Des prétraitements redressent l'image et améliorent le contraste avant la segmentation spatiale. Ensuite, la transcription et l'extraction s'enchaînent avec un score de confiance. "
        "Enfin, les données sont contrôlées avant l'injection API. Ce pipeline est 100 % local : aucune donnée nominative ne quitte la machine de l'entreprise."
    ),
    21: (
        "[Temps : 35 s | Chrono : 10:40]\n"
        "La détection spatiale est assurée par un réseau YOLOv8 Nano entraîné sur notre fonds. Il repère instantanément les cartouches, les mentions de commune, les dates et les tampons. "
        "Cela nous permet de découper des vignettes ciblées et de ne pas envoyer une image complète de 300 Mo aux moteurs d'extraction, ce qui optimise directement les temps de calcul."
    ),
    22: (
        "[Temps : 35 s | Chrono : 11:15]\n"
        "Sur ces zones découpées, les moteurs OCR et TrOCR transcrivent le texte brut, puis GLiNER isole les cinq métadonnées obligatoires : commune, section, numéro de parcelle, date d'acte et géomètre signataire. "
        "Un score de confiance est attribué à chaque extraction. Dès que ce score dépasse 0.65, la valeur est transmise directement à la fiabilisation."
    ),
    23: "[Temps : 55 s | Chrono : 12:10]\\n• **Cas réel d'archive** : mention manuscrite « **Lachapelle /s/ AUBENAS** » (« /s/ » = sous).\\n• **Piège de l'IA globale** : entraînée sur la France entière, risque de confusion avec **Lachapelle-Graillouse, Chanéac ou le Vercors**, et calcul de **8 secondes par document**.\\n• **Notre chaîne hybride** : lecture rapide en **0.2 s** + verrouillage immédiat sur les **335 communes d'Ardèche**.\\n• **Le VLM local** n'intervient qu'en **arbitre de secours** sur les cas ambigus. Résultat certifié : **zéro hallucination**.",
    24: (
        "[Temps : 5 s | Chrono : 12:15]\n"
        "Nous arrivons au Chapitre 5 avec la fiabilisation des données et le contrôle humain."
    ),
    25: "[Temps : 40 s | Chrono : 12:55]\\n• **Interface Streamlit** : conçue pour la validation contradictoire par le technicien du cabinet.\\n• **Double affichage** : scan original à gauche, données extraites à droite.\\n• **Jauges de confiance colorées** : vert pour validation 1 clic, orange à vérifier, rouge pour reprise.\\n• **Gain opérationnel** : passage de **25 minutes à moins de 2 minutes** par dossier.\\n• **Principe clé** : **l'humain dans la boucle**, le géomètre conserve l'entière décision juridique.",
    26: "[Temps : 50 s | Chrono : 13:45]\\n• **Moteur de 17 règles métier** réparties en 4 familles :\\n• **1. Cadastre** : vérification de l'INSEE et des sections réelles au PCI.\\n• **2. Dates** : cohérence temporelle avec la période d'exercice du géomètre signataire.\\n• **3. Intervenants** : contrôle du géomètre dans l'annuaire du cabinet et nettoyage des mentions.\\n• **4. Anti-doublon** : contrôle d'unicité pour ne jamais verser deux fois le même acte.\\n• **Sécurité absolue** : exécution en **0.1 seconde**, alerte visuelle si une règle est violée.",
    27: (
        "[Temps : 35 s | Chrono : 14:20]\n"
        "Le contrôle spatial s'effectue via une carte interactive Folium intégrée dans l'outil. Les parcelles de la BD Parcellaire y sont affichées en surbrillance. "
        "L'opérateur peut vérifier d'un coup d'œil que le centroïde calculé se positionne exactement sur la bonne parcelle, ou déplacer manuellement le marqueur si le parcellaire a été remanié depuis l'époque du plan."
    ),
    28: (
        "[Temps : 5 s | Chrono : 14:25]\n"
        "Passons au Chapitre 6 avec la démonstration du pipeline et les résultats sur le terrain."
    ),
    29: (
        "[Temps : 90 s | Chrono : 15:55]\n"
        "Voici la démonstration vidéo du logiciel en conditions réelles. Vous voyez ici l'opérateur charger un lot de dossiers scannés. L'algorithme détecte immédiatement le type de pièce, extrait les cartouches et pré-remplit les métadonnées. "
        "L'opérateur contrôle les valeurs, consulte la carte de localisation, et clique sur Valider. Le document est alors envoyé directement à l'API Géofoncier, qui confirme l'enregistrement et crée la pastille sur le portail national."
    ),
    30: '[Temps : 35 s | Chrono : 16:30]\\n• **Protocole API Géofoncier** articulé en 4 étapes automatisées :\\n• **1. Token JWT Bearer** via OAuth 2.0 (valide 24h).\\n• **2. Calcul du centroïde** parcellaire en coordonnées réglementaires **Lambert-93 (EPSG:2154)**.\\n• **3. Envoi du payload JSON** sur la route /rfuoge (code retour **HTTP 201 Created**).\\n• **4. Téléversement multipart du PDF** sur /dossiersoge : la pastille bleue est **publiée sur la carte nationale**.',
    31: (
        "[Temps : 50 s | Chrono : 17:20]\n"
        "La validation expérimentale a été conduite sur la commune de Prades en Ardèche. 97 dossiers historiques ont été traités : 100 % d'entre eux ont été injectés avec succès sur Géofoncier. "
        "Grâce au filtre des 17 règles métier, aucune fausse pastille n'a été créée. Le temps moyen de traitement a été ramené à moins de 3 minutes par dossier, contre 25 minutes pour la recherche physique manuelle, soit un gain de productivité d'un facteur 6."
    ),
    32: (
        "[Temps : 5 s | Chrono : 17:25]\n"
        "Voyons avec le Chapitre 7 les limites techniques rencontrées et les perspectives."
    ),
    33: (
        "[Temps : 45 s | Chrono : 18:10]\n"
        "Les limites du système sont d'abord physiques : les calques anciens créent des transparences recto-verso, et les tampons administratifs masquent parfois les numéros de parcelles. "
        "Sur le registre manuscrit du géomètre B, l'absence de colonnes fixes nécessite plus de retouches. Sur le plan matériel, l'absence de GPU dédié allonge le temps d'inférence du VLM sur CPU. "
        "Malgré ces contraintes, 82 % des dossiers sont validés directement sans aucune retouche par l'opérateur."
    ),
    34: (
        "[Temps : 45 s | Chrono : 18:55]\n"
        "Les perspectives d'évolution s'articulent autour de 3 axes concrets du mémoire : migrer l'interface vers FastAPI pour supprimer le rechargement de page sur les grands volumes, "
        "connecter les flux WFS de la DGFiP et créer un plugin QGIS pour que les géomètres accèdent aux archives depuis leur outil de production quotidien, et transposer l'outil à d'autres cabinets grâce à un simple fichier de configuration des communes et cartouches. "
        "Enfin, une boucle d'apprentissage actif permettra de réentraîner les modèles locaux au fil des corrections."
    ),
    35: (
        "[Temps : 5 s | Chrono : 19:00]\n"
        "Pour conclure, voici le Chapitre 8 avec le bilan de ce projet de fin d'études."
    ),
    36: (
        "[Temps : 50 s | Chrono : 19:50]\n"
        "Au terme de ce PFE, le bilan est très concret pour GEO-SIAPP : le cabinet dispose d'un outil fonctionnel pour valoriser ses 23 600 dossiers d'archives, avec un gain de temps d'un facteur 6 et un coût d'infrastructure nul puisque tout tourne sur les postes existants. "
        "Ce projet s'inscrit pleinement dans la Charte IA adoptée par l'Ordre des Géomètres-Experts ce 1er septembre 2026 : l'algorithme assiste l'opérateur en pré-remplissant les données, mais l'humain conserve l'entière responsabilité juridique de l'acte foncier."
    ),
    37: (
        "[Temps : 15 s | Chrono : 20:05]\n"
        "Je tiens à remercier chaleureusement Monsieur Gaëtan Hague et toute l'équipe de GEO-SIAPP pour leur accueil et leur confiance, ainsi que Monsieur Mathieu Koehl pour son encadrement à l'INSA Strasbourg. "
        "Je vous remercie pour votre attention et je suis à votre disposition pour vos questions."
    )
}

# Injection des notes dans chaque diapositive avec formattage en gras
for s_num, s_note in SPEAKER_NOTES.items():
    if s_num <= len(prs.slides):
        slide_target = prs.slides[s_num - 1]
        set_formatted_notes(slide_target, s_note)

output_pptx = "Soutenance_PFE_Adrien_TRAVAILLE.pptx"
prs.save(output_pptx)
print(f"\n=======================================================")
print(f"SUCCÈS : Présentation officielle enregistrée sous '{output_pptx}'")
print(f"Nombre total de diapositives générées : {len(prs.slides)}")
print(f"=======================================================\n")
