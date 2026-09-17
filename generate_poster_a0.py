"""
generate_poster_a0.py - Générateur du Poster PFE au format A0 vertical (84,1 x 118,9 cm)
Conforme aux consignes académiques de l'enseignant et aux directives de l'INSA Strasbourg.

Principes de conception appliqués :
- Format officiel : A0 vertical (84,1 x 118,9 cm).
- Variables visuelles limitées (2 au maximum) :
  1. Taille / Échelle de lecture (Titre lisible à 5 m, KPIs à 3 m, figures à 2 m, textes à 1 m).
  2. Couleur structurante unifiée (Bleu Nuit autorité #0C2340, Bleu INSA #1B4587, touches sobres Vert/Rouge).
- Titre synthétique et percutant avec sous-titre explicatif.
- Cadre du travail complet : références des personnes et lieux (nom, adresses, logos, encadrants).
- Organisation logique de l'exposé en 3 colonnes : Problématique -> Méthode IA -> Résultats & Impact.
- Pas de catalogue de formes : géométrie sobre (rectangles purs, bordures fines, aucun chevron/étoile).
- Respiration du regard : marges extérieures de 2.5 cm, gouttières de 2.8 cm, équilibre pleins/vides.
- Typographie proportionnée à l'A0 : textes en 14 à 20 pt, titres en 20 à 46 pt, KPIs en 50 pt.
- Respect strict des directives : aucun tiret cadratin (—) ni tiret isolé.
"""

import os
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from PIL import Image

# =============================================================================
# CONSTANTES ET DIMENSIONS DU FORMAT A0 VERTICAL
# =============================================================================
POSTER_W = Cm(84.1)
POSTER_H = Cm(118.9)

# Palette de couleurs sobre et unifiée
C_BG_PAGE       = RGBColor(248, 250, 252)   # Gris perle très doux pour fond général
C_WHITE         = RGBColor(255, 255, 255)   # Blanc pur pour les cartes
C_NAVY          = RGBColor(12, 35, 64)      # Bleu Nuit profond (autorité, structure)
C_BLEU_INSA     = RGBColor(27, 69, 135)     # Bleu officiel INSA Strasbourg
C_ROUGE_INSA    = RGBColor(214, 40, 40)     # Rouge carmin INSA (accent sobre)
C_ROUGE_BG      = RGBColor(254, 242, 242)   # Fond rouge très pâle pour boîte impasse
C_VERT_CONF     = RGBColor(15, 118, 110)    # Vert émeraude sobre (succès, pastille)
C_VERT_BG       = RGBColor(240, 253, 244)   # Fond vert très pâle pour KPIs positifs
C_BORDER        = RGBColor(203, 213, 225)   # Gris ardoise clair pour bordures fines (1 pt)
C_DARK          = RGBColor(15, 23, 42)      # Noir ardoise (lisibilité maximale)
C_MUTED         = RGBColor(71, 85, 105)     # Gris intermédiaire pour descriptions
C_ACCENT_BG     = RGBColor(241, 245, 249)   # Fond conteneur neutre

FONT_FAMILY     = "Arial"

# Chemins des répertoires et fichiers
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMG_DIR  = os.path.join(BASE_DIR, "img")

LOGO_INSA_PATH    = os.path.join(BASE_DIR, "Logo_INSAStrasbourg.jpg")
LOGO_SIAPP_PATH   = os.path.join(BASE_DIR, "Geosiapp.jpg")
IMG_CHART_PATH    = os.path.join(IMG_DIR, "repartition_archives.png")
IMG_REGISTRE_PATH = os.path.join(IMG_DIR, "registre_manuscrit.jpg")
IMG_PIPELINE_PATH = os.path.join(IMG_DIR, "architecture_tight.jpg")
IMG_YOLO_PATH     = os.path.join(IMG_DIR, "r4p_fig2_tight.png")
IMG_INTERF_PATH   = os.path.join(IMG_DIR, "interface_haut.jpg")
IMG_PASTILLE_PATH = os.path.join(IMG_DIR, "Chap2_pastilles_enhanced.jpg")

OUTPUT_PPTX = os.path.join(BASE_DIR, "poster_PFE_A0.pptx")


# =============================================================================
# UTILITAIRES GRAPHIQUES ET TYPOGRAPHIQUES
# =============================================================================
def add_card(slide, x, y, w, h, bg_color=C_WHITE, border_color=C_BORDER, border_w=Pt(1)):
    """Crée une carte rectangulaire sobre et épurée."""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = border_w
    else:
        shape.line.fill.background()
    return shape


def add_card_header(slide, x, y, w, h, title_text, num_tag=""):
    """Crée un en-tête de carte élégant et lisible avec bandeau supérieur."""
    header_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    header_box.fill.solid()
    header_box.fill.fore_color.rgb = C_NAVY
    header_box.line.fill.background()

    tf = header_box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Cm(0.8)
    tf.margin_right = Cm(0.8)
    tf.margin_top = Cm(0.1)
    tf.margin_bottom = Cm(0.1)

    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    if num_tag:
        r_tag = p.add_run()
        r_tag.text = f"{num_tag}  •  "
        r_tag.font.name = FONT_FAMILY
        r_tag.font.size = Pt(21)
        r_tag.font.bold = True
        r_tag.font.color.rgb = RGBColor(147, 197, 253)

    r_title = p.add_run()
    r_title.text = title_text.upper()
    r_title.font.name = FONT_FAMILY
    r_title.font.size = Pt(21)
    r_title.font.bold = True
    r_title.font.color.rgb = C_WHITE


def set_font(p, name=FONT_FAMILY, size=17, bold=False, color=C_DARK, align=PP_ALIGN.LEFT):
    """Applique le formatage de police standardisé à un paragraphe."""
    p.alignment = align
    for r in p.runs:
        r.font.name = name
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color


def fit_image_box(slide, img_path, box_x, box_y, box_w, box_h):
    """Insère et centre une image en conservant scrupuleusement son ratio d'aspect."""
    if not os.path.exists(img_path):
        return None
    try:
        with Image.open(img_path) as im:
            orig_w, orig_h = im.size
        ratio_box = box_w / box_h
        ratio_img = orig_w / orig_h

        if ratio_img > ratio_box:
            w = box_w
            h = box_w / ratio_img
            x = box_x
            y = box_y + (box_h - h) / 2
        else:
            h = box_h
            w = box_h * ratio_img
            x = box_x + (box_w - w) / 2
            y = box_y
        return slide.shapes.add_picture(img_path, x, y, w, h)
    except Exception as e:
        print(f"Erreur chargement image {img_path}: {e}")
        return None


# =============================================================================
# CONSTRUCTION DU POSTER A0
# =============================================================================
def build_poster():
    prs = Presentation()
    prs.slide_width = POSTER_W
    prs.slide_height = POSTER_H
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # 1. Fond général gris perle très doux
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, POSTER_W, POSTER_H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_BG_PAGE
    bg.line.fill.background()

    # -------------------------------------------------------------------------
    # 2. BANDEAU SUPÉRIEUR D'EN-TÊTE : CADRE DU TRAVAIL & TITRE SYNTHÉTIQUE
    # -------------------------------------------------------------------------
    HEADER_Y = Cm(2.0)
    HEADER_H = Cm(11.5)
    HEADER_W = Cm(79.1)
    HEADER_X = Cm(2.5)

    add_card(slide, HEADER_X, HEADER_Y, HEADER_W, HEADER_H, bg_color=C_WHITE, border_color=C_BORDER, border_w=Pt(1.2))

    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, HEADER_X, HEADER_Y, HEADER_W, Cm(0.40))
    top_bar.fill.solid(); top_bar.fill.fore_color.rgb = C_NAVY; top_bar.line.fill.background()

    if os.path.exists(LOGO_INSA_PATH):
        fit_image_box(slide, LOGO_INSA_PATH, HEADER_X + Cm(1.5), HEADER_Y + Cm(1.2), Cm(12.5), Cm(4.5))

    if os.path.exists(LOGO_SIAPP_PATH):
        fit_image_box(slide, LOGO_SIAPP_PATH, HEADER_X + HEADER_W - Cm(7.0), HEADER_Y + Cm(1.0), Cm(5.2), Cm(5.0))

    tb_title = slide.shapes.add_textbox(HEADER_X + Cm(14.5), HEADER_Y + Cm(0.6), HEADER_W - Cm(22.5), Cm(5.8))
    tf_title = tb_title.text_frame; tf_title.word_wrap = True; tf_title.margin_top = Cm(0.1)

    p1 = tf_title.paragraphs[0]
    p1.text = "TRAITEMENT AUTOMATISÉ ET INSERTION D'ARCHIVES FONCIÈRES SUR GÉOFONCIER"
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.runs[0]
    r1.font.name = FONT_FAMILY; r1.font.size = Pt(37); r1.font.bold = True; r1.font.color.rgb = C_NAVY

    p2 = tf_title.add_paragraph()
    p2.text = "Chaîne d'intelligence artificielle locale pour la valorisation du fonds patrimonial d'un cabinet de géomètre-expert"
    p2.alignment = PP_ALIGN.CENTER; p2.space_before = Pt(8)
    r2 = p2.runs[0]
    r2.font.name = FONT_FAMILY; r2.font.size = Pt(19.5); r2.font.bold = True; r2.font.color.rgb = C_BLEU_INSA

    sep_head = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, HEADER_X + Cm(1.5), HEADER_Y + Cm(6.8), HEADER_W - Cm(3.0), Pt(1))
    sep_head.fill.solid(); sep_head.fill.fore_color.rgb = C_BORDER; sep_head.line.fill.background()

    META_W = (HEADER_W - Cm(4.0)) / 3

    # Méta 1 : Candidat
    tb_m1 = slide.shapes.add_textbox(HEADER_X + Cm(1.5), HEADER_Y + Cm(7.2), META_W, Cm(3.6))
    tf_m1 = tb_m1.text_frame; tf_m1.word_wrap = True
    p = tf_m1.paragraphs[0]; p.text = "CANDIDAT"
    set_font(p, size=13, bold=True, color=C_MUTED)
    p = tf_m1.add_paragraph(); p.text = "Adrien TRAVAILLÉ"
    set_font(p, size=18, bold=True, color=C_DARK); p.space_before = Pt(2)
    p = tf_m1.add_paragraph(); p.text = "Élève-ingénieur en Topographie\nadrien.travaille@insa-strasbourg.fr"
    set_font(p, size=13.5, color=C_MUTED); p.space_before = Pt(2)

    # Méta 2 : Établissement de formation
    tb_m2 = slide.shapes.add_textbox(HEADER_X + Cm(2.0) + META_W, HEADER_Y + Cm(7.2), META_W, Cm(3.6))
    tf_m2 = tb_m2.text_frame; tf_m2.word_wrap = True
    p = tf_m2.paragraphs[0]; p.text = "ÉCOLE & ENCADREMENT"
    set_font(p, size=13, bold=True, color=C_MUTED)
    p = tf_m2.add_paragraph(); p.text = "INSA Strasbourg"
    set_font(p, size=18, bold=True, color=C_DARK); p.space_before = Pt(2)
    p = tf_m2.add_paragraph(); p.text = "Département Génie Topographique • 24 Bd de la Victoire, 67084 Strasbourg\nEnseignant-encadrant : M. Mathieu KOEHL"
    set_font(p, size=13.5, color=C_MUTED); p.space_before = Pt(2)

    # Méta 3 : Entreprise d'accueil
    tb_m3 = slide.shapes.add_textbox(HEADER_X + Cm(2.5) + 2 * META_W, HEADER_Y + Cm(7.2), META_W, Cm(3.6))
    tf_m3 = tb_m3.text_frame; tf_m3.word_wrap = True
    p = tf_m3.paragraphs[0]; p.text = "ENTREPRISE D'ACCUEIL"
    set_font(p, size=13, bold=True, color=C_MUTED)
    p = tf_m3.add_paragraph(); p.text = "Cabinet GEO-SIAPP"
    set_font(p, size=18, bold=True, color=C_DARK); p.space_before = Pt(2)
    p = tf_m3.add_paragraph(); p.text = "18 Chemin de la Plaine, 07200 Aubenas (Ardèche)\nMaître de stage : M. Gaëtan HAGUE (Géomètre-Expert)"
    set_font(p, size=13.5, color=C_MUTED); p.space_before = Pt(2)

    # -------------------------------------------------------------------------
    # 3. FIL CONDUCTEUR VISUEL (NAVIGATION LOGIQUE EN 3 PHASES)
    # -------------------------------------------------------------------------
    STRIP_Y = Cm(14.2)
    STRIP_H = Cm(1.2)
    
    COL_W = Cm(24.5)
    GAP_X = Cm(2.8)
    C1_X  = HEADER_X
    C2_X  = C1_X + COL_W + GAP_X
    C3_X  = C2_X + COL_W + GAP_X

    strip_titles = [
        (C1_X, "1. PROBLÉMATIQUE MÉTIER & LE GÎSEMENT D'ARCHIVES", C_NAVY),
        (C2_X, "2. CHAÎNE TECHNOLOGIQUE D'INTELLIGENCE ARTIFICIELLE", C_NAVY),
        (C3_X, "3. VALIDATION OPÉRATEUR, RÉSULTATS & VALORISATION", C_NAVY)
    ]
    for x_pos, title, col in strip_titles:
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x_pos, STRIP_Y, COL_W, STRIP_H)
        box.fill.solid(); box.fill.fore_color.rgb = col; box.line.fill.background()
        tf = box.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.text = title
        set_font(p, size=13.5, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    # -------------------------------------------------------------------------
    # 4. HAUTEURS HARMONISÉES DES CARTES
    # -------------------------------------------------------------------------
    BODY_Y  = Cm(16.0)
    CARD1_H = Cm(29.5)
    GAP_Y   = Cm(1.5)
    CARD2_Y = BODY_Y + CARD1_H + GAP_Y # 47.0 cm
    CARD2_H = Cm(32.5)
    CARD3_Y = CARD2_Y + CARD2_H + GAP_Y # 81.0 cm
    CARD3_H = Cm(29.0)

    # -------------------------------------------------------------------------
    # 5. COLONNE 1 : LE CONSTAT & LE VERROU MÉTIER
    # -------------------------------------------------------------------------
    # Carte 1.1 : Le stock documentaire et le cadre réglementaire + Graphique circulaire
    add_card(slide, C1_X, BODY_Y, COL_W, CARD1_H)
    add_card_header(slide, C1_X, BODY_Y, COL_W, Cm(1.5), "Le gisement d'archives et le droit foncier", "1.1")

    tb = slide.shapes.add_textbox(C1_X + Cm(0.9), BODY_Y + Cm(1.8), COL_W - Cm(1.8), Cm(15.0))
    tf = tb.text_frame; tf.word_wrap = True

    pts_1_1 = [
        ("Stock ciblé à Aubenas", "23 600 dossiers papier historiques conservés au siège (sur 40 000 au cabinet)."),
        ("Succession de 6 prédécesseurs", "Fonds Harrois, Blache, Teyssier, Reynaud (1959 à 2007) sans convention unique."),
        ("Supports hétérogènes fragiles", "Plans calques à l'encre, tirages ozalides bleutés, liasses et registres manuscrits."),
        ("Conservation trentenaire légale", "Décret 96-478 (art. 55) : obligation stricte de conserver les archives et de les transmettre."),
        ("Règle 'Bornage sur bornage ne vaut'", "Interdiction de redéfinir une limite déjà bornée : recherche d'antériorité obligatoire.")
    ]
    for i, (title, desc) in enumerate(pts_1_1):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"•  {title}"
        set_font(p, size=16.0, bold=True, color=C_NAVY)
        if i > 0: p.space_before = Pt(5)
        p_sub = tf.add_paragraph()
        p_sub.text = desc
        set_font(p_sub, size=13.5, color=C_DARK)
        p_sub.space_before = Pt(1)

    # Graphique de répartition du fonds en bas de la carte 1.1
    if os.path.exists(IMG_CHART_PATH):
        fit_image_box(slide, IMG_CHART_PATH, C1_X + Cm(0.9), BODY_Y + Cm(16.5), COL_W - Cm(1.8), Cm(10.5))
        tb_c = slide.shapes.add_textbox(C1_X + Cm(0.9), BODY_Y + Cm(27.0), COL_W - Cm(1.8), Cm(2.2))
        p = tb_c.text_frame.paragraphs[0]
        p.text = "Répartition estimée des 23 600 dossiers d'Aubenas par géomètre prédécesseur"
        set_font(p, size=11.5, color=C_MUTED, align=PP_ALIGN.CENTER)

    # Carte 1.2 : Illustration du document ancien (registre à la plume)
    add_card(slide, C1_X, CARD2_Y, COL_W, CARD2_H)
    add_card_header(slide, C1_X, CARD2_Y, COL_W, Cm(1.5), "Complexité des documents d'archives", "1.2")

    if os.path.exists(IMG_REGISTRE_PATH):
        fit_image_box(slide, IMG_REGISTRE_PATH, C1_X + Cm(0.9), CARD2_Y + Cm(1.8), COL_W - Cm(1.8), Cm(18.0))

    tb_leg1 = slide.shapes.add_textbox(C1_X + Cm(0.9), CARD2_Y + Cm(20.2), COL_W - Cm(1.8), Cm(11.5))
    tf_leg1 = tb_leg1.text_frame; tf_leg1.word_wrap = True
    p = tf_leg1.paragraphs[0]
    p.text = "Extrait de registre d'arpentage manuscrit du cabinet GEO-SIAPP"
    set_font(p, size=15.0, bold=True, color=C_NAVY, align=PP_ALIGN.CENTER)
    
    desc_leg1 = [
        "Calligraphies cursives à la plume et encres d'époque délavées.",
        "Abréviations métier condensées ('Déf.', 'Bge', 'Div.', 'Prop.').",
        "Disparité des structures de tableaux selon les géomètres prédécesseurs.",
        "Altérations physiques du papier et absence d'identifiant parcellaire unique.",
        "Document d'archives manuscrit conservé sans échelle graphique."
    ]
    for d in desc_leg1:
        p_sub = tf_leg1.add_paragraph()
        p_sub.text = f"•  {d}"
        set_font(p_sub, size=13.5, color=C_MUTED)
        p_sub.space_before = Pt(3)

    # Carte 1.3 : L'impasse manuelle et le cahier des charges
    add_card(slide, C1_X, CARD3_Y, COL_W, CARD3_H)
    add_card_header(slide, C1_X, CARD3_Y, COL_W, Cm(1.5), "Impasse du traitement manuel et objectifs", "1.3")

    # Boîte 1 : Impasse manuelle (fond rouge pâle)
    box_impasse = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, C1_X + Cm(0.9), CARD3_Y + Cm(1.8), COL_W - Cm(1.8), Cm(12.0))
    box_impasse.fill.solid(); box_impasse.fill.fore_color.rgb = C_ROUGE_BG
    box_impasse.line.color.rgb = C_ROUGE_INSA; box_impasse.line.width = Pt(1.5)
    tf_imp = box_impasse.text_frame; tf_imp.word_wrap = True
    tf_imp.margin_left = Cm(0.6); tf_imp.margin_right = Cm(0.6); tf_imp.margin_top = Cm(0.4)
    p = tf_imp.paragraphs[0]
    p.text = "LE CONSTAT : IMPASSE DU TRAITEMENT MANUEL"
    set_font(p, size=15.5, bold=True, color=C_ROUGE_INSA)
    imp_items = [
        "Temps moyen : 15 à 20 minutes par acte sur GéofoncierEXPERT.",
        "Volume cible : 23 600 dossiers accumulés sur 50 ans.",
        "Charge cumulée : 3,5 années de travail d'un technicien à plein temps (plus de 5 900 heures).",
        "Conséquence : risque d'erreur de saisie élevé et gisement documentaire inexploité."
    ]
    for it in imp_items:
        p_it = tf_imp.add_paragraph(); p_it.text = f"✕  {it}"
        set_font(p_it, size=13.5, color=C_DARK); p_it.space_before = Pt(3)

    # Boîte 2 : Cahier des charges (fond bleu ciel pâle)
    box_cdc = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, C1_X + Cm(0.9), CARD3_Y + Cm(14.5), COL_W - Cm(1.8), Cm(13.2))
    box_cdc.fill.solid(); box_cdc.fill.fore_color.rgb = RGBColor(240, 249, 255)
    box_cdc.line.color.rgb = C_BLEU_INSA; box_cdc.line.width = Pt(1.5)
    tf_cdc = box_cdc.text_frame; tf_cdc.word_wrap = True
    tf_cdc.margin_left = Cm(0.6); tf_cdc.margin_right = Cm(0.6); tf_cdc.margin_top = Cm(0.4)
    p = tf_cdc.paragraphs[0]
    p.text = "LE CAHIER DES CHARGES ASSIGNÉ AU PFE"
    set_font(p, size=15.5, bold=True, color=C_NAVY)
    cdc_items = [
        "Traitement 100 % local : données protégées, secret professionnel, aucun flux cloud.",
        "Autonomie sur PC bureautique classique du cabinet sans GPU dédié.",
        "Fiabilisation automatique contre les bases officielles (INSEE, BAN).",
        "Supervision humaine obligatoire : l'IA assiste l'opérateur, qui valide d'un clic."
    ]
    for it in cdc_items:
        p_it = tf_cdc.add_paragraph(); p_it.text = f"✓  {it}"
        set_font(p_it, size=13.5, color=C_NAVY); p_it.space_before = Pt(3)

    # -------------------------------------------------------------------------
    # 6. COLONNE 2 : LA CHAÎNE D'INTELLIGENCE ARTIFICIELLE LOCALE
    # -------------------------------------------------------------------------
    # Carte 2.1 : Architecture séquentielle en 6 étapes avec schéma 7K
    add_card(slide, C2_X, BODY_Y, COL_W, CARD1_H)
    add_card_header(slide, C2_X, BODY_Y, COL_W, Cm(1.5), "Architecture du pipeline modulaire local", "2.1")

    if os.path.exists(IMG_PIPELINE_PATH):
        fit_image_box(slide, IMG_PIPELINE_PATH, C2_X + Cm(0.9), BODY_Y + Cm(1.8), COL_W - Cm(1.8), Cm(11.5))

    tb_pipe = slide.shapes.add_textbox(C2_X + Cm(0.9), BODY_Y + Cm(13.6), COL_W - Cm(1.8), Cm(15.2))
    tf_pipe = tb_pipe.text_frame; tf_pipe.word_wrap = True

    etapes_resum = [
        ("1. Classification du document", "Tri automatique des plans graphiques, registres et procès-verbaux."),
        ("2. Détection spatiale YOLOv8", "Localisation des cartouches, tableaux récapitulatifs, tampons et signatures."),
        ("3. Lecture hybride OCR / HTR", "Binarisation Otsu, Tesseract (imprimé) et Vision Transformer TrOCR (manuscrit)."),
        ("4. Extraction d'entités (GLiNER)", "Reconnaissance ciblée des communes, dates et géomètres sans réentraînement."),
        ("5. Contrôle & arbitrage VLM", "Vérification Levenshtein (INSEE) et arbitrage par LLaVA 7B local en cas de doute."),
        ("6. Injection API Géofoncier", "Génération du JSON normalisé et versement automatique sur l'API GéofoncierEXPERT.")
    ]
    for i, (title, desc) in enumerate(etapes_resum):
        p = tf_pipe.paragraphs[0] if i == 0 else tf_pipe.add_paragraph()
        p.text = f"{title} : {desc}"
        set_font(p, size=13.5, color=C_DARK)
        if i > 0: p.space_before = Pt(4)
        if len(p.runs) > 0:
            p.runs[0].font.bold = True; p.runs[0].font.color.rgb = C_NAVY

    # Carte 2.2 : Détection spatiale YOLOv8 (Image centrale)
    add_card(slide, C2_X, CARD2_Y, COL_W, CARD2_H)
    add_card_header(slide, C2_X, CARD2_Y, COL_W, Cm(1.5), "Détection spatiale des zones clés (YOLOv8)", "2.2")

    if os.path.exists(IMG_YOLO_PATH):
        fit_image_box(slide, IMG_YOLO_PATH, C2_X + Cm(0.9), CARD2_Y + Cm(1.8), COL_W - Cm(1.8), Cm(21.0))

    tb_leg2 = slide.shapes.add_textbox(C2_X + Cm(0.9), CARD2_Y + Cm(23.2), COL_W - Cm(1.8), Cm(8.5))
    tf_leg2 = tb_leg2.text_frame; tf_leg2.word_wrap = True
    p = tf_leg2.paragraphs[0]
    p.text = "Détection instantanée par réseau YOLOv8 sur un plan DMPC"
    set_font(p, size=15.0, bold=True, color=C_NAVY, align=PP_ALIGN.CENTER)

    desc_yolo = [
        "Localisation en moins de 150 ms des cartouches, parcelles mères/filles et échelles.",
        "Insensibilité aux variations d'échelle, de cadrage et aux légères rotations.",
        "Découpage immédiat des imagettes pour nourrir les modules de lecture OCR et HTR.",
        "Plan régulier sans échelle graphique, orientation Nord en tête."
    ]
    for d in desc_yolo:
        p_sub = tf_leg2.add_paragraph(); p_sub.text = f"•  {d}"
        set_font(p_sub, size=13.5, color=C_MUTED); p_sub.space_before = Pt(2)

    # Carte 2.3 : Déploiement souverain et matrice technique
    add_card(slide, C2_X, CARD3_Y, COL_W, CARD3_H)
    add_card_header(slide, C2_X, CARD3_Y, COL_W, Cm(1.5), "Exécution souveraine et matrice décisionnelle", "2.3")

    tb = slide.shapes.add_textbox(C2_X + Cm(0.9), CARD3_Y + Cm(1.8), COL_W - Cm(1.8), Cm(7.2))
    tf = tb.text_frame; tf.word_wrap = True
    pts_2_3 = [
        ("Indépendance technologique", "Zéro flux réseau externe : secret professionnel et conformité légale garantis."),
        ("Parc bureautique standard", "Exécution fluide sur processeurs bureautiques existants du cabinet sans GPU dédié."),
        ("Temps de calcul optimisé", "Moins de 4 secondes de traitement automatique par document grand format.")
    ]
    for i, (title, desc) in enumerate(pts_2_3):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"•  {title} : {desc}"
        set_font(p, size=14.5, color=C_DARK)
        if i > 0: p.space_before = Pt(4)
        if len(p.runs) > 0:
            p.runs[0].font.bold = True; p.runs[0].font.color.rgb = C_NAVY

    # Matrice des briques en bas de carte 2.3
    matrix_y = CARD3_Y + Cm(9.6)
    briques = [
        ("YOLOv8 Nano", "Segmentation spatiale des cartouches et parcelles", "120 ms", C_NAVY),
        ("Tesseract & TrOCR", "Lecture hybride imprimé (dactylo) et manuscrit (plume)", "350 ms à 1,8 s", C_BLEU_INSA),
        ("GLiNER Multi-v2.1", "Extraction d'entités ciblées (commune, date, géomètre)", "280 ms", C_VERT_CONF),
        ("LLaVA 7B (Ollama)", "Arbitrage contextuel en secours (activé sur < 15 % des cas)", "4 à 8 s", RGBColor(107, 33, 168))
    ]
    for idx, (nom, role, tps, col_br) in enumerate(briques):
        by = matrix_y + idx * Cm(4.5)
        bbox = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, C2_X + Cm(0.9), by, COL_W - Cm(1.8), Cm(4.0))
        bbox.fill.solid(); bbox.fill.fore_color.rgb = C_ACCENT_BG
        bbox.line.color.rgb = C_BORDER; bbox.line.width = Pt(1)
        tf_b = bbox.text_frame; tf_b.word_wrap = True
        tf_b.margin_left = Cm(0.6); tf_b.margin_top = Cm(0.3)
        p = tf_b.paragraphs[0]; p.text = f"{nom}  •  {tps}"
        set_font(p, size=14.5, bold=True, color=col_br)
        p2 = tf_b.add_paragraph(); p2.text = role
        set_font(p2, size=13.0, color=C_DARK); p2.space_before = Pt(2)

    # -------------------------------------------------------------------------
    # 7. COLONNE 3 : VALIDATION OPÉRATEUR, RÉSULTATS & VALORISATION
    # -------------------------------------------------------------------------
    # Carte 3.1 : Interface Streamlit + Pastille Géofoncier (Workflow complet)
    add_card(slide, C3_X, BODY_Y, COL_W, CARD1_H)
    add_card_header(slide, C3_X, BODY_Y, COL_W, Cm(1.5), "Supervision humaine et publication Géofoncier", "3.1")

    # Image 1 : Interface Streamlit
    if os.path.exists(IMG_INTERF_PATH):
        fit_image_box(slide, IMG_INTERF_PATH, C3_X + Cm(0.9), BODY_Y + Cm(1.8), COL_W - Cm(1.8), Cm(10.5))

    # Image 2 : Pastille Géofoncier
    if os.path.exists(IMG_PASTILLE_PATH):
        fit_image_box(slide, IMG_PASTILLE_PATH, C3_X + Cm(0.9), BODY_Y + Cm(12.8), COL_W - Cm(1.8), Cm(8.8))

    tb_leg3 = slide.shapes.add_textbox(C3_X + Cm(0.9), BODY_Y + Cm(22.0), COL_W - Cm(1.8), Cm(6.8))
    tf_leg3 = tb_leg3.text_frame; tf_leg3.word_wrap = True
    p = tf_leg3.paragraphs[0]
    p.text = "Poste opérateur Streamlit et publication sur Géofoncier"
    set_font(p, size=14.5, bold=True, color=C_NAVY, align=PP_ALIGN.CENTER)

    pts_sup = [
        "L'IA préremplit le formulaire interactif à partir des détections du plan.",
        "Le géomètre-expert contrôle visuellement et valide en 1 clic.",
        "Envoi direct par API REST et pastille immédiatement active sur la carte nationale."
    ]
    for d in pts_sup:
        p_sub = tf_leg3.add_paragraph(); p_sub.text = f"•  {d}"
        set_font(p_sub, size=13.0, color=C_DARK); p_sub.space_before = Pt(2)

    # Carte 3.2 : Indicateurs de performance (Grands KPIs visibles à 3 mètres)
    add_card(slide, C3_X, CARD2_Y, COL_W, CARD2_H)
    add_card_header(slide, C3_X, CARD2_Y, COL_W, Cm(1.5), "Performances mesurées et gains opérationnels", "3.2")

    KPI_H = Cm(8.8)
    KPI_W = COL_W - Cm(1.8)
    kpis = [
        ("< 1 min 30 s", "DURÉE TOTALE PAR DOSSIER VALIDÉ", "Temps unitaire divisé par 10 (comparé à 15 à 20 minutes en saisie manuelle).", C_VERT_CONF, C_VERT_BG),
        ("88,0 %", "SCORE F1 SUR LES COMMUNES & CODES INSEE", "Exactitude élevée grâce à l'appariement flou Levenshtein sur les bases officielles.", C_NAVY, C_ACCENT_BG),
        ("100 % Local", "SOUVERAINETÉ ET SECRET PROFESSIONNEL", "Aucune donnée de propriétaire transmise à l'extérieur, zéro dépendance cloud.", C_BLEU_INSA, RGBColor(239, 246, 255))
    ]
    for i, (val, label, sub, color_num, color_bg) in enumerate(kpis):
        ky = CARD2_Y + Cm(1.8) + i * (KPI_H + Cm(0.8))
        kbox = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, C3_X + Cm(0.9), ky, KPI_W, KPI_H)
        kbox.fill.solid(); kbox.fill.fore_color.rgb = color_bg
        kbox.line.color.rgb = color_num; kbox.line.width = Pt(1.5)

        tb_k = slide.shapes.add_textbox(C3_X + Cm(1.1), ky + Cm(0.5), KPI_W - Cm(0.4), KPI_H - Cm(1.0))
        tf_k = tb_k.text_frame; tf_k.word_wrap = True
        p = tf_k.paragraphs[0]; p.text = val
        set_font(p, size=48, bold=True, color=color_num, align=PP_ALIGN.CENTER)
        p_l = tf_k.add_paragraph(); p_l.text = label
        set_font(p_l, size=15.5, bold=True, color=C_DARK, align=PP_ALIGN.CENTER)
        p_l.space_before = Pt(3)
        p_s = tf_k.add_paragraph(); p_s.text = sub
        set_font(p_s, size=13.0, color=C_MUTED, align=PP_ALIGN.CENTER)
        p_s.space_before = Pt(2)

    # Carte 3.3 : Valorisation industrielle et boucle d'apprentissage
    add_card(slide, C3_X, CARD3_Y, COL_W, CARD3_H)
    add_card_header(slide, C3_X, CARD3_Y, COL_W, Cm(1.5), "Impact industriel et boucle d'apprentissage", "3.3")

    tb = slide.shapes.add_textbox(C3_X + Cm(0.9), CARD3_Y + Cm(1.8), COL_W - Cm(1.8), Cm(7.2))
    tf = tb.text_frame; tf.word_wrap = True

    pts_3_3 = [
        ("Déploiement sur 4 agences", "Aubenas (siège), Pierrelatte, Vallon-Pont-d'Arc, Guilherand-Granges."),
        ("Réactivation du patrimoine dormant", "Accès immédiat aux dossiers historiques pour les équipes terrain lors des bornages."),
        ("Rentabilité et pérennité", "Solution souveraine sans abonnement logiciel récurrent, maîtrise intégrale du code source."),
        ("Transposition à la profession", "Méthodologie reproductible pour les cabinets confrontés aux archives papier.")
    ]
    for i, (title, desc) in enumerate(pts_3_3):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"•  {title} : {desc}"
        set_font(p, size=13.5, color=C_DARK)
        if i > 0: p.space_before = Pt(3)
        if len(p.runs) > 0:
            p.runs[0].font.bold = True; p.runs[0].font.color.rgb = C_NAVY

    # Boîte Active Learning étagée en bas de carte 3.3
    al_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, C3_X + Cm(0.9), CARD3_Y + Cm(9.6), COL_W - Cm(1.8), Cm(18.0))
    al_box.fill.solid(); al_box.fill.fore_color.rgb = RGBColor(240, 253, 244)
    al_box.line.color.rgb = C_VERT_CONF; al_box.line.width = Pt(1.5)
    tf_al = al_box.text_frame; tf_al.word_wrap = True
    tf_al.margin_left = Cm(0.6); tf_al.margin_top = Cm(0.4)
    p = tf_al.paragraphs[0]; p.text = "BOUCLE D'APPRENTISSAGE ACTIF (ACTIVE LEARNING)"
    set_font(p, size=15.0, bold=True, color=C_VERT_CONF)

    al_steps = [
        ("1. Prédiction & Extraction", "YOLOv8 et GLiNER analysent automatiquement le document scanné."),
        ("2. Validation Opérateur", "Contrôle visuel sur Streamlit, ajustement manuel si nécessaire et validation en 1 clic."),
        ("3. Collecte des Écarts", "Enregistrement local automatique des paires (image brute, correction validée)."),
        ("4. Réentraînement Incrémental", "Réajustement périodique des poids du modèle local sans coût d'annotation.")
    ]
    for step_num, (st_t, st_d) in enumerate(al_steps):
        p_st = tf_al.add_paragraph()
        p_st.text = f"{st_t}"
        set_font(p_st, size=14.0, bold=True, color=C_NAVY)
        p_st.space_before = Pt(5)
        p_sd = tf_al.add_paragraph()
        p_sd.text = f"→ {st_d}"
        set_font(p_sd, size=12.5, color=C_DARK)
        p_sd.space_before = Pt(1)

    # -------------------------------------------------------------------------
    # 8. PIED DE PAGE : MENTIONS LÉGALES, ORDRE DES GÉOMÈTRES-EXPERTS & SESSION
    # -------------------------------------------------------------------------
    FOOTER_Y = Cm(111.0)
    FOOTER_H = Cm(5.5)
    FOOTER_W = HEADER_W
    FOOTER_X = HEADER_X

    add_card(slide, FOOTER_X, FOOTER_Y, FOOTER_W, FOOTER_H, bg_color=C_WHITE, border_color=C_BORDER, border_w=Pt(1.2))

    tb_f = slide.shapes.add_textbox(FOOTER_X + Cm(1.5), FOOTER_Y + Cm(0.6), FOOTER_W - Cm(3.0), FOOTER_H - Cm(1.0))
    tf_f = tb_f.text_frame; tf_f.word_wrap = True

    p = tf_f.paragraphs[0]
    p.text = "PROJET DE FIN D'ÉTUDES INGÉNIEUR  •  DÉPARTEMENT GÉNIE TOPOGRAPHIQUE  •  INSA STRASBOURG"
    set_font(p, size=15.0, bold=True, color=C_NAVY, align=PP_ALIGN.CENTER)

    p = tf_f.add_paragraph()
    p.text = "Travaux réalisés au sein du Cabinet GEO-SIAPP (Aubenas, Ardèche)  •  Sous la direction de M. Gaëtan HAGUE (GEO-SIAPP) et M. Mathieu KOEHL (INSA Strasbourg)"
    set_font(p, size=13.5, color=C_MUTED, align=PP_ALIGN.CENTER)
    p.space_before = Pt(3)

    p = tf_f.add_paragraph()
    p.text = "Conformité stricte aux exigences de l'Ordre des Géomètres-Experts (OGE) et aux spécifications techniques de la plateforme nationale Géofoncier  •  Session de soutenance : 8 septembre 2026"
    set_font(p, size=12.5, color=C_MUTED, align=PP_ALIGN.CENTER)
    p.space_before = Pt(3)

    prs.save(OUTPUT_PPTX)
    print(f"Poster A0 vertical régénéré avec succès : {OUTPUT_PPTX}")


if __name__ == "__main__":
    build_poster()
