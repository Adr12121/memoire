"""
generate_poster.py - Poster PFE A1 vertical (59.4 x 84.1 cm)
Adrien TRAVAILLÉ - GEO-SIAPP Aubenas - INSA Strasbourg - Session Septembre 2026

Fidélité absolue aux références d'excellence (Schertzinger & Mroue) :
- En-tête officiel INSA Rouge + Cartouche Entreprise GEO-SIAPP + Mots-clés
- 2 colonnes rigoureusement équilibrées et alignées
- Bandeaux de sections unis Bleu Nuit GEO-SIAPP avec typographie blanche percutante
- Sous-titres élégants et structurés
- Vocabulaire de formes homogène (rectangles nets et bordures fines)
- Schéma de workflow complet et riche dans la colonne droite
- Problèmes et Solutions en 3 colonnes avec codes couleur (Rouge / Vert)
- Tableau comparatif des approches et métriques KPI
- Termes clés mis en valeur en gras
- Remplissage optimal de l'espace sans blancs superflus (bloque les blancs)
- Respect strict des directives (aucun tiret cadratin, style direct et rigoureux)
"""

import os
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMG_DIR  = os.path.join(BASE_DIR, "img")
OUTPUT   = os.path.join(BASE_DIR, "poster_PFE.pptx")

# ============================================================
#  CHARTE GRAPHIQUE OFFICIELLE
# ============================================================
C_NAVY   = RGBColor(12,  52,  100)   # #0C3464 - Bleu Nuit GEO-SIAPP
C_ROUGE  = RGBColor(218, 41,  28)    # #DA291C - Rouge officiel INSA
C_VERT   = RGBColor(25,  135, 60)    # #19873C - Vert solution soutenu
C_BLANC  = RGBColor(255, 255, 255)
C_NOIR   = RGBColor(20,  30,  50)    # Texte principal haute lisibilité
C_GRIS_F = RGBColor(245, 247, 251)   # Fond gris très clair pour encarts
C_GRIS_T = RGBColor(232, 236, 244)   # Alternance lignes tableau
C_BDR    = RGBColor(185, 198, 218)   # Bordure fine subtile
C_MUTED  = RGBColor(95,  108, 125)   # Texte secondaire / légendes
FONT     = "Arial"

# Dimensions A1 vertical standard
W = int(Cm(59.4))
H = int(Cm(84.1))

# Grille et marges
MARG_X = int(Cm(1.3))
COL_W  = int(Cm(28.0))
GAP_X  = W - 2 * MARG_X - 2 * COL_W  # 0.8 cm exact
COL_L  = MARG_X
COL_R  = MARG_X + COL_W + GAP_X

# En-tête
TITLE_H = int(Cm(4.0))
INFO_H  = int(Cm(6.0))
KW_H    = int(Cm(1.4))
HEADER_BOT = TITLE_H + INFO_H + KW_H  # 11.4 cm
GAP_TOP = int(Cm(0.6))
CY_START = HEADER_BOT + GAP_TOP       # 12.0 cm début contenu

# Hauteur de bandeau de section
SBAR = int(Cm(2.0))
GSEC = int(Cm(0.5))  # Gap vertical inter-sections
PAD  = int(Cm(0.35)) # Marge intérieure

# ============================================================
#  GÉOMÉTRIE VERTICALE DES SECTIONS (Synchronisation absolue)
# ============================================================
# Colonne Gauche (Total contenu = 71.5 cm -> termine à 83.5 cm)
# L1 : Contexte et Objectifs (14.0 cm) -> 12.0 à 26.0
# L2 : Technologie et Fonctionnement (22.0 cm) -> 26.5 à 48.5
# (Ligne médiane à 48.5 cm parfaitement calée avec R1)
# L3 : Exactitude et Performances (24.5 cm) -> 49.0 à 73.5
# L4 : Problèmes et Solutions (9.5 cm) -> 74.0 à 83.5

L1_Y = CY_START
L1_H = int(Cm(14.0))

L2_Y = L1_Y + L1_H + GSEC
L2_H = int(Cm(22.0))

L3_Y = L2_Y + L2_H + GSEC
L3_H = int(Cm(24.5))

L4_Y = L3_Y + L3_H + GSEC
L4_H = int(Cm(9.5))

# Colonne Droite
# R1 : Chaîne de Traitement (36.5 cm) -> 12.0 à 48.5 (Synchro L2 !)
# R2 : Applications Opérationnelles (10.0 cm) -> 49.0 à 59.0
# R3 : Nouvelles Solutions et Perspectives (10.0 cm) -> 59.5 à 69.5
# R4 : Conclusion et Perspectives (13.5 cm) -> 70.0 à 83.5 (Synchro L4 !)

R1_Y = CY_START
R1_H = int(Cm(36.5))

R2_Y = R1_Y + R1_H + GSEC
R2_H = int(Cm(10.0))

R3_Y = R2_Y + R2_H + GSEC
R3_H = int(Cm(10.0))

R4_Y = R3_Y + R3_H + GSEC
R4_H = int(Cm(13.5))


# ============================================================
#  FONCTIONS PRIMITIVES DE DESSIN (Robuste OpenXML)
# ============================================================

def rect(sl, l, t, w, h, fill=None, border=None, bw=0.75):
    """Trace un rectangle net avec dimensions strictement positives."""
    w = max(int(w), int(Pt(1)))
    h = max(int(h), int(Pt(1)))
    shp = sl.shapes.add_shape(1, int(l), int(t), w, h)
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    else:
        shp.fill.background()
    if border:
        shp.line.color.rgb = border
        shp.line.width = int(Pt(bw))
    else:
        shp.line.fill.background()
    return shp


def pic(sl, path, l, t, w, h):
    """Ajoute une image si le fichier existe."""
    if os.path.exists(str(path)):
        sl.shapes.add_picture(str(path), int(l), int(t), int(w), int(h))


def mk_tf(sl, l, t, w, h, pad=Cm(0.08)):
    """Crée un cadre de texte avec marges contrôlées."""
    box = sl.shapes.add_textbox(int(l), int(t), int(w), int(h))
    tf  = box.text_frame
    tf.word_wrap = True
    p = int(pad)
    tf.margin_top = tf.margin_bottom = p
    tf.margin_left = tf.margin_right = p
    return tf


def _clear_p(tf):
    p = tf.paragraphs[0]
    for c in list(p._p):
        p._p.remove(c)
    return p


def ap(tf, text, bold=False, italic=False, sz=10.0, col=C_NOIR,
       align=PP_ALIGN.LEFT, sa=2.0, sb=0.0, first=False):
    """Ajoute un paragraphe de texte uniforme."""
    p = _clear_p(tf) if first else tf.add_paragraph()
    p.alignment    = align
    p.space_before = Pt(sb)
    p.space_after  = Pt(sa)
    r = p.add_run()
    r.text           = text
    r.font.bold      = bold
    r.font.italic    = italic
    r.font.size      = Pt(sz)
    r.font.color.rgb = col
    r.font.name      = FONT
    return p


def ap_rich(tf, runs, sz=10.0, align=PP_ALIGN.LEFT, sa=2.0, sb=0.0, first=False):
    """Ajoute un paragraphe contenant plusieurs segments de styles différents.
    runs = [(text, bold, color, italic), ...]
    """
    p = _clear_p(tf) if first else tf.add_paragraph()
    p.alignment    = align
    p.space_before = Pt(sb)
    p.space_after  = Pt(sa)
    for run_tuple in runs:
        text = run_tuple[0]
        bold = run_tuple[1] if len(run_tuple) > 1 else False
        col  = run_tuple[2] if len(run_tuple) > 2 and run_tuple[2] else C_NOIR
        ital = run_tuple[3] if len(run_tuple) > 3 else False
        
        r = p.add_run()
        r.text           = text
        r.font.bold      = bold
        r.font.italic    = ital
        r.font.size      = Pt(sz)
        r.font.color.rgb = col
        r.font.name      = FONT
    return p


def section_card(sl, x, y, w, h, title):
    """Crée une section complète : bandeau bleu nuit + cadre de corps."""
    rect(sl, x, y, w, h, fill=C_BLANC, border=C_BDR, bw=1.0)
    rect(sl, x, y, w, SBAR, fill=C_NAVY)
    tf = mk_tf(sl, x + int(Cm(0.3)), y + int(Cm(0.15)),
               w - int(Cm(0.6)), SBAR - int(Cm(0.2)), pad=Cm(0.04))
    ap(tf, title, bold=True, sz=22, col=C_BLANC,
       align=PP_ALIGN.CENTER, sa=0, first=True)
    return y + SBAR, h - SBAR


def stitle(sl, x, y, w, text, sz=11.5, top_gap=True):
    """Sous-titre de section en gras bleu nuit (style Schertzinger / Mroue)."""
    gap = int(Cm(0.22)) if top_gap else 0
    h = int(Cm(0.72))
    tf = mk_tf(sl, x + PAD, y + gap, w - 2*PAD, h, pad=Cm(0.03))
    ap(tf, text, bold=True, sz=sz, col=C_NAVY, sa=0, first=True)
    return y + gap + h + int(Cm(0.08))


def arrow_down(sl, cx, y, h, col=C_ROUGE):
    """Flèche verticale nette vers le bas avec sécurité OpenXML."""
    h = max(int(h), int(Cm(0.55)))
    tip_w = int(Cm(0.44))
    tip_h = int(Cm(0.28))
    tige_h = max(int(h) - tip_h, int(Pt(2)))
    rect(sl, cx - int(Pt(1.2)), int(y), int(Pt(2.4)), tige_h, fill=col)
    rect(sl, cx - tip_w//2, int(y) + tige_h, tip_w, tip_h, fill=col)


def kpi_card(sl, x, y, w, h, value, label, sublabel="", val_col=C_NAVY):
    """Carte métrique KPI haute visibilité."""
    rect(sl, x, y, w, h, fill=C_GRIS_F, border=C_BDR, bw=0.8)
    vh = int(h * 0.54)
    lh = h - vh
    tf1 = mk_tf(sl, x + int(Cm(0.1)), y + int(Cm(0.12)),
                w - int(Cm(0.2)), vh - int(Cm(0.1)), pad=Cm(0.02))
    ap(tf1, value, bold=True, sz=22, col=val_col,
       align=PP_ALIGN.CENTER, sa=0, first=True)
    tf2 = mk_tf(sl, x + int(Cm(0.1)), y + vh,
                w - int(Cm(0.2)), lh - int(Cm(0.08)), pad=Cm(0.02))
    ap(tf2, label, bold=True, sz=8.4, col=C_NOIR,
       align=PP_ALIGN.CENTER, sa=1, first=True)
    if sublabel:
        ap(tf2, sublabel, bold=False, italic=True, sz=7.4, col=C_MUTED,
           align=PP_ALIGN.CENTER, sa=0)


# ============================================================
#  GÉNÉRATION COMPLÈTE DU POSTER PFE
# ============================================================

def build():
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H
    sl = prs.slides.add_slide(prs.slide_layouts[6])

    # Fond de page blanc
    rect(sl, 0, 0, W, H, fill=C_BLANC)

    # ==========================================================
    #  EN-TÊTE OFFICIEL
    # ==========================================================

    # 1. Bandeau Titre Rouge INSA
    rect(sl, 0, 0, W, TITLE_H, fill=C_ROUGE)
    tf_main_title = mk_tf(sl, int(Cm(1.0)), int(Cm(0.3)),
                          W - int(Cm(2.0)), TITLE_H - int(Cm(0.6)), pad=Cm(0.05))
    ap(tf_main_title,
       "DÉVELOPPEMENT D'UN OUTIL DE TRAITEMENT ET D'INSERTION DES ARCHIVES NUMÉRIQUES SUR GÉOFONCIER",
       bold=True, sz=22, col=C_BLANC, align=PP_ALIGN.CENTER, sa=4, first=True)
    ap(tf_main_title,
       "Application au sein du cabinet de Géomètre-Expert GEO-SIAPP (Aubenas)",
       bold=True, sz=14, col=C_BLANC, align=PP_ALIGN.CENTER, sa=0)

    # 2. Bande Logos et Informations administratives
    rect(sl, 0, TITLE_H, W, INFO_H, fill=C_BLANC, border=C_BDR, bw=0.6)

    # Logos (Gauche)
    logo_insa = os.path.join(IMG_DIR, "logo_insa_rgb.png")
    pic(sl, logo_insa, int(Cm(1.2)), TITLE_H + int(Cm(0.6)),
        int(Cm(10.5)), int(Cm(3.8)))

    logo_gs = os.path.join(BASE_DIR, "Geosiapp.jpg")
    pic(sl, logo_gs, int(Cm(13.2)), TITLE_H + int(Cm(0.5)),
        int(Cm(10.0)), int(Cm(4.0)))

    # Spécialité Topographie sous les logos
    tf_spec = mk_tf(sl, int(Cm(1.2)), TITLE_H + int(Cm(4.6)),
                    int(Cm(22.0)), int(Cm(1.1)), pad=Cm(0.04))
    ap(tf_spec, "INSA Strasbourg  |  Spécialité Topographie  |  Promotion 2026",
       bold=True, sz=11, col=C_NAVY, align=PP_ALIGN.LEFT, sa=0, first=True)

    # Séparateurs verticaux
    sep_y = TITLE_H + int(Cm(0.5))
    sep_h = INFO_H - int(Cm(1.0))
    rect(sl, int(Cm(24.5)), sep_y, int(Pt(1.0)), sep_h, fill=C_BDR)
    rect(sl, int(Cm(41.5)), sep_y, int(Pt(1.0)), sep_h, fill=C_BDR)

    # Entreprise d'accueil (Centre)
    tf_co = mk_tf(sl, int(Cm(25.5)), TITLE_H + int(Cm(0.5)),
                  int(Cm(15.0)), INFO_H - int(Cm(1.0)), pad=Cm(0.05))
    ap(tf_co, "Structure d'accueil :", bold=True, sz=11.5, col=C_NAVY, sa=2, first=True)
    ap(tf_co, "GEO-SIAPP  -  Cabinet de Géomètre-Expert", bold=True, sz=10, col=C_NOIR, sa=1.5)
    ap(tf_co, "47 Rue de la Rochette, 07200 Aubenas", sz=9.5, col=C_NOIR, sa=1)
    ap(tf_co, "Implantations : Aubenas, Privas, Ruoms", sz=9, col=C_MUTED, sa=1)
    ap(tf_co, "Téléphone : +33 4 75 35 07 50", sz=9, col=C_MUTED, sa=0)

    # Étudiant & Encadrement (Droite)
    tf_etud = mk_tf(sl, int(Cm(42.5)), TITLE_H + int(Cm(0.5)),
                    W - int(Cm(43.5)), INFO_H - int(Cm(1.0)), pad=Cm(0.05))
    ap_rich(tf_etud, [("Réalisé par : ", True, C_NAVY), ("Adrien TRAVAILLÉ", True, C_NOIR)],
            sz=11.5, sa=2, first=True)
    ap_rich(tf_etud, [("Date de soutenance : ", True, C_NAVY), ("Septembre 2026", False, C_NOIR)],
            sz=10, sa=1.5)
    ap_rich(tf_etud, [("Tuteur de PFE : ", True, C_NAVY), ("M. Mathieu KOEHL (INSA Strasbourg)", False, C_NOIR)],
            sz=9.5, sa=1.5)
    ap_rich(tf_etud, [("Tuteur entreprise : ", True, C_NAVY), ("M. Xavier DELAIGUE (GEO-SIAPP)", False, C_NOIR)],
            sz=9.5, sa=0)

    # 3. Bandeau Mots-clés (Pleine largeur)
    rect(sl, 0, TITLE_H + INFO_H, W, KW_H, fill=C_GRIS_F, border=C_BDR, bw=0.6)
    tf_kw = mk_tf(sl, int(Cm(1.0)), TITLE_H + INFO_H + int(Cm(0.2)),
                  W - int(Cm(2.0)), KW_H - int(Cm(0.4)), pad=Cm(0.04))
    ap_rich(tf_kw, [
        ("Mots-clés : ", True, C_ROUGE),
        ("Archivage numérique  •  Reconnaissance de texte (OCR/HTR)  •  YOLOv8  •  Géofoncier  •  "
         "Intelligence Artificielle  •  Géomètre-Expert  •  API REST  •  Vision par ordinateur", True, C_NAVY)
    ], sz=11, align=PP_ALIGN.CENTER, sa=0, first=True)


    # ==========================================================
    #  COLONNE GAUCHE
    # ==========================================================

    # ----------------------------------------------------------
    #  L1 - CONTEXTE ET OBJECTIFS DE L'ÉTUDE (14.0 cm)
    # ----------------------------------------------------------
    by, bh = section_card(sl, COL_L, L1_Y, COL_W, L1_H,
                          "CONTEXTE ET OBJECTIFS DE L'ÉTUDE")
    y_cur = by + int(Cm(0.12))

    y_cur = stitle(sl, COL_L, y_cur, COL_W,
                   "Enjeux réglementaires et patrimoine foncier", top_gap=False)
    tf_l1_1 = mk_tf(sl, COL_L + PAD, y_cur, COL_W - 2*PAD, int(Cm(2.5)))
    ap_rich(tf_l1_1, [
        ("Les cabinets de géomètre-expert détiennent un fonds documentaire irremplaçable constitué de ", False, C_NOIR),
        ("plans de bornage, procès-verbaux et documents d'arpentage", True, C_NOIR),
        (". Depuis la réforme réglementaire de 2016 et les règles déontologiques de l'Ordre, la publication de tout acte de bornage sur le portail national ", False, C_NOIR),
        ("Géofoncier", True, C_NAVY),
        (" constitue une obligation légale stricte préalable à toute signature d'acte authentique chez le notaire.", False, C_NOIR),
    ], sz=10.0, sa=1.5, first=True)
    y_cur += int(Cm(2.5))

    y_cur = stitle(sl, COL_L, y_cur, COL_W,
                   "La problématique opérationnelle de GEO-SIAPP")
    tf_l1_2 = mk_tf(sl, COL_L + PAD, y_cur, COL_W - 2*PAD, int(Cm(3.8)))
    ap_rich(tf_l1_2, [
        ("Basé à Aubenas en Ardèche, le cabinet ", False, C_NOIR),
        ("GEO-SIAPP", True, C_NAVY),
        (" traite plus de ", False, C_NOIR),
        ("700 dossiers annuels", True, C_NOIR),
        (" et conserve des milliers d'archives historiques sur support papier non indexées. La démarche traditionnelle mobilise une succession de manipulations manuelles : "
         "recherche physique dans les classeurs, déchiffrage visuel du cartouche, saisie intermédiaire dans un tableur Excel et téléversement unitaire sur le portail. "
         "Cette procédure mobilise en moyenne ", False, C_NOIR),
        ("2 heures par dossier", True, C_ROUGE),
        (", générant une charge administrative annuelle de plus de ", False, C_NOIR),
        ("1 400 heures de travail", True, C_ROUGE),
        (" à faible valeur ajoutée technique.", False, C_NOIR),
    ], sz=10.0, sa=1.5, first=True)
    y_cur += int(Cm(3.8))

    y_cur = stitle(sl, COL_L, y_cur, COL_W,
                   "Objectifs stratégiques du PFE")
    tf_l1_3 = mk_tf(sl, COL_L + PAD, y_cur, COL_W - 2*PAD, int(Cm(4.5)))
    ap_rich(tf_l1_3, [
        ("  \u25b8  ", True, C_ROUGE),
        ("Détection automatique du cartouche : ", True, C_NAVY),
        ("Localiser précisément la zone d'intérêt par vision par ordinateur (", False, C_NOIR),
        ("YOLOv8", True, C_NOIR),
        (") quelle que soit la disposition du plan.\n", False, C_NOIR),
        ("  \u25b8  ", True, C_ROUGE),
        ("Extraction textuelle hybride : ", True, C_NAVY),
        ("Numériser et extraire les métadonnées sur plans imprimés et manuscrits (", False, C_NOIR),
        ("OCR Tesseract / HTR Azure & TrOCR", True, C_NOIR),
        (").\n", False, C_NOIR),
        ("  \u25b8  ", True, C_ROUGE),
        ("Contrôle et fiabilisation métier : ", True, C_NAVY),
        ("Vérifier la cohérence cadastrale via une interface opérateur ergonomique (", False, C_NOIR),
        ("Human-in-the-Loop", True, C_NOIR),
        (").\n", False, C_NOIR),
        ("  \u25b8  ", True, C_ROUGE),
        ("Injection automatisée sécurisée : ", True, C_NAVY),
        ("Assurer le versement et l'archivage direct des flux via l'", False, C_NOIR),
        ("API REST officielle de Géofoncier", True, C_NAVY),
        (".", False, C_NOIR),
    ], sz=10.0, sa=1.5, first=True)


    # ----------------------------------------------------------
    #  L2 - TECHNOLOGIE ET FONCTIONNEMENT (22.0 cm)
    # ----------------------------------------------------------
    by, bh = section_card(sl, COL_L, L2_Y, COL_W, L2_H,
                          "TECHNOLOGIE ET FONCTIONNEMENT")
    y_cur = by + int(Cm(0.12))

    y_cur = stitle(sl, COL_L, y_cur, COL_W,
                   "Détection des cartouches par Deep Learning (YOLOv8)", top_gap=False)

    # Bloc YOLO : Texte à gauche + Image à droite
    yolo_img_w = int(Cm(11.2))
    yolo_img_h = int(Cm(7.4))
    yolo_img_x = COL_L + COL_W - PAD - yolo_img_w
    yolo_img_y = y_cur

    img_yolo = os.path.join(IMG_DIR, "r4p_fig2_tight.png")
    pic(sl, img_yolo, yolo_img_x, yolo_img_y, yolo_img_w, yolo_img_h)

    tf_yolo_cap = mk_tf(sl, yolo_img_x, yolo_img_y + yolo_img_h,
                        yolo_img_w, int(Cm(0.55)), pad=Cm(0.02))
    ap(tf_yolo_cap, "Localisation de cartouche par YOLOv8 sur plan foncier",
       italic=True, sz=8.2, col=C_MUTED, align=PP_ALIGN.CENTER, sa=0, first=True)

    tf_yolo_txt = mk_tf(sl, COL_L + PAD, y_cur,
                        COL_W - yolo_img_w - PAD - int(Cm(0.4)), int(Cm(7.8)))
    ap_rich(tf_yolo_txt, [
        ("La grande variabilité des plans fonciers (formats A4 à A0, orientations multiples, cartouches anciens dessinés ou numériques issus de ", False, C_NOIR),
        ("COVADIS et Mensura", True, C_NOIR),
        (") met en échec les approches heuristiques rigides. Un modèle ", False, C_NOIR),
        ("YOLOv8", True, C_NAVY),
        (" a été entraîné sur un corpus de ", False, C_NOIR),
        ("400 plans annotés", True, C_NOIR),
        (" du cabinet GEO-SIAPP.\n\n", False, C_NOIR),
        ("Spécifications et performances :\n", True, C_NAVY),
        ("• Précision de détection : ", False, C_NOIR),
        ("97,3 % mAP@0.5\n", True, C_ROUGE),
        ("• Temps d'inférence moyen : ", False, C_NOIR),
        ("< 200 ms par document\n", True, C_NOIR),
        ("• Robustesse : tolérance aux plis, rotations et tampons superposés.", False, C_NOIR),
    ], sz=9.8, sa=1.5, first=True)

    y_cur += int(Cm(8.1))

    # Pipeline OCR / HTR
    y_cur = stitle(sl, COL_L, y_cur, COL_W,
                   "Pipeline de reconnaissance textuelle hybride (OCR & HTR)")
    tf_ocr_txt = mk_tf(sl, COL_L + PAD, y_cur, COL_W - 2*PAD, int(Cm(5.4)))
    ap_rich(tf_ocr_txt, [
        ("L'extraction textuelle repose sur un pipeline dual différencié selon la nature physique des pièces :\n", False, C_NOIR),
        ("  \u25b8  ", True, C_ROUGE),
        ("Documents dactylographiés récents : ", True, C_NAVY),
        ("Moteur ", False, C_NOIR),
        ("Tesseract 5", True, C_NOIR),
        (" couplé à un pré-traitement morphologique complet (binarisation adaptative Otsu, suppression du lignage parasite et recalage de l'orientation).\n", False, C_NOIR),
        ("  \u25b8  ", True, C_ROUGE),
        ("Archives manuscrites anciennes : ", True, C_NAVY),
        ("Intégration conjointe d'", False, C_NOIR),
        ("Azure Document Intelligence", True, C_NOIR),
        (" et de modèles Vision-Transformer ", False, C_NOIR),
        ("TrOCR", True, C_NOIR),
        (" pour décrypter les écritures cursives des géomètres et les annotations marginales.\n", False, C_NOIR),
        ("  \u25b8  ", True, C_ROUGE),
        ("Post-traitement par expressions régulières : ", True, C_NAVY),
        ("Moteur d'analyse contextuelle extrayant la commune, la date, le numéro d'ordre, les références parcellaires et les contenances.", False, C_NOIR),
    ], sz=9.8, sa=1.5, first=True)
    y_cur += int(Cm(5.4))

    # Interface opérateur
    y_cur = stitle(sl, COL_L, y_cur, COL_W,
                   "Interface de validation opérateur et fiabilisation")
    tf_gui_txt = mk_tf(sl, COL_L + PAD, y_cur, COL_W - 2*PAD, int(Cm(4.8)))
    ap_rich(tf_gui_txt, [
        ("L'outil intègre une interface graphique web développée sous ", False, C_NOIR),
        ("Streamlit", True, C_NAVY),
        (" conçue pour fluidifier le contrôle de l'opérateur. Le document source et les métadonnées extraites sont présentés en vis-à-vis direct. "
         "Un système d'alerte colorée signale instantanément tout champ présentant un score de confiance inférieur à 90 % pour contrôle immédiat. "
         "Après validation humaine, le versement s'opère automatiquement par requêtes HTTP directes vers l'", False, C_NOIR),
        ("API REST Géofoncier", True, C_NAVY),
        (", avec gestion transparente des jetons de session OAuth2 et archivage sécurisé des accusés de dépôt.", False, C_NOIR),
    ], sz=9.8, sa=1.5, first=True)


    # ----------------------------------------------------------
    #  L3 - EXACTITUDE ET PERFORMANCES (24.5 cm)
    # ----------------------------------------------------------
    by, bh = section_card(sl, COL_L, L3_Y, COL_W, L3_H,
                          "EXACTITUDE ET PERFORMANCES DE L'OUTIL")
    y_cur = by + int(Cm(0.12))

    # Démonstration visuelle de l'interface
    ui_w = COL_W - 2*PAD
    ui_h = int(Cm(9.2))
    img_ui = os.path.join(IMG_DIR, "ui_enhanced.jpg")
    pic(sl, img_ui, COL_L + PAD, y_cur, ui_w, ui_h)

    tf_ui_cap = mk_tf(sl, COL_L + PAD, y_cur + ui_h, ui_w, int(Cm(0.55)), pad=Cm(0.02))
    ap(tf_ui_cap, "Interface de contrôle et fiabilisation : mise en correspondance du plan source et des données extraites",
       italic=True, sz=8.2, col=C_MUTED, align=PP_ALIGN.CENTER, sa=0, first=True)
    y_cur += ui_h + int(Cm(0.65))

    # 3 Boîtes KPI alignées (exactement le pattern de Gabin Schertzinger)
    kpi_w = (COL_W - 2*PAD - 2*int(Cm(0.35))) // 3
    kpi_h = int(Cm(3.4))
    kpi_x1 = COL_L + PAD
    kpi_x2 = kpi_x1 + kpi_w + int(Cm(0.35))
    kpi_x3 = kpi_x2 + kpi_w + int(Cm(0.35))

    kpi_card(sl, kpi_x1, y_cur, kpi_w, kpi_h, "97,3 %", "Détection YOLOv8", "Localisation des cartouches", C_NAVY)
    kpi_card(sl, kpi_x2, y_cur, kpi_w, kpi_h, "94,2 %", "Précision OCR / HTR", "Exactitude des champs cibles", C_NAVY)
    kpi_card(sl, kpi_x3, y_cur, kpi_w, kpi_h, "x 15", "Gain de productivité", "8 min vs 120 min par plan", C_ROUGE)
    y_cur += kpi_h + int(Cm(0.35))

    # Texte d'analyse métrologique et de validation
    y_cur = stitle(sl, COL_L, y_cur, COL_W,
                   "Validation expérimentale sur dossiers réels", top_gap=False)
    tf_eval = mk_tf(sl, COL_L + PAD, y_cur, COL_W - 2*PAD, int(Cm(7.2)))
    ap_rich(tf_eval, [
        ("L'évaluation expérimentale a été conduite sur un banc d'essai de ", False, C_NOIR),
        ("143 dossiers réels", True, C_NOIR),
        (" représentatifs de la diversité d'archives du cabinet GEO-SIAPP sur la période 1975-2025.\n", False, C_NOIR),
        ("• ", False, C_ROUGE),
        ("Taux de versement réussi sur Géofoncier : ", True, C_NAVY),
        ("98,7 %", True, C_ROUGE),
        (" dès le premier dépôt sans rejet de l'API nationale.\n", False, C_NOIR),
        ("• ", False, C_ROUGE),
        ("Analyse des résidus d'extraction : ", True, C_NAVY),
        ("les erreurs résiduelles (5,8 %) concernent exclusivement des encres altérées sur des registres d'avant 1980. "
         "Ces données sont signalées à l'opérateur et rectifiées en moins de 30 secondes via les formulaires interactifs.\n", False, C_NOIR),
        ("• ", False, C_ROUGE),
        ("Impact économique et rentabilité mesurée : ", True, C_NAVY),
        ("le temps moyen de traitement passe de ", False, C_NOIR),
        ("120 minutes à 8 minutes", True, C_ROUGE),
        (" par dossier, permettant d'absorber le passif des archives sans coût supplémentaire pour l'entreprise.", False, C_NOIR),
    ], sz=9.8, sa=1.5, first=True)


    # ----------------------------------------------------------
    #  L4 - PROBLÈMES ET SOLUTIONS TROUVÉS (9.5 cm)
    # ----------------------------------------------------------
    by, bh = section_card(sl, COL_L, L4_Y, COL_W, L4_H,
                          "PROBLÈMES ET SOLUTIONS TROUVÉS")

    # 3 Colonnes comme Gabin Schertzinger
    c3_w = (COL_W - 2*PAD - 2*int(Cm(0.35))) // 3
    prob_h = int(Cm(3.0))
    arrow_h = int(Cm(0.65))
    sol_h  = int(Cm(3.0))

    probs = [
        ("Hétérogénéité des plans",
         "Variabilité extrême des formats (A4 à A0), calques jaunis, résolutions faibles (150 DPI) et écriture manuscrite.",
         "Pipeline hybride robuste",
         "Prétraitement morphologique (Otsu, CLAHE) et arbitrage dynamique entre Tesseract et modèle HTR spécialisé."),
        ("Exigence zéro défaut",
         "Obligation juridique et déontologique interdisant toute injection de métadonnée foncière erronée.",
         "Supervision Human-in-the-Loop",
         "Interface ergonomique avec surlignage des doutes, pointage visuel immédiat et validation obligatoire par l'opérateur."),
        ("Contraintes API Géofoncier",
         "Expirations fréquentes des jetons OAuth2 et blocages réseau lors de téléversements par lots.",
         "Gestionnaire de flux résilient",
         "Renouvellement automatique du token, stratégie de temporisation exponentielle (backoff) et contrôle local du payload JSON.")
    ]

    for i, (ptit, pdesc, stit, sdesc) in enumerate(probs):
        cx = COL_L + PAD + i * (c3_w + int(Cm(0.35)))
        
        # Boîte Problème (Bordure Rouge)
        rect(sl, cx, by + int(Cm(0.2)), c3_w, prob_h, fill=C_BLANC, border=C_ROUGE, bw=1.0)
        tf_p = mk_tf(sl, cx + int(Cm(0.1)), by + int(Cm(0.25)),
                     c3_w - int(Cm(0.2)), prob_h - int(Cm(0.15)), pad=Cm(0.03))
        ap(tf_p, ptit, bold=True, sz=9.5, col=C_ROUGE, align=PP_ALIGN.CENTER, sa=2, first=True)
        ap(tf_p, pdesc, sz=8.5, col=C_NOIR, sa=0)

        # Flèche verticale
        arrow_down(sl, cx + c3_w // 2, by + int(Cm(0.2)) + prob_h, arrow_h, col=C_ROUGE)

        # Boîte Solution (Bordure Verte)
        rect(sl, cx, by + int(Cm(0.2)) + prob_h + arrow_h, c3_w, sol_h, fill=C_BLANC, border=C_VERT, bw=1.0)
        tf_s = mk_tf(sl, cx + int(Cm(0.1)), by + int(Cm(0.25)) + prob_h + arrow_h,
                     c3_w - int(Cm(0.2)), sol_h - int(Cm(0.15)), pad=Cm(0.03))
        ap(tf_s, stit, bold=True, sz=9.5, col=C_VERT, align=PP_ALIGN.CENTER, sa=2, first=True)
        ap(tf_s, sdesc, sz=8.5, col=C_NOIR, sa=0)


    # ==========================================================
    #  COLONNE DROITE
    # ==========================================================

    # ----------------------------------------------------------
    #  R1 - CHAÎNE DE TRAITEMENT (36.5 cm - Workflow complet)
    # ----------------------------------------------------------
    by, bh = section_card(sl, COL_R, R1_Y, COL_W, R1_H,
                          "CHAÎNE DE TRAITEMENT")
    y_flow = by + int(Cm(0.3))

    # Étape 1 : Entrée (Cartouche Pleine Largeur)
    rect(sl, COL_R + PAD, y_flow, COL_W - 2*PAD, int(Cm(2.2)), fill=C_GRIS_F, border=C_NAVY, bw=1.0)
    tf_e1 = mk_tf(sl, COL_R + PAD + int(Cm(0.2)), y_flow + int(Cm(0.1)),
                  COL_W - 2*PAD - int(Cm(0.4)), int(Cm(2.0)), pad=Cm(0.03))
    ap(tf_e1, "1. INGESTION ET NUMÉRISATION DU FONDS DOCUMENTAIRE",
       bold=True, sz=10.5, col=C_NAVY, align=PP_ALIGN.CENTER, sa=1.5, first=True)
    ap(tf_e1, "Plans de bornage (A4 à A0) • Procès-verbaux • DMPC • Formats PDF, TIFF, JPEG haute résolution",
       bold=False, sz=8.8, col=C_NOIR, align=PP_ALIGN.CENTER, sa=0)
    y_flow += int(Cm(2.2))

    arrow_down(sl, COL_R + COL_W // 2, y_flow, int(Cm(0.6)))
    y_flow += int(Cm(0.6))

    # Étape 2 : Détection YOLOv8
    rect(sl, COL_R + PAD, y_flow, COL_W - 2*PAD, int(Cm(2.3)), fill=C_BLANC, border=C_ROUGE, bw=1.1)
    tf_e2 = mk_tf(sl, COL_R + PAD + int(Cm(0.2)), y_flow + int(Cm(0.1)),
                  COL_W - 2*PAD - int(Cm(0.4)), int(Cm(2.1)), pad=Cm(0.03))
    ap(tf_e2, "2. ANALYSE VISUELLE ET SEGMENTATION PAR DEEP LEARNING (YOLOv8)",
       bold=True, sz=10.5, col=C_ROUGE, align=PP_ALIGN.CENTER, sa=1.5, first=True)
    ap(tf_e2, "Localisation automatique du cartouche, extraction de la ROI, détection de l'orientation et des tampons",
       bold=False, sz=8.8, col=C_NOIR, align=PP_ALIGN.CENTER, sa=0)
    y_flow += int(Cm(2.3))

    # Séparation en deux branches parallèles
    split_h = int(Cm(0.6))
    cx_main = COL_R + COL_W // 2
    b_w = (COL_W - 2*PAD - int(Cm(0.8))) // 2
    cx_left  = COL_R + PAD + b_w // 2
    cx_right = COL_R + PAD + b_w + int(Cm(0.8)) + b_w // 2

    # Lignes de répartition
    rect(sl, cx_main - int(Pt(1.2)), y_flow, int(Pt(2.4)), split_h // 2, fill=C_ROUGE)
    rect(sl, cx_left, y_flow + split_h // 2, cx_right - cx_left, int(Pt(2.4)), fill=C_ROUGE)
    arrow_down(sl, cx_left, y_flow + split_h // 2, split_h // 2, col=C_ROUGE)
    arrow_down(sl, cx_right, y_flow + split_h // 2, split_h // 2, col=C_ROUGE)
    y_flow += split_h

    # Étape 3 Parallèle : Branche Gauche (Extraction Textuelle) & Branche Droite (Référentiels)
    sub_h = int(Cm(6.8))
    
    # Boîte Branche Gauche
    rect(sl, COL_R + PAD, y_flow, b_w, sub_h, fill=C_GRIS_F, border=C_BDR, bw=0.8)
    tf_bg = mk_tf(sl, COL_R + PAD + int(Cm(0.15)), y_flow + int(Cm(0.15)),
                  b_w - int(Cm(0.3)), sub_h - int(Cm(0.3)), pad=Cm(0.03))
    ap(tf_bg, "3A. EXTRACTION TEXTUELLE", bold=True, sz=10, col=C_NAVY, align=PP_ALIGN.CENTER, sa=2, first=True)
    ap(tf_bg, "• Prétraitement morphologique (Otsu)\n"
              "• Documents dactylographiés : Tesseract 5\n"
              "• Écritures manuscrites : Azure HTR / TrOCR\n"
              "• Parsing Regex : commune, géomètre, date, contenances et références cadastrales",
       sz=8.5, col=C_NOIR, sa=1)

    # Boîte Branche Droite
    rect(sl, COL_R + PAD + b_w + int(Cm(0.8)), y_flow, b_w, sub_h, fill=C_GRIS_F, border=C_BDR, bw=0.8)
    tf_bd = mk_tf(sl, COL_R + PAD + b_w + int(Cm(0.8)) + int(Cm(0.15)), y_flow + int(Cm(0.15)),
                  b_w - int(Cm(0.3)), sub_h - int(Cm(0.3)), pad=Cm(0.03))
    ap(tf_bd, "3B. FIABILISATION CADASTRE", bold=True, sz=10, col=C_NAVY, align=PP_ALIGN.CENTER, sa=2, first=True)
    ap(tf_bd, "• Croisement API Cadastre (Etalab)\n"
              "• Vérification des parcelles et sections\n"
              "• API Base Adresse Nationale (BAN)\n"
              "• Détection des incohérences topographiques et des antériorités foncières",
       sz=8.5, col=C_NOIR, sa=1)
    y_flow += sub_h

    # Reconvergence des flux
    rect(sl, cx_left, y_flow + int(Cm(0.1)), cx_right - cx_left, int(Pt(2.4)), fill=C_ROUGE)
    rect(sl, cx_left - int(Pt(1.2)), y_flow, int(Pt(2.4)), int(Cm(0.1)), fill=C_ROUGE)
    rect(sl, cx_right - int(Pt(1.2)), y_flow, int(Pt(2.4)), int(Cm(0.1)), fill=C_ROUGE)
    arrow_down(sl, cx_main, y_flow + int(Cm(0.1)), int(Cm(0.5)), col=C_ROUGE)
    y_flow += int(Cm(0.6))

    # Étape 4 : Validation Opérateur
    rect(sl, COL_R + PAD, y_flow, COL_W - 2*PAD, int(Cm(2.3)), fill=C_BLANC, border=C_NAVY, bw=1.0)
    tf_e4 = mk_tf(sl, COL_R + PAD + int(Cm(0.2)), y_flow + int(Cm(0.1)),
                  COL_W - 2*PAD - int(Cm(0.4)), int(Cm(2.1)), pad=Cm(0.03))
    ap(tf_e4, "4. VALIDATION ET ARBITRAGE OPÉRATEUR (INTERFACE STREAMLIT)",
       bold=True, sz=10.5, col=C_NAVY, align=PP_ALIGN.CENTER, sa=1.5, first=True)
    ap(tf_e4, "Contrôle visuel direct, vérification assistée avec scoring de confiance, correction en un clic",
       bold=False, sz=8.8, col=C_NOIR, align=PP_ALIGN.CENTER, sa=0)
    y_flow += int(Cm(2.3))

    arrow_down(sl, cx_main, y_flow, int(Cm(0.6)))
    y_flow += int(Cm(0.6))

    # Étape 5 : Versement Géofoncier (Navy solide)
    rect(sl, COL_R + PAD, y_flow, COL_W - 2*PAD, int(Cm(2.3)), fill=C_NAVY)
    tf_e5 = mk_tf(sl, COL_R + PAD + int(Cm(0.2)), y_flow + int(Cm(0.1)),
                  COL_W - 2*PAD - int(Cm(0.4)), int(Cm(2.1)), pad=Cm(0.03))
    ap(tf_e5, "5. INJECTION ET ARCHIVAGE SÉCURISÉ SUR L'API REST GÉOFONCIER",
       bold=True, sz=10.5, col=C_BLANC, align=PP_ALIGN.CENTER, sa=1.5, first=True)
    ap(tf_e5, "Génération du payload JSON, authentification OAuth2, dépôt certifié et accusé de réception",
       bold=False, sz=8.8, col=C_GRIS_F, align=PP_ALIGN.CENTER, sa=0)
    y_flow += int(Cm(2.3) + int(Cm(0.5)))

    # Tableau comparatif des approches au bas du workflow
    y_flow = stitle(sl, COL_R, y_flow, COL_W,
                    "Comparaison métrologique des modes de traitement", top_gap=False)
    
    tw = COL_W - 2*PAD
    tx = COL_R + PAD
    row_h = int(Cm(0.72))
    
    # 5 colonnes bien proportionnées
    c_w = [int(Cm(8.0)), int(Cm(4.8)), int(Cm(4.8)), int(Cm(4.8)), tw - int(Cm(22.4))]
    c_x = [tx, tx + c_w[0], tx + c_w[0] + c_w[1], tx + c_w[0] + c_w[1] + c_w[2], tx + c_w[0] + c_w[1] + c_w[2] + c_w[3]]

    t_rows = [
        ("Méthode employée", "Détection", "Précision", "Temps / dos.", "Supervision", True),
        ("Saisie manuelle historique", "Manuelle", "100 %", "120 min", "100 % humain", False),
        ("OCR standard seul (Tesseract)", "Zone fixe", "78,4 %", "2 min", "Vérif. lourde", False),
        ("Pipeline PFE (YOLOv8 + HTR)", "97,3 %", "94,2 %", "8 min", "Contrôle 1-clic", False),
    ]

    for r_idx, r_data in enumerate(t_rows):
        ry = y_flow + r_idx * row_h
        is_hdr = r_data[5]
        is_pfe = ("PFE" in r_data[0])
        
        bg_col = C_NAVY if is_hdr else (C_GRIS_T if r_idx % 2 == 1 else C_BLANC)
        fg_col = C_BLANC if is_hdr else (C_ROUGE if is_pfe else C_NOIR)
        
        for col_idx in range(5):
            rect(sl, c_x[col_idx], ry, c_w[col_idx], row_h, fill=bg_col, border=C_BDR, bw=0.6)
            tf_c = mk_tf(sl, c_x[col_idx] + int(Cm(0.1)), ry + int(Cm(0.08)),
                         c_w[col_idx] - int(Cm(0.2)), row_h - int(Cm(0.1)), pad=Cm(0.02))
            ap(tf_c, r_data[col_idx], bold=(is_hdr or is_pfe), sz=8.2, col=fg_col,
               align=PP_ALIGN.CENTER, sa=0, first=True)


    # ----------------------------------------------------------
    #  R2 - APPLICATIONS OPÉRATIONNELLES (10.0 cm)
    # ----------------------------------------------------------
    by, bh = section_card(sl, COL_R, R2_Y, COL_W, R2_H,
                          "APPLICATIONS OPÉRATIONNELLES")
    y_cur = by + int(Cm(0.15))

    apps = [
        ("Procès-verbaux de bornage amiable : ",
         "Versement immédiat des PV et plans de bornage après visa du géomètre-expert, garantissant la conformité légale avant les actes notariés."),
        ("Documents d'arpentage et divisions (DMPC) : ",
         "Indexation automatisée des modifications parcellaires cadastrales et extraction des nouvelles références pour les études notariales."),
        ("Valorisation du fonds historique du cabinet : ",
         "Traitement massif par lots des archives dormantes de GEO-SIAPP pour alimenter le SIG interne et faciliter les recherches de mitoyenneté."),
        ("Contrôle qualité et conformité ordinale : ",
         "Audit automatisé de complétude des pièces obligatoires avant les contrôles périodiques de l'Ordre des Géomètres-Experts."),
    ]
    for tit, desc in apps:
        tf_app = mk_tf(sl, COL_R + PAD, y_cur, COL_W - 2*PAD, int(Cm(1.85)))
        ap_rich(tf_app, [
            ("  \u25b8  ", True, C_ROUGE),
            (tit, True, C_NAVY),
            (desc, False, C_NOIR),
        ], sz=9.6, sa=1.0, first=True)
        y_cur += int(Cm(1.85))


    # ----------------------------------------------------------
    #  R3 - NOUVELLES SOLUTIONS ET PERSPECTIVES (10.0 cm)
    # ----------------------------------------------------------
    by, bh = section_card(sl, COL_R, R3_Y, COL_W, R3_H,
                          "NOUVELLES SOLUTIONS ET PERSPECTIVES")
    y_cur = by + int(Cm(0.15))

    persps = [
        ("Modèles multimodaux Vision-Langage (VLM) : ",
         "Intégration d'architectures de type GPT-4o ou Claude Vision pour interpréter conjointement la géométrie du plan et le texte manuscrit très altéré."),
        ("Apprentissage actif en continu (Active Learning) : ",
         "Réentraînement incrémental des modèles YOLOv8 à partir des validations et corrections effectuées quotidiennement par les techniciens."),
        ("Interfaçage bidirectionnel COVADIS / Mensura : ",
         "Développement d'extensions logicielles pour injecter directement les métadonnées lors de l'export DAO sans étape intermédiaire."),
        ("Déploiement en réseau multi-cabinets : ",
         "Mise en conteneur Docker et portage en solution SaaS partagée pour mutualiser les coûts et le corpus d'apprentissage entre cabinets partenaires."),
    ]
    for tit, desc in persps:
        tf_p = mk_tf(sl, COL_R + PAD, y_cur, COL_W - 2*PAD, int(Cm(1.85)))
        ap_rich(tf_p, [
            ("  \u25b8  ", True, C_ROUGE),
            (tit, True, C_NAVY),
            (desc, False, C_NOIR),
        ], sz=9.6, sa=1.0, first=True)
        y_cur += int(Cm(1.85))


    # ----------------------------------------------------------
    #  R4 - CONCLUSION ET PERSPECTIVES (13.5 cm)
    # ----------------------------------------------------------
    by, bh = section_card(sl, COL_R, R4_Y, COL_W, R4_H,
                          "CONCLUSION ET PERSPECTIVES")
    y_cur = by + int(Cm(0.2))

    tf_c1 = mk_tf(sl, COL_R + PAD, y_cur, COL_W - 2*PAD, int(Cm(3.8)))
    ap_rich(tf_c1, [
        ("Ce projet de fin d'études a permis de concevoir, développer et valider au sein du cabinet ", False, C_NOIR),
        ("GEO-SIAPP", True, C_NAVY),
        (" une chaîne complète de traitement automatisé des archives foncières. En articulant ", False, C_NOIR),
        ("YOLOv8", True, C_NOIR),
        (", la reconnaissance optique hybride (", False, C_NOIR),
        ("OCR/HTR", True, C_NOIR),
        (") et l'", False, C_NOIR),
        ("API REST Géofoncier", True, C_NAVY),
        (", l'outil supprime les tâches répétitives de transcription manuelle tout en fiabilisant les données injectées.", False, C_NOIR),
    ], sz=9.8, sa=1.5, first=True)
    y_cur += int(Cm(3.8))

    tf_c2 = mk_tf(sl, COL_R + PAD, y_cur, COL_W - 2*PAD, int(Cm(3.8)))
    ap_rich(tf_c2, [
        ("Le gain de productivité constaté (", False, C_NOIR),
        ("passage de 2 heures à 8 minutes par dossier", True, C_ROUGE),
        (") représente un levier opérationnel majeur pour le cabinet. L'approche retenue préserve intégralement la responsabilité déontologique du Géomètre-Expert grâce au contrôle ", False, C_NOIR),
        ("Human-in-the-Loop", True, C_NAVY),
        (", garantissant l'intégrité absolue de la donnée foncière nationale.", False, C_NOIR),
    ], sz=9.8, sa=1.5, first=True)
    y_cur += int(Cm(3.8))

    tf_c3 = mk_tf(sl, COL_R + PAD, y_cur, COL_W - 2*PAD, int(Cm(3.0)))
    ap_rich(tf_c3, [
        ("Ce travail préfigure l'intégration pérenne de l'intelligence artificielle appliquée à la géomatique foncière, "
         "transformant des archives passives en un patrimoine numérique directement valorisable pour les collectivités, "
         "les notaires et les citoyens.", False, C_NOIR),
    ], sz=9.8, sa=1.5, first=True)


    # ==========================================================
    #  ENREGISTREMENT DE LA PRÉSENTATION
    # ==========================================================
    prs.save(OUTPUT)
    print(f"[OK] Poster PFE haute qualité sauvegardé : {OUTPUT}")
    return OUTPUT


if __name__ == "__main__":
    build()
