# -*- coding: utf-8 -*-
"""
Générateur — Partie 1 : Slides 1 à 12
Adrien TRAVAILLÉ — Diplôme d'Ingénieur Topographe INSA Strasbourg

Contenu :
  Slide  1 : Page de titre (fond bleu nuit)
  Slide  2 : Sommaire visuel (8 tuiles)
  Slide  3 : Problématique (1 question + 3 problèmes)
  Slide  4 : Transition Chapitre 1
  Slide  5 : 1.1 Structure d'accueil GEO-SIAPP
  Slide  6 : 1.2 + 1.3 Cadre légal & contexte
  Slide  7 : Transition Chapitre 2
  Slide  8 : 2.1 Cadastre & propriété foncière
  Slide  9 : 2.2 + 2.3 API Géofoncier & fonds SIAPP
  Slide 10 : 2.4 Registres manuscrits & filiation parcellaire
  Slide 11 : Transition Chapitre 3 (préparation état de l'art)
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

from generate_deck_core import (
    C_NAVY, C_NAVY_LIGHT, C_NAVY_MID, C_ACCENT, C_ACCENT_LIGHT,
    C_SUCCESS, C_SUCCESS_BG, C_WARN, C_WARN_BG,
    C_TEXT_DARK, C_TEXT_MUTED, C_TEXT_BODY, C_BORDER, C_WHITE,
    CHAPTERS, CHAPTER_COLORS,
    apply_background, apply_dark_background,
    add_navigation_bars, add_slide_header, add_footer, add_transition_slide,
    draw_styled_card, draw_framed_image, draw_kpi_card, add_centered_text,
)

TOTAL_SLIDES = 35


def build_part1_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.500)
    blank = prs.slide_layouts[6]

    # =========================================================================
    # SLIDE 1 — PAGE DE TITRE
    # Fond bleu nuit plein, logos en haut, grand titre centré, info candidat
    # =========================================================================
    s1 = prs.slides.add_slide(blank)
    apply_dark_background(s1, C_NAVY)

    # Bande de logos en haut
    top_bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                  Inches(0), Inches(0),
                                  Inches(13.333), Inches(1.20))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = C_NAVY_MID
    top_bar.line.fill.background()

    # Logo INSA (gauche)
    logo_insa = "Logo_INSAStrasbourg.jpg"
    if os.path.exists(logo_insa):
        from PIL import Image
        with Image.open(logo_insa) as img:
            w, h = img.size
        sc = min(2.60 / w, 0.90 / h)
        s1.shapes.add_picture(logo_insa,
                              Inches(0.40), Inches(0.15),
                              Inches(w * sc), Inches(h * sc))

    # Logo GEO-SIAPP (droite)
    logo_geo = "Geosiapp.jpg"
    if os.path.exists(logo_geo):
        from PIL import Image
        with Image.open(logo_geo) as img:
            w, h = img.size
        sc = min(2.00 / w, 0.90 / h)
        s1.shapes.add_picture(logo_geo,
                              Inches(10.933), Inches(0.15),
                              Inches(w * sc), Inches(h * sc))

    # Label institution centré dans la barre
    inst_box = s1.shapes.add_textbox(Inches(2.80), Inches(0.20), Inches(7.733), Inches(0.80))
    itf = inst_box.text_frame
    itf.word_wrap = True
    itf.margin_left = itf.margin_top = itf.margin_right = itf.margin_bottom = 0
    ip = itf.paragraphs[0]
    ip.alignment = PP_ALIGN.CENTER
    ip.text = "INSA STRASBOURG — SPÉCIALITÉ TOPOGRAPHIE"
    ip.font.size = Pt(9.5)
    ip.font.bold = True
    ip.font.color.rgb = RGBColor(180, 210, 240)
    ip2 = itf.add_paragraph()
    ip2.alignment = PP_ALIGN.CENTER
    ip2.text = "Projet de Fin d'Études — Soutenance du 24 septembre 2026"
    ip2.font.size = Pt(8.5)
    ip2.font.color.rgb = RGBColor(130, 170, 220)
    ip2.space_before = Pt(4)

    # Grand titre
    title_box = s1.shapes.add_textbox(Inches(0.80), Inches(1.70), Inches(11.733), Inches(1.70))
    ttf = title_box.text_frame
    ttf.word_wrap = True
    ttf.margin_left = ttf.margin_top = ttf.margin_right = ttf.margin_bottom = 0
    tp0 = ttf.paragraphs[0]
    tp0.text = "Traitement automatisé et versement"
    tp0.alignment = PP_ALIGN.CENTER
    tp0.font.size = Pt(30.0)
    tp0.font.bold = True
    tp0.font.color.rgb = C_WHITE

    tp1 = ttf.add_paragraph()
    tp1.text = "des archives foncières sur Géofoncier"
    tp1.alignment = PP_ALIGN.CENTER
    tp1.font.size = Pt(30.0)
    tp1.font.bold = True
    tp1.font.color.rgb = C_WHITE
    tp1.space_before = Pt(4)

    tp2 = ttf.add_paragraph()
    tp2.text = "au sein d'un cabinet de Géomètre-Expert"
    tp2.alignment = PP_ALIGN.CENTER
    tp2.font.size = Pt(22.0)
    tp2.font.bold = False
    tp2.font.color.rgb = RGBColor(130, 190, 240)
    tp2.space_before = Pt(8)

    # Sous-titre technique
    sub_box = s1.shapes.add_textbox(Inches(1.50), Inches(3.55), Inches(10.333), Inches(0.40))
    stf = sub_box.text_frame
    stf.margin_left = stf.margin_top = stf.margin_right = stf.margin_bottom = 0
    sp = stf.paragraphs[0]
    sp.text = "Vision par ordinateur, modèles de langue légers, fiabilisation et intégration API REST"
    sp.alignment = PP_ALIGN.CENTER
    sp.font.size = Pt(11.0)
    sp.font.color.rgb = RGBColor(100, 155, 210)

    # Ligne séparatrice
    sep = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                               Inches(3.8), Inches(4.05), Inches(5.733), Inches(0.025))
    sep.fill.solid()
    sep.fill.fore_color.rgb = C_ACCENT
    sep.line.fill.background()

    # Bloc info candidat / encadrement / entreprise (3 colonnes)
    infos = [
        ("CANDIDAT", "Adrien TRAVAILLÉ", "Élève-ingénieur Topographe\nINSA Strasbourg — Promo 2026"),
        ("ENCADREMENT", "Mathieu KOEHL", "Directeur du PFE\nEnseignant-chercheur INSA\nLaboratoire ICube (TRIO)"),
        ("ENTREPRISE D'ACCUEIL", "Cabinet GEO-SIAPP", "Tuteur : Gaëtan HAGUE\nGéomètre-Expert Associé\nAubenas (Ardèche)"),
    ]
    col_w = 3.60
    col_starts = [1.00, 4.867, 8.733]
    for (label, name, detail), cx in zip(infos, col_starts):
        box = s1.shapes.add_textbox(Inches(cx), Inches(4.20), Inches(col_w), Inches(2.80))
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        lp = tf.paragraphs[0]
        lp.text = label
        lp.alignment = PP_ALIGN.CENTER
        lp.font.size = Pt(8.0)
        lp.font.bold = True
        lp.font.color.rgb = C_ACCENT
        np_ = tf.add_paragraph()
        np_.text = name
        np_.alignment = PP_ALIGN.CENTER
        np_.font.size = Pt(13.5)
        np_.font.bold = True
        np_.font.color.rgb = C_WHITE
        np_.space_before = Pt(5)
        dp = tf.add_paragraph()
        dp.text = detail
        dp.alignment = PP_ALIGN.CENTER
        dp.font.size = Pt(8.5)
        dp.font.color.rgb = RGBColor(140, 175, 220)
        dp.space_before = Pt(5)

    # Notes
    s1.notes_slide.notes_text_frame.text = (
        "Monsieur le Président du jury, mesdames et messieurs les membres du jury.\n"
        "J'ai le plaisir de vous présenter aujourd'hui mes travaux de Projet de Fin d'Études intitulé : "
        "Traitement automatisé et versement des archives foncières sur Géofoncier au sein d'un cabinet de Géomètre-Expert.\n"
        "Ce projet a été réalisé de février à juillet 2026 au sein du cabinet GEO-SIAPP à Aubenas, en Ardèche, "
        "sous la direction de M. Mathieu Koehl et le tutorat de M. Gaëtan Hague."
    )

    # =========================================================================
    # SLIDE 2 — SOMMAIRE VISUEL (8 tuiles)
    # =========================================================================
    s2 = prs.slides.add_slide(blank)
    apply_dark_background(s2, C_NAVY)

    # Titre
    add_centered_text(s2, "PLAN DE LA PRÉSENTATION", y=0.30, font_size=11,
                      bold=True, color=RGBColor(130, 190, 240))
    add_centered_text(s2, "Sommaire", y=0.55, font_size=26, bold=True, color=C_WHITE)

    # 8 tuiles (2 rangées × 4)
    chap_labels = [
        ("1", "Introduction\n& Contexte"),
        ("2", "Analyse\nMétier & Données"),
        ("3", "État\nde l'Art"),
        ("4", "Architecture\n& Pipeline"),
        ("5", "Validation\n& Résultats"),
        ("6", "Intégration\nGéofoncier"),
        ("7", "Limites\n& Perspectives"),
        ("8", "Conclusion"),
    ]
    tile_w = 2.90
    tile_h = 2.20
    gap = 0.12
    start_x = (13.333 - (4 * tile_w + 3 * gap)) / 2
    rows = [(1.28, 0), (3.66, 4)]  # (y_start, offset_idx)

    for row_y, offset in rows:
        for col in range(4):
            idx = offset + col
            if idx >= len(chap_labels):
                break
            num, label = chap_labels[idx]
            x = start_x + col * (tile_w + gap)
            color = CHAPTER_COLORS[idx]
            # Pour les tuiles 1 et 8 (bleu nuit), ajouter une bordure visible
            has_border = (color == C_NAVY)

            tile = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                       Inches(x), Inches(row_y),
                                       Inches(tile_w), Inches(tile_h))
            tile.fill.solid()
            tile.fill.fore_color.rgb = color
            if has_border:
                tile.line.color.rgb = C_ACCENT
                tile.line.width = Pt(2.0)
            else:
                tile.line.fill.background()

            # Numéro
            num_box = s2.shapes.add_textbox(Inches(x + 0.14), Inches(row_y + 0.14),
                                             Inches(0.50), Inches(0.44))
            ntf = num_box.text_frame
            ntf.margin_left = ntf.margin_top = ntf.margin_right = ntf.margin_bottom = 0
            np2 = ntf.paragraphs[0]
            np2.text = num
            np2.font.size = Pt(22.0)
            np2.font.bold = True
            np2.font.color.rgb = RGBColor(255, 255, 255)

            # Label
            lbl_box = s2.shapes.add_textbox(Inches(x + 0.14), Inches(row_y + 0.66),
                                             Inches(tile_w - 0.28), Inches(1.30))
            ltf = lbl_box.text_frame
            ltf.word_wrap = True
            ltf.margin_left = ltf.margin_top = ltf.margin_right = ltf.margin_bottom = 0
            lp2 = ltf.paragraphs[0]
            lp2.text = label
            lp2.font.size = Pt(11.5)
            lp2.font.bold = True
            lp2.font.color.rgb = C_WHITE

    s2.notes_slide.notes_text_frame.text = (
        "Voici le plan de ma présentation, organisé en 8 chapitres fidèles à mon mémoire.\n"
        "Je commencerai par le contexte et la structure d'accueil, puis je détaillerai l'analyse "
        "du métier et des données, avant de présenter l'état de l'art des technologies mobilisées.\n"
        "J'entrerai ensuite dans le cœur technique avec l'architecture du pipeline, "
        "puis les résultats de validation et la démonstration vidéo.\n"
        "Je terminerai par l'intégration à l'API Géofoncier, les limites du projet et la conclusion."
    )

    # =========================================================================
    # SLIDE 3 — PROBLÉMATIQUE
    # Grande question centrale + 3 problèmes concrets illustrés
    # =========================================================================
    s3 = prs.slides.add_slide(blank)
    apply_dark_background(s3, C_NAVY)

    # Grande question
    q_box = s3.shapes.add_textbox(Inches(0.80), Inches(0.55), Inches(11.733), Inches(1.10))
    qtf = q_box.text_frame
    qtf.word_wrap = True
    qtf.margin_left = qtf.margin_top = qtf.margin_right = qtf.margin_bottom = 0
    qp = qtf.paragraphs[0]
    qp.text = "Comment automatiser le traitement et le versement"
    qp.alignment = PP_ALIGN.CENTER
    qp.font.size = Pt(22.0)
    qp.font.bold = True
    qp.font.color.rgb = C_WHITE
    qp2 = qtf.add_paragraph()
    qp2.text = "de 50 ans d'archives foncières manuscrites sur Géofoncier ?"
    qp2.alignment = PP_ALIGN.CENTER
    qp2.font.size = Pt(22.0)
    qp2.font.bold = True
    qp2.font.color.rgb = C_ACCENT
    qp2.space_before = Pt(4)

    # Séparateur
    sep2 = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                Inches(4.5), Inches(1.80), Inches(4.333), Inches(0.025))
    sep2.fill.solid()
    sep2.fill.fore_color.rgb = C_ACCENT
    sep2.line.fill.background()

    # 3 blocs-problèmes
    problems = [
        (C_ACCENT, "Registres illisibles",
         "Les actes fonciers manuscrits cumulent abréviations, encre fanée et écriture cursive.",
         "Ecriture sans IA illisible"),
        (RGBColor(109, 40, 217), "Communes inconnues",
         "Des noms de communes abrégés (ex. La Chap./A.) sont incompréhensibles hors contexte ardéchois.",
         "Ambiguite commune"),
        (C_SUCCESS, "Versement 100% manuel",
         "Chaque dossier nécessite 15 à 30 min de recherche et une saisie manuelle sur Géofoncier.",
         "Cout manuel versement"),
    ]
    card_w = 3.80
    card_h = 3.60
    gap3 = 0.27
    total_cards = 3 * card_w + 2 * gap3
    start_x3 = (13.333 - total_cards) / 2

    for i, (col, title_p, desc_p, _) in enumerate(problems):
        cx = start_x3 + i * (card_w + gap3)
        card = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                   Inches(cx), Inches(2.10),
                                   Inches(card_w), Inches(card_h))
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(25, 50, 90)
        card.line.color.rgb = col
        card.line.width = Pt(2.0)

        # Numéro
        nb_box = s3.shapes.add_textbox(Inches(cx + 0.15), Inches(2.18), Inches(0.50), Inches(0.50))
        ntf3 = nb_box.text_frame
        ntf3.margin_left = ntf3.margin_top = ntf3.margin_right = ntf3.margin_bottom = 0
        np3 = ntf3.paragraphs[0]
        np3.text = str(i + 1)
        np3.font.size = Pt(26.0)
        np3.font.bold = True
        np3.font.color.rgb = col

        # Titre problème
        tb = s3.shapes.add_textbox(Inches(cx + 0.15), Inches(2.72),
                                   Inches(card_w - 0.30), Inches(0.44))
        tf3 = tb.text_frame
        tf3.word_wrap = True
        tf3.margin_left = tf3.margin_top = tf3.margin_right = tf3.margin_bottom = 0
        tp3 = tf3.paragraphs[0]
        tp3.text = title_p
        tp3.font.size = Pt(13.5)
        tp3.font.bold = True
        tp3.font.color.rgb = C_WHITE

        # Description
        db = s3.shapes.add_textbox(Inches(cx + 0.15), Inches(3.22),
                                   Inches(card_w - 0.30), Inches(1.80))
        dtf = db.text_frame
        dtf.word_wrap = True
        dtf.margin_left = dtf.margin_top = dtf.margin_right = dtf.margin_bottom = 0
        dp3 = dtf.paragraphs[0]
        dp3.text = desc_p
        dp3.font.size = Pt(10.5)
        dp3.font.color.rgb = RGBColor(190, 210, 240)

    # Réponse en bas
    ans_box = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Inches(1.20), Inches(5.90), Inches(10.933), Inches(0.80))
    ans_box.fill.solid()
    ans_box.fill.fore_color.rgb = C_ACCENT
    ans_box.line.fill.background()
    atf = ans_box.text_frame
    atf.word_wrap = True
    atf.margin_top = Inches(0.16)
    atf.margin_left = Inches(0.20)
    ap = atf.paragraphs[0]
    ap.alignment = PP_ALIGN.CENTER
    ap.text = ("Réponse : un pipeline IA local — OCR, NER, VLM et API REST — "
               "sans données en cloud, entièrement déployable au cabinet.")
    ap.font.size = Pt(10.5)
    ap.font.bold = True
    ap.font.color.rgb = C_WHITE

    add_footer(s3, 3, TOTAL_SLIDES)

    s3.notes_slide.notes_text_frame.text = (
        "La problématique centrale de ce PFE est la suivante :\n"
        "Comment automatiser le traitement et le versement de 50 ans d'archives foncières manuscrites sur Géofoncier ?\n\n"
        "Trois obstacles concrets rendent cette tâche difficile :\n"
        "Premier obstacle : les registres sont manuscrits, avec des abréviations, une encre parfois fanée et une écriture "
        "cursive qui résiste aux OCR classiques.\n"
        "Deuxième obstacle : des noms de communes abrégés — comme La Chap./A. pour La Chapelle-sous-Aubenas — "
        "sont incompréhensibles pour quelqu'un qui ne connaît pas l'Ardèche.\n"
        "Troisième obstacle : sans outil, chaque dossier nécessite 15 à 30 minutes de recherche physique "
        "puis une saisie manuelle sur Géofoncier.\n\n"
        "Ma réponse est un pipeline IA entièrement local : OCR pour la lecture, NER pour l'extraction sémantique, "
        "VLM pour l'arbitrage en cas d'ambiguïté, et une API REST pour le versement automatisé."
    )

    # =========================================================================
    # SLIDE 4 — TRANSITION CHAPITRE 1
    # =========================================================================
    add_transition_slide(prs,
                         chap_num=1,
                         chap_title="Introduction & Contexte",
                         chap_subtitle="Structure d'accueil, cadre légal et besoins du cabinet GEO-SIAPP",
                         slide_num=4, total_slides=TOTAL_SLIDES, chap_idx=0)

    # =========================================================================
    # SLIDE 5 — 1.1 Structure d'accueil GEO-SIAPP
    # =========================================================================
    s5 = prs.slides.add_slide(blank)
    apply_background(s5)
    add_navigation_bars(s5, active_chap_idx=0, active_sub_idx=0)
    add_slide_header(s5, "Chapitre 1 — Introduction",
                     "Le Cabinet GEO-SIAPP : une structure ardéchoise aux archives étendues")
    add_footer(s5, 5, TOTAL_SLIDES)

    c5_1 = [
        ("Fondé en 1992, siège à Aubenas", "Société de géomètres-experts couvrant l'Ardèche et la Drôme via 4 agences."),
        ("4 agences : Aubenas, Pierrelatte, Vallon-Pont-d'Arc, Guilherand-Granges", "Territoire rural étendu imposant une forte proximité avec les propriétaires."),
        ("Missions : foncier, urbanisme, ingénierie VRD, drone & laser 3D", "Équipe pluridisciplinaire d'une quinzaine de collaborateurs dont plusieurs géomètres-experts."),
    ]
    draw_styled_card(s5, 0.30, 1.74, 6.00, 2.60, "Profil du cabinet", c5_1,
                     banner_color=C_NAVY, tag_str="DEPUIS 1992",
                     bottom_callout="Mission confiée : 6 mois de stage — février à juillet 2026 (siège Aubenas)")

    c5_2 = [
        ("Rachats successifs d'études locales", "Chaque acquisition intègre les collaborateurs et l'intégralité des fonds d'archives physiques."),
        ("Fonds historique : 1959 à 2007", "Près de 23 600 dossiers à Aubenas, non numérisés, conservés en sous-sol."),
        ("Objectif du PFE", "Concevoir un outil automatisé pour extraire et verser ces dossiers sur Géofoncier."),
    ]
    draw_styled_card(s5, 0.30, 4.48, 6.00, 2.52, "Le patrimoine d'archives", c5_2,
                     banner_color=C_ACCENT, tag_str="23 600 DOSSIERS",
                     bottom_callout="Enjeu : valoriser 50 ans d'histoire foncière encore dormante")

    draw_framed_image(s5, 6.50, 1.74, 6.533, 4.00,
                      "img/Chap2_pastilles.jpg",
                      "Pastilles d'intervention Géofoncier — Aubenas",
                      "Vue du registre de bornages sur le territoire du cabinet (source : GéofoncierEXPERT)")

    draw_kpi_card(s5, 6.50, 5.82, 2.10, 1.06, "4 Agences", "MAILLAGE", "Ardèche / Drôme", C_NAVY, C_NAVY_LIGHT)
    draw_kpi_card(s5, 8.72, 5.82, 2.10, 1.06, "23 600", "DOSSIERS", "Fonds non numérisés", C_ACCENT, C_ACCENT_LIGHT)
    draw_kpi_card(s5, 10.94, 5.82, 2.09, 1.06, "1959–2007", "PÉRIODE", "48 ans d'archives", C_SUCCESS, C_SUCCESS_BG)

    s5.notes_slide.notes_text_frame.text = (
        "Pour débuter, je vous présente la structure d'accueil : le cabinet GEO-SIAPP, fondé en 1992 à Aubenas.\n"
        "Avec ses 4 agences, le cabinet couvre l'Ardèche et la Drôme. Il réalise des missions de foncier, "
        "d'urbanisme, d'ingénierie VRD ainsi que des levés drone et scanner laser.\n"
        "Au fil des acquisitions de cabinets locaux, GEO-SIAPP a hérité d'un fonds d'archives considérable : "
        "près de 23 600 dossiers pour le seul site d'Aubenas, couvrant de 1959 à 2007.\n"
        "Ces dossiers, conservés physiquement en sous-sol, ne sont pas numérisés. C'est précisément l'objectif "
        "qui m'a été confié pour ce PFE de 6 mois."
    )

    # =========================================================================
    # SLIDE 6 — 1.2 + 1.3 Cadre légal & contexte opérationnel
    # =========================================================================
    s6 = prs.slides.add_slide(blank)
    apply_background(s6)
    add_navigation_bars(s6, active_chap_idx=0, active_sub_idx=1)
    add_slide_header(s6, "Chapitre 1 — Introduction",
                     "Cadre légal et contexte opérationnel : pourquoi ce projet est urgent")
    add_footer(s6, 6, TOTAL_SLIDES)

    c6_1 = [
        ("Loi n° 46-942 du 7 mai 1946", "Monopole légal du géomètre-expert pour la délimitation de la propriété privée."),
        ("Décret n° 96-478 — Art. 55", "Obligation de conservation minimale 30 ans. Actes opposables aux tiers."),
        ("Article 646 du Code civil", "Si une limite a déjà été fixée, le bornage doit se fonder sur l'acte antérieur."),
    ]
    draw_styled_card(s6, 0.30, 1.74, 5.90, 2.60, "Le cadre légal de la profession", c6_1,
                     banner_color=C_NAVY, tag_str="LOI 1946",
                     bottom_callout="La recherche d'antériorité sur Géofoncier est une obligation légale avant toute intervention.")

    c6_2 = [
        ("15 à 30 min par dossier", "Retrouver une pochette dans les archives physiques prend entre 15 et 30 minutes, sans certitude."),
        ("Géofoncier non alimenté", "Les pastilles Géofoncier ne reflètent pas les actes anciens du cabinet tant qu'ils ne sont pas versés."),
        ("Sollicitations des confrères", "D'autres cabinets demandent régulièrement copie des anciens plans."),
    ]
    draw_styled_card(s6, 6.40, 1.74, 6.633, 2.60, "Le besoin opérationnel", c6_2,
                     banner_color=C_ACCENT, tag_str="BESOIN MÉTIER",
                     bottom_callout="Bénéfice : accès instantané au dossier + alimentation du portail national")

    # KPIs
    draw_kpi_card(s6, 0.30, 4.48, 4.06, 1.06, "Art. 646", "CODE CIVIL", "Antériorité obligatoire", C_NAVY, C_NAVY_LIGHT)
    draw_kpi_card(s6, 4.48, 4.48, 4.06, 1.06, "30 min", "RECHERCHE PHYSIQUE", "Coût par dossier", C_ACCENT, C_ACCENT_LIGHT)
    draw_kpi_card(s6, 8.66, 4.48, 4.27, 1.06, "GEODEMAT", "LIMITE DGFiP", "N'inclut pas les archives privées", C_SUCCESS, C_SUCCESS_BG)

    # Encadré GEODEMAT
    geo_box = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Inches(0.30), Inches(5.70), Inches(12.733), Inches(1.26))
    geo_box.fill.solid()
    geo_box.fill.fore_color.rgb = C_ACCENT_LIGHT
    geo_box.line.color.rgb = C_ACCENT
    geo_box.line.width = Pt(1.2)
    gtf = geo_box.text_frame
    gtf.word_wrap = True
    gtf.margin_top = Inches(0.10)
    gtf.margin_left = Inches(0.20)
    gp = gtf.paragraphs[0]
    gp.text = ("Opération GEODEMAT (DGFiP, 2021) : numérise les DMPC fiscaux depuis 1956 — "
               "mais sans accès aux plans ni PV de bornage, qui restent dans les cabinets. "
               "L'angle mort des archives privées reste entier.")
    gp.font.size = Pt(9.5)
    gp.font.color.rgb = C_NAVY

    s6.notes_slide.notes_text_frame.text = (
        "Le cadre légal impose deux choses au géomètre-expert.\n"
        "Premièrement, la loi de 1946 lui confère le monopole de la délimitation de la propriété. "
        "Ses actes sont opposables aux tiers.\n"
        "Deuxièmement, l'article 646 du Code civil interdit de rebouger une borne déjà fixée sans retrouver l'acte antérieur.\n"
        "Ce qui signifie qu'avant chaque bornage, le géomètre doit consulter Géofoncier et ses archives. "
        "Sans numérisation, cette recherche prend 15 à 30 minutes dans les sous-sols du cabinet.\n"
        "L'État mène l'opération GEODEMAT avec la DGFiP pour numériser les DMPC fiscaux — mais cette opération "
        "ne donne pas accès aux plans ni aux PV de bornage, qui restent la propriété exclusive des cabinets.\n"
        "Il y a donc un angle mort complet sur les archives privées que ce PFE vise à combler."
    )

    # =========================================================================
    # SLIDE 7 — TRANSITION CHAPITRE 2
    # =========================================================================
    add_transition_slide(prs,
                         chap_num=2,
                         chap_title="Analyse Métier & Données",
                         chap_subtitle="Cadastre, API Géofoncier, fonds d'archives SIAPP et registres manuscrits",
                         slide_num=7, total_slides=TOTAL_SLIDES, chap_idx=1)

    # =========================================================================
    # SLIDE 8 — 2.1 Cadastre & propriété foncière
    # =========================================================================
    s8 = prs.slides.add_slide(blank)
    apply_background(s8)
    add_navigation_bars(s8, active_chap_idx=1, active_sub_idx=0)
    add_slide_header(s8, "Chapitre 2 — Analyse Métier & Données",
                     "Cadastre et propriété foncière : deux systèmes distincts")
    add_footer(s8, 8, TOTAL_SLIDES)

    c8_1 = [
        ("Le cadastre : outil fiscal, non opposable", "Géré par la DGFiP, il identifie les parcelles à des fins d'imposition. Il ne constitue pas un titre de propriété."),
        ("Propriété foncière : définie par les actes", "La limite réelle est fixée par le PV de bornage signé d'un géomètre-expert. C'est le seul document opposable."),
        ("Divergences fréquentes", "Entre cadastre et réalité terrain, des erreurs historiques existent — justifiant l'existence de Géofoncier."),
    ]
    draw_styled_card(s8, 0.30, 1.74, 6.20, 3.10, "Cadastre vs propriété foncière", c8_1,
                     banner_color=C_NAVY, tag_str="DISTINCTION CLÉ",
                     bottom_callout="Géofoncier réconcilie les deux en centralisant les actes de bornage.")

    c8_2 = [
        ("Portail Géofoncier (Ordre des GE)", "Créé par l'OGE, Géofoncier recense toutes les interventions foncières sous forme de pastilles géolocalisées."),
        ("Obligation de versement des actes", "Depuis 2011, tout acte signé doit être versé sur Géofoncier. Les actes antérieurs restent une lacune."),
        ("Interface GéofoncierEXPERT", "Accès réservé aux géomètres-experts pour verser et consulter les dossiers. API REST disponible."),
    ]
    draw_styled_card(s8, 6.70, 1.74, 6.333, 3.10, "Géofoncier : le registre national", c8_2,
                     banner_color=C_ACCENT, tag_str="OGE 2011",
                     bottom_callout="L'API REST de Géofoncier permet l'automatisation du versement.")

    draw_kpi_card(s8, 0.30, 4.98, 3.10, 1.00, "DGFiP", "GESTIONNAIRE CADASTRE", "Usage fiscal uniquement", C_NAVY, C_NAVY_LIGHT)
    draw_kpi_card(s8, 3.52, 4.98, 3.10, 1.00, "OGE", "GESTIONNAIRE GÉOFONCIER", "Ordre des GE", C_ACCENT, C_ACCENT_LIGHT)
    draw_kpi_card(s8, 6.74, 4.98, 3.10, 1.00, "2011", "OBLIGATION DE VERSEMENT", "Actes postérieurs", C_SUCCESS, C_SUCCESS_BG)
    draw_kpi_card(s8, 9.96, 4.98, 3.07, 1.00, "API REST", "VERSEMENT AUTOMATISÉ", "JWT + Swagger", C_WARN, C_WARN_BG)

    s8.notes_slide.notes_text_frame.text = (
        "Avant d'entrer dans la technique, il faut comprendre la distinction fondamentale entre cadastre et propriété foncière.\n"
        "Le cadastre, géré par la DGFiP, est un outil fiscal. Il identifie les parcelles pour calculer l'impôt. "
        "Il ne constitue pas un titre de propriété et n'est pas opposable en justice.\n"
        "La limite réelle, juridiquement opposable, est fixée par le procès-verbal de bornage signé par un géomètre-expert.\n"
        "C'est pour réconcilier ces deux systèmes que l'Ordre des géomètres-experts a créé Géofoncier. "
        "Depuis 2011, tout acte de bornage doit y être versé. "
        "Mais les actes antérieurs à 2011 restent une lacune que ce PFE vise à combler pour le cabinet GEO-SIAPP."
    )

    # =========================================================================
    # SLIDE 9 — 2.2 + 2.3 API Géofoncier & fonds d'archives SIAPP
    # =========================================================================
    s9 = prs.slides.add_slide(blank)
    apply_background(s9)
    add_navigation_bars(s9, active_chap_idx=1, active_sub_idx=1)
    add_slide_header(s9, "Chapitre 2 — Analyse Métier & Données",
                     "L'API Géofoncier et la caractérisation du fonds SIAPP")
    add_footer(s9, 9, TOTAL_SLIDES)

    c9_1 = [
        ("API REST documentée sous Swagger", "Le portail GéofoncierEXPERT expose des endpoints REST pour créer, lire et mettre à jour les dossiers."),
        ("Authentification JWT (token 1h)", "Chaque session nécessite un token JWT signé. Le script renouvelle le token automatiquement."),
        ("Structure du versement JSON", "Le payload attend 5 champs obligatoires : commune INSEE, date, surface, nature de l'opération et type d'acte."),
    ]
    draw_styled_card(s9, 0.30, 1.74, 5.90, 3.20, "API Géofoncier : protocole de versement", c9_1,
                     banner_color=C_NAVY, tag_str="REST / JWT",
                     bottom_callout="Endpoint principal : POST /api/v2/affaires — authentification Bearer Token.")

    c9_2 = [
        ("23 600 dossiers sur le seul site d'Aubenas", "Pochettes physiques contenant : plans calque/mylar, PV de bornage, devis, annexes cadastrales."),
        ("4 types de documents identifiés", "Plans DAO (vectoriels), DMPC scannés, registres manuscrits et pièces annexes diverses."),
        ("Registres : le cas le plus difficile", "Pages manuscrites multi-opérations, souvent une page = plusieurs affaires différentes sur une même commune."),
    ]
    draw_styled_card(s9, 6.40, 1.74, 6.633, 3.20, "Le fonds d'archives GEO-SIAPP", c9_2,
                     banner_color=C_ACCENT, tag_str="23 600 DOSSIERS",
                     bottom_callout="Priorité traitement : les registres manuscrits (plus denses, plus ambigus).")

    draw_kpi_card(s9, 0.30, 5.10, 3.06, 1.00, "5 Champs", "PAYLOAD OBLIGATOIRE", "INSEE + date + surface...", C_NAVY, C_NAVY_LIGHT)
    draw_kpi_card(s9, 3.48, 5.10, 3.06, 1.00, "1h / Token", "DURÉE JWT", "Renouvellement auto", C_ACCENT, C_ACCENT_LIGHT)
    draw_kpi_card(s9, 6.66, 5.10, 3.06, 1.00, "4 Types", "DOCUMENTS", "Plan/DMPC/Reg./Annexe", C_SUCCESS, C_SUCCESS_BG)
    draw_kpi_card(s9, 9.84, 5.10, 3.09, 1.00, "300 DPI", "NUMÉRISATION", "Résolution cible scanner", C_WARN, C_WARN_BG)

    s9.notes_slide.notes_text_frame.text = (
        "La slide précédente a présenté pourquoi Géofoncier existe. Celle-ci entre dans le comment du versement.\n"
        "L'API Géofoncier est documentée sous Swagger. Elle expose des endpoints REST standard. "
        "Le versement d'un dossier nécessite un token JWT valable une heure — mon script le renouvelle automatiquement.\n"
        "Le payload attendu comporte 5 champs obligatoires : la commune INSEE, la date de l'acte, la surface, "
        "la nature de l'opération et le type d'acte. C'est précisément ces métadonnées que mon système extrait automatiquement.\n\n"
        "Du côté du fonds SIAPP, j'ai identifié 4 types de documents : "
        "les plans DAO vectoriels, les DMPC scannés, les registres manuscrits et les pièces annexes. "
        "Les registres manuscrits sont le cas le plus complexe : une seule page peut contenir plusieurs affaires "
        "sur des communes différentes, avec des abréviations et une écriture cursive."
    )

    # =========================================================================
    # SLIDE 10 — 2.4 Registres manuscrits & filiation parcellaire
    # =========================================================================
    s10 = prs.slides.add_slide(blank)
    apply_background(s10)
    add_navigation_bars(s10, active_chap_idx=1, active_sub_idx=3)
    add_slide_header(s10, "Chapitre 2 — Analyse Métier & Données",
                     "Registres manuscrits et filiation parcellaire : les cas les plus difficiles")
    add_footer(s10, 10, TOTAL_SLIDES)

    c10_1 = [
        ("Structure tabulaire manuscrite", "Les registres sont organisés en colonnes : référence, date, commune, surface, propriétaires, parcelles."),
        ("Difficultés de lecture", "Abréviations locales, encre fanée, corrections manuscrites et écriture cursive dégradée."),
        ("Multi-affaires par page", "Une même page peut référencer 5 à 15 affaires différentes. La segmentation est critique."),
    ]
    draw_styled_card(s10, 0.30, 1.74, 5.90, 2.90, "Caractéristiques des registres", c10_1,
                     banner_color=C_NAVY, tag_str="MANUSCRIT",
                     bottom_callout="Défi OCR : cursive + tableau + abréviations = 3 problèmes combinés.")

    c10_2 = [
        ("Filiation parcellaire : le bornage sur bornage", "Quand une parcelle a déjà été bornée, tout acte ultérieur doit s'y référer. La filiation crée des chaînes d'actes."),
        ("Numérotation évolutive du cadastre", "Les numéros de parcelles changent lors des remaniements cadastraux. L'acte historique peut citer un numéro obsolète."),
        ("Contrôle de cohérence nécessaire", "Mon système vérifie les numéros INSEE, les dates et la surface via des règles déterministes (Levenshtein + regex)."),
    ]
    draw_styled_card(s10, 6.40, 1.74, 6.633, 2.90, "La filiation parcellaire", c10_2,
                     banner_color=C_ACCENT, tag_str="CHAÎNE D'ACTES",
                     bottom_callout="La filiation impose une traçabilité des numéros de parcelles dans le temps.")

    draw_framed_image(s10, 0.30, 4.78, 12.733, 2.30,
                      "img/4513_DA_124_centered_crop.png",
                      "Extrait d'un registre manuscrit — Cabinet GEO-SIAPP",
                      "Structure tabulaire en colonnes avec abréviations de communes (ex. La Chap./A. = La Chapelle-sous-Aubenas)")

    s10.notes_slide.notes_text_frame.text = (
        "Cette slide montre concrètement à quoi ressemble un registre manuscrit du cabinet.\n"
        "On voit en bas un extrait réel : une structure en colonnes avec les références d'affaires, "
        "les communes abrégées, les dates, les surfaces et les propriétaires.\n"
        "La difficulté est triple : l'écriture est cursive, l'encre peut être fanée, "
        "et les noms de communes sont abrégés. 'La Chap./A.' signifie La Chapelle-sous-Aubenas — "
        "incompréhensible pour un modèle de langue qui ne connaît pas l'Ardèche.\n\n"
        "Autre complexité : la filiation parcellaire. Quand une parcelle a déjà été bornée, "
        "l'acte ultérieur doit s'y référer. Les numéros de parcelles ont par ailleurs changé au fil des "
        "remaniements cadastraux, ce qui impose un contrôle de cohérence supplémentaire."
    )

    # =========================================================================
    # SLIDE 11 — TRANSITION CHAPITRE 3
    # =========================================================================
    add_transition_slide(prs,
                         chap_num=3,
                         chap_title="État de l'Art",
                         chap_subtitle="OCR/HTR, extraction d'entités, VLM et choix technologiques",
                         slide_num=11, total_slides=TOTAL_SLIDES, chap_idx=2)

    return prs
