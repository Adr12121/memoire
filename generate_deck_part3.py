# -*- coding: utf-8 -*-
"""
Générateur — Partie 3 : Slides 16 à 26
Adrien TRAVAILLÉ — Diplôme d'Ingénieur Topographe INSA Strasbourg

Chapitre 4 — Architecture & Pipeline (slides 16-19)
  Slide 16 : 4.1 Architecture globale (6 scripts)
  Slide 17 : 4.2 Prétraitement image (Hough, Otsu, 300 DPI)
  Slide 18 : 4.3 YOLOv8 + GLiNER
  Slide 19 : 4.4 Arbitrage VLM — La Chapelle-sous-Aubenas

Chapitre 5 — Validation & Résultats (slides 20-23)
  Slide 20 : Transition Chapitre 5
  Slide 21 : 5.1 Interface Streamlit & vue miroir
  Slide 22 : 5.2 Répertoire & carte Folium
  Slide 23 : 5.3 Démonstration vidéo

Chapitre 6 — Intégration Géofoncier (slides 24-26)
  Slide 24 : Transition Chapitre 6
  Slide 25 : 6.1 + 6.2 API REST & versement par lot
  Slide 26 : 6.3 + 6.4 Cas de Prades & évaluation F1
"""

import os
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

from generate_deck_core import (
    C_NAVY, C_NAVY_LIGHT, C_ACCENT, C_ACCENT_LIGHT,
    C_SUCCESS, C_SUCCESS_BG, C_WARN, C_WARN_BG,
    C_TEXT_DARK, C_TEXT_MUTED, C_TEXT_BODY, C_BORDER, C_WHITE,
    apply_background, apply_dark_background,
    add_navigation_bars, add_slide_header, add_footer, add_transition_slide,
    draw_styled_card, draw_framed_image, draw_kpi_card,
)

TOTAL_SLIDES = 35


def add_slides_part3(prs):
    blank = prs.slide_layouts[6]

    # =========================================================================
    # SLIDE 16 — 4.1 Architecture logicielle (6 scripts)
    # =========================================================================
    s16 = prs.slides.add_slide(blank)
    apply_background(s16)
    add_navigation_bars(s16, active_chap_idx=3, active_sub_idx=0)
    add_slide_header(s16, "Chapitre 4 — Architecture & Pipeline",
                     "Architecture logicielle : 6 scripts Python modulaires")
    add_footer(s16, 16, TOTAL_SLIDES)

    scripts = [
        (C_NAVY, "main.py", "Orchestrateur", "Lance la chaîne, gère la mémoire, nettoie les fichiers temporaires."),
        (C_ACCENT, "plan_classifieur.py", "Classification", "Distingue plan DAO, DMPC, registre manuscrit, pièce annexe — en < 1 seconde."),
        (RGBColor(6, 95, 70), "spatial_extracteur.py", "Segmentation YOLOv8", "Détecte et découpe les zones utiles (cartouche, tableau parcelles, annotations)."),
        (RGBColor(154, 52, 18), "outil_ocr.py", "OCR + HTR + GLiNER", "Lit le texte (EasyOCR + TrOCR) et extrait les 5 champs cibles."),
        (RGBColor(109, 40, 217), "coherence_controle.py", "Contrôle qualité", "17 règles métier, Levenshtein INSEE, validation ISO 8601 pour les dates."),
        (RGBColor(15, 118, 110), "app_validation.py + geofoncier_api.py", "IHM & API", "Interface Streamlit de relecture + génération du payload JSON REST."),
    ]

    script_w = 2.07
    script_h = 2.26
    gap16 = 0.07
    start_x16 = 0.30

    for i, (color, name, role, desc) in enumerate(scripts):
        x = start_x16 + i * (script_w + gap16)
        card = s16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                    Inches(x), Inches(1.74),
                                    Inches(script_w), Inches(script_h))
        card.fill.solid()
        card.fill.fore_color.rgb = C_WHITE
        card.line.color.rgb = color
        card.line.width = Pt(2.0)

        # Bandeau couleur
        ban = s16.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   Inches(x), Inches(1.74),
                                   Inches(script_w), Inches(0.34))
        ban.fill.solid()
        ban.fill.fore_color.rgb = color
        ban.line.fill.background()

        # Numéro
        nb16 = s16.shapes.add_textbox(Inches(x + 0.08), Inches(1.76),
                                       Inches(0.30), Inches(0.30))
        ntf16 = nb16.text_frame
        ntf16.margin_left = ntf16.margin_top = ntf16.margin_right = ntf16.margin_bottom = 0
        np16 = ntf16.paragraphs[0]
        np16.text = str(i + 1) if i < 5 else "5-6"
        np16.font.size = Pt(11.0)
        np16.font.bold = True
        np16.font.color.rgb = C_WHITE

        # Nom du script
        sn16 = s16.shapes.add_textbox(Inches(x + 0.08), Inches(2.14),
                                       Inches(script_w - 0.16), Inches(0.32))
        stf16 = sn16.text_frame
        stf16.word_wrap = True
        stf16.margin_left = stf16.margin_top = stf16.margin_right = stf16.margin_bottom = 0
        sp16 = stf16.paragraphs[0]
        sp16.text = name
        sp16.font.size = Pt(8.5)
        sp16.font.bold = True
        sp16.font.color.rgb = color

        # Rôle
        rl16 = s16.shapes.add_textbox(Inches(x + 0.08), Inches(2.50),
                                       Inches(script_w - 0.16), Inches(0.26))
        rtf16 = rl16.text_frame
        rtf16.word_wrap = True
        rtf16.margin_left = rtf16.margin_top = rtf16.margin_right = rtf16.margin_bottom = 0
        rp16 = rtf16.paragraphs[0]
        rp16.text = role
        rp16.font.size = Pt(8.0)
        rp16.font.bold = True
        rp16.font.color.rgb = C_TEXT_DARK

        # Description
        db16 = s16.shapes.add_textbox(Inches(x + 0.08), Inches(2.80),
                                       Inches(script_w - 0.16), Inches(1.10))
        dtf16 = db16.text_frame
        dtf16.word_wrap = True
        dtf16.margin_left = dtf16.margin_top = dtf16.margin_right = dtf16.margin_bottom = 0
        dp16 = dtf16.paragraphs[0]
        dp16.text = desc
        dp16.font.size = Pt(8.0)
        dp16.font.color.rgb = C_TEXT_BODY

    # Flèche-flux en bas
    flow_box = s16.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    Inches(0.30), Inches(4.12),
                                    Inches(12.733), Inches(0.38))
    flow_box.fill.solid()
    flow_box.fill.fore_color.rgb = C_NAVY_LIGHT
    flow_box.line.color.rgb = C_BORDER
    flow_box.line.width = Pt(0.8)
    ftf16 = flow_box.text_frame
    ftf16.margin_top = Inches(0.08)
    fp16 = ftf16.paragraphs[0]
    fp16.alignment = PP_ALIGN.CENTER
    fp16.text = "Scan brut  →  Classification  →  Segmentation YOLOv8  →  OCR+NER  →  Contrôle qualité  →  IHM relecture  →  Versement Géofoncier"
    fp16.font.size = Pt(8.5)
    fp16.font.bold = True
    fp16.font.color.rgb = C_NAVY

    draw_framed_image(s16, 0.30, 4.62, 12.733, 2.44,
                      "img/architecture_globale.png",
                      "Architecture complète du pipeline — 6 modules Python 3.10",
                      "Environnement : Windows 10, venv dédié, GPU local (NVIDIA) — aucune dépendance cloud")

    s16.notes_slide.notes_text_frame.text = (
        "J'arrive maintenant au cœur technique : l'architecture du système que j'ai développé.\n\n"
        "Le pipeline est entièrement écrit en Python 3.10 sous Windows, sans aucune dépendance cloud. "
        "Il comporte 6 scripts modulaires orchestrés par un fichier principal.\n\n"
        "Le flux est le suivant : "
        "un scan brut entre dans le pipeline, le classifieur l'identifie en moins d'une seconde, "
        "YOLOv8 segmente les zones utiles, EasyOCR et TrOCR lisent le texte, "
        "GLiNER extrait les 5 champs cibles, le contrôleur de cohérence applique 17 règles métier, "
        "puis l'interface Streamlit permet une relecture humaine avant le versement via l'API Géofoncier."
    )

    # =========================================================================
    # SLIDE 17 — 4.2 Prétraitement image
    # =========================================================================
    s17 = prs.slides.add_slide(blank)
    apply_background(s17)
    add_navigation_bars(s17, active_chap_idx=3, active_sub_idx=1)
    add_slide_header(s17, "Chapitre 4 — Architecture & Pipeline",
                     "Prétraitement image : redressement, binarisation et 300 DPI")
    add_footer(s17, 17, TOTAL_SLIDES)

    c17_1 = [
        ("Résolution minimale : 300 DPI", "Seuil empirique retenu pour une OCR fiable sur documents manuscrits anciens. En dessous, le CER explose."),
        ("Détection de l'inclinaison (Hough)", "Transformée de Hough sur les lignes du tableau pour estimer l'angle de rotation. Correction automatique."),
        ("Binarisation adaptative (Otsu)", "Seuillage automatique qui maximise la variance inter-classe. Résistant aux variations d'encre et d'éclairage."),
    ]
    draw_styled_card(s17, 0.30, 1.74, 5.90, 3.10, "Les 3 étapes du prétraitement", c17_1,
                     banner_color=C_NAVY, tag_str="300 DPI",
                     bottom_callout="Tous les traitements se font en mémoire (PIL + OpenCV) sans réécriture disque intermédiaire.")

    c17_2 = [
        ("Résultat : gain de 18 % de CER", "Sur le corpus test de 50 registres, le prétraitement Hough + Otsu réduit le taux d'erreur de 18 points."),
        ("Découpe des zones (crop YOLOv8)", "Après redressement, YOLOv8 découpe les boîtes englobantes des zones d'intérêt pour l'OCR ciblé."),
        ("Normalisation des contrastes", "CLAHE (Contrast Limited Adaptive Histogram Equalization) appliqué aux zones à fort contraste réduit."),
    ]
    draw_styled_card(s17, 6.40, 1.74, 6.633, 3.10, "Impact mesurable sur la qualité", c17_2,
                     banner_color=C_ACCENT, tag_str="-18% CER",
                     bottom_callout="Pipeline PIL → OpenCV → Pillow : traitement in-memory, compatible GPU via CUDA.")

    draw_framed_image(s17, 0.30, 4.98, 6.10, 2.40,
                      "img/4513_DA_124_crop.png",
                      "Document avant prétraitement",
                      "Registre brut scanné — inclinaison et contraste variable")

    draw_framed_image(s17, 6.60, 4.98, 6.433, 2.40,
                      "img/4513_DA_124_centered_crop.png",
                      "Document après redressement Hough + Otsu",
                      "Registre redressé, binarisé — prêt pour OCR/HTR")

    s17.notes_slide.notes_text_frame.text = (
        "Avant de passer aux modèles d'IA, le document brut doit être prétraité.\n\n"
        "Trois opérations sont appliquées systématiquement.\n"
        "Premièrement, la résolution est vérifiée : 300 DPI est le seuil minimal empiriquement validé pour l'OCR "
        "sur documents manuscrits anciens.\n"
        "Deuxièmement, la transformée de Hough détecte les lignes du tableau pour estimer et corriger l'angle "
        "d'inclinaison du scan.\n"
        "Troisièmement, la binarisation d'Otsu seuille automatiquement l'image en maximisant la variance "
        "inter-classe — ce qui le rend robuste aux variations d'encre et d'éclairage.\n\n"
        "Le résultat : sur notre corpus test de 50 registres, ce prétraitement réduit le taux d'erreur "
        "caractère de 18 points de pourcentage."
    )

    # =========================================================================
    # SLIDE 18 — 4.3 YOLOv8 + GLiNER
    # =========================================================================
    s18 = prs.slides.add_slide(blank)
    apply_background(s18)
    add_navigation_bars(s18, active_chap_idx=3, active_sub_idx=2)
    add_slide_header(s18, "Chapitre 4 — Architecture & Pipeline",
                     "Détection des zones (YOLOv8) et extraction sémantique (GLiNER)")
    add_footer(s18, 18, TOTAL_SLIDES)

    c18_1 = [
        ("YOLOv8n — 8M paramètres", "Ultralytics (2023). Variante nano, ~20 ms/image sur GPU. Entraîné sur 400 images annotées du fonds SIAPP."),
        ("3 classes de zones détectées", "Cartouche d'identification, tableau de parcelles, annotations marginales et cachets."),
        ("mAP@0.5 = 0.87 sur corpus test", "Performances mesurées sur 80 images de test (split 80/20). Seuil de confiance : 0.45."),
    ]
    draw_styled_card(s18, 0.30, 1.74, 6.00, 3.00, "YOLOv8 : segmentation spatiale", c18_1,
                     banner_color=C_NAVY, tag_str="mAP@0.5=0.87",
                     bottom_callout="Ultralytics YOLOv8 : https://github.com/ultralytics/ultralytics — AGPL-3.0")

    c18_2 = [
        ("GLiNER en zéro-shot NER", "Modèle : urchade/gliner-multitask-large-v0.5 (Hugging Face). Extraction des 5 champs obligatoires pour Géofoncier."),
        ("5 entités extraites", "Commune (code INSEE), date de l'acte, surface en m², nature de l'opération, type d'acte."),
        ("Score de confiance par entité", "GLiNER retourne un score [0, 1] par entité. Si score < 0.65 sur commune ou date, le VLM est activé."),
    ]
    draw_styled_card(s18, 6.50, 1.74, 6.533, 3.00, "GLiNER : extraction NER zéro-shot", c18_2,
                     banner_color=C_ACCENT, tag_str="NER ZÉRO-SHOT",
                     bottom_callout="GLiNER : urchade/GLiNER — Apache 2.0 — Hugging Face Hub")

    draw_kpi_card(s18, 0.30, 4.88, 3.06, 1.00, "8M", "PARAMÈTRES YOLO", "Nano — 20ms/image GPU", C_NAVY, C_NAVY_LIGHT)
    draw_kpi_card(s18, 3.48, 4.88, 3.06, 1.00, "mAP 0.87", "PRÉCISION YOLO", "Sur corpus test 80 img.", C_ACCENT, C_ACCENT_LIGHT)
    draw_kpi_card(s18, 6.66, 4.88, 3.06, 1.00, "5 Entités", "EXTRACTION NER", "Commune, date, surface...", C_SUCCESS, C_SUCCESS_BG)
    draw_kpi_card(s18, 9.84, 4.88, 3.09, 1.00, "Seuil 0.65", "VLM ACTIVÉ SI", "Score GLiNER < 0.65", RGBColor(109, 40, 217), RGBColor(237, 233, 254))

    draw_framed_image(s18, 0.30, 6.00, 12.733, 1.38,
                      "img/Chap2_pastilles_enhanced.jpg",
                      "Exemple de zones détectées par YOLOv8 sur un registre SIAPP",
                      "Boîtes englobantes sur cartouche + tableau parcelles (source : GEO-SIAPP)")

    s18.notes_slide.notes_text_frame.text = (
        "Le pipeline comporte deux briques majeures de traitement.\n\n"
        "Première brique : YOLOv8 segmente le document pour repérer les zones utiles. "
        "J'ai utilisé la variante nano d'Ultralytics, la plus légère, qui traite une image en 20 ms sur GPU. "
        "Après annotation manuelle de 400 images du fonds SIAPP, le modèle atteint un mAP@0.5 de 0.87.\n\n"
        "Deuxième brique : GLiNER extrait les 5 champs obligatoires pour Géofoncier "
        "— commune, date, surface, nature et type d'acte — "
        "en mode zéro-shot, sans entraînement supplémentaire. "
        "Il retourne un score de confiance par entité. "
        "Si ce score descend sous 0.65 sur la commune ou la date, le VLM est activé en arbitre."
    )

    # =========================================================================
    # SLIDE 19 — 4.4 Arbitrage VLM — La Chapelle-sous-Aubenas
    # =========================================================================
    s19 = prs.slides.add_slide(blank)
    apply_background(s19)
    add_navigation_bars(s19, active_chap_idx=3, active_sub_idx=3)
    add_slide_header(s19, "Chapitre 4 — Architecture & Pipeline",
                     "Arbitrage VLM : le cas concret de La Chapelle-sous-Aubenas")
    add_footer(s19, 19, TOTAL_SLIDES)

    c19_1 = [
        ("Problème : abréviation locale incompréhensible", "Dans les registres, 'La Chap./A.' désigne La Chapelle-sous-Aubenas — "
         "une commune de 1 800 habitants en Ardèche que nul modèle généraliste ne connaît."),
        ("GLiNER renvoie score < 0.65 sur commune", "Sans contexte géographique local, aucun modèle de langue ne peut résoudre 'La Chap./A.' de façon certaine."),
        ("Activation du VLM (LLaVA 7B / Qwen2-VL)", "Le VLM reçoit l'image de la cellule + le contexte de la page. Il identifie la commune à 97 % de précision."),
    ]
    draw_styled_card(s19, 0.30, 1.74, 7.20, 3.60, "Le problème des abréviations ardéchoises", c19_1,
                     banner_color=C_NAVY, tag_str="CAS RÉEL",
                     bottom_callout="Résultat : Le VLM retourne 'La Chapelle-sous-Aubenas' et son code INSEE (07035).")

    # Cadre droit avec illustration du problème
    illus_box = s19.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                     Inches(7.70), Inches(1.74),
                                     Inches(5.333), Inches(3.60))
    illus_box.fill.solid()
    illus_box.fill.fore_color.rgb = RGBColor(237, 233, 254)
    illus_box.line.color.rgb = RGBColor(109, 40, 217)
    illus_box.line.width = Pt(1.5)

    illus_title = s19.shapes.add_textbox(Inches(7.85), Inches(1.82),
                                          Inches(5.10), Inches(0.30))
    ittf = illus_title.text_frame
    ittf.margin_left = ittf.margin_top = ittf.margin_right = ittf.margin_bottom = 0
    itp = ittf.paragraphs[0]
    itp.text = "POURQUOI LE VLM ?"
    itp.font.size = Pt(9.0)
    itp.font.bold = True
    itp.font.color.rgb = RGBColor(109, 40, 217)

    examples_box = s19.shapes.add_textbox(Inches(7.85), Inches(2.18),
                                           Inches(5.10), Inches(2.80))
    etf19 = examples_box.text_frame
    etf19.word_wrap = True
    etf19.margin_left = etf19.margin_top = etf19.margin_right = etf19.margin_bottom = 0
    examples = [
        ("Abréviations rencontrées dans les registres :", True, C_TEXT_DARK),
        ("• La Chap./A.  →  La Chapelle-sous-Aubenas (07035)", False, C_TEXT_DARK),
        ("• St-Prix  →  Saint-Prix (07)", False, C_TEXT_DARK),
        ("• Champ.  →  Champagne (07033) ou Champagne-le-Sec (86)", False, RGBColor(180, 50, 50)),
        ("• Vals  →  Vals-les-Bains (07324)", False, C_TEXT_DARK),
        ("", False, C_TEXT_DARK),
        ("Sans contexte géographique ardéchois, impossible de trancher 'Champ.' entre 2 communes.", False, C_TEXT_MUTED),
        ("", False, C_TEXT_DARK),
        ("Le VLM analyse l'image + contexte de la page pour lever l'ambiguïté.", False, RGBColor(109, 40, 217)),
    ]
    first19 = True
    for (text, bold, color) in examples:
        if first19:
            pe = etf19.paragraphs[0]
            first19 = False
        else:
            pe = etf19.add_paragraph()
        pe.text = text
        pe.font.size = Pt(8.5)
        pe.font.bold = bold
        pe.font.color.rgb = color
        pe.space_before = Pt(2)

    # KPIs bas
    draw_kpi_card(s19, 0.30, 5.50, 3.06, 1.00, "< 0.65", "SEUIL ACTIVATION VLM", "Score GLiNER commune", RGBColor(109, 40, 217), RGBColor(237, 233, 254))
    draw_kpi_card(s19, 3.48, 5.50, 3.06, 1.00, "97 %", "PRÉCISION VLM", "Sur communes ambiguës", C_SUCCESS, C_SUCCESS_BG)
    draw_kpi_card(s19, 6.66, 5.50, 3.06, 1.00, "~8s", "TEMPS ARBITRAGE VLM", "LLaVA 7B sur GPU local", C_WARN, C_WARN_BG)
    draw_kpi_card(s19, 9.84, 5.50, 3.09, 1.00, "0% Cloud", "DÉPLOIEMENT", "100% local — Ollama", C_NAVY, C_NAVY_LIGHT)

    s19.notes_slide.notes_text_frame.text = (
        "Voici la justification concrète du VLM dans le pipeline.\n\n"
        "Dans les registres du cabinet, les noms de communes sont systématiquement abrégés. "
        "'La Chap./A.' désigne La Chapelle-sous-Aubenas, une commune de 1 800 habitants en Ardèche. "
        "Aucun modèle de langue généraliste ne connaît cette abréviation.\n\n"
        "GLiNER retourne un score de confiance inférieur à 0.65 sur ces cellules, "
        "ce qui déclenche l'activation du VLM.\n\n"
        "Le VLM reçoit deux entrées : l'image de la cellule concernée "
        "et le contexte de la page entière — ce qui lui permet de repérer d'autres communes de la même page "
        "et de déduire qu'on est en Ardèche.\n\n"
        "Sur notre corpus de validation, le VLM identifie correctement la commune à 97 % des cas. "
        "Temps d'arbitrage : environ 8 secondes par page sur GPU local — Ollama, sans aucun appel cloud."
    )

    # =========================================================================
    # SLIDE 20 — TRANSITION CHAPITRE 5
    # =========================================================================
    add_transition_slide(prs,
                         chap_num=5,
                         chap_title="Validation & Résultats",
                         chap_subtitle="Interface Streamlit, vue miroir, carte Folium et démonstration vidéo",
                         slide_num=20, total_slides=TOTAL_SLIDES, chap_idx=4)

    # =========================================================================
    # SLIDE 21 — 5.1 Interface Streamlit & vue miroir
    # =========================================================================
    s21 = prs.slides.add_slide(blank)
    apply_background(s21)
    add_navigation_bars(s21, active_chap_idx=4, active_sub_idx=0)
    add_slide_header(s21, "Chapitre 5 — Validation & Résultats",
                     "Interface Streamlit : validation humaine avant versement")
    add_footer(s21, 21, TOTAL_SLIDES)

    c21_1 = [
        ("Application locale (port 8501)", "Interface web Streamlit déployée localement. Aucune donnée ne sort du réseau du cabinet."),
        ("Vue miroir : scan vs extraction", "Côte à côte : le scan du document original à gauche, les métadonnées extraites à droite. Correction manuelle possible."),
        ("Validation par champ", "Chaque champ peut être corrigé individuellement. Un bouton 'Valider' finalise le JSON avant versement."),
    ]
    draw_styled_card(s21, 0.30, 1.74, 5.90, 3.00, "Interface de validation Streamlit", c21_1,
                     banner_color=C_NAVY, tag_str="PORT 8501",
                     bottom_callout="Déploiement : streamlit run app_validation.py — aucun serveur externe requis.")

    draw_framed_image(s21, 6.40, 1.74, 6.633, 3.00,
                      "img/Etape 1_recherche_reference.jpg",
                      "Interface Streamlit — Recherche de référence",
                      "Étape 1 : saisie de la référence d'affaire et chargement du scan")

    draw_framed_image(s21, 0.30, 4.88, 6.10, 2.40,
                      "img/5.5_reprtoire_A_infos.jpg",
                      "Vue miroir — Métadonnées extraites vs scan original",
                      "Affichage côte à côte permettant la correction manuelle champ par champ")

    draw_framed_image(s21, 6.60, 4.88, 6.433, 2.40,
                      "img/Etape_2_localisation_cartographie.jpg",
                      "Étape 2 — Localisation cartographique",
                      "Positionnement de l'affaire sur la carte avant versement API")

    s21.notes_slide.notes_text_frame.text = (
        "Avant tout versement automatisé sur Géofoncier, l'opérateur valide les données extraites "
        "via une interface Streamlit déployée localement sur le poste du cabinet.\n\n"
        "L'interface propose une vue miroir : le scan original s'affiche à gauche, "
        "les métadonnées extraites à droite. L'opérateur peut corriger n'importe quel champ "
        "d'un simple clic avant de valider.\n\n"
        "Ce choix de garder un humain dans la boucle est volontaire : "
        "l'objectif n'est pas de remplacer le géomètre, mais d'éliminer le travail de copie manuelle "
        "tout en conservant le contrôle professionnel."
    )

    # =========================================================================
    # SLIDE 22 — 5.2 Répertoire & carte Folium
    # =========================================================================
    s22 = prs.slides.add_slide(blank)
    apply_background(s22)
    add_navigation_bars(s22, active_chap_idx=4, active_sub_idx=1)
    add_slide_header(s22, "Chapitre 5 — Validation & Résultats",
                     "Répertoire d'affaires et cartographie Folium")
    add_footer(s22, 22, TOTAL_SLIDES)

    c22_1 = [
        ("Répertoire alphabétique des affaires", "Liste filtrée et recherchable de tous les dossiers traités, avec accès direct au scan et aux métadonnées."),
        ("Export CSV / Excel", "Génération automatique d'un fichier de synthèse pour le cabinet, avec tous les champs validés."),
        ("Traçabilité des modifications", "Journal horodaté de chaque correction manuelle — conforme aux exigences de l'article 55 du décret 1996."),
    ]
    draw_styled_card(s22, 0.30, 1.74, 5.90, 2.90, "Répertoire numérique des affaires", c22_1,
                     banner_color=C_NAVY, tag_str="EXPORT CSV",
                     bottom_callout="Le répertoire remplace les anciens fichiers Excel manuels tenus par le cabinet.")

    c22_2 = [
        ("Carte interactive Folium (Leaflet.js)", "Chaque affaire validée apparaît comme un marqueur géolocalisé sur une carte OpenStreetMap."),
        ("Informations au clic", "Clic sur un marqueur : référence, commune, date, surface, lien vers le scan PDF."),
        ("Export HTML autonome", "La carte est exportée en HTML autonome — consultable sans connexion internet."),
    ]
    draw_styled_card(s22, 6.40, 1.74, 6.633, 2.90, "Cartographie des affaires (Folium)", c22_2,
                     banner_color=C_ACCENT, tag_str="FOLIUM / LEAFLET",
                     bottom_callout="Superposition possible avec les couches cadastrales WMS du Géoportail IGN.")

    draw_framed_image(s22, 0.30, 4.76, 6.10, 2.60,
                      "img/5.5_reprtoire_A_liste.jpg",
                      "Répertoire des affaires — liste filtrée",
                      "Recherche par commune, date ou référence — export CSV en 1 clic")

    draw_framed_image(s22, 6.60, 4.76, 6.433, 2.60,
                      "img/6.4_etape2_parcelles.jpg",
                      "Carte Folium — localisation des affaires",
                      "Marqueurs géolocalisés sur fond OpenStreetMap + couche cadastrale IGN")

    s22.notes_slide.notes_text_frame.text = (
        "En parallèle de l'interface de validation, j'ai développé deux outils de consultation.\n\n"
        "Premièrement, un répertoire numérique des affaires : liste filtrée et recherchable de tous "
        "les dossiers traités, avec export CSV pour le cabinet.\n\n"
        "Deuxièmement, une carte interactive générée avec la bibliothèque Folium, "
        "qui s'appuie sur Leaflet.js et OpenStreetMap. "
        "Chaque affaire validée apparaît comme un marqueur géolocalisé. "
        "Un clic affiche la référence, la date, la surface et un lien vers le scan. "
        "La carte est exportée en HTML autonome — consultable sans connexion."
    )

    # =========================================================================
    # SLIDE 23 — 5.3 VIDÉO DE DÉMONSTRATION
    # =========================================================================
    s23 = prs.slides.add_slide(blank)
    apply_dark_background(s23, RGBColor(5, 15, 30))

    # Titre
    vid_title = s23.shapes.add_textbox(Inches(0.50), Inches(0.20), Inches(12.333), Inches(0.50))
    vtf = vid_title.text_frame
    vtf.margin_left = vtf.margin_top = vtf.margin_right = vtf.margin_bottom = 0
    vp = vtf.paragraphs[0]
    vp.text = "DÉMONSTRATION — Pipeline complet sur un registre réel"
    vp.alignment = PP_ALIGN.CENTER
    vp.font.size = Pt(14.0)
    vp.font.bold = True
    vp.font.color.rgb = C_ACCENT

    # Cadre vidéo
    vid_path = "Demonstration_PFE_Condensee_1m25.mp4"
    vid_w = 11.80
    vid_h = 6.18
    vid_left = (13.333 - vid_w) / 2
    vid_top = 0.78

    if os.path.exists(vid_path):
        try:
            # Insérer la vidéo
            from pptx.util import Inches as I
            media = s23.shapes.add_movie(
                vid_path,
                I(vid_left), I(vid_top),
                I(vid_w), I(vid_h),
                poster_frame_image=None,
                mime_type="video/mp4"
            )
        except Exception:
            # Fallback : cadre avec message
            _add_video_placeholder(s23, vid_path, vid_left, vid_top, vid_w, vid_h)
    else:
        _add_video_placeholder(s23, vid_path, vid_left, vid_top, vid_w, vid_h)

    add_footer(s23, 23, TOTAL_SLIDES)

    s23.notes_slide.notes_text_frame.text = (
        "Je vous propose maintenant de regarder une démonstration condensée du pipeline complet.\n\n"
        "Cette vidéo de 1 minute 25 montre le traitement d'un registre réel du cabinet GEO-SIAPP :\n"
        "- Chargement du scan brut\n"
        "- Prétraitement et redressement\n"
        "- Segmentation YOLOv8 des zones\n"
        "- Extraction GLiNER des métadonnées\n"
        "- Validation dans l'interface Streamlit\n"
        "- Versement API vers Géofoncier\n\n"
        "(Cliquez sur la vidéo pour la lancer)"
    )

    # =========================================================================
    # SLIDE 24 — TRANSITION CHAPITRE 6
    # =========================================================================
    add_transition_slide(prs,
                         chap_num=6,
                         chap_title="Intégration Géofoncier",
                         chap_subtitle="API REST, versement par lot, cas de Prades et évaluation F1",
                         slide_num=24, total_slides=TOTAL_SLIDES, chap_idx=5)

    # =========================================================================
    # SLIDE 25 — 6.1 + 6.2 API REST & versement par lot
    # =========================================================================
    s25 = prs.slides.add_slide(blank)
    apply_background(s25)
    add_navigation_bars(s25, active_chap_idx=5, active_sub_idx=0)
    add_slide_header(s25, "Chapitre 6 — Intégration Géofoncier",
                     "Protocole API REST et versement automatisé par lot")
    add_footer(s25, 25, TOTAL_SLIDES)

    c25_1 = [
        ("Authentification JWT Bearer", "Token obtenu via POST /api/auth — validité 1 heure. Renouvellement automatique par le script."),
        ("Endpoint de versement", "POST /api/v2/affaires — payload JSON avec les 5 champs + pièce jointe PDF (multipart/form-data)."),
        ("Vérification de doublon (GET)", "Avant versement, GET /api/v2/affaires?commune=... détecte les doublons par commune + date."),
    ]
    draw_styled_card(s25, 0.30, 1.74, 5.90, 3.00, "Protocole API Géofoncier", c25_1,
                     banner_color=C_NAVY, tag_str="REST / JWT",
                     bottom_callout="La documentation Swagger est accessible sur GéofoncierEXPERT pour les membres OGE.")

    c25_2 = [
        ("Versement en lot (batch mode)", "Traitement de plusieurs dossiers en séquence : délai d'1 seconde entre requêtes pour respecter le rate-limit."),
        ("Gestion des erreurs HTTP", "Codes 400/409/500 : journalisation dans un fichier log + notification Streamlit. Reprise possible."),
        ("Rapport de versement automatique", "Fichier CSV horodaté résumant : référence, commune, date, statut HTTP, ID Géofoncier retourné."),
    ]
    draw_styled_card(s25, 6.40, 1.74, 6.633, 3.00, "Versement par lot et gestion d'erreurs", c25_2,
                     banner_color=C_ACCENT, tag_str="BATCH MODE",
                     bottom_callout="Débit constaté : 80 à 120 dossiers versés par heure selon la connexion.")

    draw_framed_image(s25, 0.30, 4.88, 6.10, 2.38,
                      "img/Etape 3_Versement_geofoncier.jpg",
                      "Interface Géofoncier — Confirmation de versement",
                      "Pastille créée sur Géofoncier après versement réussi (HTTP 201)")

    draw_framed_image(s25, 6.60, 4.88, 6.433, 2.38,
                      "img/6_succes_pastille.jpg",
                      "Pastille Géofoncier — vue cartographique",
                      "L'acte historique apparaît désormais sur la carte Géofoncier de la commune")

    s25.notes_slide.notes_text_frame.text = (
        "Le versement sur Géofoncier se fait via l'API REST officielle, accessible aux membres de l'Ordre.\n\n"
        "Le protocole est le suivant : "
        "d'abord une authentification JWT pour obtenir un token valable 1 heure, "
        "puis un GET pour vérifier l'absence de doublon, "
        "et enfin un POST multipart pour verser le dossier avec ses métadonnées et le PDF.\n\n"
        "En mode batch, le script traite plusieurs dossiers en séquence avec un délai d'une seconde entre "
        "chaque requête pour respecter le rate-limit de l'API. "
        "Le débit constaté est de 80 à 120 dossiers versés par heure."
    )

    # =========================================================================
    # SLIDE 26 — 6.3 + 6.4 Cas de Prades & évaluation F1
    # =========================================================================
    s26 = prs.slides.add_slide(blank)
    apply_background(s26)
    add_navigation_bars(s26, active_chap_idx=5, active_sub_idx=2)
    add_slide_header(s26, "Chapitre 6 — Intégration Géofoncier",
                     "Cas de Prades et évaluation quantitative des performances")
    add_footer(s26, 26, TOTAL_SLIDES)

    c26_1 = [
        ("Commune de Prades (Ardèche) — test grandeur réelle", "50 dossiers des années 1970-1990 versés en conditions réelles, avec validation humaine à chaque étape."),
        ("Résultat : 47 / 50 dossiers versés avec succès", "3 échecs : 2 sur qualité scan trop dégradée (< 200 DPI), 1 sur doublon existant dans Géofoncier."),
        ("Validation par le cabinet", "M. Hague a validé la cohérence des métadonnées extraites pour chacun des 47 dossiers versés."),
    ]
    draw_styled_card(s26, 0.30, 1.74, 6.00, 2.90, "Cas opérationnel : commune de Prades", c26_1,
                     banner_color=C_NAVY, tag_str="47/50 SUCCÈS",
                     bottom_callout="Gain de temps estimé : 4h de saisie manuelle évitée pour 50 dossiers.")

    # Tableau F1 scores par champ
    fields_perf = [
        ("Commune (code INSEE)", "F1 = 0.94", C_SUCCESS),
        ("Date de l'acte", "F1 = 0.91", C_SUCCESS),
        ("Surface (m²)", "F1 = 0.88", C_ACCENT),
        ("Nature opération", "F1 = 0.85", C_ACCENT),
        ("Type d'acte", "F1 = 0.82", C_WARN),
        ("Score global", "F1 = 0.88", C_SUCCESS),
    ]

    perf_title = s26.shapes.add_textbox(Inches(6.40), Inches(1.74), Inches(6.633), Inches(0.26))
    ptf26 = perf_title.text_frame
    pp26 = ptf26.paragraphs[0]
    pp26.text = "SCORES F1 PAR CHAMP OBLIGATOIRE GÉOFONCIER"
    pp26.font.size = Pt(8.0)
    pp26.font.bold = True
    pp26.font.color.rgb = C_ACCENT

    bar_y_start = 2.06
    bar_h = 0.44
    bar_gap = 0.06
    for i, (field, score, color) in enumerate(fields_perf):
        y = bar_y_start + i * (bar_h + bar_gap)
        is_global = (i == len(fields_perf) - 1)

        # Nom du champ
        fn = s26.shapes.add_textbox(Inches(6.40), Inches(y), Inches(3.10), Inches(bar_h))
        ftf26 = fn.text_frame
        ftf26.margin_top = Inches(0.10)
        fp26 = ftf26.paragraphs[0]
        fp26.text = field
        fp26.font.size = Pt(9.0 if not is_global else 9.5)
        fp26.font.bold = is_global
        fp26.font.color.rgb = C_TEXT_DARK

        # Barre de score
        score_val = float(score.split("=")[1].strip())
        bar_total_w = 2.60
        bar_filled_w = bar_total_w * score_val
        bg_bar = s26.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                      Inches(9.62), Inches(y + 0.07),
                                      Inches(bar_total_w), Inches(bar_h - 0.14))
        bg_bar.fill.solid()
        bg_bar.fill.fore_color.rgb = C_BORDER
        bg_bar.line.fill.background()

        fg_bar = s26.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                      Inches(9.62), Inches(y + 0.07),
                                      Inches(bar_filled_w), Inches(bar_h - 0.14))
        fg_bar.fill.solid()
        fg_bar.fill.fore_color.rgb = color
        fg_bar.line.fill.background()

        # Score texte
        sc_box = s26.shapes.add_textbox(Inches(12.32), Inches(y), Inches(0.93), Inches(bar_h))
        sctf = sc_box.text_frame
        sctf.margin_top = Inches(0.08)
        scp = sctf.paragraphs[0]
        scp.text = score
        scp.font.size = Pt(9.5 if not is_global else 10.0)
        scp.font.bold = True
        scp.font.color.rgb = color

    draw_framed_image(s26, 0.30, 4.78, 12.733, 2.48,
                      "img/6.4_match_confirm.jpg",
                      "Confirmation de versement Géofoncier — Commune de Prades",
                      "Capture de la pastille créée sur Géofoncier après versement des 47 dossiers (HTTP 201)")

    s26.notes_slide.notes_text_frame.text = (
        "Pour valider le système en conditions réelles, j'ai mené une expérimentation complète "
        "sur la commune de Prades, en Ardèche.\n\n"
        "Sur 50 dossiers sélectionnés dans les archives de 1970 à 1990, "
        "47 ont été versés avec succès sur Géofoncier. "
        "Les 3 échecs s'expliquent : 2 scans trop dégradés (moins de 200 DPI) "
        "et 1 doublon déjà présent dans Géofoncier.\n\n"
        "Les scores F1 par champ sont les suivants :\n"
        "- Commune INSEE : 0.94\n"
        "- Date de l'acte : 0.91\n"
        "- Surface : 0.88\n"
        "- Nature de l'opération : 0.85\n"
        "- Type d'acte : 0.82\n"
        "Score global : F1 = 0.88\n\n"
        "Ces résultats ont été validés par M. Hague, le géomètre-expert tuteur."
    )


def _add_video_placeholder(slide, vid_path, left, top, w, h):
    """Affiche un cadre de remplacement quand la vidéo ne peut pas être insérée."""
    frame = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   Inches(left), Inches(top), Inches(w), Inches(h))
    frame.fill.solid()
    frame.fill.fore_color.rgb = RGBColor(5, 20, 45)
    frame.line.color.rgb = C_ACCENT
    frame.line.width = Pt(2.0)

    msg = slide.shapes.add_textbox(Inches(left + 0.40), Inches(top + h / 2 - 0.60),
                                    Inches(w - 0.80), Inches(1.20))
    mtf = msg.text_frame
    mtf.word_wrap = True
    mtf.margin_left = mtf.margin_top = mtf.margin_right = mtf.margin_bottom = 0
    mp = mtf.paragraphs[0]
    mp.text = f"Vidéo : {vid_path}"
    mp.alignment = PP_ALIGN.CENTER
    mp.font.size = Pt(14.0)
    mp.font.color.rgb = C_ACCENT
    mp2 = mtf.add_paragraph()
    mp2.text = "Cliquez pour lancer la démonstration"
    mp2.alignment = PP_ALIGN.CENTER
    mp2.font.size = Pt(11.0)
    mp2.font.color.rgb = RGBColor(130, 170, 220)
    mp2.space_before = Pt(8)
