# -*- coding: utf-8 -*-
"""
Générateur — Partie 2 : Slides 12 à 14
Adrien TRAVAILLÉ — Diplôme d'Ingénieur Topographe INSA Strasbourg

Contenu (Chapitre 3 — État de l'Art) :
  Slide 12 : 3.1 OCR & HTR — pipeline et modèles
  Slide 13 : 3.2 + 3.3 GLiNER / LayoutLM + VLM & tableaux
  Slide 14 : 3.4 Synthèse et choix technologiques
  Slide 15 : Transition Chapitre 4
"""

import os
from pptx.util import Inches, Pt
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


def add_slides_part2(prs):
    blank = prs.slide_layouts[6]

    # =========================================================================
    # SLIDE 12 — 3.1 OCR & HTR : pipeline de lecture des registres
    # =========================================================================
    s12 = prs.slides.add_slide(blank)
    apply_background(s12)
    add_navigation_bars(s12, active_chap_idx=2, active_sub_idx=0)
    add_slide_header(s12, "Chapitre 3 — État de l'Art",
                     "OCR & HTR : lire les documents manuscrits anciens")
    add_footer(s12, 12, TOTAL_SLIDES)

    c12_1 = [
        ("OCR classique (Tesseract / EasyOCR)", "Adapté aux documents imprimés. Performances insuffisantes sur cursive historique (CER > 30 %)."),
        ("HTR — Handwritten Text Recognition", "Reconnaissance spécifique de l'écriture cursive. Modèles de référence : TrOCR (Microsoft), Kraken, Pylaia."),
        ("Pipeline retenu", "EasyOCR pour le texte imprimé + TrOCR fine-tuné pour les zones manuscrites identifiées par YOLOv8."),
    ]
    draw_styled_card(s12, 0.30, 1.74, 6.00, 3.00, "Du OCR classique au HTR spécialisé", c12_1,
                     banner_color=C_NAVY, tag_str="BAEK ET AL. 2019",
                     bottom_callout="Référence : What Is Wrong With Scene Text Recognition Model Comparisons? Baek et al. (2019) — ICCV")

    c12_2 = [
        ("Modèle Baek et al. (2019)", "Taxonomy unifiée de 4 étapes : transformation spatiale, extraction de features, modélisation séquentielle, prédiction."),
        ("TrOCR — Microsoft (2021)", "Transformer pur (ViT encoder + GPT-2 decoder). Fine-tunable sur corpus manuscrits historiques."),
        ("Kraken / Pylaia (HTR-United)", "Modèles open-source pré-entraînés sur archives françaises du XVIIIe–XIXe siècle via HTR-United."),
    ]
    draw_styled_card(s12, 6.50, 1.74, 6.533, 3.00, "Les modèles de référence", c12_2,
                     banner_color=C_ACCENT, tag_str="MICROSOFT 2021",
                     bottom_callout="TrOCR : pré-entraîné sur 684M documents, fine-tunable sur corpus français.")

    draw_framed_image(s12, 0.30, 4.90, 12.733, 2.18,
                      "img/baek_ocr_pipeline_crop.png",
                      "Pipeline de reconnaissance OCR/HTR — Baek et al. (ICCV 2019)",
                      "Les 4 étapes : transformation spatiale, extraction features, modélisation séq., prédiction (CTC/Attn)")

    s12.notes_slide.notes_text_frame.text = (
        "L'état de l'art commence par la question fondamentale : comment lire les documents manuscrits anciens ?\n\n"
        "Les OCR classiques comme Tesseract ou EasyOCR sont calibrés pour le texte imprimé. "
        "Sur cursive historique, le taux d'erreur caractère dépasse souvent 30 %, ce qui est inacceptable.\n\n"
        "La reconnaissance d'écriture manuscrite (HTR) est une discipline spécifique. "
        "Le modèle de référence dans le domaine est la taxonomie de Baek et al. (ICCV 2019), "
        "qui décompose la reconnaissance en 4 étapes indépendantes : "
        "transformation spatiale, extraction de caractéristiques, modélisation séquentielle et prédiction.\n\n"
        "J'ai retenu EasyOCR pour les zones imprimées et TrOCR de Microsoft pour les zones manuscrites, "
        "les deux étant orchestrés par YOLOv8 qui identifie d'abord les zones à lire."
    )

    # =========================================================================
    # SLIDE 13 — 3.2 + 3.3 GLiNER & LayoutLM + VLM
    # =========================================================================
    s13 = prs.slides.add_slide(blank)
    apply_background(s13)
    add_navigation_bars(s13, active_chap_idx=2, active_sub_idx=1)
    add_slide_header(s13, "Chapitre 3 — État de l'Art",
                     "Extraction d'entités (NER) et traitement des tableaux complexes")
    add_footer(s13, 13, TOTAL_SLIDES)

    c13_1 = [
        ("GLiNER (Zaratiana et al., 2023)", "NER zéro-shot généraliste. Entraîné sur PILE (825 GB), permet l'extraction sans fine-tuning sur un corpus spécifique."),
        ("LayoutLM (Microsoft, 2020)", "Transformer multimodal intégrant texte + coordonnées spatiales des boîtes. Performant sur formulaires structurés."),
        ("Choix GLiNER pour les registres", "Supérieur sur texte libre cursive. LayoutLM requiert une mise en forme tabulaire stricte, inadaptée aux registres variables."),
    ]
    draw_styled_card(s13, 0.30, 1.74, 6.00, 3.00, "NER : extraction des métadonnées clés", c13_1,
                     banner_color=C_NAVY, tag_str="GLINER 2023",
                     bottom_callout="GitHub : urchade/GLiNER — Apache 2.0 — Entraîné sur PILE (825 GB de texte)")

    c13_2 = [
        ("VLM — Vision Language Model", "Modèles capables de comprendre simultanément image et texte (ex. LLaVA, Qwen-VL, Phi-3-Vision)."),
        ("Ollama — inférence locale GPU", "Permet d'exécuter des VLM sur GPU local sans dépendance cloud. Modèles : llava:7b, qwen2-vl:7b."),
        ("Rôle dans le pipeline", "Arbitre de dernier recours quand OCR + GLiNER échouent à identifier la commune ou la date. N'est pas systématique."),
    ]
    draw_styled_card(s13, 6.50, 1.74, 6.533, 3.00, "VLM : arbitrage visuel local", c13_2,
                     banner_color=RGBColor(109, 40, 217), tag_str="OLLAMA LOCAL",
                     bottom_callout="Ollama : https://ollama.ai — MIT License — GPU local requis (≥ 8 Go VRAM)")

    # Tableau comparatif en bas
    table_title = s13.shapes.add_textbox(Inches(0.30), Inches(4.88), Inches(12.733), Inches(0.26))
    ttf13 = table_title.text_frame
    tp13 = ttf13.paragraphs[0]
    tp13.text = "COMPARAISON DES APPROCHES"
    tp13.font.size = Pt(8.0)
    tp13.font.bold = True
    tp13.font.color.rgb = C_ACCENT

    headers = ["Technologie", "Type", "Force principale", "Limite"]
    rows_data = [
        ["EasyOCR", "OCR", "Rapide, multilingue", "Cursive < 70 % CER"],
        ["TrOCR", "HTR", "Manuscrit historique", "Fine-tuning requis"],
        ["GLiNER", "NER", "Zéro-shot, flexible", "Contexte long < LayoutLM"],
        ["LayoutLM", "NER+Layout", "Formulaires structurés", "Inadapté cursive libre"],
        ["VLM (LLaVA)", "Multimodal", "Ambiguïté visuelle", "Lent (~8s/page GPU)"],
    ]
    col_widths = [2.40, 1.80, 4.00, 4.00]
    col_starts = [0.30, 2.82, 4.74, 8.86]
    row_h = 0.32
    row_start_y = 5.16

    # Entêtes
    for j, (header, cx, cw) in enumerate(zip(headers, col_starts, col_widths)):
        hb = s13.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   Inches(cx), Inches(row_start_y),
                                   Inches(cw), Inches(row_h))
        hb.fill.solid()
        hb.fill.fore_color.rgb = C_NAVY
        hb.line.fill.background()
        htf = hb.text_frame
        htf.margin_top = Inches(0.05)
        hp = htf.paragraphs[0]
        hp.alignment = PP_ALIGN.CENTER
        hp.text = header
        hp.font.size = Pt(8.0)
        hp.font.bold = True
        hp.font.color.rgb = C_WHITE

    # Lignes
    for r, row in enumerate(rows_data):
        y = row_start_y + (r + 1) * row_h
        bg = C_NAVY_LIGHT if r % 2 == 0 else C_WHITE
        for j, (cell, cx, cw) in enumerate(zip(row, col_starts, col_widths)):
            cb = s13.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                       Inches(cx), Inches(y),
                                       Inches(cw), Inches(row_h))
            cb.fill.solid()
            cb.fill.fore_color.rgb = bg
            cb.line.color.rgb = C_BORDER
            cb.line.width = Pt(0.6)
            ctf2 = cb.text_frame
            ctf2.margin_top = Inches(0.05)
            ctf2.margin_left = Inches(0.06)
            cp2 = ctf2.paragraphs[0]
            cp2.text = cell
            cp2.font.size = Pt(7.5)
            cp2.font.bold = (j == 0)
            cp2.font.color.rgb = C_TEXT_DARK

    s13.notes_slide.notes_text_frame.text = (
        "Pour extraire les métadonnées cibles, j'ai comparé deux approches de NER (Named Entity Recognition).\n\n"
        "GLiNER est un modèle zéro-shot généraliste, entraîné sur 825 GB de texte (corpus PILE). "
        "Il permet l'extraction de n'importe quelle entité nommée sans fine-tuning sur un corpus spécifique. "
        "Il provient du dépôt urchade/GLiNER sur GitHub, sous licence Apache 2.0.\n\n"
        "LayoutLM de Microsoft intègre texte et coordonnées spatiales. Il est performant sur des formulaires "
        "structurés comme des factures ou des DMPC. Mais il requiert une mise en page tabulaire stricte, "
        "incompatible avec la variabilité des registres manuscrits.\n\n"
        "J'ai donc retenu GLiNER pour l'extraction sémantique.\n\n"
        "Pour les cas d'ambiguïté — notamment les noms de communes abrégés — j'ai ajouté un VLM en arbitre de dernier recours. "
        "J'utilise Ollama pour exécuter LLaVA ou Qwen-VL localement sur GPU, sans aucun appel cloud."
    )

    # =========================================================================
    # SLIDE 14 — 3.4 Synthèse et justification des choix
    # =========================================================================
    s14 = prs.slides.add_slide(blank)
    apply_background(s14)
    add_navigation_bars(s14, active_chap_idx=2, active_sub_idx=3)
    add_slide_header(s14, "Chapitre 3 — État de l'Art",
                     "Synthèse des choix technologiques et justification")
    add_footer(s14, 14, TOTAL_SLIDES)

    # 4 cartes en 2x2
    choices = [
        (C_NAVY, "EasyOCR + TrOCR", "LECTURE", [
            "EasyOCR pour le texte imprimé (rapide, multilingue)",
            "TrOCR fine-tuné pour les zones manuscrites identifiées par YOLO",
            "Combinaison : couverture totale sans perdre en précision",
        ]),
        (C_ACCENT, "YOLOv8 (Ultralytics)", "SEGMENTATION", [
            "Détection des zones : cartouche, tableau parcelles, notes marginales",
            "Modèle le plus léger de la famille YOLO (8M paramètres, ~20ms/image)",
            "Fine-tuné sur 400 images annotées du fonds SIAPP",
        ]),
        (RGBColor(6, 95, 70), "GLiNER (zéro-shot)", "EXTRACTION NER", [
            "Extraction des 5 champs cibles sans données d'entraînement spécifiques",
            "Supérieur à LayoutLM sur texte libre cursive non structuré",
            "Intégré via Hugging Face — modèle urchade/gliner-multitask-large-v0.5",
        ]),
        (RGBColor(109, 40, 217), "VLM Ollama (local)", "ARBITRAGE", [
            "Résolution des ambiguïtés de communes (La Chap./A.) par analyse visuelle",
            "LLaVA 7B ou Qwen2-VL 7B selon disponibilité GPU (≥ 8 Go VRAM)",
            "Dernier recours uniquement : activé si score GLiNER < 0.65",
        ]),
    ]
    positions = [
        (0.30, 1.74, 6.20, 2.80),
        (6.60, 1.74, 6.433, 2.80),
        (0.30, 4.68, 6.20, 2.70),
        (6.60, 4.68, 6.433, 2.70),
    ]
    for (color, title14, tag14, items14), (l, t, w, h) in zip(choices, positions):
        draw_styled_card(s14, l, t, w, h, title14,
                         [(it, "") for it in items14],
                         banner_color=color, tag_str=tag14)

    s14.notes_slide.notes_text_frame.text = (
        "Cette slide synthétise les 4 choix technologiques retenus après l'état de l'art.\n\n"
        "Pour la lecture : EasyOCR sur les zones imprimées, TrOCR fine-tuné sur les zones manuscrites. "
        "Les deux sont orchestrés par YOLOv8 qui identifie les zones à lire.\n\n"
        "Pour la segmentation : YOLOv8 d'Ultralytics, dans sa variante la plus légère. "
        "J'ai annoté 400 images du fonds pour l'entraîner à repérer les cartouches, tableaux et annotations marginales.\n\n"
        "Pour l'extraction sémantique : GLiNER en mode zéro-shot. "
        "Il extrait les 5 champs obligatoires pour Géofoncier sans données d'entraînement spécifiques.\n\n"
        "Enfin, le VLM Ollama intervient comme arbitre de dernier recours pour les communes ambiguës. "
        "Il n'est activé que lorsque le score de confiance de GLiNER descend sous 0.65."
    )

    # =========================================================================
    # SLIDE 15 — TRANSITION CHAPITRE 4
    # =========================================================================
    add_transition_slide(prs,
                         chap_num=4,
                         chap_title="Architecture & Pipeline",
                         chap_subtitle="6 scripts modulaires, prétraitement, YOLOv8, GLiNER et arbitrage VLM",
                         slide_num=15, total_slides=TOTAL_SLIDES, chap_idx=3)
