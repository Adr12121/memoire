# -*- coding: utf-8 -*-
"""
Générateur — Partie 4 : Slides 27 à 35
Adrien TRAVAILLÉ — Diplôme d'Ingénieur Topographe INSA Strasbourg

Chapitre 7 — Limites & Perspectives (slides 27-28)
  Slide 27 : Transition Chapitre 7
  Slide 28 : 7.1 + 7.2 Limites physiques & perspectives

Chapitre 8 — Conclusion (slides 29-30)
  Slide 29 : Transition Chapitre 8
  Slide 30 : 8.1 + 8.2 Bilan & déontologie

  Slide 31 : Remerciements & clôture

Annexes jury (slides 32-35)
  Slide 32 : Fiche GLiNER
  Slide 33 : Fiche YOLOv8
  Slide 34 : Fiche Pipeline OCR/HTR
  Slide 35 : Références bibliographiques
"""

import os
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

from generate_deck_core import (
    C_NAVY, C_NAVY_LIGHT, C_NAVY_MID, C_ACCENT, C_ACCENT_LIGHT,
    C_SUCCESS, C_SUCCESS_BG, C_WARN, C_WARN_BG,
    C_TEXT_DARK, C_TEXT_MUTED, C_TEXT_BODY, C_BORDER, C_WHITE, C_CALLOUT_BG,
    CHAPTER_COLORS,
    apply_background, apply_dark_background,
    add_navigation_bars, add_slide_header, add_footer, add_transition_slide,
    draw_styled_card, draw_framed_image, draw_kpi_card, add_centered_text,
)

TOTAL_SLIDES = 35


def add_slides_part4(prs):
    blank = prs.slide_layouts[6]

    # =========================================================================
    # SLIDE 27 — TRANSITION CHAPITRE 7
    # =========================================================================
    add_transition_slide(prs,
                         chap_num=7,
                         chap_title="Limites & Perspectives",
                         chap_subtitle="Limites physiques du système, retours terrain et pistes d'évolution",
                         slide_num=27, total_slides=TOTAL_SLIDES, chap_idx=6)

    # =========================================================================
    # SLIDE 28 — 7.1 + 7.2 Limites physiques & perspectives
    # =========================================================================
    s28 = prs.slides.add_slide(blank)
    apply_background(s28)
    add_navigation_bars(s28, active_chap_idx=6, active_sub_idx=0)
    add_slide_header(s28, "Chapitre 7 — Limites & Perspectives",
                     "Ce que le système ne sait pas faire (encore)")
    add_footer(s28, 28, TOTAL_SLIDES)

    c28_1 = [
        ("Scans dégradés (< 200 DPI)", "3 % du corpus test est irrécupérable : papier déchiré, taches d'humidité, encre quasi-inexistante."),
        ("Écriture libre hors tableau", "Les annotations marginales non structurées (notes cursives libres) résistent à l'extraction NER."),
        ("Géomètre B — fonds inconnu", "Si un fonds acquis provient d'un géomètre avec une cursive très personnelle, le modèle HTR nécessite un fine-tuning supplémentaire."),
    ]
    draw_styled_card(s28, 0.30, 1.74, 5.90, 3.00, "Limites techniques identifiées", c28_1,
                     banner_color=RGBColor(180, 50, 50), tag_str="LIMITES",
                     bottom_callout="Ces cas représentent environ 5 % du fonds — traitement manuel résiduel maintenu.")

    c28_2 = [
        ("Fine-tuning HTR sur corpus ardéchois", "Constituer un corpus annoté de 500+ pages de registres pour fine-tuner TrOCR sur la cursive locale."),
        ("Extension à d'autres cabinets", "Le pipeline est générique : déploiement possible dans n'importe quel cabinet disposant d'un GPU."),
        ("Intégration continue sur Géofoncier", "Versement en temps réel des nouveaux actes dès signature, sans geste manuel supplémentaire."),
    ]
    draw_styled_card(s28, 6.40, 1.74, 6.633, 3.00, "Perspectives d'évolution", c28_2,
                     banner_color=C_ACCENT, tag_str="PERSPECTIVES",
                     bottom_callout="Priorité : le fine-tuning HTR sur corpus manuscrit ardéchois (hiver 2026).")

    draw_kpi_card(s28, 0.30, 4.88, 3.06, 1.00, "5 %", "ECHECS TRAITEMENT", "Scans irrécupérables", RGBColor(180, 50, 50), RGBColor(254, 242, 242))
    draw_kpi_card(s28, 3.48, 4.88, 3.06, 1.00, "500+", "PAGES À ANNOTER", "Pour fine-tuning HTR", C_ACCENT, C_ACCENT_LIGHT)
    draw_kpi_card(s28, 6.66, 4.88, 3.06, 1.00, "N Cabinets", "GÉNÉRICITÉ", "Déployable partout", C_SUCCESS, C_SUCCESS_BG)
    draw_kpi_card(s28, 9.84, 4.88, 3.09, 1.00, "En Temps Réel", "OBJECTIF FUTUR", "Versement à la signature", C_WARN, C_WARN_BG)

    s28.notes_slide.notes_text_frame.text = (
        "Toute solution technique a ses limites, et il est important d'en être transparent.\n\n"
        "Trois limites ont été identifiées lors des tests sur le terrain.\n"
        "Premièrement, les scans très dégradés — moins de 200 DPI, papier déchiré ou taches importantes — "
        "représentent environ 3 à 5 % du corpus et sont irrécupérables par le pipeline. Ils nécessitent un traitement manuel résiduel.\n"
        "Deuxièmement, les annotations marginales libres, hors structure tabulaire, résistent à l'extraction GLiNER.\n"
        "Troisièmement, si un fonds acquis provient d'un géomètre avec une cursive très personnelle, "
        "le modèle HTR peut nécessiter un fine-tuning supplémentaire.\n\n"
        "En termes de perspectives, la priorité est le fine-tuning de TrOCR sur un corpus manuscrit ardéchois annoté. "
        "À terme, le pipeline est suffisamment générique pour être déployé dans n'importe quel cabinet de géomètre-expert."
    )

    # =========================================================================
    # SLIDE 29 — TRANSITION CHAPITRE 8
    # =========================================================================
    add_transition_slide(prs,
                         chap_num=8,
                         chap_title="Conclusion",
                         chap_subtitle="Bilan du projet, apport ingénieur et déontologie professionnelle",
                         slide_num=29, total_slides=TOTAL_SLIDES, chap_idx=7)

    # =========================================================================
    # SLIDE 30 — 8.1 + 8.2 Bilan & déontologie
    # =========================================================================
    s30 = prs.slides.add_slide(blank)
    apply_background(s30)
    add_navigation_bars(s30, active_chap_idx=7, active_sub_idx=0)
    add_slide_header(s30, "Chapitre 8 — Conclusion",
                     "Bilan du projet et responsabilité de l'ingénieur-géomètre")
    add_footer(s30, 30, TOTAL_SLIDES)

    c30_1 = [
        ("Objectif atteint", "Pipeline opérationnel : scan brut → versement Géofoncier en un flux automatisé, sans donnée en cloud."),
        ("47/50 dossiers versés sur Prades", "F1 score global de 0.88 sur les 5 champs obligatoires. Validé par le tuteur M. Hague."),
        ("Gain de temps démontré", "4 heures de saisie manuelle évitées pour 50 dossiers — soit 80 % du temps gagné."),
    ]
    draw_styled_card(s30, 0.30, 1.74, 5.90, 2.80, "Ce que ce PFE a livré", c30_1,
                     banner_color=C_NAVY, tag_str="OBJECTIF ATTEINT",
                     bottom_callout="Le pipeline est livré et déployé au cabinet GEO-SIAPP depuis juillet 2026.")

    c30_2 = [
        ("Secret professionnel (Art. 18 CGP)", "Toutes les données restent sur le réseau local du cabinet. Aucune transmission à un tiers non autorisé."),
        ("Charte OGE sur l'IA (mars 2024)", "L'IA est utilisée en assistant — le géomètre valide chaque dossier avant versement. Responsabilité non déléguée."),
        ("Traçabilité conforme Art. 55", "Journal horodaté de chaque extraction et correction. Conforme à l'obligation de conservation décennale."),
    ]
    draw_styled_card(s30, 6.40, 1.74, 6.633, 2.80, "Déontologie & Éthique professionnelle", c30_2,
                     banner_color=C_ACCENT, tag_str="OGE / CHARTE IA",
                     bottom_callout="Principe directeur : l'IA augmente le géomètre, elle ne le remplace pas.")

    draw_kpi_card(s30, 0.30, 4.68, 2.07, 1.00, "F1 = 0.88", "SCORE GLOBAL", "5 champs Géofoncier", C_SUCCESS, C_SUCCESS_BG)
    draw_kpi_card(s30, 2.49, 4.68, 2.07, 1.00, "80 %", "GAIN DE TEMPS", "vs saisie manuelle", C_ACCENT, C_ACCENT_LIGHT)
    draw_kpi_card(s30, 4.68, 4.68, 2.07, 1.00, "0 %", "DONNÉES EN CLOUD", "Déploiement 100% local", C_NAVY, C_NAVY_LIGHT)
    draw_kpi_card(s30, 6.87, 4.68, 2.07, 1.00, "47/50", "DOSSIERS VERSÉS", "Commune de Prades", C_SUCCESS, C_SUCCESS_BG)
    draw_kpi_card(s30, 9.06, 4.68, 4.00, 1.00, "Charte OGE 2024", "ÉTHIQUE IA", "Contrôle humain maintenu", C_ACCENT, C_ACCENT_LIGHT)

    s30.notes_slide.notes_text_frame.text = (
        "En conclusion, ce PFE a atteint son objectif principal : développer un pipeline opérationnel "
        "permettant de traiter automatiquement les archives foncières manuscrites et de les verser sur Géofoncier.\n\n"
        "Les résultats quantitatifs le confirment : F1 global de 0.88, 47 dossiers sur 50 versés avec succès "
        "sur la commune de Prades, et un gain de temps de 80 % par rapport à la saisie manuelle.\n\n"
        "Sur le plan de la responsabilité professionnelle, je tiens à souligner que ce système a été conçu "
        "en stricte conformité avec la déontologie du géomètre-expert.\n"
        "Toutes les données restent sur le réseau local du cabinet — conformément à l'article 18 du Code de géomètre.\n"
        "La Charte OGE sur l'IA de mars 2024 encadre l'usage de l'IA dans la profession : "
        "l'IA est un assistant, le géomètre-expert valide et signe chaque acte. La responsabilité n'est pas déléguée."
    )

    # =========================================================================
    # SLIDE 31 — REMERCIEMENTS & CLÔTURE
    # =========================================================================
    s31 = prs.slides.add_slide(blank)
    apply_dark_background(s31, C_NAVY)

    # Logo INSA
    logo_insa = "Logo_INSAStrasbourg.jpg"
    if os.path.exists(logo_insa):
        from PIL import Image
        with Image.open(logo_insa) as img:
            w, h = img.size
        sc = min(2.00 / w, 0.90 / h)
        s31.shapes.add_picture(logo_insa, Inches(0.50), Inches(0.20),
                               Inches(w * sc), Inches(h * sc))

    add_centered_text(s31, "Merci pour votre attention", y=1.50,
                      font_size=32, bold=True, color=C_WHITE)
    add_centered_text(s31, "Je suis disponible pour répondre à vos questions.", y=2.40,
                      font_size=14, bold=False, color=RGBColor(130, 190, 240))

    # Séparateur
    sep31 = s31.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                  Inches(4.5), Inches(3.10), Inches(4.333), Inches(0.025))
    sep31.fill.solid()
    sep31.fill.fore_color.rgb = C_ACCENT
    sep31.line.fill.background()

    # Remerciements
    merci_box = s31.shapes.add_textbox(Inches(1.00), Inches(3.30), Inches(11.333), Inches(1.60))
    mtf31 = merci_box.text_frame
    mtf31.word_wrap = True
    mtf31.margin_left = mtf31.margin_top = mtf31.margin_right = mtf31.margin_bottom = 0
    remerciements = [
        ("Mathieu KOEHL", "Directeur du PFE — INSA Strasbourg / Laboratoire ICube"),
        ("Gaëtan HAGUE", "Tuteur entreprise — Géomètre-Expert Associé, Cabinet GEO-SIAPP"),
        ("L'équipe GEO-SIAPP", "Pour l'accueil, la confiance accordée et l'accès au fonds d'archives"),
        ("Membres du jury", "Pour le temps consacré à l'évaluation de ce travail"),
    ]
    first31 = True
    for name, role in remerciements:
        if first31:
            pm = mtf31.paragraphs[0]
            first31 = False
        else:
            pm = mtf31.add_paragraph()
            pm.space_before = Pt(5)
        pm.text = f"{name}  —  {role}"
        pm.alignment = PP_ALIGN.CENTER
        pm.font.size = Pt(11.0)
        pm.font.color.rgb = RGBColor(180, 210, 240)

    add_footer(s31, 31, TOTAL_SLIDES)

    s31.notes_slide.notes_text_frame.text = (
        "Je vous remercie de votre attention.\n\n"
        "Je souhaite remercier particulièrement :\n"
        "- M. Mathieu Koehl pour sa direction scientifique tout au long de ce PFE\n"
        "- M. Gaëtan Hague pour son accueil au sein de GEO-SIAPP et son soutien opérationnel\n"
        "- Toute l'équipe du cabinet pour la confiance accordée et l'accès au fonds d'archives\n\n"
        "Je reste disponible pour répondre à vos questions."
    )

    # =========================================================================
    # SLIDE 32 — FICHE JURY : GLiNER
    # =========================================================================
    s32 = prs.slides.add_slide(blank)
    apply_background(s32)
    add_navigation_bars(s32, active_chap_idx=9, active_sub_idx=0)
    add_slide_header(s32, "Annexes Techniques — Questions du jury",
                     "Fiche : GLiNER — Generalist and Lightweight Named Entity Recognizer")
    add_footer(s32, 32, TOTAL_SLIDES)

    c32_1 = [
        ("Auteurs & publication", "Zaratiana et al. — arXiv 2311.08526 (2023). Equipe : Universités de Paris-Saclay & Paris Cité."),
        ("Dépôt GitHub", "github.com/urchade/GLiNER — Licence Apache 2.0. Plus de 3 000 étoiles."),
        ("Architecture", "Encoder BERT bidirectionnel. Les entités sont représentées comme des vecteurs de type + span."),
    ]
    draw_styled_card(s32, 0.30, 1.74, 6.00, 2.90, "Origine et architecture de GLiNER", c32_1,
                     banner_color=C_NAVY, tag_str="NLP 2023",
                     bottom_callout="Hugging Face : urchade/gliner-multitask-large-v0.5")

    c32_2 = [
        ("Données d'entraînement", "Corpus PILE (825 GB de texte anglais + multilingue). Aucune annotation spécifique au domaine foncier."),
        ("Mode de fonctionnement", "Zéro-shot : les types d'entités sont définis à l'inférence, pas à l'entraînement."),
        ("Avantage principal", "Aucun besoin de données d'entraînement spécifiques au fonds SIAPP."),
    ]
    draw_styled_card(s32, 6.50, 1.74, 6.533, 2.90, "Données et fonctionnement", c32_2,
                     banner_color=RGBColor(109, 40, 217), tag_str="ZÉRO-SHOT",
                     bottom_callout="Corpus PILE : disponible sur Hugging Face — EleutherAI/pile")

    # Schéma de fonctionnement
    func_box = s32.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                    Inches(0.30), Inches(4.78),
                                    Inches(12.733), Inches(2.54))
    func_box.fill.solid()
    func_box.fill.fore_color.rgb = C_NAVY_LIGHT
    func_box.line.color.rgb = C_BORDER
    func_box.line.width = Pt(0.8)

    steps_gliner = [
        ("1. Entrée", "Texte brut de la cellule\n+ Types d'entités cibles"),
        ("2. Encodage", "BERT encode le texte\net les labels de types"),
        ("3. Scoring", "Score de similarité\nspan × type"),
        ("4. Sortie", "Entités avec score\nde confiance [0, 1]"),
    ]
    step_w = 3.00
    step_gap = 0.24
    step_start = 0.50
    for i, (step_num, step_desc) in enumerate(steps_gliner):
        sx = step_start + i * (step_w + step_gap)
        step_box = s32.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                        Inches(sx), Inches(4.98),
                                        Inches(step_w), Inches(2.14))
        step_box.fill.solid()
        step_box.fill.fore_color.rgb = C_WHITE
        step_box.line.color.rgb = RGBColor(109, 40, 217)
        step_box.line.width = Pt(1.2)

        st_tf = step_box.text_frame
        st_tf.word_wrap = True
        st_tf.margin_top = Inches(0.15)
        st_tf.margin_left = Inches(0.10)
        stp = st_tf.paragraphs[0]
        stp.text = step_num
        stp.font.size = Pt(10.0)
        stp.font.bold = True
        stp.font.color.rgb = RGBColor(109, 40, 217)
        std = st_tf.add_paragraph()
        std.text = step_desc
        std.font.size = Pt(9.0)
        std.font.color.rgb = C_TEXT_BODY
        std.space_before = Pt(6)

    s32.notes_slide.notes_text_frame.text = (
        "Si on me pose des questions sur GLiNER :\n\n"
        "GLiNER (Generalist and Lightweight Named Entity Recognizer) a été développé par Zaratiana et al. "
        "en 2023, publiés sur arXiv (2311.08526). Le code est open-source sur GitHub (urchade/GLiNER, Apache 2.0).\n\n"
        "L'architecture est basée sur un encodeur BERT bidirectionnel. "
        "Les entités sont représentées comme des vecteurs de type + span, et un score de similarité est calculé.\n\n"
        "Les données d'entraînement : le corpus PILE d'EleutherAI, qui contient 825 GB de texte multilingue. "
        "Aucune annotation spécifique au domaine foncier n'a été nécessaire.\n\n"
        "Le fonctionnement est zéro-shot : les types d'entités (commune, date, surface...) sont définis au moment "
        "de l'inférence, pas lors de l'entraînement. C'est l'avantage principal pour notre usage."
    )

    # =========================================================================
    # SLIDE 33 — FICHE JURY : YOLOv8
    # =========================================================================
    s33 = prs.slides.add_slide(blank)
    apply_background(s33)
    add_navigation_bars(s33, active_chap_idx=9, active_sub_idx=1)
    add_slide_header(s33, "Annexes Techniques — Questions du jury",
                     "Fiche : YOLOv8 — Ultralytics Object Detection")
    add_footer(s33, 33, TOTAL_SLIDES)

    c33_1 = [
        ("Ultralytics (2023) — Glenn Jocher et al.", "YOLOv8 est la huitième génération de l'architecture YOLO (You Only Look Once)."),
        ("Dépôt GitHub", "github.com/ultralytics/ultralytics — Licence AGPL-3.0. Plus de 30 000 étoiles."),
        ("Variante retenue : YOLOv8n (nano)", "8 millions de paramètres. Temps d'inférence ~20 ms/image sur GPU NVIDIA."),
    ]
    draw_styled_card(s33, 0.30, 1.74, 6.00, 2.90, "Origine et architecture YOLOv8", c33_1,
                     banner_color=C_NAVY, tag_str="2023",
                     bottom_callout="Installation : pip install ultralytics — Docs : docs.ultralytics.com")

    c33_2 = [
        ("Données d'entraînement initiales", "Pré-entraîné sur COCO (118 000 images annotées, 80 classes). Fine-tuné sur le fonds SIAPP."),
        ("Fine-tuning sur corpus SIAPP", "400 images annotées manuellement avec LabelImg. Split 80/20 (train/test). 50 epochs."),
        ("3 classes annotées", "Classe 0 : cartouche. Classe 1 : tableau parcelles. Classe 2 : annotations marginales/cachets."),
    ]
    draw_styled_card(s33, 6.50, 1.74, 6.533, 2.90, "Entraînement et données", c33_2,
                     banner_color=C_ACCENT, tag_str="COCO + SIAPP",
                     bottom_callout="Outil d'annotation : LabelImg (tzutalin/labelImg) — MIT License")

    draw_kpi_card(s33, 0.30, 4.78, 2.07, 1.00, "8M", "PARAMÈTRES NANO", "Modèle le plus léger", C_NAVY, C_NAVY_LIGHT)
    draw_kpi_card(s33, 2.49, 4.78, 2.07, 1.00, "20ms", "INFÉRENCE GPU", "NVIDIA RTX 3060", C_ACCENT, C_ACCENT_LIGHT)
    draw_kpi_card(s33, 4.68, 4.78, 2.07, 1.00, "0.87", "mAP@0.5", "Sur corpus test 80 img.", C_SUCCESS, C_SUCCESS_BG)
    draw_kpi_card(s33, 6.87, 4.78, 2.07, 1.00, "400", "IMAGES ANNOTÉES", "Split 80/20", C_WARN, C_WARN_BG)
    draw_kpi_card(s33, 9.06, 4.78, 2.07, 1.00, "50", "EPOCHS FINE-TUNING", "Batch size 16", C_NAVY, C_NAVY_LIGHT)
    draw_kpi_card(s33, 11.25, 4.78, 1.78, 1.00, "3 Classes", "SEGMENTATION", "Cartouche/Tableau/Marg.", C_ACCENT, C_ACCENT_LIGHT)

    draw_framed_image(s33, 0.30, 5.90, 12.733, 1.48,
                      "img/architecture_pfe_finale_fixed.png",
                      "Intégration de YOLOv8 dans le pipeline global de traitement",
                      "YOLOv8 intervient entre le prétraitement (Hough/Otsu) et l'extraction OCR/NER")

    s33.notes_slide.notes_text_frame.text = (
        "Si on me pose des questions sur YOLOv8 :\n\n"
        "YOLOv8 est développé par Ultralytics (Glenn Jocher et al.), sorti en 2023. "
        "Le dépôt est open-source sur GitHub sous licence AGPL-3.0.\n\n"
        "J'ai utilisé la variante nano (8M paramètres) pour la performance sur GPU modeste.\n\n"
        "Les données d'entraînement : YOLOv8 est pré-entraîné sur COCO (118 000 images, 80 classes). "
        "J'ai ensuite réalisé un fine-tuning sur 400 images annotées manuellement avec LabelImg, "
        "en 3 classes : cartouche, tableau de parcelles, et annotations marginales. "
        "L'entraînement a duré 50 epochs avec un batch size de 16.\n\n"
        "Le mAP@0.5 atteint 0.87 sur le corpus de test (80 images)."
    )

    # =========================================================================
    # SLIDE 34 — FICHE JURY : Pipeline OCR/HTR
    # =========================================================================
    s34 = prs.slides.add_slide(blank)
    apply_background(s34)
    add_navigation_bars(s34, active_chap_idx=9, active_sub_idx=2)
    add_slide_header(s34, "Annexes Techniques — Questions du jury",
                     "Fiche : Pipeline OCR/HTR — EasyOCR + TrOCR")
    add_footer(s34, 34, TOTAL_SLIDES)

    c34_1 = [
        ("EasyOCR (JaidedAI, 2020)", "OCR multilingue open-source. 80+ langues. Pré-entraîné sur millions de documents imprimés. Rapide sur CPU."),
        ("TrOCR (Microsoft, 2021)", "Transformer pur : ViT encoder + GPT-2 decoder. Pré-entraîné sur 684M paires image-texte. Fine-tunable."),
        ("Stratégie de routage", "La zone est d'abord classifiée 'imprimée' ou 'manuscrite' par le classifieur binaire. Chaque zone est routée vers le bon modèle."),
    ]
    draw_styled_card(s34, 0.30, 1.74, 6.00, 3.00, "Les deux moteurs de lecture", c34_1,
                     banner_color=C_NAVY, tag_str="OCR + HTR",
                     bottom_callout="Référence : Baek et al. (ICCV 2019) — arXiv 1904.01906 — Taxonomy OCR/HTR")

    c34_2 = [
        ("Données TrOCR (Microsoft)", "684M paires image-texte. Base : english-printed et english-handwritten. Fine-tuning possible via Hugging Face Trainer."),
        ("Performance sur registres SIAPP", "Après fine-tuning sur 200 pages annotées : CER = 8.4 % (vs 31.2 % sans fine-tuning)."),
        ("Références clés", "Baek et al. (ICCV 2019) : taxonomie OCR. Li et al. (AAAI 2022) : TrOCR. HTR-United : modèles français."),
    ]
    draw_styled_card(s34, 6.50, 1.74, 6.533, 3.00, "Données et performances mesurées", c34_2,
                     banner_color=C_ACCENT, tag_str="CER = 8.4%",
                     bottom_callout="Modèle HuggingFace : microsoft/trocr-large-handwritten — MIT License")

    draw_framed_image(s34, 0.30, 4.88, 12.733, 2.44,
                      "img/baek_p4_pipeline.png",
                      "Pipeline OCR unifié — Baek et al. (ICCV 2019) — Référence académique de base",
                      "Taxonomy en 4 étapes : Trans. spatiale → Features → Modélisation séq. → Prédiction (CTC/Attn)")

    s34.notes_slide.notes_text_frame.text = (
        "Si on me pose des questions sur le pipeline OCR/HTR :\n\n"
        "J'utilise deux moteurs de lecture distincts selon la nature de la zone.\n\n"
        "EasyOCR de JaidedAI (2020) pour les zones imprimées — rapide, multilingue, sans GPU requis.\n\n"
        "TrOCR de Microsoft (2021) pour les zones manuscrites — c'est un Transformer pur, "
        "avec un ViT comme encodeur et GPT-2 comme décodeur. "
        "Il a été pré-entraîné sur 684 millions de paires image-texte par Microsoft. "
        "Disponible sur Hugging Face sous licence MIT (microsoft/trocr-large-handwritten).\n\n"
        "Après fine-tuning sur 200 pages annotées du fonds SIAPP, le CER descend à 8.4 %, "
        "contre 31.2 % sans fine-tuning.\n\n"
        "La référence académique de base pour comprendre l'architecture est l'article de Baek et al. "
        "présenté à ICCV 2019, qui propose une taxonomie unifiée des systèmes OCR en 4 étapes."
    )

    # =========================================================================
    # SLIDE 35 — RÉFÉRENCES BIBLIOGRAPHIQUES
    # =========================================================================
    s35 = prs.slides.add_slide(blank)
    apply_background(s35)
    add_navigation_bars(s35, active_chap_idx=8, active_sub_idx=0)
    add_slide_header(s35, "Références",
                     "Bibliographie et sources utilisées dans ce mémoire")
    add_footer(s35, 35, TOTAL_SLIDES)

    refs = [
        ("Académiques", C_NAVY, [
            "Baek, J. et al. (2019). What Is Wrong With Scene Text Recognition Model Comparisons? ICCV 2019. arXiv 1904.01906.",
            "Li, M. et al. (2021). TrOCR: Transformer-based Optical Character Recognition with Pre-trained Models. AAAI 2022. arXiv 2109.10282.",
            "Zaratiana, U. et al. (2023). GLiNER: Generalist and Lightweight Model for Named Entity Recognition. arXiv 2311.08526.",
            "Xu, Y. et al. (2020). LayoutLM: Pre-training of Text and Layout for Document Image Understanding. KDD 2020.",
            "Jocher, G. et al. (2023). YOLO by Ultralytics. github.com/ultralytics/ultralytics. AGPL-3.0.",
        ]),
        ("Juridiques & professionnelles", C_ACCENT, [
            "Loi n° 46-942 du 7 mai 1946 instituant l'Ordre des Géomètres-Experts.",
            "Décret n° 96-478 du 31 mai 1996 portant règlement de la profession de géomètre-expert (Art. 55).",
            "Code civil, Article 646 (bornage et antériorité).",
            "Ordre des Géomètres-Experts. (2024). Charte sur l'usage de l'intelligence artificielle. Mars 2024.",
            "API Géofoncier — Documentation Swagger — GéofoncierEXPERT — OGE (accès réservé membres).",
        ]),
        ("Logiciels open-source", C_SUCCESS, [
            "Ollama. (2023). Run LLMs locally. ollama.ai — MIT License.",
            "HTR-United. (2022). Collaborative repository for HTR ground truth. github.com/HTR-United/htr-united.",
            "Streamlit Inc. (2019). Streamlit — The fastest way to build data apps. streamlit.io — Apache 2.0.",
            "OpenStreetMap contributors. (2004). OpenStreetMap. openstreetmap.org — ODbL.",
            "Python Software Foundation. (2023). Python 3.10. python.org. PSF License.",
        ]),
    ]

    ref_w = 4.10
    ref_gap = 0.07
    ref_start = 0.30
    ref_top = 1.74
    for i, (cat_title, cat_color, ref_items) in enumerate(refs):
        x = ref_start + i * (ref_w + ref_gap)
        # Bandeau
        ban = s35.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    Inches(x), Inches(ref_top),
                                    Inches(ref_w), Inches(0.34))
        ban.fill.solid()
        ban.fill.fore_color.rgb = cat_color
        ban.line.fill.background()
        btf = ban.text_frame
        btf.margin_top = Inches(0.07)
        btf.margin_left = Inches(0.10)
        bp = btf.paragraphs[0]
        bp.text = cat_title.upper()
        bp.font.size = Pt(9.5)
        bp.font.bold = True
        bp.font.color.rgb = C_WHITE

        # Fond
        bg_ref = s35.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                       Inches(x), Inches(ref_top + 0.34),
                                       Inches(ref_w), Inches(5.60))
        bg_ref.fill.solid()
        bg_ref.fill.fore_color.rgb = C_WHITE
        bg_ref.line.color.rgb = C_BORDER
        bg_ref.line.width = Pt(0.8)

        # Texte
        rtb = s35.shapes.add_textbox(Inches(x + 0.10), Inches(ref_top + 0.40),
                                      Inches(ref_w - 0.20), Inches(5.50))
        rtf35 = rtb.text_frame
        rtf35.word_wrap = True
        rtf35.margin_left = rtf35.margin_top = rtf35.margin_right = rtf35.margin_bottom = 0
        first35 = True
        for ref_text in ref_items:
            if first35:
                rp35 = rtf35.paragraphs[0]
                first35 = False
            else:
                rp35 = rtf35.add_paragraph()
                rp35.space_before = Pt(7)
            rp35.text = f"\u2022 {ref_text}"
            rp35.font.size = Pt(7.5)
            rp35.font.color.rgb = C_TEXT_BODY

    s35.notes_slide.notes_text_frame.text = (
        "Cette slide récapitule les références bibliographiques du mémoire, organisées en trois catégories.\n\n"
        "Références académiques : les articles de Baek (OCR taxonomy), Li (TrOCR), Zaratiana (GLiNER), "
        "Xu (LayoutLM) et Jocher (YOLOv8).\n\n"
        "Références juridiques : la loi de 1946, le décret de 1996, l'article 646 du Code civil "
        "et la Charte OGE sur l'IA de mars 2024.\n\n"
        "Logiciels open-source : Ollama, HTR-United, Streamlit, OpenStreetMap et Python 3.10."
    )
