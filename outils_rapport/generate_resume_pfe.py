# -*- coding: utf-8 -*-
"""
generate_resume_pfe.py
Génère le résumé PFE au format .docx (consignes Journées de la Topographie INSA)
  - Titres de section  : Arial 12 gras
  - Corps de texte     : Arial 10, justifié, interligne 1,0
  - Légendes figures   : Arial 9, italique
  - Pas de numérotation de pages
  - Marges : 2,5 cm partout
"""

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUTPUT_PATH = r"C:\Users\Topo_4\Documents\AT_PFE\TRAVAILLE_Resume_PFE_JourneesTopo.docx"

# ─── Helpers ─────────────────────────────────────────────────────────────────

def set_font(run, name="Arial", size_pt=10, bold=False, italic=False):
    run.font.name = name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    r = run._r
    rPr = r.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    rFonts.set(qn("w:cs"), name)
    ex = rPr.find(qn("w:rFonts"))
    if ex is not None:
        rPr.remove(ex)
    rPr.insert(0, rFonts)

def set_spacing(para, before=0, after=4, line=1.0):
    pf = para.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pPr = para._p.get_or_add_pPr()
    lSp = OxmlElement("w:spacing")
    lSp.set(qn("w:line"), str(int(line * 240)))
    lSp.set(qn("w:lineRule"), "auto")
    ex = pPr.find(qn("w:spacing"))
    if ex is not None:
        pPr.remove(ex)
    pPr.append(lSp)

def titre(doc, text):
    """Titre de section Arial 12 gras."""
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_spacing(p, before=10, after=4, line=1.0)
    r = p.add_run(text)
    set_font(r, size_pt=12, bold=True)

def body(doc, parts, after=4):
    """
    Corps de texte Arial 10 justifié.
    parts = [(texte, bold), ...]  — permet le gras inline.
    """
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_spacing(p, before=0, after=after, line=1.0)
    for text, bold in parts:
        r = p.add_run(text)
        set_font(r, size_pt=10, bold=bold)

def figure(doc, n, caption):
    """Emplacement figure + légende Arial 9 italique."""
    p1 = doc.add_paragraph()
    p1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_spacing(p1, before=6, after=2, line=1.0)
    r1 = p1.add_run(f"[ Insérer ici : Figure {n} ]")
    set_font(r1, size_pt=10, italic=True)
    p2 = doc.add_paragraph()
    p2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_spacing(p2, before=0, after=8, line=1.0)
    r2 = p2.add_run(f"Fig. {n} : {caption}")
    set_font(r2, size_pt=9, italic=True)


def build_document():
    doc = Document()

    # Marges 2,5 cm
    for section in doc.sections:
        section.top_margin    = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)
        for para in section.footer.paragraphs:
            for run in para.runs:
                run.text = ""

    # ── EN-TÊTE ──────────────────────────────────────────────────────────────
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_spacing(p, before=0, after=4, line=1.0)
    r = p.add_run(
        "Développement d'un outil permettant le traitement et l'insertion "
        "des archives numériques sur Géofoncier au sein d'un cabinet de Géomètre-Expert"
    )
    set_font(r, size_pt=12, bold=True)

    for line in [
        "TRAVAILLÉ Adrien  —  INSA Strasbourg, Spécialité Topographie",
        "Structure d'accueil : Cabinet GEO-SIAPP, Ardèche",
        "Encadrant pédagogique : Mathieu KOEHL",
    ]:
        p = doc.add_paragraph()
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_spacing(p, before=0, after=2, line=1.0)
        r = p.add_run(line)
        set_font(r, size_pt=10)

    # ── SECTION 1 — Contextualisation ────────────────────────────────────────
    titre(doc, "1. Contextualisation et problématique")

    body(doc, [(
        "La profession de géomètre-expert est soumise, depuis le décret n°96-478 du 31 mai 1996, "
        "à une obligation de conservation et de mise à jour des archives topographiques fixant les "
        "limites des propriétés foncières. Ces documents — procès-verbaux de bornage, plans de "
        "division, documents d'arpentage — constituent une mémoire indispensable pour toute nouvelle "
        "mission foncière. Avant d'intervenir, les opérateurs consultent systématiquement Géofoncier, "
        "le portail national de l'Ordre des Géomètres-Experts, afin de vérifier si une opération "
        "foncière antérieure concerne les parcelles visées.", False)])

    body(doc, [(
        "Or, si les documents récents sont aujourd'hui versés sur ce portail, la très grande "
        "majorité des archives historiques des cabinets n'y figure pas. Le cabinet GEO-SIAPP, "
        "basé en Ardèche, dispose d'un fonds estimé à plusieurs milliers de documents couvrant "
        "la période 1959–2007. Ces archives sont extrêmement hétérogènes : plans modernes numérisés, "
        "registres manuscrits anciens, livrets de plusieurs géomètres-experts. Aucune de ces pièces "
        "n'est indexée, ni géoréférencée sur Géofoncier. La recherche manuelle dans ce fonds est "
        "longue, non facturable, et source d'erreurs de saisie.", False)])

    body(doc, [(
        "Ce projet vise à concevoir un outil automatisé capable de lire ces archives brutes "
        "numérisées, d'en extraire les métadonnées essentielles (commune, date, numéro de dossier, "
        "géomètre responsable...) et d'automatiser leur versement sur Géofoncier (voir Fig. 1).", False)], after=6)

    figure(doc, 1,
        "Carte de localisation de la zone d'étude — département de l'Ardèche (07) "
        "et communes couvertes par le fonds d'archives GEO-SIAPP (flèche Nord + échelle)")

    # ── SECTION 2 — Méthodologie ─────────────────────────────────────────────
    titre(doc, "2. De l'archive à la pastille : méthodologie et choix techniques")

    body(doc, [
        ("Cependant, la saisie manuelle systématique de ce passif documentaire s'avère économiquement peu viable. "
         "En estimant à 30 minutes le temps de traitement par dossier, "
         "la charge totale d'un seul de ces fonds historiques (4 535 dossiers, 1977–2007) dépasse ", False),
        ("2 200 heures de saisie", True),
        (" — soit plus d'un an de travail à temps plein, non facturable. L'outil d'automatisation développé réduit ce délai à ", False),
        ("environ 3 minutes par document", True),
        (" (scan + traitement + contrôle humain), soit un gain de productivité d'un facteur "
         "10 à 15, réinvestissable en missions facturables.", False)])

    body(doc, [(
        "Ce résultat suppose toutefois de surmonter une difficulté technique fondamentale : les archives ne sont "
        "pas homogènes. Un plan DMPC numérisé récent, un registre en tableau manuscrit "
        "des années 1950, et un livret des années 1980 n'ont rien en commun visuellement. "
        "Le traitement de documents hétérogènes par une solution unique et générique est reconnu "
        "comme un défi majeur dans la littérature [AlKendi et al., 2024 ; Sang, 2024]. "
        "C'est ce constat qui a conduit à l'architecture de traitement globale (voir Fig. 2), "
        "allant du document brut jusqu'à la création de la pastille sur Géofoncier.", False)])

    body(doc, [
        ("Étape 1 — Classification. ", True),
        ("Le système identifie automatiquement la nature du document entrant (plan vectoriel "
         "récent, scan numérisé ou registre manuscrit) pour orienter le traitement vers le module "
         "adapté. Ce tri conditionne l'ensemble du traitement et évite d'appliquer des modèles "
         "lourds à des documents simples.", False)])

    body(doc, [
        ("Étape 2 — Segmentation visuelle (YOLOv8). ", True),
        ("L'algorithme de détection d'objets YOLOv8 [Jiang et al., 2022 ; Jocher et al., 2023] "
         "localise les zones d'intérêt sur le document (cartouche, tableau de données, bloc de "
         "signatures). Ce choix est justifié par ses performances en temps réel sur CPU standard "
         "— compatibles avec le matériel du cabinet — et par sa robustesse aux variations de qualité "
         "de scan. Des architectures plus lourdes, comme LayoutLM [Xu et al., 2020], bien que très "
         "performantes sur des documents standardisés [Bajrami et al., 2023], ont été écartées en "
         "raison de leurs exigences en GPU incompatibles avec l'environnement cible et du manque "
         "de données d'entraînement annotées sur ce type de documents cadastraux historiques.", False)])

    body(doc, [
        ("Étape 3 — Extraction hybride. ", True),
        ("Pour les plans modernes, un modèle de reconnaissance d'entités nommées (NER) de type "
         "GLiNER [Zaratiana et al., 2024] identifie les informations clés (commune, géomètre, "
         "date, numéro de dossier) en comprenant le contexte sémantique de la phrase, malgré les "
         "erreurs typiques de l'OCR. GLiNER présente l'avantage d'être entraînable sur de petits "
         "corpus sans GPU dédié, ce qui correspond aux contraintes du projet. Pour les registres "
         "manuscrits, une chaîne de modèles HTR (TrOCR [Li et al., 2023] pour la transcription, "
         "PyLaia [Tarride et al., 2024] pour les scores de confiance) génère plusieurs hypothèses "
         "de lecture en parallèle. En dernier recours, un modèle Vision-Langage (LLaMA-Vision, "
         "exécuté localement via Ollama) tranche les cas les plus ambigus en raisonnant sur "
         "l'image dans sa globalité.", False)])

    body(doc, [
        ("Étape 4 — Vérification de cohérence (4 couches). ", True),
        ("Avant toute validation humaine, les métadonnées traversent un module de vérification "
         "multicritères inspiré de l'approche de Sang (2024) sur la fiabilisation post-HTR. "
         "La ", False),
        ("couche A", True),
        (" applique des règles métier instantanées : cohérence entre l'année encodée dans le "
         "numéro de dossier et la date du document, géomètre confondu avec un nom de ville, "
         "échelle hors des normes cadastrales (1/200 à 1/10 000), propriétaires anciens et "
         "nouveaux identiques — signe d'un glissement OCR. La ", False),
        ("couche B", True),
        (" valide le schéma formel des données par type de plan (Pydantic). La ", False),
        ("couche C", True),
        (" (optionnelle) soumet les champs à un LLM local (Ollama/LLaMA) pour un raisonnement "
         "sémantique approfondi. La ", False),
        ("couche D", True),
        (" croise les résultats avec une base de connaissances locale — enrichie automatiquement "
         "à chaque validation humaine — pour vérifier que la section cadastrale détectée est "
         "cohérente avec la commune identifiée.", False)])

    body(doc, [
        ("Étape 5 — Résolution du numéro de dossier historique. ", True),
        ("Pour l'un des fonds historiques, commune et parcelles extraites sont croisées avec "
         "le répertoire numérisé du cabinet (4 535 dossiers, 1977–2007). Un algorithme de "
         "correspondance floue (fuzzy matching) identifie le numéro de dossier d'époque, dont "
         "le format normalise l'année et le numéro séquentiel de prise en charge "
         "(ex. : 97050 = dossier n°50 de 1997). Ce mécanisme constitue le lien entre "
         "la référence historique d'archive et la parcelle cadastrale d'époque.", False)])

    body(doc, [
        ("Étape 6 — Filiation cadastrale. ", True),
        ("Les numéros de parcelles extraits ne correspondent souvent plus au plan cadastral "
         "actuel, en raison de divisions et fusions successives. Le moteur de filiation exploite "
         "les données DFI (Documents de Filiation Informatisés) de la DGFiP, structurées en un "
         "graphe de ", False),
        ("187 287 nœuds", True),
        (" et ", False),
        ("384 306 liens", True),
        (" pour le seul département de l'Ardèche. Un parcours en profondeur récursif remonte la "
         "chaîne des divisions pour identifier les parcelles actuellement en vigueur, permettant "
         "de positionner la pastille au bon endroit sur la carte Géofoncier.", False)])

    body(doc, [
        ("Étape 7 — Validation humaine et versement API. ", True),
        ("Les métadonnées consolidées sont soumises à l'opérateur via une interface Streamlit, "
         "n'affichant que les champs incertains avec la zone du document zoomée en regard. "
         "Cette approche dite \"human-in-the-loop\" [Monarch, 2021] garantit la qualité de la "
         "donnée versée. Après validation, la requête est envoyée à l'API Géofoncier "
         "(authentification par token, payload JSON en coordonnées Lambert 93, upload du PDF "
         "en multipart), créant la pastille et indexant le document sur le portail national.", False)], after=6)

    figure(doc, 2,
        "Architecture globale de traitement — du document brut numérisé à la pastille Géofoncier")

    # ── SECTION 3 — Résultats ────────────────────────────────────────────────
    titre(doc, "3. Résultats et performances")

    body(doc, [(
        "La chaîne de traitement est actuellement opérationnelle pour deux grandes familles "
        "d'archives historiques du cabinet. La figure 3 illustre un exemple de détection automatique sur un "
        "Document Modificatif du Parcellaire Cadastral (DMPC) : les zones d'intérêt identifiées "
        "par YOLOv8 sont encadrées et étiquetées avant d'être transmises au moteur d'extraction.", False)])

    body(doc, [
        ("Le taux de détection automatique correcte, sans intervention humaine, avoisine 50 % "
         "des champs sur l'ensemble des documents. Après contrôle de l'opérateur via l'interface "
         "de validation — qui n'affiche que les zones incertaines avec le texte déchiffré en "
         "regard —, ce taux est considérablement plus élevé. Le gain de temps est significatif : "
         "là où la saisie manuelle mobilise plusieurs dizaines de minutes par pièce, le flux "
         "automatisé ramène ce temps à ", False),
        ("environ 3 minutes par plan", True),
        (", incluant le scan, le traitement OCR et le contrôle humain.", False)], after=6)

    figure(doc, 3,
        "Exemple de détection automatique des zones d'intérêt sur un plan cadastral moderne (DMPC) "
        "— bounding boxes YOLOv8 et métadonnées extraites")

    # ── SECTION 4 — Géofoncier & perspectives ───────────────────────────────
    titre(doc, "4. Intégration Géofoncier et perspectives")

    body(doc, [(
        "L'API Géofoncier est aujourd'hui pleinement intégrée à la chaîne de traitement. "
        "Une fois les métadonnées validées par l'opérateur, le système génère automatiquement "
        "la requête d'indexation au format attendu par le portail : création de la pastille "
        "géographique (coordonnées Lambert 93), rattachement des références cadastrales et "
        "upload du document PDF, le tout sans saisie manuelle supplémentaire.", False)])

    body(doc, [(
        "La prochaine étape, prévue pour juillet 2026, est la numérisation et l'intégration "
        "de nouveaux livrets anciens, ainsi que d'autres fonds manuscrits du cabinet. "
        "Leur traitement permettra d'élargir la couverture du passif et de valider la robustesse "
        "de l'outil sur des écritures plus complexes. À terme, l'objectif est d'obtenir une "
        "couverture complète du fonds historique GEO-SIAPP sur Géofoncier, offrant à l'ensemble "
        "des opérateurs un accès immédiat et géoréférencé à l'historique foncier de l'Ardèche.", False)])

    # ── CONCLUSION ───────────────────────────────────────────────────────────
    titre(doc, "5. Conclusion")

    body(doc, [(
        "Ce projet démontre la faisabilité d'une automatisation du traitement des archives "
        "foncières historiques au sein d'un cabinet de géomètre-expert. En combinant détection "
        "d'objets (YOLOv8), extraction sémantique (GLiNER), reconnaissance de l'écriture "
        "manuscrite (TrOCR, PyLaia) et modèles Vision-Langage, l'outil développé réduit "
        "considérablement le temps de saisie tout en maintenant le contrôle humain indispensable "
        "à la qualité de la donnée foncière. L'intégration directe à l'API Géofoncier, couplée "
        "au moteur de filiation cadastrale, ouvre la voie à une alimentation massive et "
        "automatisée du portail national, au bénéfice de l'ensemble de la profession.", False)])

    # ── RÉFÉRENCES ───────────────────────────────────────────────────────────
    titre(doc, "Références")

    refs = [
        "AlKendi W. et al. (2024). Advancements and Challenges in Handwritten Text Recognition. "
        "Journal of Imaging, 10(1), 18. https://doi.org/10.3390/jimaging10010018",

        "Bajrami M. et al. (2023). A Comprehensive Analysis of LayoutLM and Donut for Document "
        "Classification.",

        "Jiang P. et al. (2022). A Review of Yolo Algorithm Developments. "
        "Procedia Computer Science, 199, 1066–1073.",

        "Jocher G. et al. (2023). Ultralytics YOLOv8. https://github.com/ultralytics/ultralytics",

        "Li M. et al. (2023). TrOCR: Transformer-based Optical Character Recognition with "
        "Pre-trained Models. AAAI 2023.",

        "Monarch R. (2021). Human-in-the-Loop Machine Learning. Manning Publications.",

        "Sang E.T.K. (2024). REE-HDSC: Recognizing Extracted Entities for the Historical "
        "Database Suriname Curacao. arXiv:2401.02972.",

        "Tarride S. et al. (2024). Improving Automatic Text Recognition with Language Models "
        "in the PyLaia Open-Source Library. arXiv:2404.18722.",

        "Xu Y. et al. (2020). LayoutLM: Pre-training of Text and Layout for Document Image "
        "Understanding. KDD '20, 1192–1200.",

        "Zaratiana U. et al. (2024). GLiNER: Generalist Model for Named Entity Recognition "
        "using Bidirectional Transformer. NAACL 2024, 5364–5376.",
    ]

    for ref in refs:
        p = doc.add_paragraph()
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        set_spacing(p, before=0, after=2, line=1.0)
        r = p.add_run(ref)
        set_font(r, size_pt=9)

    doc.save(OUTPUT_PATH)
    print("[OK] Document généré : " + OUTPUT_PATH)


if __name__ == "__main__":
    build_document()
