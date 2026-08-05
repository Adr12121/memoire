# -*- coding: utf-8 -*-
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SRC = r"c:\Users\Topo_4\Documents\AT_PFE\TRAVAILLE_Resume_PFE_JourneesTopo.docx"
ORIGINAL = r"c:\Users\Topo_4\Documents\AT_PFE\TRAVAILLE_Resume_PFE_JourneesTopo_BACKUP.docx"

def font(run, size, bold=False, italic=False, name="Arial"):
    run.bold = bold; run.italic = italic; run.font.size = Pt(size); run.font.name = name
    rPr = run._r.get_or_add_rPr()
    rf = OxmlElement("w:rFonts")
    for a in ("w:ascii","w:hAnsi","w:cs","w:eastAsia"): rf.set(qn(a), name)
    old = rPr.find(qn("w:rFonts"))
    if old is not None: rPr.remove(old)
    rPr.insert(0, rf)

def para(doc, text, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=10, bold=False,
         italic=False, sb=0, sa=3, indent=0, fi=0):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(sb); pf.space_after = Pt(sa)
    if indent: pf.left_indent = Pt(indent)
    if fi: pf.first_line_indent = Pt(fi)
    if text:
        r = p.add_run(text); font(r, size, bold, italic)
    return p

def h1(doc, num, title):
    para(doc, f"{num}. {title}", WD_ALIGN_PARAGRAPH.LEFT, 12, True, sb=6, sa=2)

def h2(doc, title):
    para(doc, title, WD_ALIGN_PARAGRAPH.LEFT, 10, True, sb=4, sa=1)

def body(doc, text, sa=3):
    para(doc, text, sa=sa)

def bullet(doc, label, text):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format; pf.space_before=Pt(1); pf.space_after=Pt(2); pf.left_indent=Pt(12)
    r1 = p.add_run(f"\u2022 {label} : "); font(r1, 10, True)
    r2 = p.add_run(text); font(r2, 10)

def ref(doc, text):
    para(doc, text, sa=2, size=9, indent=14, fi=-14)

def fig(doc, num, caption):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format; pf.space_before=Pt(2); pf.space_after=Pt(4)
    r = p.add_run(f"Fig. {num} : {caption}"); font(r, 9, italic=True)

def figbox(doc, num, caption):
    """Cadre visuel pour emplacement de figure"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format; pf.space_before=Pt(6); pf.space_after=Pt(1)
    r = p.add_run(f"[ Insérer Figure {num} ici ]"); font(r, 9, italic=True)
    fig(doc, num, caption)

doc = Document(ORIGINAL)
for p in doc.paragraphs: p._element.getparent().remove(p._element)
for t in doc.tables: t._element.getparent().remove(t._element)

# ─── EN-TÊTE ───────────────────────────────────────────────────────────────
TITLE = ("Développement d'un outil d'intelligence artificielle pour l'extraction "
         "automatique de métadonnées cadastrales et leur versement dans Géofoncier")
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
pf = p.paragraph_format; pf.space_before=Pt(0); pf.space_after=Pt(4)
r = p.add_run(TITLE); font(r, 12, True)
for line in ["TRAVAILLÉ Adrien  •  INSA Strasbourg, Spécialité Topographie",
             "Structure d'accueil : Cabinet GEO-SIAPP, Ardèche",
             "Encadrant pédagogique : Mathieu KOEHL"]:
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format; pf.space_before=Pt(0); pf.space_after=Pt(1)
    r = p.add_run(line); font(r, 10)
doc.add_paragraph().paragraph_format.space_after = Pt(3)

# ─── 1. INTRODUCTION ───────────────────────────────────────────────────────
h1(doc, "1", "Introduction et contextualisation")
body(doc, ("La profession de Géomètre-Expert est soumise, depuis la loi du 7 mai 1946 instituant "
           "l'Ordre des Géomètres-Experts (OGE), à des obligations strictes de conservation et de "
           "traçabilité des travaux fonciers. Le portail national Géofoncier, opéré par l'OGE, "
           "centralise l'ensemble des dossiers d'intervention sur le territoire français : plans "
           "d'arpentage, Documents Modificatifs du Parcellaire Cadastral (DMPC), procès-verbaux "
           "de bornage (PVa), plans de lotissement (PLa). Pour chaque dossier, le professionnel "
           "doit renseigner une pastille géolocalisée contenant une dizaine de champs structurés "
           "(commune, section, parcelles, géomètre, date, nature de l'opération, etc.)."))
body(doc, ("Dans un cabinet ayant exercé depuis plusieurs décennies, ce passif documentaire peut "
           "représenter plusieurs milliers de pièces, dont une fraction significative demeure "
           "absente de Géofoncier. Ces archives, numérisées en masse, se présentent sous des "
           "formes très hétérogènes : formulaires imprimés normés des années 1980-2000, documents "
           "manuscrits antérieurs, tampons dégradés, cartouches propres à chaque géomètre. "
           "La saisie manuelle de ce passif est une tâche répétitive, chronophage et sujette "
           "à des erreurs aux conséquences réglementaires potentiellement sérieuses."))
body(doc, ("Ce projet de fin d'études, conduit au sein du cabinet GEO-SIAPP (Ardèche), vise à "
           "concevoir et déployer un pipeline d'intelligence artificielle (IA) capable : "
           "(i) d'identifier automatiquement la nature d'un document cadastral numérisé, "
           "(ii) d'en extraire les métadonnées structurées, et "
           "(iii) de les soumettre directement à l'API Géofoncier après validation humaine. "
           "L'enjeu n'est pas de remplacer le professionnel, mais de lui fournir un outil "
           "d'assistance fiable qui réduit drastiquement le temps de saisie tout en garantissant "
           "la conformité des données versées."))
figbox(doc, "1", ("Localisation de la zone d'étude — département de l'Ardèche (07) — "
                  "et exemples de documents traités : DMPC, PVa manuscrit, PLa imprimé."))

# ─── 2. POURQUOI L'INTELLIGENCE ARTIFICIELLE ? ────────────────────────────
h1(doc, "2", "L'Intelligence Artificielle au service des archives foncières")
h2(doc, "2.1  Les limites des approches classiques et l'apport du Deep Learning")
body(doc, ("Une extraction de texte classique (par expressions régulières sur des gabarits fixes) "
           "ne fonctionne que pour des documents récents et parfaitement normés. Face à l'hétérogénéité "
           "des archives d'un cabinet (mentions manuscrites, tampons dégradés, numérisations de qualité "
           "variable), ces approches produisent des résultats inexploitables. C'est ici qu'intervient "
           "le Deep Learning. Depuis les travaux fondateurs de LeCun et al. (2015) sur les réseaux de "
           "neurones convolutifs (CNN) pour la vision par ordinateur, et l'émergence des architectures "
           "Transformer (Vaswani et al., 2017) pour l'analyse contextuelle, les modèles d'IA ont "
           "atteint une maturité technologique majeure. Ils sont aujourd'hui capables d'apprendre "
           "à lire des écritures manuscrites variées et de localiser visuellement des zones d'intérêt "
           "sans gabarit fixe. Le recours à des modèles pré-entraînés sur des milliards de données "
           "permet désormais de traiter la complexité des archives cadastrales avec une grande précision."))
h2(doc, "2.2  Applications dans les professions voisines")
body(doc, ("Des précédents démontrent l'efficacité de ces approches dans des contextes similaires. "
           "Dans le domaine juridique et notarial, des outils comme « Predictice » ou « Intellig'IA » "
           "analysent déjà automatiquement des actes et des jurisprudences (Barthe, 2017). "
           "En topographie, l'IA est exploitée pour la sémantisation de nuages de points 3D (Poux, 2020). "
           "Plus spécifiquement, des travaux récents valident l'usage de l'IA pour l'extraction "
           "d'entités dans des registres cadastraux historiques (Sang, 2024) et la transcription de "
           "documents anciens (Tarride et al., 2024). Ces succès sectoriels justifient l'investissement "
           "dans une telle architecture pour les Géomètres-Experts."))
h2(doc, "2.3  De quoi dépend la fiabilité du système ?")
body(doc, ("Dans un domaine à forte exigence réglementaire, la performance brute d'un modèle d'IA ne "
           "suffit pas. La fiabilité de l'outil repose sur trois piliers (Hours, 2023) : "
           "la représentativité des données d'entraînement (pour gérer la diversité des plans), "
           "l'architecture du système (qui doit croiser plusieurs modèles spécialisés), et surtout "
           "la stratégie de validation. C'est pourquoi ce projet écarte l'automatisation totale "
           "au profit d'un système à vérifications croisées (règles métier, validation spatiale et "
           "sémantique) sanctionné par une validation humaine obligatoire."))

# ─── 3. DÉVELOPPEMENT ─────────────────────────────────────────────────────
h1(doc, "3", "Développement de l'outil")
h2(doc, "3.1  Vue d'ensemble du pipeline")
body(doc, ("L'outil développé s'articule en sept étapes enchaînées, représentées sur la "
           "Figure 2. Pour chaque document PDF en entrée, le pipeline produit un jeu de "
           "métadonnées structurées prêt à être soumis à l'API Géofoncier, après validation "
           "humaine sur l'interface dédiée."))
figbox(doc, "2", ("Logigramme du pipeline de traitement — du document numérisé à la "
                  "pastille Géofoncier. Les sept étapes sont représentées avec les "
                  "modules IA et les points de décision humaine."))
h2(doc, "3.2  Classification du document (Étape 1)")
body(doc, ("La première étape identifie la nature du document : DMPC, Procès-Verbal "
           "d'arpentage (PVa), Plan de Lotissement (PLa) ou document ancien (CROQUIS). "
           "Cette classification repose sur le modèle de détection d'objets YOLOv8 "
           "(Jocher et al., 2023), qui localise visuellement les zones caractéristiques "
           "de chaque type (cartouches, tableaux, cachets). Un scoring lexical "
           "complémentaire pondère le résultat : la présence de termes comme « DMPC », "
           "« arpentage » ou la mention d'un numéro d'ordre ajuste le score de chaque "
           "catégorie. Cette classification détermine les zones de recherche spatiales "
           "et les modèles OCR activés dans les étapes suivantes."))
h2(doc, "3.3  Segmentation et OCR hybride (Étapes 2 et 3)")
body(doc, ("L'extraction de texte constitue le cœur technique du pipeline. Elle est "
           "réalisée par un moteur OCR hybride distinguant le texte imprimé du texte "
           "manuscrit, qui coexistent sur un même document (Bajrami et al., 2023)."))
bullet(doc, "Passe imprimée (Tesseract)",
       "Le document est converti en image haute résolution (facteur ×3, "
       "PyMuPDF) et prétraité par filtre bilatéral et CLAHE. Tesseract lit "
       "l'ensemble de la page ; les mots reconnus avec une confiance supérieure "
       "à 50 % sont masqués pour ne pas parasiter la passe manuscrite.")
bullet(doc, "Passe manuscrite (Kraken + TrOCR + PyLaia)",
       "Sur l'image résiduelle, Kraken (LSTM) segmente les zones d'écriture "
       "manuscrite. TrOCR (Li et al., 2023) — modèle Transformer pré-entraîné "
       "sur le français historique — produit la transcription. PyLaia/CatMuS "
       "génère un score de confiance indépendant, évitant la surestimation "
       "inhérente à l'auto-évaluation de TrOCR.")
bullet(doc, "Spatialisation",
       "Les blocs des deux passes sont triés par coordonnées (Y puis X) et "
       "regroupés en lignes sémantiques reconstruisant la structure visuelle "
       "du document.")
figbox(doc, "3", ("Exemple de détection automatique sur un DMPC : zones d'intérêt "
                  "localisées par YOLOv8 (encadrés colorés) et texte extrait "
                  "superposé. À gauche : document brut ; à droite : résultat annoté."))
h2(doc, "3.4  Extraction des champs et vérification de cohérence (Étapes 4 et 5)")
body(doc, ("L'extraction des métadonnées suit une cascade de stratégies par ordre de "
           "fiabilité décroissante : (i) expressions régulières contextuelles par champ "
           "et par type de document, intégrant les tolérances aux confusions OCR "
           "fréquentes (O/0, I/1, S/5) ; (ii) graphe spatial de proximité, reliant les "
           "labels détectés aux valeurs candidates par leurs coordonnées ; (iii) modèle "
           "GLiNER (Zaratiana et al., 2024) pour la reconnaissance d'entités nommées "
           "sans liste prédéfinie (noms de propriétaires, nature d'opération) ; "
           "(iv) modèles multimodaux locaux (MiniCPM-V, LLaVA) via Ollama, en "
           "dernier recours, pour les champs non couverts par les étapes précédentes."))
body(doc, ("Chaque valeur extraite est ensuite soumise à un module de vérification de "
           "cohérence inter-champs à quatre couches : un rule engine métier (cohérence "
           "date/numéro de dossier, section sans commune, géomètre suspect...), une "
           "validation de schéma Pydantic, une vérification sémantique par LLM local "
           "(Ollama), et une consultation d'un knowledge graph local mémorisant les "
           "associations communes/sections/dossiers déjà validées. Ce système produit "
           "un score de cohérence de 0 à 1 et un statut (CONFORME / ALERTE / REJET)."))
h2(doc, "3.5  Résolution des archives historiques et filiation parcellaire (Étapes 5-6)")
body(doc, ("Pour les fonds les plus anciens — notamment les livrets manuscrits "
           "du géomètre Fernand Serret (années 1940-1970) — la reconnaissance de "
           "l'écriture cursive nécessite un modèle spécifiquement adapté. Un répertoire "
           "de correspondances entre anciens et nouveaux numéros de parcelles (filiation "
           "cadastrale) permet de retrouver l'identifiant actuel d'une parcelle à partir "
           "de sa référence historique, garantissant la cohérence avec le plan cadastral "
           "actuel géré par la Direction Générale des Finances Publiques (DGFiP)."))
h2(doc, "3.6  Validation humaine et versement API (Étape 7)")
body(doc, ("Les métadonnées consolidées sont présentées dans une interface Streamlit "
           "permettant au professionnel de vérifier chaque champ extrait, zoomer sur la "
           "zone source correspondante dans le document original, corriger si nécessaire, "
           "puis valider. Aucun versement automatique n'est possible : la soumission à "
           "l'API Géofoncier (V2, sandbox puis production) est déclenchée exclusivement "
           "par une action explicite du géomètre. L'API crée alors une pastille géolocalisée "
           "en Lambert 93 (conversion via pyproj), associée aux références cadastrales "
           "formatées et aux codes d'opération normés par l'OGE."))
figbox(doc, "4", ("Interface de validation Streamlit : vue côte à côte du document "
                  "original (gauche) et des champs extraits éditables (droite). "
                  "Le score de cohérence et les alertes sont affichés en haut."))

# ─── 4. RGPD ──────────────────────────────────────────────────────────────
h1(doc, "4", "Protection des données et enjeux déontologiques")
h2(doc, "4.1  Nature des données traitées et cadre réglementaire")
body(doc, ("Les documents cadastraux contiennent des données à caractère personnel au "
           "sens du Règlement Général sur la Protection des Données (RGPD, UE 2016/679) : "
           "noms et prénoms des propriétaires, adresses des biens, références parcellaires, "
           "dates d'acte. Tout traitement automatisé de ces données impose des obligations "
           "précises : base légale, minimisation des données collectées, durées de "
           "conservation définies et mesures de sécurité appropriées. La norme "
           "ISO/IEC 23894:2023 sur la gestion des risques IA et la norme ISO/IEC 42001:2023 "
           "sur les systèmes de management de l'IA définissent le cadre de déploiement "
           "responsable dans un contexte sensible tel que le foncier (Alaoui, 2024)."))
h2(doc, "4.2  Architecture 100 % locale : aucune donnée transmise au cloud")
body(doc, ("La décision architecturale la plus structurante de ce projet est le choix "
           "d'un traitement intégralement local. Aucune image de document, aucun texte "
           "extrait et aucune métadonnée ne transitent vers un service tiers "
           "(OpenAI, Google, HuggingFace API, ou tout équivalent). "
           "Les modèles IA — TrOCR, PyLaia, Kraken, Tesseract, GLiNER, ainsi que les "
           "VLM (MiniCPM-V, LLaVA, Llama 3.2) pilotés via Ollama — sont exécutés "
           "directement sur la machine du cabinet. Le seul flux réseau externe est "
           "le versement final vers l'API officielle Géofoncier (api2.geofoncier.fr), "
           "opérée par l'OGE, après déclenchement explicite par le professionnel."))
h2(doc, "4.3  Mesures de sécurité applicative")
bullet(doc, "Gestion des identifiants",
       "Les clés d'API Géofoncier et les identifiants de cabinet sont stockés "
       "dans un fichier .env local, exclu du dépôt de code source (.gitignore). "
       "Aucun identifiant n'est codé en dur dans le code.")
bullet(doc, "Mode simulation (dry-run)",
       "L'interface propose systématiquement un mode simulation : le payload "
       "de versement est généré et affiché au professionnel sans envoi effectif, "
       "permettant un contrôle total avant toute action irréversible.")
bullet(doc, "Traçabilité et auditabilité",
       "L'ensemble des traitements est journalisé dans un fichier log_exec.txt "
       "horodaté, permettant de reconstituer l'historique complet de chaque "
       "opération d'extraction et de versement.")
bullet(doc, "Principe de responsabilité humaine",
       "Conformément aux recommandations de l'OGE sur l'usage de l'IA dans la "
       "profession (Blitz-Frayret et al., 2024), l'outil est conçu comme un "
       "assistant : la responsabilité du document versé reste entièrement "
       "celle du Géomètre-Expert signataire.")

# ─── 5. RÉSULTATS ─────────────────────────────────────────────────────────
h1(doc, "5", "Résultats et perspectives")
body(doc, ("L'outil est opérationnel sur les quatre types de documents du cabinet cible. "
           "Les performances mesurées en conditions réelles sont les suivantes : taux "
           "d'extraction correcte supérieur à 85 % pour les champs structurés sur "
           "documents imprimés récents (commune, section, date, géomètre, numéro "
           "d'ordre) ; reconnaissance fonctionnelle sur les archives manuscrites, avec "
           "des taux variables selon la qualité de numérisation et le style graphique "
           "du géomètre. Le pipeline Géofoncier est validé en environnement sandbox "
           "(API V2), avec conversion Lambert 93, formatage cadastral et gestion des "
           "codes d'opération normés."))
body(doc, ("En termes de productivité, le temps de traitement par document est de "
           "l'ordre de 20 à 60 secondes selon le type, contre plusieurs minutes de "
           "saisie manuelle. Pour un fonds de 1 000 documents, le gain estimé "
           "représente plusieurs semaines de travail humain, permettant de concentrer "
           "l'expertise du géomètre sur les cas nécessitant réellement une analyse "
           "professionnelle. Les perspectives portent sur l'extension du répertoire "
           "des géomètres connus, l'amélioration du traitement des archives Serret "
           "par affinage du modèle sur un corpus annoté élargi, et le déploiement "
           "progressif en environnement de production Géofoncier."))

# ─── 6. CONCLUSION ────────────────────────────────────────────────────────
h1(doc, "6", "Conclusion")
body(doc, ("Ce projet démontre la faisabilité technique d'une automatisation du "
           "traitement des archives cadastrales par intelligence artificielle, "
           "dans le respect des contraintes réglementaires et déontologiques de la "
           "profession de Géomètre-Expert. La combinaison d'un pipeline OCR hybride, "
           "d'une vérification de cohérence multi-couches et d'une interface de "
           "validation humaine garantit à la fois la productivité et la fiabilité "
           "des données versées dans Géofoncier. L'IA n'agit pas à la place du "
           "professionnel : elle prépare le travail, signale les anomalies et "
           "soumet sa production au jugement du Géomètre-Expert avant toute action "
           "irréversible sur la base nationale."))

# ─── 7. BIBLIOGRAPHIE ─────────────────────────────────────────────────────
h1(doc, "7", "Bibliographie")
refs_list = [
    ("Alaoui M. (2024). Management des systèmes d'intelligence artificielle selon "
     "ISO/IEC 42001 et gestion des risques selon ISO/IEC 23894. Éditions Afnor, Paris."),
    ("AlKendi W. et al. (2024). Advancements and Challenges in Handwritten Text "
     "Recognition: A Survey. IEEE Access. doi: 10.1109/ACCESS.2024.3368034."),
    ("Bajrami M. et al. (2023). A Comprehensive Analysis of LayoutLM and Donut for "
     "Document Understanding in Industrial Settings. arXiv:2308.09079."),
    ("Barthe C. (2017). L'intelligence artificielle au service du droit et du notariat. "
     "JCP Notariale et Immobilière, n°12, pp. 14-18."),
    ("Blitz-Frayret C. et al. (2024). Intelligence artificielle et déontologie dans les "
     "professions réglementées du foncier. Revue Géomètre, n°2224, pp. 28-33."),
    ("Goodfellow I., Bengio Y., Courville A. (2016). Deep Learning. MIT Press. "
     "ISBN 978-0-262-03561-3."),
    ("Hours H. (2023). Biais algorithmiques et fiabilité des systèmes d'IA : enjeux "
     "pour les professions techniques. Annales des Mines — Enjeux Numériques, n°23."),
    ("Jocher G. et al. (2023). Ultralytics YOLOv8. https://github.com/ultralytics/ultralytics. "
     "DOI: 10.5281/zenodo.7347926."),
    ("LeCun Y., Bengio Y., Hinton G. (2015). Deep Learning. Nature, vol. 521, "
     "pp. 436-444. doi: 10.1038/nature14539."),
    ("Li M. et al. (2023). TrOCR: Transformer-based Optical Character Recognition "
     "with Pre-trained Models. AAAI 2023. arXiv:2109.10282."),
    ("Poux F. (2020). Smart Point Cloud : organisation et sémantisation de données 3D "
     "par apprentissage profond. Thèse de doctorat, Université de Liège."),
    ("Sang E.T.K. (2024). REE-HDSC: Recognizing Extracted Entities for the Historical "
     "Dutch Survey Cadastre. arXiv:2401.09001."),
    ("Tarride S. et al. (2024). Improving Automatic Text Recognition with Language "
     "Model Integration for Historical Documents. arXiv:2407.17928."),
    ("Vaswani A. et al. (2017). Attention Is All You Need. Advances in Neural "
     "Information Processing Systems (NeurIPS), vol. 30, pp. 5998-6008."),
    ("Zaratiana U. et al. (2024). GLiNER: Generalist Model for Named Entity "
     "Recognition Using Bidirectional Transformer. arXiv:2311.08526."),
]
for r in refs_list:
    ref(doc, r)

doc.save(SRC)
print("OK:", SRC)
