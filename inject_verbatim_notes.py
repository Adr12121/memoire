# -*- coding: utf-8 -*-
"""
Inject verbatim spoken-word speaker notes with speech durations into generate_final_180s_deck.py.
No meta headers ('MOTS-CLÉS', etc.), direct spoken words, exact timings.
"""

with open('generate_final_180s_deck.py', 'r', encoding='utf-8') as f:
    code = f.read()

notes_start_marker = "# ==================== NOTES DU PRÉSENTATEUR POUR L'ORAL (37 DIAPOSITIVES) ===================="
notes_end_marker = 'output_pptx = "Soutenance_PFE_Adrien_TRAVAILLE.pptx"'

start_pos = code.find(notes_start_marker)
end_pos = code.find(notes_end_marker)

if start_pos == -1 or end_pos == -1:
    print(f"Error: markers not found (start: {start_pos}, end: {end_pos})")
    exit(1)

new_notes_section = '''# ==================== NOTES DU PRÉSENTATEUR POUR L'ORAL (37 DIAPOSITIVES) ====================
# Discours direct mot à mot avec timing de parole par diapositive (Total : 20 min)
SPEAKER_NOTES = {
    1: (
        "[Temps : 30 s | Chrono : 00:30]\\n"
        "Bonjour à tous, membres du jury. Je vous présente aujourd'hui mon projet de fin d'études d'ingénieur topographe à l'INSA Strasbourg, intitulé : "
        "Développement d'un outil permettant le traitement et l'insertion des archives numériques sur Géofoncier. "
        "Ce travail a été réalisé au sein du cabinet GEO-SIAPP sous le tutorat de Monsieur Gaëtan Hague, géomètre-expert associé, et la direction de Monsieur Mathieu Koehl."
    ),
    2: (
        "[Temps : 30 s | Chrono : 01:00]\\n"
        "Notre présentation s'articule en 8 chapitres. Nous partirons du cadre réglementaire du cabinet et de la problématique foncière, avant d'aborder l'état de l'art des technologies de lecture. "
        "Nous détaillerons ensuite l'architecture développée, le moteur de fiabilisation et l'expérimentation réelle sur le terrain, pour terminer par l'analyse des limites et le bilan pour le cabinet."
    ),
    3: (
        "[Temps : 5 s | Chrono : 01:05]\\n"
        "Commençons par le premier chapitre avec la présentation de la structure d'accueil et du contexte métier."
    ),
    4: (
        "[Temps : 30 s | Chrono : 01:35]\\n"
        "Le cabinet GEO-SIAPP est une structure historique d'Ardèche et de Drôme qui compte cinq géomètres-experts associés et quatre agences. "
        "Le seul siège d'Aubenas conserve plus de 23 600 dossiers d'archives physiques. Ce fonds rassemble près de 50 ans de plans de bornage, "
        "de divisions et d'actes d'arpentage issus des géomètres prédécesseurs, aujourd'hui stockés sous forme papier dans les réserves."
    ),
    5: (
        "[Temps : 40 s | Chrono : 02:15]\\n"
        "Le géomètre-expert dispose d'un monopole légal institué par la loi de 1946 : il est le seul habilité à fixer les limites réelles de propriété via un bornage contradictoire. "
        "Le décret de 1996 nous impose de conserver ces archives au moins 30 ans. En parallèle, l'Ordre a créé le portail Géofoncier en 2010 pour référencer ces actes. "
        "Attention à la confusion fréquente avec le programme public GEODÉMAT : l'État ne numérise que le cadastre fiscal. La valorisation de nos archives privées repose donc entièrement sur l'initiative du cabinet."
    ),
    6: (
        "[Temps : 40 s | Chrono : 02:55]\\n"
        "Sur le plan technique et juridique, la règle d'or est : Bornage sur bornage ne vaut. Avant toute intervention sur le terrain, nous devons obligatoirement rechercher si un confrère ou prédécesseur n'a pas déjà borné la parcelle. "
        "Actuellement, cette recherche manuelle dans les cartons d'archives prend 25 à 30 minutes par dossier, avec un risque permanent d'oublier un plan ancien. "
        "L'objectif est de transformer cette fouille physique en une recherche cartographique instantanée."
    ),
    7: (
        "[Temps : 45 s | Chrono : 03:40]\\n"
        "La problématique centrale du mémoire est la suivante : Dans quelle mesure est-il possible de concevoir un outil logiciel capable d'extraire avec fiabilité les données d'archives foncières très variées, tout en fonctionnant localement sur un ordinateur de bureau ? "
        "Trois contraintes fortes encadrent ce travail : l'hétérogénéité d'un fonds de 50 ans avec des encres pâlies et des calques, l'échec des logiciels OCR du marché sur l'écriture manuscrite des géomètres, "
        "et surtout le secret professionnel qui interdit formellement d'envoyer les plans clients vers le cloud. Tout doit tourner localement sur les PC du cabinet."
    ),
    8: (
        "[Temps : 5 s | Chrono : 03:45]\\n"
        "Abordons le Chapitre 2 avec l'analyse métier et les spécifications des données que nous devons traiter."
    ),
    9: (
        "[Temps : 35 s | Chrono : 04:20]\\n"
        "Il faut rappeler une distinction juridique fondamentale : le cadastre est un document administratif et fiscal à valeur purement indicative. Il ne prouve pas la propriété et ne garantit pas les limites. "
        "Seul le procès-verbal de bornage contradictoire dressé par le géomètre-expert et signé par les riverains a une valeur juridique opposable. "
        "Numériser ces archives privées, c'est garantir la sécurité foncière des propriétaires contre les erreurs du plan cadastral."
    ),
    10: (
        "[Temps : 35 s | Chrono : 04:55]\\n"
        "Pour verser un dossier sur Géofoncier, l'API REST de l'Ordre exige des métadonnées strictes : une référence unique, la date de l'acte, la commune, le code postal, et obligatoirement un centroïde calculé en coordonnées Lambert-93. "
        "Le plan ou procès-verbal doit être joint au format PDF pour consultation par la profession. Ces contraintes définissent précisément le format de sortie de notre chaîne."
    ),
    11: (
        "[Temps : 35 s | Chrono : 05:30]\\n"
        "Le fonds documentaire d'Aubenas se divise en trois familles. D'un côté, les DMPC récents : des formulaires Cerfa imprimés très standardisés. "
        "De l'autre, les répertoires historiques des anciens géomètres A et B : des registres grands formats entièrement manuscrits. "
        "Et enfin, les actes libres : des plans sur calque avec des annotations manuelles. Cette diversité interdit une solution de lecture unique et impose un découpage modulaire."
    ),
    12: (
        "[Temps : 35 s | Chrono : 06:05]\\n"
        "Une parcelle cadastrale est une entité vivante : au fil des décennies, elle est divisée, réunie ou remembrée. Une parcelle mère donne naissance à plusieurs parcelles filles. "
        "Pour retrouver un bornage de 1970 sur un terrain d'aujourd'hui, notre outil doit impérativement tracer cette filiation cadastrale, afin de rattacher l'acte historique au numéro de parcelle actuellement visible sur le plan."
    ),
    13: (
        "[Temps : 5 s | Chrono : 06:10]\\n"
        "Passons au Chapitre 3 avec l'état de l'art des technologies d'analyse de documents."
    ),
    14: (
        "[Temps : 35 s | Chrono : 06:45]\\n"
        "Pour le texte imprimé, nous avons testé les moteurs OCR conventionnels comme Tesseract et EasyOCR. Tesseract est très rapide sur CPU mais décroche dès que l'image est penchée. "
        "EasyOCR apporte une meilleure détection spatiale grâce à son réseau CRAFT. Ces moteurs sont très efficaces sur les formulaires récents, mais ils échouent complètement sur les mentions manuscrites où le taux d'erreur dépasse 30 %."
    ),
    15: (
        "[Temps : 35 s | Chrono : 07:20]\\n"
        "Pour le manuscrit, l'approche HTR par Transformer change la donne. Le modèle TrOCR développé par Microsoft associe un encodeur d'image et un décodeur linguistique. "
        "Contrairement à l'OCR qui analyse des caractères isolés, TrOCR lit l'écriture cursive de manière séquentielle et continue. Il atteint un taux d'erreur de caractères inférieur à 6 %, à condition de lui fournir des lignes de texte correctement découpées."
    ),
    16: (
        "[Temps : 35 s | Chrono : 07:55]\\n"
        "Une fois le texte transcrit, il faut extraire les champs cibles : commune, date, parcelles. Nous avons comparé LayoutLMv3 et GLiNER. "
        "GLiNER a été retenu : il fonctionne en zero-shot sans nécessiter des milliers d'exemples d'entraînement, et son architecture bidirectionnelle s'exécute en quelques millisecondes sur CPU avec une excellente précision d'extraction."
    ),
    17: (
        "[Temps : 35 s | Chrono : 08:30]\\n"
        "Les modèles Vision-Langage comme LLaVA ou MiniCPM permettent d'analyser conjointement l'image et le texte. Exécutés localement via Ollama, ils sont capables de comprendre des contextes visuels difficiles. "
        "Cependant, leur coût d'inférence sur CPU est élevé : 5 à 8 secondes par document. Nous avons donc fait le choix de ne pas les utiliser sur tout le flux, mais de les réserver comme arbitres de secours."
    ),
    18: (
        "[Temps : 40 s | Chrono : 09:10]\\n"
        "Ce tableau comparatif résume les forces et faiblesses de chaque technologie. On constate qu'aucun outil ne résout la problématique à lui seul. "
        "La force de notre démarche d'ingénieur a été de combiner le meilleur de chaque brique : YOLOv8 pour la segmentation, TrOCR pour le manuscrit, GLiNER pour les entités, et le VLM local uniquement en arbitre sur les cas ambigus. C'est l'architecture que nous découvrons au chapitre 4."
    ),
    19: (
        "[Temps : 5 s | Chrono : 09:15]\\n"
        "Voici le Chapitre 4 détaillant l'architecture logicielle et la chaîne de traitement."
    ),
    20: (
        "[Temps : 50 s | Chrono : 10:05]\\n"
        "Voici la chaîne de traitement complète articulée en 6 étapes indépendantes. Le document brut entre à gauche, il est orienté selon sa nature : DMPC, registre ou acte libre. "
        "Des prétraitements redressent l'image et améliorent le contraste avant la segmentation spatiale. Ensuite, la transcription et l'extraction s'enchaînent avec un score de confiance. "
        "Enfin, les données sont contrôlées avant l'injection API. Ce pipeline est 100 % local : aucune donnée nominative ne quitte la machine de l'entreprise."
    ),
    21: (
        "[Temps : 35 s | Chrono : 10:40]\\n"
        "La détection spatiale est assurée par un réseau YOLOv8 Nano entraîné sur notre fonds. Il repère instantanément les cartouches, les mentions de commune, les dates et les tampons. "
        "Cela nous permet de découper des vignettes ciblées et de ne pas envoyer une image complète de 300 Mo aux moteurs d'extraction, ce qui optimise directement les temps de calcul."
    ),
    22: (
        "[Temps : 35 s | Chrono : 11:15]\\n"
        "Sur ces zones découpées, les moteurs OCR et TrOCR transcrivent le texte brut, puis GLiNER isole les cinq métadonnées obligatoires : commune, section, numéro de parcelle, date d'acte et géomètre signataire. "
        "Un score de confiance est attribué à chaque extraction. Dès que ce score dépasse 0.65, la valeur est transmise directement à la fiabilisation."
    ),
    23: (
        "[Temps : 55 s | Chrono : 12:10]\\n"
        "Cette diapositive illustre pourquoi nous avons fait le choix d'une chaîne hybride plutôt que de tout confier à une IA générative. Voici un cas réel manuscrit : « Lachapelle /s/ AUBENAS ». « /s/ » est l'abréviation locale pour « sous ». "
        "Si l'on envoie cette image à une IA globale entraînée sur la France entière, elle risque de confondre avec Lachapelle-Graillouse, Lachapelle-sous-Chanéac en Ardèche ou La Chapelle-en-Vercors dans la Drôme, voire d'inventer un nom. De plus, son calcul prend 8 secondes sur CPU. "
        "Notre approche hybride associe une lecture rapide sans hallucination et le référentiel des 335 communes d'Ardèche. Le VLM n'intervient qu'en arbitre si la confiance est faible, et sa sortie est strictement verrouillée par la base locale. Résultat : zéro hallucination et conformité cadastrale absolue."
    ),
    24: (
        "[Temps : 5 s | Chrono : 12:15]\\n"
        "Nous arrivons au Chapitre 5 avec la fiabilisation des données et le contrôle humain."
    ),
    25: (
        "[Temps : 40 s | Chrono : 12:55]\\n"
        "Aucune donnée n'est envoyée à l'API sans validation humaine. L'interface Streamlit permet à l'opérateur de visualiser côte à côte le scan original et les champs pré-remplis par l'algorithme. "
        "Les incertitudes sont signalées en couleur. L'opérateur valide en un clic ou ajuste manuellement une valeur douteuse, ce qui divise le travail de saisie par dix tout en conservant une validation juridique rigoureuse."
    ),
    26: (
        "[Temps : 50 s | Chrono : 13:45]\\n"
        "Chaque fiche passe par un moteur automatique de 17 règles métier réparties en 4 familles : les contrôles cadastraux pour vérifier que la commune et la section existent bien dans la base INSEE, "
        "les contrôles temporels pour s'assurer que la date est cohérente avec la période d'exercice du géomètre, les contrôles de personnes pour vérifier le signataire, et les contrôles anti-doublon pour éviter d'enregistrer deux fois le même acte. "
        "Si une règle échoue, une pastille alerte l'opérateur avant toute exportation."
    ),
    27: (
        "[Temps : 35 s | Chrono : 14:20]\\n"
        "Le contrôle spatial s'effectue via une carte interactive Folium intégrée dans l'outil. Les parcelles de la BD Parcellaire y sont affichées en surbrillance. "
        "L'opérateur peut vérifier d'un coup d'œil que le centroïde calculé se positionne exactement sur la bonne parcelle, ou déplacer manuellement le marqueur si le parcellaire a été remanié depuis l'époque du plan."
    ),
    28: (
        "[Temps : 5 s | Chrono : 14:25]\\n"
        "Passons au Chapitre 6 avec la démonstration du pipeline et les résultats sur le terrain."
    ),
    29: (
        "[Temps : 90 s | Chrono : 15:55]\\n"
        "Voici la démonstration vidéo du logiciel en conditions réelles. Vous voyez ici l'opérateur charger un lot de dossiers scannés. L'algorithme détecte immédiatement le type de pièce, extrait les cartouches et pré-remplit les métadonnées. "
        "L'opérateur contrôle les valeurs, consulte la carte de localisation, et clique sur Valider. Le document est alors envoyé directement à l'API Géofoncier, qui confirme l'enregistrement et crée la pastille sur le portail national."
    ),
    30: (
        "[Temps : 35 s | Chrono : 16:30]\\n"
        "Le module d'export communique avec l'API REST Géofoncier via des requêtes sécurisées. Il convertit automatiquement les coordonnées géographiques en projection Lambert-93 réglementaire, "
        "formate le payload JSON avec les attributs requis, et téléverse le fichier PDF source. L'API retourne un identifiant unique qui certifie la publication de l'antériorité."
    ),
    31: (
        "[Temps : 50 s | Chrono : 17:20]\\n"
        "La validation expérimentale a été conduite sur la commune de Prades en Ardèche. 97 dossiers historiques ont été traités : 100 % d'entre eux ont été injectés avec succès sur Géofoncier. "
        "Grâce au filtre des 17 règles métier, aucune fausse pastille n'a été créée. Le temps moyen de traitement a été ramené à moins de 3 minutes par dossier, contre 25 minutes pour la recherche physique manuelle, soit un gain de productivité d'un facteur 6."
    ),
    32: (
        "[Temps : 5 s | Chrono : 17:25]\\n"
        "Voyons avec le Chapitre 7 les limites techniques rencontrées et les perspectives."
    ),
    33: (
        "[Temps : 45 s | Chrono : 18:10]\\n"
        "Les limites du système sont d'abord physiques : les calques anciens créent des transparences recto-verso, et les tampons administratifs masquent parfois les numéros de parcelles. "
        "Sur le registre manuscrit du géomètre B, l'absence de colonnes fixes nécessite plus de retouches. Sur le plan matériel, l'absence de GPU dédié allonge le temps d'inférence du VLM sur CPU. "
        "Malgré ces contraintes, 82 % des dossiers sont validés directement sans aucune retouche par l'opérateur."
    ),
    34: (
        "[Temps : 45 s | Chrono : 18:55]\\n"
        "Les perspectives d'évolution s'articulent autour de 3 axes concrets du mémoire : migrer l'interface vers FastAPI pour supprimer le rechargement de page sur les grands volumes, "
        "connecter les flux WFS de la DGFiP et créer un plugin QGIS pour que les géomètres accèdent aux archives depuis leur outil de production quotidien, et transposer l'outil à d'autres cabinets grâce à un simple fichier de configuration des communes et cartouches. "
        "Enfin, une boucle d'apprentissage actif permettra de réentraîner les modèles locaux au fil des corrections."
    ),
    35: (
        "[Temps : 5 s | Chrono : 19:00]\\n"
        "Pour conclure, voici le Chapitre 8 avec le bilan de ce projet de fin d'études."
    ),
    36: (
        "[Temps : 50 s | Chrono : 19:50]\\n"
        "Au terme de ce PFE, le bilan est très concret pour GEO-SIAPP : le cabinet dispose d'un outil fonctionnel pour valoriser ses 23 600 dossiers d'archives, avec un gain de temps d'un facteur 6 et un coût d'infrastructure nul puisque tout tourne sur les postes existants. "
        "Ce projet s'inscrit pleinement dans la Charte IA adoptée par l'Ordre des Géomètres-Experts ce 1er septembre 2026 : l'algorithme assiste l'opérateur en pré-remplissant les données, mais l'humain conserve l'entière responsabilité juridique de l'acte foncier."
    ),
    37: (
        "[Temps : 15 s | Chrono : 20:05]\\n"
        "Je tiens à remercier chaleureusement Monsieur Gaëtan Hague et toute l'équipe de GEO-SIAPP pour leur accueil et leur confiance, ainsi que Monsieur Mathieu Koehl pour son encadrement à l'INSA Strasbourg. "
        "Je vous remercie pour votre attention et je suis à votre disposition pour vos questions."
    )
}

# Injection des notes dans chaque diapositive
for s_num, s_note in SPEAKER_NOTES.items():
    if s_num <= len(prs.slides):
        slide_target = prs.slides[s_num - 1]
        notes_frame = slide_target.notes_slide.notes_text_frame
        notes_frame.text = s_note

'''

code = code[:start_pos] + new_notes_section + code[end_pos:]

with open('generate_final_180s_deck.py', 'w', encoding='utf-8') as f:
    f.write(code)

print("Verbatim notes injected successfully into generate_final_180s_deck.py!")
