# -*- coding: utf-8 -*-
"""
Script to:
1. Replace any remaining jargon in generate_final_180s_deck.py (e.g. 'Paradigme Vision-Langage' -> 'Modèles Vision-Langage').
2. Add comprehensive Speaker Notes to all 37 slides with:
   - MOTS-CLÉS
   - DÉBUTS DE PHRASE POUR RELANCER À L'ORAL
   - MESSAGE CENTRAL
"""

with open('generate_final_180s_deck.py', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Clean remaining jargon
code = code.replace('"Paradigme Vision-Langage (LLaVA / MiniCPM)"', '"Modèles Vision-Langage (LLaVA / MiniCPM)"')

# 2. Build the notes dictionary code
notes_code = '''
# ==================== NOTES DU PRÉSENTATEUR POUR L'ORAL (37 DIAPOSITIVES) ====================
SPEAKER_NOTES = {
    1: (
        "MOTS-CLÉS : Soutenance PFE, INSA Strasbourg Topographie, Cabinet GEO-SIAPP Aubenas, Archives foncières, API Géofoncier, Gaëtan Hague, Mathieu Koehl.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Bonjour à tous. Je vous présente aujourd'hui mon projet de fin d'études... »\\n"
        "• « Ce travail a été réalisé au sein du cabinet GEO-SIAPP à Aubenas... »\\n"
        "• « L'objectif central est de concevoir un outil capable de traiter et d'insérer 50 ans d'archives privées sur Géofoncier... »\\n\\n"
        "MESSAGE CLÉ : Présentation officielle et annonce du sujet d'ingénieur."
    ),
    2: (
        "MOTS-CLÉS : 8 chapitres, fil conducteur, du besoin métier au déploiement réel.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « La présentation suit fidèlement la logique du mémoire en 8 étapes... »\\n"
        "• « Nous partirons du cadre réglementaire et des contraintes du cabinet... »\\n"
        "• « Avant d'aborder l'état de l'art, l'architecture développée, la fiabilisation des données... »\\n"
        "• « Pour finir sur la validation sur le terrain et le bilan pour l'entreprise. »\\n\\n"
        "MESSAGE CLÉ : Plan clair, direct et sans détour."
    ),
    3: (
        "MOTS-CLÉS : Chapitre 1, Introduction, Structure d'accueil, Contexte foncier.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Commençons par le premier chapitre avec la présentation de la structure d'accueil et du contexte métier... »\\n\\n"
        "MESSAGE CLÉ : Transition rapide (1 seconde)."
    ),
    4: (
        "MOTS-CLÉS : GEO-SIAPP Aubenas, 4 agences (Guilherand-Granges, Vallon, Pierrelatte), 5 géomètres-experts associés, 23 600 dossiers au siège, 1959 à 2007, fonds A et B.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Le cabinet GEO-SIAPP est une structure historique d'Ardèche et de Drôme... »\\n"
        "• « Le siège d'Aubenas conserve à lui seul plus de 23 000 dossiers d'archives physiques... »\\n"
        "• « Ce fonds rassemble près de 50 ans d'interventions foncières des géomètres prédécesseurs... »\\n\\n"
        "MESSAGE CLÉ : L'entreprise possède un patrimoine documentaire immense mais dormant."
    ),
    5: (
        "MOTS-CLÉS : Monopole loi du 7 mai 1946, Article 646 Code civil, Bornage contradictoire, Décret 96-478 trentenaire, Portail Géofoncier (2010), Programme GEODÉMAT (DGFiP).\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Le cadre juridique du géomètre-expert repose sur un monopole d'État... »\\n"
        "• « L'article 646 du Code civil fonde le droit imprescriptible au bornage contradictoire... »\\n"
        "• « Attention à la distinction clé avec GEODÉMAT : l'État ne numérise que le cadastre fiscal public... »\\n"
        "• « La valorisation des archives privées du cabinet relève exclusivement de notre initiative... »\\n\\n"
        "MESSAGE CLÉ : Les archives privées sont sous la responsabilité exclusive du géomètre-expert."
    ),
    6: (
        "MOTS-CLÉS : « Bornage sur bornage ne vaut », recherche d'antériorité obligatoire, cartons d'archives, 25 à 30 minutes de fouille, risque de double bornage conflictuel.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « En matière de délimitation foncière, le principe fondamental est : Bornage sur bornage ne vaut... »\\n"
        "• « Avant chaque nouvelle affaire, le géomètre est légalement tenu de rechercher les actes anciens... »\\n"
        "• « Actuellement, cette recherche manuelle dans les cartons d'archives prend 25 à 30 minutes par dossier... »\\n"
        "• « L'enjeu est d'éliminer ce temps perdu et de supprimer tout risque d'omission d'un plan antérieur. »\\n\\n"
        "MESSAGE CLÉ : La recherche d'antériorité est une obligation légale chronophage."
    ),
    7: (
        "MOTS-CLÉS : Hétérogénéité des pièces (1959-2007), plans calques, registres manuscrits, secret professionnel, exécution 100% locale, zéro cloud, machine de bureau.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « La problématique centrale du projet se résume ainsi... »\\n"
        "• « Comment extraire automatiquement des données fiables sur des documents très variés, sans aucune infrastructure cloud ? »\\n"
        "• « Le secret professionnel nous interdit formellement d'envoyer les plans clients sur des serveurs externes... »\\n"
        "• « Tout doit tourner localement sur les ordinateurs de bureau existants du cabinet. »\\n\\n"
        "MESSAGE CLÉ : Défi technique = documents dégradés + contrainte stricte de confidentialité locale."
    ),
    8: (
        "MOTS-CLÉS : Chapitre 2, Données et métier, Enjeux documentaires, Spécifications API.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Abordons le Chapitre 2 avec l'analyse métier et les données que nous devons traiter... »\\n\\n"
        "MESSAGE CLÉ : Transition rapide."
    ),
    9: (
        "MOTS-CLÉS : Distinction juridique fondamentale, Cadastre fiscal = inventaire administratif et fiscal non opposable, PV de bornage = limite réelle contradictoire opposable aux tiers.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Il est crucial de rappeler la différence entre cadastre et bornage... »\\n"
        "• « Le cadastre est un document fiscal qui n'a pas valeur de preuve de propriété... »\\n"
        "• « Seul l'acte de bornage contradictoire signé par les riverains fixe juridiquement la limite réelle... »\\n"
        "• « C'est pour cela que les archives du géomètre ont une valeur juridique inestimable. »\\n\\n"
        "MESSAGE CLÉ : L'archive foncière privée prime sur le plan cadastral public."
    ),
    10: (
        "MOTS-CLÉS : Spécifications API REST Géofoncier, schéma JSON, coordonnées Lambert-93, PDF source obligatoire, pastille cartographique, référence dossier.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Pour verser un dossier sur le portail Géofoncier, l'Ordre impose une structure de données stricte... »\\n"
        "• « Le fichier JSON doit obligatoirement comporter la référence, la date, la commune et un centroïde en Lambert-93... »\\n"
        "• « L'acte scanné doit être joint au format PDF pour pouvoir être consulté par les confrères... »\\n\\n"
        "MESSAGE CLÉ : Modèle de données normalisé imposé par l'API nationale."
    ),
    11: (
        "MOTS-CLÉS : 3 familles d'archives, DMPC formulaires Cerfa imprimés, Répertoires manuscrits des prédécesseurs (A et B), Actes libres et plans d'arpentage.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Le fonds documentaire d'Aubenas se divise en trois grandes familles... »\\n"
        "• « Les DMPC récents : formulaires dactylographiés faciles à découper... »\\n"
        "• « Les répertoires des anciens géomètres : registres entièrement manuscrits... »\\n"
        "• « Et les actes libres : plans sur calque avec encres pâlies et ratures. »\\n\\n"
        "MESSAGE CLÉ : Une diversité documentaire qui empêche une solution logicielle unique."
    ),
    12: (
        "MOTS-CLÉS : Filiation parcellaire, arbre généalogique cadastral, parcelles mères et filles, division, réunion, remembrement, traçabilité temporelle.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Une parcelle cadastrale évolue en permanence au fil des successions et des ventes... »\\n"
        "• « Lorsqu'une parcelle se divise, la parcelle mère disparaît au profit de nouvelles parcelles filles... »\\n"
        "• « Pour retrouver un acte de 1970 sur une parcelle actuelle, notre outil doit reconstituer toute la chaîne de filiation. »\\n\\n"
        "MESSAGE CLÉ : Nécessité de remonter la filiation parcellaire pour géolocaliser l'archive ancienne."
    ),
    13: (
        "MOTS-CLÉS : Chapitre 3, État de l'art, Technologies de lecture et d'analyse.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Passons au Chapitre 3 avec l'évaluation comparative des technologies existantes... »\\n\\n"
        "MESSAGE CLÉ : Transition rapide."
    ),
    14: (
        "MOTS-CLÉS : OCR imprimé, Tesseract v5, EasyOCR CRAFT + CRNN, segmentation par caractères, performant sur Cerfa, échec dès qu'il y a du manuscrit.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Pour les documents imprimés, nous avons testé les moteurs OCR classiques... »\\n"
        "• « Tesseract est très léger mais sensible au bruit et aux documents penchés... »\\n"
        "• « EasyOCR s'avère plus robuste pour détecter les cartouches Cerfa... »\\n"
        "• « Mais dès que l'on rencontre une mention manuscrite, le taux d'erreur dépasse 30%. »\\n\\n"
        "MESSAGE CLÉ : L'OCR classique est adapté aux formulaires récents, pas aux archives anciennes."
    ),
    15: (
        "MOTS-CLÉS : HTR manuscrit, TrOCR Microsoft (DeiT + RoBERTa), PyLaia, analyse séquentielle continue, CER < 6% sur lignes isolées.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Pour l'écriture manuscrite, nous sommes passés à l'approche HTR... »\\n"
        "• « Le modèle TrOCR utilise des Transformers pour lire les mots de façon continue... »\\n"
        "• « Il atteint d'excellents résultats sur l'écriture cursive, sous réserve de lui fournir des zones bien découpées. »\\n\\n"
        "MESSAGE CLÉ : TrOCR est l'outil adapté pour déchiffrer les mentions cursives."
    ),
    16: (
        "MOTS-CLÉS : Extraction d'entités nommées (NER), GLiNER bidirectionnel zero-shot, LayoutLMv3 multimodal, légèreté CPU, absence d'entraînement lourd.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Après la transcription, il faut extraire les champs cibles : commune, section, parcelles, date... »\\n"
        "• « Nous avons comparé LayoutLMv3 et GLiNER... »\\n"
        "• « GLiNER a été choisi car il fonctionne en zero-shot sans phase d'entraînement lourde et tourne en quelques millisecondes sur CPU. »\\n\\n"
        "MESSAGE CLÉ : GLiNER permet une extraction ciblée rapide et sans réentraînement."
    ),
    17: (
        "MOTS-CLÉS : Modèles Vision-Langage (VLM), MiniCPM-V, LLaVA sous Ollama local, compréhension globale image-texte, latence CPU 5 à 8s, réservé en secours.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Les modèles Vision-Langage apportent une capacité de raisonnement direct sur l'image... »\\n"
        "• « Déployés en local via Ollama, ils sont capables de déchiffrer des abréviations complexes... »\\n"
        "• « Mais leur temps de calcul sur CPU (5 à 8 secondes) nous interdit de les utiliser sur l'ensemble du flux. »\\n\\n"
        "MESSAGE CLÉ : Le VLM est puissant mais trop lent pour être utilisé partout."
    ),
    18: (
        "MOTS-CLÉS : Tableau comparatif 3.1 du mémoire, complémentarité des briques, YOLOv8 pour segmenter, TrOCR/EasyOCR pour transcrire, GLiNER pour extraire, VLM en arbitre.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Ce tableau de synthèse résume notre stratégie technologique... »\\n"
        "• « Aucun modèle ne peut tout résoudre à lui seul... »\\n"
        "• « La solution consiste à assembler ces briques de façon complémentaire dans une chaîne modulaire. »\\n\\n"
        "MESSAGE CLÉ : Justification rigoureuse du choix de chaque composant."
    ),
    19: (
        "MOTS-CLÉS : Chapitre 4, Architecture logicielle, Chaîne de traitement modulaire.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Nous entrons dans le Chapitre 4 pour détailler l'architecture logicielle développée... »\\n\\n"
        "MESSAGE CLÉ : Transition rapide."
    ),
    20: (
        "MOTS-CLÉS : Figure 4.1 du mémoire, pipeline modulaire 6 étapes, orientation flux (DMPC / registres / actes), exécution 100% locale, secret professionnel garanti.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Voici le schéma directeur de notre chaîne de traitement en 6 étapes... »\\n"
        "• « Le document brut est orienté automatiquement selon sa typologie... »\\n"
        "• « Chaque étape est indépendante et produit un résultat intermédiaire contrôlable... »\\n"
        "• « Et surtout, 100% du traitement s'exécute sur le poste de travail du cabinet, sans aucun flux externe. »\\n\\n"
        "MESSAGE CLÉ : Chaîne industrielle transparente, modulaire et strictement confidentielle."
    ),
    21: (
        "MOTS-CLÉS : Segmentation spatiale YOLOv8, Figure 4.2 du mémoire, détection cartouche, date, tampons, parcelles, découpage précis en vignettes.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « La première étape visuelle repose sur YOLOv8... »\\n"
        "• « Le modèle localise avec précision les différentes zones du document : cartouche, tampons, signatures... »\\n"
        "• « Cela évite de lire toute la page et permet de focaliser les OCR sur des zones nettes. »\\n\\n"
        "MESSAGE CLÉ : Découpage intelligent de la page avant tout traitement textuel."
    ),
    22: (
        "MOTS-CLÉS : Transcription textuelle et GLiNER, OCR/TrOCR, extraction des 5 entités clés, score de confiance, seuil de déclenchement à 0.65.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Une fois les zones isolées, le moteur de lecture extrait le texte brut... »\\n"
        "• « GLiNER identifie les entités cibles et leur attribue un score de confiance... »\\n"
        "• « Si ce score est supérieur à 0.65, la valeur est directement retenue. »\\n\\n"
        "MESSAGE CLÉ : Filtrage automatique basé sur un seuil de confiance statistique."
    ),
    23: (
        "MOTS-CLÉS : Cas réel Lachapelle /s/ Aubenas, piège d'une IA globale nationale (Lachapelle-Graillouse, Chanéac, Vercors), chaîne hybride OCR/HTR + 335 communes Ardèche, VLM en arbitre.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Voici l'exemple le plus parlant qui justifie toute notre démarche d'ingénieur... »\\n"
        "• « Si on confie cette mention à une IA générative globale, elle risque de confondre avec Lachapelle-Graillouse ou La Chapelle-en-Vercors... »\\n"
        "• « Notre chaîne hybride associe un OCR rapide sans hallucination et le référentiel des 335 communes d'Ardèche... »\\n"
        "• « Le VLM n'intervient qu'en arbitre sur l'image si la confiance est faible, et sa réponse reste bridée par la base locale. »\\n\\n"
        "MESSAGE CLÉ : Rigueur cadastrale absolue et vitesse d'exécution garantie par la chaîne hybride."
    ),
    24: (
        "MOTS-CLÉS : Chapitre 5, Fiabilisation des données, Interface opérateur Streamlit.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Nous abordons le Chapitre 5 avec la fiabilisation des données et la place centrale de l'opérateur... »\\n\\n"
        "MESSAGE CLÉ : Transition rapide."
    ),
    25: (
        "MOTS-CLÉS : Interface web Streamlit, relecture contradictoire humaine, pré-remplissage en 1 clic, comparaison visuelle scan / texte, validation obligatoire.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « L'interface Streamlit met en œuvre le principe de validation contradictoire... »\\n"
        "• « L'opérateur visualise côte à côte le scan original et les données extraites... »\\n"
        "• « Les champs incertains sont surlignés, et un clic suffit pour valider ou corriger une valeur. »\\n\\n"
        "MESSAGE CLÉ : L'IA pré-remplit, mais l'humain garde le contrôle complet et valide chaque acte."
    ),
    26: (
        "MOTS-CLÉS : 17 règles métier, 4 familles (cadastre 5, dates 4, personnes 4, dossiers 4), base 335 communes Ardèche, Levenshtein, anti-doublon, alertes visuelles.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Avant tout versement, chaque dossier est soumis à un filtre de 17 règles de cohérence métier... »\\n"
        "• « Contrôles cadastraux avec l'INSEE, vérification des dates avec la période d'exercice du géomètre, anti-doublon... »\\n"
        "• « Toute anomalie déclenche une pastille colorée qui oriente immédiatement l'opérateur vers l'erreur. »\\n\\n"
        "MESSAGE CLÉ : Sécurisation totale du processus par un moteur de règles déterministes."
    ),
    27: (
        "MOTS-CLÉS : Contrôle spatial Folium, carte interactive, superposition OpenStreetMap / BD Parcellaire, positionnement du centroïde, validation géographique.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Le contrôle géographique s'appuie sur une carte interactive Folium intégrée... »\\n"
        "• « L'opérateur vérifie d'un coup d'œil que la pastille se positionne exactement sur la parcelle concernée... »\\n"
        "• « Cela élimine tout risque d'erreur d'aiguillage spatial avant l'injection vers Géofoncier. »\\n\\n"
        "MESSAGE CLÉ : Vérification cartographique immédiate avant publication."
    ),
    28: (
        "MOTS-CLÉS : Chapitre 6, Intégration Géofoncier, Expérimentation réelle.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Nous arrivons au Chapitre 6 avec la démonstration du système et les tests sur le terrain... »\\n\\n"
        "MESSAGE CLÉ : Transition rapide."
    ),
    29: (
        "MOTS-CLÉS : Démonstration vidéo, conditions réelles, pipeline complet, formulaire DMPC Cerfa, registre manuscrit, validation Streamlit, injection API.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Voici la démonstration de l'outil en conditions réelles d'utilisation... »\\n"
        "• « On observe le traitement automatique d'un dossier, la détection des champs, la relecture opérateur... »\\n"
        "• « Et enfin la transmission directe vers l'API avec confirmation de publication. »\\n\\n"
        "MESSAGE CLÉ : Preuve concrète du fonctionnement fluide de la chaîne."
    ),
    30: (
        "MOTS-CLÉS : Protocole API REST Géofoncier, authentification OAuth2, conversion Lambert-93, payload JSON, téléversement PDF, certificat d'injection.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Le protocole d'injection respecte à la lettre les spécifications de l'Ordre des géomètres-experts... »\\n"
        "• « Le script convertit les coordonnées en Lambert-93, génère le JSON et téléverse le PDF source... »\\n"
        "• « L'API retourne un identifiant unique qui certifie l'antériorité sur le portail national. »\\n\\n"
        "MESSAGE CLÉ : Communication réseau robuste et conforme aux normes nationales."
    ),
    31: (
        "MOTS-CLÉS : Expérimentation Prades (Ardèche), 97 dossiers historiques traités, 100% succès injection, 0 fausse pastille, gain facteur 6 (25 min -> < 3 min).\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « L'expérimentation en conditions réelles a été menée sur la commune de Prades... »\\n"
        "• « Sur 97 dossiers traités, le taux de réussite d'injection est de 100%... »\\n"
        "• « Le filtre des 17 règles a bloqué toute erreur : zéro fausse pastille sur Géofoncier... »\\n"
        "• « Et le temps passé par dossier a été divisé par 6 par rapport à la recherche manuelle. »\\n\\n"
        "MESSAGE CLÉ : Validation empirique irréfutable de la rentabilité et de la fiabilité du système."
    ),
    32: (
        "MOTS-CLÉS : Chapitre 7, Analyse des limites, Perspectives d'évolution.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Abordons avec le Chapitre 7 les limites objectives rencontrées et les pistes d'évolution... »\\n\\n"
        "MESSAGE CLÉ : Transition rapide."
    ),
    33: (
        "MOTS-CLÉS : Altérations matérielles, calques transparents recto/verso, tampons recouvrants (chiffres 3/8, 0/6), cursive libre Registre B, contrainte CPU, 82% conformes sans retouche.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « La principale limite du système reste la qualité physique des archives anciennes... »\\n"
        "• « Les calques transparents et les tampons administratifs masquent parfois les numéros... »\\n"
        "• « Sur le Registre B, l'écriture manuscrite libre nécessite plus d'attention humaine... »\\n"
        "• « Malgré cela, 82% des fiches sont validées sans la moindre retouche par l'opérateur. »\\n\\n"
        "MESSAGE CLÉ : Diagnostic lucide des difficultés matérielles et excellente résilience globale."
    ),
    34: (
        "MOTS-CLÉS : Perspectives réelles du mémoire, migration FastAPI / frontend HTML, flux WFS DGFiP et plugin QGIS, transposabilité autres cabinets, apprentissage actif local.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Les perspectives d'évolution s'articulent autour de 3 axes très concrets... »\\n"
        "• « Migrer l'interface vers FastAPI pour éliminer les rechargements de page sur les grands lots... »\\n"
        "• « Développer un plugin QGIS pour permettre aux géomètres d'accéder aux archives depuis leur logiciel métier... »\\n"
        "• « Et transposer la méthode à d'autres cabinets grâce à un simple fichier de configuration des cartouches et des communes. »\\n\\n"
        "MESSAGE CLÉ : Feuille de route technique réaliste et directement applicable."
    ),
    35: (
        "MOTS-CLÉS : Chapitre 8, Conclusion générale, Bilan personnel et professionnel.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Pour terminer, voici le Chapitre 8 avec le bilan général du projet... »\\n\\n"
        "MESSAGE CLÉ : Transition rapide."
    ),
    36: (
        "MOTS-CLÉS : Bilan PFE, 23 600 dossiers valorisables, 100% local, gain facteur 6, 0 € coût cloud, Charte IA OGE (septembre 2026), 4 principes, responsabilité juridique humaine.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « En conclusion de ce travail de six mois, le bilan est extrêmement positif... »\\n"
        "• « Le cabinet GEO-SIAPP dispose désormais d'un outil opérationnel pour valoriser ses 23 600 dossiers dormants... »\\n"
        "• « La recherche d'antériorité passe de 25 minutes à moins de 3 minutes, avec un coût d'infrastructure nul... »\\n"
        "• « Ce travail s'inscrit parfaitement dans la Charte IA adoptée par l'Ordre ce mois-ci : l'IA assiste l'opérateur, mais l'humain conserve la pleine responsabilité de l'acte foncier. »\\n\\n"
        "MESSAGE CLÉ : Double réussite technique et déontologique pour le cabinet et l'élève-ingénieur."
    ),
    37: (
        "MOTS-CLÉS : Remerciements, Gaëtan Hague, GEO-SIAPP Aubenas, Mathieu Koehl, INSA Strasbourg Topographie, ouverture aux questions du jury.\\n\\n"
        "DÉBUTS DE PHRASE POUR RELANCER :\\n"
        "• « Je tiens à adresser mes sincères remerciements à M. Gaëtan Hague et à toute l'équipe de GEO-SIAPP... »\\n"
        "• « Ainsi qu'à M. Mathieu Koehl et au corps professoral de l'INSA Strasbourg pour leur accompagnement... »\\n"
        "• « Je vous remercie pour votre attention et je me tiens à votre disposition pour répondre à toutes vos questions. »\\n\\n"
        "MESSAGE CLÉ : Clôture polie, assurée et ouverture professionnelle aux échanges."
    )
}

# Injection des notes dans chaque diapositive
for s_num, s_note in SPEAKER_NOTES.items():
    if s_num <= len(prs.slides):
        slide_target = prs.slides[s_num - 1]
        notes_frame = slide_target.notes_slide.notes_text_frame
        notes_frame.text = s_note

'''

# Insert the notes logic right before prs.save(output_pptx)
save_marker = 'output_pptx = "Soutenance_PFE_Adrien_TRAVAILLE.pptx"'
if save_marker in code:
    code = code.replace(save_marker, notes_code + '\n' + save_marker)
    print("Notes code inserted before save!")
else:
    print("ERROR: save_marker not found!")

with open('generate_final_180s_deck.py', 'w', encoding='utf-8') as f:
    f.write(code)

print("generate_final_180s_deck.py successfully updated with speaker notes!")
