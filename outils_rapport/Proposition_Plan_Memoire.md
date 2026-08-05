Plan provisoire détaillé du Mémoire de PFE

Titre : Développement d'un outil permettant le traitement et l'insertion des archives numériques sur Geofoncier au sein d'un cabinet de Géomètre-Expert

I. INTRODUCTION
1.1. Contexte du projet
- L’importance des archives dans le métier de Géomètre-Expert (foncier, bornage, historique).
- La transition numérique des cabinets.
1.2. La problématique
- La difficulté d'exploitation des registres anciens (manuscrits) et la diversité structurelle des plans modernes numérisés.
- Le coût humain et le risque d'erreur de la saisie manuelle pour Geofoncier.
1.3. Objectifs et livrables
- Créer un processus hybride et autonome de détection, de lecture sémantique et de validation interactive.

II. ANALYSE DU MÉTIER ET DES DONNÉES
2.1. Le cadre de Geofoncier
- Rôle et fonctionnement du portail, intérêt du versement des données pour les confrères.
- Spécifications techniques des données attendues par l'API de l'Ordre (format, tri de pertinence…).
2.2. Typologie des archives à traiter
- Analyse de la bivalence des flux documentaires :
  - Les registres et livrets anciens (format manuscrit, papier, encres).
  - Les plans cadastraux modernes (documents DGFIP, DMPC, PVa).
- Analyse de la complexité des écritures et des variations de mise en page.

III. ÉTAT DE L'ART 
3.1. Vision par ordinateur 
- Détection d’objets : Pourquoi YOLO pour segmenter la structure des documents.
3.2. Reconnaissance de l'écriture manuscrite (HTR)
- Les réseaux de neurones récurrents vs les Transformers.
- Étude comparative : TrOCR vs Kraken vs EasyOCR (en open source).
3.3. L'avènement des modèles Vision-Langage (VLM)
- L'apport décisif des modèles multimodaux (ex: Qwen2-VL, LLaVA) pour la compréhension sémantique du texte dans son contexte visuel.

IV. ARCHITECTURE DU PROJET
4.1. Étape 1 : Classification et routage des documents
- Identification automatique de la nature de l'archive (plan moderne vs registre ancien) pour orienter le pipeline de traitement.
4.2. Étape 2 : Extraction géométrique et segmentation
- Détection des colonnes et cellules via YOLO pour les livrets anciens (algorithme "Ghost Lines" pour reconstruire la structure).
- Recherche de correspondances géométriques et par ancrage de mots-clés pour le recadrage des plans modernes.
4.3. Étape 3 : Moteur d'extraction hybride (HTR & VLM)
- Stratégie multi-hypothèses : Génération de plusieurs prédictions par TrOCR pour les écritures très dégradées.
- Utilisation des modèles Vision-Langage (VLM) en tant que moteur d'extraction sémantique direct pour les métadonnées complexes (N° DA, Section).
4.4. Protection de données sensibles et dispositions mises en place

V. FIABILISATION ET POST-TRAITEMENT DES DONNÉES
5.1. Décodage et correction à la volée
- Utilisation des LogitsProcessors pour interdire les prédictions hors-dictionnaire ou hors communal.
- Normalisation syntaxique des toponymes (gestion des accents, majuscules).
5.2. Matching intelligent et croisement d'archives
- Matching Fuzzy et similarité visuelle : pondération de la distance de Levenshtein par les confusions de lettres courantes.
- Croisement avec des répertoires historiques externes (ex: index des archives Racat & Ceyte) pour consolider la fiabilité des extractions.
5.3. Interface "Human-in-the-Loop" : l'application de validation
- Développement d'une interface de contrôle réactive (Streamlit) garantissant l'intégrité de la donnée.
- Validation des métadonnées via les référentiels officiels (INSEE, liste nationale des géomètres) et auto-sauvegarde des corrections en temps réel.

VI. INTÉGRATION SUR GEOFONCIER ET RÉSULTATS
6.1. Le connecteur API Geofoncier
- Développement d'une communication directe avec les serveurs de Geofoncier (mise en place d'un mode Dry Run de simulation, requêtes multipart).
- Structuration, contrôle et formatage automatiques des paquets de données (JSON/CSV) pour l'API.
6.2. Évaluation des performances
- Définition des métriques : Taux de succès par commune etc ...
- Comparaison de productivité : Gain de temps réel estimé pour le cabinet.
6.3. Analyse des limites
- Gestion des faux positifs et nécessité du contrôle humain final, et à quel degré.

VII. CONCLUSION ET PERSPECTIVES
- Synthèse des contributions techniques.
- Apports personnels du projet (IA, développement, métier de géomètre).
- Ouvertures : automatisation de la lecture des limites de propriétés sur les plans parcellaires ?
