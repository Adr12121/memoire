# Architecture Complète - Pipeline d'Extraction Cadastrale (PFE)

Ce document décrit l'architecture complète du système automatisé d'extraction et d'analyse des archives cadastrales (Plans modernes DA/DMPC et Livrets historiques), en détaillant chaque étape de la chaîne de traitement.

## Schéma d'Architecture (Éditable)

Voici le diagramme d'architecture au format **Mermaid**. Vous pouvez éditer ce code directement pour mettre à jour le schéma. De nombreux éditeurs (comme VS Code, Obsidian, Notion ou GitHub) supportent l'affichage direct des schémas Mermaid.

```mermaid
flowchart TD
    %% Définition des couleurs
    classDef input fill:#fff9c4,stroke:#fbc02d,stroke-width:2px,color:#000
    classDef yolo fill:#b3e5fc,stroke:#0288d1,stroke-width:2px,color:#000
    classDef ocr fill:#e0f7fa,stroke:#006064,stroke-width:2px,color:#000
    classDef validate fill:#fef3c7,stroke:#f59e0b,stroke-width:2px,color:#000
    classDef human fill:#d1fae5,stroke:#10b981,stroke-width:2px,color:#000
    classDef geo fill:#ffe0b2,stroke:#e65100,stroke-width:2px,color:#000

    %% ÉTAPE 1 : ENTRÉE ET CLASSIFICATION
    subgraph Etape_1 [1. Pré-traitement & Classification]
        A1[Documents Bruts<br/>PDF / Images] ::: input --> A2{Classification Intelligente<br/>plan_classifier.py} ::: input
        A2 -->|DA / DMPC| B1
        A2 -->|Livrets historiques| B1
    end

    %% ÉTAPE 2 : DÉTECTION (YOLO)
    subgraph Etape_2 [2. Détection Spatiale - YOLOv8]
        B1[Découpage & Localisation<br/>modern_plan_extractor.py<br/>spatial_extractor.py] ::: yolo
    end

    %% ÉTAPE 3 : PIPELINE OCR HYBRIDE
    subgraph Etape_3 [3. Pipeline OCR Hybride & VLM]
        B1 --> C1{Type de texte} ::: ocr
        C1 -->|Texte Tapé (Imprimé)| C2[Tesseract / TrOCR] ::: ocr
        C1 -->|Texte Manuscrit| C3[Kraken / PyLaia / Modèles VLM] ::: ocr
        
        C2 --> C4[Moteurs d'Extraction<br/>semantic_ocr_engine.py] ::: ocr
        C3 --> C4
    end

    %% ÉTAPE 4 : COHÉRENCE ET NORMALISATION
    subgraph Etape_4 [4. Cohérence et Bases de Connaissance]
        C4 --> D1[Vérification Multi-couches<br/>coherence_checker.py] ::: validate
        D1 <--> D2[(Bases de données<br/>Communes, Géomètres)] ::: validate
    end

    %% ÉTAPE 5 : VALIDATION HUMAINE
    subgraph Etape_5 [5. Human-in-the-Loop]
        D1 --> E1[Interface Streamlit<br/>app_validation.py] ::: human
        E1 -->|Correction & Validation| E2[Dossier Consolidé JSON/CSV] ::: human
    end

    %% ÉTAPE 6 : GÉOFONCIER
    subgraph Etape_6 [6. Export Géofoncier]
        E2 --> F1[API Géofoncier V2<br/>geofoncier_api.py] ::: geo
    end
```

---

## Détail des Étapes du Projet

### 1. Pré-traitement et Classification (`plan_classifier.py`)
Le pipeline ingère des documents bruts (images ou PDF). Le script de classification détermine le format du document reçu pour orienter le traitement. Le système discrimine les plans modernes (Documents d'Arpentage, DMPC) des registres historiques complexes (archives papier des géomètres).

### 2. Détection Spatiale par Intelligence Artificielle (`modern_plan_extractor.py`, `spatial_extractor.py`)
Utilisation d'un modèle d'IA de vision **YOLOv8** (réseau de neurones convolutif entraîné spécifiquement sur des archives cadastrales) pour détecter et découper l'image en zones d'intérêt :
- Cartouches d'en-tête (Commune, Section, etc.)
- Tableaux de filiation (Anciennes/Nouvelles parcelles)
- Signatures
- Mentions d'indication de travaux.
L'IA intervient ici pour "voir" le document et isoler spatialement les blocs pertinents.

### 3. Pipeline OCR Hybride et Modèles d'IA Locaux (`semantic_ocr_engine.py`, `color_ocr_engine.py`)
Le routage dynamique envoie les zones découpées au moteur de lecture adapté :
- **Texte tapé/imprimé** : Géré par des solutions OCR classiques (Tesseract).
- **Texte manuscrit & contextuel** : Fait appel à l'Intelligence Artificielle générative. Le système utilise des modèles de langage multimodaux (VLM) exploités localement via le moteur **Ollama** (ex: *Llama 3* ou *LLaVA* selon la tâche) pour lire et comprendre les écritures cursives ou les paragraphes d'indications. Un modèle **PyLaia** entraîné peut également être invoqué. Un mécanisme de *"survivabilité"* assure un résultat même si l'IA locale échoue.

### 4. Vérification de Cohérence et BDD (`coherence_checker.py`, `repertoire_lookup.py`)
Les métadonnées extraites subissent une double vérification :
- **Raisonnement LLM (Ollama)** : L'IA intervient une seconde fois comme vérificateur sémantique (couche de raisonnement local) pour statuer sur la logique des informations extraites.
- **Bases de données métier** : Nettoyage, recherche des géomètres (`GEOMETRES_CONNUS`) et alignement des noms de communes avec la base de données INSEE officielle pour contrer les hallucinations OCR.
- Une règle métier compile ces retours (IA + Bases de données) pour générer un score de confiance (Statut CONFORME, ALERTE, REJET).

### 5. Validation "Human-in-the-loop" (`app_validation.py`)
Afin de répondre à une contrainte de qualité stricte (Privacy-by-Design et fiabilité), une application **Streamlit** (Split-Screen) offre à l'utilisateur :
- La visionneuse de documents avec le zoom interactif et l'affichage des bounding boxes (boîtes de détection).
- Les champs extraits pré-remplis, colorés par taux de confiance.
- La validation, la correction manuelle et un auto-apprentissage des corrections fréquentes.

### 6. Intégration Géofoncier (`geofoncier_api.py`, `export_geofoncier.py`)
Le "dossier consolidé" validé par l'humain est converti au format attendu par le portail national. 
- Mapping des codes d'opération (ex: "Division" devient l'Operation Code `Ec`).
- Soumission automatique et sécurisée du document (JSON / PDF) à la nouvelle API V2 Géofoncier avec géolocalisation correcte.

---

> **Note à l'utilisateur :**
> Ce fichier `.md` (Markdown) est le format standard pour documenter le code. Il vous permet de modifier le texte explicatif et de mettre à jour le diagramme Mermaid au fur et à mesure des évolutions de votre projet. Les éditeurs Markdown vous offrent la possibilité d'exporter cela facilement en PDF si nécessaire.
