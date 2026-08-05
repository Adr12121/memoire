import docx

def create_clean_word():
    # Création d'un document Word Vierge et propre
    doc = docx.Document()
    
    doc.add_heading("Explication détaillée du pipeline de traitement", level=1)
    
    intro = doc.add_paragraph("Le traitement et la structuration des archives cadastrales (Plans modernes et Livrets historiques) reposent sur une architecture modulaire découpée en six étapes successives. Ce processus automatisé garantit la fiabilité des données extraites tout en respectant les contraintes matérielles et déontologiques de la profession.")

    # Etape 1
    p1 = doc.add_paragraph()
    r1 = p1.add_run("1. Pré-traitement et Classification\n")
    r1.bold = True
    p1.add_run("Le flux de traitement débute par la réception des fichiers bruts (scans PDF ou images). Cette étape est structurante car l'hétérogénéité des archives foncières exige des stratégies d'extraction différentes. L'algorithme de tri (plan_classifier.py) analyse la géométrie globale du document pour déterminer son format. Ce premier routage discrimine les plans modernes (Documents d'Arpentage, DMPC) des registres historiques complexes, orientant chaque archive vers le sous-système d'analyse le plus pertinent.")

    # Etape 2
    p2 = doc.add_paragraph()
    r2 = p2.add_run("2. Détection Spatiale par Intelligence Artificielle (YOLOv8)\n")
    r2.bold = True
    p2.add_run("Une fois le document classifié, l'extraction de l'information s'opère par une approche spatiale. Plutôt que de chercher du texte à l'aveugle, le système emploie le réseau de neurones convolutifs YOLOv8 [Jocher et al., 2023] (via spatial_extractor.py) pour localiser visuellement et découper les zones d'intérêt (cartouches, tableaux de filiation, signatures). Ce choix se justifie par la robustesse de YOLOv8 face à la dégradation des scans historiques et sa rapidité d'inférence. L'IA intervient ici pour « voir » le document et isoler les blocs, sans nécessiter de lourdes ressources de calcul GPU.")

    # Etape 3
    p3 = doc.add_paragraph()
    r3 = p3.add_run("3. Pipeline OCR Hybride et Modèles d'IA Locaux\n")
    r3.bold = True
    p3.add_run("Après le découpage, le système (semantic_ocr_engine.py) procède à l'extraction textuelle par un routage dynamique. Face aux archives mêlant textes dactylographiés et manuscrits, une approche hybride a été implémentée. Pour les textes imprimés, un moteur de reconnaissance optique classique (Tesseract) est privilégié pour sa fiabilité. En revanche, pour le texte manuscrit cursif, le système fait appel à des Modèles de Vision-Langage (VLM) exécutés localement via le moteur Ollama (Llama 3 ou LLaVA). Ces modèles multimodaux interprètent le texte sémantiquement, offrant une robustesse inédite face à l'écriture manuscrite humaine [Tarride et al., 2024].")

    # Etape 4
    p4 = doc.add_paragraph()
    r4 = p4.add_run("4. Vérification de Cohérence et Bases de Connaissance\n")
    r4.bold = True
    p4.add_run("L'extraction par IA générative nécessitant un contrôle strict (pour éviter les hallucinations), une double vérification est opérée (coherence_checker.py). D'abord, une couche de raisonnement local (LLM Ollama) vérifie la sémantique et la logique des données. Ensuite, le système croise ces métadonnées avec des bases de données de référence : validation des communes via le répertoire de l'INSEE et authentification des géomètres via l'annuaire de l'OGE. Une règle métier compile ces retours pour générer un score de confiance (Statut CONFORME, ALERTE, REJET) [Sang, 2024].")

    # Etape 5
    p5 = doc.add_paragraph()
    r5 = p5.add_run("5. Validation Humaine (Human-in-the-Loop)\n")
    r5.bold = True
    p5.add_run("Afin de répondre à une contrainte de qualité stricte et de Privacy-by-Design, l'IA n'agit que comme un assistant ; la responsabilité appartient à l'expert. Les dossiers sont présentés à l'opérateur via une interface visuelle interactive (développée sous Streamlit). Le géomètre peut y inspecter le document d'origine face aux champs extraits (colorés selon leur taux de confiance), puis valider ou corriger les erreurs. Le système intègre un auto-apprentissage pour affiner les modèles sur ces corrections.")

    # Etape 6
    p6 = doc.add_paragraph()
    r6 = p6.add_run("6. Export et Intégration Géofoncier\n")
    r6.bold = True
    p6.add_run("Une fois le dossier consolidé et validé par le professionnel, le système prépare l'intégration. Il associe les données aux codes d'opération attendus (ex: « Division » converti en code Ec). Le dossier, formaté (JSON/PDF) et géolocalisé, est alors téléversé de manière automatisée et sécurisée sur la nouvelle API V2 de Géofoncier (geofoncier_api.py). Cette ultime étape clôture la dématérialisation de l'archive.")

    out_path = "C:/Users/Topo_4/Documents/AT_PFE/Explication_Etapes_Architecture_Propre.docx"
    doc.save(out_path)
    print("Fichier généré :", out_path)

if __name__ == "__main__":
    create_clean_word()
