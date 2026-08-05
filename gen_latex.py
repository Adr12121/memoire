import os
import shutil

memoire_dir = r'C:\Users\Topo_4\Documents\AT_PFE\memoire'
os.chdir(memoire_dir)

# Backup main.tex
if os.path.exists('main.tex'):
    shutil.copy('main.tex', 'main_old_report_class.tex')

# Ensure img directory exists
os.makedirs('img', exist_ok=True)

main_content = r'''\documentclass[a4paper,11pt]{article}

\usepackage[T1]{fontenc}
\usepackage[utf8]{inputenc}
\usepackage{csquotes}
\usepackage[authoryear,round]{natbib}
\usepackage[french]{babel}
\usepackage{geometry}
\usepackage{hyperref}
\usepackage{xcolor}
\usepackage{graphicx}
\usepackage{fancyhdr}
\usepackage{array}
\usepackage{booktabs}
\usepackage{enumitem}
\usepackage{tikz}
\usetikzlibrary{positioning,arrows.meta,shapes.geometric}
\usepackage[expansion=false]{microtype}
\usepackage{amsmath}
\usepackage{float}

\addto\captionsfrench{\renewcommand{\contentsname}{Sommaire}}

\geometry{
  a4paper,
  left=2.5cm,
  right=2.5cm,
  top=2.5cm,
  bottom=2.5cm,
  centering
}

\setlength{\headheight}{15pt}
\setlength{\headsep}{12pt}
\hypersetup{hidelinks}

\renewcommand{\sectionmark}[1]{\markright{#1}}
\fancyhf{}
\fancyhead[R]{\nouppercase{\rightmark}}
\fancyfoot[C]{\thepage}
\renewcommand{\headrulewidth}{0pt}
\renewcommand{\footrulewidth}{0.4pt}
\fancypagestyle{plain}{%
  \fancyhf{}
  \fancyhead[R]{\nouppercase{\rightmark}}
  \fancyfoot[C]{\thepage}
  \renewcommand{\headrulewidth}{0pt}
  \renewcommand{\footrulewidth}{0.4pt}
}

\newcommand{\titreEtude}{Développement d'un outil permettant le traitement et l'insertion des archives numériques sur Geofoncier au sein d'un cabinet de Géomètre-Expert}
\newcommand{\auteur}{TRAVAILLÉ Adrien}
\newcommand{\specialite}{INSA Strasbourg --- Spécialité Topographie}
\newcommand{\structure}{Cabinet de Géomètre-Expert (structure d'accueil : GEO-SIAPP)}
\newcommand{\encadrement}{Encadrant : Mathieu KOEHL}
\newcommand{\dateRapport}{2026}

\title{\textbf{Mémoire de Fin d'Études (PFE)}\\[0.4em]\large \titreEtude}
\author{\auteur\\\specialite\\\structure\\\encadrement}
\date{\dateRapport}

\begin{document}

\begin{titlepage}
  \thispagestyle{empty}
  \noindent% <-- important : le % supprime l'espace parasite
  \begin{minipage}[c]{0.45\textwidth}
    \includegraphics[height=2.8cm]{Logo_INSAStrasbourg.jpg}
  \end{minipage}% <-- important : supprime le saut de ligne
  \hfill
  \begin{minipage}[c]{0.45\textwidth}
    \raggedleft
    \includegraphics[height=2.8cm]{Geosiapp.jpg}
  \end{minipage}

  \vspace{4em}
  \begin{center}
    {\Huge\bfseries\sffamily Mémoire de Fin d'Études\par}
    \vspace{0.6em}
    {\Large\bfseries\sffamily (PFE)\par}
    \vspace{1em}
    
    {\rule{0.35\linewidth}{1pt}\par}
    \vspace{2.5em}
    
    % Mise en valeur du titre : en gras, plus grand, et bien espacé
    {\LARGE\bfseries \titreEtude\par}
    
    \vspace{3.5em}
    {\large \textbf{\auteur}\par}
    \vspace{0.5em}
    {\large \specialite\par}
    {\large \structure\par}
    {\large \encadrement\par}
    \vfill
    {\large \dateRapport\par}
  \end{center}
\end{titlepage}

\cleardoublepage
\pagenumbering{gobble}
\pdfbookmark[1]{\contentsname}{toc}
\tableofcontents
\clearpage

\input{00_Remerciements.tex}

\clearpage
\pagestyle{fancy}
\pagenumbering{arabic}
\setcounter{page}{1}

\input{01_Introduction.tex}
\clearpage
\input{02_Analyse_Metier_Donnees.tex}
\clearpage
\input{03_Etat_de_lart.tex}
\clearpage
\input{04_Architecture.tex}
\clearpage
\input{05_Fiabilisation.tex}
\clearpage
\input{06_Integration_Geofoncier.tex}
\clearpage
\input{07_Conclusion.tex}
\clearpage

\addcontentsline{toc}{section}{Bibliographie}
\bibliographystyle{plainnat}
\bibliography{PFE_AT}
\clearpage

\input{08_Annexes.tex}

\end{document}
'''

with open('main.tex', 'w', encoding='utf-8') as f:
    f.write(main_content)

files_content = {
    '00_Remerciements.tex': r'''\section*{Remerciements}
\addcontentsline{toc}{section}{Remerciements}

Je tiens tout d'abord à remercier...
''',

    '01_Introduction.tex': r'''\section{INTRODUCTION}

\subsection{Contexte du projet}
La profession de géomètre-expert est réglementée, tout comme la gestion de la donnée foncière qu'elle produit et dont elle dispose de l'exclusivité. Selon l'article 55 du décret n°96-478 du 31 mai 1996 (portant sur les devoirs professionnels), le géomètre-expert a l'obligation de conserver et de tenir à jour les archives et documents topographiques fixant les limites des biens fonciers. Ces documents (procès-verbaux de bornage, plans de division, documents d'arpentage) sont alors classifiés et conservés par les cabinets.

\subsection{La problématique}
Le cabinet Géo-Siapp dispose d'un fonds d'archives très important, accumulé au fil des décennies (estimé pour la période 1959-2007) et réparti sur ses différentes agences en Ardèche. Ces dossiers sont extrêmement hétérogènes : plans papiers manuscrits, registres anciens, ou vieux documents scannés sous forme d'images. 
La majorité du passif n'est ni indexée, ni géo-référencée, encore moins versée sur Geofoncier. La recherche manuelle dans ces papiers est une tâche très chronophage.

\subsection{Objectifs et livrables}
L'objectif de ce projet est donc de concevoir un système automatisé capable d'analyser ces archives brutes numérisées pour en extraire les informations clés (la commune, la date, le numéro de document d'arpentage...). Ce système doit permettre de structurer l'information pour générer automatiquement des dossiers locaux, et préparer un versement massif sur la plateforme Géofoncier. Pour garantir la qualité de la donnée foncière, chaque étape d'extraction doit être vérifiable par les opérateurs grâce à une interface de validation, permettant de corriger les erreurs de la machine.
''',

    '02_Analyse_Metier_Donnees.tex': r'''\section{ANALYSE DU MÉTIER ET DES DONNÉES}

\subsection{Le cadre de Geofoncier}
Lorsqu'une nouvelle mission foncière est confiée au cabinet, la première étape pour l'opérateur consiste à effectuer une recherche ou une reconnaissance des limites. Pour ce faire, il se connecte systématiquement sur \textbf{Géofoncier}, le portail national de l'Ordre des géomètres-experts, afin de vérifier si une opération foncière a déjà eu lieu sur les parcelles concernées ou à proximité. Si l'opération est ancienne, il doit rechercher physiquement dans le stock d'archives du cabinet.

\subsection{Typologie des archives à traiter}
Le périmètre d'entrée englobe de nombreux types de documents différents, allant du document PDF vectoriel récent aux images scannées (raster) de qualité parfois médiocre, formant des lots d'archives non homogènes. En sortie, le système doit générer des données toujours identiques (fichiers CSV, JSON ou Excel), accompagnées de preuves visuelles pour le contrôle, comme des images annotées.

La dégradation des documents sources (bruit numérique, compression, mauvaise rotation, faible contraste, mauvaise écriture) complique la tâche de segmentation et de lecture. D'autre part, la variabilité des mises en page demande une analyse différente en fonction du type de document.
''',

    '03_Etat_de_lart.tex': r'''\section{ÉTAT DE L'ART}

L'étude bibliographique a été importante pour orienter les choix architecturaux du projet. Les recherches récentes démontrent sans équivoque que le traitement des documents anciens et de leurs mises en page complexes nécessite une combinaison d'approches, allant de la simple détection de zones aux modèles d'intelligence artificielle avancés.

\subsection{Vision par ordinateur (Détection d'objets)}
Pour l'extraction spécifique des mots-clés, la détection purement visuelle et géométrique a été considérablement renforcée par l'évolution des algorithmes YOLO. Excellente détection spatiale des blocs (en-têtes, cases), robuste aux variations de fond de l'image.

\subsection{Reconnaissance de l'écriture manuscrite (HTR)}
Contrairement aux documents modernes générés numériquement, les registres et plans cadastraux anciens présentent d'importantes variations visuelles et structurelles. La reconnaissance de texte manuscrit (HTR) est fortement dépendante de la qualité des données d'entraînement. Les moteurs d'extraction classiques basés uniquement sur la reconnaissance optique de caractères (OCR) affichent très souvent leurs limites face aux écritures anciennes. Plusieurs études récentes recommandent l'utilisation d'outils complémentaires comme PyLaia ou des procédés hybrides (TrOCR, etc.).

\subsection{L'avènement des modèles Vision-Langage (VLM)}
L'utilisation de modèles Vision-Langage (VLM) permet d'apporter des capacités de raisonnement et de correction d'erreurs (joue le rôle d'arbitre ou de \enquote{troisième avis}) directement sur l'image du document. Ces modèles interviennent en dernier recours pour trancher un doute et garantir à l'opérateur final la proposition d'extraction la plus pertinente possible.
''',

    '04_Architecture.tex': r'''\section{ARCHITECTURE DU PROJET}

\subsection{Étape 1 : Segmentation et géométrie des pages}
Le script principal analyse le document d'entrée pour identifier sa nature et le rediriger vers le module d'extraction adéquat. L'algorithme YOLO a été intégré pour détecter les différentes zones d'intérêt (tableaux, lignes). Ce découpage est utile pour ne pas polluer l'OCR avec des informations parasites.

\subsection{Étape 2 : Moteur de transcription Multi-modèles}
Le programme isole les informations ciblées. Pour les plans cadastraux récents, un module dédié (\texttt{modern\_plan\_extractor.py}) est mis en place. Pour les registres anciens, l'extraction repose sur des modèles d'intelligence artificielle. Le modèle GLiNER a été privilégié en raison de son fonctionnement hors-ligne performant sur CPU. Il permet d'extraire la commune ou le numéro de document en se basant sur le sens de la phrase, contournant ainsi les fautes de frappe de l'OCR classique.

\subsection{Étape 3 : Arbitrage par VLM}
Inspiré de travaux récents, l'usage d'un modèle VLM a été mis en place pour jouer un rôle de dernier recours sur les cas où le texte est mal reconnu par l'OCR, ou en cas d'ambiguïté, en se basant sur le contexte visuel global.

\subsection{Protection de données sensibles}
Tous les traitements (lecture d'image, intelligence artificielle, extraction) sont exécutés localement pour des raisons de confidentialité. Les données sensibles du cabinet ne transitent pas par des serveurs tiers.
''',

    '05_Fiabilisation.tex': r'''\section{FIABILISATION ET POST-TRAITEMENT DES DONNÉES}

\subsection{Décodage quand l'écriture est difficile}
Afin de pallier les erreurs de lecture liées à la machine (confusion de caractères, zones de texte non reconnues), le développement s'est orienté vers la création d'un moteur de vérification. 

\subsection{Matching intelligent et Validation via les référentiels officiels}
Ce module analyse le texte brut extrait par l'intelligence artificielle pour le corriger automatiquement. Il effectue notamment un croisement entre les noms de communes détectés et la base de données officielle de l'INSEE pour en rectifier l'orthographe, et s'assure du bon formatage des éléments. 

\subsection{Interface de validation et contrôle humain}
Une interface visuelle interactive a été mise en place pour les opérateurs du cabinet avec le framework Python \textbf{Streamlit}. L'application affiche uniquement un encadré zoomé sur la zone précise du document où l'ordinateur a trouvé l'information, avec le texte déchiffré affiché juste à côté. L'opérateur valide (via la touche Entrée) ou corrige si l'information est erronée. Chaque modification apportée par l'humain est enregistrée par le logiciel afin d'améliorer progressivement la précision de l'outil.
''',

    '06_Integration_Geofoncier.tex': r'''\section{INTÉGRATION SUR GEOFONCIER ET RÉSULTATS}

\subsection{Structuration et export}
Les données extraites sont exportées (fichiers CSV placés dans le dossier \texttt{outputs/}) et serviront à générer des dossiers locaux selon la commune et le numéro de parcelle identifiés.

\subsection{Préparation au versement API Géofoncier}
Le système devra communiquer avec l'API pour générer la \enquote{pastille} correspondante directement sur la carte de la plateforme. Cela nécessite de structurer les informations validées dans un format JSON, et de s'authentifier de manière sécurisée auprès des serveurs de l'Ordre des géomètres-experts via la génération d'un jeton d'authentification (Token).

\subsection{Évaluation des performances et limites}
Le temps de traitement informatique pur chute à quelques minutes par document. Avec le temps nécessaire à la validation humaine via l'interface (estimé à une dizaine de secondes en moyenne par information), le gain de temps par rapport à une saisie manuelle est très important.
Les limites résident dans les \enquote{faux positifs} et la nécessité d'un contrôle humain final avant tout versement définitif.
''',

    '07_Conclusion.tex': r'''\section{CONCLUSION ET PERSPECTIVES}

\subsection{Synthèse des contributions techniques}
Ce projet a permis de développer un système d'extraction complet et hybride, capable de traiter à la fois des plans modernes structurés et des livrets d'archives entièrement manuscrits. Grâce à la combinaison de YOLO pour la segmentation visuelle, GLiNER pour l'extraction sémantique et des VLM pour l'arbitrage, l'outil propose une approche fiable face aux erreurs d'OCR classiques.

\subsection{Perspectives}
L'objectif final est de permettre la création automatisée des dossiers locaux et le versement sans friction des informations sur le portail des géomètres-experts (Géofoncier).
''',

    '08_Annexes.tex': r'''\appendix
\section*{Annexes}
\addcontentsline{toc}{section}{Annexes}

\subsection{Schéma d'architecture globale}
\begin{figure}[h!]
  \centering
  \includegraphics[width=1\textwidth]{img/architecture_pfe_complet.pdf}
  \caption{Vue détaillée de l'architecture globale du projet}
\end{figure}

\subsection{Planning et suivi hebdomadaire}
\begin{figure}[h!]
  \centering
  \includegraphics[width=1\textwidth]{img/planning.jpg}
  \caption{Vue détaillée du fichier de suivi hebdomadaire}
\end{figure}
'''
}

for filename, content in files_content.items():
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(content)

print("All LaTeX files created successfully.")
