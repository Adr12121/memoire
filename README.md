## Mémoire PFE (LaTeX)

Ce dossier contient un squelette LaTeX prêt à compiler.

### Prérequis (Windows)

- **TeX Live** (recommandé) ou **MiKTeX**
- Optionnel mais pratique: **latexmk**

### Compiler

Depuis ce dossier (`memoire/`) :

```powershell
latexmk -pdf -interaction=nonstopmode -halt-on-error main.tex
```

Pour nettoyer:

```powershell
latexmk -c
```

### Structure

- `main.tex`: document principal
- `tex/`: préambule + infos (titre, auteurs, etc.)
- `chapitres/`: chapitres (introduction, état de l’art, conception, implémentation, résultats, conclusion)
- `annexes/`: annexes
- `biblio.bib`: bibliographie (BibTeX)
- `assets/`: images (schémas, captures, etc.)

### À faire en premier

1) Éditer `tex/00_infos.tex` (titre, auteur, encadrant, établissement, dates).
2) Éditer `chapitres/01_introduction.tex` avec ton contexte + objectifs.
3) Mettre tes images dans `assets/` puis utiliser `\includegraphics`.

