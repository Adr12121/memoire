import matplotlib.pyplot as plt
import matplotlib.patheffects as patheffects
import numpy as np
import os
from docx import Document
from docx.shared import Inches

# Configuration police globale
plt.rcParams['font.family'] = 'sans-serif'

# Données Anonymisées
noms = ['Cabinet A', 'Cabinet B', 'Cabinet C', 'Cabinet D', 'Cabinet E']
dossiers = [6600, 6000, 4500, 3500, 3000]
total = sum(dossiers)

# Couleurs élégantes et modernes (Palette Tailwind : Blue, Emerald, Amber, Rose, Purple)
couleurs = ['#3b82f6', '#10b981', '#f59e0b', '#e11d48', '#8b5cf6']
explode = (0.05, 0.05, 0.05, 0.05, 0.05) # Léger détachement pour tous

# Création de la figure
fig, ax = plt.subplots(figsize=(14, 10), facecolor='#ffffff')
ax.set_facecolor('#ffffff')

# Graphe en anneau (Donut Chart)
wedges, texts, autotexts = ax.pie(
    dossiers, 
    explode=explode, 
    labels=noms,
    autopct=lambda pct: f"{pct:.1f}%\n({int(pct/100.*total):,} doss.)".replace(',', ' '),
    pctdistance=0.82, 
    colors=couleurs, 
    startangle=140,
    textprops=dict(color="white", fontsize=14, fontweight='bold'),
    wedgeprops=dict(width=0.45, edgecolor='#ffffff', linewidth=4)
)

# Style du texte externe (labels)
for text in texts:
    text.set_fontsize(18)
    text.set_color('#1e293b')
    text.set_fontweight('bold')

# Style du texte interne (pourcentages)
for autotext in autotexts:
    # Ajouter une légère ombre portée au texte blanc pour lisibilité
    autotext.set_path_effects([patheffects.withStroke(linewidth=3, foreground='#00000044')])

# Ajout d'un cercle central pour accentuer l'effet Donut
centre_circle = plt.Circle((0,0), 0.55, fc='#ffffff')
fig.gca().add_artist(centre_circle)

# Texte au centre du Donut
plt.text(0, 0.15, "Volume Total", ha='center', va='center', fontsize=20, color='#64748b', style='italic')
plt.text(0, -0.05, f"{total:,}".replace(',', ' '), ha='center', va='center', fontsize=40, color='#0f172a', fontweight='heavy')
plt.text(0, -0.25, "dossiers", ha='center', va='center', fontsize=22, color='#64748b')

# Titre
plt.title(
    "RÉPARTITION DU PASSIF DOCUMENTAIRE", 
    fontsize=28, 
    fontweight='heavy', 
    color='#1e293b',
    pad=30
)

plt.tight_layout()

# Sauvegarde
output_path = r"c:\Users\Topo_4\Documents\AT_PFE\repartition_geometres_premium_anonyme.png"
plt.savefig(output_path, dpi=300, bbox_inches='tight', facecolor=fig.get_facecolor())
print(f"IMAGE GENEREE : {output_path}")

# Génération du document Word
docx_path = r"c:\Users\Topo_4\Documents\AT_PFE\Repartition_Geometres.docx"
doc = Document()
doc.add_heading('Répartition des Dossiers par Cabinet', 0)
doc.add_paragraph("Le graphique ci-dessous illustre la répartition du volume documentaire pour chaque cabinet partenaire de façon anonymisée.")
doc.add_picture(output_path, width=Inches(6.0))
doc.save(docx_path)
print(f"DOCUMENT WORD GENERE : {docx_path}")
