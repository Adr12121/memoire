import matplotlib.pyplot as plt
import numpy as np

# Data (sorted by size for better aesthetic)
labels_original = ['Dupuy', 'Lacour', 'Serret', 'Ract & Ceyte', 'Harrois']
labels = [
    'Géomètre A\n(Actif ~1990-2010)', 
    'Géomètre B\n(Actif ~2000-2020)', 
    'Géomètre C\n(Actif ~1940-1970)', 
    'Géomètre D\n(Actif ~1980-2000)', 
    'Géomètre E\n(Actif ~1970-1990)'
]
sizes = [6600, 6000, 4500, 3500, 3000]

# Elegant color palette
colors = ['#2A4B7C', '#4A7c59', '#D98324', '#A63446', '#5C4A72']

# Create figure and axis
fig, ax = plt.subplots(figsize=(10, 7), facecolor='white')

# Explode the largest slice slightly
explode = (0.05, 0, 0, 0, 0)

# Create a donut chart
wedges, texts, autotexts = ax.pie(
    sizes, 
    explode=explode,
    labels=labels, 
    autopct=lambda pct: f"{pct:.1f}%\n({int(pct/100.*sum(sizes)):,} doss.)".replace(',', ' '),
    startangle=90, 
    colors=colors, 
    pctdistance=0.75,
    labeldistance=1.1,
    wedgeprops=dict(width=0.4, edgecolor='white', linewidth=2)
)

# Style the texts
for t in texts:
    t.set_color('#333333')
    t.set_fontsize(10)
    t.set_weight('bold')

for autotext in autotexts:
    autotext.set_color('white')
    autotext.set_fontsize(9)
    autotext.set_weight('bold')

# Title
plt.title(
    "Répartition estimée du passif documentaire par prédécesseur\n(Volume total : ~23 600 dossiers)", 
    fontsize=14, 
    weight='bold', 
    color='#2c3e50',
    pad=30
)

# Equal aspect ratio ensures that pie is drawn as a circle
ax.axis('equal')  

plt.tight_layout()

# Save the plot
output_path = r'c:\Users\Topo_4\Documents\AT_PFE\repartition_archives_esthetique.png'
plt.savefig(output_path, dpi=300, bbox_inches='tight')
print(f"Diagramme sauvegardé : {output_path}")
