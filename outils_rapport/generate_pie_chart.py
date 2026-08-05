import matplotlib.pyplot as plt

# Data
labels = ['Géomètre A', 'Géomètre B', 'Géomètre C', 'Géomètre D', 'Géomètre E']
sizes = [6600, 4500, 3000, 3500, 6000]

# Colors (professional palette for PFE)
colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']

# Create pie chart
fig, ax = plt.subplots(figsize=(8, 6))
wedges, texts, autotexts = ax.pie(sizes, labels=labels, autopct='%1.1f%%',
                                  startangle=140, colors=colors, textprops=dict(color="w", weight="bold"))

# Adjust text properties
for t in texts:
    t.set_color('black')
    t.set_fontsize(11)
for autotext in autotexts:
    autotext.set_fontsize(10)

ax.set_title("Répartition estimée du passif documentaire par prédécesseur\n(Total : ~23 600 dossiers)", 
             fontsize=14, weight='bold', pad=20)

# Add a legend to the side
ax.legend(wedges, labels,
          title="Prédécesseurs (Anonymisés)",
          loc="center left",
          bbox_to_anchor=(1, 0, 0.5, 1))

plt.tight_layout()

# Save the plot
output_path = r'c:\Users\Topo_4\Documents\AT_PFE\repartition_archives.png'
plt.savefig(output_path, dpi=300, bbox_inches='tight')
print(f"Diagramme sauvegardé : {output_path}")
