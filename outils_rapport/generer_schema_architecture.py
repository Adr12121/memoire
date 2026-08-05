import matplotlib.pyplot as plt
import matplotlib.patches as patches
import textwrap
import os

# Configuration globale
plt.rcParams['font.family'] = 'sans-serif'
fig, ax = plt.subplots(figsize=(24, 14))
ax.set_xlim(0, 24)
ax.set_ylim(0, 14)
ax.axis('off')

# Arrière-plan
fig.patch.set_facecolor('#ffffff')
ax.set_facecolor('#ffffff')

def draw_box(x, y, w, h, bg_color, border_color, title, subtitle, desc, step_num):
    # Ombre douce
    for i in range(1, 8):
        alpha = 0.03 if i == 1 else 0.015
        ax.add_patch(patches.FancyBboxPatch((x + i*0.05, y - i*0.05), w, h, boxstyle="round,pad=0.2", fc='#334155', ec='none', alpha=alpha))
    
    # Boîte principale
    ax.add_patch(patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.2", fc=bg_color, ec=border_color, lw=2.5))
    
    # En-tête de la boîte (pour le numéro)
    ax.add_patch(patches.Circle((x + 0.6, y + h - 0.6), 0.7, fc=border_color, ec='white', lw=3, zorder=10))
    ax.text(x + 0.6, y + h - 0.6, str(step_num), color='white', fontweight='bold', fontsize=22, ha='center', va='center', zorder=11)
    
    # Titre
    ax.text(x + 1.6, y + h - 0.4, title, color='#0f172a', fontweight='heavy', fontsize=18, ha='left', va='center')
    
    # Sous-titre (nom du fichier)
    ax.text(x + 1.6, y + h - 0.9, subtitle, color=border_color, fontweight='bold', fontsize=13, ha='left', va='center', style='italic')
    
    # Ligne de séparation
    ax.plot([x + 0.4, x + w - 0.4], [y + h - 1.5, y + h - 1.5], color=border_color, lw=1.5, alpha=0.3)
    
    # Description
    wrapped = textwrap.fill(desc, width=42)
    ax.text(x + w/2, y + h - 1.9, wrapped, color='#334155', fontsize=14, ha='center', va='top', linespacing=1.5)

def draw_arrow(x1, y1, x2, y2, rad=0.0):
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1), 
                arrowprops=dict(arrowstyle="-|>,head_width=0.8,head_length=1.0", 
                                color='#94a3b8', lw=4.5, 
                                connectionstyle=f"arc3,rad={rad}"))

# Titre principal
ax.text(12, 12.8, "TRAITEMENT DES ARCHIVES CADASTRALES", fontsize=32, fontweight='heavy', color='#1e293b', ha='center', va='center')
ax.text(12, 12.0, "Automatisation de l'extraction des données", fontsize=18, color='#64748b', ha='center', va='center', style='italic')

# --- DEFINITION DES BLOCS ---
# Ligne 1 : Etapes 1, 2, 3
draw_box(0.8, 7.5, 6.0, 3.8, '#f0fdf4', '#16a34a', "Classification", "plan_classifier.py", "Réception des fichiers bruts (PDF ou images).\nSéparation entre les plans modernes et les livrets historiques.", 1)
draw_box(9.0, 7.5, 6.0, 3.8, '#eff6ff', '#2563eb', "Détection spatiale", "YOLOv8 & spatial_extractor.py", "Analyse visuelle du document (réseau de neurones YOLOv8).\nDécoupage des zones : cartouches, tableaux, signatures, indications.", 2)
draw_box(17.2, 7.5, 6.0, 3.8, '#fdf4ff', '#c026d3', "Lecture hybride", "semantic_ocr_engine.py", "Choix automatique de l'outil de lecture.\nTesseract pour l'imprimé.\nModèle de langage (Ollama/LLaVA) pour l'écriture manuscrite.", 3)

# Ligne 2 : Etapes 4, 5, 6 (De droite à gauche pour le flow en U)
draw_box(17.2, 1.8, 6.0, 3.8, '#fffbeb', '#d97706', "Contrôle de cohérence", "coherence_checker.py", "Vérification de la logique par l'IA (Ollama LLM).\nRecherche en base des communes (INSEE) et des géomètres.\nCalcul du score de confiance.", 4)
draw_box(9.0, 1.8, 6.0, 3.8, '#fee2e2', '#dc2626', "Validation manuelle", "Streamlit (app_validation.py)", "Interface visuelle de contrôle.\nCorrection des erreurs de lecture.\nApprentissage automatique des corrections.", 5)
draw_box(0.8, 1.8, 6.0, 3.8, '#ffedd5', '#ea580c', "Export Géofoncier", "geofoncier_api.py", "Association finale des codes d'opération.\nSoumission du dossier structuré au portail officiel Géofoncier.", 6)

# --- FLECHES ---
draw_arrow(7.0, 9.4, 8.8, 9.4)      # 1 -> 2
draw_arrow(15.2, 9.4, 17.0, 9.4)    # 2 -> 3

# Flèche descendante avec une légère courbe pour adoucir l'angle
draw_arrow(20.2, 7.3, 20.2, 5.8)    # 3 -> 4

draw_arrow(17.0, 3.7, 15.2, 3.7)    # 4 -> 5
draw_arrow(8.8, 3.7, 7.0, 3.7)      # 5 -> 6

# Footer
plt.tight_layout()
filename = "Architecture_PFE_Premium.png"
plt.savefig(filename, dpi=300, bbox_inches='tight', facecolor=fig.get_facecolor())
print(f"IMAGE GENEREE : {os.path.abspath(filename)}")

filename_jpg = "Architecture_PFE_Premium.jpg"
plt.savefig(filename_jpg, dpi=300, bbox_inches='tight', facecolor=fig.get_facecolor())
print(f"IMAGE GENEREE : {os.path.abspath(filename_jpg)}")
