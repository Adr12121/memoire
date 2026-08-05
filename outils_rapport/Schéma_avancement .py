import matplotlib.pyplot as plt
import matplotlib.patches as patches
import textwrap
import os

# 1. CONFIGURATION 
fig, ax = plt.subplots(figsize=(22, 12))
ax.set_xlim(0, 22)
ax.set_ylim(0, 12)
ax.axis('off')

# 2. FONCTIONS DE DESSIN 
def draw_box_simple(x, y, w, h, color, title, desc, step_num):
    # Ombre
    ax.add_patch(patches.FancyBboxPatch((x+0.1, y-0.1), w, h, boxstyle="round,pad=0.1", fc='#cccccc', ec='none'))
  
    ax.add_patch(patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.1", fc=color, ec='#333', lw=1.5))
    # titre
    ax.text(x + w/2, y + h - 0.5, title, ha='center', fontweight='bold', fontsize=11)
    
    wrapped = textwrap.fill(desc, width=30)
    ax.text(x + w/2, y + h - 1.5, wrapped, ha='center', va='top', fontsize=9, style='italic', color='#444')
    
    ax.add_patch(patches.Circle((x, y+h), 0.4, fc='#d32f2f', ec='white', zorder=10))
    ax.text(x, y+h, str(step_num), color='white', fontweight='bold', ha='center', va='center', zorder=11)

def draw_arrow(x1, y1, x2, y2, color='#444'):
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1), 
                arrowprops=dict(arrowstyle="->", color=color, lw=2.5, mutation_scale=25))

def draw_curved_arrow(x1, y1, x2, y2, rad=0.3):
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1), 
                arrowprops=dict(arrowstyle="->", color='#333', lw=2, 
                                connectionstyle=f"arc3,rad={rad}", mutation_scale=20))



# TITRE
plt.title("SCHEMA GENERAL DES DIFFERENTES ETAPES DU PROJET", fontsize=20, fontweight='bold', pad=20, color='#1565c0')

# ÉTAPE 1 : PRÉPARATION
draw_box_simple(0.5, 4.5, 3.5, 5, '#fff9c4', "1. PRÉ-TRAITEMENT", "Améliorer la qualité du scan (binarisation, redressement, nettoyage).", 1)
ax.text(2.25, 5.5, "[+] Outils à définir", ha='center', fontsize=9, fontweight='bold', color='#fbc02d')

# ÉTAPE 4 : GÉOFONCIER
draw_box_simple(18.0, 4.5, 3.5, 5, '#ffe0b2', "4. GÉOFONCIER", "Renseigner tout ce dont on a besoin (JSON/PDF).", 4)


# CŒUR DU SYSTÈME (DOCKER)


# Cadre Docker
ax.add_patch(patches.Rectangle((4.8, 2.0), 12.5, 9.0, linewidth=2, ec='#0277bd', fc='#e1f5fe', linestyle='--'))
ax.text(11.0, 11.2, "ENVIRONNEMENT DOCKER", ha='center', fontsize=12, fontweight='bold', color='#0277bd', backgroundcolor='white')

# GRANDE BOITE ÉTAPE 2 
ax.add_patch(patches.FancyBboxPatch((5.5, 3.0), 8.5, 7.5, boxstyle="round,pad=0.1", fc='#b3e5fc', ec='#333', lw=2))
ax.add_patch(patches.Circle((5.5, 10.5), 0.5, fc='#d32f2f', ec='white', zorder=10))
ax.text(5.5, 10.5, "2", color='white', fontweight='bold', ha='center', va='center', zorder=11, fontsize=14)

ax.text(9.75, 10.0, "DETECTER ET LIRE LES CARACTERES", ha='center', fontweight='bold', fontsize=13) 

# 2A. LE ROUTAGE 
ax.add_patch(patches.Rectangle((6.0, 4.5), 2.0, 4.5, fc='#ffffcc', ec='#fbc02d', lw=2))
ax.text(7.0, 8.6, "DETECTION DES\nZONES", ha='center', va='center', fontweight='bold', fontsize=9)
ax.text(7.0, 8.0, "(Découpage)", ha='center', fontsize=8, style='italic')

# Texte wrappé pour 2A
desc_2a = "Sépare le manuscrit de l'imprimé pour envoyer vers le bon algo."
wrapped_2a = textwrap.fill(desc_2a, width=15) # Force le retour à la ligne
ax.text(7.0, 6.5, wrapped_2a, ha='center', va='center', fontsize=8, color='#444')

ax.text(7.0, 5.0, "[+] YOLO\nou LayoutLM", ha='center', fontsize=8, fontweight='bold', color='#e65100')

# --- 2B. BRANCHE HAUTE (TEXTE TAPPÉ) ---
ax.add_patch(patches.Rectangle((9.0, 7.0), 4.5, 2.0, fc='#ffffff', ec='#0288d1', lw=1))
ax.text(9.1, 8.7, "TEXTE TAPÉ", ha='left', va='top', fontweight='bold', color='#0288d1', fontsize=9)

# Ton texte spécifique pour le haut
text_haut = "Seulement comment a été établi le document, le mot commune pour le repérer, section, feuille, n° DA, dressé par..."
wrapped_haut = textwrap.fill(text_haut, width=48)
ax.text(9.1, 8.4, wrapped_haut, ha='left', va='top', fontsize=7.5, style='italic', color='#555')

ax.text(13.4, 7.1, "TESSERACT ? un autre ?", ha='right', fontweight='bold', fontsize=10)


# 2C. BRANCHE BASSE (MANUSCRIT)
ax.add_patch(patches.Rectangle((9.0, 4.5), 4.5, 2.0, fc='#e0f7fa', ec='#006064', lw=2))
ax.text(9.1, 6.2, "TEXTE MANUSCRIT", ha='left', va='top', fontweight='bold', color='#006064', fontsize=9)

# Ton texte spécifique pour le bas
text_bas = "Nom de la commune, n° section, N° du DA, date, nom du GE, si possible limite. Autour de 2010 : pas besoin"
wrapped_bas = textwrap.fill(text_bas, width=48)
ax.text(9.1, 5.9, wrapped_bas, ha='left', va='top', fontsize=7.5, style='italic', color='#555')

ax.text(13.4, 4.6, "PyLaia / Kraken", ha='right', fontweight='bold', fontsize=10, color='#d81b60')

# FLÈCHES INTERNES  

draw_arrow(4.0, 6.75, 6.0, 6.75)
draw_curved_arrow(8.0, 7.5, 9.0, 8.0, rad=-0.3)
draw_curved_arrow(8.0, 6.0, 9.0, 5.5, rad=0.3)

# ÉTAPE 3 : FUSION 
desc_3 = "Réassemblage texte + position. Indiquer ce qui correspond à quoi (Identifier chaque donnée)."
draw_box_simple(14.5, 4.5, 3.0, 5, '#c8e6c9', "3. FUSION", desc_3, 3)
ax.text(16.0, 6.5, "Format cible\nà définir", fontweight='bold', ha='center', color='#2e7d32')

# FLÈCHES EXTERNES

draw_arrow(4.0, 7.0, 5.5, 7.0)    
draw_arrow(13.5, 8.0, 14.5, 8.0)  
draw_arrow(13.5, 5.5, 14.5, 5.5)  
draw_arrow(17.5, 7.0, 18.0, 7.0)  

# SAUVEGARDE
            
plt.tight_layout()
filename = "architecture_pfe_finale_fixed.png"
plt.savefig(filename, dpi=300, bbox_inches='tight')

print(f" IMAGE GÉNÉRÉE : {os.path.abspath(filename)}")
plt.show()