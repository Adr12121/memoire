import matplotlib.pyplot as plt
import matplotlib.patches as patches
from PIL import Image

# Charger l'image crop
im = Image.open('img/chapelle_sous_aubenas_crop.png')
# Crop précis sur la case commune : Lachapelle /s/ AUBENAS
crop_commune = im.crop((305, 20, 615, 165))
crop_commune.save('img/crop_lachapelle_tight.png')

fig = plt.figure(figsize=(12, 6.0), dpi=300)
fig.patch.set_facecolor('#FFFFFF')

# Ax principal pour les formes
ax = fig.add_axes([0, 0, 1, 1])
ax.set_facecolor('#FFFFFF')
ax.set_xlim(0, 100)
ax.set_ylim(0, 100)
ax.axis('off')

# Boîte gauche : Entrée manuscrite
b_left = patches.FancyBboxPatch((3, 4), 44, 92, boxstyle='round,pad=0.5,rounding_size=2',
                               facecolor='#F8FAFC', edgecolor='#CBD5E1', linewidth=1.2)
ax.add_patch(b_left)
ax.text(6, 90, '1. Mention manuscrite dans le registre', fontsize=12, fontweight='bold', color='#102C57')

# Cadre autour de l'image insérée
b_im_frame = patches.FancyBboxPatch((5.5, 48), 39, 36, boxstyle='round,pad=0.3,rounding_size=1',
                                   facecolor='#FFFFFF', edgecolor='#0284C7', linewidth=1.5)
ax.add_patch(b_im_frame)

# Insérer le zoom de l'écriture
ax_im = fig.add_axes([0.065, 0.49, 0.37, 0.34])
ax_im.imshow(crop_commune)
ax_im.axis('off')

ax.text(6, 42, 'Mention : « Lachapelle /s/ AUBENAS »', fontsize=11, fontweight='bold', color='#0F172A')
ax.text(6, 34, '• Difficulté métier : abréviation ardéchoise (/s/ = sous)', fontsize=9.2, color='#475569')
ax.text(6, 27, '• Échec OCR conventionnel : lit « Aubenas » seul ou échoue', fontsize=9.2, color='#475569')
ax.text(6, 20, '• Confiance GLiNER < 0.65 : déclenchement de l\'arbitre local', fontsize=9.2, color='#D97706', fontweight='bold')
ax.text(6, 11, 'Statut : Alerte de cohérence levée après arbitrage', fontsize=9.2, fontweight='bold', color='#16A34A')

# Flèche centrale
ax.annotate('', xy=(52, 50), xytext=(48, 50),
            arrowprops=dict(arrowstyle='->', lw=3.0, color='#0284C7'))

# Boîte droite : Arbitrage VLM local
b_right = patches.FancyBboxPatch((53, 4), 44, 92, boxstyle='round,pad=0.5,rounding_size=2',
                                facecolor='#F0FDF4', edgecolor='#86EFAC', linewidth=1.2)
ax.add_patch(b_right)
ax.text(56, 90, '2. Arbitrage visuel local (VLM & Base Ardèche)', fontsize=12, fontweight='bold', color='#166534')

ax.text(56, 81, 'Résolution toponymique contextuelle :', fontsize=10.5, color='#475569')

b_res = patches.FancyBboxPatch((56, 55), 38, 22, boxstyle='round,pad=0.4,rounding_size=1.5',
                              facecolor='#FFFFFF', edgecolor='#22C55E', linewidth=1.8)
ax.add_patch(b_res)
ax.text(75, 68, '« La Chapelle-sous-Aubenas »', fontsize=13.5, fontweight='bold', color='#15803D', ha='center', va='center')
ax.text(75, 60, 'Code INSEE officiel : 07058 (100% certifié)', fontsize=10, fontweight='bold', color='#0284C7', ha='center', va='center')

ax.text(56, 45, '• Analyse conjointe du scan et du tableau cadastral', fontsize=9.2, color='#334155')
ax.text(56, 38, '• Rapprochement instantané avec le référentiel INSEE (COG)', fontsize=9.2, color='#334155')
ax.text(56, 31, '• Validation automatique du dossier sans blocage opérateur', fontsize=9.2, color='#334155')

# Bandeau bas de la boîte droite
b_bot = patches.FancyBboxPatch((56, 9), 38, 14, boxstyle='round,pad=0.3,rounding_size=1',
                              facecolor='#DCFCE7', edgecolor='#4ADE80', linewidth=1.0)
ax.add_patch(b_bot)
ax.text(75, 17.5, 'CONFIDENTIALITÉ ET TRAITEMENT LOCAL', fontsize=9.5, fontweight='bold', color='#166534', ha='center', va='center')
ax.text(75, 12.5, 'Exécution 100% locale (Ollama) : secret professionnel garanti', fontsize=8.2, color='#15803D', ha='center', va='center')

out_vlm = 'img/vlm_arbitrage_chapelle_zoom.png'
plt.savefig(out_vlm, bbox_inches='tight', dpi=300)
print('Graphic generated:', out_vlm)
