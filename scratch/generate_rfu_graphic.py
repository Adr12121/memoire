import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches

os.makedirs('img', exist_ok=True)
fig, ax = plt.subplots(figsize=(8.5, 7.0), dpi=300)
fig.patch.set_facecolor('#FFFFFF')
ax.set_facecolor('#FFFFFF')
ax.set_xlim(0, 100)
ax.set_ylim(0, 100)
ax.axis('off')

# Header
rect_head = patches.FancyBboxPatch((3, 86), 94, 11, boxstyle='round,pad=0.5,rounding_size=2',
                                    facecolor='#0F172A', edgecolor='none')
ax.add_patch(rect_head)
ax.text(6, 92.2, 'API REST GÉOFONCIER : SPÉCIFICATION DU PAYLOAD /rfuoge', fontsize=11, fontweight='bold', color='#FFFFFF', va='center')
ax.text(6, 88.5, 'Structure JSON obligatoire pour l\'enregistrement et le versement d\'une archive au RFU', fontsize=8.0, color='#94A3B8', va='center')

# Outer Card
rect_card = patches.FancyBboxPatch((3, 3), 94, 81, boxstyle='round,pad=0.5,rounding_size=2',
                                    facecolor='#F8FAFC', edgecolor='#CBD5E1', linewidth=1.2)
ax.add_patch(rect_card)

# Tree fields
fields = [
    ('code_insee', '"07019"', 'string (5 car.)', 'Code Officiel Géographique (COG) de la commune ardéchoise', '#2563EB'),
    ('date_operation', '"1984-05-12"', 'ISO 8601', 'Date d\'acte normée (format AAAA-MM-JJ)', '#0D9488'),
    ('type_acte', '"Bo"', 'code OGE', 'Nature d\'opération : Bo (Bornage) ou DA (Document d\'Arpentage)', '#D97706'),
    ('reference_dossier', '"GEO-1984-089"', 'string', 'Identifiant unique de l\'affaire conservé au cabinet SIAPP', '#6366F1'),
    ('geometry', '{ type, coordinates }', 'GeoJSON', 'Centroïde ou polygone géoréférencé en Lambert-93 (EPSG:2154)', '#DC2626'),
]

y = 75
for name, val, ftype, desc, col in fields:
    # Field card
    f_box = patches.FancyBboxPatch((6, y-8.5), 88, 10, boxstyle='round,pad=0.4,rounding_size=1.5',
                                   facecolor='#FFFFFF', edgecolor='#E2E8F0', linewidth=1)
    ax.add_patch(f_box)
    
    # Left accent bar
    ax.plot([6.2, 6.2], [y-8.2, y+1.2], color=col, linewidth=4.0, solid_capstyle='round')
    
    # Field name and value
    ax.text(9, y-1.5, f'"{name}" :', fontsize=10.5, fontweight='bold', color='#1E293B', fontfamily='monospace')
    ax.text(37, y-1.5, val, fontsize=10.5, fontweight='bold', color=col, fontfamily='monospace')
    
    # Type pill
    pill = patches.FancyBboxPatch((68, y-3.8), 24, 4.5, boxstyle='round,pad=0.2,rounding_size=1',
                                  facecolor='#F1F5F9', edgecolor='#CBD5E1', linewidth=0.7)
    ax.add_patch(pill)
    ax.text(80, y-1.5, ftype, fontsize=7.5, fontweight='bold', color='#475569', ha='center', va='center')
    
    # Description
    ax.text(9, y-6.0, desc, fontsize=7.8, color='#64748B')
    
    y -= 12.5

# Subfield geometry detail
sub_box = patches.FancyBboxPatch((10, 6), 80, 13, boxstyle='round,pad=0.3,rounding_size=1.2',
                                 facecolor='#EFF6FF', edgecolor='#BFDBFE', linewidth=0.9)
ax.add_patch(sub_box)
ax.text(12, 14.5, '└── geometry.type : "Point"', fontsize=8.8, fontweight='bold', color='#1D4ED8', fontfamily='monospace')
ax.text(12, 9.5, '└── geometry.coordinates : [ 801452.35,  6391204.18 ]  (Lambert-93 officiel)', fontsize=8.8, fontweight='bold', color='#1E3A8A', fontfamily='monospace')

plt.tight_layout()
out_path = 'img/arborescence_versement_geofoncier.png'
plt.savefig(out_path, dpi=300, bbox_inches='tight')
print(f'Graphic generated: {out_path}')
