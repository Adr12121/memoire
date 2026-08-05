import sys
from docx import Document
from docx.shared import Inches, Pt
from docx.chart.data import ChartData
from docx.enum.chart import XL_CHART_TYPE
from docx.enum.chart import XL_LABEL_POSITION

doc = Document()
doc.add_heading('Répartition des archives (Graphique Modifiable)', 0)

doc.add_paragraph(
    "Ce graphique est natif à Word. Faites un clic droit dessus et choisissez "
    "« Modifier les données » pour changer les valeurs ou les textes (noms, années) "
    "directement dans Excel."
)

chart_data = ChartData()
chart_data.categories = [
    'Géomètre A\n(≈ 6 600 doss.)', 
    'Géomètre B\n(≈ 6 000 doss.)', 
    'Géomètre C\n(≈ 4 500 doss.)', 
    'Géomètre D\n(≈ 3 500 doss.)', 
    'Géomètre E\n(≈ 3 000 doss.)'
]
# We use the approximate numbers as the data points
chart_data.add_series('Nombre de dossiers (estimatif)', [6600, 6000, 4500, 3500, 3000])

# Add a pie chart
chart = doc.add_chart(
    XL_CHART_TYPE.PIE, Inches(1), Inches(1), Inches(5.5), Inches(4), chart_data
).chart

chart.has_title = True
chart.chart_title.text_frame.text = "Répartition estimée du passif documentaire par prédécesseur\n(Volume total : ~23 600 dossiers)"
chart.chart_title.text_frame.paragraphs[0].font.size = Pt(12)

# Enable data labels and show percentage
for series in chart.series:
    series.has_data_labels = True
    labels = series.data_labels
    labels.show_percentage = True
    labels.show_value = False
    labels.position = XL_LABEL_POSITION.OUTSIDE_END

output_path = r'c:\Users\Topo_4\Documents\AT_PFE\Graphique_Archives_Modifiable.docx'
doc.save(output_path)
print(f"Document Word créé : {output_path}")
