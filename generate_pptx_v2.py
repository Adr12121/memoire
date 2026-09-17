# -*- coding: utf-8 -*-
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image

C_NAVY=RGBColor(15,37,70); C_ACCENT=RGBColor(2,132,199); C_WHITE=RGBColor(255,255,255)
C_DARK=RGBColor(15,23,42); C_MUTED=RGBColor(100,116,139); C_LIGHT=RGBColor(248,250,252)
C_BORDER=RGBColor(226,232,240); C_SUCCESS=RGBColor(22,163,74); C_WARN=RGBColor(217,119,6)
C_RED=RGBColor(220,38,38)
CHAP_COLORS=[RGBColor(15,37,70),RGBColor(2,132,199),RGBColor(88,28,135),RGBColor(6,95,70),RGBColor(154,52,18),RGBColor(15,118,110),RGBColor(109,40,217),RGBColor(15,37,70)]
CHAP_SHORT=["1. INTRO","2. METIER","3. ETAT ART","4. ARCHI","5. FIAB.","6. INTEGR.","7. LIMITES","8. CONCL."]
SLIDE_W=13.333; SLIDE_H=7.5; TOTAL=30
prs=Presentation(); prs.slide_width=Inches(SLIDE_W); prs.slide_height=Inches(SLIDE_H); blank=prs.slide_layouts[6]

def bg_w(s): f=s.background.fill; f.solid(); f.fore_color.rgb=C_WHITE
def bg_d(s,c=None): f=s.background.fill; f.solid(); f.fore_color.rgb=(c or C_NAVY)
def rect(s,x,y,w,h,fill=None,line=None,lw=None):
    shp=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x),Inches(y),Inches(w),Inches(h))
    if fill: shp.fill.solid(); shp.fill.fore_color.rgb=fill
    else: shp.fill.background()
    if line: shp.line.color.rgb=line; shp.line.width=(lw or Pt(0.5))
    else: shp.line.fill.background()
    return shp
def tb(s,x,y,w,h): return s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h))
def txt(s,x,y,w,h,t,sz=11,b=False,c=None,al=PP_ALIGN.LEFT,it=False):
    bx=tb(s,x,y,w,h); tf=bx.text_frame; tf.word_wrap=True
    tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=Inches(0.02)
    p=tf.paragraphs[0]; p.text=t; p.alignment=al; r=p.runs[0]
    r.font.size=Pt(sz); r.font.bold=b; r.font.italic=it; r.font.color.rgb=(c or C_DARK)
    return bx
def addimg(s,path,x,y,mw,mh,cap=None):
    if not path or not os.path.exists(path): return
    try:
        with Image.open(path) as im: iw,ih=im.size
        ratio=iw/ih
        if ratio>(mw/mh): w=mw; h=mw/ratio
        else: h=mh; w=mh*ratio
        cx=x+(mw-w)/2; cy=y+(mh-h)/2
        s.shapes.add_picture(path,Inches(cx),Inches(cy),Inches(w),Inches(h))
        if cap: txt(s,x,y+mh+0.02,mw,0.26,cap,sz=8,c=C_MUTED,it=True,al=PP_ALIGN.CENTER)
    except Exception as e: print(f"  img err {path}: {e}")
def ftr(s,n):
    bx=tb(s,12.80,7.16,0.45,0.24); tf=bx.text_frame; p=tf.paragraphs[0]
    p.text=f"{n}/{TOTAL}"; p.alignment=PP_ALIGN.RIGHT; r=p.runs[0]
    r.font.size=Pt(8); r.font.color.rgb=C_MUTED
def frieze(s,act):
    cw=SLIDE_W/8
    for i,lb in enumerate(CHAP_SHORT):
        x=i*cw; col=CHAP_COLORS[i] if i==act else C_NAVY
        shp=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x),Inches(0),Inches(cw),Inches(0.34))
        shp.fill.solid(); shp.fill.fore_color.rgb=col; shp.line.fill.background()
        bx=tb(s,x+0.02,0.0,cw-0.04,0.34); tf=bx.text_frame
        tf.margin_left=tf.margin_top=tf.margin_right=tf.margin_bottom=Inches(0.01)
        p=tf.paragraphs[0]; p.text=lb; p.alignment=PP_ALIGN.CENTER; r=p.runs[0]
        r.font.size=Pt(7.5); r.font.bold=(i==act); r.font.color.rgb=C_WHITE
def ptitle(s,t,sub=None):
    if sub: txt(s,0.25,0.38,12.80,0.22,sub,sz=9,b=True,c=C_ACCENT)
    txt(s,0.25,0.58,12.80,0.56,t,sz=20,b=True,c=C_DARK)
    sep=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(0.25),Inches(1.17),Inches(12.80),Inches(0.025))
    sep.fill.solid(); sep.fill.fore_color.rgb=C_ACCENT; sep.line.fill.background()
def blist(s,x,y,w,h,items,sz=10.5):
    bx=tb(s,x,y,w,h); tf=bx.text_frame; tf.word_wrap=True
    tf.margin_left=Inches(0.05); tf.margin_top=Inches(0.05); first=True
    for item in items:
        p=(tf.paragraphs[0] if first else tf.add_paragraph()); first=False; p.space_before=Pt(5)
        if isinstance(item,tuple):
            t1,t2=item; r1=p.add_run(); r1.text=f"\u25cf  {t1}"
            r1.font.size=Pt(sz); r1.font.bold=True; r1.font.color.rgb=C_DARK
            if t2:
                p2=tf.add_paragraph(); p2.space_before=Pt(1); r2=p2.add_run()
                r2.text=f"    {t2}"; r2.font.size=Pt(sz-1.5); r2.font.color.rgb=C_MUTED
        else:
            r=p.add_run(); r.text=f"\u25cf  {item}"; r.font.size=Pt(sz); r.font.color.rgb=C_DARK
def kbar(s,y,kpis):
    n=len(kpis); w=SLIDE_W/n
    for i,(val,lbl) in enumerate(kpis):
        x=i*w; rect(s,x,y,w,0.85,fill=C_NAVY)
        txt(s,x+0.10,y+0.06,w-0.20,0.40,val,sz=22,b=True,c=C_ACCENT,al=PP_ALIGN.CENTER)
        txt(s,x+0.10,y+0.47,w-0.20,0.30,lbl,sz=8.5,c=C_WHITE,al=PP_ALIGN.CENTER)
def trans(lab,n,ci,sn):
    s=prs.slides.add_slide(blank); bg_d(s,CHAP_COLORS[ci])
    txt(s,0.50,2.68,12.33,0.50,f"CHAPITRE {n}",sz=11,b=True,c=C_WHITE,al=PP_ALIGN.CENTER)
    txt(s,0.50,3.16,12.33,1.20,lab,sz=34,b=True,c=C_WHITE,al=PP_ALIGN.CENTER)
    sep=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(5.50),Inches(4.50),Inches(2.33),Inches(0.04))
    sep.fill.solid(); sep.fill.fore_color.rgb=C_WHITE; sep.line.fill.background()
    ftr(s,sn); return s
def c2col(ci,sub,title,items,img=None,icap=None,sn=None,kpis=None):
    s=prs.slides.add_slide(blank); bg_w(s); frieze(s,ci); ptitle(s,title,sub); ftr(s,sn)
    ky=SLIDE_H-0.90 if kpis else None; ch=(ky-1.30-0.05 if kpis else SLIDE_H-1.30-0.05)
    if kpis: kbar(s,ky,kpis)
    if img and os.path.exists(img):
        blist(s,0.25,1.30,6.10,ch,items)
        addimg(s,img,6.55,1.30,6.65,ch-(0.30 if icap else 0),icap)
    else: blist(s,0.25,1.30,12.80,ch,items)
    return s

# ============ SLIDE 1 — TITRE ============
s1=prs.slides.add_slide(blank); bg_w(s1)
if os.path.exists("Logo_INSAStrasbourg.jpg"): s1.shapes.add_picture("Logo_INSAStrasbourg.jpg",Inches(0.30),Inches(0.18),width=Inches(2.60))
addimg(s1,"Geosiapp.jpg",10.60,0.10,2.40,1.00)
card=s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(0.50),Inches(1.50),Inches(12.33),Inches(5.50))
card.fill.solid(); card.fill.fore_color.rgb=RGBColor(245,247,250)
card.line.color.rgb=C_BORDER; card.line.width=Pt(1.0); card.adjustments[0]=0.04
txt(s1,0.80,1.68,11.80,0.30,"PROJET DE FIN D'ETUDES — INSA STRASBOURG",sz=11,b=True,c=C_ACCENT,al=PP_ALIGN.CENTER)
txt(s1,0.80,1.98,11.80,0.25,"Specialite Topographie — Annee 2025-2026",sz=10,c=C_MUTED,al=PP_ALIGN.CENTER)
bxt=tb(s1,0.80,2.30,11.80,1.70); tft=bxt.text_frame; tft.word_wrap=True
tft.margin_left=tft.margin_right=Inches(0.10)
pt=tft.paragraphs[0]; pt.alignment=PP_ALIGN.CENTER
pt.text="Developpement d'un outil permettant le traitement\net l'insertion des archives numeriques sur Geofoncier"
rt=pt.runs[0]; rt.font.size=Pt(24); rt.font.bold=True; rt.font.color.rgb=C_DARK
txt(s1,0.80,4.02,11.80,0.30,"au sein du Cabinet GEO-SIAPP — Aubenas (Ardeche)",sz=12,c=C_ACCENT,al=PP_ALIGN.CENTER)
sep=s1.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(2.0),Inches(4.42),Inches(9.33),Inches(0.025))
sep.fill.solid(); sep.fill.fore_color.rgb=C_BORDER; sep.line.fill.background()
cols=[("CANDIDAT","TRAVAILLE Adrien","Eleve-ingenieur Topographe\nINSA Strasbourg — Promo 2026"),
      ("CORRECTEUR","M. Mathieu KOEHL","Directeur du PFE\nEnseignant-chercheur, INSA / ICube"),
      ("ENTREPRISE","Cabinet GEO-SIAPP","Tuteur : M. Gaetan HAGUE\nGeometre-Expert, Aubenas")]
for i,(role,name,det) in enumerate(cols):
    x=0.80+i*4.00
    txt(s1,x,4.48,3.70,0.22,role,sz=8,b=True,c=C_MUTED,al=PP_ALIGN.CENTER)
    txt(s1,x,4.70,3.70,0.34,name,sz=13,b=True,c=C_NAVY,al=PP_ALIGN.CENTER)
    txt(s1,x,5.06,3.70,0.60,det,sz=9,c=C_MUTED,al=PP_ALIGN.CENTER)
txt(s1,0.80,5.74,11.80,0.22,"Soutenance du 24 septembre 2026",sz=9,c=C_MUTED,al=PP_ALIGN.CENTER,it=True)
print("S1 OK")

# ============ SLIDE 2 — SOMMAIRE ============
s2=prs.slides.add_slide(blank); bg_d(s2)
txt(s2,0.50,0.28,12.33,0.30,"PLAN DE LA PRESENTATION",sz=10,b=True,c=C_ACCENT,al=PP_ALIGN.CENTER)
txt(s2,0.50,0.54,12.33,0.62,"Sommaire",sz=30,b=True,c=C_WHITE,al=PP_ALIGN.CENTER)
labels=[("1","Introduction\n& Contexte"),("2","Analyse\nMetier & Donnees"),("3","Etat\nde l'Art"),
        ("4","Architecture\n& Pipeline"),("5","Fiabilisation\n& Interface"),("6","Integration\nGeofoncier"),
        ("7","Limites\n& Perspectives"),("8","Conclusion")]
tw,th,gap=2.90,2.48,0.16; sx=0.24
for idx,(num,lab) in enumerate(labels):
    row=idx//4; col=idx%4; x=sx+col*(tw+gap); y=1.34+row*(th+gap); color=CHAP_COLORS[idx]
    tile=s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(tw),Inches(th))
    tile.fill.solid(); tile.fill.fore_color.rgb=color
    if color==C_NAVY: tile.line.color.rgb=C_ACCENT; tile.line.width=Pt(2.0)
    else: tile.line.fill.background()
    txt(s2,x+0.14,y+0.14,0.60,0.52,num,sz=24,b=True,c=C_WHITE)
    txt(s2,x+0.14,y+0.66,tw-0.28,1.60,lab,sz=12,b=True,c=C_WHITE)
print("S2 OK")

# ============ SLIDE 3 — PROBLEMATIQUE ============
s3=prs.slides.add_slide(blank); bg_w(s3)
addimg(s3,"img/registre_crop_opt.jpg",0.00,0.00,5.70,7.50)
sep_v=s3.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(5.72),Inches(0.0),Inches(0.03),Inches(7.50))
sep_v.fill.solid(); sep_v.fill.fore_color.rgb=C_ACCENT; sep_v.line.fill.background()
txt(s3,6.00,0.55,7.10,0.28,"PROBLEMATIQUE",sz=10,b=True,c=C_ACCENT)
bq=tb(s3,6.00,0.86,7.10,1.55); tfq=bq.text_frame; tfq.word_wrap=True
pq=tfq.paragraphs[0]; pq.alignment=PP_ALIGN.LEFT
pq.text="Comment automatiser le traitement et le versement\nde 50 ans d'archives foncieres manuscrites sur Geofoncier ?"
rq=pq.runs[0]; rq.font.size=Pt(16.5); rq.font.bold=True; rq.font.color.rgb=C_DARK
sep_h=s3.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(6.00),Inches(2.48),Inches(7.10),Inches(0.03))
sep_h.fill.solid(); sep_h.fill.fore_color.rgb=C_ACCENT; sep_h.line.fill.background()
verrous=[("Registres illisibles","Ecriture manuscrite ancienne, abreviations, encre fanee, scans inclins."),
         ("Noms de communes ambigus","Ex. 'La Chap./A.' n'a aucun sens hors contexte ardechois."),
         ("Versement 100% manuel","15 a 30 min par dossier. GEO-SIAPP : 23 600 dossiers en attente.")]
for i,(t,c) in enumerate(verrous):
    y=2.64+i*1.18
    ns=s3.shapes.add_shape(MSO_SHAPE.OVAL,Inches(6.00),Inches(y),Inches(0.36),Inches(0.36))
    ns.fill.solid(); ns.fill.fore_color.rgb=C_ACCENT; ns.line.fill.background()
    txt(s3,6.00,y,0.36,0.36,str(i+1),sz=12,b=True,c=C_WHITE,al=PP_ALIGN.CENTER)
    txt(s3,6.48,y,6.55,0.30,t,sz=12,b=True,c=C_DARK)
    txt(s3,6.48,y+0.32,6.55,0.70,c,sz=10,c=C_MUTED)
rb=s3.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(6.00),Inches(6.34),Inches(7.10),Inches(0.96))
rb.fill.solid(); rb.fill.fore_color.rgb=C_NAVY; rb.line.fill.background()
txt(s3,6.10,6.44,6.90,0.70,"Reponse : un pipeline IA local — OCR, NER, VLM et API REST — deployable sur PC de bureau Windows.",sz=10.5,b=True,c=C_WHITE)
ftr(s3,3); print("S3 OK")

# ============ SLIDES 4-10 ============
trans("Introduction & Contexte",1,0,4); print("S4 OK")
c2col(0,"Chapitre 1 — Structure d'accueil","Le Cabinet GEO-SIAPP",
    [("Depuis 1992, Aubenas (Ardeche)","4 agences : Aubenas, Pierrelatte, Vallon-Pont-d'Arc, Guilherand-Granges."),
     ("Missions","Foncier, urbanisme, ingenierie VRD, topographie drone et laser 3D."),
     ("Mission confiee : 6 mois","Concevoir un outil automatise d'extraction et de versement des archives sur Geofoncier.")],
    img="img/bureau_cropped.jpg",icap="Siege GEO-SIAPP, Aubenas",sn=5,
    kpis=[("4","AGENCES"),("23 600","DOSSIERS"),("1959-2007","PERIODE")]); print("S5 OK")
c2col(0,"Chapitre 1 — Cadre legal","Le cadre reglementaire de la profession",
    [("Loi du 7 mai 1946","Seul le geometre-expert fixe les limites de propriete. Valeur juridique opposable aux tiers."),
     ("Decret n 96-478 — Art. 55","Obligation de conservation 30 ans minimum, meme apres cessation d'activite."),
     ("Portail Geofoncier (OGE)","Centralise toutes les interventions foncieres. GEODEMAT (DGFiP, 2021) ne couvre pas les archives privees.")],
    img="img/Chap2_pastilles.jpg",icap="Pastilles d'intervention Geofoncier — territoire GEO-SIAPP",sn=6); print("S6 OK")
c2col(0,"Chapitre 1 — Contexte","Pourquoi ce projet ?",
    [("'Bornage sur bornage ne vaut'","Toute recherche de borne necessite de retrouver l'acte ancien. Recherche d'anteriorite obligatoire."),
     ("Une recherche manuelle chronophage","15 a 30 min par dossier. 23 600 dossiers non verses. Certains cabinets facturent la recherche."),
     ("GEODEMAT ne couvre pas les archives privees","La responsabilite de versement reste entierement celle des geometres-experts.")],
    img="img/chapelle_sous_aubenas_crop.png",icap="'La Chap./A.' — ambigu hors contexte ardechois",sn=7); print("S7 OK")
trans("Analyse Metier & Donnees",2,1,8); print("S8 OK")
c2col(1,"Chapitre 2 — Fonds d'archives","Les archives : diversite et contraintes",
    [("50 ans de documents tres heterogenes","Plans sur calque, registres manuscrits, tapuscripts. Pas de standard commun."),
     ("Limites des OCR classiques","Abreviations locales, encre fanee, scans inclins. CER > 30% sur manuscrits anciens avec EasyOCR."),
     ("Traitement 100% local obligatoire","Archives foncieres = donnees personnelles. Aucun envoi vers serveur distant autorise.")],
    img="img/4513_DA_124_crop.png",icap="Registre manuscrit du fonds SIAPP (annees 1970)",sn=9); print("S9 OK")
c2col(1,"Chapitre 2 — API Geofoncier","L'API REST Geofoncier",
    [("5 champs obligatoires","Code INSEE commune, date (ISO 8601), surface m2, nature de l'operation, type d'acte."),
     ("Authentification JWT","Credentials cabinet → token Bearer. POST /dossiers avec payload JSON. HTTP 201 = succes."),
     ("3 types de documents","Plan DAO (DMPC), registre manuscrit, piece annexe. Chacun a sa mise en page propre.")],
    img="img/Etape 3_Versement_geofoncier.jpg",icap="Interface Geofoncier apres versement",sn=10); print("S10 OK")

# ============ SLIDES 11-15 ============
trans("Etat de l'Art",3,2,11); print("S11 OK")
c2col(2,"Chapitre 3 — OCR & HTR","Reconnaissance de texte imprime et manuscrit",
    [("EasyOCR — OCR classique","Rapide, multilingue, bon sur texte imprime. CER > 30% sur ecriture ancienne. Retenu pour en-tetes."),
     ("TrOCR (Microsoft, 2021)","Transformer encoder-decoder pre-entraine (IAM, IIIT5K). Tres bon sur manuscrits historiques."),
     ("HTR-United","Depot collaboratif de modeles HTR. Dataset de reference pour ecriture francaise ancienne.")],
    img="img/trocr_architecture_figure1.png",icap="Architecture TrOCR (Li et al., 2021)",sn=12); print("S12 OK")
c2col(2,"Chapitre 3 — NER & Extraction","Extraction d'entites : GLiNER vs LayoutLM",
    [("GLiNER (Zaratiana et al., 2023)","NER zero-shot, entraine sur PILE (825 Go). GitHub: urchade/GLiNER. Apache 2.0. Aucun fine-tuning requis."),
     ("LayoutLM (Microsoft, 2020)","Transformer multimodal (texte + coordonnees spatiales). Performant sur formulaires structures."),
     ("Choix retenu : GLiNER","Superieur sur texte libre manuscrit. LayoutLM inadapte aux registres non tabulaires du fonds SIAPP.")],
    img="img/gliner_prompt_render.png",icap="GLiNER : extraction zero-shot par types definis a l'inference",sn=13); print("S13 OK")
c2col(2,"Chapitre 3 — VLM & Arbitrage","Vision Language Models en arbitre local",
    [("Ollama — inference locale GPU","Execute LLaVA, Qwen2-VL, Phi-3-Vision sur GPU local (8 Go VRAM). Licence MIT. Zero cloud."),
     ("Role : arbitre de dernier recours","Si score GLiNER < 0.65 sur commune ou date, le VLM est interroge sur une capture de la zone."),
     ("Exemple : La Chapelle-sous-Aubenas","Abreviee 'La Chap./A.' dans les registres. Le VLM retourne le nom complet + code INSEE.")],
    img="img/vlm_arbitrage_chapelle_exact.png",icap="'La Chap./A.' → 'La Chapelle-sous-Aubenas' via VLM",sn=14); print("S14 OK")

# Slide 15 — Tableau comparatif
s15=prs.slides.add_slide(blank); bg_w(s15); frieze(s15,2)
ptitle(s15,"Tableau comparatif — justification des choix","Chapitre 3 — Synthese"); ftr(s15,15)
headers=["Technologie","Type","Force principale","Limite","Statut"]
rows15=[["EasyOCR","OCR","Rapide, multilingue","Cursive > 30% CER","Partiel"],
        ["TrOCR","HTR","Manuscrit historique","Fine-tuning requis","Partiel"],
        ["GLiNER","NER","Zero-shot flexible","Contexte long","Retenu"],
        ["LayoutLM","NER+Layout","Formulaires structures","Inadapte cursive","Rejete"],
        ["VLM (LLaVA)","Multimodal","Ambiguite visuelle","Lent ~8s/page","Arbitre"]]
cw15=[2.20,1.35,3.00,2.90,1.30]; rh=0.52; ty=1.42; tx=0.25
hx=tx
for i,h in enumerate(headers):
    rect(s15,hx,ty,cw15[i],rh,fill=C_NAVY)
    txt(s15,hx+0.06,ty+0.10,cw15[i]-0.12,rh-0.12,h,sz=10,b=True,c=C_WHITE,al=PP_ALIGN.CENTER)
    hx+=cw15[i]
for ri,row in enumerate(rows15):
    ry=ty+(ri+1)*rh; fc=RGBColor(245,247,250) if ri%2==0 else C_WHITE; rx=tx
    for ci,cell in enumerate(row):
        rect(s15,rx,ry,cw15[ci],rh,fill=fc,line=C_BORDER)
        cc=(C_SUCCESS if cell=="Retenu" else (C_RED if cell=="Rejete" else (C_ACCENT if cell in ("Partiel","Arbitre") else C_DARK)))
        bd=cell in ("Retenu","Rejete","Partiel","Arbitre")
        txt(s15,rx+0.06,ry+0.10,cw15[ci]-0.12,rh-0.14,cell,sz=9.5,b=bd,c=cc,al=PP_ALIGN.CENTER); rx+=cw15[ci]
fy=ty+(len(rows15)+1)*rh+0.14
rect(s15,0.25,fy,12.80,0.70,fill=C_NAVY)
txt(s15,0.40,fy+0.12,12.50,0.50,"Pipeline retenu : EasyOCR/TrOCR → GLiNER → Ollama VLM en arbitre de dernier recours.",sz=11,b=True,c=C_WHITE,al=PP_ALIGN.CENTER)
print("S15 OK")

# ============ SLIDES 16-22 ============
trans("Architecture & Pipeline",4,3,16); print("S16 OK")

# Slide 17 — 6 scripts
s17=prs.slides.add_slide(blank); bg_w(s17); frieze(s17,3)
ptitle(s17,"Architecture : 6 scripts Python modulaires","Chapitre 4 — Architecture"); ftr(s17,17)
scripts17=[("main.py","Orchestrateur","Lance la chaine, gere memoire et fichiers temporaires.",C_NAVY),
           ("plan_classifieur.py","Classification","Distingue DAO, DMPC, registre, piece annexe en < 1 seconde.",C_ACCENT),
           ("spatial_extractor.py","Segmentation YOLOv8","Detecte cartouche, tableau parcelles, annotations.",RGBColor(6,95,70)),
           ("outil_ocr.py","OCR + HTR + GLiNER","Lit le texte (EasyOCR + TrOCR) et extrait les 5 champs cibles.",RGBColor(154,52,18)),
           ("coherence_controle.py","Controle qualite","17 regles metier, Levenshtein INSEE, validation ISO 8601.",RGBColor(88,28,135)),
           ("geofoncier_api.py","IHM & API","Interface Streamlit de relecture + payload JSON REST vers Geofoncier.",RGBColor(15,118,110))]
cw17=(SLIDE_W-0.50)/6
for i,(nm,rl,ds,col) in enumerate(scripts17):
    x=0.25+i*cw17
    c=s17.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.38),Inches(cw17-0.08),Inches(4.80))
    c.fill.solid(); c.fill.fore_color.rgb=C_LIGHT; c.line.color.rgb=col; c.line.width=Pt(2.0)
    txt(s17,x+0.06,1.48,cw17-0.14,0.28,str(i+1),sz=18,b=True,c=col)
    txt(s17,x+0.06,1.80,cw17-0.14,0.28,nm,sz=8.5,b=True,c=col)
    txt(s17,x+0.06,2.14,cw17-0.14,0.28,rl,sz=9,b=True,c=C_DARK)
    txt(s17,x+0.06,2.48,cw17-0.14,3.40,ds,sz=8.5,c=C_MUTED)
rect(s17,0.25,6.26,12.80,0.56,fill=RGBColor(241,245,249))
txt(s17,0.35,6.33,12.60,0.42,"Scan brut  →  Classification  →  Segmentation YOLOv8  →  OCR+NER  →  Controle qualite  →  Versement Geofoncier",sz=9.5,b=True,c=C_NAVY,al=PP_ALIGN.CENTER)
print("S17 OK")

c2col(3,"Chapitre 4 — Pretraitement","Pretraitement de l'image",
    [("1. Verification resolution (seuil 300 DPI)","En-dessous de 300 DPI, signal a l'operateur. Le pipeline continue mais alertes activees."),
     ("2. Redressement par transformee de Hough","Detection des lignes du tableau pour corriger l'angle d'inclinaison du scan."),
     ("3. Binarisation d'Otsu","Seuillage automatique, robuste aux variations d'encre et d'eclairage. -18 pts de CER sur corpus test.")],
    img="img/4513_DA_124_centered_crop.png",icap="Apres redressement Hough + binarisation Otsu",sn=18); print("S18 OK")
c2col(3,"Chapitre 4 — Detection & NER","YOLOv8 et GLiNER : detection et extraction",
    [("YOLOv8n — 3 classes, mAP@0.5 = 0.87","Cartouche, tableau parcelles, annotations. ~20ms/image GPU. 400 images annotees du fonds SIAPP."),
     ("GLiNER — 5 entites zero-shot","Commune (INSEE), date (ISO 8601), surface m2, nature, type d'acte. Score [0,1] par entite."),
     ("Arbitrage VLM si score < 0.65","Sur commune ou date ambigues : capture zone → Ollama VLM → nom complet + code INSEE.")],
    img="img/composite_yolo_extraction.png",icap="Zones detectees YOLOv8 + entites GLiNER",sn=19); print("S19 OK")
trans("Fiabilisation & Interface",5,4,20); print("S20 OK")
c2col(4,"Chapitre 5 — Interface Streamlit","Interface de validation et controle qualite",
    [("Formulaire de relecture assistee","Affichage cote-a-cote scan / formulaire pre-rempli. Validation en 1-2 minutes."),
     ("Code couleur de confiance","Vert > 0.85 (fiable), orange 0.65-0.85 (a verifier), rouge < 0.65 (saisie manuelle)."),
     ("Carte Folium interactive","Chaque dossier valide sur OpenStreetMap. Clic : reference, date, surface, lien vers scan.")],
    img="img/interface_haut.jpg",icap="Interface de relecture Streamlit",sn=21); print("S21 OK")

# Slide 22 — Video
s22=prs.slides.add_slide(blank); bg_d(s22,RGBColor(5,15,30))
txt(s22,0.50,0.16,12.33,0.42,"DEMONSTRATION — Pipeline complet sur des registres reels GEO-SIAPP",sz=13,b=True,c=C_ACCENT,al=PP_ALIGN.CENTER)
vid="Demo_soutenance_final.mp4"
if not os.path.exists(vid): vid="Demonstration_PFE_Toutes_Videos.mp4"
vw=11.80; vh=6.38; vl=(SLIDE_W-vw)/2; vt=0.70
if os.path.exists(vid):
    try:
        s22.shapes.add_movie(vid,Inches(vl),Inches(vt),Inches(vw),Inches(vh),poster_frame_image=None,mime_type="video/mp4")
        print(f"  Video: {vid}")
    except Exception as e:
        print(f"  Video err: {e}"); rect(s22,vl,vt,vw,vh,fill=RGBColor(20,30,50))
        txt(s22,vl+0.5,vt+vh/2-0.2,vw-1.0,0.40,f"[VIDEO] {vid}",sz=14,c=C_MUTED,al=PP_ALIGN.CENTER)
else:
    rect(s22,vl,vt,vw,vh,fill=RGBColor(20,30,50))
    txt(s22,vl+0.5,vt+vh/2-0.3,vw-1.0,0.60,f"INSERER : {vid}",sz=14,c=C_MUTED,al=PP_ALIGN.CENTER)
ftr(s22,22); print("S22 OK")

# ============ SLIDES 23-30 ============
trans("Integration Geofoncier",6,5,23); print("S23 OK")
c2col(5,"Chapitre 6 — API Geofoncier","Versement via l'API REST Geofoncier",
    [("Authentification JWT","Credentials → token Bearer. POST /dossiers avec payload JSON (5 champs + scan). HTTP 201 = succes."),
     ("Versement par lot","main.py orchestre plusieurs dossiers en serie. Log CSV genere a chaque session."),
     ("Detection de doublon","GET /dossiers?ref=... avant tout versement. Reference existante = dossier marque doublon, ignore.")],
    img="img/6_succes_pastille.jpg",icap="Pastille Geofoncier creee apres versement reussi",sn=24); print("S24 OK")

# Slide 25 — Resultats Prades
s25=prs.slides.add_slide(blank); bg_w(s25); frieze(s25,5)
ptitle(s25,"Validation : commune de Prades (Ardeche)","Chapitre 6 — Resultats"); ftr(s25,25)
blist(s25,0.25,1.38,6.00,4.20,
    [("50 dossiers 1970-1990 — test grandeur reelle","Conditions operationnelles reelles, validation humaine a chaque etape."),
     ("47/50 dossiers verses avec succes","3 echecs : 2 scans < 200 DPI, 1 doublon existant dans Geofoncier."),
     ("Validation par M. Hague (GEO-SIAPP)","Coherence des metadonnees verifiee pour chacun des 47 dossiers."),
     ("Gain estime : 4h de saisie evitee","Pour 50 dossiers. Soit 5 min vs 30 min/dossier en saisie manuelle.")])
txt(s25,6.30,1.38,6.85,0.28,"SCORES F1 PAR CHAMP OBLIGATOIRE",sz=9,b=True,c=C_ACCENT)
fields25=[("Commune (code INSEE)",0.94,C_SUCCESS),("Date de l'acte",0.91,C_SUCCESS),
          ("Surface (m2)",0.88,C_ACCENT),("Nature operation",0.85,C_ACCENT),
          ("Type d'acte",0.82,C_WARN),("Score global",0.88,C_SUCCESS)]
by=1.74
for field,score,col in fields25:
    txt(s25,6.30,by,3.20,0.38,field,sz=9.5,c=C_DARK)
    rect(s25,9.58,by+0.06,2.60,0.26,fill=C_BORDER)
    rect(s25,9.58,by+0.06,2.60*score,0.26,fill=col)
    txt(s25,12.25,by,0.95,0.38,f"F1={score:.2f}",sz=9.5,b=True,c=col); by+=0.52
addimg(s25,"img/6.4_match_confirm.jpg",6.30,5.10,6.85,2.26); print("S25 OK")

trans("Limites & Perspectives",7,6,26); print("S26 OK")
c2col(6,"Chapitre 7 — Limites","Limites actuelles et perspectives",
    [("Limite — qualite des scans","En-dessous de 200 DPI, le pipeline echoue. Numerisation physique indispensable."),
     ("Limite — ecriture tres degradee","Registres pre-1960 parfois illisibles. Fine-tuning de TrOCR sur corpus SIAPP ameliorerait les resultats."),
     ("Perspective — extension multi-cabinet","Architecture modulaire : adaptable a d'autres fonds avec un module de configuration."),
     ("Perspective — auto-validation","Si F1 > 0.95 sur tous les champs, validation humaine optionnelle. Sous reserve accord OGE.")],
    sn=27); print("S27 OK")

trans("Conclusion",8,7,28); print("S28 OK")

# Slide 29 — Bilan
s29=prs.slides.add_slide(blank); bg_w(s29); frieze(s29,7)
ptitle(s29,"Bilan du projet","Chapitre 8 — Conclusion"); ftr(s29,29)
blist(s29,0.25,1.38,9.10,5.00,
    [("Un pipeline IA deployable localement","OCR + HTR + GLiNER + VLM en arbitre. 6 scripts Python 3.10. Environnement Windows, zero cloud."),
     ("Des resultats concrets sur un cas reel","47/50 dossiers verses. F1 global = 0.88. Gain de temps x6 par dossier. Valide par GEO-SIAPP."),
     ("Une architecture reproductible","Modulaire, open-source, adaptable. Conforme OGE : l'IA est en appui, le geometre reste decideur.")],
    sz=12)
kbar(s29,6.55,[("F1 = 0.88","Score global"),("94%","Communes correctes"),("x6","Gain de temps"),("23 600","Dossiers restants")])
print("S29 OK")

# Slide 30 — Remerciements
s30=prs.slides.add_slide(blank); bg_d(s30)
txt(s30,1.0,1.90,11.33,0.72,"Merci pour votre attention.",sz=36,b=True,c=C_WHITE,al=PP_ALIGN.CENTER)
txt(s30,1.0,2.68,11.33,0.44,"Je reste disponible pour vos questions.",sz=16,c=C_ACCENT,al=PP_ALIGN.CENTER)
sep30=s30.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(4.50),Inches(3.22),Inches(4.33),Inches(0.04))
sep30.fill.solid(); sep30.fill.fore_color.rgb=C_ACCENT; sep30.line.fill.background()
thanks=["M. Gaetan HAGUE — Tuteur entreprise, Cabinet GEO-SIAPP",
        "M. Mathieu KOEHL — Directeur du PFE, INSA Strasbourg / ICube (TRIO)",
        "L'equipe GEO-SIAPP d'Aubenas pour leur accueil et disponibilite"]
for i,t in enumerate(thanks): txt(s30,1.0,3.42+i*0.48,11.33,0.44,t,sz=11.5,c=C_WHITE,al=PP_ALIGN.CENTER)
if os.path.exists("Logo_INSAStrasbourg.jpg"): s30.shapes.add_picture("Logo_INSAStrasbourg.jpg",Inches(3.30),Inches(5.65),width=Inches(2.20))
addimg(s30,"Geosiapp.jpg",7.80,5.65,2.30,0.92)
ftr(s30,30); print("S30 OK")

# ============ SAVE ============
out="Soutenance_PFE_Adrien_TRAVAILLE.pptx"
prs.save(out)
print(f"\nSucces ! {out} — {len(prs.slides)} slides")
