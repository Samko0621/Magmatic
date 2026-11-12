import re, json, sys
from pypdf import PdfReader
from playwright.sync_api import sync_playwright

def chip(t,l): return f'<span class="tag {l}">{ {"c":"Confirmé","d":"Déduit","a":"À compléter"}[l] }</span>'
def todo(t): return f'<div class="todo"><b>À compléter</b> — {t}</div>'

def tree():
    rows=[("010270","Charpente",2,[("609509","Cadre réa lev",4),("609458","U support crémaillères",20),("609457","Cadre inf. crémaillère",2)]),
    ("002531","Équipement cage 1",2,[("001972","S/ens REA",4),("002886","S/ens contrepoids",4),("705719","Contrepoids béton 188 kg",3),("705834","Contrepoids béton 84 kg",1)]),
    ("001089","Crémaillères",2,[("001217","S/ens crémaillères",4),("000284","Tendeur de crémaillère",4),("002523","Chape supérieure",4),("002154","Accessoire support",10),("002793","Partie fixe",1),("002025","Butée escamotable",10),("002719","Fixation cadre inf. (chevilles chimiques)",1)]),
    ("080143","Levmatic",2,[("010358","Nacelle AS12",1),("002534","Chaîne porte-câbles",1),("002545","Capteur à monter sur site",1)]),
    ("002532","Équipement cage 2",2,[("002784","Butée méca. haute",4),("002806","Butée méca. basse",4),("000951","Came hors zone",1),("002324","Embase câble de guidage",4),("002903","Accrochage câbles",8),("002807","Élingue de guidage",1)]),
    ("002533","Équipement cage 3",2,[("002796","Accrochage câble CP",4),("002808","Élingue de CP",4)])]
    h='<div class="tree">'
    for c,n,q,ch in rows:
        h+=f'<div class="trow"><div class="tp"><small>{c}</small><b>{n}</b><i>× {q}</i></div><div class="tc">'
        for cc,cn,cq in ch: h+=f'<div class="ch"><small>{cc}</small>{cn}<i>× {cq}</i></div>'
        h+='</div></div>'
    return h+'</div>'

H=[]  # headings for toc: (level,text)
def sec(num,title):
    H.append((1,f"{num}. {title}")); return f'<div class="banner" id="s{num}">{num}. {title}</div>'
def sub(txt, cls=""):
    return f'<h2 class="{cls}">{txt}</h2>'

CSS="""
@page{size:A4;margin:0}
*{box-sizing:border-box}
body{font-family:Arial,Helvetica,sans-serif;color:#333;font-size:10.2pt;line-height:1.38;margin:0}
.page{width:210mm;height:297mm;padding:30mm 16mm 24mm 16mm;page-break-after:always;position:relative;overflow:hidden}
.page:last-child{page-break-after:auto}
.top{position:absolute;top:0;left:70mm;right:16mm;height:5mm;background:#4f81bd}
.bot{position:absolute;bottom:0;left:70mm;right:16mm;height:6mm;background:#c0504d}
.foot{position:absolute;bottom:9mm;left:16mm;right:16mm;font-size:8pt;color:#777;display:flex;justify-content:space-between}
.banner{background:#6d6d6d;color:#fff;font-weight:bold;text-transform:uppercase;font-size:17pt;padding:3mm 5mm;margin:0 -16mm 6mm -16mm;padding-left:16mm;letter-spacing:.2px}
h2{color:#d9232d;font-size:14pt;margin:5mm 0 2mm 0;font-weight:bold}
h3{color:#555;font-size:11pt;margin:4mm 0 1.5mm 0}
p{margin:0 0 2.5mm 0}
table{border-collapse:collapse;width:100%;margin:2mm 0 4mm 0;font-size:9.2pt}
th{background:#6d6d6d;color:#fff;text-align:left;padding:1.6mm 2.4mm;font-weight:bold}
td{padding:1.5mm 2.4mm;border-bottom:1px solid #bbb;vertical-align:top}
tr:nth-child(even) td{background:#f3f3f3}
.tag{display:inline-block;font-size:7.4pt;font-weight:bold;padding:.4mm 1.8mm;border-radius:2mm;color:#fff;white-space:nowrap}
.tag.c{background:#2e7d32}.tag.d{background:#e08a00}.tag.a{background:#fff;color:#c62828;border:1px solid #c62828}
.todo{border:1.4px dashed #c62828;background:#fff5f5;color:#8a1c1c;padding:2mm 3mm;margin:2mm 0 4mm 0;font-size:9.2pt}
.warn{color:#d9232d;font-weight:bold;font-size:12.5pt;text-align:center;margin-top:16mm;line-height:1.35}
.toc{margin:6mm 0 0 55mm}
.toc div{display:flex;justify-content:space-between;color:#555;font-size:11.5pt;margin:0 0 3.6mm 0;text-decoration:underline}
.toc div.l2{font-size:9.8pt;margin:0 0 1.6mm 6mm;text-decoration:none;color:#666}
.cover h1{color:#e31e24;font-size:38pt;margin:55mm 0 0 62mm;line-height:1.05}
.cover .s1{color:#e31e24;font-size:22pt;font-style:italic;margin:6mm 0 0 62mm}
.cover .s2{color:#e31e24;font-size:21pt;font-weight:bold;margin:20mm 0 0 62mm;line-height:1.2}
.cover .ref{color:#666;font-size:21pt;font-weight:bold;margin:22mm 0 0 62mm}
.cover .draw{margin:22mm 0 0 8mm}
.tree{margin:2mm 0 3mm 0}
.trow{display:flex;gap:3mm;margin-bottom:2.2mm;align-items:stretch}
.tp{width:34mm;flex:none;background:#2f4a6d;color:#fff;padding:1.6mm 2.2mm;font-size:8.6pt;display:flex;flex-direction:column;justify-content:center}
.tp small{opacity:.8;font-size:7.4pt}.tp i{font-size:8pt}
.tc{display:flex;flex-wrap:wrap;gap:1.4mm;flex:1}
.ch{background:#dff1f7;border:1px solid #7fb8cf;padding:1mm 1.8mm;font-size:7.8pt;line-height:1.25;min-width:24mm}
.ch small{display:block;color:#33607a;font-size:7pt}.ch i{display:block;font-size:7.4pt;color:#333}
ul{margin:0 0 3mm 4mm;padding-left:4mm}li{margin-bottom:1mm}
.legend{font-size:8.6pt;color:#555;margin:0 0 3mm 0}
.k{font-family:Consolas,monospace;background:#eee;padding:0 1mm}
"""

def page(body,foot=True,cls=""):
    f='<div class="foot"><span>Magmatic — Principes généraux et règles de conception — Ind A (brouillon)</span><span class="pn"></span></div>' if foot else ''
    return f'<div class="page {cls}"><div class="top"></div><div class="bot"></div>{body}{f}</div>'

P=[]
# cover
P.append(page('''<h1>MAGMATIC</h1><div class="s1">Stockage automatisé de palettes</div>
<div class="s2">Principes généraux et règles de conception</div>
<div class="ref">Réf. à attribuer (format BEEM)&nbsp; Ind A</div>
<div class="draw"></div>''',False,"cover"))
# toc placeholder
P.append("@@TOC@@")
# historique
P.append(page(sec(1,"Historique des modifications")+'''
<table><tr><th style="width:16mm">Version</th><th style="width:26mm">Date</th><th style="width:36mm">Auteur</th><th>Modifications</th></tr>
<tr><td><b>A</b></td><td>10/11/2026</td><td>MIHCI S.</td><td>Création du document (brouillon V1) : données d'entrée et nomenclature de lancement d'une cage Levmatic, règles déduites de l'historique des affaires P&amp;G, nommage BEEM, structure cible calquée sur le guide d'implantation Intelis Convey.</td></tr></table>
<p class="legend"><b>Niveau de confiance des règles :</b> '''+chip('','c')+' visible dans une source écrite &nbsp; '+chip('','d')+' tendance relevée dans l\'historique des affaires, à valider par le tuteur &nbsp; '+chip('','a')+''' information absente des sources disponibles.</p>
<h2>Sources utilisées</h2>
<table><tr><th>Source</th><th>Apport</th></tr>
<tr><td>Classeur de lancement EHP01718LEVM001 (cage Levmatic NEXT EXPANSION, Magmatic 2A)</td><td>Données d'entrée, contrepoids, nomenclature de lancement</td></tr>
<tr><td>Bilan affaire P&amp;G (120 affaires, 5 sites) et ses onglets « Moteurs Conv Lev » et « Suivi voies de roulement »</td><td>Historique des évolutions, fournisseurs de convoyeurs, voies de roulement</td></tr>
<tr><td>Récapitulatif clients Magmatic (29 clients)</td><td>Types 2A / 2C, hauteurs, nombre de niveaux</td></tr>
<tr><td>Nommage des documents BEEM</td><td>Codification des plans et documents</td></tr>
<tr><td>Plans d'implantation AutoCAD (site type)</td><td>Lecture d'une implantation</td></tr>
<tr><td>Guide Intelis Convey ST1100-10-00009 Ind M</td><td>Modèle de structure uniquement : ses valeurs ne s'appliquent pas à Magmatic</td></tr></table>
<div class="warn">Les cas non décrits dans ce document, ou qui ne respectent pas ces règles, seront soumis à la validation du bureau d'études.</div>'''))
# données générales
P.append(page(sec(2,"Données générales et charges admissibles")+'''
<h2>Données d'entrée d'une cage Levmatic</h2>
<p>Valeurs de la cage <b>NEXT EXPANSION</b> (affaire EHP01718, code FLEVE, Magmatic 2A). Elles servent d'exemple de saisie ; une autre cage aura d'autres valeurs.</p>
<table><tr><th>Donnée</th><th>Valeur</th><th>Niveau</th></tr>
<tr><td>Type de Magmatic</td><td>2A</td><td>'''+chip('','c')+'''</td></tr>
<tr><td>Ambiance</td><td>Chaud (table du classeur : chaud 2×25, froid 2×50)</td><td>'''+chip('','c')+'''</td></tr>
<tr><td>Débordement de charge, sens 1200</td><td>2 × 100 mm</td><td>'''+chip('','c')+'''</td></tr>
<tr><td>Poids Levmatic</td><td>1 960 kg</td><td>'''+chip('','c')+'''</td></tr>
<tr><td>Poids VM</td><td>1 550 kg</td><td>'''+chip('','c')+'''</td></tr>
<tr><td>Poids maxi palette</td><td>1 000 kg</td><td>'''+chip('','c')+'''</td></tr>
<tr><td>Nombre d'étages</td><td>10</td><td>'''+chip('','c')+'''</td></tr>
<tr><td>Nombre de niveaux de convoyeurs</td><td>2</td><td>'''+chip('','c')+'''</td></tr>
<tr><td>Nombre de Levmatic / cages optionnelles</td><td>2 / 0</td><td>'''+chip('','c')+'''</td></tr></table>
<h2>Contrepoids</h2>
<table><tr><th>Donnée</th><th>Valeur</th><th>Niveau</th></tr>
<tr><td>Contrepoids béton</td><td>4 CP, poids par CP 808,75 kg, 3 lests standard de 188 kg</td><td>'''+chip('','c')+'''</td></tr>
<tr><td>Références de lest béton</td><td>705719 (650×650×200, 188 kg) ; 705833 (650×650×110, 103,5 kg) ; 705834 (84,5 kg)</td><td>'''+chip('','c')+'''</td></tr>
<tr><td>Contrepoids acier</td><td>4 CP, cadre 143 kg, 86 lests standard (344 pour 4 CP)</td><td>'''+chip('','c')+'''</td></tr>
<tr><td>Déséquilibre après choix des lests</td><td>9,75 kg (reste à équilibrer : 216,75 kg)</td><td>'''+chip('','c')+'''</td></tr></table>
'''+todo("règle de choix entre contrepoids béton et acier, déséquilibre admissible, effet de la hauteur de cage. Source : tuteur ou notice de calcul.")+
'''<h2>Types de palettes et charges acceptées</h2>'''+todo("formats de palette admis (dimensions mini et maxi, hauteur, matériau), tolérances de déformation. Aucun document des sources ne les donne ; le guide Intelis a un chapitre équivalent pour les colis et les bacs.")))
# FAQ
P.append(page(sec(3,"Règles générales : FAQ")+'''
<h2>Comment lire une implantation Magmatic ?</h2>
<ul><li>Le plan « First floor » examiné montre un rack dont les allées sont numérotées de 1 à 54, des véhicules dans les allées, des convoyeurs d'entrée et de sortie le long du rack (flux « B » et « C ») et le bâtiment du client en arrière-plan. '''+chip('','c')+'''</li>
<li>Les plans sont organisés par niveaux : Ground floor, First floor, Second floor, Third floor, plus des coupes (« Magmatic platforms », « Opening in Z axis »). '''+chip('','c')+'''</li>
<li>Un système de navettes « X-PTS » apparaît sur le plan du rez-de-chaussée. '''+chip('','c')+'''</li></ul>
'''+todo("signification des couleurs et calques, cartouche, règle de numérotation des Levmatic (chiffres) et des VM, sens des allées, repère d'origine. À décrire avec le tuteur à partir d'un plan réel.")+
'''<h2>Comment identifier un Levmatic ou un VM sur un site ?</h2>'''+todo("numérotation (ex. Lev 1 à 14 à Londres, Lev A à K à Amiens : deux conventions coexistent selon le site), correspondance avec les plans et le parc VM du Bilan P&amp;G.")+
'''<h2>Où retrouver les plans d'une affaire ?</h2>
<p>Sur le Bilan P&amp;G, 105 affaires sur 120 n'ont pas de plan d'ensemble renseigné et 46 n'ont pas de plan de sous-ensemble. Le fichier « Historique des affaires Magmatic » sert de point d'entrée : une ligne par affaire, avec le n° d'origine, le n° BEEM et l'emplacement des plans. '''+chip('','c')))
# règles de conception
P.append(page(sec(4,"Règles de conception")+'''
<h2>4.1 Cage Levmatic : nomenclature de lancement</h2>
<p>Le classeur de lancement calcule les quantités de chaque ensemble à partir des données d'entrée (section 2). Exemple pour <b>2 Levmatic</b>, 10 étages, 2 niveaux de convoyeurs :</p>'''+tree()+
'''<p class="legend">Chaque ensemble est lancé avec sa quantité calculée ; « × n » est la quantité pour la cage NEXT. '''+chip('','c')+'''</p>
'''))
P.append(page('''<h2>4.1 (suite) Calculs de géométrie de la cage</h2>
<p>Le classeur EHP01718LEVM001 enchaîne cinq onglets de calcul à partir des données d'entrée de la section 2. Valeurs de l'exemple NEXT (chaud, 10 étages, palette 1500 mm). '''+chip('','c')+'''</p>
<h3>Hauteur du cadre de réa</h3>
<table><tr><th>Grandeur</th><th>Règle (classeur)</th><th>NEXT</th></tr>
<tr><td>Fond de fosse → dessous toit</td><td>dernier niveau de pose + 1993 + 1100 + profondeur de fosse</td><td>21 743 mm</td></tr>
<tr><td><b>Fond de fosse → dessous cadre réa</b></td><td>profondeur de fosse + dernier niveau de pose + 1993</td><td><b>20 643 mm</b></td></tr>
<tr><td>Surcourse haute + basse du Lev</td><td>600 + 600</td><td>1 200 mm</td></tr>
<tr><td>Hauteur sous cadre pour VM (batterie en haut)</td><td>VM chaud 1920 / froid 1750 + jeu 50 + surcourse/2 − 165</td><td>HCR pratique 2 470 mm</td></tr></table>
<h3>Longueur des câbles (modèles 612122 et 612123)</h3>
<table><tr><th>Câble</th><th>Règle</th><th>NEXT</th></tr>
<tr><td>Câble de levage</td><td>(H axe réa − H plan roulement − H1 attache − surcourse haute) + π·D/2 + (25 + traverse 300 + jeu cadre/CP) + course (nominale + 2 surcourses) + 900 (boîte à coin), arrondi à la centaine supérieure</td><td><b>21 900 mm</b> (21 872 avant arrondi)</td></tr>
<tr><td>Câble de guidage</td><td>H cadre réa − 35 − 440 + 400 (boîte à coin), arrondi à la centaine supérieure</td><td><b>20 600 mm</b> (20 568 avant arrondi)</td></tr></table>
<p class="legend">Contrôle : le jeu cadre de réa / contrepoids doit être positif (NEXT : −280 mm, le classeur indique « si négatif le CP ne rentre pas dans la cage »). Diamètre primitif de la poulie : 582 mm. '''+chip('','c')+'''</p>
<h3>Chaîne porte-câble (CPC)</h3>
<p>Longueur théorique = course verticale / 2 + π·R + rehausse (200) + longueur supplémentaire (300). Nombre de maillons = longueur théorique / pas (91 mm), arrondi au supérieur. NEXT : course 18 150 mm → 10 800 mm théorique → <b>119 maillons, 10 829 mm</b>. Le classeur vérifie aussi l'interférence de la plaque avec la traverse. '''+chip('','c')+'''</p>
'''))
P.append(page('''<h2>4.1 (fin) Crémaillères et repères automate</h2>
<h3>Lancement des crémaillères</h3>
<table><tr><th>Étape</th><th>Règle</th><th>NEXT</th></tr>
<tr><td>Longueur pratique</td><td>hauteur de cage − 79,5 − 72,5 − 50</td><td>20 441 mm</td></tr>
<tr><td>Crémaillères de 3000 (710241)</td><td>partie entière de longueur pratique / 3000, plus un morceau additionnel pour le reste</td><td>6 + morceau 2020 mm</td></tr>
<tr><td>Pistes</td><td>1 piste 2250 (710238, −60 mm en bas) + pistes de 3000 + morceau additionnel</td><td>1 × 2250 + 5 × 3000 + 2770</td></tr>
<tr><td>Profils de 6000 (440109)</td><td>longueur du profil = longueur crémaillère</td><td>3 profils + morceau 2020 mm</td></tr>
<tr><td>Visserie</td><td>plaques de jonction et goupilles : nb de crémaillères ; vis CHC M12x35 : 9 par crémaillère</td><td>6 plaques, 63 vis, 57 rondelles épaisses</td></tr>
<tr><td>Conditionnement</td><td>troncons de 12 m ; caissons d'extrémité 703450 et central 703449 ; pointes 375000</td><td>2 + 3 caissons, 24 pointes</td></tr></table>
<h3>Repères automate (Registres_automates)</h3>
<p>Pas d'étage 1850 mm, premier niveau palettier 1050 mm, niveaux batterie espacés de 2000 mm (premier 1194,5 mm), niveau convoyeur 1 à 1090 mm. Ces registres se déduisent de la feuille « Niveaux ». '''+chip('','d')+'''</p>
<div class="todo"><b>À valider</b> — Poids Levmatic : 1 500 kg dans « Données Palettier » (« 1500 sur 2A et 1200 ») et 1 960 kg dans « CP et organigramme ». Pas d'étage : 350 mm de plus que la hauteur de palette (1850) alors que le tableau palettier note 1500/1800. À confirmer avec le tuteur.</div>
'''))
P.append(page('''<h2>4.2 Convoyeurs de cage et motoréducteurs</h2>
<p>Le Bilan P&amp;G recense les fournisseurs et motoréducteurs des convoyeurs de Levmatic par affaire. Il n'existe pas de données pour certains fournisseurs, ce qui complique les remplacements. '''+chip('','d')+'''</p>
<table><tr><th>Site</th><th>Affaire</th><th>Levmatic</th><th>Fournisseur</th><th>Remarque</th></tr>
<tr><td>Londres</td><td>PGLO0</td><td>1 ; 2 ; 4 ; 5 ; 7 ; 8</td><td>Schulz</td><td>Pas de données Schulz</td></tr>
<tr><td>Londres</td><td>B6082</td><td>11 ; 12 ; 14</td><td>Matrex</td><td>Pas de données Matrex</td></tr>
<tr><td>Mechelen</td><td>10451</td><td>1 à 6</td><td>Schulz</td><td>Pas de données Schulz</td></tr>
<tr><td>Amiens</td><td>30144</td><td>A, E, K</td><td>Matrex</td><td>Spécification : motoréducteur à frein SEW R37 DT80</td></tr>
<tr><td>Jijona</td><td>24206</td><td>6</td><td>Citex</td><td>Pas de données Citex</td></tr>
<tr><td>Euskirchen</td><td>EHP0000488</td><td>5</td><td>Matrex</td><td>SEW R37DRN80MK4BE1HR, 74 tr/min ; plan Matrex</td></tr></table>
<p><b>Règle proposée :</b> à chaque affaire, archiver le plan et la référence du motoréducteur du convoyeur de cage dans le dossier de l'affaire. '''+chip('','d')+'''</p>
<h2>4.3 Voies de roulement des Levmatic</h2>
<p>Un problème de fissuration a été constaté sur les voies de roulement ; elles sont remplacées par des voies ne créant pas de sur-contraintes. Les pièces d'origine sont remplacées par des sous-ensembles : '''+chip('','c')+'''</p>
<table><tr><th>Pièces d'origine</th><th>Nouveau sous-ensemble</th><th>Sites concernés</th></tr>
<tr><td>711108 – 711110</td><td>005561 (711462 et 711463)</td><td>Londres (Lev 3 et 6)</td></tr>
<tr><td>710892 – 710893</td><td>005563 (711466 et 711467)</td><td>Londres, Mechelen, Amiens, Euskirchen</td></tr>
<tr><td>711591 – 711594</td><td>005560 (711438 et 711439)</td><td>Euskirchen (Lev 13, 14 et 16)</td></tr></table>
<p>Les installations suivent un calendrier par Levmatic (mai 2026 à mai 2027 selon le site) ; Jijona n'a aucun Levmatic remplacé.</p>
'''+todo("cause identifiée de la fissuration, critère de choix entre voie d'origine et voie rapportée (voir B6082 à Londres), conditions de dimensionnement.")))
P.append(page('''<h2>4.4 Véhicules motorisés (VM)</h2>
<ul><li>Poids d'un VM : 1 550 kg (cage NEXT). '''+chip('','c')+'''</li>
<li>Deux types de moteur coexistent : DY (origine) et CMP (après migration, ex. affaires N0212 + E1146 et BJ020 à Londres). '''+chip('','c')+'''</li></ul>
'''+todo("règles de choix du type de moteur, limites de charge, dimensions et gabarit d'un VM. Source possible : parc VM du Bilan P&amp;G (128 lignes), à exploiter avec le tuteur.")+
'''<h2>4.5 Sécurité</h2>
<p>La sécurité des Levmatic est traitée sur chaque site par des affaires dont les désignations se répètent : '''+chip('','d')+'''</p>
<table><tr><th>Sujet</th><th>Affaires relevées</th></tr>
<tr><td>Porte voie X (portes coulissantes)</td><td>Londres BE050 ; Mechelen BE022 (proto, Lev 1) puis BE048 (Lev 2 à 6) ; Amiens BE044 (11 Lev) ; Jijona BE046 (Lev 1 à 4) ; Euskirchen BE042 (7 Lev)</td></tr>
<tr><td>Amélioration sécurité, phase 1</td><td>Équipement de palette cage : Londres BC070, Mechelen BC084, Amiens BC072, Jijona BC082, Euskirchen BC068</td></tr>
<tr><td>Amélioration sécurité, phase 2 (« Bucket 2 »)</td><td>Détection de déformation de rail : BE012, BE010, BE016, BE014, BE008</td></tr>
<tr><td>Traitement des obsolescences (« Bucket 3 »)</td><td>BH006, BH002, BH004, BH010, BH008</td></tr></table>
'''+todo("règles de sécurité à intégrer dès la conception d'une nouvelle cage (portes, détection, butées, protections), pour ne plus les ajouter après coup. À formuler avec le tuteur.")+
'''<h2>4.6 Chargeurs de batteries et lecteurs de balises</h2>
<ul><li>Chargeurs : migration avec support de chargeur Enersys (Londres BE032, Mechelen BD028, Amiens BE038, Jijona BE026, Euskirchen BE024 et BF002) ; ajout de 2 chargeurs à Mechelen (BG028). '''+chip('','d')+'''</li>
<li>Lecteur de balises : antenne Nedap (Mechelen BG022), migrations BG026 et BG030. '''+chip('','d')+'''</li></ul>'''))
# nommage + implantation
P.append(page(sec(5,"Implantation et nommage des documents")+'''
<h2>Nommage des documents (BEEM)</h2>
<p>Format : <span class="k">N° d'affaire – Service – Type – Famille – Produit – Incrément – Langue – Désignation libre – Indice</span>, par exemple <span class="k">NPJ0000001 – M – IM – INT – CON – 001 – FR – Désignation libre – 01</span>. '''+chip('','c')+'''</p>
<table><tr><th style="width:42mm">Champ</th><th>Valeur utile pour Magmatic</th></tr>
<tr><td>Famille de produit</td><td>MAG (Magmatic)</td></tr>
<tr><td>Produits</td><td>ALV armoire élévateur Levmatic ; CON convoyeurs ; VMA véhicule Magmatic ; PCA palette cage ; PCF porte coupe-feu ; PGR protection grillagée ; RAC rack ; STR structure ; SEC sécurité ; HAC chargeur ; BAT batterie VM</td></tr>
<tr><td>Types de document</td><td>IG plan d'implantation général ; IM plan d'implantation ; EN implantation point énergie ; PS plans spécifiques ; NO nomenclature ; ZO zonage ; CG bilan des charges</td></tr>
<tr><td>Service</td><td>A BEA ; B BEEM ; E BEE ; M BMI : à confirmer pour le BE mécanique</td></tr></table>
<h2>Numéros d'affaire</h2>
<p>Sur les 120 affaires du Bilan P&amp;G, 25 seulement sont au format BEEM (3 lettres et 7 chiffres). Les autres suivent des formats plus anciens (ex. B5068, BE012, 30144). '''+chip('','c')+'''</p>
<p><b>Règle proposée :</b> toute nouvelle affaire Magmatic est numérotée au format BEEM ; l'ancien n° reste noté dans l'historique des affaires pour retrouver les plans. '''+chip('','d')+'''</p>
<h2>Implantations standards</h2>'''+todo("implantations types Magmatic 2A et 2C, ambiance chaude ou froide, positions des convoyeurs d'entrée et de sortie, règles de distance (rack, bâtiment, convoyeurs). À écrire à partir de plans de sites réels.")))
P.append(page(sec(6,"Adaptation aux anciens sites")+'''
<p>Les sites P&amp;G montrent les mêmes opérations d'adaptation sur du matériel de générations différentes : '''+chip('','d')+'''</p>
<table><tr><th>Opération</th><th>Contenu relevé</th><th>Exemples</th></tr>
<tr><td>Rétrofit Levmatic</td><td>Changement des moteurs et variateurs, configuration SEW, armoire, partie fixe, capteurs, nacelle, convoyeurs</td><td>E3750, E3754, E3748 ; EHP0000488, EHP0000499, EHP0000527</td></tr>
<tr><td>Rétrofit VM</td><td>Changement de moteur et platine ; remplacement par des VM neufs</td><td>E3756, E3758, E3760 ; EHP0000512, EHP0000528</td></tr>
<tr><td>Remplacement des nacelles</td><td>Nacelles de plusieurs Levmatic par site</td><td>ASO0000080 (Lev 2, 5, 7), ASO0000078 (C, D, F), ASO0000079 (4, 6, 8, 9)</td></tr>
<tr><td>Relais de convoyeurs</td><td>Remplacement des relais par des variateurs, chaînes porte-câbles</td><td>ASO0000071 à ASO0000075</td></tr>
<tr><td>Sécurisation de la butée d'entrée de palettes</td><td>Ajout d'un capteur sur la butée</td><td>EHP0000691, EHP0000871, EHP0000987</td></tr></table>
'''+todo("compatibilité entre pièces anciennes et nouvelles (équivalences de références, interfaces), règles de conservation des anciennes pièces au catalogue. À formuler avec le tuteur."))+
'''''')
P.append(page('<div class="banner" id="s7">Annexe : points à valider avec le tuteur</div>'+'''
<table><tr><th style="width:8mm">N°</th><th>Point à valider</th><th style="width:40mm">Chapitre concerné</th></tr>
<tr><td>1</td><td>Signification des valeurs TOR du classeur de lancement, et part des quantités qui dépend du site (hauteur, nombre d'étages)</td><td>4.1</td></tr>
<tr><td>2</td><td>Règles générales (tous sites) et règles propres à un client</td><td>4.2 à 4.6</td></tr>
<tr><td>3</td><td>Règle de choix et d'équilibrage des contrepoids (béton ou acier, déséquilibre admissible)</td><td>2</td></tr>
<tr><td>4</td><td>Cause de la fissuration des voies de roulement et critère de choix d'une voie rapportée</td><td>4.3</td></tr>
<tr><td>5</td><td>Règles de sécurité à intégrer dès la conception d'une cage (portes, détection, butées, protections)</td><td>4.5</td></tr>
<tr><td>6</td><td>Service BEEM à utiliser pour les plans du BE mécanique (B, E ou M)</td><td>5</td></tr>
<tr><td>7</td><td>Lecture d'un plan d'implantation : calques, cartouche, numérotation des Levmatic et des VM</td><td>3</td></tr>
<tr><td>8</td><td>Poids Levmatic (1500 ou 1960 kg) et pas d'étage entre les onglets du classeur de lancement</td><td>4.1</td></tr>
<tr><td>9</td><td>Formats de palette admis et tolérances de déformation</td><td>2</td></tr></table>
<p class="legend">Une fois ces points validés, chaque règle « déduite » passe à « confirmée », et les chapitres « à compléter » sont rédigés à partir des plans de sites réels (version V2).</p>'''))

# remove empty trailing space; write function to assemble
def toc_html(pages):
    items=[(1,"Historique des modifications","1"),(1,"Données générales et charges admissibles","2"),(1,"Règles générales : FAQ","3"),(1,"Règles de conception","4"),(2,"Cage Levmatic : nomenclature de lancement","4.1"),(2,"Convoyeurs de cage et motoréducteurs","4.2"),(2,"Voies de roulement des Levmatic","4.3"),(2,"Véhicules motorisés (VM)","4.4"),(2,"Sécurité","4.5"),(2,"Chargeurs de batteries et lecteurs de balises","4.6"),(1,"Implantation et nommage des documents","5"),(1,"Adaptation aux anciens sites","6"),(1,"Annexe : points à valider avec le tuteur","7")]
    h='<div class="banner" style="margin-top:0">MAGMATIC</div><div class="toc">'
    for lvl,t,n in items:
        pg=pages.get(n,"")
        h+=f'<div class="{"l2" if lvl==2 else ""}"><span>{t}</span><span>{pg}</span></div>'
    return page(h+'</div>',True)
def assemble(pages):
    out=[]
    for i,p in enumerate(P):
        h=toc_html(pages) if p=="@@TOC@@" else p
        h=h.replace('<span class="pn"></span>',f'<span class="pn">{i+1}</span>')
        out.append(h)
    body="".join(out)
    return f'<!doctype html><html lang="fr"><head><meta charset="utf-8"><style>{CSS}</style></head><body>{body}</body></html>'
def render(html,out):
    open("/tmp/claude-0/g/g.html","w").write(html)
    with sync_playwright() as p:
        b=p.chromium.launch(); pg=b.new_page(); pg.goto("file:///tmp/claude-0/g/g.html")
        pg.pdf(path=out,format="A4",print_background=True,margin=dict(top="0",bottom="0",left="0",right="0"),prefer_css_page_size=True)
        b.close()
out="/mnt/user-data/outputs/Guide_regles_conception_Magmatic_V1.pdf"
render(assemble({}),out)
# find pages
rd=PdfReader(out); txt=[p.extract_text() or "" for p in rd.pages]
pages={}
keys={"1":"1. Historique","2":"2. Données générales","3":"3. Règles générales","4":"4. Règles de conception","4.1":"4.1 Cage","4.2":"4.2 Convoyeurs","4.3":"4.3 Voies","4.4":"4.4 Véhicules","4.5":"4.5 Sécurité","4.6":"4.6 Chargeurs","5":"5. Implantation","6":"6. Adaptation","7":"Annexe : points"}
for k,v in keys.items():
    for i,t in enumerate(txt):
        if i<2: continue
        if v.lower() in t.lower().replace("\n"," "): pages[k]=str(i+1);break
print(pages,len(txt))
render(assemble(pages),out)
print(len(PdfReader(out).pages))
