# Projet : Capitalisation documentaire Magmatic (Savoye, BE mécanique)

Utilisateur : Samih Mihci, BUT GMP en alternance chez Savoye (voie SINREV).
Tuteur entreprise : CHABOT Philippe. Tuteur école : WADOUX Laurine.
Langue de travail : français.
Style de réponse : court et direct, pas trop verbeux. Aller à l'essentiel, éviter les longues explications.

## Objectif
Projet d'alternance en 5 phases. Phase 2 :
1. Historique et organisation Magmatic (affaires initiales et évolutions par client, sur le modèle du Bilan affaire P&G).
2. Guide des règles de conception Magmatic V1 (structure calquée sur le guide Intelis ST1100-10-00009 Ind M ; valeurs Intelis NON reprises).

## Guide sur Canva (V3, source actuelle)
- Design Canva : https://www.canva.com/d/SOxtdxd7dM_qq9S (id DAHaoSVTKwQ) - 23 pages 1920x1080, charte rouge Savoye.
- Export : livrables/Guide_regles_conception_Magmatic_V3_Canva.pdf.
- 16 illustrations tirées de images/, téléversées dans Canva (assets MAHao...).
- ATTENTION : la génération Canva avait produit une coquille vide (cellules « À reprendre », règles génériques inventées). Tout le contenu technique a été réinjecté à la main page par page. Ne jamais laisser une valeur générée par Canva dans ce guide.
- Mise en page relue et corrigée page par page (11/11/2026) : plus aucun texte tronqué ni superposé. Les cellules de tableau Canva grandissent avec le texte et débordent sur la ligne suivante : vérifier la hauteur de chaque cellule après remplissage, et réduire la taille de police plutôt que d'allonger le texte.

## Arborescence
- livrables/ : Guide V2 (PDF 18 pages + DOCX 23 pages, Ind B, illustrés) ; Guide V1 (PDF + DOCX, ancien) ; Historique_affaires_Magmatic_V1.xlsx ; Recap clients (vide, V1, V2)
- sources/guide.html : SOURCE ACTUELLE du guide (V2, Ind B). Une <section class="page"> par page A4 ; sommaire et n° de page générés par le script en fin de fichier.
- sources/build_guide.ps1 : découpe les illustrations (images/ -> sources/figures/) et imprime guide.html en PDF avec Edge headless. Pas besoin de Python. Option -SkipFigures.
- sources/build_docx.ps1 : convertit guide.html en DOCX via Word (COM). Pièges : ne jamais ouvrir un document en lecture seule par COM (Word se bloque), passer les chemins en [string], donner largeur ET hauteur aux images.
- Contrôle de débordement : msedge --headless --dump-dom guide.html, lire l'attribut data-reste-mm de <body> (négatif = page qui déborde).
- sources/build.py (OBSOLÈTE, guide V1) : génère le guide en HTML puis PDF (Playwright/Chromium). Liste P de pages, helpers chip/todo/sec/page.
- sources/todocx.py : convertit le HTML en DOCX (bs4 + python-docx). Réutilise P de build.py.
- sources/hist.py : génère Historique_affaires_Magmatic_V1.xlsx (openpyxl).
- donnees_entree/ : classeur de lancement EHP01718LEVM001 (cage Levmatic NEXT EXPANSION), converti en .xlsx.
Les scripts ont des chemins en dur (/tmp/claude-0/..., /mnt/user-data/outputs) : à adapter en local.

## Conventions
- Niveaux de confiance dans le guide : Confirmé / Déduit / À compléter. Ne jamais présenter une déduction comme confirmée.
- N° d'affaire BEEM = 3 lettres + 7 chiffres (ex. EHP01718). Dans le Bilan P&G, seules 25 affaires sur 120 sont au format BEEM.
- Nommage documents BEEM : 9 champs, famille MAG ; codes service A BEA / B BEEM / E BEE / M BMI (à confirmer).

## Points ouverts (à valider avec le tuteur)
- Poids Levmatic : 1500 kg (Données Palettier) vs 1960 kg (CP et organigramme).
- Pas d'étage : 1850 = charge 1500 + 350 (confirmé par le détail « Pallets in a bay », images 05_04). Reste à expliquer : hauteur 1800 du tableau palettier, et débordement 2x100 vs charge 1400 sur palette 1000.
- Sens du sigle TOR (TOR 450 sur convoyeur, TOR 300 en rayonnage) ; définition des types 1 / 2 / 2A / 2C.
- Règles de mise en plan BEMI (présentation 16/08/2023) : applicables telles quelles à Magmatic ?
- Jeu cadre de réa / contrepoids négatif (-280 mm) sur NEXT.
- Sens de « chaud / froid » (ambiance) et de « EURO 1000 » (1000 = charge en kg, déduit).
- Onglet « CP et organigramme » : calcul des contrepoids acier en #NAME?.
- 13 clients sans n° d'affaire ; Pliva = B6066 (second n° manquant) ; service BEEM à utiliser (B, E ou M).

## Prochaines étapes
- Remplir l'historique des autres clients (statut, plans).
- (11/11/2026) Espaces « Valeursde » : bug réel visible sur images/10_captures_guide_magmatic/10_01.png (ancienne version du tableau, 2 colonnes). La source actuelle (build.py, tableau 3 colonnes avec « Niveau »), le DOCX et le Claude Doc ont bien les espaces. PDF livré non vérifié visuellement : le régénérer dès que Python est installé.
- Dossier images/ : 77 captures triées en 10 sous-dossiers (plan d'action, serveur PDM, plans BEMI, nomenclature cage, implantations, nommage BEEM, livret IUT, recap clients P&G, guide Intelis, captures guide). Index dans images/README.md.
- (11/11/2026) Claude Doc en ligne (https://claude.ai/artifact/LV8uQnXcMo11s8oEn2SvfW) : calculs de géométrie, crémaillères et repères automate du §4.1 ajoutés dans la section 2.
- Python n'est pas installé sur le PC (seul l'alias Microsoft Store existe) : l'installer avant de régénérer PDF/DOCX/XLSX.
- (11/11/2026) Guide refait intégralement en V2 à partir des 77 captures. DOCX V2 généré (mise en page simplifiée, non relue à l'écran). Puis V3 refaite sur Canva (voir en haut). Claude Doc en ligne pas aligné sur V2 ni V3.
- Rédiger les chapitres « à compléter » après réponses du tuteur (liste en annexe du guide, 11 points).
