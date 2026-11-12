import openpyxl,re
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as L
R=openpyxl.load_workbook("/mnt/user-data/outputs/Recap_clients_magmatic_rempli_V2.xlsx",data_only=True)["Récapitulatif clients"]
B=openpyxl.load_workbook("/mnt/user-data/uploads/1794214255380_Bilan_affaire_PG.xlsx",data_only=True)["Feuil1"]
def cat(t):
    t=(t or "").lower()
    if "projet initial" in t: return "Projet initial"
    if re.search(r"ajout (de |d')?\d|ajout \d|ajout de|supply|neuf|extension|intégration|integration|\bvm \d",t) and not re.search(r"caméra|chargeur|porte|capteur|coffret",t): return "Ajout d'équipement"
    if re.search(r"retrofit|rétrofit|upgrade|migration|remplacement|obsolescence|changement|réparation|renovation|rénovation|bucket",t): return "Rétrofit / mise à niveau"
    if re.search(r"secu|sécu|porte|cyber|butée|capteur|coffret|protection|caméra",t): return "Sécurité / accessoires"
    return "Autre / à qualifier"
rows=[]
clients=[]
pgn={"P&G AMIENS","P&G EUSKIRCHEN","P&G EUSKIRCHEN AVOCADO","P&G JIJONA","P&G LONDRES","P&G MECHELEN"}
for r in range(2,R.max_row+1):
    c=R.cell(r,1).value
    if not c: continue
    clients.append([c,R.cell(r,2).value,R.cell(r,4).value,R.cell(r,9).value,R.cell(r,11).value])
    if c in pgn: continue
    v=R.cell(r,3).value
    if not v:
        rows.append([c,None,"","À qualifier",None,None,None,"Récap","N° d'affaire non renseigné"]); continue
    for line in str(v).split("\n"):
        line=line.strip()
        if not line: continue
        m=re.match(r"^(\S+)\s*(?:\(([^)]*)\))?\s*(?::\s*(.*))?$",line)
        if line.startswith("B6066"):
            rows.append([c,"B6066","Projet initial","Projet initial",None,None,None,"Récap","2e numéro non renseigné dans le récap (« B6066 / »)"]);continue
        if not m:
            rows.append([c,line,"","À qualifier",None,None,None,"Récap","Format à vérifier"]);continue
        n,d,mat=m.groups()
        init= d and "site initial" in d
        rows.append([c,n,"Projet initial" if init else "Extension",("Rétrofit / mise à niveau" if "emplacement" in (mat or "").lower() else "Ajout d'équipement" if re.search(r"\d+\s*(vm|lev|map)",(mat or "").lower()) else cat(mat)) if not init else "Projet initial",None if init else d,mat,None,"Récap",""])
# columns: client, n°, type, categorie, date, equip, designation, source, note
pg=[]
for r in B.iter_rows(min_row=2,values_only=True):
    if not r[0]: continue
    n=str(r[0]).strip(); des=(r[1] or "").strip()
    cl=(r[3] or "").upper().replace("P&G ","P&G ")
    pg.append([cl,n,"Projet initial" if des.lower().startswith("projet initial") else "Extension",cat(des),None,(r[2] or "").replace("\n"," ") or None,des,"Bilan P&G",""])
# ensure P&G rows' extra info
pgx={}
for r in B.iter_rows(min_row=2,values_only=True):
    if r[0]: pgx[str(r[0]).strip()]=(r[4],r[5])
wb=openpyxl.Workbook()
navy=PatternFill("solid",fgColor="1F3864");yel=PatternFill("solid",fgColor="FFF2CC")
thin=Side(style="thin",color="BFBFBF");bd=Border(top=thin,bottom=thin,left=thin,right=thin)
def head(ws,cols,widths):
    ws.append(cols)
    for c in ws[1]:
        c.font=Font(name="Arial",bold=True,color="FFFFFF",size=10);c.fill=navy;c.alignment=Alignment(wrap_text=True,vertical="center",horizontal="center");c.border=bd
    ws.row_dimensions[1].height=40
    for i,w in enumerate(widths,1): ws.column_dimensions[L(i)].width=w
    ws.freeze_panes="C2"
def body(ws,yc):
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font=Font(name="Arial",size=9);c.alignment=Alignment(wrap_text=True,vertical="top");c.border=bd
            if c.column in yc: c.fill=yel
# Lisez-moi
w0=wb.active;w0.title="Lisez-moi"
txt=["Historique des affaires Magmatic par client – V1 (brouillon)","",
"Objet : retracer, pour chaque client, l'affaire initiale et les évolutions apportées, sur le modèle du Bilan affaire P&G.",
"Onglet Clients : un client par ligne ; mettre à jour la colonne Statut (en service / arrêté / à confirmer).",
"Onglet Affaires : une ligne par affaire. Les 120 affaires P&G viennent du Bilan P&G ; les autres clients sont préremplis depuis le récapitulatif clients (n° d'origine, date, matériel).",
"Cellules jaunes : à compléter (n° BEEM, affaire mère, plans, emplacement des plans, statut de vérification).",
"Les catégories sont proposées automatiquement à partir des mots de la désignation : à corriger.",
"Ordre conseillé : clients en service d'abord, affaires récentes d'abord.",
"Format de n° d'affaire cible : BEEM (3 lettres + 7 chiffres, ex. NPJ0000001). Garder l'ancien n° dans la colonne d'origine."]
for i,t in enumerate(txt,1):
    w0.cell(i,1).value=t;w0.cell(i,1).font=Font(name="Arial",size=14 if i==1 else 10,bold=i==1);w0.cell(i,1).alignment=Alignment(wrap_text=True)
w0.column_dimensions["A"].width=120
# Clients
wc=wb.create_sheet("Clients")
head(wc,["Client","Business","Date d'installation (approx.)","Type Magmatic","Nb VM, LEV, MAP","Nb d'affaires listées","Statut","Remarques"],[28,26,14,12,22,12,16,40])
for i,c in enumerate(clients,2):
    wc.append([c[0],c[1],c[2],c[3],c[4],f"=COUNTIF(Affaires!$A:$A,A{i})",None,None])
    if c[0]=="P&G EUSKIRCHEN AVOCADO": wc.cell(i,8).value="N° 50520 compté dans P&G EUSKIRCHEN (Bilan P&G)"
body(wc,{7,8})
dv=DataValidation(type="list",formula1='"En service,Arrêté,À confirmer"',allow_blank=True);wc.add_data_validation(dv);dv.add(f"G2:G{wc.max_row}")
wc.auto_filter.ref=f"A1:H{wc.max_row}"
last=wc.max_row;n=last+2
wc.cell(n,1).value="Lignes listées (dont n° à renseigner)";wc.cell(n,6).value=f"=SUM(F2:F{last})"
wc.cell(n+1,1).value="Clients en service";wc.cell(n+1,6).value=f'=COUNTIF(G2:G{last},"En service")'
wc.cell(n+2,1).value="Statut non renseigné";wc.cell(n+2,6).value=f'=COUNTA(A2:A{last})-COUNTA(G2:G{last})'
for k in range(3):
    wc.cell(n+k,1).font=Font(name="Arial",bold=True,size=9);wc.cell(n+k,6).font=Font(name="Arial",bold=True,size=9)
# Affaires
wa=wb.create_sheet("Affaires")
head(wa,["Client","N° d'affaire (origine)","N° BEEM normalisé","Affaire mère","Type","Catégorie (proposée)","Date","Équipement / matériel","Désignation projet","Plan d'ensemble","Plans de sous-ensemble","Emplacement des plans (serveur / PDM)","Source","Remarque","Vérifiée ?"],[24,16,16,14,14,22,12,26,36,34,28,28,11,24,10])
for r in pg+rows:
    ex=pgx.get(r[1],(None,None)) if r[7]=="Bilan P&G" else (None,None)
    wa.append([r[0],r[1],None,None,r[2],r[3],r[4],r[5],r[6],ex[0],ex[1],None,r[7],r[8],None])
body(wa,{3,4,12,15})
dv2=DataValidation(type="list",formula1="=Listes!$A$2:$A$7",allow_blank=True);wa.add_data_validation(dv2);dv2.add(f"F2:F{wa.max_row}")
dv3=DataValidation(type="list",formula1='"Oui,Non"',allow_blank=True);wa.add_data_validation(dv3);dv3.add(f"O2:O{wa.max_row}")
dv4=DataValidation(type="list",formula1='"Projet initial,Extension,À qualifier"',allow_blank=True);wa.add_data_validation(dv4);dv4.add(f"E2:E{wa.max_row}")
wa.auto_filter.ref=f"A1:O{wa.max_row}"
for i in range(2,wa.max_row+1):
    v=wa.cell(i,2).value
    if v and re.fullmatch(r"[A-Z]{3}\d{7}",str(v)): wa.cell(i,3).value=v
wl=wb.create_sheet("Listes")
wl.append(["Catégories"])
for t in ["Projet initial","Ajout d'équipement","Rétrofit / mise à niveau","Sécurité / accessoires","Autre / à qualifier","À qualifier"]: wl.append([t])
wl.column_dimensions["A"].width=28
# synthèse sur Lisez-moi
w0.cell(11,1).value="Synthèse";w0.cell(11,1).font=Font(name="Arial",bold=True)
w0.cell(12,1).value='="Affaires listées : "&COUNTA(Affaires!B2:B'+str(wa.max_row)+')&" dont "&COUNTIF(Affaires!A2:A'+str(wa.max_row)+',"P&G*")&" P&G"'
w0.cell(13,1).value='="N° BEEM renseignés : "&COUNTA(Affaires!C2:C'+str(wa.max_row)+')'
w0.cell(14,1).value='="Affaires vérifiées : "&COUNTIF(Affaires!O2:O'+str(wa.max_row)+',"Oui")'
wa.page_setup.orientation="landscape";wa.sheet_properties.pageSetUpPr=openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True);wa.page_setup.fitToHeight=0
out="/mnt/user-data/outputs/Historique_affaires_Magmatic_V1.xlsx";wb.save(out)
print(len(pg),"P&G",len(rows),"autres lignes", len(clients),"clients")
import collections;print(collections.Counter(r[0] for r in rows))
