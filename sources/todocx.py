import re,sys
sys.argv=["x"]
src=open("build.py").read().split("out=\"/mnt/user-data/outputs/Guide_regles")[0]
exec(src)  # defines P, tree etc.
from bs4 import BeautifulSoup,NavigableString
from docx import Document
from docx.shared import Pt,Cm,RGBColor,Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH,WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
D=Document()
s=D.sections[0];s.page_width=Mm(210);s.page_height=Mm(297)
s.left_margin=s.right_margin=Cm(1.8);s.top_margin=Cm(2.2);s.bottom_margin=Cm(2.0)
st=D.styles
st["Normal"].font.name="Arial";st["Normal"].font.size=Pt(10)
st["Normal"].element.rPr.rFonts.set(qn("w:eastAsia"),"Arial")
st["Normal"].paragraph_format.space_after=Pt(4)
def hs(name,size,color,bold=True,caps=False):
    h=st[name];h.font.name="Arial";h.font.size=Pt(size);h.font.bold=bold;h.font.color.rgb=RGBColor(*color);h.font.all_caps=caps
    h.element.rPr.rFonts.set(qn("w:ascii"),"Arial");h.element.rPr.rFonts.set(qn("w:hAnsi"),"Arial")
    h.paragraph_format.keep_with_next=True
hs("Heading 1",15,(255,255,255),caps=True);hs("Heading 2",13,(0xD9,0x23,0x2D));hs("Heading 3",11,(0x55,0x55,0x55))
st["Heading 1"].paragraph_format.space_before=Pt(0);st["Heading 1"].paragraph_format.space_after=Pt(10)
st["Heading 2"].paragraph_format.space_before=Pt(10);st["Heading 2"].paragraph_format.space_after=Pt(4)
def shade(el,fill):
    pr=el.get_or_add_pPr() if hasattr(el,"get_or_add_pPr") else el
    sh=OxmlElement("w:shd");sh.set(qn("w:val"),"clear");sh.set(qn("w:color"),"auto");sh.set(qn("w:fill"),fill);pr.append(sh)
def cell_shade(c,fill):
    tcPr=c._tc.get_or_add_tcPr();sh=OxmlElement("w:shd");sh.set(qn("w:val"),"clear");sh.set(qn("w:color"),"auto");sh.set(qn("w:fill"),fill);tcPr.append(sh)
def borders(tbl,color="BBBBBB"):
    t=tbl._tbl;pr=t.tblPr;b=OxmlElement("w:tblBorders")
    for e in("top","left","bottom","right","insideH","insideV"):
        x=OxmlElement(f"w:{e}");x.set(qn("w:val"),"single");x.set(qn("w:sz"),"4");x.set(qn("w:color"),color);b.append(x)
    pr.append(b)
TAGCOL={"c":RGBColor(0x2E,0x7D,0x32),"d":RGBColor(0xE0,0x8A,0),"a":RGBColor(0xC6,0x28,0x28)}
def runs(par,node,size=None,color=None,bold=False,italic=False):
    for ch in node.children:
        if isinstance(ch,NavigableString):
            t=str(ch).replace("\n"," ")
            if t.strip()=="" and not t: continue
            r=par.add_run(t);r.bold=bold;r.italic=italic
            if size:r.font.size=Pt(size)
            if color:r.font.color.rgb=color
        elif ch.name=="span" and "tag" in ch.get("class",[]):
            k=[c for c in ch["class"] if c in TAGCOL][0]
            r=par.add_run("["+ch.get_text()+"]");r.bold=True;r.font.size=Pt(8);r.font.color.rgb=TAGCOL[k]
        elif ch.name=="div" and "foot" in ch.get("class",[]): continue
        elif ch.name in("b","strong"): runs(par,ch,size,color,True,italic)
        elif ch.name=="i" or ch.name=="em": runs(par,ch,size,color,bold,True)
        elif ch.name=="span" and "k" in ch.get("class",[]):
            r=par.add_run(ch.get_text());r.font.name="Consolas";r.font.size=Pt(9)
        else: runs(par,ch,size,color,bold,italic)
def addtable(node):
    rows=node.find_all("tr");ncol=max(len(r.find_all(["td","th"])) for r in rows)
    t=D.add_table(rows=0,cols=ncol);t.alignment=WD_TABLE_ALIGNMENT.CENTER;borders(t);t.autofit=True
    for ri,r in enumerate(rows):
        cells=r.find_all(["td","th"]);row=t.add_row()
        for ci,c in enumerate(cells):
            cc=row.cells[ci];p=cc.paragraphs[0];p.paragraph_format.space_after=Pt(1)
            if c.name=="th":
                cell_shade(cc,"6D6D6D");runs(p,c,9,RGBColor(255,255,255),True)
            else:
                if ri%2==0:cell_shade(cc,"F3F3F3")
                runs(p,c,9)
    import statistics
    w=[]
    for ci in range(ncol):
        L=[len(r.find_all(["td","th"])[ci].get_text()) for r in rows if len(r.find_all(["td","th"]))>ci]
        w.append(min(max(statistics.mean(L),13),60))
    tot=sum(w);t.autofit=False
    for ci in range(ncol): t.columns[ci].width=Cm(17.4*w[ci]/tot)
    for row in t.rows:
        for ci,c in enumerate(row.cells): c.width=Cm(17.4*w[ci]/tot)
    D.add_paragraph().paragraph_format.space_after=Pt(2)
def addtree(node):
    t=D.add_table(rows=0,cols=2);borders(t,"7FB8CF")
    for tr in node.select(".trow"):
        row=t.add_row();a=row.cells[0];b=row.cells[1]
        cell_shade(a,"2F4A6D");a.width=Cm(3.6);b.width=Cm(13.8)
        tp=tr.select_one(".tp");p=a.paragraphs[0]
        parts=[x.get_text() for x in tp.children if getattr(x,"name",None)]
        r=p.add_run(parts[1]);r.bold=True;r.font.size=Pt(9);r.font.color.rgb=RGBColor(255,255,255)
        p.add_run("\n");r=p.add_run(f"{parts[0]}  {parts[2]}");r.font.size=Pt(8);r.font.color.rgb=RGBColor(255,255,255)
        cell_shade(b,"DFF1F7");p=b.paragraphs[0];first=True
        for ch in tr.select(".ch"):
            kids=[x for x in ch.children]
            code=ch.find("small").get_text();q=ch.find("i").get_text()
            name=ch.get_text().replace(code,"").replace(q,"").strip()
            if not first: p=b.add_paragraph()
            first=False
            p.paragraph_format.space_after=Pt(0)
            r=p.add_run(f"{code}  ");r.font.size=Pt(7.5);r.font.color.rgb=RGBColor(0x33,0x60,0x7A)
            r=p.add_run(f"{name}  ");r.font.size=Pt(8.5);r=p.add_run(q);r.font.size=Pt(8.5);r.bold=True
    D.add_paragraph()
def footer():
    sec=D.sections[0];fp=sec.footer.paragraphs[0];fp.text=""
    r=fp.add_run("Magmatic — Principes généraux et règles de conception — Ind A (brouillon)    Page ");r.font.size=Pt(8);r.font.color.rgb=RGBColor(0x77,0x77,0x77)
    r2=fp.add_run();r2.font.size=Pt(8)
    for typ,txt in(("begin",None),(None," PAGE "),("end",None)):
        if typ:
            e=OxmlElement("w:fldChar");e.set(qn("w:fldCharType"),typ);r2._r.append(e)
        else:
            e=OxmlElement("w:instrText");e.set(qn("xml:space"),"preserve");e.text=txt;r2._r.append(e)
footer()
# cover
def para(text,size,color,bold=False,italic=False,before=0,after=6,align=None):
    p=D.add_paragraph();r=p.add_run(text);r.font.size=Pt(size);r.font.color.rgb=color;r.bold=bold;r.italic=italic
    p.paragraph_format.space_before=Pt(before);p.paragraph_format.space_after=Pt(after)
    if align:p.alignment=align
    return p
red=RGBColor(0xE3,0x1E,0x24)
para("MAGMATIC",40,red,True,before=120)
para("Stockage automatisé de palettes",22,red,italic=True,after=40)
para("Principes généraux et règles de conception",22,red,True,after=40)
para("Réf. à attribuer (format BEEM)  Ind A",18,RGBColor(0x66,0x66,0x66),True,after=0)
D.add_page_break()
# TOC
D.add_heading("Sommaire",1)
sh=D.paragraphs[-1];shade(sh._p,"6D6D6D")
p=D.add_paragraph();r=p.add_run();
for typ,txt in(("begin",None),(None,' TOC \\o "1-2" \\h \\z \\u '),("separate",None),("end",None)):
    if typ:
        e=OxmlElement("w:fldChar");e.set(qn("w:fldCharType"),typ);r._r.append(e)
    else:
        e=OxmlElement("w:instrText");e.set(qn("xml:space"),"preserve");e.text=txt;r._r.append(e)
t=OxmlElement("w:t");t.text="Clique droit puis « Mettre à jour les champs » pour afficher le sommaire.";r._r.append(t)
D.add_page_break()
# content pages (skip cover and toc)
for idx,pg in enumerate(P[2:]):
    soup=BeautifulSoup(pg,"html.parser");root=soup.select_one(".page")
    for el in root.children:
        if not getattr(el,"name",None): continue
        cl=el.get("class",[])
        if "top" in cl or "bot" in cl or "foot" in cl: continue
        if "banner" in cl:
            h=D.add_heading(el.get_text(),1);shade(h._p,"6D6D6D")
        elif el.name=="h2": D.add_heading(el.get_text(),2)
        elif el.name=="h3": D.add_heading(el.get_text(),3)
        elif el.name=="p":
            p=D.add_paragraph()
            if "legend" in cl: runs(p,el,9,RGBColor(0x55,0x55,0x55))
            else: runs(p,el)
        elif el.name=="ul":
            for li in el.find_all("li"):
                p=D.add_paragraph(style="List Bullet");runs(p,li)
        elif el.name=="table": addtable(el)
        elif "tree" in cl: addtree(el)
        elif "todo" in cl:
            p=D.add_paragraph();shade(p._p,"FFF5F5")
            pb=OxmlElement("w:pBdr")
            for e in("top","left","bottom","right"):
                x=OxmlElement(f"w:{e}");x.set(qn("w:val"),"dashed");x.set(qn("w:sz"),"6");x.set(qn("w:color"),"C62828");x.set(qn("w:space"),"3");pb.append(x)
            p._p.get_or_add_pPr().append(pb)
            runs(p,el,9,RGBColor(0x8A,0x1C,0x1C))
        elif "warn" in cl:
            p=D.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_before=Pt(30)
            r=p.add_run(el.get_text());r.bold=True;r.font.size=Pt(12);r.font.color.rgb=RGBColor(0xD9,0x23,0x2D)
    if idx<len(P[2:])-1: D.add_page_break()
out="/mnt/user-data/outputs/Guide_regles_conception_Magmatic_V1.docx"
st_=D.settings.element;u=OxmlElement("w:updateFields");u.set(qn("w:val"),"true");st_.append(u)
D.save(out);print("ok")
