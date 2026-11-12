# Genere le DOCX du guide a partir de guide.html, via Word (COM). Pas besoin de Python.
# Usage : powershell -ExecutionPolicy Bypass -File sources\build_docx.ps1
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$src  = Join-Path $PSScriptRoot 'guide.html'
$tmp  = Join-Path $PSScriptRoot '_guide_word.html'
$out  = Join-Path $root 'livrables\Guide_regles_conception_Magmatic_V2.docx'
Add-Type -AssemblyName System.Drawing
$t = [IO.File]::ReadAllText($src,[Text.Encoding]::UTF8)
$opt = [Text.RegularExpressions.RegexOptions]::Singleline

# 1. corps seul, sans script
$body = [regex]::Match($t,'<body>(.*)</body>',$opt).Groups[1].Value
$body = [regex]::Replace($body,'<script>.*?</script>','',$opt)
$body = [regex]::Replace($body,'<!--.*?-->','',$opt)

# 2. couverture
$body = [regex]::Replace($body,'<section class="page cover" data-nofoot>(.*?)</section>',{
  param($m) $c=$m.Groups[1].Value
  $c=$c -replace '<h1>(.*?)</h1>','<p class="cv1">$1</p>'
  $c=$c -replace '<div class="s1">(.*?)</div>','<p class="cv2">$1</p>'
  $c=[regex]::Replace($c,'<div class="s2">(.*?)</div>','<p class="cv3">$1</p>',$opt)
  $c=$c -replace '<div class="ref">(.*?)</div>','<p class="cv4">$1</p>'
  $c=$c -replace '<img src="([^"]+)" alt="">','<p align="center"><img src="$1" width="380" height="389"></p>'
  $c=$c -replace '<div class="cap">(.*?)</div>','<p class="cap" align="center">$1</p>'
  $c
},$opt)

# 3. sommaire : remplace par un repere, le champ TOC est insere par Word
$body = $body -replace '<div class="banner">MAGMATIC</div>','<p class="tt">Sommaire</p>'
$body = $body -replace '<div class="toc" id="toc"></div>','<p>[[TOC]]</p>'

# 4. titres
$body = [regex]::Replace($body,'<div class="banner"[^>]*>(.*?)</div>','<h1>$1</h1>')
$body = [regex]::Replace($body,'<h2[^>]*>','<h2>')

# 5. arbre de nomenclature -> tableau
$body = [regex]::Replace($body,'<div class="ch"><small>(.*?)</small>(.*?)<i>(.*?)</i></div>','$1 $2 ($3)<br>')
$body = [regex]::Replace($body,'<div class="trow"><div class="tp"><small>(.*?)</small><b>(.*?)</b><i>(.*?)</i></div><div class="tc">(.*?)</div></div>','<tr><td width="150" bgcolor="#2f4a6d"><font color="#ffffff"><b>$1 $2</b> $3</font></td><td bgcolor="#dff1f7">$4</td></tr>',$opt)
$body = [regex]::Replace($body,'<div class="tree">(.*?)</div>\s*(?=<p class="legend">)','<table border="1" cellspacing="0" cellpadding="4">$1</table>',$opt)

# 6. images : largeur explicite (Word ignore max-height et flex)
$body = [regex]::Replace($body,'<figure([^>]*)><img src="([^"]+)"([^>]*)>',{
  param($m)
  $fa=$m.Groups[1].Value; $p=$m.Groups[2].Value; $ia=$m.Groups[3].Value
  $im=[System.Drawing.Image]::FromFile((Join-Path $PSScriptRoot $p)); $ratio=$im.Height/$im.Width; $im.Dispose()
  $wmm=170.0
  $f=[regex]::Match($fa,'flex:0 0 (\d+)mm'); if($f.Success){ $wmm=[double]$f.Groups[1].Value }
  $h=[regex]::Match($ia,'max-height:(\d+)mm'); if($h.Success){ $hm=[double]$h.Groups[1].Value; if($wmm*$ratio -gt $hm){ $wmm=$hm/$ratio } }
  if($wmm*$ratio -gt 120){ $wmm=120/$ratio }
  '<figure><img src="{0}" width="{1}" height="{2}">' -f $p,[int]($wmm*3.78),[int]($wmm*$ratio*3.78)
},$opt)
$body = $body -replace '<figure>','<p align="center">' -replace '</figure>','</p>'
$body = $body -replace '<figcaption>','<br><span class="cap">' -replace '</figcaption>','</span>'

# 7. encadres et etiquettes
$body = $body -replace '<div class="todo">','<p class="todo">'
$body = $body -replace '<div class="rule">','<p class="rule">'
$body = $body -replace '<div class="warn">','<p class="warn">'
$body = [regex]::Replace($body,'(<p class="(?:todo|rule|warn)">.*?)</div>','$1</p>',$opt)
$body = $body -replace '<span class="tag c">(.*?)</span>','<b><span style="color:#2e7d32;font-size:8pt">[$1]</span></b>'
$body = $body -replace '<span class="tag d">(.*?)</span>','<b><span style="color:#e08a00;font-size:8pt">[$1]</span></b>'
$body = $body -replace '<span class="tag a">(.*?)</span>','<b><span style="color:#c62828;font-size:8pt">[$1]</span></b>'

# 8. sauts de page et nettoyage des conteneurs
$n=0
$body = [regex]::Replace($body,'<section class="page[^>]*>\s*(?=<(h1|p class="tt"))','<br clear="all" style="page-break-before:always">')
$body = [regex]::Replace($body,'<section class="page[^>]*>','')
$body = $body -replace '</section>',''
$body = [regex]::Replace($body,'<div class="cols">|<div>|</div>','')
$body = $body -replace '<table>','<table border="1" cellspacing="0" cellpadding="4" width="100%">'
$body = [regex]::Replace($body,'<th([^>]*)>','<th$1 bgcolor="#6d6d6d"><font color="#ffffff">') -replace '</th>','</font></th>'

$css = @'
body{font-family:Arial;font-size:10pt;color:#333}
h1{font-family:Arial;font-size:15pt;color:#ffffff;background:#6d6d6d;text-transform:uppercase;padding:4pt}
h2{font-family:Arial;font-size:13pt;color:#d9232d}
h3{font-family:Arial;font-size:11pt;color:#555555}
table{border-collapse:collapse;font-size:9pt}
td,th{border:1px solid #bbbbbb;vertical-align:top;font-family:Arial;font-size:9pt}
p.todo{border:1.5pt dashed #c62828;background:#fff5f5;color:#8a1c1c;padding:4pt;font-size:9pt}
p.rule{border-left:4pt solid #d9232d;background:#f6f6f6;padding:4pt}
p.warn{color:#d9232d;font-weight:bold;font-size:12pt;text-align:center}
p.legend{font-size:8.5pt;color:#555555}
.cap{font-size:8pt;color:#666666;font-style:italic}
.k{font-family:Consolas;font-size:9pt}
p.tt{font-size:15pt;font-weight:bold;color:#ffffff;background:#6d6d6d;padding:4pt}
p.cv1{font-size:38pt;font-weight:bold;color:#e31e24;margin-top:60pt}
p.cv2{font-size:21pt;font-style:italic;color:#e31e24}
p.cv3{font-size:20pt;font-weight:bold;color:#e31e24}
p.cv4{font-size:17pt;font-weight:bold;color:#666666}
'@
$html = '<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"><style>'+$css+'</style></head><body>'+$body+'</body></html>'
[IO.File]::WriteAllText($tmp,$html,(New-Object Text.UTF8Encoding $true))

# 9. Word : ouvrir, incorporer les images, sommaire, pied de page, enregistrer
$word = New-Object -ComObject Word.Application
$word.Visible = $false; $word.DisplayAlerts = 0
try {
  [Console]::WriteLine("ouverture..."); $doc = $word.Documents.Open($tmp,$false,$false); [Console]::WriteLine("ouvert")
  $ps = $doc.PageSetup; $ps.PaperSize = 7
  $ps.LeftMargin=51; $ps.RightMargin=51; $ps.TopMargin=57; $ps.BottomMargin=57
  [Console]::WriteLine("images : "+$doc.InlineShapes.Count)
  foreach($s in @($doc.InlineShapes)){ if($s.LinkFormat){ $s.LinkFormat.SavePictureWithDocument=$true } }
  [Console]::WriteLine("tableaux : "+$doc.Tables.Count)
  foreach($tb in @($doc.Tables)){ $tb.PreferredWidthType=2; $tb.PreferredWidth=100 }
  [Console]::WriteLine("sommaire...")
  $r = $doc.Content; $f = $r.Find
  if($f.Execute('[[TOC]]')){ $r.Text=''; [void]$doc.TablesOfContents.Add($r,$true,1,2) }
  $ft = $doc.Sections.Item(1).Footers.Item(1).Range
  $ft.Text = "Magmatic - Principes g$([char]233)n$([char]233)raux et r$([char]232)gles de conception - Ind B (brouillon)   -   page "
  $ft.Font.Size = 8; $ft.Font.Name='Arial'
  $e = $doc.Sections.Item(1).Footers.Item(1).Range; $e.Collapse(0); [void]$e.Fields.Add($e,33)
  [Console]::WriteLine("enregistrement...")
  if(Test-Path $out){ Remove-Item $out -Force }
  [string]$tmpdocx = "$env:TEMP\_guide_magmatic.docx"; $doc.SaveAs([ref]$tmpdocx,[ref]16)
  "pages : " + $doc.ComputeStatistics(2) + " ; images : " + $doc.InlineShapes.Count + " ; tableaux : " + $doc.Tables.Count
  $doc.Close([ref]0); Copy-Item $tmpdocx $out -Force; Remove-Item $tmpdocx -Force
} finally { $word.Quit([ref]0); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
Remove-Item $tmp -Force
"DOCX : $out ({0} Ko)" -f [int]((Get-Item $out).Length/1KB)
