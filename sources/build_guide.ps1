# Genere le guide Magmatic : decoupe les illustrations puis imprime guide.html en PDF (Edge headless).
# Usage : powershell -ExecutionPolicy Bypass -File sources\build_guide.ps1 [-SkipFigures]
param([switch]$SkipFigures)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$img  = Join-Path $root 'images'
$fig  = Join-Path $PSScriptRoot 'figures'
New-Item -ItemType Directory -Force $fig | Out-Null
Add-Type -AssemblyName System.Drawing

# nom de sortie ; image source ; x ; y ; largeur ; hauteur (0 = image entiere) ; en pixels de l'image d'origine
$crops = @(
 @('cage_plan',        '04_levmatic_cage_nomenclature\04_01.png', 288, 53, 2058, 1464),
 @('cage_elevation',   '04_levmatic_cage_nomenclature\04_01.png', 400, 118, 1300, 1330),
 @('donnees_entree',   '04_levmatic_cage_nomenclature\04_02.png', 24, 142, 1090, 492),
 @('organigramme',     '04_levmatic_cage_nomenclature\04_03.png', 174, 37, 2266, 922),
 @('nommage_exemple',  '06_nommage_documents_BEEM\06_01.png', 30, 356, 1730, 290),
 @('nommage_tables',   '06_nommage_documents_BEEM\06_01.png', 30, 646, 1939, 1089),
 @('implant_first',    '05_plans_implantation_sites\05_09.png', 126, 10, 1812, 1096),
 @('implant_ground',   '05_plans_implantation_sites\05_03.png', 138, 106, 1748, 1176),
 @('implant_coupes',   '05_plans_implantation_sites\05_02.png', 0, 0, 0, 0),
 @('alveole',          '05_plans_implantation_sites\05_04.png', 160, 55, 515, 900),
 @('legende_plan',     '05_plans_implantation_sites\05_07.png', 110, 40, 760, 690),
 @('transfert_tor',    '05_plans_implantation_sites\05_08.png', 62, 75, 997, 479),
 @('bemi_zones',       '03_uniformisation_plans_BEMI\03_02.png', 0, 0, 0, 0),
 @('bemi_couleurs',    '03_uniformisation_plans_BEMI\03_03.png', 0, 0, 0, 0),
 @('bemi_vues',        '03_uniformisation_plans_BEMI\03_04.png', 0, 0, 0, 0),
 @('pdm_coffre',       '02_arborescence_serveur_PDM\02_01.png', 840, 200, 434, 335),
 @('serveur_codif',    '02_arborescence_serveur_PDM\02_03.png', 863, 193, 1046, 528),
 @('serveur_081000',   '02_arborescence_serveur_PDM\02_04.png', 877, 194, 1041, 928),
 @('serveur_next',     '02_arborescence_serveur_PDM\02_05.png', 870, 195, 1085, 1555),
 @('recap_clients',    '08_recap_clients_PG\08_01.png', 0, 344, 3575, 1591)
)
if (-not $SkipFigures) {
  foreach ($c in $crops) {
    $src = [System.Drawing.Image]::FromFile((Join-Path $img $c[1]))
    $x=$c[2]; $y=$c[3]; $w=$c[4]; $h=$c[5]
    if ($w -eq 0) { $x=0; $y=0; $w=$src.Width; $h=$src.Height }
    $w=[Math]::Min($w,$src.Width-$x); $h=[Math]::Min($h,$src.Height-$y)
    $scale=[Math]::Min(1.0,1800.0/$w)
    $ow=[int]($w*$scale); $oh=[int]($h*$scale)
    $bmp = New-Object System.Drawing.Bitmap $ow,$oh
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = 'HighQualityBicubic'
    $g.DrawImage($src,(New-Object System.Drawing.Rectangle 0,0,$ow,$oh),(New-Object System.Drawing.Rectangle $x,$y,$w,$h),[System.Drawing.GraphicsUnit]::Pixel)
    $bmp.Save((Join-Path $fig ($c[0]+'.png')),[System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose(); $src.Dispose()
    "{0}.png  {1}x{2}" -f $c[0],$ow,$oh
  }
}

$edge = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
$html = Join-Path $PSScriptRoot 'guide.html'
$pdf  = Join-Path $root 'livrables\Guide_regles_conception_Magmatic_V2.pdf'
if (Test-Path $html) {
  $uri = ([System.Uri]$html).AbsoluteUri
  $p = Start-Process -FilePath $edge -ArgumentList '--headless','--disable-gpu','--no-pdf-header-footer',"--print-to-pdf=`"$pdf`"",$uri -PassThru -Wait
  Start-Sleep -Seconds 1
  if (Test-Path $pdf) { "PDF : $pdf ({0} Ko)" -f [int]((Get-Item $pdf).Length/1KB) } else { "ECHEC PDF" }
}
