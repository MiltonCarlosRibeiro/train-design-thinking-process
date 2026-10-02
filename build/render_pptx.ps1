param([string]$pptx, [string]$outdir)
New-Item -ItemType Directory -Force $outdir | Out-Null
Get-ChildItem $outdir -Filter *.png | Remove-Item -Force
$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Open($pptx, $true, $false, $false)
foreach ($s in $pres.Slides) { $s.Export((Join-Path $outdir ("s{0:D3}.png" -f $s.SlideIndex)), "PNG", 1333, 750) }
$pres.Close(); $pp.Quit()
