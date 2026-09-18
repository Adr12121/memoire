$app = New-Object -ComObject PowerPoint.Application
$p = (Resolve-Path "Presentation_PFE_180s.pptx").Path
$pres = $app.Presentations.Open($p, 1, 0, 0)
$outDir = Join-Path (Get-Location).Path "preview_180s_all"
if (!(Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $numStr = $i.ToString("00")
    $outFile = Join-Path $outDir "slide_$numStr.png"
    $pres.Slides.Item($i).Export($outFile, "PNG")
}

$pres.Close()
$app.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
Write-Host "Exported all 180s slides to $outDir"
