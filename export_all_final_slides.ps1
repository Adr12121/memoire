$app = New-Object -ComObject PowerPoint.Application
$p = (Resolve-Path "Soutenance_PFE_Adrien_TRAVAILLE.pptx").Path
$pres = $app.Presentations.Open($p, 1, 0, 0)
$outDir = Join-Path (Get-Location).Path "final_slides_preview"
if (!(Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

$count = $pres.Slides.Count
for ($i = 1; $i -le $count; $i++) {
    $numStr = $i.ToString("00")
    $outFile = Join-Path $outDir "slide_$numStr.png"
    $pres.Slides.Item($i).Export($outFile, "PNG")
}

$pres.Close()
$app.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
Write-Host "Exported $count slides to $outDir"
