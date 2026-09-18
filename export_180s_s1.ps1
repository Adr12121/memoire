$app = New-Object -ComObject PowerPoint.Application
$p = (Resolve-Path "Presentation_PFE_180s.pptx").Path
$pres = $app.Presentations.Open($p, 1, 0, 0)
$out = Join-Path (Get-Location).Path "preview_180s_s1.png"
$pres.Slides.Item(1).Export($out, "PNG")
$pres.Close()
$app.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
Write-Host "Exported to $out"
