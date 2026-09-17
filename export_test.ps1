$app = New-Object -ComObject PowerPoint.Application
$p = (Resolve-Path "test_180s_header.pptx").Path
$pres = $app.Presentations.Open($p, 1, 0, 0)
$out = Join-Path (Get-Location).Path "test_180s_header.png"
$pres.Slides.Item(1).Export($out, "PNG")
$pres.Close()
$app.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
Write-Host "Exported test_180s_header.png"
