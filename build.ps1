# VERY 대사관 발표: 발표는 index.html(GitHub Pages)로 한다. 이 스크립트는 비상용 PDF와 검수용 장별 PNG만 만든다.
# 실행: pwsh -File build.ps1
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$out = Join-Path (Split-Path -Parent $here) 'out'
New-Item -ItemType Directory -Force $out | Out-Null
$chrome = 'C:\Program Files\Google\Chrome\Application\chrome.exe'
$html = (Join-Path $here 'index.html')
$pdf = Join-Path $out 'VERY_대사관발표_비상용.pdf'
# ?print 를 붙이면 발표 모드(한 장씩, 모션) 대신 장을 늘어놓은 정적 배치로 인쇄된다
$url = 'file:///' + ($html -replace '\\', '/') + '?print=1'
& $chrome --headless=new --disable-gpu --no-pdf-header-footer --run-all-compositor-stages-before-draw --virtual-time-budget=8000 "--print-to-pdf=$pdf" $url 2>$null | Out-Null
Start-Sleep -Milliseconds 500
python -W ignore (Join-Path $here 'topng.py') $pdf $out
