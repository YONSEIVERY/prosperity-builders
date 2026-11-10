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
# 대사관 요청 예비용 PPTX: 장마다 PNG 한 장을 꽉 채운다(글 수정 불가, 클리커로 넘기기용)
$pptx = Join-Path $out 'VERY_대사관발표_비상용.pptx'
$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Add(0)
$pres.PageSetup.SlideWidth = 960; $pres.PageSetup.SlideHeight = 540
$i = 0
foreach ($png in Get-ChildItem (Join-Path $out 'slides') -Filter 's*.png' | Sort-Object Name) {
  $i++; $s = $pres.Slides.Add($i, 12)
  $s.Shapes.AddPicture($png.FullName, 0, -1, 0, 0, 960, 540) | Out-Null
}
$pres.SaveAs($pptx, 24); $pres.Close(); $pp.Quit()
"pptx $i"
