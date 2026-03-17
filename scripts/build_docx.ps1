$template = 'D:\skills\活动方案\主题党日活动方案.docx'
$outDir = $env:OUTPUT_DIR
if (-not $outDir) { $outDir = 'E:\05-党员活动\2026年支部活动' }
$yearMonth = $env:YEAR_MONTH
if (-not $yearMonth) { throw 'YEAR_MONTH env var required, e.g. 2026年1月' }
$outDocx = Join-Path $outDir ('智能制造管控党支部-主题党日记录表（' + $yearMonth + '）.docx')
$work = 'd:\cursor project\_tmp_docx_build'

if (Test-Path $work) { Remove-Item -Recurse -Force $work }
New-Item -ItemType Directory -Path $work | Out-Null
Copy-Item $template "$work\template.zip" -Force
Expand-Archive -Path "$work\template.zip" -DestinationPath "$work\unzipped" -Force

$xmlPath = "$work\unzipped\word\document.xml"
$bytes = [System.IO.File]::ReadAllBytes($xmlPath)
$text = [System.Text.Encoding]::UTF8.GetString($bytes)

function Escape-Xml([string]$s) {
  return $s.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;')
}

function Replace-Para([string]$id, [string]$newText) {
  $pattern = '(<w:p[^>]*w14:paraId="' + $id + '"[^>]*>)(.*?)(</w:p>)'
  $m = [regex]::Match($script:text, $pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
  if (-not $m.Success) { return }
  $pOpen = $m.Groups[1].Value
  $pBody = $m.Groups[2].Value
  $pClose = $m.Groups[3].Value
  $pPr = [regex]::Match($pBody, '(<w:pPr>.*?</w:pPr>)', [System.Text.RegularExpressions.RegexOptions]::Singleline)
  $rPr = [regex]::Match($pBody, '(<w:rPr>.*?</w:rPr>)', [System.Text.RegularExpressions.RegexOptions]::Singleline)
  if (-not $pPr.Success) { return }
  $rPrText = if ($rPr.Success) { $rPr.Groups[1].Value } else { '' }
  $escaped = Escape-Xml $newText
  $newBody = $pPr.Groups[1].Value + '<w:r>' + $rPrText + '<w:t>' + $escaped + '</w:t></w:r>'
  $script:text = $script:text.Substring(0, $m.Index) + $pOpen + $newBody + $pClose + $script:text.Substring($m.Index + $m.Length)
}

$TITLE = $env:TITLE
$BACKGROUND = $env:BACKGROUND
$TIME = $env:TIME
$LOCATION = $env:LOCATION
$CONTENT1 = $env:CONTENT1
$CONTENT2 = $env:CONTENT2
$CONTENT3 = $env:CONTENT3
$CONTENT4 = $env:CONTENT4
$CONTENT5 = $env:CONTENT5
$REQ1 = $env:REQ1
$REQ2 = $env:REQ2
$REQ3 = $env:REQ3

if (-not $TITLE) { throw 'TITLE env var required' }
if (-not $BACKGROUND) { throw 'BACKGROUND env var required' }
if (-not $TIME) { throw 'TIME env var required' }
if (-not $LOCATION) { throw 'LOCATION env var required' }
if (-not $CONTENT1) { throw 'CONTENT1 env var required' }
if (-not $CONTENT2) { throw 'CONTENT2 env var required' }
if (-not $CONTENT3) { throw 'CONTENT3 env var required' }
if (-not $CONTENT4) { throw 'CONTENT4 env var required' }
if (-not $CONTENT5) { throw 'CONTENT5 env var required' }
if (-not $REQ1) { throw 'REQ1 env var required' }
if (-not $REQ2) { throw 'REQ2 env var required' }
if (-not $REQ3) { throw 'REQ3 env var required' }

Replace-Para '6E3305A9' $TITLE
Replace-Para '0E411C12' $BACKGROUND
Replace-Para '6D9297A0' $TIME
Replace-Para '3D443CF2' $LOCATION
Replace-Para '7E1BC118' $CONTENT1
Replace-Para '73E5F5AE' $CONTENT2
Replace-Para '002FA72D' $CONTENT3
Replace-Para '055920AF' $CONTENT4
Replace-Para '42BED460' $CONTENT5
Replace-Para '2E475C59' $REQ1
Replace-Para '6EA4F3EC' $REQ2
Replace-Para '0BBF58E1' $REQ3

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllBytes($xmlPath, $utf8NoBom.GetBytes($text))

$outZip = "$work\out.zip"
if (Test-Path $outZip) { Remove-Item $outZip -Force }
Compress-Archive -Path "$work\unzipped\*" -DestinationPath $outZip -Force
Copy-Item $outZip $outDocx -Force

Write-Output $outDocx