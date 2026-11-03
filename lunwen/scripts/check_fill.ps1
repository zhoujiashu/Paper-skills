# check_fill.ps1 - lunwen skill
# Scan a generated or filled manuscript docx for:
#   - remaining [[placeholders]]
#   - GUIDE paragraphs that must be deleted before submission
#   - AGU length limits: Key Points (3 x <=100 chars), Abstract (<=250 words),
#     Plain Language Summary (100-200 words)
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File check_fill.ps1 -Path "C:\path\paper.docx"

param(
    [Parameter(Mandatory = $true)][string]$Path
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$zip = [System.IO.Compression.ZipFile]::OpenRead($Path)

function GetText([string]$name) {
    $e = $script:zip.GetEntry($name)
    if (-not $e) { return $null }
    $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
    try { return $sr.ReadToEnd() } finally { $sr.Close() }
}

try {
    $xml = GetText 'word/document.xml'
    if (-not $xml) { Write-Output 'ERROR: word/document.xml not found'; exit 1 }
    $doc = [xml]$xml
    $ns = [System.Xml.XmlNamespaceManager]::new($doc.NameTable)
    $ns.AddNamespace('w', $W)

    $placeholderCount = 0
    $guideCount = 0
    $keyCount = 0
    $keyMax = 0
    $absText = ''
    $plsText = ''
    $idx = 0

    Write-Output '== REMAINING ITEMS =='
    foreach ($p in $doc.SelectNodes('//w:body//w:p', $ns)) {
        $idx++
        $txt = ''
        foreach ($t in $p.SelectNodes('.//w:t', $ns)) { $txt += $t.InnerText }
        if ($txt.Trim().Length -eq 0) { continue }
        $style = ''
        $ps = $p.SelectSingleNode('w:pPr/w:pStyle', $ns)
        if ($ps) { $style = $ps.GetAttribute('val', $W) }

        $m = [regex]::Matches($txt, '\[\[.+?\]\]')
        if ($m.Count -gt 0) {
            $placeholderCount += $m.Count
            $show = $txt
            if ($show.Length -gt 160) { $show = $show.Substring(0, 160) + '...' }
            Write-Output ('P' + $idx + ' [' + $style + '] ' + $m.Count + ' placeholder(s): ' + $show)
        }
        if ($txt.Trim().StartsWith('GUIDE')) {
            $guideCount++
            $show = $txt
            if ($show.Length -gt 120) { $show = $show.Substring(0, 120) + '...' }
            Write-Output ('P' + $idx + ' [GUIDE] ' + $show)
        }
        if ($style -eq 'KeyPoints') {
            $keyCount++
            $n = ([regex]::Replace($txt, '\[\[(.+?)\]\]', '$1')).Length
            if ($n -gt $keyMax) { $keyMax = $n }
        }
        if ($style -eq 'Abstract') { $absText += ' ' + $txt }
        if ($style -eq 'PLS') { $plsText += ' ' + $txt }
    }

    $absWords = 0
    if ($absText.Trim().Length -gt 0) { $absWords = ($absText.Trim() -split '\s+').Count }
    $plsWords = 0
    if ($plsText.Trim().Length -gt 0) { $plsWords = ($plsText.Trim() -split '\s+').Count }

    Write-Output ''
    Write-Output '== SUMMARY =='
    Write-Output ('placeholders`t' + $placeholderCount)
    Write-Output ('guides`t' + $guideCount)
    Write-Output ('keyPoints`t' + $keyCount)
    Write-Output ('keyPointMaxChars`t' + $keyMax)
    Write-Output ('abstractWords`t' + $absWords)
    Write-Output ('plsWords`t' + $plsWords)

    Write-Output ''
    Write-Output '== CHECKS =='
    if ($placeholderCount -gt 0) { Write-Output ('WARNING: ' + $placeholderCount + ' placeholder(s) still highlighted') }
    if ($guideCount -gt 0) { Write-Output ('WARNING: ' + $guideCount + ' GUIDE paragraph(s) must be deleted') }
    if ($keyCount -ne 3) { Write-Output ('WARNING: Key Points must be exactly 3, found ' + $keyCount) }
    if ($keyMax -gt 100) { Write-Output ('WARNING: a Key Point exceeds 100 characters (' + $keyMax + ')') }
    if ($absWords -gt 250) { Write-Output ('WARNING: Abstract exceeds 250 words (' + $absWords + ')') }
    if ($plsWords -gt 0 -and ($plsWords -lt 100 -or $plsWords -gt 200)) {
        Write-Output ('WARNING: Plain Language Summary should be 100-200 words (' + $plsWords + ')')
    }
    if ($placeholderCount -eq 0 -and $guideCount -eq 0) { Write-Output 'OK: no placeholders or guides left' }
}
finally {
    $zip.Dispose()
}
