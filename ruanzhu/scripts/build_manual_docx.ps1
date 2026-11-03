# build_manual_docx.ps1 - ruanzhu skill
# Build a software-copyright manual docx by template-package replacement:
# copy the reference docx, rewrite word/document.xml from a content DSL file,
# optionally update docProps/core.xml title, then validate by opening in Word.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File build_manual_docx.ps1 `
#     -Template "C:\path\ref.docx" -Content "C:\path\content.txt" `
#     -Out "C:\path\out.docx" [-Title "..."] [-SkipValidation]
#
# IMPORTANT: keep this script pure ASCII. Put all Chinese text in the content
# DSL file (UTF-8) or pass via -Title on the command line.

param(
    [Parameter(Mandatory = $true)][string]$Template,
    [Parameter(Mandatory = $true)][string]$Content,
    [Parameter(Mandatory = $true)][string]$Out,
    [string]$Title = '',
    [switch]$SkipValidation
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---- template style IDs (verify against the template's styles.xml, tweak if needed) ----
$STYLE_H1   = '1'                    # heading 1
$STYLE_H2   = '2'                    # heading 2
$STYLE_H3   = 'ad'                   # heading 3 / Title alias
$STYLE_LBL  = 'a8'                   # List Paragraph (labels, unnumbered section head)
$STYLE_EQ   = 'af'                   # equation paragraph: center tab 4536 + right tab 9072
$STYLE_TBL  = 'a9'                   # Table Grid
$STYLE_BIB  = 'EndNoteBibliography'  # references
$NUMID_SEC1 = '1'                    # auto number for H1A (unnumbered section like background)
$NUMID_LBL  = '8'                    # auto number for BLAB labels

$song = [string][char]0x5B8B + [string][char]0x4F53   # SimSun in Chinese

function EscX([string]$s) {
    return $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

function RunXml([string]$t, [bool]$it, [bool]$sbv, [bool]$spv, [bool]$bd, [string]$mode) {
    $rpr = '<w:rFonts w:cs="Times New Roman" w:hint="eastAsia"/>'
    if ($bd) { $rpr += '<w:b/>' }
    if ($it) { $rpr += '<w:i/>' }
    if ($mode -eq 'small') { $rpr += '<w:sz w:val="21"/><w:szCs w:val="21"/>' }
    else { $rpr += '<w:szCs w:val="24"/>' }
    if ($sbv) { $rpr += '<w:vertAlign w:val="subscript"/>' }
    if ($spv) { $rpr += '<w:vertAlign w:val="superscript"/>' }
    return '<w:r><w:rPr>' + $rpr + '</w:rPr><w:t xml:space="preserve">' + $t + '</w:t></w:r>'
}

function MarkedRuns([string]$s, [string]$mode) {
    $out = New-Object System.Text.StringBuilder
    $rx = New-Object System.Text.RegularExpressions.Regex('\[(/?)(i|sb|sp|b)\]')
    $ms = $rx.Matches($s)
    $it = $false; $sb = $false; $sp = $false; $bd = $false
    $pos = 0
    foreach ($m in $ms) {
        $piece = $s.Substring($pos, $m.Index - $pos)
        if ($piece.Length -gt 0) { [void]$out.Append((RunXml (EscX $piece) $it $sb $sp $bd $mode)) }
        $closing = ($m.Groups[1].Value -ne '')
        switch ($m.Groups[2].Value) {
            'i'  { $it = -not $closing }
            'sb' { $sb = -not $closing }
            'sp' { $sp = -not $closing }
            'b'  { $bd = -not $closing }
        }
        $pos = $m.Index + $m.Length
    }
    $tail = $s.Substring($pos)
    if ($tail.Length -gt 0) { [void]$out.Append((RunXml (EscX $tail) $it $sb $sp $bd $mode)) }
    return $out.ToString()
}

function ParaTitle([string]$t) {
    return '<w:p><w:pPr><w:spacing w:afterLines="50" w:after="156"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="center"/><w:rPr><w:rFonts w:cs="Times New Roman"/><w:b/><w:sz w:val="32"/><w:szCs w:val="36"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:cs="Times New Roman" w:hint="eastAsia"/><w:b/><w:sz w:val="32"/><w:szCs w:val="36"/></w:rPr><w:t xml:space="preserve">' + (EscX $t) + '</w:t></w:r></w:p>'
}

function ParaH1([string]$t) {
    return '<w:p><w:pPr><w:pStyle w:val="' + $script:STYLE_H1 + '"/><w:rPr><w:rFonts w:hint="eastAsia"/><w:szCs w:val="24"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t xml:space="preserve">' + (EscX $t) + '</w:t></w:r></w:p>'
}

function ParaH2([string]$t) {
    return '<w:p><w:pPr><w:pStyle w:val="' + $script:STYLE_H2 + '"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t xml:space="preserve">' + (EscX $t) + '</w:t></w:r></w:p>'
}

function ParaH3([string]$t) {
    return '<w:p><w:pPr><w:pStyle w:val="' + $script:STYLE_H3 + '"/><w:rPr><w:szCs w:val="24"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t xml:space="preserve">' + (EscX $t) + '</w:t></w:r></w:p>'
}

function ParaH1A([string]$t) {
    return '<w:p><w:pPr><w:pStyle w:val="' + $script:STYLE_LBL + '"/><w:keepNext/><w:keepLines/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="' + $script:NUMID_SEC1 + '"/></w:numPr><w:spacing w:before="240" w:after="240"/><w:ind w:firstLineChars="0"/><w:outlineLvl w:val="0"/><w:rPr><w:rFonts w:cs="Times New Roman"/><w:b/><w:bCs/><w:kern w:val="44"/><w:sz w:val="30"/><w:szCs w:val="44"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:cs="Times New Roman" w:hint="eastAsia"/><w:b/><w:bCs/><w:kern w:val="44"/><w:sz w:val="30"/><w:szCs w:val="44"/></w:rPr><w:t xml:space="preserve">' + (EscX $t) + '</w:t></w:r></w:p>'
}

function ParaBLAB([string]$t) {
    return '<w:p><w:pPr><w:pStyle w:val="' + $script:STYLE_LBL + '"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="' + $script:NUMID_LBL + '"/></w:numPr><w:ind w:firstLineChars="0"/></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t xml:space="preserve">' + (EscX $t) + '</w:t></w:r></w:p>'
}

function ParaBLIT([string]$t) {
    return '<w:p><w:pPr><w:pStyle w:val="' + $script:STYLE_LBL + '"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="left"/><w:rPr><w:rFonts w:cs="Times New Roman"/><w:szCs w:val="24"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:cs="Times New Roman" w:hint="eastAsia"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">' + (EscX $t) + '</w:t></w:r></w:p>'
}

function ParaBody([string]$t) {
    return '<w:p><w:pPr><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:spacing w:line="400" w:lineRule="exact"/><w:ind w:firstLine="480"/><w:rPr><w:rFonts w:cs="Times New Roman"/><w:szCs w:val="24"/></w:rPr></w:pPr>' + (MarkedRuns $t 'body') + '</w:p>'
}

function ParaEqn([string]$num, [string]$latex) {
    $mid = ''
    if ($latex -ne '') {
        $rpr = '<w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math" w:cs="Times New Roman" w:hint="eastAsia"/><w:i/><w:color w:val="808080"/><w:sz w:val="18"/><w:szCs w:val="18"/>'
        $mid = '<w:r><w:rPr>' + $rpr + '</w:rPr><w:t xml:space="preserve">' + (EscX $latex) + '</w:t></w:r>'
    }
    return '<w:p><w:pPr><w:pStyle w:val="' + $script:STYLE_EQ + '"/><w:rPr><w:rFonts w:cs="Times New Roman"/><w:szCs w:val="24"/></w:rPr></w:pPr><w:r><w:tab/></w:r>' + $mid + '<w:r><w:tab/></w:r><w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">' + (EscX $num) + '</w:t></w:r></w:p>'
}

function ParaFig([string]$cap) {
    return '<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p><w:p><w:pPr><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="center"/><w:rPr><w:rFonts w:cs="Times New Roman"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:pPr>' + (MarkedRuns $cap 'small') + '</w:p>'
}

function ParaTcap([string]$cap) {
    return '<w:p><w:pPr><w:ind w:firstLine="420"/><w:jc w:val="center"/><w:rPr><w:rFonts w:cs="Times New Roman"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:pPr>' + (MarkedRuns $cap 'small') + '</w:p>'
}

function ParaRef([string]$t) {
    $rp = '<w:rFonts w:ascii="Times New Roman" w:eastAsia="' + $script:song + '" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="21"/><w:szCs w:val="21"/>'
    return '<w:p><w:pPr><w:pStyle w:val="' + $script:STYLE_BIB + '"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="left"/><w:rPr>' + $rp + '</w:rPr></w:pPr><w:r><w:rPr>' + $rp + '</w:rPr><w:t xml:space="preserve">' + (EscX $t) + '</w:t></w:r></w:p>'
}

function EmitTable($rows) {
    $sb2 = New-Object System.Text.StringBuilder
    [void]$sb2.Append('<w:tbl><w:tblPr><w:tblStyle w:val="' + $script:STYLE_TBL + '"/><w:tblW w:w="0" w:type="auto"/><w:jc w:val="center"/><w:tblBorders><w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:tblBorders><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="1922"/><w:gridCol w:w="1659"/><w:gridCol w:w="1659"/><w:gridCol w:w="1659"/></w:tblGrid>')
    $n = $rows.Count
    for ($r = 0; $r -lt $n; $r++) {
        $cells = $rows[$r] -split '##'
        [void]$sb2.Append('<w:tr><w:trPr><w:jc w:val="center"/></w:trPr>')
        for ($c = 0; $c -lt 4; $c++) {
            $w = 1659; if ($c -eq 0) { $w = 1922 }
            $txt = ''; if ($c -lt $cells.Count) { $txt = $cells[$c] }
            $borders = ''
            if ($r -eq 0) {
                $borders = '<w:tcBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="auto"/>'
                if ($c -eq 0) { $borders += '<w:left w:val="nil"/>' }
                $borders += '<w:bottom w:val="single" w:sz="12" w:space="0" w:color="auto"/></w:tcBorders>'
            } elseif ($r -eq 1) {
                $borders = '<w:tcBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="auto"/>'
                if ($c -eq 0) { $borders += '<w:left w:val="nil"/>' }
                $borders += '</w:tcBorders>'
            } elseif ($r -eq ($n - 1)) {
                $borders = '<w:tcBorders><w:top w:val="nil"/>'
                if ($c -eq 0) { $borders += '<w:left w:val="nil"/>' }
                $borders += '<w:bottom w:val="single" w:sz="12" w:space="0" w:color="auto"/></w:tcBorders>'
            }
            [void]$sb2.Append('<w:tc><w:tcPr><w:tcW w:w="' + $w + '" w:type="dxa"/>' + $borders + '</w:tcPr><w:p><w:pPr><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="center"/><w:rPr><w:rFonts w:cs="Times New Roman"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:pPr>' + (MarkedRuns $txt 'small') + '</w:p></w:tc>')
        }
        [void]$sb2.Append('</w:tr>')
    }
    [void]$sb2.Append('</w:tbl>')
    return $sb2.ToString()
}

# ---- extract root start tag and sectPr from the template package itself ----
$zipR = [System.IO.Compression.ZipFile]::OpenRead($Template)
try {
    $entry = $zipR.GetEntry('word/document.xml')
    $sr = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    $raw = $sr.ReadToEnd(); $sr.Close()
} finally { $zipR.Dispose() }

$i0 = $raw.IndexOf('<w:document')
$i1 = $raw.IndexOf('>', $i0)
if ($i0 -lt 0 -or $i1 -lt 0) { Write-Output 'TEMPLATE_INVALID: <w:document> tag not found'; exit 1 }
$rootTag = $raw.Substring($i0, $i1 - $i0 + 1)

$xmlT = [xml]$raw
$ns = New-Object System.Xml.XmlNamespaceManager($xmlT.NameTable)
$ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
$sectNode = $xmlT.SelectSingleNode('//w:body/w:sectPr', $ns)
if ($sectNode) { $sectPr = $sectNode.OuterXml } else { $sectPr = '' }

# ---- parse content DSL and build body ----
$lines = [System.IO.File]::ReadAllLines($Content, [System.Text.Encoding]::UTF8)
$body = New-Object System.Text.StringBuilder
$rows = New-Object System.Collections.Generic.List[string]
$eqnCount = 0
$tblCount = 0

foreach ($line in $lines) {
    if ($line.Trim().Length -eq 0) { continue }
    $sep = $line.IndexOf('|')
    if ($sep -lt 1) { continue }
    $tag = $line.Substring(0, $sep)
    $val = $line.Substring($sep + 1)
    if (($tag -ne 'TBLH') -and ($tag -ne 'TBLR') -and ($rows.Count -gt 0)) {
        [void]$body.Append((EmitTable $rows)); $rows.Clear(); $tblCount++
    }
    switch ($tag) {
        'TITLE' { [void]$body.Append((ParaTitle $val)) }
        'H1A'   { [void]$body.Append((ParaH1A $val)) }
        'H1'    { [void]$body.Append((ParaH1 $val)) }
        'H2'    { [void]$body.Append((ParaH2 $val)) }
        'H3'    { [void]$body.Append((ParaH3 $val)) }
        'BLAB'  { [void]$body.Append((ParaBLAB $val)) }
        'BLIT'  { [void]$body.Append((ParaBLIT $val)) }
        'B'     { [void]$body.Append((ParaBody $val)) }
        'EQN'   {
            $num = $val; $latex = ''
            $s2 = $val.IndexOf('|')
            if ($s2 -ge 0) { $num = $val.Substring(0, $s2); $latex = $val.Substring($s2 + 1) }
            [void]$body.Append((ParaEqn $num $latex)); $eqnCount++
        }
        'FIG'   { [void]$body.Append((ParaFig $val)) }
        'TCAP'  { [void]$body.Append((ParaTcap $val)) }
        'TBLH'  { if ($rows.Count -gt 0) { [void]$body.Append((EmitTable $rows)); $rows.Clear(); $tblCount++ }; $rows.Add($val) }
        'TBLR'  { $rows.Add($val) }
        'REFH'  { [void]$body.Append((ParaH2 $val)) }
        'REF'   { [void]$body.Append((ParaRef $val)) }
    }
}
if ($rows.Count -gt 0) { [void]$body.Append((EmitTable $rows)); $rows.Clear(); $tblCount++ }

$xmlStr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + $rootTag + '<w:body>' + $body.ToString() + $sectPr + '</w:body></w:document>'

# ---- well-formedness check ----
try { $chk = [xml]$xmlStr } catch {
    Write-Output ('XML_BUILD_FAILED: ' + $_.Exception.Message)
    exit 1
}
Write-Output ('XML OK. Equations: ' + $eqnCount + '  Tables: ' + $tblCount)

# ---- package: copy template, replace word/document.xml (+ core.xml title) ----
try {
    Copy-Item -LiteralPath $Template -Destination $Out -Force
} catch {
    Write-Output ('OUTPUT_LOCKED_OR_UNWRITABLE: ' + $_.Exception.Message)
    exit 1
}

$zip = [System.IO.Compression.ZipFile]::Open($Out, 'Update')
try {
    $e = $zip.GetEntry('word/document.xml')
    if ($e) { $e.Delete() }
    $ne = $zip.CreateEntry('word/document.xml')
    $sw = New-Object System.IO.StreamWriter($ne.Open(), (New-Object System.Text.UTF8Encoding($false)))
    $sw.Write($xmlStr)
    $sw.Flush(); $sw.Close()

    if ($Title -ne '') {
        $ce = $zip.GetEntry('docProps/core.xml')
        if ($ce) {
            $sr2 = New-Object System.IO.StreamReader($ce.Open(), [System.Text.Encoding]::UTF8)
            $core = $sr2.ReadToEnd(); $sr2.Close()
            if ($core -match '<dc:title>.*?</dc:title>') {
                $core = [regex]::Replace($core, '<dc:title>.*?</dc:title>', '<dc:title>' + (EscX $Title) + '</dc:title>')
                $ce.Delete()
                $nce = $zip.CreateEntry('docProps/core.xml')
                $sw2 = New-Object System.IO.StreamWriter($nce.Open(), (New-Object System.Text.UTF8Encoding($false)))
                $sw2.Write($core); $sw2.Flush(); $sw2.Close()
                Write-Output 'core.xml title updated'
            }
        }
    }
} finally { $zip.Dispose() }
Write-Output ('PACKAGE OK: ' + $Out)

# ---- Word COM validation ----
if (-not $SkipValidation) {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try {
        $doc = $word.Documents.Open($Out, $false, $true)
        $pc = $doc.Paragraphs.Count
        $tc = $doc.Tables.Count
        Write-Output ("WORD OPEN OK. paras=$pc tables=$tc")
        $doc.Close($false)
    } catch {
        Write-Output ('WORD OPEN FAILED: ' + $_.Exception.Message)
    } finally {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}
Write-Output 'BUILD DONE'
