# build_paper_skeleton.ps1 - lunwen skill
# Build an AGU/WRR-style manuscript docx from a content DSL:
#   Letter page, 1-inch margins, continuous line numbers, Times New Roman 12 pt,
#   double-spaced justified body, 14 pt bold title, MathType-style equation slots.
# Content in [[double brackets]] is highlighted yellow so the author can find and
# fill every placeholder. Paragraphs starting with GUIDE are grey italic hints.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File build_paper_skeleton.ps1 `
#     -Out "C:\path\paper_skeleton.docx" [-Content "C:\path\content.txt"] `
#     [-Journal "Water Resources Research"] [-SkipValidation]
#
# Keep this script pure ASCII. All manuscript text lives in the UTF-8 DSL file.

param(
    [Parameter(Mandatory = $true)][string]$Out,
    [string]$Content = '',
    [string]$Journal = 'Water Resources Research',
    [switch]$SkipValidation
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

if ([string]::IsNullOrWhiteSpace($Content)) {
    $Content = Join-Path $PSScriptRoot 'content_skeleton.txt'
}
if (-not (Test-Path -LiteralPath $Content)) {
    Write-Output ('CONTENT_NOT_FOUND: ' + $Content); exit 1
}

$song = [string][char]0x5B8B + [string][char]0x4F53   # SimSun
$tnr = 'Times New Roman'
$textWidth = 9360                                   # 12240 - 2 * 1440

# ---------------------------------------------------------------- helpers ----

function EscX([string]$s) {
    return $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

function FontXml([bool]$small) {
    $f = '<w:rFonts w:ascii="' + $script:tnr + '" w:hAnsi="' + $script:tnr + '" w:eastAsia="' + $script:song + '" w:cs="' + $script:tnr + '" w:hint="eastAsia"/>'
    if ($small) { return $f + '<w:sz w:val="21"/><w:szCs w:val="21"/>' }
    return $f + '<w:sz w:val="24"/><w:szCs w:val="24"/>'
}

function RunXml([string]$t, [bool]$it, [bool]$sbv, [bool]$spv, [bool]$bd, [bool]$hl, [string]$mode) {
    $rpr = FontXml ($mode -eq 'small')
    if ($bd) { $rpr += '<w:b/><w:bCs/>' }
    if ($it) { $rpr += '<w:i/>' }
    if ($hl) { $rpr += '<w:highlight w:val="yellow"/>' }
    if ($sbv) { $rpr += '<w:vertAlign w:val="subscript"/>' }
    if ($spv) { $rpr += '<w:vertAlign w:val="superscript"/>' }
    return '<w:r><w:rPr>' + $rpr + '</w:rPr><w:t xml:space="preserve">' + $t + '</w:t></w:r>'
}

function MarkedRuns([string]$s, [string]$mode) {
    $out = New-Object System.Text.StringBuilder
    $rx = New-Object System.Text.RegularExpressions.Regex('\[\[(.+?)\]\]|\[(/?)(i|sb|sp|b)\]')
    $ms = $rx.Matches($s)
    $it = $false; $sb = $false; $sp = $false; $bd = $false
    $pos = 0
    foreach ($m in $ms) {
        $piece = $s.Substring($pos, $m.Index - $pos)
        if ($piece.Length -gt 0) { [void]$out.Append((RunXml (EscX $piece) $it $sb $sp $bd $false $mode)) }
        if ($m.Groups[1].Success) {
            # Keep the [[ ]] markers visible: they are the fill-in contract and
            # let check_fill.ps1 find every remaining placeholder.
            [void]$out.Append((RunXml (EscX $m.Value) $it $sb $sp $bd $true $mode))
        } else {
            $closing = ($m.Groups[2].Value -ne '')
            switch ($m.Groups[3].Value) {
                'i'  { $it = -not $closing }
                'sb' { $sb = -not $closing }
                'sp' { $sp = -not $closing }
                'b'  { $bd = -not $closing }
            }
        }
        $pos = $m.Index + $m.Length
    }
    $tail = $s.Substring($pos)
    if ($tail.Length -gt 0) { [void]$out.Append((RunXml (EscX $tail) $it $sb $sp $bd $false $mode)) }
    return $out.ToString()
}

function ParaStyled([string]$style, [string]$t, [string]$extraPpr) {
    return '<w:p><w:pPr><w:pStyle w:val="' + $style + '"/>' + $extraPpr + '<w:rPr>' + (FontXml $false) + '</w:rPr></w:pPr>' + (MarkedRuns $t 'body') + '</w:p>'
}

function ParaTitle([string]$t) {
    # 14 pt bold title; the run must carry the same rPr as the paragraph mark,
    # otherwise Word reports and renders it at the Normal size.
    $rp = '<w:rFonts w:ascii="' + $script:tnr + '" w:hAnsi="' + $script:tnr + '" w:eastAsia="' + $script:song + '" w:cs="' + $script:tnr + '" w:hint="eastAsia"/><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/>'
    return '<w:p><w:pPr><w:spacing w:before="0" w:after="120" w:line="360" w:lineRule="auto"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:rPr>' + $rp + '</w:rPr></w:pPr><w:r><w:rPr>' + $rp + '<w:highlight w:val="yellow"/></w:rPr><w:t xml:space="preserve">' + (EscX $t) + '</w:t></w:r></w:p>'
}

function ParaGuide([string]$t) {
    $rp = '<w:rFonts w:ascii="' + $script:tnr + '" w:hAnsi="' + $script:tnr + '" w:eastAsia="' + $script:song + '" w:cs="' + $script:tnr + '" w:hint="eastAsia"/><w:i/><w:color w:val="808080"/><w:sz w:val="21"/><w:szCs w:val="21"/>'
    return '<w:p><w:pPr><w:spacing w:before="60" w:after="120" w:line="360" w:lineRule="auto"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="left"/><w:rPr>' + $rp + '</w:rPr></w:pPr>' + (RunRawLocal $t $rp) + '</w:p>'
}

function RunRawLocal([string]$t, [string]$rpr) {
    return '<w:r><w:rPr>' + $rpr + '</w:rPr><w:t xml:space="preserve">' + (EscX $t) + '</w:t></w:r>'
}

function ParaEq([string]$num, [string]$latex) {
    $mid = ''
    if ($latex -ne '') {
        $rp = '<w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math" w:eastAsia="' + $script:song + '" w:cs="' + $script:tnr + '" w:hint="eastAsia"/><w:i/><w:color w:val="808080"/><w:sz w:val="22"/><w:szCs w:val="22"/>'
        $mid = '<w:r><w:rPr>' + $rp + '</w:rPr><w:t xml:space="preserve">' + (EscX $latex) + '</w:t></w:r>'
    }
    $numRun = '<w:r><w:rPr><w:rFonts w:ascii="' + $script:tnr + '" w:hAnsi="' + $script:tnr + '" w:eastAsia="' + $script:song + '" w:hint="eastAsia"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">' + (EscX $num) + '</w:t></w:r>'
    return '<w:p><w:pPr><w:pStyle w:val="MTDisplayEquation"/><w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:ind w:firstLineChars="0" w:firstLine="0"/><w:rPr><w:rFonts w:ascii="' + $script:tnr + '" w:hAnsi="' + $script:tnr + '" w:eastAsia="' + $script:song + '"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:pPr><w:r><w:tab/></w:r>' + $mid + '<w:r><w:tab/></w:r>' + $numRun + '</w:p>'
}

function ParaFig([string]$cap) {
    $small = FontXml $true
    return '<w:p><w:pPr><w:spacing w:before="120" w:after="0"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="center"/></w:pPr></w:p>' +
           '<w:p><w:pPr><w:pStyle w:val="FigureorTableCaption"/><w:rPr>' + $small + '</w:rPr></w:pPr>' + (MarkedRuns $cap 'small') + '</w:p>'
}

function ParaRef([string]$t) {
    return '<w:p><w:pPr><w:pStyle w:val="Reference"/><w:rPr>' + (FontXml $false) + '</w:rPr></w:pPr>' + (MarkedRuns $t 'body') + '</w:p>'
}

function ParaBreak {
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
}

function EmitTable($rows) {
    $n = $rows.Count
    $cols = 1
    foreach ($r in $rows) {
        $c = ($r -split '##').Count
        if ($c -gt $cols) { $cols = $c }
    }
    $colW = [int]([math]::Floor($script:textWidth / $cols))
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<w:tbl><w:tblPr><w:tblW w:w="' + ($colW * $cols) + '" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>')
    for ($c = 0; $c -lt $cols; $c++) { [void]$sb.Append('<w:gridCol w:w="' + $colW + '"/>') }
    [void]$sb.Append('</w:tblGrid>')
    for ($r = 0; $r -lt $n; $r++) {
        $cells = $rows[$r] -split '##'
        [void]$sb.Append('<w:tr>')
        for ($c = 0; $c -lt $cols; $c++) {
            $txt = ''
            if ($c -lt $cells.Count) { $txt = $cells[$c] }
            [void]$sb.Append('<w:tc><w:tcPr><w:tcW w:w="' + $colW + '" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="left"/><w:rPr>' + (FontXml $true) + '</w:rPr></w:pPr>' + (MarkedRuns $txt 'small') + '</w:p></w:tc>')
        }
        [void]$sb.Append('</w:tr>')
    }
    [void]$sb.Append('</w:tbl>')
    return $sb.ToString()
}

# ------------------------------------------------------------- doc parts ----

$rprBody = '<w:rFonts w:ascii="' + $tnr + '" w:hAnsi="' + $tnr + '" w:eastAsia="' + $song + '" w:cs="' + $tnr + '"/><w:sz w:val="24"/><w:szCs w:val="24"/>'

$STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
'<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="' + $tnr + '" w:hAnsi="' + $tnr + '" w:eastAsia="' + $song + '" w:cs="' + $tnr + '"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="en-US"/></w:rPr></w:rPrDefault>' +
'<w:pPrDefault><w:pPr><w:spacing w:line="480" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>' +
'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
'<w:style w:type="paragraph" w:customStyle="1" w:styleId="Text"><w:name w:val="Text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="120"/><w:ind w:firstLine="720"/><w:jc w:val="both"/></w:pPr><w:rPr>' + $rprBody + '</w:rPr></w:style>' +
'<w:style w:type="paragraph" w:customStyle="1" w:styleId="Heading-Main"><w:name w:val="Heading-Main"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr>' + $rprBody + '<w:b/><w:bCs/></w:rPr></w:style>' +
'<w:style w:type="paragraph" w:customStyle="1" w:styleId="Heading-Secondary"><w:name w:val="Heading-Secondary"/><w:basedOn w:val="Heading-Main"/><w:qFormat/><w:pPr><w:ind w:left="720"/></w:pPr><w:rPr><w:b w:val="0"/><w:bCs w:val="0"/></w:rPr></w:style>' +
'<w:style w:type="paragraph" w:customStyle="1" w:styleId="Authors"><w:name w:val="Authors"/><w:basedOn w:val="Normal"/><w:rPr>' + $rprBody + '<w:b/><w:bCs/></w:rPr></w:style>' +
'<w:style w:type="paragraph" w:customStyle="1" w:styleId="Affiliation"><w:name w:val="Affiliation"/><w:basedOn w:val="Normal"/><w:rPr>' + $rprBody + '</w:rPr></w:style>' +
'<w:style w:type="paragraph" w:customStyle="1" w:styleId="KeyPoints"><w:name w:val="Key Points"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="120"/></w:pPr><w:rPr>' + $rprBody + '</w:rPr></w:style>' +
'<w:style w:type="paragraph" w:customStyle="1" w:styleId="Abstract"><w:name w:val="Abstract"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="120"/><w:ind w:firstLine="480"/></w:pPr><w:rPr>' + $rprBody + '</w:rPr></w:style>' +
'<w:style w:type="paragraph" w:customStyle="1" w:styleId="PLS"><w:name w:val="Plain Language Summary"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="120"/><w:ind w:firstLine="480"/></w:pPr><w:rPr>' + $rprBody + '</w:rPr></w:style>' +
'<w:style w:type="paragraph" w:customStyle="1" w:styleId="Where"><w:name w:val="Where"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="60" w:after="120"/><w:ind w:left="0" w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="both"/></w:pPr><w:rPr>' + $rprBody + '</w:rPr></w:style>' +
'<w:style w:type="paragraph" w:customStyle="1" w:styleId="FigureorTableCaption"><w:name w:val="Figure or Table Caption"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240"/></w:pPr><w:rPr>' + $rprBody + '</w:rPr></w:style>' +
'<w:style w:type="paragraph" w:customStyle="1" w:styleId="Reference"><w:name w:val="Reference"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="120"/><w:ind w:left="720" w:hanging="720"/><w:jc w:val="left"/></w:pPr><w:rPr>' + $rprBody + '</w:rPr></w:style>' +
'<w:style w:type="paragraph" w:customStyle="1" w:styleId="MTDisplayEquation"><w:name w:val="MTDisplayEquation"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:ind w:firstLine="0"/></w:pPr><w:rPr>' + $rprBody + '</w:rPr></w:style>' +
'<w:style w:type="paragraph" w:customStyle="1" w:styleId="Guide"><w:name w:val="Guide"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="60" w:after="120"/><w:ind w:firstLine="0"/><w:jc w:val="left"/></w:pPr><w:rPr>' + $rprBody + '<w:i/><w:color w:val="808080"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:style>' +
'</w:styles>'

$CONTENT_TYPES = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>
'@

$ROOT_RELS = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>
'@

$DOC_RELS = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>
'@

$SETTINGS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
'<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'

function HeaderXml([string]$journal) {
    $rp = '<w:rFonts w:ascii="' + $tnr + '" w:hAnsi="' + $tnr + '" w:eastAsia="' + $song + '" w:cs="' + $tnr + '"/><w:sz w:val="20"/><w:szCs w:val="20"/>'
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="center"/><w:rPr>' + $rp + '</w:rPr></w:pPr>' +
    '<w:r><w:rPr>' + $rp + '</w:rPr><w:t xml:space="preserve">manuscript submitted to ' + (EscX $journal) + '</w:t></w:r></w:p></w:hdr>'
}

function CoreXml([string]$title) {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    '<dc:title>' + (EscX $title) + '</dc:title><dc:creator>lunwen</dc:creator><cp:lastModifiedBy>lunwen</cp:lastModifiedBy></cp:coreProperties>'
}

# --------------------------------------------------------------- parsing ----

$docs = New-Object System.Text.StringBuilder
$rows = New-Object System.Collections.Generic.List[string]
$eqCount = 0
$tblCount = 0
$keyCount = 0
$abstractText = ''
$plsText = ''
$guideCount = 0
$curSection = ''
$titleText = ''

function FlushTable {
    if ($script:rows.Count -gt 0) {
        [void]$script:docs.Append((EmitTable $script:rows))
        $script:rows.Clear()
        $script:tblCount++
    }
}

$lines = [System.IO.File]::ReadAllLines($Content, [System.Text.Encoding]::UTF8)
foreach ($line in $lines) {
    if ($line.Trim().Length -eq 0) { continue }
    $sep = $line.IndexOf('|')
    if ($sep -lt 1) { continue }
    $tag = $line.Substring(0, $sep).Trim().ToUpper()
    $val = $line.Substring($sep + 1)
    if (($tag -ne 'TBLH') -and ($tag -ne 'TBLR')) { FlushTable }
    switch ($tag) {
        'TITLE' { $titleText = [regex]::Replace($val, '\[\[(.+?)\]\]', '$1'); [void]$docs.Append((ParaTitle $val)) }
        'AUTH'  { [void]$docs.Append((ParaStyled 'Authors' $val '')) }
        'AFF'   { [void]$docs.Append((ParaStyled 'Affiliation' $val '')) }
        'CORR'  { [void]$docs.Append((ParaStyled 'Affiliation' $val '')) }
        'H1'    { $curSection = $val; [void]$docs.Append((ParaStyled 'Heading-Main' $val '')) }
        'H2'    { [void]$docs.Append((ParaStyled 'Heading-Secondary' $val '')) }
        'B'     { [void]$docs.Append((ParaStyled 'Text' $val '')) }
        'BQ'    { [void]$docs.Append((ParaStyled 'Normal' $val '')) }
        'W'     {
            # Symbol definitions after an equation: flush left, lowercase "where".
            $t = [regex]::Replace($val, '^\s*where\b', 'where', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            [void]$docs.Append((ParaStyled 'Where' $t ''))
        }
        'KEY'   { $keyCount++; [void]$docs.Append((ParaStyled 'KeyPoints' $val '')) }
        'ABS'   { $abstractText += ' ' + [regex]::Replace($val, '\[\[(.+?)\]\]', '$1'); [void]$docs.Append((ParaStyled 'Abstract' $val '')) }
        'PLS'   { $plsText += ' ' + [regex]::Replace($val, '\[\[(.+?)\]\]', '$1'); [void]$docs.Append((ParaStyled 'PLS' $val '')) }
        'NOTE'  { $guideCount++; [void]$docs.Append((ParaGuide $val)) }
        'EQ'    {
            $num = ''; $latex = ''
            $s2 = $val.IndexOf('|')
            if ($s2 -ge 0) { $num = $val.Substring(0, $s2).Trim(); $latex = $val.Substring($s2 + 1) }
            else { $latex = $val }
            [void]$docs.Append((ParaEq $num $latex)); $eqCount++
        }
        'FIG'   { [void]$docs.Append((ParaFig $val)) }
        'TCAP'  { [void]$docs.Append((ParaStyled 'FigureorTableCaption' $val '')) }
        'TBLH'  { if ($script:rows.Count -gt 0) { FlushTable }; $script:rows.Add($val) }
        'TBLR'  { $script:rows.Add($val) }
        'REF'   { [void]$docs.Append((ParaRef $val)) }
        'PAGEBREAK' { [void]$docs.Append((ParaBreak)) }
        default { }
    }
}
FlushTable

# ------------------------------------------------------------- assemble ----

$sectPr = '<w:sectPr><w:headerReference w:type="default" r:id="rId3"/><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="432" w:footer="720" w:gutter="0"/><w:lnNumType w:countBy="1" w:restart="continuous"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>'
$rootTag = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
$xmlStr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + $rootTag + '<w:body>' + $docs.ToString() + $sectPr + '</w:body></w:document>'

try { $chk = [xml]$xmlStr } catch {
    Write-Output ('XML_BUILD_FAILED: ' + $_.Exception.Message); exit 1
}

# word / character checks
$absWords = 0
if ($abstractText.Trim().Length -gt 0) { $absWords = ($abstractText.Trim() -split '\s+').Count }
$plsWords = 0
if ($plsText.Trim().Length -gt 0) { $plsWords = ($plsText.Trim() -split '\s+').Count }
Write-Output ('COUNTS: equations=' + $eqCount + ' tables=' + $tblCount + ' keyPoints=' + $keyCount + ' guides=' + $guideCount)
if ($absWords -gt 0) { Write-Output ('ABSTRACT WORDS: ' + $absWords + ' (AGU limit 250)') }
if ($plsWords -gt 0) { Write-Output ('PLS WORDS: ' + $plsWords + ' (AGU 100-200)') }

# --------------------------------------------------------------- package ----

# Word COM resolves relative paths against its own working directory, so make
# the output path absolute before packing and validating.
if (-not [System.IO.Path]::IsPathRooted($Out)) {
    $Out = Join-Path (Get-Location).Path $Out
}
$Out = [System.IO.Path]::GetFullPath($Out)

$outDir = Split-Path -Parent $Out
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}
if (Test-Path -LiteralPath $Out) { Remove-Item -LiteralPath $Out -Force }

$zip = [System.IO.Compression.ZipFile]::Open($Out, 'Create')
function AddEntry([string]$name, [string]$text) {
    $e = $script:zip.CreateEntry($name)
    $sw = New-Object System.IO.StreamWriter($e.Open(), (New-Object System.Text.UTF8Encoding($false)))
    $sw.Write($text); $sw.Flush(); $sw.Close()
}
try {
    AddEntry '[Content_Types].xml' $CONTENT_TYPES
    AddEntry '_rels/.rels' $ROOT_RELS
    AddEntry 'docProps/core.xml' (CoreXml $titleText)
    AddEntry 'word/document.xml' $xmlStr
    AddEntry 'word/styles.xml' $STYLES
    AddEntry 'word/settings.xml' $SETTINGS
    AddEntry 'word/header1.xml' (HeaderXml $Journal)
    AddEntry 'word/_rels/document.xml.rels' $DOC_RELS
} finally { $zip.Dispose() }
Write-Output ('PACKAGE OK: ' + $Out)

if (-not $SkipValidation) {
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        try {
            $doc = $word.Documents.Open($Out, $false, $true)
            Write-Output ('WORD OPEN OK: paras=' + $doc.Paragraphs.Count + ' tables=' + $doc.Tables.Count)
            $doc.Close($false)
        } finally {
            $word.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
    } catch {
        Write-Output ('WORD VALIDATION SKIPPED/FAILED: ' + $_.Exception.Message)
    }
}
Write-Output 'BUILD DONE'
