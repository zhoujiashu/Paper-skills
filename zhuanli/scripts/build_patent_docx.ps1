# build_patent_docx.ps1 - zhuanli skill
# Build CN patent application documents (specification / claims / abstract)
# by template-package replacement:
#   copy the reference docx -> rewrite word/document.xml from a content DSL file
#   -> repack -> optionally validate by opening in Word.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File build_patent_docx.ps1 `
#     -Template "C:\path\ref.docx" -Content "C:\path\content.txt" `
#     -OutDir "C:\path\out" -BaseName "CN2024xxxxxx" [-SingleFile] [-SkipValidation]
#
# DSL: one document per "DOC|name" block. Each block becomes one docx
# (or all blocks are merged into one file with -SingleFile).
#
# IMPORTANT: keep this script pure ASCII. Put all Chinese text in the content
# DSL file (UTF-8) or pass it through -BaseName on the command line.

param(
    [string]$Template = '',
    [Parameter(Mandatory = $true)][string]$Content,
    [Parameter(Mandatory = $true)][string]$OutDir,
    [string]$BaseName = 'patent',
    [switch]$SingleFile,
    [switch]$SkipValidation
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

# No -Template means: build a CNIPA standard-layout package from scratch.
$standard = [string]::IsNullOrWhiteSpace($Template)

$song = [string][char]0x5B8B + [string][char]0x4F53   # SimSun
$asciiFont = 'Times New Roman'
$totalWidth = 8300

# ---------------------------------------------------------------- helpers ----

function EscX([string]$s) {
    return $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

function FontXml([bool]$small) {
    $f = '<w:rFonts w:ascii="' + $script:asciiFont + '" w:hAnsi="' + $script:asciiFont + '" w:eastAsia="' + $script:song + '" w:cs="' + $script:asciiFont + '" w:hint="eastAsia"/>'
    if ($small) { return $f + '<w:sz w:val="21"/><w:szCs w:val="21"/>' }
    return $f + '<w:sz w:val="24"/><w:szCs w:val="24"/>'
}

function RunRaw([string]$t, [string]$rpr) {
    return '<w:r><w:rPr>' + $rpr + '</w:rPr><w:t xml:space="preserve">' + (EscX $t) + '</w:t></w:r>'
}

function RunXml([string]$t, [bool]$it, [bool]$sbv, [bool]$spv, [bool]$bd, [string]$mode) {
    $rpr = FontXml ($mode -eq 'small')
    if ($bd) { $rpr += '<w:b/><w:bCs/>' }
    if ($it) { $rpr += '<w:i/>' }
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

# ------------------------------------------------------------- paragraphs ----

function ParaTitle([string]$t) {
    # 12 pt (sz24) keeps every glyph within the 3.5-4.5 mm CNIPA limit.
    $rp = '<w:rFonts w:ascii="' + $script:asciiFont + '" w:hAnsi="' + $script:asciiFont + '" w:eastAsia="' + $script:song + '" w:cs="' + $script:asciiFont + '" w:hint="eastAsia"/>' +
          '<w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/>'
    return '<w:p><w:pPr><w:spacing w:before="0" w:after="240" w:line="360" w:lineRule="auto"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="center"/><w:rPr>' + $rp + '</w:rPr></w:pPr>' + (RunRaw $t $rp) + '</w:p>'
}

function ParaSec([string]$t) {
    $rp = FontXml $false + '<w:b/><w:bCs/>'
    return '<w:p><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="120" w:after="120" w:line="360" w:lineRule="auto"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:rPr>' + $rp + '</w:rPr></w:pPr>' + (RunRaw $t $rp) + '</w:p>'
}

function ParaBody([string]$t, [bool]$indent) {
    $ind = '<w:ind w:firstLine="480"/>'
    if (-not $indent) { $ind = '<w:ind w:left="0" w:firstLineChars="0" w:firstLine="0"/>' }
    return '<w:p><w:pPr><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:spacing w:line="360" w:lineRule="auto"/><w:jc w:val="both"/>' + $ind + '<w:rPr>' + (FontXml $false) + '</w:rPr></w:pPr>' + (MarkedRuns $t 'body') + '</w:p>'
}

function ParaClaim([string]$t) {
    return '<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/><w:ind w:left="0" w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="both"/><w:rPr>' + (FontXml $false) + '</w:rPr></w:pPr>' + (MarkedRuns $t 'body') + '</w:p>'
}

function ParaEqn([string]$num, [string]$latex) {
    if ($num -eq '' -and $latex -eq '') { return '' }
    $mid = ''
    if ($latex -ne '') {
        $rp = '<w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math" w:eastAsia="' + $script:song + '" w:cs="' + $script:asciiFont + '" w:hint="eastAsia"/><w:i/><w:color w:val="808080"/><w:sz w:val="20"/><w:szCs w:val="20"/>'
        $mid = RunRaw $latex $rp
    }
    if ($num -eq '') {
        return '<w:p><w:pPr><w:spacing w:before="60" w:after="60" w:line="360" w:lineRule="auto"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="center"/></w:pPr>' + $mid + '</w:p>'
    }
    $numRun = RunRaw $num ('<w:rFonts w:ascii="' + $script:asciiFont + '" w:hAnsi="' + $script:asciiFont + '" w:eastAsia="' + $script:song + '" w:hint="eastAsia"/><w:sz w:val="24"/><w:szCs w:val="24"/>')
    return '<w:p><w:pPr><w:spacing w:before="60" w:after="60" w:line="360" w:lineRule="auto"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:tabs><w:tab w:val="center" w:pos="4150"/><w:tab w:val="right" w:pos="8300"/></w:tabs></w:pPr><w:r><w:tab/></w:r>' + $mid + '<w:r><w:tab/></w:r>' + $numRun + '</w:p>'
}

function ParaFig([string]$cap) {
    $small = FontXml $true
    return '<w:p><w:pPr><w:spacing w:before="120" w:after="0"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="center"/></w:pPr></w:p>' +
           '<w:p><w:pPr><w:spacing w:after="120" w:line="360" w:lineRule="auto"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="center"/><w:rPr>' + $small + '</w:rPr></w:pPr>' + (MarkedRuns $cap 'small') + '</w:p>'
}

function ParaTcap([string]$cap) {
    $small = FontXml $true
    return '<w:p><w:pPr><w:keepNext/><w:spacing w:before="120" w:after="60"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="center"/><w:rPr>' + $small + '</w:rPr></w:pPr>' + (MarkedRuns $cap 'small') + '</w:p>'
}

function ParaRef([string]$t) {
    $rp = FontXml $true
    return '<w:p><w:pPr><w:spacing w:line="300" w:lineRule="auto"/><w:ind w:left="0" w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="left"/><w:rPr>' + $rp + '</w:rPr></w:pPr>' + (RunRaw $t $rp) + '</w:p>'
}

function ParaBreak {
    return '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>'
}

function EmitTable($rows) {
    $n = $rows.Count
    $cols = 1
    foreach ($r in $rows) {
        $c = ($r -split '##').Count
        if ($c -gt $cols) { $cols = $c }
    }
    $colW = [int]([math]::Floor($script:totalWidth / $cols))
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<w:tbl><w:tblPr><w:tblW w:w="' + ($colW * $cols) + '" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>')
    for ($c = 0; $c -lt $cols; $c++) { [void]$sb.Append('<w:gridCol w:w="' + $colW + '"/>') }
    [void]$sb.Append('</w:tblGrid>')
    for ($r = 0; $r -lt $n; $r++) {
        $cells = $rows[$r] -split '##'
        [void]$sb.Append('<w:tr><w:trPr><w:jc w:val="center"/></w:trPr>')
        for ($c = 0; $c -lt $cols; $c++) {
            $txt = ''
            if ($c -lt $cells.Count) { $txt = $cells[$c] }
            [void]$sb.Append('<w:tc><w:tcPr><w:tcW w:w="' + $colW + '" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:line="300" w:lineRule="auto"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:jc w:val="center"/><w:rPr>' + (FontXml $true) + '</w:rPr></w:pPr>' + (MarkedRuns $txt 'small') + '</w:p></w:tc>')
        }
        [void]$sb.Append('</w:tr>')
    }
    [void]$sb.Append('</w:tbl>')
    return $sb.ToString()
}

function SanitizeName([string]$s) {
    $bad = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($ch in $bad) { $s = $s.Replace([string]$ch, '_') }
    $s = $s.Trim()
    if ($s -eq '') { $s = 'doc' }
    return $s
}

# ------------------------------------------------- template or standard ----

if ($standard) {
    # CNIPA submission layout: A4, 25/15/25/15 mm margins, centered page number.
    $rootTag = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    $sectPr = '<w:sectPr><w:footerReference w:type="default" r:id="rId3"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="850" w:bottom="850" w:left="1417" w:header="709" w:footer="567" w:gutter="0"/><w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>'
    Write-Output 'MODE: standard CNIPA layout (A4, 25/15/25/15 mm, SimSun 12pt, 1.5 line, centered page numbers)'
} else {
    $zipR = [System.IO.Compression.ZipFile]::OpenRead($Template)
    try {
        $entry = $zipR.GetEntry('word/document.xml')
        if (-not $entry) { Write-Output 'TEMPLATE_INVALID: word/document.xml not found'; exit 1 }
        $sr = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
        $raw = $sr.ReadToEnd(); $sr.Close()
    } finally { $zipR.Dispose() }

    $i0 = $raw.IndexOf('<w:document')
    $i1 = $raw.IndexOf('>', $i0)
    if ($i0 -lt 0 -or $i1 -lt 0) { Write-Output 'TEMPLATE_INVALID: <w:document> tag not found'; exit 1 }
    $rootTag = $raw.Substring($i0, $i1 - $i0 + 1)

    $xmlT = [xml]$raw
    # Build inline: an XmlNamespaceManager returned from a function would be
    # unrolled into its namespace-prefix strings (it is IEnumerable in .NET FW).
    $ns = [System.Xml.XmlNamespaceManager]::new($xmlT.NameTable)
    $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $sectNode = $xmlT.SelectSingleNode('//w:body/w:sectPr', $ns)
    if ($sectNode) { $sectPr = $sectNode.OuterXml } else { $sectPr = '' }
    Write-Output ('MODE: template layout (' + $Template + ')')
}

# ---------------------------------------------------------- content parser ----

$script:docs = New-Object System.Collections.ArrayList
$script:rows = New-Object System.Collections.Generic.List[string]
$script:curDoc = $null

function StartDoc([string]$name) {
    $d = [PSCustomObject]@{
        Name  = $name
        Body  = (New-Object System.Text.StringBuilder)
        Eqs   = 0
        Tbls  = 0
        Chars = 0
    }
    [void]$script:docs.Add($d)
    $script:curDoc = $d
}

function FlushTable {
    if ($script:rows.Count -gt 0 -and $null -ne $script:curDoc) {
        [void]$script:curDoc.Body.Append((EmitTable $script:rows))
        $script:rows.Clear()
        $script:curDoc.Tbls++
    }
}

function AddChars([string]$s) {
    if ($null -eq $script:curDoc) { return }
    $plain = [regex]::Replace($s, '\[/?[a-z]+\]', '')
    $script:curDoc.Chars += $plain.Length
}

StartDoc 'Document'

$lines = [System.IO.File]::ReadAllLines($Content, [System.Text.Encoding]::UTF8)
foreach ($line in $lines) {
    if ($line.Trim().Length -eq 0) { continue }
    $sep = $line.IndexOf('|')
    if ($sep -lt 1) { continue }
    $tag = $line.Substring(0, $sep).Trim().ToUpper()
    $val = $line.Substring($sep + 1)
    if (($tag -ne 'TBLH') -and ($tag -ne 'TBLR')) { FlushTable }
    switch ($tag) {
        'DOC' {
            $newName = $val.Trim()
            if ($script:docs.Count -eq 1 -and $script:curDoc.Body.Length -eq 0) {
                $script:curDoc.Name = $newName
            } else {
                StartDoc $newName
            }
        }
        'TITLE'     { [void]$script:curDoc.Body.Append((ParaTitle $val)) }
        'SEC'       { [void]$script:curDoc.Body.Append((ParaSec $val)) }
        'REFH'      { [void]$script:curDoc.Body.Append((ParaSec $val)) }
        'B'         { [void]$script:curDoc.Body.Append((ParaBody $val $true)); AddChars $val }
        'BQ'        { [void]$script:curDoc.Body.Append((ParaBody $val $false)); AddChars $val }
        'CLAIM'     { [void]$script:curDoc.Body.Append((ParaClaim $val)); AddChars $val }
        'EQN'       {
            $num = ''; $latex = ''
            $s2 = $val.IndexOf('|')
            if ($s2 -ge 0) {
                $num = $val.Substring(0, $s2).Trim()
                $latex = $val.Substring($s2 + 1)
            } else {
                $latex = $val
            }
            [void]$script:curDoc.Body.Append((ParaEqn $num $latex))
            $script:curDoc.Eqs++
        }
        'FIG'       { [void]$script:curDoc.Body.Append((ParaFig $val)) }
        'TCAP'      { [void]$script:curDoc.Body.Append((ParaTcap $val)) }
        'TBLH'      { if ($script:rows.Count -gt 0) { FlushTable }; $script:rows.Add($val) }
        'TBLR'      { $script:rows.Add($val) }
        'PAGEBREAK' { [void]$script:curDoc.Body.Append((ParaBreak)) }
        'REF'       { [void]$script:curDoc.Body.Append((ParaRef $val)) }
        default     { }
    }
}
FlushTable

# CNIPA abstract: no more than 300 characters.
$zhaiyao = [string][char]0x6458 + [string][char]0x8981
foreach ($d in $script:docs) {
    if ($d.Name.Contains($zhaiyao)) {
        if ($d.Chars -gt 300) {
            Write-Output ('WARNING: abstract ' + $d.Chars + ' chars exceeds the 300-char limit')
        } else {
            Write-Output ('ABSTRACT OK: ' + $d.Chars + ' chars (limit 300)')
        }
    }
}

# ------------------------------------------------------------ build output ----

if (-not (Test-Path -LiteralPath $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}
# Word COM resolves relative paths against its own working directory, so make
# every output path absolute before packing and validating.
$OutDir = (Resolve-Path -LiteralPath $OutDir).Path

function BuildXml([string]$body) {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + $rootTag + '<w:body>' + $body + $script:sectPr + '</w:body></w:document>'
}

function WritePackage([string]$outPath, [string]$xmlStr, [string]$title) {
    try {
        Copy-Item -LiteralPath $script:Template -Destination $outPath -Force
    } catch {
        throw ('OUTPUT_LOCKED_OR_UNWRITABLE: ' + $_.Exception.Message)
    }
    $zip = [System.IO.Compression.ZipFile]::Open($outPath, 'Update')
    try {
        $e = $zip.GetEntry('word/document.xml')
        if ($e) { $e.Delete() }
        $ne = $zip.CreateEntry('word/document.xml')
        $sw = New-Object System.IO.StreamWriter($ne.Open(), (New-Object System.Text.UTF8Encoding($false)))
        $sw.Write($xmlStr); $sw.Flush(); $sw.Close()

        if ($title -ne '') {
            $ce = $zip.GetEntry('docProps/core.xml')
            if ($ce) {
                $sr2 = New-Object System.IO.StreamReader($ce.Open(), [System.Text.Encoding]::UTF8)
                $core = $sr2.ReadToEnd(); $sr2.Close()
                if ($core -match '<dc:title>.*?</dc:title>') {
                    $core = [regex]::Replace($core, '<dc:title>.*?</dc:title>', '<dc:title>' + (EscX $title) + '</dc:title>')
                    $ce.Delete()
                    $nce = $zip.CreateEntry('docProps/core.xml')
                    $sw2 = New-Object System.IO.StreamWriter($nce.Open(), (New-Object System.Text.UTF8Encoding($false)))
                    $sw2.Write($core); $sw2.Flush(); $sw2.Close()
                }
            }
        }
    } finally { $zip.Dispose() }
}

# ------------------------------------------- standard package (no template) ----

$STD_CONTENT_TYPES = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>
'@

$STD_ROOT_RELS = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>
'@

$STD_DOC_RELS = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>
'@

$STD_STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
'<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="' + $song + '" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/></w:rPr></w:rPrDefault>' +
'<w:pPrDefault><w:pPr><w:spacing w:line="360" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>' +
'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
'</w:styles>'

$STD_SETTINGS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
'<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
'<w:zoom w:percent="100"/><w:defaultTabStop w:val="420"/>' +
'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' +
'</w:settings>'

$footerRpr = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="' + $song + '" w:cs="Times New Roman"/><w:sz w:val="21"/><w:szCs w:val="21"/>'
$STD_FOOTER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
'<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
'<w:p><w:pPr><w:jc w:val="center"/><w:rPr>' + $footerRpr + '</w:rPr></w:pPr>' +
'<w:r><w:rPr>' + $footerRpr + '</w:rPr><w:fldChar w:fldCharType="begin"/></w:r>' +
'<w:r><w:rPr>' + $footerRpr + '</w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
'<w:r><w:rPr>' + $footerRpr + '</w:rPr><w:fldChar w:fldCharType="separate"/></w:r>' +
'<w:r><w:rPr>' + $footerRpr + '</w:rPr><w:t>1</w:t></w:r>' +
'<w:r><w:rPr>' + $footerRpr + '</w:rPr><w:fldChar w:fldCharType="end"/></w:r>' +
'</w:p></w:ftr>'

function CoreXml([string]$title) {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    '<dc:title>' + (EscX $title) + '</dc:title><dc:creator>zhuanli</dc:creator><cp:lastModifiedBy>zhuanli</cp:lastModifiedBy>' +
    '</cp:coreProperties>'
}

function AddZipEntry($zip, [string]$name, [string]$text) {
    $e = $zip.CreateEntry($name)
    $sw = New-Object System.IO.StreamWriter($e.Open(), (New-Object System.Text.UTF8Encoding($false)))
    $sw.Write($text); $sw.Flush(); $sw.Close()
}

function WriteStandardPackage([string]$outPath, [string]$xmlStr, [string]$title) {
    if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }
    $zip = [System.IO.Compression.ZipFile]::Open($outPath, 'Create')
    try {
        AddZipEntry $zip '[Content_Types].xml' $script:STD_CONTENT_TYPES
        AddZipEntry $zip '_rels/.rels' $script:STD_ROOT_RELS
        AddZipEntry $zip 'docProps/core.xml' (CoreXml $title)
        AddZipEntry $zip 'word/document.xml' $xmlStr
        AddZipEntry $zip 'word/styles.xml' $script:STD_STYLES
        AddZipEntry $zip 'word/settings.xml' $script:STD_SETTINGS
        AddZipEntry $zip 'word/footer1.xml' $script:STD_FOOTER
        AddZipEntry $zip 'word/_rels/document.xml.rels' $script:STD_DOC_RELS
    } finally { $zip.Dispose() }
}

$outputs = New-Object System.Collections.Generic.List[string]
$safeBase = SanitizeName $BaseName

if ($SingleFile) {
    $merged = New-Object System.Text.StringBuilder
    for ($k = 0; $k -lt $script:docs.Count; $k++) {
        if ($k -gt 0) { [void]$merged.Append((ParaBreak)) }
        [void]$merged.Append($script:docs[$k].Body.ToString())
    }
    $xmlStr = BuildXml $merged.ToString()
    try { $chk = [xml]$xmlStr } catch {
        Write-Output ('XML_BUILD_FAILED: ' + $_.Exception.Message); exit 1
    }
    $outPath = Join-Path $OutDir ($safeBase + '.docx')
    if ($standard) { WriteStandardPackage $outPath $xmlStr $safeBase }
    else { WritePackage $outPath $xmlStr $safeBase }
    $outputs.Add($outPath)
    $eqTotal = 0; $tblTotal = 0
    foreach ($d in $script:docs) { $eqTotal += $d.Eqs; $tblTotal += $d.Tbls }
    Write-Output ('XML OK. docs=' + $script:docs.Count + ' equations=' + $eqTotal + ' tables=' + $tblTotal + ' -> ' + $outPath)
} else {
    foreach ($d in $script:docs) {
        $docName = SanitizeName $d.Name
        if ($script:docs.Count -eq 1) { $fileName = $safeBase + '.docx' }
        else { $fileName = $safeBase + '_' + $docName + '.docx' }
        $xmlStr = BuildXml $d.Body.ToString()
        try { $chk = [xml]$xmlStr } catch {
            Write-Output ('XML_BUILD_FAILED [' + $d.Name + ']: ' + $_.Exception.Message); exit 1
        }
        $outPath = Join-Path $OutDir $fileName
        if ($standard) { WriteStandardPackage $outPath $xmlStr ($safeBase + ' ' + $d.Name) }
        else { WritePackage $outPath $xmlStr ($safeBase + ' ' + $d.Name) }
        $outputs.Add($outPath)
        Write-Output ('XML OK [' + $d.Name + '] equations=' + $d.Eqs + ' tables=' + $d.Tbls + ' -> ' + $outPath)
    }
}

# ------------------------------------------------------- Word validation ----

if (-not $SkipValidation) {
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        try {
            foreach ($o in $outputs) {
                $doc = $word.Documents.Open($o, $false, $true)
                Write-Output ('WORD OPEN OK: ' + (Split-Path $o -Leaf) + ' paras=' + $doc.Paragraphs.Count + ' tables=' + $doc.Tables.Count)
                $doc.Close($false)
            }
        } finally {
            $word.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
    } catch {
        Write-Output ('WORD VALIDATION SKIPPED/FAILED: ' + $_.Exception.Message)
    }
}

Write-Output ('BUILD DONE. files=' + $outputs.Count)
