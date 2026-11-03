# inspect_docx.ps1 - zhuanli skill
# Dump the structural facts needed to reproduce a patent docx layout:
# package entries, document root tag, sectPr, docDefaults, styles,
# numbering ids, and body style usage.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File inspect_docx.ps1 -Path "C:\path\ref.docx"
#
# IMPORTANT: keep this script pure ASCII. Output is tab separated.

param(
    [Parameter(Mandatory = $true)][string]$Path
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch { }

$WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

$script:zip = [System.IO.Compression.ZipFile]::OpenRead($Path)

function ReadEntry([string]$name) {
    $e = $script:zip.GetEntry($name)
    if (-not $e) { return $null }
    $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
    try { return $sr.ReadToEnd() } finally { $sr.Close() }
}

# NOTE: build every namespace manager inline, never return one from a helper:
# in .NET Framework it is IEnumerable, so PowerShell unrolls it into prefix
# strings on function return.

function Attr($node, [string]$local) {
    if ($null -eq $node) { return '' }
    return $node.GetAttribute($local, $WNS)
}

try {
    Write-Output '== PACKAGE ENTRIES =='
    foreach ($e in $script:zip.Entries) { Write-Output $e.FullName }

    $docXml = ReadEntry 'word/document.xml'
    if (-not $docXml) { Write-Output 'ERROR: word/document.xml not found'; exit 1 }

    $i0 = $docXml.IndexOf('<w:document')
    $i1 = $docXml.IndexOf('>', $i0)
    Write-Output ''
    Write-Output '== ROOT TAG =='
    if ($i0 -ge 0 -and $i1 -gt $i0) { Write-Output $docXml.Substring($i0, $i1 - $i0 + 1) }

    $xmlT = [xml]$docXml
    $ns = [System.Xml.XmlNamespaceManager]::new($xmlT.NameTable)
    $ns.AddNamespace('w', $WNS)

    $sect = $xmlT.SelectSingleNode('//w:body/w:sectPr', $ns)
    Write-Output ''
    Write-Output '== SECTPR =='
    if ($sect) { Write-Output $sect.OuterXml } else { Write-Output 'none' }

    $paras = $xmlT.SelectNodes('//w:body//w:p', $ns)
    $tbls = $xmlT.SelectNodes('//w:body//w:tbl', $ns)
    Write-Output ''
    Write-Output '== COUNTS =='
    Write-Output ("paragraphs`t" + $paras.Count)
    Write-Output ("tables`t" + $tbls.Count)
    Write-Output ("sectPr`t" + $xmlT.SelectNodes('//w:sectPr', $ns).Count)

    $stylesXml = ReadEntry 'word/styles.xml'
    if ($stylesXml) {
        $st = [xml]$stylesXml
        $ns2 = [System.Xml.XmlNamespaceManager]::new($st.NameTable)
        $ns2.AddNamespace('w', $WNS)

        Write-Output ''
        Write-Output '== DOC DEFAULTS (outer xml) =='
        $dd = $st.SelectSingleNode('//w:docDefaults', $ns2)
        if ($dd) { Write-Output $dd.OuterXml } else { Write-Output 'none' }

        Write-Output ''
        Write-Output "== STYLES: id`ttype`tname`tdefault`tbasedOn`tascii`teastAsia`tsz =="
        foreach ($s in $st.SelectNodes('//w:style', $ns2)) {
            $id = Attr $s 'styleId'
            $type = Attr $s 'type'
            $def = Attr $s 'default'
            $name = Attr ($s.SelectSingleNode('w:name', $ns2)) 'val'
            $based = Attr ($s.SelectSingleNode('w:basedOn', $ns2)) 'val'
            $ascii = ''; $ea = ''; $sz = ''
            $rpr = $s.SelectSingleNode('w:rPr', $ns2)
            if ($rpr) {
                $rf = $rpr.SelectSingleNode('w:rFonts', $ns2)
                if ($rf) {
                    $ascii = Attr $rf 'ascii'
                    $ea = Attr $rf 'eastAsia'
                }
                $sz = Attr ($rpr.SelectSingleNode('w:sz', $ns2)) 'val'
            }
            Write-Output ($id + "`t" + $type + "`t" + $name + "`t" + $def + "`t" + $based + "`t" + $ascii + "`t" + $ea + "`t" + $sz)
        }

        Write-Output ''
        Write-Output "== BODY STYLE USAGE: pStyleId`tcount =="
        $uses = @{}
        foreach ($p in $xmlT.SelectNodes('//w:body//w:pPr/w:pStyle', $ns)) {
            $v = Attr $p 'val'
            if ($uses.ContainsKey($v)) { $uses[$v]++ } else { $uses[$v] = 1 }
        }
        foreach ($k in ($uses.Keys | Sort-Object)) { Write-Output ($k + "`t" + $uses[$k]) }
    } else {
        Write-Output ''
        Write-Output '== STYLES =='
        Write-Output 'word/styles.xml not found'
    }

    $numXml = ReadEntry 'word/numbering.xml'
    if ($numXml) {
        $nb = [xml]$numXml
        $ns3 = [System.Xml.XmlNamespaceManager]::new($nb.NameTable)
        $ns3.AddNamespace('w', $WNS)
        Write-Output ''
        Write-Output "== NUMBERING: numId`tabstractNumId`tlvl0Fmt`tlvl0Text =="
        foreach ($n in $nb.SelectNodes('//w:num', $ns3)) {
            $nid = Attr $n 'numId'
            $abs = Attr ($n.SelectSingleNode('w:abstractNumId', $ns3)) 'val'
            $fmt = ''; $txt = ''
            if ($abs -ne '') {
                $absNode = $nb.SelectSingleNode("//w:abstractNum[@w:abstractNumId='" + $abs + "']", $ns3)
                if ($absNode) {
                    $lvl0 = $absNode.SelectSingleNode("w:lvl[@w:ilvl='0']", $ns3)
                    if ($lvl0) {
                        $fmt = Attr ($lvl0.SelectSingleNode('w:numFmt', $ns3)) 'val'
                        $txt = Attr ($lvl0.SelectSingleNode('w:lvlText', $ns3)) 'val'
                    }
                }
            }
            Write-Output ($nid + "`t" + $abs + "`t" + $fmt + "`t" + $txt)
        }
    }
}
finally {
    $script:zip.Dispose()
}
