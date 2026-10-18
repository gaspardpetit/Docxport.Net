$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression

$probeDirectory = Join-Path $PSScriptRoot '../fixtures/docx/probes'
$utf8 = [System.Text.UTF8Encoding]::new($false)

function Write-Probe($name, $documentXml, $extraParts) {
    $path = Join-Path $probeDirectory $name
    $stream = [System.IO.File]::Create($path)
    try {
        $archive = [System.IO.Compression.ZipArchive]::new($stream, 'Create', $true)
        try {
            $parts = @{
                '[Content_Types].xml' = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="html" ContentType="text/html; charset=windows-1252"/><Default Extension="txt" ContentType="text/plain"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'
                '_rels/.rels' = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
                'word/document.xml' = $documentXml
            }
            foreach ($part in $extraParts.GetEnumerator()) { $parts[$part.Key] = $part.Value }
            foreach ($part in ($parts.GetEnumerator() | Sort-Object Key)) {
                $entry = $archive.CreateEntry($part.Key)
                $encoding = if ($part.Key -eq 'word/chunk.html') { [System.Text.Encoding]::Latin1 } else { $utf8 }
                $writer = [System.IO.StreamWriter]::new($entry.Open(), $encoding)
                try { $writer.Write($part.Value) } finally { $writer.Dispose() }
            }
        } finally { $archive.Dispose() }
    } finally { $stream.Dispose() }
}

$prefix = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>'
$suffix = '<w:sectPr/></w:body></w:document>'

Write-Probe 'shared-hyperlink-carriage-return.docx' ($prefix + '<w:p><w:r><w:t>Before</w:t></w:r><w:hyperlink r:id="rId2"><w:r><w:t>First</w:t><w:cr/><w:t>Second</w:t></w:r></w:hyperlink><w:r><w:t>After</w:t></w:r></w:p>' + $suffix) @{
    'word/_rels/document.xml.rels' = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/" TargetMode="External"/></Relationships>'
}

Write-Probe 'shared-repeated-highlight-stories.docx' ($prefix + '<w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>Repeated</w:t></w:r><w:r><w:t xml:space="preserve"> plain </w:t></w:r><w:r><w:rPr><w:highlight w:val="green"/></w:rPr><w:t>Repeated</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/></w:r><w:r><w:t>Second page</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId2"/></w:sectPr></w:body></w:document>') @{
    'word/_rels/document.xml.rels' = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>'
    'word/header1.xml' = '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:highlight w:val="cyan"/></w:rPr><w:t>Repeated header</w:t></w:r></w:p></w:hdr>'
}

Write-Probe 'shared-altchunk-charset.docx' ($prefix + '<w:p><w:r><w:t>Before chunk</w:t></w:r></w:p><w:altChunk r:id="rId2"/><w:p><w:r><w:t>After chunk</w:t></w:r></w:p>' + $suffix) @{
    'word/_rels/document.xml.rels' = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="chunk.html"/></Relationships>'
    'word/chunk.html' = '<html><head><meta charset="windows-1252"></head><body><p>Café from chunk</p><table><tr><td rowspan="2">Tall cell</td><td>First row</td></tr><tr><td>Second row</td></tr></table></body></html>'
}

Write-Probe 'shared-altchunk-table-cell-html.docx' ($prefix + '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Before</w:t></w:r></w:p><w:altChunk r:id="rId2"/><w:p><w:r><w:t>After</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' + $suffix) @{
    'word/_rels/document.xml.rels' = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="chunk.html"/></Relationships>'
    'word/chunk.html' = '<html><body><p>Imported cell</p></body></html>'
}

Write-Probe 'shared-altchunk-table-cell-text.docx' ($prefix + '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Before</w:t></w:r></w:p><w:altChunk r:id="rId2"/><w:p><w:r><w:t>After</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' + $suffix) @{
    'word/_rels/document.xml.rels' = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="chunk.txt"/></Relationships>'
    'word/chunk.txt' = 'Imported cell'
}

Write-Probe 'shared-line-number-continuous-sections.docx' ($prefix + '<w:p><w:pPr><w:sectPr><w:lnNumType w:countBy="1"/></w:sectPr></w:pPr><w:r><w:t>First section</w:t></w:r></w:p><w:p><w:r><w:t>Second section</w:t></w:r></w:p><w:sectPr><w:lnNumType w:countBy="1" w:restart="continuous"/></w:sectPr></w:body></w:document>') @{}
