#requires -Version 7.2
<#
.SYNOPSIS
Converts batches of PDF, DOC/DOCX, PPT/PPTX, XLS/XLSX and text files to structured JSON.
.DESCRIPTION
PDFs use existing text first; -Ocr enables local OCR for empty pages. On Windows, auto
detects full Acrobat; otherwise uses Poppler. PDF OCR needs Poppler and Tesseract.
Legacy .doc/.ppt/.xls files are converted in a temp folder with installed Word/PowerPoint/
Excel (macros disabled); originals are never modified. Acrobat failures in Auto backend mode
fall back to Poppler when Poppler is on PATH. -PlainJson writes .json instead of .json.gz.
OneDrive cloud-only files are downloaded on demand before conversion; progress is shown per file.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory, Position=0)][string[]] $InputPath,
    [string] $OutputDirectory = (Join-Path $PWD 'converted-json'),
    [switch] $Recurse,
    [ValidateSet('Auto','Poppler','Acrobat')][string] $PdfBackend = 'Auto',
    [string] $AcrobatPath,
    [switch] $Ocr,
    [ValidateSet('Always','Auto','Never')][string] $PdfOcr = 'Auto',
    [string] $OcrLanguage = 'eng',
    [ValidateRange(72,600)][int] $OcrDpi = 200,
    [switch] $PlainJson,
    [switch] $Force
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# An explicit advanced mode also enables that OCR policy; the default is opt-in.
if (-not $Ocr -and -not $PSBoundParameters.ContainsKey('PdfOcr')) {$PdfOcr='Never'}

function Invoke-Tool([string] $Name, [string[]] $Arguments) {
    $tool = Get-Command $Name -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $tool) { throw "Missing '$Name' on PATH. See README.md for installation." }
    $result = & $tool.Source @Arguments 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) { throw "$Name failed (exit $LASTEXITCODE): $result" }
    return $result
}
function New-Block([string] $Type, [string] $Text) {
    return [ordered]@{ type=$Type; text=$Text.Trim() }
}
function Read-ZipXml($Zip, [string] $Name) {
    $entry = $Zip.GetEntry($Name)
    if (-not $entry) { throw "Missing Office package entry: $Name" }
    $stream = $entry.Open()
    try {
        $settings = [System.Xml.XmlReaderSettings]::new()
        $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
        $settings.XmlResolver = $null
        $reader = [System.Xml.XmlReader]::Create($stream, $settings)
        try { $xml = [System.Xml.XmlDocument]::new(); $xml.XmlResolver = $null; $xml.Load($reader) }
        finally { $reader.Dispose() }
        return ,$xml
    } finally { $stream.Dispose() }
}
function Get-WordParagraph($Node) {
    $parts = foreach ($item in $Node.SelectNodes('.//*[local-name()="t" or local-name()="tab" or local-name()="br"]')) {
        switch ($item.LocalName) { 't' { $item.InnerText }; 'tab' { "`t" }; 'br' { "`n" } }
    }
    return ($parts -join '').Trim()
}
function Convert-Word($Zip) {
    $xml = Read-ZipXml $Zip 'word/document.xml'
    $blocks = [System.Collections.Generic.List[object]]::new()
    foreach ($node in $xml.SelectNodes('//*[local-name()="body"]/*')) {
        if ($node.LocalName -eq 'p') {
            $text = Get-WordParagraph $node
            if ($text) {
                $style = $node.SelectSingleNode('./*[local-name()="pPr"]/*[local-name()="pStyle"]')
                $block = New-Block 'paragraph' $text
                if ($style) {
                    $block['style'] = $style.GetAttribute('val', $style.NamespaceURI)
                    if ($block.style -match '^Heading\s*([1-9])$') { $block.type='heading'; $block['level']=[int]$Matches[1] }
                }
                $blocks.Add($block)
            }
        } elseif ($node.LocalName -eq 'tbl') {
            $rows = @(foreach ($row in $node.SelectNodes('./*[local-name()="tr"]')) {
                $cells = @(foreach ($cell in $row.SelectNodes('./*[local-name()="tc"]')) {
                    (@(foreach ($p in $cell.SelectNodes('./*[local-name()="p"]')) { Get-WordParagraph $p }) -join "`n")
                })
                ,$cells
            })
            $blocks.Add([ordered]@{type='table'; rows=$rows})
        }
        foreach ($image in $node.SelectNodes('.//*[local-name()="docPr"]')) {
            $blocks.Add([ordered]@{type='image'; alt_text=$image.GetAttribute('descr'); name=$image.GetAttribute('name'); extracted=$false})
        }
    }
    return [ordered]@{kind='document'; number=1; blocks=@($blocks.ToArray())}
}
function Convert-Presentation($Zip) {
    # Follow presentation relationships: numeric filenames need not match slide order.
    $presentation = Read-ZipXml $Zip 'ppt/presentation.xml'
    $rels = Read-ZipXml $Zip 'ppt/_rels/presentation.xml.rels'
    $targets = @{}
    foreach ($rel in $rels.DocumentElement.ChildNodes) { $targets[$rel.GetAttribute('Id')] = $rel.GetAttribute('Target') }
    $number = 0
    foreach ($id in $presentation.SelectNodes('//*[local-name()="sldId"]')) {
        $number++
        $rid = $id.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $uri = [uri]::new([uri]'https://package.invalid/ppt/presentation.xml', $targets[$rid])
        $xml = Read-ZipXml $Zip $uri.AbsolutePath.TrimStart('/')
        $blocks = [System.Collections.Generic.List[object]]::new()
        foreach ($shape in $xml.SelectNodes('//*[local-name()="spTree"]/*')) {
            $table = $shape.SelectSingleNode('.//*[local-name()="tbl"]')
            if ($table) {
                $rows = @(foreach ($row in $table.SelectNodes('./*[local-name()="tr"]')) {
                    $cells = @(foreach ($cell in $row.SelectNodes('./*[local-name()="tc"]')) {
                        (@(foreach ($p in $cell.SelectNodes('.//*[local-name()="p"]')) {
                            @($p.SelectNodes('.//*[local-name()="t"]') | ForEach-Object {$_.InnerText}) -join ''
                        }) -join "`n")
                    })
                    ,$cells
                })
                $blocks.Add([ordered]@{type='table'; rows=$rows})
                continue
            }
            $placeholder = $shape.SelectSingleNode('.//*[local-name()="ph"]')
            foreach ($p in $shape.SelectNodes('.//*[local-name()="p"]')) {
                $text = (@($p.SelectNodes('.//*[local-name()="t"]') | ForEach-Object { $_.InnerText }) -join '').Trim()
                if ($text) {
                    $block=New-Block 'paragraph' $text
                    if ($placeholder -and $placeholder.GetAttribute('type') -in 'title','ctrTitle') {$block.type='heading'; $block['level']=1}
                    $blocks.Add($block)
                }
            }
            foreach ($pic in $shape.SelectNodes('descendant-or-self::*[local-name()="pic"]')) {
                $props=$pic.SelectSingleNode('.//*[local-name()="cNvPr"]')
                $blocks.Add([ordered]@{type='image'; alt_text=$(if ($props) {$props.GetAttribute('descr')} else {''}); extracted=$false})
            }
        }
        [ordered]@{kind='slide'; number=$number; blocks=@($blocks.ToArray())}
    }
}
function Convert-Spreadsheet($Zip) {
    $strings = @()
    if ($Zip.GetEntry('xl/sharedStrings.xml')) {
        $shared = Read-ZipXml $Zip 'xl/sharedStrings.xml'
        $strings = @(foreach ($si in $shared.SelectNodes('//*[local-name()="si"]')) {
            (@($si.SelectNodes('.//*[local-name()="t"]') | ForEach-Object {$_.InnerText}) -join '')
        })
    }
    $book = Read-ZipXml $Zip 'xl/workbook.xml'
    $rels = Read-ZipXml $Zip 'xl/_rels/workbook.xml.rels'
    $targets=@{}; foreach ($rel in $rels.DocumentElement.ChildNodes) { $targets[$rel.GetAttribute('Id')]=$rel.GetAttribute('Target') }
    $number=0
    foreach ($sheet in $book.SelectNodes('//*[local-name()="sheet"]')) {
        $number++
        $rid=$sheet.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $uri=[uri]::new([uri]'https://package.invalid/xl/workbook.xml',$targets[$rid])
        $xml=Read-ZipXml $Zip $uri.AbsolutePath.TrimStart('/')
        $rows=@(foreach ($row in $xml.SelectNodes('//*[local-name()="sheetData"]/*[local-name()="row"]')) {
            $cells=@(foreach ($cell in $row.SelectNodes('./*[local-name()="c"]')) {
                $v=$cell.SelectSingleNode('./*[local-name()="v"]')
                $value=if ($v) {$v.InnerText} else {''}
                if ($cell.GetAttribute('t') -eq 's') {$value=$strings[[int]$value]}
                elseif ($cell.GetAttribute('t') -eq 'inlineStr') {$value=(@($cell.SelectNodes('.//*[local-name()="t"]') | ForEach-Object {$_.InnerText}) -join '')}
                $result=[ordered]@{cell=$cell.GetAttribute('r'); value=$value}
                $formula=$cell.SelectSingleNode('./*[local-name()="f"]')
                if ($formula) {$result['formula']=$formula.InnerText}
                $result
            })
            [ordered]@{number=[int]$row.GetAttribute('r'); cells=$cells}
        })
        [ordered]@{kind='sheet'; number=$number; name=$sheet.GetAttribute('name'); blocks=@([ordered]@{type='table'; rows=$rows})}
    }
}
function Invoke-PageOcr([string] $Path, [int] $Page, [string] $Temp) {
    $prefix=Join-Path $Temp "ocr-$Page"
    $null=Invoke-Tool 'pdftoppm' @('-f',"$Page",'-l',"$Page",'-singlefile','-r',"$OcrDpi",'-png',$Path,$prefix)
    $null=Invoke-Tool 'tesseract' @("$prefix.png",$prefix,'-l',$OcrLanguage)
    $text=[IO.File]::ReadAllText("$prefix.txt").Trim()
    Remove-Item -LiteralPath "$prefix.png","$prefix.txt" -Force
    return $text
}
function Convert-Pdf([string] $Path, [string] $Temp, $Warnings) {
    $info = Invoke-Tool 'pdfinfo' @($Path)
    if ($info -notmatch '(?m)^Pages:\s+(\d+)') { throw 'Cannot determine PDF page count.' }
    $pages=[int]$Matches[1]
    for ($page=1; $page -le $pages; $page++) {
        $textFile=Join-Path $Temp "page-$page.txt"
        $null=Invoke-Tool 'pdftotext' @('-f',"$page",'-l',"$page",'-layout','-enc','UTF-8',$Path,$textFile)
        $native=[IO.File]::ReadAllText($textFile).Trim([char[]]" `t`r`n`f")
        $text=$native; $method='native'
        if ($PdfOcr -eq 'Always' -or ($PdfOcr -eq 'Auto' -and [string]::IsNullOrWhiteSpace($native))) {
            # Render one page at a time to bound disk and memory use.
            $ocr=Invoke-PageOcr $Path $page $Temp
            if ($ocr) { $text=$ocr; $method='ocr' }
            elseif ($native) { $Warnings.Add("Page ${page}: OCR returned no text; used native text.") }
            else { $method='ocr'; $Warnings.Add("Page ${page}: no text detected; visual review may be needed.") }
        }
        $blocks=@(); if ($text) {$blocks=@((New-Block 'text' $text))}
        if (-not $text -and $method -eq 'native') {$Warnings.Add("Page ${page}: no native text found; enable OCR for scanned pages.")}
        $unit=[ordered]@{kind='page'; number=$page; extraction=$method; blocks=$blocks}
        # Preserve differing existing text for comparison rather than silently discarding it.
        if ($method -eq 'ocr' -and $native -and $native -ne $text) {$unit['native_text']=$native}
        $unit
    }
}

function Find-Acrobat {
        if (-not $IsWindows) {return $null}
        $executable=$AcrobatPath
        if (-not $executable) {
            foreach ($root in @([Environment]::GetEnvironmentVariable('ProgramFiles'), [Environment]::GetEnvironmentVariable('ProgramFiles(x86)'))) {
                if (-not $root) {continue}
                foreach ($relative in @('Adobe/Acrobat DC/Acrobat/Acrobat.exe', 'Adobe/Acrobat 2020/Acrobat/Acrobat.exe')) {
                    $candidate=Join-Path $root $relative
                    if (Test-Path -LiteralPath $candidate -PathType Leaf) {$executable=$candidate; break}
                }
                if ($executable) {break}
            }
        }
        return $executable
}
function Invoke-AcroJs($Jso, [string] $Method, [object[]] $Arguments) {
    # Acrobat's JavaScript bridge only answers late-bound IDispatch calls; PowerShell's COM binder fails on it.
    return $Jso.GetType().InvokeMember($Method, [Reflection.BindingFlags]::InvokeMethod, $null, $Jso, $Arguments)
}
function Convert-AcrobatPdf([string] $Path, [string] $Temp, $Warnings) {
    if (-not $IsWindows) {throw 'The Acrobat backend requires Windows and the full desktop Adobe Acrobat installation.'}
    if (-not $script:AcrobatStarted) {
        $executable=Find-Acrobat
        if (-not $executable -or -not (Test-Path -LiteralPath $executable -PathType Leaf)) {throw 'Cannot locate Acrobat.exe. Supply -AcrobatPath with its full installed path.'}
        if ([IO.Path]::GetFileName($executable) -ine 'Acrobat.exe') {throw '-AcrobatPath must point to Acrobat.exe from the full Adobe Acrobat desktop installation.'}
        # Launch Acrobat explicitly instead of asking Windows to open a PDF by association.
        $process=Start-Process -FilePath $executable -PassThru
        # Give Acrobat time to initialize its automation server before the first COM call.
        try { $null=$process.WaitForInputIdle(30000) } catch { Start-Sleep -Seconds 5 }
        $script:AcrobatStarted=$true
    }
    # Open a local short-path copy: avoids Acrobat failing on OneDrive, long, or sandbox-blocked paths.
    $local=Join-Path $Temp 'acrobat-source.pdf'
    [IO.File]::Copy($Path, $local, $true)
    $pdf=$null; $js=$null
    try {
        $pdf=New-Object -ComObject AcroExch.PDDoc
        if (-not $pdf.Open($local)) {throw 'Acrobat could not open the PDF. Check encryption and the desktop installation.'}
        $js=$pdf.GetJSObject()
        if ($null -eq $js) {throw 'Acrobat JavaScript text-extraction interface is unavailable.'}
        $count=$pdf.GetNumPages()
        for ($page=0; $page -lt $count; $page++) {
            $words=[System.Collections.Generic.List[string]]::new()
            $wordCount=[int](Invoke-AcroJs $js 'getPageNumWords' @($page))
            for ($word=0; $word -lt $wordCount; $word++) {
                $words.Add([string](Invoke-AcroJs $js 'getPageNthWord' @($page,$word,$false)))
            }
            $text=($words.ToArray() -join ' ').Trim()
            $native=$text; $method='acrobat_existing_text'
            if ($PdfOcr -eq 'Always' -or ($PdfOcr -eq 'Auto' -and -not $text)) {
                $ocr=Invoke-PageOcr $Path ($page+1) $Temp
                if ($ocr) {$text=$ocr; $method='ocr'}
                elseif ($native) {$Warnings.Add("Page $($page+1): OCR returned no text; used Acrobat text.")}
                else {$method='ocr'; $Warnings.Add("Page $($page+1): no text detected; visual review may be needed.")}
            }
            $blocks=@(); if ($text) {$blocks=@((New-Block 'text' $text))}
            elseif ($PdfOcr -eq 'Never') {$Warnings.Add("Page $($page+1): Acrobat found no text; enable OCR for scanned pages.")}
            $unit=[ordered]@{kind='page'; number=($page+1); extraction=$method; blocks=$blocks}
            if ($method -eq 'ocr' -and $native -and $native -ne $text) {$unit['native_text']=$native}
            $unit
        }
        $Warnings.Add('Acrobat supplies existing text; fresh OCR, when needed, uses Poppler/Tesseract. Word order and layout may differ from the original.')
    } finally {
        if ($js -and [Runtime.InteropServices.Marshal]::IsComObject($js)) {$null=[Runtime.InteropServices.Marshal]::FinalReleaseComObject($js)}
        if ($pdf) {$null=$pdf.Close(); $null=[Runtime.InteropServices.Marshal]::FinalReleaseComObject($pdf)}
        # Leave the Acrobat application running; it may contain the user's other documents.
    }
}
function Convert-LegacyWord([string] $Path, [string] $Temp) {
    if (-not $IsWindows) { throw 'Legacy .doc conversion requires Windows with Microsoft Word installed.' }
    if ($null -eq $script:Word) {
        try { $script:Word = New-Object -ComObject Word.Application }
        catch { throw 'Legacy .doc conversion requires Microsoft Word. Save the file as .docx with Word or LibreOffice first.' }
        $script:Word.Visible = $false
        $script:Word.DisplayAlerts = 0          # wdAlertsNone
        $script:Word.AutomationSecurity = 3     # msoAutomationSecurityForceDisable: macros never run
    }
    $converted = Join-Path $Temp 'legacy-converted.docx'
    $doc = $null
    try {
        # Read-only open. Dummy password: unprotected files ignore it; protected files fail
        # fast instead of raising an invisible prompt that would hang the hidden Word instance.
        $doc = $script:Word.Documents.Open($Path, $false, $true, $false, 'no-password-supplied')
        $doc.SaveAs2($converted, 16)            # wdFormatDocumentDefault (.docx)
    } finally {
        if ($null -ne $doc) {
            try { $doc.Close(0) } catch { }     # wdDoNotSaveChanges
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
        }
    }
    if (-not (Test-Path -LiteralPath $converted -PathType Leaf)) { throw 'Word did not produce a .docx copy.' }
    return $converted
}
function Convert-LegacyExcel([string] $Path, [string] $Temp) {
    if (-not $IsWindows) { throw 'Legacy .xls conversion requires Windows with Microsoft Excel installed.' }
    if ($null -eq $script:Excel) {
        try { $script:Excel = New-Object -ComObject Excel.Application }
        catch { throw 'Legacy .xls conversion requires Microsoft Excel. Save the file as .xlsx with Excel or LibreOffice first.' }
        $script:Excel.Visible = $false
        $script:Excel.DisplayAlerts = $false
        $script:Excel.AskToUpdateLinks = $false
        $script:Excel.EnableEvents = $false
        $script:Excel.AutomationSecurity = 3    # msoAutomationSecurityForceDisable: macros never run
    }
    $converted = Join-Path $Temp 'legacy-converted.xlsx'
    $book = $null
    try {
        # UpdateLinks=0, ReadOnly, dummy password (fails fast on protected files), ignore read-only recommendation.
        $book = $script:Excel.Workbooks.Open($Path, 0, $true, [Type]::Missing, 'no-password-supplied', [Type]::Missing, $true)
        $book.SaveAs($converted, 51)            # xlOpenXMLWorkbook (.xlsx)
    } finally {
        if ($null -ne $book) {
            try { $book.Close($false) } catch { }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)
        }
    }
    if (-not (Test-Path -LiteralPath $converted -PathType Leaf)) { throw 'Excel did not produce an .xlsx copy.' }
    return $converted
}
function Convert-LegacyPowerPoint([string] $Path, [string] $Temp) {
    if (-not $IsWindows) { throw 'Legacy .ppt conversion requires Windows with Microsoft PowerPoint installed.' }
    if ($null -eq $script:PowerPoint) {
        try { $script:PowerPoint = New-Object -ComObject PowerPoint.Application }
        catch { throw 'Legacy .ppt conversion requires Microsoft PowerPoint. Save the file as .pptx with PowerPoint or LibreOffice first.' }
        $script:PowerPoint.DisplayAlerts = 1    # ppAlertsNone (PowerPoint cannot be made invisible; WithWindow=false hides the file)
        $script:PowerPoint.AutomationSecurity = 3
    }
    $converted = Join-Path $Temp 'legacy-converted.pptx'
    $deck = $null
    try {
        $deck = $script:PowerPoint.Presentations.Open($Path, -1, 0, 0)   # ReadOnly, not Untitled, no window
        $deck.SaveAs($converted, 24)            # ppSaveAsOpenXMLPresentation (.pptx)
    } finally {
        if ($null -ne $deck) {
            try { $deck.Close() } catch { }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($deck)
        }
    }
    if (-not (Test-Path -LiteralPath $converted -PathType Leaf)) { throw 'PowerPoint did not produce a .pptx copy.' }
    return $converted
}
function Close-OfficeApp($App, [string] $Collection) {
    # Quit only if this instance holds nothing else; never close a user's open files.
    if ($null -eq $App) { return }
    try { if ($App.$Collection.Count -eq 0) { $App.Quit() } } catch { }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($App) } catch { }
}
function Confirm-LocalFile($File, [string] $Tag) {
    # OneDrive Files On-Demand placeholders: RecallOnDataAccess | RecallOnOpen | Offline.
    $cloudOnly = 0x400000 -bor 0x40000 -bor 0x1000
    if (([int]$File.Attributes -band $cloudOnly) -eq 0) { return $false }
    $sizeMb = [math]::Round($File.Length / 1MB, 1)
    Write-Host "$Tag DOWNLOADING from OneDrive: $($File.Name) ($sizeMb MB)"
    Write-Progress -Activity 'Converting documents' -Status "$Tag downloading $($File.Name) from OneDrive"
    # Reading the whole file forces OneDrive to hydrate it, so download errors surface clearly here.
    $buffer = [byte[]]::new(1MB)
    try {
        $stream = [IO.File]::OpenRead($File.FullName)
        try { while ($stream.Read($buffer, 0, $buffer.Length) -gt 0) { } } finally { $stream.Dispose() }
    } catch {
        throw "OneDrive download failed (check OneDrive is running, signed in, and not paused): $($_.Exception.Message)"
    }
    return $true
}

$script:AcrobatStarted=$false
$script:Word=$null
$script:Excel=$null
$script:PowerPoint=$null
$extensions=@('.pdf','.docx','.pptx','.xlsx','.txt','.md','.csv','.json','.doc','.ppt','.xls')
$output=[IO.Path]::GetFullPath($OutputDirectory)
$null=[IO.Directory]::CreateDirectory($output)
$files=[System.Collections.Generic.Dictionary[string,object]]::new([StringComparer]::Ordinal)
foreach ($path in $InputPath) {
    $item=Get-Item -LiteralPath $path
    if ($item.PSIsContainer) {
        foreach ($file in Get-ChildItem -LiteralPath $item.FullName -File -Recurse:$Recurse) {
            if ($file.Extension.ToLowerInvariant() -in $extensions -and -not $file.FullName.StartsWith($output+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { $files[$file.FullName]=$file }
        }
    } else { $files[$item.FullName]=$item }
}
if ($files.Count -eq 0) { throw 'No supported input files found.' }
$failed=0; $ok=0; $skipped=0; $downloaded=0; $index=0; $total=$files.Count
try {
foreach ($file in ($files.Values | Sort-Object FullName)) {
    $index++
    $tag="[$index/$total]"
    Write-Progress -Activity 'Converting documents' -Status "$tag $($file.Name)" -PercentComplete ([int](($index-1)*100/$total))
    $temp=Join-Path ([IO.Path]::GetTempPath()) ('doc-json-'+[guid]::NewGuid().ToString('N'))
    $pending=$null
    try {
        if ($file.DirectoryName -eq $output) { throw 'Input files must be outside the output directory.' }
        $ext=$file.Extension.ToLowerInvariant()
        if ($ext -notin $extensions) { throw "Unsupported format: $ext" }
        $pathHash=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($file.FullName))).Substring(0,12).ToLowerInvariant()
        $suffix=if ($PlainJson) {'.json'} else {'.json.gz'}
        $destination=Join-Path $output ($file.Name+".$pathHash"+$suffix)
        if ((Test-Path -LiteralPath $destination) -and -not $Force) { $skipped++; Write-Host "$tag SKIP $($file.Name): output exists (use -Force to refresh)"; continue }
        if (Confirm-LocalFile $file $tag) { $downloaded++ }
        Write-Progress -Activity 'Converting documents' -Status "$tag converting $($file.Name)" -PercentComplete ([int](($index-1)*100/$total))
        $null=[IO.Directory]::CreateDirectory($temp)
        $warnings=[System.Collections.Generic.List[string]]::new()
        $units=@()
        if ($ext -eq '.pdf') {
            $warnings.Add('OCR and PDF reading order may be inaccurate. Charts and diagrams are not interpreted.')
            $backend=$PdfBackend
            if ($backend -eq 'Auto') {$backend=if (Find-Acrobat) {'Acrobat'} else {'Poppler'}}
            if ($backend -eq 'Acrobat') {
                try { $units=@(Convert-AcrobatPdf $file.FullName $temp $warnings) }
                catch {
                    # Explicit -PdfBackend Acrobat, or no Poppler available: report the Acrobat failure.
                    if ($PdfBackend -ne 'Auto' -or -not (Get-Command pdftotext -CommandType Application -ErrorAction SilentlyContinue)) { throw }
                    $reason=$_.Exception.Message
                    Write-Host "$tag Acrobat failed ($reason); retrying with Poppler"
                    $warnings.Add("Acrobat extraction failed ($reason); text extracted with Poppler instead.")
                    $units=@(Convert-Pdf $file.FullName $temp $warnings)
                }
            }
            else {$units=@(Convert-Pdf $file.FullName $temp $warnings)}
        } elseif ($ext -in '.doc','.docx','.ppt','.pptx','.xls','.xlsx') {
            $package=$file.FullName
            switch ($ext) {
                '.doc' { $package=Convert-LegacyWord $file.FullName $temp; $warnings.Add('Converted from legacy .doc to .docx with Microsoft Word before extraction.') }
                '.ppt' { $package=Convert-LegacyPowerPoint $file.FullName $temp; $warnings.Add('Converted from legacy .ppt to .pptx with Microsoft PowerPoint before extraction.') }
                '.xls' { $package=Convert-LegacyExcel $file.FullName $temp; $warnings.Add('Converted from legacy .xls to .xlsx with Microsoft Excel before extraction; cached values may have been recalculated.') }
            }
            $zip=[IO.Compression.ZipFile]::OpenRead($package)
            try {
                $warnings.Add('Embedded images are not OCRed. Headers, footers, notes, comments, charts, and embedded objects may be omitted; Office layout is not preserved.')
                $units=@(switch ($ext) { {$_ -in '.doc','.docx'} {Convert-Word $zip}; {$_ -in '.ppt','.pptx'} {Convert-Presentation $zip}; {$_ -in '.xls','.xlsx'} {Convert-Spreadsheet $zip} })
            } finally { $zip.Dispose() }
        } else {
            $units=@([ordered]@{kind='document'; number=1; blocks=@((New-Block 'text' ([IO.File]::ReadAllText($file.FullName))))})
        }
        $doc=[ordered]@{
            schema_version='1.0'
            source=[ordered]@{name=$file.Name; path=$file.FullName; format=$ext.TrimStart('.'); sha256=(Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()}
            units=$units
            warnings=@($warnings.ToArray())
        }
        $json=ConvertTo-Json -InputObject $doc -Depth 100 -Compress
        $pending=Join-Path $output ('.'+[guid]::NewGuid().ToString('N')+'.tmp')
        $bytes=[Text.Encoding]::UTF8.GetBytes($json)
        $stream=[IO.File]::Create($pending)
        try {
            if ($PlainJson) { $stream.Write($bytes,0,$bytes.Length) }
            else {
                $gzip=[IO.Compression.GZipStream]::new($stream,[IO.Compression.CompressionLevel]::Optimal,$true)
                try {$gzip.Write($bytes,0,$bytes.Length)} finally {$gzip.Dispose()}
            }
        } finally {$stream.Dispose()}
        [IO.File]::Move($pending,$destination,[bool]$Force)
        $ok++; Write-Host "$tag OK $($file.Name) -> $destination"
    } catch {
        $failed++
        Write-Warning "$tag FAILED $($file.FullName): $($_.Exception.Message)"
    } finally {
        if ($pending -and (Test-Path -LiteralPath $pending)) {Remove-Item -LiteralPath $pending -Force -ErrorAction SilentlyContinue}
        if (Test-Path -LiteralPath $temp) {Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue}
    }
}
} finally {
    Write-Progress -Activity 'Converting documents' -Completed
    Close-OfficeApp $script:Word 'Documents'
    Close-OfficeApp $script:Excel 'Workbooks'
    Close-OfficeApp $script:PowerPoint 'Presentations'
    $script:Word=$null; $script:Excel=$null; $script:PowerPoint=$null
}
Write-Host "Done: $ok converted, $skipped skipped, $failed failed (of $total); $downloaded downloaded from OneDrive."
if ($failed) { Write-Warning "$failed document(s) failed."; exit 1 }
exit 0