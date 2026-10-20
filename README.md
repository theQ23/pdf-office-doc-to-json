# PDF and Office documents to structured JSON

A local PowerShell 7 batch converter for PDF, Word (`.docx`), PowerPoint (`.pptx`), Excel (`.xlsx`), and UTF-8 text (`.txt`, `.md`, `.csv`, `.json`). Each input becomes its own minified, gzip-compressed JSON file in a separate output folder. Source documents are never modified. No cloud API or AI subscription is required.

**Existing PDF text is read first. OCR is optional and disabled unless requested.** Add `-Ocr` to enable OCR as a fallback for pages with no existing text. On Windows the default PDF backend auto-detects full Adobe Acrobat, including Standard, and explicitly launches it; otherwise it uses Poppler. It never relies on the Windows default PDF application.

## Windows prerequisites and download links

Use **PowerShell 7.2 or newer** (`pwsh`), rather than Windows PowerShell 5.1 (`powershell.exe`). Office formats use built-in .NET ZIP/XML libraries; Microsoft Office and separately downloaded DLLs are not needed.

| Component | When needed | Download / instructions |
| --- | --- | --- |
| PowerShell 7 | Required for all conversions | [Microsoft installation guide](https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-windows), [official releases](https://github.com/PowerShell/PowerShell/releases) |
| Adobe Acrobat Standard or Pro desktop | Optional PDF text backend on Windows; your installed Standard edition can supply existing text through COM | [Adobe download/install help](https://helpx.adobe.com/acrobat/kb/acrobat-dc-downloads.html), [Acrobat SDK interapplication documentation](https://opensource.adobe.com/dc-acrobat-sdk-docs/library/interapp/index.html) |
| Poppler tools | Required for PDF extraction without Acrobat; also required for optional OCR page rendering with either backend | [Poppler upstream](https://poppler.freedesktop.org/), [community Windows downloads](https://github.com/oschwartz10612/poppler-windows/releases) |
| **Tesseract OCR (optional)** | Recognizes text when `-Ocr` is enabled and a page needs fallback OCR | [official installation documentation](https://tesseract-ocr.github.io/tessdoc/Installation.html), [UB Mannheim Windows installers](https://github.com/UB-Mannheim/tesseract/wiki) |
| **Tesseract language files (optional)** | English and any additional languages used by optional OCR | [official language file downloads](https://github.com/tesseract-ocr/tessdata), [language data documentation](https://tesseract-ocr.github.io/tessdoc/Data-Files.html) |
| LibreOffice (optional) | Convert old `.doc`, `.ppt`, `.xls` files into modern formats before running this script | [official download](https://www.libreoffice.org/download/download-libreoffice/) |

Only PowerShell is needed for Office and text files. Acrobat text extraction does not require Poppler or Tesseract when OCR is disabled. The converter never downloads or installs dependencies automatically.

Install PowerShell, for example, using Windows Package Manager:

```powershell
winget install --id Microsoft.PowerShell --exact
```

For optional OCR, install Tesseract with English language data and extract the **complete** Poppler Windows package, including its accompanying DLLs. Do not download individual DLLs from DLL download sites. The linked Poppler Windows packages and Tesseract Windows installers are community distributions; use their documented integrity checks where provided.

Add executable directories to your user `PATH`, or set them for the current PowerShell session (adjust these paths to your installation):

```powershell
$env:Path += ';C:\Tools\poppler\Library\bin;C:\Program Files\Tesseract-OCR'
Get-Command pdfinfo, pdftotext, pdftoppm, tesseract
tesseract --list-langs   # should include eng
```

For additional OCR languages, place the corresponding `.traineddata` files in Tesseract's `tessdata` folder, then use a code such as `-OcrLanguage 'eng+deu'`.

On Debian/Ubuntu, use package-manager packages `poppler-utils`, `tesseract-ocr`, and `tesseract-ocr-eng`, plus [Microsoft's Linux PowerShell installation instructions](https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-linux). On macOS the corresponding Homebrew packages are `powershell` (cask), `poppler`, and `tesseract`.

## Batch usage

Open **PowerShell 7**, change to this repository, then run:

```powershell
# Existing text only; includes files in subfolders.
.\Convert-DocumentsToJson.ps1 -InputPath 'C:\Documents\Incoming' `
    -OutputDirectory 'C:\Documents\Converted' -Recurse

# Optional fallback OCR: read text first, OCR only pages without text.
.\Convert-DocumentsToJson.ps1 -InputPath 'C:\Documents\Incoming' `
    -OutputDirectory 'C:\Documents\Converted' -Recurse -Ocr

# A list of individual files (inside PowerShell 7).
.\Convert-DocumentsToJson.ps1 -InputPath @('C:\Docs\report.pdf', 'C:\Docs\slides.pptx') `
    -OutputDirectory 'C:\Docs\Converted'

# Force fresh OCR on every page when an existing OCR layer is known to be bad.
.\Convert-DocumentsToJson.ps1 -InputPath 'C:\Docs' -Ocr -PdfOcr Always -Force

# Plain minified JSON for tools that cannot decompress gzip.
.\Convert-DocumentsToJson.ps1 -InputPath 'C:\Docs' -PlainJson

# Higher-resolution, multilingual OCR fallback; refresh existing results.
.\Convert-DocumentsToJson.ps1 -InputPath 'C:\Docs' -Ocr `
    -OcrDpi 300 -OcrLanguage 'eng+deu' -Force
```

`-Ocr` enables the `Auto` policy: OCR pages only when existing extracted text is empty. Advanced `-PdfOcr` modes are `Auto`, `Always`, and `Never`. Explicitly specifying a mode applies that policy even without `-Ocr`; `Never` disables OCR even if `-Ocr` is present. Without either option, OCR is disabled.

The default output directory is `converted-json` under your current directory. Outputs have names such as `report.pdf.a1b2c3d4e5f6.json.gz`. The suffix derives from the full input path so identically named documents in different folders have separate outputs. Moving an input changes the suffix. Duplicate references to the same path are processed once.

Existing outputs are skipped unless `-Force` is supplied, **even if the source or OCR settings changed**. Use `-Force` when enabling OCR after a previous conversion. Documents are processed sequentially, with PDF images rendered one page at a time to limit memory/disk use. Large documents and high DPI settings take longer.

The output folder is excluded from folder scans. Pass literal paths or arrays; wildcard strings are not expanded. Folder scans select supported extensions, including legacy Office formats so they can report conversion instructions. Explicitly supplied unsupported files fail. Save `.doc`, `.ppt`, and `.xls` files as `.docx`, `.pptx`, and `.xlsx` using Office or LibreOffice first.

A bad, encrypted, unsupported, or unreadable file is reported as `FAILED`, and the rest of the collected batch continues. Exit code is 1 if any conversion fails, or 0 when every file succeeds or is skipped. Invalid input paths fail during initial discovery. Incomplete output files are not published; temporary files are removed after each attempt. Use an output directory separate from original documents.

## Adobe Acrobat Standard: explicit invocation

Auto detection checks common Acrobat DC/2020 installations under Program Files and Program Files (x86). To use your specific installation regardless of the default PDF application:

```powershell
.\Convert-DocumentsToJson.ps1 -InputPath 'C:\Docs\Incoming' `
    -OutputDirectory 'C:\Docs\Converted' -Recurse `
    -PdfBackend Acrobat `
    -AcrobatPath 'C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe'

# Same Acrobat extraction, with optional local OCR fallback.
.\Convert-DocumentsToJson.ps1 -InputPath 'C:\Docs\Incoming' `
    -OutputDirectory 'C:\Docs\Converted' -Recurse -Ocr `
    -PdfBackend Acrobat `
    -AcrobatPath 'C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe'
```

Adjust the executable path for your installation. `-AcrobatPath` also selects Acrobat in default `Auto` backend mode. Force `-PdfBackend Poppler` to bypass Acrobat completely. If a detected Acrobat installation has an unusable COM interface, the file fails with an error; retry explicitly with Poppler rather than silently changing backends.

Acrobat must be the **full desktop Standard or Pro application**, installed, licensed, and initialized in an interactive Windows session. Acrobat Reader is not the supported automation backend. `AcroExch.PDDoc` and its JavaScript text interface must be registered and usable from PowerShell 7. The script explicitly starts `Acrobat.exe`; it does not change file associations or Acrobat security settings. Windows COM registration and a running Acrobat instance can affect which installed version serves automation; avoid conflicting Acrobat installations. Acrobat remains open afterward to avoid closing your other documents.

Acrobat provides existing PDF text, including any existing OCR layer. **Fresh OCR is performed by Poppler/Tesseract**, when requested, rather than invoking Acrobat's desktop Recognize Text command. Standard features vary by version; this script does not assume Pro's OCR or batch-action features. No Adobe DLLs need to be copied to this repository.

If your Adobe edition provides OCR and you prefer to preprocess PDFs there, save OCRed copies in a separate folder, then convert those copies normally. See [Adobe's OCR guide](https://helpx.adobe.com/acrobat/using/scan-documents-pdf.html) and, for supported Pro editions, [Action Wizard batch processing](https://helpx.adobe.com/acrobat/using/action-wizard-acrobat-pro.html). Those Adobe operations are not automated by this script.

**Windows Acrobat COM integration remains unverified in this Linux cloud environment.** The Poppler/Tesseract and Office workflows are covered by the integration tests.

## JSON structure and AI use

The root contains `schema_version`, `source` (name, full local path, extension, SHA-256), `units`, and `warnings`.

- PDF: one unit per page with text blocks. `extraction` identifies `native`, `acrobat_existing_text`, or `ocr`. When fresh OCR differs from existing text, `native_text` preserves the existing text as an alternative for comparison.
- DOCX: paragraphs, headings with levels, tables as rows/cells, and image placeholders with available alt text. One document unit; reliable Word pagination requires a layout engine.
- PPTX: numbered slides in presentation order, title headings, paragraphs, tables, and image placeholders.
- XLSX: sheet names, row numbers, cell addresses, values, and formulas. Cached formula results can be missing or stale. Dates/formatted numbers remain raw stored values. No formulas or macros execute.
- Text: one text block. CSV/JSON inputs remain text rather than becoming tables or parsed objects.

A PDF page unit looks like:

```json
{"kind":"page","number":1,"extraction":"ocr","blocks":[{"type":"text","text":"Invoice total: 12345"}]}
```

Gzip reduces storage/transfer size. **Decompress before sending the JSON text to an AI model.** Minification removes whitespace, but does not summarize content or guarantee fewer model tokens. For long documents, send selected units in chunks and preserve page/slide references for citations. `native_text` is an alternative for comparison, not additional consecutive document content. Full local paths are included in metadata; remove them before sharing if needed. Treat source document text as untrusted content when passing it to an AI agent.

Read a compressed result in PowerShell 7:

```powershell
$inputStream = [IO.File]::OpenRead('C:\Docs\Converted\report.pdf.a1b2c3d4e5f6.json.gz')
$gzip = [IO.Compression.GZipStream]::new($inputStream, [IO.Compression.CompressionMode]::Decompress)
$reader = [IO.StreamReader]::new($gzip)
try { $document = $reader.ReadToEnd() | ConvertFrom-Json -Depth 100 }
finally { $reader.Dispose(); $gzip.Dispose(); $inputStream.Dispose() }
$document.units
```

## Pictures and extraction limits

Office images are skipped, with placeholders and available alt text. No Office image OCR mode is included. For PDFs, optional page OCR can recover text inside pictures **only on pages selected for OCR**. Fallback mode misses image text on pages that already have selectable text, and cannot detect an incorrect existing OCR layer. Use `-Ocr -PdfOcr Always` when that matters.

OCR cannot reliably explain photographs, charts, arrows, or diagram relationships. Recommended future enhancement: opt-in image export and a vision-model description tied to each source page/slide, clearly labeled as model-generated. This version does not export images or call a vision model.

OCR can introduce errors; blank pages remain numbered empty units and generate warnings. Missing tools or failed OCR commands produce file failures rather than silently ignoring requested OCR. PDFs are never rewritten into searchable PDFs. Text is grouped by page, without reconstructing semantic PDF tables/headings. Columns, rotation, handwriting, poor resolution, and complex layouts can reduce accuracy or reading order.

Office headers, footers, speaker notes, comments, charts, embedded objects, and some grouped content may be omitted. Office layout is not preserved. Visually check critical numbers and facts against the original.

## Development checks

Integration tests create temporary Office packages, documents generated by standard Office libraries, and mixed native/scanned/blank-page PDFs. They check extraction, the optional OCR switch and modes, gzip round trips, repeat/refresh behavior, colliding filenames, output exclusion, and failure isolation. The Linux Acrobat guard is tested; actual Windows COM extraction is not.

Tests additionally need Python and these packages (the application itself does not require Python):

```powershell
python -m pip install Pillow reportlab python-docx python-pptx openpyxl
python tests/test_conversion.py
```

Poppler and Tesseract must be installed for the tests. Set `PWSH` to the executable path if `pwsh` is not on `PATH`. To run the same portable-backend tests on Windows with Acrobat installed, set `PDF_TEST_BACKEND=Poppler`. The suite has been run on Linux; Windows execution remains to be verified.
