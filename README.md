# mht2pdf

MHT/MHTML archival conversion with metadata enrichment.

## What it does
- Converts `.mht` / `.mhtml` to PDF via headless Chrome/Edge.
- Extracts metadata from MHT headers + HTML content.
- Embeds metadata into standard PDF fields and XMP.
- Writes JSON sidecars for traceability.
- Auto-shortens overlong output PDF names (deterministic truncate + hash suffix) to avoid Windows/Chromium path-length write failures.

## Metadata strategy
Priority order:
1. In-content metadata (HTML meta tags / JSON-LD / canonical URL)
2. MHT headers (`Snapshot-Content-Location`, `Date`, etc.)
3. Fallbacks:
   - Title: filename stem
   - Date: file creation time

Embedded fields:
- Info dictionary: `Title`, `Author`, `Subject`, `Keywords`, `CreationDate`, `ModDate`, `Creator`, `Producer`, plus custom `SourceURL`, `SourceFile`
- XMP: `dc:title`, `dc:creator`, `dc:date`, `dc:identifier`, `dc:description`, `dc:subject`, `dc:language`, `dc:publisher`

## Run
```powershell
python <project-root>\scripts\convert_mht_to_pdf.py `
  --source-root "<source-root>" `
  --recurse-subdirs `
  --skip-existing
```

Defaults if omitted:
- `--output-root`:
  - without `--recurse-subdirs`: `<source-root>\_pdf_archive`
  - with `--recurse-subdirs`: each source directory gets its own `<that-directory>\_pdf_archive`
- `--log-path`: `<output-root>\logs\convert.log` (or `MHT2PDF_LOG_PATH` if set)
- Recursion: off (top-level files only) unless `--recurse-subdirs` is provided

### Pilot run
```powershell
python <project-root>\scripts\convert_mht_to_pdf.py `
  --source-root "<source-root>" `
  --output-root "<project-root>\output" `
  --log-path "<project-root>\logs\pilot.log" `
  --recurse-subdirs `
  --max-files 25
```

### Environment override for log path
```powershell
$env:MHT2PDF_LOG_PATH = "<project-root>\logs\custom.log"
```

### Process specific files only
```powershell
python <project-root>\scripts\convert_mht_to_pdf.py `
  --source-file "D:\path\one.mht" `
  --source-file "D:\path\two.mhtml"
```

## Output
- PDFs: `output\...` (mirrors source tree)
- Sidecars: `*.metadata.json` next to each PDF
- Log: `logs\convert.log`
