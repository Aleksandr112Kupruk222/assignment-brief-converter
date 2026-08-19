# Assignment Brief Converter

A small FastAPI web application that migrates content from an old college assignment brief into the structure of a completed example using a newer Word template. The old document is the content source of truth; the newer document is the layout and formatting source of truth.

> **Teacher review is mandatory.** Every generated brief must be checked by the teacher before official use, especially dates, assessment criteria, learning outcomes, qualification details, and any missing fields.

## What it does

- Accepts two `.docx` files: an old assignment and a completed new-template example.
- Inspects paragraphs, Word headings, tables, and table cells.
- Detects common academic fields using aliases rather than fixed template headings.
- Suggests old-to-new mappings with a confidence score.
- Lets the teacher change mappings or enter missing values manually.
- Generates a downloadable `.docx` based on a copy of the new template.
- Preserves the new document's styles, tables, page layout, headers, and footers wherever `python-docx` can preserve them.

## Run locally

Python 3.11 or newer is recommended.

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements-dev.txt
uvicorn app.main:app --reload
```

Open <http://127.0.0.1:8000>.

## Tests

```powershell
pytest
```

The tests cover upload validation, corrupt/empty/duplicate documents, paragraph and table extraction, section detection, mapping, missing fields, and generated document content.

## Privacy behaviour

There is no database, account system, analytics service, or third-party document processor. Uploads are stored in a random directory under the operating system's temporary directory for the active conversion only. Files are removed after generation, and abandoned sessions are removed after one hour when the next analysis starts. Document contents are not logged.

For an internet-facing deployment, add transport encryption, request-rate limits, malware scanning, and a scheduled cleanup job appropriate to the hosting platform.

## Current limitations

- `.docx` is the only supported format; legacy `.doc`, PDFs, and protected documents are not supported.
- Detection is heuristic. Unusual labels and complex nested tables may need manual mapping.
- The MVP replaces detected value cells and the first content paragraph following detected headings. It does not rebuild multi-row task tables or merge/split template sections intelligently.
- Inline formatting within a replaced value is simplified to the formatting of the destination's first run. The surrounding table, paragraph style, headers, footers, media, and page settings remain those of the new template.
- Word fields, content controls, text boxes, tracked changes, macros, and embedded objects are not interpreted.
- Missing information is never invented, but requiredness cannot always be inferred from a completed example. Teachers must review all unmapped fields.

## Project structure

```text
app/
  main.py              FastAPI routes, temporary-session lifecycle, validation
  documents.py         DOCX extraction, detection, mapping, generation
  static/               Lightweight HTML/CSS/JavaScript interface
tests/                  Unit and end-to-end API tests
requirements*.txt      Runtime and development dependencies
```

## Recommended next steps

1. Add explicit support for content controls and more complex task/criteria tables.
2. Add a visual, page-level preview before download.
3. Learn per-college label aliases from approved user corrections without storing document contents.
4. Add deployment hardening and automated retention cleanup.
