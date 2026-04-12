# Checklist Extractor

Extracts audit procedure checklists from CaseWare Cloud SE author templates and generates formatted Excel workbooks. Each checklist becomes a separate sheet with procedures, assertions, guidance, settings, and visibility rules.

## Quick Start

```bash
# Install dependencies
pip install -r requirements.txt

# Copy and fill in credentials
cp .env.example .env

# Web UI (recommended)
python web/app.py
# → http://localhost:5003

# CLI — single checklist or full engagement
python tools/checklist_extract.py --url "<caseware-url>"

# Mock data (no credentials needed)
python tools/checklist_extract.py --mock
```

## Authentication

Two methods, checked in order:

1. **OAuth** (preferred) — set per-region credentials in `.env`:
   - `CW_CA_CLIENT_ID` / `CW_CA_CLIENT_SECRET` for Canadian environments
   - `CW_US_CLIENT_ID` / `CW_US_CLIENT_SECRET` for US environments
2. **Cookie fallback** — set `CW_COOKIES` from browser DevTools

The region prefix is derived from the hostname (e.g. `ca.cwcloudpartner.com` → `CA`).

## URL Formats

- **Engagement URL** → extracts ALL checklists in the engagement
  `https://{host}/{tenant}/e/eng/{engagementId}/index.jsp`
- **Document URL** → extracts a single checklist
  `...#/checklist/{documentId}` or `...#/efinancials/{documentId}`

## Output

Excel workbook with columns A–S per sheet:
- **A**: Procedure text (hierarchical)
- **B**: Authoritative reference (AU-C citations)
- **C**: Assertions (C, E, A, V, PD)
- **D**: Lightbulb guidance
- **E–H**: Response settings
- **I–K**: Cloud settings (notes, sign-offs)
- **L–P**: Visibility conditions 1–5
- **Q–S**: Additional cloud settings

## Key API Quirks

- Procedures use the document's `content` field as `checklistId`, not the document `id`. The tool handles this mapping internally.
- Assertions come from `proc.tagging.baseassertion` — 14 API tags map down to 5 display assertions (C, E, A, V, PD). See the workflow for the full mapping rules.
- Citations live in `proc.attachables` with `kind: "citation"`, not a top-level field.
- Settings field names differ from what you'd guess: `allowSignOffs` (not `allowSignoffs`), `allowNote` (not `allowInputNotes`), `showResponsesBelow` (not `showResponseBeneathProcedure`).
- Procedures without explicit settings inherit from checklist-level defaults via `checklist/get`.

## Project Structure

```
tools/checklist_extract.py   # Core extraction engine
web/app.py                   # Flask backend (port 5003)
web/static/                  # Frontend (vanilla HTML/JS/CSS)
workflows/extract_checklists.md  # Full workflow SOP
.env.example                 # Credential template
```
