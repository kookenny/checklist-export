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

Excel workbook with one **procedure sheet per checklist** (columns A–S), plus a trailing **`Glossary & Dynamic Text`** reference sheet when the template contains any dynamic text or glossary references.

Procedure sheet columns:
- **A**: Procedure text (hierarchical; inline `[[value]]`/`[[?]]` markers in blue mark each dynamic-text span, and the cell hyperlinks to the reference sheet)
- **B**: Authoritative reference (AU-C citations)
- **C**: Assertions (C, E, A, V, PD)
- **D**: Lightbulb guidance
- **E–H**: Response settings
- **I–K**: Cloud settings (notes, sign-offs)
- **L–P**: Visibility conditions 1–5
- **Q–S**: Additional cloud settings

The reference sheet has two stacked sections: GLOSSARY (unique wording terms referenced by any extracted checklist, with all their conditions and outputs) and DYNAMIC TEXT (every `<span formula="">` occurrence expanded into one row per condition). Wording-type rows in the Dynamic Text section hyperlink back up to the matching Glossary row. See [workflows/extract_checklists.md](workflows/extract_checklists.md) for the full layout.

## Key API Quirks

- Procedures use the document's `content` field as `checklistId`, not the document `id`. The tool handles this mapping internally.
- Assertions come from `proc.tagging.baseassertion` — 14 API tags map down to 5 display assertions (C, E, A, V, PD). See the workflow for the full mapping rules.
- Citations live in `proc.attachables` with `kind: "citation"`, not a top-level field.
- Settings field names differ from what you'd guess: `allowSignOffs` (not `allowSignoffs`), `allowNote` (not `allowInputNotes`), `showResponsesBelow` (not `showResponseBeneathProcedure`).
- Procedures without explicit settings inherit from checklist-level defaults via `checklist/get`.
- **Dynamic text** (`<span formula="refId">` in procedure HTML) resolves via `proc.attachables[refId]` — two flavours: local `values[]` arrays of condition/output pairs, or global glossary references `wording("@<tag_id>")`.
- **Glossary terms** are NOT a dedicated endpoint. They are tags with `subKind: "wording"` (fetch via `tag/get` with that filter). A wording tag's `parent` can point to another wording tag used as a UI group header (e.g. "Accounting framework"). The `organization_type` condition type appears almost exclusively inside wording tag values.

## Project Structure

```
tools/checklist_extract.py   # Core extraction engine
web/app.py                   # Flask backend (port 5003)
web/static/                  # Frontend (vanilla HTML/JS/CSS)
workflows/extract_checklists.md  # Full workflow SOP
.env.example                 # Credential template
```
