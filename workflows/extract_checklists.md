# Extract Checklists from Caseware Template

## Objective

Extract full checklist content (procedures, settings, visibility rules) from a CaseWare Cloud SE author template and produce an Excel workbook. Each checklist becomes a separate sheet.

## Prerequisites

- Python 3.10+
- Dependencies: `pip install -r requirements.txt`
- OAuth credentials in `.env` (copy from `.env.example`)

## How to Run

### Option A: Web UI (Recommended)

```bash
cd "Checklist extractor"
python web/app.py
```

Open http://localhost:5003 in your browser. Paste the Caseware URL and click "Extract Checklists".

### Option B: CLI

```bash
# Single checklist (URL includes #/checklist/<documentId> or #/efinancials/<documentId>)
python tools/checklist_extract.py --url "https://us.cwcloudpartner.com/cwus-develop/e/eng/<engId>/index.jsp#/checklist/<docId>"

# All checklists in engagement (URL without document fragment)
python tools/checklist_extract.py --url "https://us.cwcloudpartner.com/cwus-develop/e/eng/<engId>/index.jsp"

# Custom output path
python tools/checklist_extract.py --url "<url>" --output "my_report.xlsx"
```

### Option C: Mock Data (Testing)

```bash
python tools/checklist_extract.py --mock
```

Generates a sample report in `.tmp/checklist_extract.xlsx` with mock data covering all row types and visibility patterns.

### Option D: Discovery Mode (Field Exploration)

```bash
python tools/checklist_extract.py --discover --url "<url>"
```

Dumps raw procedure JSON to the console for identifying field paths.

## URL Format

The tool accepts two URL patterns:

- **Engagement URL**: `https://{host}/{tenant}/e/eng/{engagementId}/...`
  - Extracts ALL checklists in the engagement (one sheet per checklist)
- **Document URL**: `...#/checklist/{documentId}` or `...#/efinancials/{documentId}`
  - Extracts only that specific checklist (single sheet)

**Note**: The URL fragment references the document `id`, but procedures use the document's `content` field as their `checklistId`. The tool handles this mapping automatically.

## Output Format

Excel workbook with columns:

| Column | Content |
|--------|---------|
| A | Procedure Text (full hierarchy) |
| B | Authoritative reference (e.g. AU-C 520.05) |
| C | Assertions Tested (C, E, A, V, PD) |
| D | Lightbulb Guidance |
| E-N | Cloud Settings (override, sign-offs, response types, options) |
| O | Gap column |
| P-T | Visibility Conditions 1-5 |

### Row Types

- **Section headers** (blue background): `type=group` procedures — "Procedures", "Assessed risks", etc. Can have guidance and visibility.
- **Procedures**: `type=procedure` items with settings (explicit or inherited from checklist defaults)
- **Group procedures**: `type=procedure` with children but no explicit response sets
- **Sub-procedures**: Lettered (a, b, c) child procedures

### Data Sources

| Column | API Source |
|--------|-----------|
| Authoritative reference | `proc.attachables` where `kind="citation"` → `labels.en` |
| Assertions | `proc.tagging.baseassertion` → resolved via `tag/get` (subKind=baseassertion) → mapped to C,E,A,V,PD |
| Guidance | `proc.guidances.en` or `proc.guidance` |
| Settings E-N | `proc.settings` or inherited from `checklist/get` defaults |
| Visibility P-T | `proc.visibility.conditions[]` — types: response, rmm_rank, enum_value, boolean_value, condition_group, organization_type, consolidation |

### Assertion Mapping

Authors select from 5 assertions (C, E, A, V, PD). The API stores 14 "baseassertion" tags. Mapping rules:
- **C** = `Completeness` present
- **E** = `Existence` present
- **A** = `Accuracy` (CoT group) OR `Accuracy, valuation and allocation` (AB group) present
- **V** = `Acc,val,alloc` present AND `Accuracy` NOT present; OR all 6 AB tags present
- **PD** = `Presentation` present

### Visibility Patterns

- **None**: No visibility conditions
- **Inherited from above.**: Child inherits parent's visibility
- **Show when: [checklist] / [procedure] = [response]**: Response-based condition
- **Show when ANY/ALL of the following assertions of [area] have RMM >= [level]**: RMM rank condition
- **VISIBILITYFORM: [area] = [value]**: Enum configuration condition
- **ACCOUNTINGEST.[key] = TRUE/FALSE**: Boolean flag condition
- **Show when ALL are met: [cond1] AND [cond2]**: Multi-condition across columns P+Q+...

## Authentication

Priority order:
1. **OAuth** (preferred): Set `CW_CA_CLIENT_ID` + `CW_CA_CLIENT_SECRET` (or `CW_US_*` for US region) in `.env`
2. **Cookie fallback**: Set `CW_COOKIES` in `.env` (copy from browser DevTools)

## API Reference

For detailed API documentation, see the global knowledge docs:
- `Note visibility/docs/caseware-data-checklist-procedures.md` — full procedure data model
- `Note visibility/docs/caseware-cloud-api.md` — authentication, endpoints, URL structure
- `Note visibility/docs/caseware-data-sections-visibility.md` — visibility condition types
- `Note visibility/docs/caseware-data-components-tags.md` — tag resolution

## Troubleshooting

- **401 Unauthorised**: OAuth token expired or cookies stale. Re-generate credentials.
- **No checklists found**: The engagement may not contain checklist-type documents.
- **0 procedures**: The document ID from the URL must be mapped to its `content` field. Check the `--discover` output.
- **Permission error on output**: Close the Excel file before re-running.
- **Slow execution**: Large engagements with many checklists require many API calls for visibility ID resolution. This is expected.

## Edge Cases

- Sheet names are truncated to 31 characters (Excel limit)
- Procedures with multiple response sets produce multiple rows with merged cells
- Procedures without explicit settings inherit checklist-level defaults (response sets, note placeholder, etc.)
- Unknown visibility condition types are output as raw JSON for manual review
