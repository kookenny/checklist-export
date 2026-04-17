# CaseWare Cloud Checklist Extractor

A tool that extracts audit procedure checklists from CaseWare Cloud SE author templates and exports them as formatted Excel workbooks. Each checklist becomes a separate sheet with complete metadata including procedures, assertions, guidance, settings, and visibility rules.

## Features

- **Web UI and CLI** — browser-based interface or command-line for automation
- **Single or bulk extraction** — extract one checklist by document URL or all checklists in an engagement
- **Complete metadata** — procedures, authoritative references, assertions, guidance, response settings, visibility conditions
- **Hierarchical structure** — section headers, procedures, sub-procedures rendered with proper indentation
- **Mock mode** — generate sample output without credentials for testing

## Prerequisites

- Python 3.9+
- CaseWare Cloud account with API access (OAuth credentials or browser cookies)

## Setup

```bash
# Install dependencies
pip install -r requirements.txt

# Configure credentials
cp .env.example .env
# Edit .env with your OAuth credentials or browser cookies
```

### Authentication

| Method | Environment Variables | Notes |
|--------|----------------------|-------|
| OAuth (preferred) | `CW_CA_CLIENT_ID`, `CW_CA_CLIENT_SECRET` | Per-region (CA, US, etc.) |
| Cookie fallback | `CW_COOKIES` | Copy from browser DevTools |

## Usage

### Web UI (recommended)

```bash
python web/app.py
```

Open http://localhost:5003, paste a CaseWare URL, and click **Extract Checklists**. The Excel file downloads automatically.

### Command Line

```bash
# Extract all checklists from an engagement
python tools/checklist_extract.py --url "https://ca.cwcloudpartner.com/{tenant}/e/eng/{engId}/index.jsp"

# Extract a single checklist
python tools/checklist_extract.py --url "https://ca.cwcloudpartner.com/{tenant}/e/eng/{engId}/index.jsp#/checklist/{docId}"

# Custom output path
python tools/checklist_extract.py --url "<url>" --output "my_report.xlsx"

# Mock data (no credentials needed)
python tools/checklist_extract.py --mock

# Discovery mode (dump raw JSON for debugging)
python tools/checklist_extract.py --discover --url "<url>"
```

## Output Format

The workbook contains one procedure sheet per checklist plus (when applicable) a trailing `Glossary & Dynamic Text` reference sheet.

Each procedure sheet contains columns A through S:

| Column | Content |
|--------|---------|
| A | Procedure text (hierarchical; dynamic-text markers styled blue and hyperlinked) |
| B | Authoritative reference (e.g., AU-C 520.05) |
| C | Assertions (C, E, A, V, PD) |
| D | Lightbulb guidance |
| E–H | Response settings (placeholder, type, options, display) |
| I–K | Cloud settings (input notes, notes placeholder, sign-offs) |
| L–P | Visibility conditions 1–5 |
| Q–S | Additional settings (multiple rows, response beneath, override) |

Row types are visually distinguished:
- **Blue background** — section headers
- **Regular rows** — procedures with settings
- **Lettered items (a, b, c)** — sub-procedures

### Dynamic Text

Procedures in CaseWare templates can include dynamic text spans that resolve based on engagement state (organization type, responses to other procedures, consolidation status, etc.). These appear in column A as:

- `[[value]]` — formula resolved to a value (shown in blue)
- `[[?]]` — formula is present but its condition isn't met in this engagement (still visible so the gap is discoverable)

The procedure cell hyperlinks to the `Glossary & Dynamic Text` sheet, which has two sections:
- **Glossary** — every referenced global term, one row per condition (Group / Term / Condition / Output)
- **Dynamic Text** — every `<span formula>` occurrence across all checklists, one row per condition

## Project Structure

```
tools/checklist_extract.py       # Core extraction engine
web/
  app.py                         # Flask backend (port 5003)
  static/
    index.html                   # Single-page UI
    app.js                       # Client-side logic
    styles.css                   # Styling
workflows/extract_checklists.md  # Workflow documentation
.env.example                     # Credential template
```

## License

Internal use only.
