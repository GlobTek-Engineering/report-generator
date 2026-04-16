npm install docx

# GTM965500P Test Report Generator

Pulls live data from the GitHub project board and generates test reports in HTML, DOCX, and PDF formats — split by output voltage (12V, 24V, 54V, Model Level) and as a combined all-models report.

---

## Requirements

| Tool | Purpose |
|------|---------|
| Node.js + `npm install docx` | DOCX generation |
| GitHub CLI (`gh`) | Fetching project data |
| Python 3 | Pretty-printing the raw JSON |
| LibreOffice (`libreoffice`) | Converting DOCX → PDF |

---

## Setup

**1. Install Node dependencies:**
```bash
npm install docx
```

**2. Set your GitHub token** in `generate_report.js`:
```js
const GITHUB_TOKEN = 'your_token_here';
```
Get your token with:
```bash
gh auth token
```
Without a token, images in comments will appear as `[image]` placeholders — all other content still generates normally.

---

## Make Commands

### `make`
Fetches fresh data from GitHub, then generates **all output formats** — DOCX, PDF, and HTML.
```bash
make
```

### `make html`
Fetches fresh data from GitHub, then generates **HTML reports only**. DOCX and PDF are skipped. Images are still downloaded and embedded.
```bash
make html
```
Use this for a fast preview or when you only need the browser-viewable version.

### `make pdf`
Fetches fresh data from GitHub, then generates **DOCX and PDF reports only**. HTML is skipped.
```bash
make pdf
```

### `make fetch`
Fetches GitHub project data and writes the JSON files, but does **not** generate any reports. Useful for inspecting the raw data before running a report.
```bash
make fetch
```
Outputs:
- `project_issues.json` — raw API response
- `GTM965500P_pretty.json` — formatted version used by the generator

### `make clean`
Deletes all generated output folders and the cached JSON files.
```bash
make clean
```
Removes: `docx/`, `pdf/`, `html/`, `.img_cache/`, `project_issues.json`, `GTM965500P_pretty.json`

---

## Output Files

Every `make` command that generates reports always fetches fresh data first. Outputs are written to:

```
docx/
  GTM965500P_Test_Report.docx           ← All models combined
  GTM965500P_Test_Report_12V.docx
  GTM965500P_Test_Report_24V.docx
  GTM965500P_Test_Report_54V.docx
  GTM965500P_Test_Report_Model_Level.docx

pdf/
  GTM965500P_Test_Report.pdf            ← All models combined
  GTM965500P_Test_Report_12V.pdf
  ...

html/
  GTM965500P_Test_Report.html           ← All models combined
  GTM965500P_Test_Report_12V.html
  ...
```

Voltage-specific reports only include items tagged to that voltage. The combined report includes everything.

---

## Report Contents

Each report contains:

- **Cover** — product name, subtitle, date generated
- **Test Summary** — status count table (OK/Resolved, For Review, Regression Req'd, In Progress, Has Issue, Invalid/Incomplete Test, Unknown) and per-voltage summary tables with links
- **Changes Required** — all comments tagged `[CHANGE]`, grouped with author and date
- **Tests Not Yet Run** — items with no comments or Unknown status
- **Detailed Test Results** — full comment threads per test, with Results, Changes, and other comments shown in collapsible sections
- **Untested Specifications** — spec body text for items that haven't been run yet

---

## Image Caching

Downloaded images are stored in `.img_cache/` and reused on subsequent runs. To force a fresh download, run `make clean` first.
