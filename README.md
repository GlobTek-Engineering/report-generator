# Test Report Generator
**https://globtek-engineering.github.io/report-generator/**

Pulls live data from a GitHub project board and generates test reports in HTML, DOCX, and PDF formats — split by output voltage and as a combined all-models report.

---

## First-Time Setup

### 1. Install dependencies

**Windows (PowerShell):**
```powershell
winget install GnuWin32.Make
winget install OpenJS.NodeJS
winget install GitHub.cli
winget install Python.Python.3
```
Close and reopen PowerShell after installing so the PATH updates.

**Linux / WSL:**
```bash
sudo apt install make nodejs npm gh python3
```

---

### 2. Allow PowerShell scripts (Windows only)
By default Windows blocks scripts from running. Run this once:
```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```
Say `Y` when prompted. This only affects your user account.

---

### 3. Clone the repo
```bash
git clone https://github.com/GlobTek-Engineering/report-generator.git
cd report-generator
```

---

### 4. Install Node dependencies
```bash
npm install docx
```

---

### 5. Create your `.env` file
```bash
make setup
```
This copies `.env.example` to `.env`. Open `.env` and set `PROJECT_ORG` and `PROJECT_NUMBER` to match the GitHub org and project number you want to generate reports for.

---

### 6. Set up your GitHub token
Your token needs the `read:project` scope to access the GitHub project board.

**Check your token:**
1. Go to https://github.com/settings/tokens
2. Click your token
3. Make sure **`read:project`** is checked
4. Scroll down and make sure your org is authorized under **Organization access** — click **Grant** if not
5. Save

**Add the missing scope if needed:**
```bash
gh auth refresh -s read:project
```

**Save your token to `.env`:**
```bash
make auth
```
This writes your token into `.env`, which is gitignored and never touches the repo. Run this once after setup, or again if your token expires.

---

## Make Commands

| Command | Description |
|---|---|
| `make` | Fetch fresh data + generate all formats (DOCX, PDF, HTML) |
| `make html` | Fetch fresh data + generate HTML reports only |
| `make pdf` | Fetch fresh data + generate DOCX and PDF only |
| `make setup` | Create `.env` from `.env.example` (run once after cloning) |
| `make auth` | Update GitHub token in `.env` |
| `make fetch` | Fetch GitHub project data only, no report generation |
| `make pages` | Copy the main HTML report to `index.html` for GitHub Pages |
| `make clean` | Delete all generated reports and cached data |

### Details

**`make html`** — fastest option for a quick preview. Images are still downloaded and embedded; DOCX and PDF are skipped.

**`make fetch`** — writes `project_raw.json` and `project_pretty.json`. Useful for inspecting the raw data without generating reports.

**`make clean`** — removes `docx/`, `pdf/`, `html/`, `.img_cache/`, `project_raw.json`, `project_pretty.json`. Does not touch `.env`.

---

## Output Files

Reports are named after the GitHub project title and written to:

```
docx/
  <ProjectName>_Test_Report.docx             ← All models combined
  <ProjectName>_Test_Report_12V.docx
  <ProjectName>_Test_Report_24V.docx
  <ProjectName>_Test_Report_54V.docx
  <ProjectName>_Test_Report_Model_Level.docx

pdf/
  <ProjectName>_Test_Report.pdf
  <ProjectName>_Test_Report_12V.pdf
  ...

html/
  <ProjectName>_Test_Report.html
  <ProjectName>_Test_Report_12V.html
  ...
```

Voltage-specific reports only include items tagged to that voltage. The combined report includes everything.

---

## Report Contents

Each report contains:

- **Cover** — project name, subtitle, date generated
- **Test Summary** — status count table and per-voltage summary tables with links
- **Changes Required** — all comments tagged `[CHANGE]`, grouped with author and date
- **Tests Not Yet Run** — items with no comments or unknown status
- **Detailed Test Results** — full comment threads per test, with Results, Changes, and other comments shown in collapsible sections
- **Untested Specifications** — spec body text for items that haven't been run yet

---

## PDF Generation

PDF conversion requires LibreOffice. If it's not installed, DOCX files still generate normally — PDF is skipped with a warning.

**Windows:** Download from https://www.libreoffice.org and make sure `soffice` is on your PATH.

**Linux / WSL:**
```bash
sudo apt install libreoffice
```

---

## Image Caching

Downloaded images are stored in `.img_cache/` and reused on subsequent runs. To force a fresh download, run `make clean` first.
