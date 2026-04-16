# GTM965500P Test Report Generator

Pulls live data from the GitHub project board and generates test reports in HTML, DOCX, and PDF formats — split by output voltage (12V, 24V, 54V, Model Level) and as a combined all-models report.

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
```powershell
git clone https://github.com/GlobTek-Engineering/report-generator.git
cd report-generator
```

---

### 4. Install Node dependencies
```powershell
npm install docx
```

---

### 5. Set up your GitHub token
Your token needs the `read:project` scope to access the GitHub project board.

**Check your token:**
1. Go to https://github.com/settings/tokens
2. Click your token
3. Make sure **`read:project`** is checked
4. Scroll to the bottom and make sure **GlobTek Engineering** is authorized under **Organization access** — click **Grant** if not
5. Save

**Add the missing scope if needed:**
```powershell
gh auth refresh -s read:project
```

**Save your token locally:**
```powershell
make auth
```
This writes your token to a local `.env` file that is gitignored — it never touches the repo. Run this once after cloning, or again if your token expires.

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

### `make auth`
Saves your GitHub token to a local `.env` file. Run this once after cloning — all subsequent `make` commands read the token from there automatically.
```bash
make auth
```

### `make fetch`
Fetches GitHub project data and writes the JSON files, but does **not** generate any reports.
```bash
make fetch
```
Outputs:
- `project_issues.json` — raw API response
- `GTM965500P_pretty.json` — formatted version used by the generator

### `make clean`
Deletes all generated output folders and cached JSON files. Does **not** delete `.env`.
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

## PDF Generation

PDF conversion requires LibreOffice. If it's not installed, DOCX files will still generate normally — PDF will just be skipped with a warning.

**Windows:** Download from https://www.libreoffice.org and make sure `soffice` is on your PATH after installing.

**Linux / WSL:**
```bash
sudo apt install libreoffice
```

---

## Image Caching

Downloaded images are stored in `.img_cache/` and reused on subsequent runs. To force a fresh download, run `make clean` first.