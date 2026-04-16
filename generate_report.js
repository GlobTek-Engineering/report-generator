/**
 * GTM965500P Test Report Generator
 *
 * Usage:
 *   node generate_report.js
 *
 * Requirements:
 *   npm install docx node-fetch   (or: npm install docx axios)
 *
 * Place this file in the same folder as GTM965500P_pretty.json.
 * Output: GTM965500P_Test_Report.docx
 *
 * IMAGE SUPPORT:
 *   Images in comments are GitHub asset URLs that require authentication.
 *   To include them, run:
 *     gh auth token
 *   and paste the result into GITHUB_TOKEN below.
 *   Leave blank to skip images (they'll show as [image] placeholders).
 */

const GITHUB_TOKEN = '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'; // <-- paste your token here: gh auth token

const fs   = require('fs');
const path = require('path');
const https = require('https');
const { execSync } = require('child_process');

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, Footer, PageNumber, PageBreak,
  InternalHyperlink, Bookmark, ImageRun
} = require('docx');

// ── Config ────────────────────────────────────────────────────────────────────
const INPUT_FILE  = path.join(__dirname, 'GTM965500P_pretty.json');
const IMG_CACHE   = path.join(__dirname, '.img_cache');
const DOCX_DIR    = path.join(__dirname, 'docx');
const PDF_DIR     = path.join(__dirname, 'pdf');
const HTML_DIR    = path.join(__dirname, 'html');
const OUTPUT_FILE = path.join(DOCX_DIR, 'GTM965500P_Test_Report.docx');

// ── Mode flag  ────────────────────────────────────────────────────────────────
//   node generate_report.js            → all  (docx + pdf + html)
//   node generate_report.js --html     → html only  (images still prefetched)
//   node generate_report.js --pdf      → docx + pdf only
const _modeArg = process.argv.slice(2).find(a => a.startsWith('--'));
const MODE = _modeArg === '--html' ? 'html'
           : _modeArg === '--pdf'  ? 'pdf'
           : 'all';

// Create output folders if they don't exist
if (MODE !== 'html') {
  if (!fs.existsSync(DOCX_DIR)) fs.mkdirSync(DOCX_DIR);
  if (!fs.existsSync(PDF_DIR))  fs.mkdirSync(PDF_DIR);
}
if (MODE !== 'pdf') {
  if (!fs.existsSync(HTML_DIR)) fs.mkdirSync(HTML_DIR);
}

// ── Load & parse JSON ─────────────────────────────────────────────────────────
const raw  = fs.readFileSync(INPUT_FILE, 'utf-8').replace(/^\uFEFF/, '');
const data = JSON.parse(raw);
const nodes = data.data.organization.projectV2.items.nodes;

function extractFields(node) {
  const fv = (node.fieldValues || {}).nodes || [];
  let title = '', status = '', voltage = '', category = '';
  for (const f of fv) {
    if (f && f.field) {
      if (f.field.name === 'Title')         title    = f.text || '';
      if (f.field.name === 'Status')        status   = f.name || '';
      if (f.field.name === 'Output Voltage') voltage  = f.name || '';
      if (f.field.name === 'Category')      category = f.name || '';
    }
  }
  const c = node.content || {};
  return {
    number:   c.number || 0,
    title:    title || c.title || '',
    status, voltage, category,
    state:    c.state || '',
    body:     c.body || '',
    comments: (c.comments || {}).nodes || []
  };
}

const allItems = nodes.map(extractFields);
const items = allItems.filter(item => item.comments.length > 0);

// ── Image downloading ─────────────────────────────────────────────────────────
if (!fs.existsSync(IMG_CACHE)) fs.mkdirSync(IMG_CACHE);

function downloadImage(url) {
  return new Promise((resolve, reject) => {
    const filename = url.split('/').pop().split('?')[0];
    const cachePath = path.join(IMG_CACHE, filename);

    // Re-check if cached file is valid image (not HTML/XML from a bad previous run)
    if (fs.existsSync(cachePath)) {
      const buf = Buffer.alloc(4);
      const fd = fs.openSync(cachePath, 'r');
      fs.readSync(fd, buf, 0, 4, 0);
      fs.closeSync(fd);
      const isImage = (buf[0] === 0x89 && buf[1] === 0x50) || // PNG
                      (buf[0] === 0xFF && buf[1] === 0xD8);    // JPEG
      if (isImage) { resolve(cachePath); return; }
      fs.unlinkSync(cachePath); // delete bad cache
    }

    try {
      execSync(
        `curl -s -L -H "Authorization: token ${GITHUB_TOKEN}" ` +
        `"${url}" -o "${cachePath}"`,
        { stdio: 'pipe' }
      );
      resolve(cachePath);
    } catch (e) {
      reject(e);
    }
  });
}

function getImgType(url) {
  // First try URL extension
  const ext = url.split('.').pop().toLowerCase().split('?')[0];
  const extMap = { jpg: 'jpg', jpeg: 'jpg', png: 'png', gif: 'gif', webp: 'png' };
  if (extMap[ext]) return extMap[ext];

  // Fall back to reading magic bytes from cached file
  const filename = url.split('/').pop().split('?')[0];
  const cachePath = path.join(IMG_CACHE, filename);
  if (fs.existsSync(cachePath)) {
    const buf = Buffer.alloc(4);
    const fd = fs.openSync(cachePath, 'r');
    fs.readSync(fd, buf, 0, 4, 0);
    fs.closeSync(fd);
    if (buf[0] === 0x89 && buf[1] === 0x50) return 'png';  // PNG
    if (buf[0] === 0xFF && buf[1] === 0xD8) return 'jpg';  // JPEG
    if (buf[0] === 0x47 && buf[1] === 0x49) return 'gif';  // GIF
    if (buf[0] === 0x52 && buf[1] === 0x49) return 'png';  // WEBP (treat as png)
  }
  return 'jpg'; // GitHub assets are usually JPEG
}

// Extract all image URLs from a text block
function extractImgUrls(text) {
  const matches = [];
  const re = /<img[^>]+src="([^"]+)"/gi;
  let m;
  while ((m = re.exec(text)) !== null) matches.push(m[1]);
  return matches;
}

// Pre-download all images referenced across all comments
async function prefetchAllImages() {
  if (!GITHUB_TOKEN) {
    console.log('No GITHUB_TOKEN set — skipping image downloads. Images will show as [image] placeholders.');
    return;
  }
  const allUrls = new Set();
  for (const item of items) {
    for (const comment of item.comments) {
      for (const url of extractImgUrls(comment.body || '')) allUrls.add(url);
    }
    for (const url of extractImgUrls(item.body || '')) allUrls.add(url);
  }
  console.log(`Downloading ${allUrls.size} images...`);
  let done = 0;
  for (const url of allUrls) {
    try {
      await downloadImage(url);
      done++;
      process.stdout.write(`\r  ${done}/${allUrls.size}`);
    } catch (e) {
      console.warn(`\n  Failed: ${url} — ${e.message}`);
    }
  }
  console.log(`\nDone downloading images.`);
}

// ── Colors ────────────────────────────────────────────────────────────────────
const C = {
  darkBlue: '1F3864', medBlue: '2E75B6', lightBlue: 'D6E4F0',
  headerBg: '1F3864', white: 'FFFFFF', altRow: 'EBF3FB',
  passGreen: 'E2EFDA', passText: '375623',
  reviewAmber: 'FFF2CC', reviewText: '7F6000',
  regressionRed: 'FCE4D6', regressionText: '843C0C',
  issueRed: 'FCE4D6', issueText: '843C0C',
  progressBlue: 'DEEAF1', progressText: '1F3864',
  unknownGray: 'F2F2F2', unknownText: '444444',
  invalidBg: 'EDE7F6', invalidText: '4527A0',
  border: 'BFBFBF', dataBorder: 'CCCCCC', dataHeader: 'D6E4F0'
};

function statusColors(s) {
  if (s === 'OK / Resolved')     return { bg: C.passGreen,      fg: C.passText };
  if (s === 'Marked for review') return { bg: C.reviewAmber,    fg: C.reviewText };
  if (s === "Regression Req'd")  return { bg: C.regressionRed,  fg: C.regressionText };
  if (s === 'Has Issue')         return { bg: C.issueRed,        fg: C.issueText };
  if (s === 'In Progress')       return { bg: C.progressBlue,   fg: C.progressText };
  if (s === 'Invalid/Incomplete Test') return { bg: C.invalidBg,     fg: C.invalidText };
  return { bg: C.unknownGray, fg: C.unknownText };
}

// ── Comment tag detection ─────────────────────────────────────────────────────
function getCommentTag(body) {
  const m = (body || '').trim().match(/^\[(RESULT|CHANGE)\]/i);
  return m ? m[1].toUpperCase() : null;
}

function stripTag(body) {
  return (body || '').replace(/^\[(RESULT|CHANGE)\]\s*/i, '').trim();
}
const bdr         = { style: BorderStyle.SINGLE, size: 4, color: C.border };
const allBorders  = { top: bdr, bottom: bdr, left: bdr, right: bdr };
const dataBdr     = { style: BorderStyle.SINGLE, size: 2, color: C.dataBorder };
const dataAllBorders = { top: dataBdr, bottom: dataBdr, left: dataBdr, right: dataBdr };

function cell(text, { bold=false, fg='000000', bg='FFFFFF', width=1000, align=AlignmentType.LEFT, size=18 }={}) {
  return new TableCell({
    borders: allBorders, width: { size: width, type: WidthType.DXA },
    shading: { fill: bg, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ alignment: align, children: [
      new TextRun({ text: String(text || ''), bold, color: fg, size, font: 'Arial' })
    ]})]
  });
}

function hdrCell(text, width) {
  return cell(text, { bold: true, fg: C.white, bg: C.headerBg, width });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 120 },
    children: [new TextRun({ text, bold: true, size: 36, font: 'Arial', color: C.darkBlue })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 80 },
    children: [new TextRun({ text, bold: true, size: 26, font: 'Arial', color: C.medBlue })]
  });
}

function para(text, { bold=false, size=20, color='000000', indent=0, italic=false }={}) {
  return new Paragraph({
    spacing: { before: 40, after: 40 },
    indent: indent ? { left: indent } : undefined,
    children: [new TextRun({ text: String(text||''), bold, size, font: 'Arial', color, italics: italic })]
  });
}

function divider() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.medBlue, space: 1 } },
    spacing: { before: 100, after: 100 }, children: []
  });
}

function thinRule() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: 'CCCCCC', space: 1 } },
    spacing: { before: 120, after: 120 }, children: []
  });
}

function spacer() {
  return new Paragraph({ spacing: { before: 60, after: 60 }, children: [] });
}

// ── Markdown → DOCX renderer ──────────────────────────────────────────────────
function isTableRow(line) {
  // Matches | col | col | style AND col | col | col style (no leading pipe)
  return /^\s*\|/.test(line) || /^[^\|#\-\s>][^\n]*\|[^\n]*\|/.test(line.trim());
}
function isSepRow(line) {
  const t = line.replace(/┬á/g, ' ').trim();
  // Must contain at least one '-' and consist entirely of |, -, :, spaces
  return /^[\|\s\-:]+$/.test(t) && t.includes('-') && t.includes('|');
}

function parseMarkdownTableBlock(lines) {
  const dataRows = lines.filter(l => !isSepRow(l));
  if (dataRows.length < 1) return null;

  // Parse each row into cells, keeping img tags intact
  const parsedRows = dataRows.map(line => {
    // Check BEFORE any cleaning whether this is a continuation row
    // Only \xa0 (non-breaking space) or ┬á (garbled nbsp) signal an empty first cell
    // Regular leading spaces before | are just sloppy formatting — NOT continuation rows
    // Must test on the raw line (before trim) because trim() strips \xa0
    const isContinuationRow = /^[\s]*[\u00a0┬á]/.test(line);

    let l = line
      .replace(/┬á/g, ' ')
      .replace(/\u00a0/g, ' ')
      .replace(/ΓåÆ/g, '\u2192')
      .replace(/┬▒/g, '\u00b1')
      .replace(/┬╡/g, '\u00b5')
      .trim();

    if (isContinuationRow) {
      // Strip any leading pipe artifact, rebuild with explicit empty first cell
      l = l.replace(/^\|/, '').trim();
      l = '| | ' + l;
    }

    // Normalize: ensure leading AND trailing pipes
    if (!l.startsWith('|')) l = '|' + l;
    if (!l.endsWith('|')) l = l + '|';

    // Split on | — drop first and last empty strings from outer pipes
    const parts = l.split('|');
    parts.shift();
    parts.pop();
    return parts.map(c => c.trim());
  });

  // Use header row column count as the canonical column count
  const numCols = parsedRows[0] ? parsedRows[0].length : Math.max(...parsedRows.map(r => r.length));
  if (numCols === 0) return null;

  // Determine column widths — give image columns 3x more space than text columns
  const hasImage = Array(numCols).fill(false);
  for (const row of parsedRows) {
    for (let ci = 0; ci < row.length; ci++) {
      if (/<img/i.test(row[ci] || '')) hasImage[ci] = true;
    }
  }
  const imgCols   = hasImage.filter(Boolean).length;
  const textCols  = numCols - imgCols;
  const totalDxa  = 9000;
  // Image cols get 3 units, text cols get 1 unit
  const unit      = Math.floor(totalDxa / (textCols + imgCols * 3));
  const colWidths = hasImage.map(isImg => isImg ? unit * 3 : unit);
  // Fix rounding so columns sum exactly to totalDxa
  const diff = totalDxa - colWidths.reduce((a, b) => a + b, 0);
  colWidths[colWidths.length - 1] += diff;

  const tableRows = parsedRows.map((row, rowIdx) => {
    const isHeader = rowIdx === 0;
    const isLastRow = rowIdx === parsedRows.length - 1;
    const cells = [];
    for (let ci = 0; ci < numCols; ci++) {
      const cellText = row[ci] || '';
      // Check if this cell contains an image
      const imgMatch = cellText.match(/<img[^>]+src="([^"]+)"/i);
      let cellChildren;
      if (imgMatch && !isHeader) {
        const imgPara = renderImg(imgMatch[1], 0, true, colWidths[ci]);
        cellChildren = [imgPara];
      } else {
        const cleanText = cellText.replace(/<[^>]+>/g, '').trim();
        cellChildren = [new Paragraph({ keepNext: !isLastRow, alignment: AlignmentType.CENTER, children: [
          new TextRun({ text: cleanText, bold: isHeader, size: 17, font: 'Arial',
            color: isHeader ? C.darkBlue : '222222' })
        ]})];
      }
      cells.push(new TableCell({
        borders: dataAllBorders,
        width: { size: colWidths[ci], type: WidthType.DXA },
        shading: { fill: isHeader ? C.dataHeader : (rowIdx % 2 === 0 ? 'FFFFFF' : C.altRow), type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        verticalAlign: VerticalAlign.CENTER,
        children: cellChildren
      }));
    }
    return new TableRow({ cantSplit: true, children: cells });
  });

  return new Table({
    width: { size: colWidths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: colWidths,
    rows: tableRows
  });
}

// Render an <img> tag as an ImageRun if we have the file, else a placeholder paragraph
function renderImg(url, indent, inTable = false, cellWidthDxa = 0) {
  const filename = url.split('/').pop().split('?')[0];
  const cachePath = path.join(IMG_CACHE, filename);
  if (fs.existsSync(cachePath)) {
    try {
      const imgData = fs.readFileSync(cachePath);
      const type = getImgType(url);
      // If inside a table cell, fit to cell width (DXA → points: divide by 20, subtract padding)
      // Otherwise use full page width
      let maxWidthPt;
      if (inTable && cellWidthDxa > 0) {
        maxWidthPt = Math.max(120, Math.floor(cellWidthDxa / 20) - 12); // min 120pt, subtract cell margin
      } else if (inTable) {
        maxWidthPt = 280;
      } else {
        maxWidthPt = 500;
      }
      return new Paragraph({
        spacing: { before: 60, after: 60 },
        indent: indent ? { left: indent } : undefined,
        children: [new ImageRun({
          type,
          data: imgData,
          transformation: { width: maxWidthPt, height: Math.round(maxWidthPt * 0.6) },
          altText: { title: 'Oscilloscope screenshot', description: url, name: filename }
        })]
      });
    } catch (e) {
      // fall through to placeholder
    }
  }
  // Placeholder if image not available
  return new Paragraph({
    spacing: { before: 30, after: 30 },
    indent: indent ? { left: indent } : undefined,
    children: [new TextRun({ text: `[image: ${url}]`, size: 16, font: 'Arial', color: '999999', italics: true })]
  });
}

function renderMarkdown(text, indent = 0) {
  if (!text) return [];
  // Fix garbled characters
  text = text.replace(/ΓåÆ/g, '\u2192')
             .replace(/┬▒/g, '\u00b1')
             .replace(/┬╡/g, '\u00b5')
             .replace(/┬á/g, ' ');

  const lines = text.split(/\r?\n/);
  const elements = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];

    // Image tag
    const imgMatch = line.match(/<img[^>]+src="([^"]+)"/i);
    if (imgMatch) {
      elements.push(renderImg(imgMatch[1], indent));
      i++; continue;
    }

    // Markdown table
    if (isTableRow(line)) {
      const tableLines = [];
      while (i < lines.length && (isTableRow(lines[i]) || isSepRow(lines[i]))) {
        tableLines.push(lines[i]);
        i++;
      }
      const tbl = parseMarkdownTableBlock(tableLines);
      if (tbl) {
        elements.push(new Paragraph({ spacing: { before: 60, after: 4 }, children: [] }));
        elements.push(tbl);
        elements.push(new Paragraph({ spacing: { before: 4, after: 60 }, children: [] }));
      }
      continue;
    }

    // Horizontal rule
    if (/^[-=]{3,}$/.test(line.trim())) { i++; continue; }

    // Headings
    const hMatch = line.match(/^(#{1,4})\s+(.+)/);
    if (hMatch) {
      const level = hMatch[1].length;
      const htxt = hMatch[2].replace(/\*\*/g, '').trim();

      if (level >= 3) {
        // ### = table title: centered, bold, dark blue, underlined
        // #### = sub-title: slightly smaller
        const sz = level === 3 ? 22 : 19;
        const topSpace = level === 3 ? 200 : 140;
        elements.push(new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: topSpace, after: 60 },
          border: {
            bottom: { style: BorderStyle.SINGLE, size: level === 3 ? 6 : 3, color: C.medBlue, space: 4 }
          },
          children: [new TextRun({ text: htxt, bold: true, size: sz, font: 'Arial', color: C.darkBlue })]
        }));
      } else {
        // # / ## = regular section heading
        const sz = level === 1 ? 22 : 20;
        elements.push(new Paragraph({
          spacing: { before: 100, after: 40 },
          indent: indent ? { left: indent } : undefined,
          children: [new TextRun({ text: htxt, bold: true, size: sz, font: 'Arial', color: C.darkBlue })]
        }));
      }
      i++; continue;
    }

    // Blank line
    if (!line.trim()) {
      elements.push(new Paragraph({ spacing: { before: 20, after: 20 }, children: [] }));
      i++; continue;
    }

    // Strip remaining HTML, clean up bold markers
    const cleaned = line.replace(/<[^>]+>/g, '').replace(/\*\*/g, '').trim();
    if (cleaned) {
      elements.push(new Paragraph({
        spacing: { before: 30, after: 30 },
        indent: indent ? { left: indent } : undefined,
        children: [new TextRun({ text: cleaned, size: 18, font: 'Arial', color: '222222' })]
      }));
    }
    i++;
  }
  return elements;
}

// ── Document assembly ─────────────────────────────────────────────────────────
// ── Write helper: pack buffer, fix bookmark IDs, save to outPath ─────────────
async function writeDoc(doc, outPath) {
  const buf    = await Packer.toBuffer(doc);
  const tmpDoc = outPath + '.tmp.docx';
  const tmpDir = outPath + '.tmp_unpack';
  fs.writeFileSync(tmpDoc, buf);

  execSync(`rmdir /s /q "${tmpDir}" 2>nul || rm -rf "${tmpDir}"; mkdir "${tmpDir}" 2>nul || mkdir -p "${tmpDir}"`);

  try {
    execSync(`cd "${tmpDir}" && unzip -q "${tmpDoc}"`);
  } catch {
    execSync(`powershell -Command "Expand-Archive -Path '${tmpDoc}' -DestinationPath '${tmpDir}' -Force"`);
  }

  let xml = fs.readFileSync(path.join(tmpDir, 'word', 'document.xml'), 'utf-8');
  let startId = 1000, endId = 1000;
  xml = xml.replace(/(<w:bookmarkStart[^>]*?)w:id="[^"]*"/g, (m, pre) => `${pre}w:id="${startId++}"`);
  xml = xml.replace(/(<w:bookmarkEnd[^>]*)w:id="[^"]*"/g,   (m, pre) => `${pre}w:id="${endId++}"`);
  fs.writeFileSync(path.join(tmpDir, 'word', 'document.xml'), xml);

  try {
    execSync(`cd "${tmpDir}" && zip -q -r "${outPath}" .`);
  } catch {
    execSync(`powershell -Command "Compress-Archive -Path '${tmpDir}\\*' -DestinationPath '${outPath}' -Force"`);
  }

  try { fs.rmSync(tmpDir, { recursive: true }); } catch {}
  try { fs.unlinkSync(tmpDoc); } catch {}

  // Convert to PDF using LibreOffice, output to pdf/ folder
  try {
    execSync(`libreoffice --headless --convert-to pdf "${outPath}" --outdir "${PDF_DIR}"`, { stdio: 'pipe' });
    const pdfName = path.basename(outPath).replace('.docx', '.pdf');
    console.log(`  PDF:     ${path.join(PDF_DIR, pdfName)}`);
  } catch (e) {
    console.warn(`  PDF conversion failed: ${e.message}`);
  }

  console.log(`  Written: ${outPath}`);
}

async function buildDoc(voltFilter = null) {
  const voltOrder = ['12V', '24V', '54V', 'Model Level'];
  const grouped = {};
  for (const item of items) {
    const volt = item.voltage || 'Model Level';
    if (!grouped[volt]) grouped[volt] = [];
    grouped[volt].push(item);
  }

  // When filtering to a single voltage, only use that voltage's items for counts
  const activeVolts = voltFilter ? [voltFilter] : voltOrder;
  const activeItems = voltFilter ? (grouped[voltFilter] || []) : items;

  const allActiveItems_counts = voltFilter ? allItems.filter(i => (i.voltage || 'Model Level') === voltFilter) : allItems;
  const counts = {};
  for (const item of allActiveItems_counts) counts[item.status || 'Unknown'] = (counts[item.status || 'Unknown'] || 0) + 1;
  const total = allActiveItems_counts.length;

  const subtitle = voltFilter ? `${voltFilter} Model` : 'All Models';

  const children = [];

  // ── Cover ───────────────────────────────────────────────────────────────────
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 720, after: 160 },
    children: [new TextRun({ text: 'GTM965500P', bold: true, size: 80, font: 'Arial', color: C.darkBlue })]
  }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 120 },
    children: [new TextRun({ text: 'Test Report', size: 52, font: 'Arial', color: C.medBlue })]
  }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 },
    children: [new TextRun({ text: subtitle, size: 32, font: 'Arial', color: C.medBlue })]
  }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: 'GlobTek Engineering', size: 28, font: 'Arial', color: '555555' })]
  }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 600 },
    children: [new TextRun({ text: `Generated: ${new Date().toLocaleDateString('en-US', {year:'numeric',month:'long',day:'numeric'})}`, size: 22, font: 'Arial', color: '888888' })]
  }));
  children.push(divider());

  // ── Summary ─────────────────────────────────────────────────────────────────
  children.push(h1('Test Summary'));
  children.push(para(`Project: GTM965500P  |  ${subtitle}  |  Total Test Items: ${total}`, { bold: true }));
  children.push(spacer());

  // Status count table
  const sumColW = [1260, 1260, 1260, 1260, 1260, 1260, 1260, 1260];
  children.push(new Table({
    width: { size: 10080, type: WidthType.DXA }, columnWidths: sumColW,
    rows: [
      new TableRow({ cantSplit: true, children: [
        hdrCell('Status', 1260), hdrCell('OK / Resolved', 1260), hdrCell('For Review', 1260),
        hdrCell("Regression Req'd", 1260), hdrCell('In Progress', 1260), hdrCell('Has Issue', 1260), hdrCell('Invalid/Incomplete Test', 1260), hdrCell('Unknown', 1260),
      ]}),
      new TableRow({ cantSplit: true, children: [
        cell('Count', { bold: true, bg: C.altRow, width: 1260 }),
        cell(String(counts['OK / Resolved']||0),     { align: AlignmentType.CENTER, bg: C.passGreen,      fg: C.passText,       bold: true, width: 1260 }),
        cell(String(counts['Marked for review']||0),  { align: AlignmentType.CENTER, bg: C.reviewAmber,    fg: C.reviewText,     bold: true, width: 1260 }),
        cell(String(counts["Regression Req'd"]||0),   { align: AlignmentType.CENTER, bg: C.regressionRed,  fg: C.regressionText, bold: true, width: 1260 }),
        cell(String(counts['In Progress']||0),        { align: AlignmentType.CENTER, bg: C.progressBlue,   fg: C.progressText,   bold: true, width: 1260 }),
        cell(String(counts['Has Issue']||0),          { align: AlignmentType.CENTER, bg: C.issueRed,       fg: C.issueText,      bold: true, width: 1260 }),
        cell(String(counts['Invalid/Incomplete Test']||0), { align: AlignmentType.CENTER, bg: C.invalidBg,      fg: C.invalidText,    bold: true, width: 1260 }),
        cell(String(counts['Unknown']||0),            { align: AlignmentType.CENTER, bg: C.unknownGray,    fg: C.unknownText,    bold: true, width: 1260 }),
      ]}),
    ]
  }));
  children.push(spacer());

  // ── Page break: cover + stats on page 1, voltage tables start on page 2 ──────
  children.push(new Paragraph({ children: [new PageBreak()] }));

  // Per-voltage summary tables
  const catOrder = ['Input', 'Main Output', 'Standby Output', 'Fan Output', 'Protections',
                    'Environmental / Reliability', 'Safety', 'EMC', 'PFC'];

  for (const volt of activeVolts) {
    const catItems = (grouped[volt] || []).slice().sort((a, b) => a.number - b.number);
    if (!catItems.length) continue;

    children.push(h2(volt));
    const colW = [6239, 2320, 1521];
    const rows = [new TableRow({ cantSplit: true, tableHeader: true, children: [
      hdrCell('Test Item', 6239), hdrCell('Status', 2320), hdrCell('Link', 1521),
    ]})];

    // Group items by category, preserve catOrder
    const byCategory = {};
    for (const item of catItems) {
      const cat = item.category || 'Other';
      if (!byCategory[cat]) byCategory[cat] = [];
      byCategory[cat].push(item);
    }
    const orderedCats = [...catOrder.filter(c => byCategory[c]), ...Object.keys(byCategory).filter(c => !catOrder.includes(c))];

    let rowIdx = 0;
    for (const cat of orderedCats) {
      const catGroup = byCategory[cat];
      if (!catGroup || !catGroup.length) continue;

      // Category header row — spans all columns
      rows.push(new TableRow({ cantSplit: true, children: [
        new TableCell({
          borders: allBorders,
          columnSpan: 3,
          width: { size: 10080, type: WidthType.DXA },
          shading: { fill: C.medBlue, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 120, right: 120 },
          children: [new Paragraph({ children: [
            new TextRun({ text: cat, bold: true, size: 17, font: 'Arial', color: C.white })
          ]})]
        })
      ]}));

      for (const item of catGroup) {
        const bg = rowIdx % 2 === 0 ? 'FFFFFF' : C.altRow;
        rowIdx++;
        const bookmarkId = `issue_${item.number}`;
        const { bg: sBg, fg: sFg } = statusColors(item.status);
        const linkCell = new TableCell({
          borders: allBorders, width: { size: 1521, type: WidthType.DXA },
          shading: { fill: bg, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new InternalHyperlink({
              anchor: bookmarkId,
              children: [new TextRun({ text: 'View \u2192', style: 'Hyperlink', size: 16, font: 'Arial' })]
            })]
          })]
        });
        rows.push(new TableRow({ cantSplit: true, children: [
          cell(item.title,              { bg, width: 6239, size: 18 }),
          cell(item.status || 'Unknown', { bg: sBg, fg: sFg, width: 2320, size: 16 }),
          linkCell,
        ]}));
      }
    }

    children.push(new Table({ width: { size: 10080, type: WidthType.DXA }, columnWidths: colW, rows }));
    children.push(spacer());
    children.push(new Paragraph({ children: [new PageBreak()] }));
  }

  // ── Changes Required section ──────────────────────────────────────────────────
  const allChanges = [];
  for (const volt of activeVolts) {
    for (const item of (grouped[volt] || [])) {
      for (const comment of item.comments) {
        if (getCommentTag(comment.body) === 'CHANGE') {
          allChanges.push({ item, comment });
        }
      }
    }
  }

  if (allChanges.length > 0) {
    children.push(h1('Changes Required'));
    children.push(divider());
    children.push(para(`${allChanges.length} change${allChanges.length > 1 ? 's' : ''} flagged across all tests.`));
    children.push(spacer());

    const chgColW = [800, 3800, 4200];
    const chgRows = [new TableRow({ cantSplit: true, tableHeader: true, children: [
      hdrCell('Voltage', 800), hdrCell('Test', 3800), hdrCell('Change Description', 4200),
    ]})];

    allChanges.forEach(({ item, comment }, i) => {
      const bg = i % 2 === 0 ? 'FFFFFF' : C.altRow;
      const bookmarkId = `issue_${item.number}`;
      const author = (comment.author || {}).login || '';
      const date   = (comment.createdAt || '').substring(0, 10);
      const desc   = stripTag(comment.body)
        .replace(/<[^>]+>/g, '').replace(/\*\*/g, '').replace(/\n+/g, ' ').trim();
      const shortDesc = desc.length > 300 ? desc.substring(0, 297) + '...' : desc;

      // Test name cell with hyperlink
      const testCell = new TableCell({
        borders: allBorders, width: { size: 3800, type: WidthType.DXA },
        shading: { fill: bg, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ children: [
          new InternalHyperlink({
            anchor: bookmarkId,
            children: [new TextRun({ text: item.title, style: 'Hyperlink', size: 17, font: 'Arial' })]
          }),
          new TextRun({ text: `  ·  ${author}  ·  ${date}`, size: 15, font: 'Arial', color: '888888', italics: true })
        ]})]
      });

      chgRows.push(new TableRow({ cantSplit: true, children: [
        cell(item.voltage || 'Model Level', { bg: C.regressionRed, fg: C.regressionText, width: 800, align: AlignmentType.CENTER, size: 16, bold: true }),
        testCell,
        cell(shortDesc, { bg, width: 4200, size: 16 }),
      ]}));
    });

    children.push(new Table({ width: { size: 8800, type: WidthType.DXA }, columnWidths: chgColW, rows: chgRows }));
    children.push(spacer());
  }

  // ── Not Yet Run section ───────────────────────────────────────────────────────
  // Items with no comments OR Unknown status — filtered to active voltages
  const allActiveItems = voltFilter
    ? allItems.filter(i => (i.voltage || 'Model Level') === voltFilter)
    : allItems;
  const unknownItems = allActiveItems.filter(i => !i.comments.length || !i.status || i.status === 'Unknown');

  // Build nyrGrouped outside the if so the untested specs section can also use it
  const nyrGrouped = {};
  for (const item of unknownItems) {
    const volt = item.voltage || 'Model Level';
    if (!nyrGrouped[volt]) nyrGrouped[volt] = [];
    nyrGrouped[volt].push(item);
  }

  if (unknownItems.length > 0) {
    children.push(new Paragraph({ children: [new PageBreak()] }));
    children.push(h1('Tests Not Yet Run'));
    children.push(divider());
    children.push(para(`${unknownItems.length} test${unknownItems.length > 1 ? 's' : ''} have not yet been run.`));
    children.push(spacer());

    const nyrColW = [7440, 1320, 1320];

    for (const volt of voltOrder) {
      const voltItems = (nyrGrouped[volt] || []).slice().sort((a, b) => a.number - b.number);
      if (!voltItems.length) continue;

      children.push(h2(volt));

      const rows = [new TableRow({ cantSplit: true, tableHeader: true, children: [
        hdrCell('Test Item', 7440), hdrCell('Status', 1320), hdrCell('Link', 1320),
      ]})];

      // Group by category
      const byCategory = {};
      for (const item of voltItems) {
        const cat = item.category || 'Other';
        if (!byCategory[cat]) byCategory[cat] = [];
        byCategory[cat].push(item);
      }
      const orderedCats = [...catOrder.filter(c => byCategory[c]), ...Object.keys(byCategory).filter(c => !catOrder.includes(c))];

      let rowIdx = 0;
      for (const cat of orderedCats) {
        const catGroup = byCategory[cat];
        if (!catGroup || !catGroup.length) continue;

        // Category header row
        rows.push(new TableRow({ cantSplit: true, children: [
          new TableCell({
            borders: allBorders, columnSpan: 3,
            width: { size: 10080, type: WidthType.DXA },
            shading: { fill: C.medBlue, type: ShadingType.CLEAR },
            margins: { top: 60, bottom: 60, left: 120, right: 120 },
            children: [new Paragraph({ children: [
              new TextRun({ text: cat, bold: true, size: 17, font: 'Arial', color: C.white })
            ]})]
          })
        ]}));

        for (const item of catGroup) {
          const bg = rowIdx % 2 === 0 ? 'FFFFFF' : C.altRow;
          rowIdx++;
          const linkCell = new TableCell({
            borders: allBorders, width: { size: 1320, type: WidthType.DXA },
            shading: { fill: bg, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new InternalHyperlink({
                anchor: `issue_${item.number}`,
                children: [new TextRun({ text: 'View \u2192', style: 'Hyperlink', size: 16, font: 'Arial' })]
              })]
            })]
          });
          rows.push(new TableRow({ cantSplit: true, children: [
            cell(item.title, { bg, width: 7440, size: 18 }),
            (() => { const { bg: sBg, fg: sFg } = statusColors(item.status); return cell(item.status || 'Unknown', { bg: sBg, fg: sFg, width: 1320, size: 16, align: AlignmentType.CENTER }); })(),
            linkCell,
          ]}));
        }
      }

      children.push(new Table({ width: { size: 10080, type: WidthType.DXA }, columnWidths: nyrColW, rows }));
      children.push(spacer());
      children.push(new Paragraph({ children: [new PageBreak()] }));
    }
  }

  // ── Detail section ───────────────────────────────────────────────────────────
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(h1('Detailed Test Results'));
  children.push(divider());

  for (const volt of activeVolts) {
    const catItems = (grouped[volt] || []).slice().sort((a, b) => a.number - b.number);
    if (!catItems.length) continue;

    children.push(h2(volt));

    for (const item of catItems) {
      const { bg: sBg, fg: sFg } = statusColors(item.status);
      const bookmarkId = `issue_${item.number}`;

      // Issue header with bookmark anchor
      children.push(new Table({
        width: { size: 9600, type: WidthType.DXA },
        columnWidths: [800, 5300, 1400, 2100],
        rows: [new TableRow({ cantSplit: true, children: [
          cell(item.voltage || 'Model Level', { bold: true, bg: C.darkBlue, fg: C.white, width: 800, align: AlignmentType.CENTER, size: 20 }),
          new TableCell({
            borders: allBorders, width: { size: 5300, type: WidthType.DXA },
            shading: { fill: C.lightBlue, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({ children: [
              new Bookmark({ id: bookmarkId, children: [
                new TextRun({ text: item.title, bold: true, size: 20, font: 'Arial', color: C.darkBlue })
              ]})
            ]})]
          }),
          cell(item.state || '',         { bg: C.lightBlue, fg: C.darkBlue, width: 1400, align: AlignmentType.CENTER, size: 16 }),
          cell(item.status || 'Unknown', { bg: sBg, fg: sFg, width: 2100, size: 16 }),
        ]})]
      }));

      // Spec
      if (item.body && item.body.trim()) {
        children.push(spacer());
        children.push(sectionLabel('Specification / Procedure', C.medBlue, 'EBF3FB'));
        children.push(...renderMarkdown(item.body, 200));
      }

      // Section label helper — colored left-border accent bar
      function sectionLabel(text, color, bgColor) {
        return new Table({
          width: { size: 9600, type: WidthType.DXA },
          columnWidths: [9600],
          rows: [new TableRow({ cantSplit: true, children: [
            new TableCell({
              borders: {
                top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
                right: { style: BorderStyle.NONE },
                left: { style: BorderStyle.SINGLE, size: 16, color },
              },
              shading: { fill: bgColor, type: ShadingType.CLEAR },
              margins: { top: 60, bottom: 60, left: 160, right: 120 },
              children: [new Paragraph({ children: [
                new TextRun({ text, bold: true, size: 19, font: 'Arial', color })
              ]})]
            })
          ]})]
        });
      }

      // Comments — split by tag
      const changeComments = item.comments.filter(c => getCommentTag(c.body) === 'CHANGE');
      const resultComments = item.comments.filter(c => getCommentTag(c.body) === 'RESULT');
      const otherComments  = item.comments.filter(c => !getCommentTag(c.body));

      function renderCommentBlock(comments, label, labelColor, labelBg, threadBg) {
        if (!comments.length) return;

        children.push(spacer());
        children.push(sectionLabel(label, labelColor, labelBg));

        // Wrap all comments in a single gray-background table cell
        const threadChildren = [];
        comments.forEach((comment, idx) => {
          const isLast  = idx === comments.length - 1;
          const author  = (comment.author || {}).login || 'unknown';
          const date    = (comment.createdAt || '').substring(0, 10);
          const byline  = isLast ? `${author}  ·  ${date}  [LATEST]` : `${author}  ·  ${date}`;

          // Separator between comments (not before the first)
          if (idx > 0) {
            threadChildren.push(new Paragraph({
              border: { top: { style: BorderStyle.SINGLE, size: 2, color: 'CCCCCC', space: 1 } },
              spacing: { before: 80, after: 0 }, children: []
            }));
          }

          threadChildren.push(new Paragraph({
            spacing: { before: idx === 0 ? 0 : 80, after: 10 },
            children: [new TextRun({ text: byline, bold: isLast, size: 17, font: 'Arial',
              color: isLast ? C.darkBlue : '555555', italics: !isLast })]
          }));
          threadChildren.push(...renderMarkdown(stripTag(comment.body), 0));
        });

        children.push(new Table({
          width: { size: 9600, type: WidthType.DXA },
          columnWidths: [9600],
          rows: [new TableRow({ cantSplit: false, children: [
            new TableCell({
              borders: {
                top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
                left: { style: BorderStyle.SINGLE, size: 8, color: labelColor },
                right: { style: BorderStyle.NONE },
              },
              shading: { fill: threadBg, type: ShadingType.CLEAR },
              margins: { top: 120, bottom: 120, left: 240, right: 240 },
              children: threadChildren
            })
          ]})]
        }));
      }

      renderCommentBlock(changeComments, 'Changes Required', C.regressionText, 'FCE4D6', 'FEF4F1');
      renderCommentBlock(resultComments, 'Results',          C.passText,       'E2EFDA', 'F4FAF0');
      if (otherComments.length) {
        renderCommentBlock(otherComments, `Comments (${otherComments.length})`, C.medBlue, 'D6E4F0', 'F2F6FB');
      }

      children.push(new Paragraph({ children: [new PageBreak()] }));
    }

    children.push(spacer());
  }

  // ── Untested items — same format as detail section above ─────────────────────
  if (unknownItems.length > 0) {
    children.push(new Paragraph({ children: [new PageBreak()] }));
    children.push(h1('Untested — Specifications'));
    children.push(divider());

    for (const volt of voltOrder) {
      const voltItems = (nyrGrouped[volt] || []).slice().sort((a, b) => a.number - b.number);
      if (!voltItems.length) continue;

      children.push(h2(volt));

      for (const item of voltItems) {
        const { bg: sBg, fg: sFg } = statusColors(item.status);
        const bookmarkId = `issue_${item.number}`;

        children.push(new Table({
          width: { size: 9600, type: WidthType.DXA },
          columnWidths: [800, 5300, 1400, 2100],
          rows: [new TableRow({ cantSplit: true, children: [
            cell(item.voltage || 'Model Level', { bold: true, bg: C.darkBlue, fg: C.white, width: 800, align: AlignmentType.CENTER, size: 20 }),
            new TableCell({
              borders: allBorders, width: { size: 5300, type: WidthType.DXA },
              shading: { fill: C.lightBlue, type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              verticalAlign: VerticalAlign.CENTER,
              children: [new Paragraph({ children: [
                new Bookmark({ id: bookmarkId, children: [
                  new TextRun({ text: item.title, bold: true, size: 20, font: 'Arial', color: C.darkBlue })
                ]})
              ]})]
            }),
            cell(item.state || '',         { bg: C.lightBlue, fg: C.darkBlue, width: 1400, align: AlignmentType.CENTER, size: 16 }),
            cell(item.status || 'Unknown', { bg: sBg, fg: sFg, width: 2100, size: 16 }),
          ]})]
        }));

        if (item.body && item.body.trim()) {
          children.push(spacer());
          children.push(sectionLabel('Specification / Procedure', C.medBlue, 'EBF3FB'));
          children.push(...renderMarkdown(item.body, 200));
        }

        children.push(spacer());
      }
    }
  }

  // ── Assemble document ────────────────────────────────────────────────────────
  const doc = new Document({
    styles: {
      default: { document: { run: { font: 'Arial', size: 20 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 36, bold: true, font: 'Arial', color: C.darkBlue },
          paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 26, bold: true, font: 'Arial', color: C.medBlue },
          paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 1 } },
      ]
    },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
        }
      },
      footers: {
        default: new Footer({ children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: 'GTM965500P Test Report  |  GlobTek Engineering  |  Page ', size: 16, font: 'Arial', color: '888888' }),
            new TextRun({ children: [PageNumber.CURRENT], size: 16, font: 'Arial', color: '888888' }),
            new TextRun({ text: ' of ', size: 16, font: 'Arial', color: '888888' }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, font: 'Arial', color: '888888' }),
          ]
        })]})
      },
      children
    }]
  });

  return doc;
}

function imgToBase64(cachePath) {
  try {
    const data = fs.readFileSync(cachePath);
    const buf = Buffer.alloc(4);
    const fd = fs.openSync(cachePath, 'r');
    fs.readSync(fd, buf, 0, 4, 0);
    fs.closeSync(fd);
    const mime = (buf[0] === 0x89) ? 'image/png' : 'image/jpeg';
    return `data:${mime};base64,${data.toString('base64')}`;
  } catch { return null; }
}

// isSep: line consists entirely of |, -, :, spaces — no real data
function isSepHtml(line) {
  return /^[\|\s\-:]+$/.test(line.replace(/┬á/g, ' ').trim()) && line.includes('-') && line.includes('|');
}

function mdToHtml(text) {
  if (!text) return '';
  text = text
    .replace(/ΓåÆ/g, '→').replace(/┬▒/g, '±').replace(/┬╡/g, 'µ')
    .replace(/[\u00a0]/g, ' ');

  const lines = text.split(/\r?\n/);
  let html = '';
  let i = 0;
  const isRow = l => /^\s*\|/.test(l) || /^[^\|#\-\s>][^\n]*\|[^\n]*\|/.test(l.trim());

  while (i < lines.length) {
    const line = lines[i];

    // ### / #### = table title (centered, underlined)
    const hMatch = line.match(/^(#{1,4})\s+(.+)/);
    if (hMatch) {
      const level = hMatch[1].length;
      const htxt = hMatch[2].replace(/\*\*/g, '').trim();
      if (level >= 3) {
        html += `<p class="tbl-title">${htxt}</p>\n`;
      } else {
        html += `<p class="md-heading"><strong>${htxt}</strong></p>\n`;
      }
      i++; continue;
    }

    // Markdown table
    if (isRow(line)) {
      const tableLines = [];
      while (i < lines.length && (isRow(lines[i]) || isSepHtml(lines[i]))) {
        tableLines.push(lines[i++]);
      }
      const dataRows = tableLines.filter(l => !isSepHtml(l));
      if (dataRows.length > 0) {
        html += '<table class="md-table"><tbody>\n';
        dataRows.forEach((row, ri) => {
          const isNbsp = /^[\s]*[\u00a0┬á]/.test(row);
          let l = row.replace(/[\u00a0┬á]/g, ' ').trim();
          if (isNbsp) { l = l.replace(/^\|/, '').trim(); l = '| | ' + l; }
          if (!l.startsWith('|')) l = '|' + l;
          if (!l.endsWith('|')) l = l + '|';
          const parts = l.split('|'); parts.shift(); parts.pop();
          const cells = parts.map(c => c.trim());
          const tag = ri === 0 ? 'th' : 'td';
          const rowClass = ri === 0 ? '' : (ri % 2 === 0 ? ' class="row-even"' : ' class="row-odd"');
          html += `<tr${rowClass}>` + cells.map(c => {
            const imgM = c.match(/<img[^>]+src="([^"]+)"/i);
            if (imgM && ri > 0) {
              const filename = imgM[1].split('/').pop().split('?')[0];
              const b64 = imgToBase64(path.join(IMG_CACHE, filename));
              return b64 ? `<${tag}><img src="${b64}" style="max-width:300px;height:auto;display:block;margin:4px auto"></${tag}>` : `<${tag}>[image]</${tag}>`;
            }
            return `<${tag}>${c.replace(/<[^>]+>/g, '')}</${tag}>`;
          }).join('') + '</tr>\n';
        });
        html += '</tbody></table>\n';
      }
      continue;
    }

    // HR
    if (/^[-=]{3,}$/.test(line.trim())) { html += '<hr class="thin-rule">\n'; i++; continue; }

    // Blank
    if (!line.trim()) { i++; continue; }

    // Standalone image
    const imgM = line.match(/<img[^>]+src="([^"]+)"/i);
    if (imgM) {
      const filename = imgM[1].split('/').pop().split('?')[0];
      const b64 = imgToBase64(path.join(IMG_CACHE, filename));
      html += b64 ? `<img src="${b64}" style="max-width:100%;height:auto;display:block;margin:8px 0;border-radius:2px">\n` : '';
      i++; continue;
    }

    // Normal line
    const cleaned = line.replace(/<[^>]+>/g, '').replace(/\*\*/g, '').trim();
    if (cleaned) html += `<p>${cleaned}</p>\n`;
    i++;
  }
  return html;
}

function buildHtml(voltFilter = null) {
  const voltOrder = ['12V', '24V', '54V', 'Model Level'];
  const grouped = {};
  for (const item of items) {
    const volt = item.voltage || 'Model Level';
    if (!grouped[volt]) grouped[volt] = [];
    grouped[volt].push(item);
  }

  const activeVolts    = voltFilter ? [voltFilter] : voltOrder;
  const activeItems    = voltFilter ? (grouped[voltFilter] || []) : items;
  const allActiveItems = voltFilter ? allItems.filter(i => (i.voltage || 'Model Level') === voltFilter) : allItems;
  const counts = {};
  for (const item of allActiveItems) counts[item.status || 'Unknown'] = (counts[item.status || 'Unknown'] || 0) + 1;
  const total    = allActiveItems.length;
  const subtitle = voltFilter ? `${voltFilter} Model` : 'All Models';
  const catOrder = ['Input', 'Main Output', 'Standby Output', 'Fan Output', 'Protections',
                    'Environmental / Reliability', 'Safety', 'EMC', 'PFC'];
  const date = new Date().toLocaleDateString('en-US', {year:'numeric',month:'long',day:'numeric'});

  const badge = (s) => {
    const { bg, fg } = statusColors(s || 'Unknown');
    return `<span class="status-badge" style="background:#${bg};color:#${fg}">${s || 'Unknown'}</span>`;
  };

  const sectionLabel = (text, color, bg) =>
    `<div class="section-label" style="border-left:4px solid #${color};background:#${bg};color:#${color}">${text}</div>`;

  const unknownItems = allActiveItems.filter(i => !i.comments.length || !i.status || i.status === 'Unknown');
  const nyrGrouped = {};
  for (const item of unknownItems) {
    const v = item.voltage || 'Model Level';
    if (!nyrGrouped[v]) nyrGrouped[v] = [];
    nyrGrouped[v].push(item);
  }

  let nav = `<nav class="sidebar">
    <div class="nav-title">Navigation</div>
    <ul>
      <li><a href="#summary">Test Summary</a></li>
      <li><a href="#detail">Detailed Results</a></li>
      ${unknownItems.length > 0 ? '<li><a href="#untested-specs">Untested</a></li>' : ''}
    </ul>
    <div class="nav-section">Tests</div>
    ${activeVolts.map(v => {
      const vi = (grouped[v]||[]).slice().sort((a,b)=>a.number-b.number);
      if (!vi.length) return '';
      const byCategory = {};
      for (const item of vi) {
        const cat = item.category || 'Other';
        if (!byCategory[cat]) byCategory[cat] = [];
        byCategory[cat].push(item);
      }
      const orderedCats = [...catOrder.filter(c => byCategory[c]), ...Object.keys(byCategory).filter(c => !catOrder.includes(c))];
      return `<details class="nav-group" open>
        <summary class="nav-volt">${v}</summary>` +
        orderedCats.map(cat =>
          `<div class="nav-cat">${cat}</div><ul>${byCategory[cat].map(item =>
            `<li><a href="#issue_${item.number}">${item.title}</a></li>`).join('')}</ul>`
        ).join('') +
        `</details>`;
    }).join('')}
    ${unknownItems.length > 0 ? `
    <details class="nav-group">
      <summary class="nav-section nav-section-summary">Untested</summary>
      ${voltOrder.map(v => {
        const vi = (nyrGrouped[v]||[]).slice().sort((a,b)=>a.number-b.number);
        if (!vi.length) return '';
        const byCategory = {};
        for (const item of vi) {
          const cat = item.category || 'Other';
          if (!byCategory[cat]) byCategory[cat] = [];
          byCategory[cat].push(item);
        }
        const orderedCats = [...catOrder.filter(c => byCategory[c]), ...Object.keys(byCategory).filter(c => !catOrder.includes(c))];
        return `<details class="nav-group" open>
          <summary class="nav-volt">${v}</summary>` +
          orderedCats.map(cat =>
            `<div class="nav-cat">${cat}</div><ul>${byCategory[cat].map(item =>
              `<li><a href="#issue_${item.number}">${item.title}</a></li>`).join('')}</ul>`
          ).join('') +
          `</details>`;
      }).join('')}
    </details>` : ''}
  </nav>`;

  let body = '';

  // ── Cover ──────────────────────────────────────────────────────────────────
  body += `<div class="cover">
    <div class="cover-product">GTM965500P</div>
    <div class="cover-report">Test Report</div>
    <div class="cover-model">${subtitle}</div>
    <div class="cover-org">GlobTek Engineering</div>
    <div class="cover-date">Generated: ${date}</div>
  </div>
  <div class="h-divider"></div>`;

  // ── Summary ────────────────────────────────────────────────────────────────
  body += `<div class="section-card" id="summary">
  <div class="section-card-header">Test Summary</div>
  <div class="section-card-body">
  <p class="meta-line"><strong>Project: GTM965500P &nbsp;|&nbsp; ${subtitle} &nbsp;|&nbsp; Total Test Items: ${total}</strong></p>
  <table class="status-count-table">
    <thead><tr>
      <th>Status</th>
      <th>OK / Resolved</th><th>For Review</th><th>Regression Req'd</th>
      <th>In Progress</th><th>Has Issue</th><th>Invalid/Incomplete Test</th><th>Unknown</th>
    </tr></thead>
    <tbody><tr>
      <td><strong>Count</strong></td>
      <td class="cnt pass">${counts['OK / Resolved']||0}</td>
      <td class="cnt review">${counts['Marked for review']||0}</td>
      <td class="cnt regression">${counts["Regression Req'd"]||0}</td>
      <td class="cnt progress">${counts['In Progress']||0}</td>
      <td class="cnt issue">${counts['Has Issue']||0}</td>
      <td class="cnt invalid">${counts['Invalid/Incomplete Test']||0}</td>
      <td class="cnt unknown">${counts['Unknown']||0}</td>
    </tr></tbody>
  </table>`;

  // Per-voltage summary tables
  for (const volt of activeVolts) {
    const catItems = (grouped[volt]||[]).slice().sort((a,b)=>a.number-b.number);
    if (!catItems.length) continue;
    const byCategory = {};
    for (const item of catItems) {
      const cat = item.category || 'Other';
      if (!byCategory[cat]) byCategory[cat] = [];
      byCategory[cat].push(item);
    }
    const orderedCats = [...catOrder.filter(c => byCategory[c]), ...Object.keys(byCategory).filter(c => !catOrder.includes(c))];

    body += `<h2>${volt}</h2>
    <table class="summary-table">
      <thead><tr><th class="col-test">Test Item</th><th class="col-status">Status</th><th class="col-link">Link</th></tr></thead>
      <tbody>`;
    for (const cat of orderedCats) {
      const grp = byCategory[cat];
      if (!grp || !grp.length) continue;
      body += `<tr class="cat-row"><td colspan="3">${cat}</td></tr>`;
      grp.forEach((item, idx) => {
        const { bg, fg } = statusColors(item.status);
        const rowBg = idx % 2 === 0 ? '#ffffff' : '#EBF3FB';
        body += `<tr style="background:${rowBg}">
          <td>${item.title}</td>
          <td><span class="status-badge" style="background:#${bg};color:#${fg}">${item.status||'Unknown'}</span></td>
          <td class="link-cell"><a href="#issue_${item.number}">View →</a></td>
        </tr>`;
      });
    }
    body += `</tbody></table>`;
  }
  body += `</div></div>`; // section-card-body, section-card

  // ── Changes Required ───────────────────────────────────────────────────────
  const allChanges = [];
  for (const volt of activeVolts) {
    for (const item of (grouped[volt]||[])) {
      for (const comment of item.comments) {
        if (getCommentTag(comment.body) === 'CHANGE') allChanges.push({ item, comment });
      }
    }
  }
  if (allChanges.length > 0) {
    body += `<div class="section-card section-card-red">
    <div class="section-card-header section-card-header-red">Changes Required
      <span class="section-card-count">${allChanges.length} change${allChanges.length > 1 ? 's' : ''} flagged</span>
    </div>
    <div class="section-card-body">
    <table class="summary-table">
      <thead><tr><th style="width:80px">Voltage</th><th>Test</th><th>Change Description</th></tr></thead>
      <tbody>`;
    allChanges.forEach(({ item, comment }, idx) => {
      const desc = stripTag(comment.body).replace(/<[^>]+>/g,'').replace(/\*\*/g,'').replace(/\n+/g,' ').trim();
      const short = desc.length > 300 ? desc.substring(0,297)+'...' : desc;
      const rowBg = idx % 2 === 0 ? '#ffffff' : '#FEF4F1';
      body += `<tr style="background:${rowBg}">
        <td style="text-align:center"><span class="status-badge" style="background:#FCE4D6;color:#843C0C">${item.voltage||'Model Level'}</span></td>
        <td><a href="#issue_${item.number}">${item.title}</a><br><span class="byline">${(comment.author||{}).login||''} · ${(comment.createdAt||'').substring(0,10)}</span></td>
        <td>${short}</td>
      </tr>`;
    });
    body += `</tbody></table>
    </div></div>`; // section-card-body, section-card
  }

  // ── Tests Not Yet Run ──────────────────────────────────────────────────────
  if (unknownItems.length > 0) {
    body += `<div class="section-card section-card-gray">
    <div class="section-card-header section-card-header-gray">Tests Not Yet Run
      <span class="section-card-count">${unknownItems.length} test${unknownItems.length > 1 ? 's' : ''} pending</span>
    </div>
    <div class="section-card-body">`;
    for (const volt of voltOrder) {
      const voltItems = (nyrGrouped[volt]||[]).slice().sort((a,b)=>a.number-b.number);
      if (!voltItems.length) continue;
      const byCategory = {};
      for (const item of voltItems) {
        const cat = item.category || 'Other';
        if (!byCategory[cat]) byCategory[cat] = [];
        byCategory[cat].push(item);
      }
      const orderedCats = [...catOrder.filter(c => byCategory[c]), ...Object.keys(byCategory).filter(c => !catOrder.includes(c))];
      body += `<h2>${volt}</h2>
      <table class="summary-table">
        <thead><tr><th class="col-test">Test Item</th><th class="col-status">Status</th><th class="col-link">Link</th></tr></thead>
        <tbody>`;
      for (const cat of orderedCats) {
        const grp = byCategory[cat];
        if (!grp || !grp.length) continue;
        body += `<tr class="cat-row"><td colspan="3">${cat}</td></tr>`;
        grp.forEach((item, idx) => {
          const rowBg = idx % 2 === 0 ? '#ffffff' : '#EBF3FB';
          body += `<tr style="background:${rowBg}"><td>${item.title}</td><td>${badge(item.status)}</td><td class="link-cell"><a href="#issue_${item.number}">View →</a></td></tr>`;
        });
      }
      body += `</tbody></table>`;
    }
    body += `</div></div>`; // section-card-body, section-card
  }

  // ── Detailed Results ───────────────────────────────────────────────────────
  body += `<div class="section-card">
  <div class="section-card-header" id="detail">Detailed Test Results</div>
  <div class="section-card-body">`;

  for (const volt of activeVolts) {
    const catItems = (grouped[volt]||[]).slice().sort((a,b)=>a.number-b.number);
    if (!catItems.length) continue;
    body += `<h2>${volt}</h2>`;

    for (const item of catItems) {
      const { bg: sBg, fg: sFg } = statusColors(item.status);

      // Header bar: matches the docx 4-column header table
      body += `<div class="test-card" id="issue_${item.number}">
        <table class="test-header-table">
          <tr>
            <td class="th-volt">${item.voltage||'Model Level'}</td>
            <td class="th-title">${item.title}</td>
            <td class="th-state">${item.state||''}</td>
            <td class="th-status" style="background:#${sBg};color:#${sFg}">${item.status||'Unknown'}</td>
          </tr>
        </table>`;

      // Spec / Procedure
      if (item.body && item.body.trim()) {
        body += sectionLabel('Specification / Procedure', '2E75B6', 'EBF3FB');
        body += `<div class="comment-thread" style="background:#EBF3FB;border-left:none;padding:12px 16px">${mdToHtml(item.body)}</div>`;
      }

      // Comment block renderer — mirrors docx exactly
      const renderBlock = (comments, label, labelColor, labelBg, threadBg, defaultOpen = true) => {
        if (!comments.length) return '';
        const openAttr = defaultOpen ? ' open' : '';
        let h = `<details class="collapsible-block"${openAttr}>`;
        h += `<summary class="section-label" style="border-left:4px solid #${labelColor};background:#${labelBg};color:#${labelColor}">${label}</summary>`;
        h += `<div class="thread-block" style="border-left:3px solid #${labelColor};background:#${threadBg}">`;
        comments.forEach((comment, idx) => {
          const isLast = idx === comments.length - 1;
          const author = (comment.author||{}).login || 'unknown';
          const date   = (comment.createdAt||'').substring(0,10);
          if (idx > 0) h += `<div class="thread-sep"></div>`;
          h += `<div class="thread-comment">
            <div class="byline ${isLast ? 'byline-latest' : ''}">${author} · ${date}${isLast ? ' <strong>[LATEST]</strong>' : ''}</div>
            <div class="comment-body">${mdToHtml(stripTag(comment.body))}</div>
          </div>`;
        });
        h += `</div></details>`;
        return h;
      };

      const changeComments = item.comments.filter(c => getCommentTag(c.body) === 'CHANGE');
      const resultComments = item.comments.filter(c => getCommentTag(c.body) === 'RESULT');
      const otherComments  = item.comments.filter(c => !getCommentTag(c.body));

      body += renderBlock(changeComments, 'Changes Required', '843C0C', 'FCE4D6', 'FEF4F1', true);
      body += renderBlock(resultComments, 'Results',          '375623', 'E2EFDA', 'F4FAF0', true);
      if (otherComments.length) {
        body += renderBlock(otherComments, `Comments (${otherComments.length})`, '2E75B6', 'D6E4F0', 'F2F6FB', false);
      }

      body += `</div>`; // .test-card
    }
  }

  body += `</div></div>`; // section-card-body, section-card (Detailed Results)

  // ── Untested items — same format as detail section above (HTML) ───────────────
  if (unknownItems.length > 0) {
    body += `<div class="section-card">
    <div class="section-card-header" id="untested-specs">Untested — Specifications</div>
    <div class="section-card-body">`;
    for (const volt of voltOrder) {
      const voltItems = (nyrGrouped[volt]||[]).slice().sort((a,b)=>a.number-b.number);
      if (!voltItems.length) continue;
      body += `<h2>${volt}</h2>`;
      for (const item of voltItems) {
        const { bg: sBg, fg: sFg } = statusColors(item.status);
        body += `<div class="test-card" id="issue_${item.number}">
          <table class="test-header-table"><tr>
            <td class="th-volt">${item.voltage||'Model Level'}</td>
            <td class="th-title">${item.title}</td>
            <td class="th-state">${item.state||''}</td>
            <td class="th-status" style="background:#${sBg};color:#${sFg}">${item.status||'Unknown'}</td>
          </tr></table>`;
        if (item.body && item.body.trim()) {
          body += `<div class="section-label" style="border-left:4px solid #2E75B6;background:#EBF3FB;color:#2E75B6">Specification / Procedure</div>`;
          body += `<div class="thread-block" style="background:#EBF3FB;border-left:none;padding:12px 16px">${mdToHtml(item.body)}</div>`;
        }
        body += `</div>`;
      }
    }
    body += `</div></div>`;
  }

  // ── CSS — mirrors the PDF exactly ─────────────────────────────────────────
  const css = `
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: Arial, sans-serif; font-size: 13px; color: #222; background: #f0f2f5; }
    a { color: #2E75B6; text-decoration: none; }
    a:hover { text-decoration: underline; }

    /* Layout */
    .page { display: flex; max-width: 1200px; margin: 0 auto; padding: 24px 16px; gap: 20px; align-items: flex-start; }
    .main  { flex: 1; min-width: 0; }

    /* Section cards */
    .section-card { background: #fff; border: 1px solid #d0d4db; border-radius: 5px; margin-bottom: 20px; overflow: hidden; }
    .section-card-header { background: #1F3864; color: #fff; font-size: 16px; font-weight: 700;
      padding: 10px 16px; display: flex; align-items: center; justify-content: space-between; }
    .section-card-header-red  { background: #843C0C; }
    .section-card-header-gray { background: #555; }
    .section-card-count { font-size: 11px; font-weight: 400; opacity: 0.85; }
    .section-card-body { padding: 16px; }

    /* Sidebar */
    .sidebar { width: 240px; flex-shrink: 0; position: sticky; top: 16px; background: #fff;
      border: 1px solid #ddd; border-radius: 4px; padding: 12px; font-size: 13px; max-height: 92vh; overflow-y: auto; }
    .nav-title { font-size: 11px; text-transform: uppercase; letter-spacing:.06em; color:#999; margin-bottom:6px; font-weight:600; }
    .nav-section { margin-top:10px; font-weight:700; font-size:13px; color:#1F3864; }
    .nav-section-summary { cursor:pointer; list-style:none; display:flex; align-items:center; gap:4px; }
    .nav-section-summary::-webkit-details-marker { display:none; }
    .nav-section-summary::before { content:'▶'; font-size:9px; transition:transform .15s; color:#1F3864; }
    details.nav-group[open] > summary.nav-section-summary::before { transform:rotate(90deg); }
    .nav-group { border:none; }
    .nav-volt { margin-top:6px; font-weight:700; font-size:13px; color:#2E75B6; cursor:pointer;
      list-style:none; display:flex; align-items:center; gap:4px; user-select:none; }
    .nav-volt::-webkit-details-marker { display:none; }
    .nav-volt::before { content:'▶'; font-size:9px; transition:transform .15s; color:#2E75B6; }
    details.nav-group[open] > summary.nav-volt::before { transform:rotate(90deg); }
    .nav-cat { margin-top:4px; font-size:12px; font-weight:600; color:#2E75B6; padding-left:6px; border-left:2px solid #D6E4F0; }
    .sidebar ul { list-style:none; padding-left:4px; }
    .sidebar li { margin:4px 0; }
    .sidebar a { color:#333; font-size:13px; }
    .sidebar a:hover { color:#2E75B6; }

    /* Cover */
    .cover { text-align:center; padding:40px 0 28px; background:#fff; border-radius:4px; margin-bottom:16px; border:1px solid #ddd; }
    .cover-product { font-size:48px; font-weight:800; color:#1F3864; letter-spacing:-1px; }
    .cover-report  { font-size:28px; color:#2E75B6; margin-top:4px; }
    .cover-model   { font-size:20px; color:#2E75B6; margin-top:2px; }
    .cover-org     { font-size:16px; color:#555; margin-top:6px; }
    .cover-date    { font-size:13px; color:#888; margin-top:4px; }

    /* Content blocks */
    h1 { font-size:18px; color:#1F3864; margin: 16px 0 6px; }
    h2 { font-size:15px; color:#2E75B6; margin: 16px 0 6px; font-weight:700; border-bottom:2px solid #D6E4F0; padding-bottom:4px; }
    .h-divider { border:none; border-bottom:3px solid #2E75B6; margin:8px 0 12px; }
    .meta-line { font-size:12px; margin:4px 0 10px; color:#333; }
    .thin-rule { border:none; border-top:1px solid #ccc; margin:8px 0; }
    p { margin: 3px 0 5px; line-height:1.5; }

    /* Status count table */
    .status-count-table { border-collapse:collapse; width:100%; margin-bottom:14px; font-size:12px; }
    .status-count-table th { background:#1F3864; color:#fff; padding:7px 10px; text-align:center; border:1px solid #bbb; }
    .status-count-table td { padding:7px 10px; border:1px solid #bbb; text-align:center; vertical-align:middle; }
    .status-count-table td:first-child { text-align:left; }
    .cnt { font-weight:700; font-size:15px; text-align:center; vertical-align:middle; }
    .cnt.pass       { background:#E2EFDA; color:#375623; }
    .cnt.review     { background:#FFF2CC; color:#7F6000; }
    .cnt.regression { background:#FCE4D6; color:#843C0C; }
    .cnt.progress   { background:#DEEAF1; color:#1F3864; }
    .cnt.issue      { background:#FCE4D6; color:#843C0C; }
    .cnt.invalid    { background:#EDE7F6; color:#4527A0; }
    .cnt.unknown    { background:#F2F2F2; color:#444444; }

    /* Summary / not-yet-run tables */
    .summary-table { border-collapse:collapse; width:100%; margin-bottom:14px; font-size:12px; }
    .summary-table th { background:#1F3864; color:#fff; padding:7px 10px; text-align:left; border:1px solid #bbb; }
    .summary-table td { padding:7px 10px; border:1px solid #bbb; vertical-align:middle; }
    .summary-table .col-test   { width:62%; }
    .summary-table .col-status { width:23%; }
    .summary-table .col-link   { width:15%; }
    .summary-table .link-cell  { text-align:center; }
    .cat-row td { background:#2E75B6 !important; color:#fff; font-weight:600; font-size:11px; padding:4px 10px; }

    /* Status badge */
    .status-badge { display:inline-block; padding:2px 7px; border-radius:2px; font-size:11px; font-weight:600; white-space:nowrap; }

    /* Test card */
    .test-card { border:1px solid #ccc; border-radius:3px; margin-bottom:20px; overflow:hidden; background:#fff; }
    .test-header-table { border-collapse:collapse; width:100%; }
    .test-header-table td { padding:8px 12px; border:none; vertical-align:middle; }
    .th-volt   { background:#1F3864; color:#fff; font-weight:700; font-size:13px; white-space:nowrap; width:80px; text-align:center; }
    .th-title  { background:#D6E4F0; color:#1F3864; font-weight:700; font-size:13px; }
    .th-state  { background:#D6E4F0; color:#555; font-size:11px; white-space:nowrap; text-align:center; width:80px; }
    .th-status { font-weight:600; font-size:12px; white-space:nowrap; text-align:center; width:140px; }

    /* Section labels */
    .section-label { padding:6px 12px 6px 14px; font-weight:700; font-size:12px; margin-top:1px; }
    .collapsible-block { border:none; }
    .collapsible-block > summary {
      list-style:none; cursor:pointer; display:flex; align-items:center; gap:6px;
      padding:6px 12px 6px 14px; font-weight:700; font-size:12px; margin-top:1px;
      user-select:none;
    }
    .collapsible-block > summary::-webkit-details-marker { display:none; }
    .collapsible-block > summary::before { content:'▶'; font-size:9px; transition:transform .15s; flex-shrink:0; }
    .collapsible-block[open] > summary::before { transform:rotate(90deg); }

    /* Thread blocks */
    .thread-block { padding:0; }
    .thread-comment { padding:10px 16px; }
    .thread-sep { border-top:1px solid rgba(0,0,0,.1); margin:0 16px; }
    .byline { font-size:11px; color:#666; font-style:italic; margin-bottom:5px; }
    .byline-latest { color:#1F3864; font-style:normal; font-weight:600; }
    .comment-body p { margin:3px 0 5px; }
    .comment-body p:last-child { margin-bottom:0; }

    /* Inline markdown tables */
    .md-table { border-collapse:collapse; width:100%; margin:8px 0 10px; font-size:12px; }
    .md-table th { background:#D6E4F0; color:#1F3864; padding:6px 10px; text-align:left; border:1px solid #bbb; font-weight:700; }
    .md-table td { padding:6px 10px; border:1px solid #ccc; }
    .md-table tr.row-odd  td { background:#F2F2F2; }
    .md-table tr.row-even td { background:#ffffff; }

    /* Table title / heading */
    .tbl-title { text-align:center; font-weight:700; color:#1F3864; border-bottom:2px solid #2E75B6;
      padding-bottom:3px; margin:14px 0 4px; font-size:13px; }
    .md-heading { color:#1F3864; font-size:13px; margin:8px 0 3px; text-decoration:underline; }

    /* Byline in changes table */
    .byline { font-size:11px; color:#888; }

    img { max-width:100%; height:auto; border-radius:2px; display:block; margin:6px 0; }
  `;

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>GTM965500P Test Report — ${subtitle}</title>
<style>${css}</style>
</head>
<body>
<div class="page">
  ${nav}
  <div class="main">${body}</div>
</div>
</body>
</html>`;
}

function writeHtml(voltFilter, outPath) {
  const html = buildHtml(voltFilter);
  fs.writeFileSync(outPath, html, 'utf-8');
  console.log(`  HTML:    ${outPath}`);
}

async function main() {
  console.log(`\nMode: ${MODE}`);
  await prefetchAllImages();

  const voltOrder = ['12V', '24V', '54V', 'Model Level'];

  console.log('\nBuilding full report...');
  if (MODE !== 'html') {
    const fullDoc = await buildDoc(null);
    await writeDoc(fullDoc, OUTPUT_FILE);
  }
  if (MODE !== 'pdf') {
    writeHtml(null, path.join(HTML_DIR, 'GTM965500P_Test_Report.html'));
  }

  for (const volt of voltOrder) {
    const voltItems = items.filter(i => (i.voltage || 'Model Level') === volt);
    if (!voltItems.length) continue;
    const safeName = volt.replace(/\s+/g, '_');
    console.log(`\nBuilding ${volt} report...`);
    if (MODE !== 'html') {
      const outPath = path.join(DOCX_DIR, `GTM965500P_Test_Report_${safeName}.docx`);
      const doc = await buildDoc(volt);
      await writeDoc(doc, outPath);
    }
    if (MODE !== 'pdf') {
      writeHtml(volt, path.join(HTML_DIR, `GTM965500P_Test_Report_${safeName}.html`));
    }
  }

  console.log('\nAll reports done!');
}

main().catch(e => { console.error(e); process.exit(1); });