const fs = require('fs/promises');
const path = require('path');
const { spawn } = require('child_process');
const { parse } = require('csv-parse/sync');
const fontkit = require('@pdf-lib/fontkit');
const {
  PDFDocument,
  rgb,
} = require('pdf-lib');

const workspaceRoot = path.resolve(__dirname, '..');
const templatePdfPath = path.join(workspaceRoot, 'resources', 'cards-template.pdf');
const factsCsvPath = path.join(workspaceRoot, 'resources', "That's Interesting But - Facts - Sheet1.csv");
const hostedFactsCsvUrl = 'https://docs.google.com/spreadsheets/d/105bSNBbS60Sqg0BMVTrDRfX0kc9AZZBgS3wF-ecXke4/export?format=csv&gid=0';
const defaultFontPath = path.join(
  workspaceRoot,
  'resources',
  'fonts',
  'GuardianTextEgyptian-Regular-Web.woff',
);
const outputDir = path.join(workspaceRoot, 'output');

const textArea = {
  x: 40.5,
  y: 37,
  width: 220,
  height: 90,
};

const drawStyle = {
  maxFontSize: 13,
  minFontSize: 8,
  lineGapMultiplier: 1.23,
};

const terminalPunctuationPattern = /[.!?…][)\]}'"”’»]*$/;
const trailingClosersPattern = /[)\]}'"”’»]+$/;
const trailingWeakPunctuationPattern = /[,;:–—-]+$/;
const requiredIncludedFactCount = 75;

function tokenize(text) {
  return text
    .replace(/\s+/g, ' ')
    .trim()
    .split(' ')
    .filter(Boolean);
}

function breakLongWord(word, maxWidth, font, fontSize) {
  if (font.widthOfTextAtSize(word, fontSize) <= maxWidth) {
    return [word];
  }

  const parts = [];
  let start = 0;

  while (start < word.length) {
    let end = start + 1;
    while (
      end <= word.length &&
      font.widthOfTextAtSize(word.slice(start, end), fontSize) <= maxWidth
    ) {
      end += 1;
    }

    if (end === start + 1) {
      end = Math.min(start + 2, word.length);
    } else {
      end -= 1;
    }

    parts.push(word.slice(start, end));
    start = end;
  }

  return parts;
}

function wrapText(text, maxWidth, font, fontSize) {
  const words = tokenize(text).flatMap((word) => breakLongWord(word, maxWidth, font, fontSize));
  const lines = [];
  let current = '';

  for (const word of words) {
    const candidate = current ? `${current} ${word}` : word;
    if (font.widthOfTextAtSize(candidate, fontSize) <= maxWidth) {
      current = candidate;
      continue;
    }

    if (current) {
      lines.push(current);
    }
    current = word;
  }

  if (current) {
    lines.push(current);
  }

  return lines;
}

function fitText(text, font, maxWidth, maxHeight) {
  for (let fontSize = drawStyle.maxFontSize; fontSize >= drawStyle.minFontSize; fontSize -= 0.5) {
    const lines = wrapText(text, maxWidth, font, fontSize);
    const lineHeight = fontSize * drawStyle.lineGapMultiplier;
    const totalHeight = lines.length * lineHeight;

    if (totalHeight <= maxHeight) {
      return { lines, fontSize, lineHeight };
    }
  }

  const fontSize = drawStyle.minFontSize;
  const lineHeight = fontSize * drawStyle.lineGapMultiplier;
  const maxLines = Math.max(1, Math.floor(maxHeight / lineHeight));
  const lines = wrapText(text, maxWidth, font, fontSize).slice(0, maxLines);

  if (lines.length > 0) {
    let finalLine = lines[lines.length - 1];
    while (font.widthOfTextAtSize(`${finalLine}…`, fontSize) > maxWidth && finalLine.length > 1) {
      finalLine = finalLine.slice(0, -1).trimEnd();
    }
    lines[lines.length - 1] = `${finalLine}…`;
  }

  return { lines, fontSize, lineHeight };
}

function timestamp() {
  const now = new Date();
  const parts = [
    now.getFullYear(),
    String(now.getMonth() + 1).padStart(2, '0'),
    String(now.getDate()).padStart(2, '0'),
    String(now.getHours()).padStart(2, '0'),
    String(now.getMinutes()).padStart(2, '0'),
    String(now.getSeconds()).padStart(2, '0'),
  ];
  return `${parts[0]}${parts[1]}${parts[2]}-${parts[3]}${parts[4]}${parts[5]}`;
}

function ensureTerminalPunctuation(text) {
  const trimmed = text.trim();
  if (!trimmed) {
    return { text: trimmed, changed: false };
  }

  if (terminalPunctuationPattern.test(trimmed)) {
    return { text: trimmed, changed: false };
  }

  let closers = '';
  let base = trimmed;
  const trailingClosers = trimmed.match(trailingClosersPattern);
  if (trailingClosers) {
    closers = trailingClosers[0];
    base = trimmed.slice(0, -closers.length).trimEnd();
  }

  const baseWithoutWeakPunctuation = base.replace(trailingWeakPunctuationPattern, '').trimEnd();
  if (!baseWithoutWeakPunctuation) {
    return { text: `${base}.${closers}`, changed: false };
  }

  return { text: `${baseWithoutWeakPunctuation}.${closers}`, changed: true };
}

async function writePunctuationLog(changes) {
  const logPath = path.join(outputDir, `punctuation-fixes-${timestamp()}.log`);
  const lines = [
    `Punctuation fixes applied: ${changes.length}`,
    '',
    ...changes.map((change) => (
      `Row ${change.csvRow}: ${change.before} => ${change.after}`
    )),
    '',
  ];

  await fs.writeFile(logPath, lines.join('\n'), 'utf8');
  return logPath;
}

function isIncludeTrue(value) {
  const normalized = String(value ?? '').trim().toLowerCase();
  return normalized === 'true' || normalized === '1' || normalized === 'yes' || normalized === 'y';
}

function isOptionEnabled(value) {
  const normalized = String(value ?? '').trim().toLowerCase();
  return normalized === 'true' || normalized === '1' || normalized === 'yes' || normalized === 'y';
}

function openFileWithDefaultApp(filePath) {
  if (process.platform === 'darwin') {
    const child = spawn('open', [filePath], { detached: true, stdio: 'ignore' });
    child.unref();
    return;
  }

  if (process.platform === 'win32') {
    const child = spawn('cmd', ['/c', 'start', '', filePath], {
      detached: true,
      stdio: 'ignore',
    });
    child.unref();
    return;
  }

  if (process.platform === 'linux') {
    const child = spawn('xdg-open', [filePath], { detached: true, stdio: 'ignore' });
    child.unref();
    return;
  }

  throw new Error(`Auto-open is not supported on platform: ${process.platform}`);
}

function parseCliOptions() {
  const args = process.argv.slice(2);
  const options = {};

  for (const arg of args) {
    const normalizedArg = arg.startsWith('--') ? arg.slice(2) : arg;
    const separatorIndex = normalizedArg.indexOf('=');
    if (separatorIndex === -1) {
      const key = normalizedArg.trim();
      if (key) {
        options[key] = 'true';
      }
      continue;
    }

    const key = normalizedArg.slice(0, separatorIndex).trim();
    const value = normalizedArg.slice(separatorIndex + 1).trim();
    if (!key) {
      continue;
    }

    options[key] = value;
  }

  return options;
}

async function loadCsvRaw(options) {
  const hosted = String(options.hosted || '').trim().toLowerCase() === 'true';
  if (!hosted) {
    return { csvRaw: await fs.readFile(factsCsvPath, 'utf8'), source: factsCsvPath };
  }

  const response = await fetch(hostedFactsCsvUrl);
  if (!response.ok) {
    throw new Error(`Failed to fetch hosted CSV (${response.status} ${response.statusText}).`);
  }

  return { csvRaw: await response.text(), source: hostedFactsCsvUrl };
}

function resolveFontPath(options) {
  if (!options.font) {
    return defaultFontPath;
  }

  const providedPath = String(options.font).trim();
  if (!providedPath) {
    return defaultFontPath;
  }

  return path.isAbsolute(providedPath)
    ? providedPath
    : path.resolve(workspaceRoot, providedPath);
}

async function loadFacts(options) {
  const { csvRaw, source } = await loadCsvRaw(options);
  const rows = parse(csvRaw, {
    columns: true,
    skip_empty_lines: true,
    relax_column_count: true,
    trim: true,
  });

  if (rows.length === 0) {
    throw new Error('CSV file has no fact rows.');
  }

  const availableColumns = Object.keys(rows[0]);
  const includeColumnName = availableColumns.find(
    (columnName) => columnName.trim().toLowerCase() === 'include',
  );
  if (!includeColumnName) {
    throw new Error(
      `CSV is missing required "include" column. Found columns: ${availableColumns.join(', ')}`,
    );
  }

  const punctuationChanges = [];
  const facts = rows
    .map((row, index) => {
      if (!isIncludeTrue(row[includeColumnName])) {
        return '';
      }

      const normalizedFact = String(row.Fact || '').replace(/\s+/g, ' ').trim();
      if (!normalizedFact) {
        return '';
      }

      const fixedFact = ensureTerminalPunctuation(normalizedFact);
      if (fixedFact.changed) {
        punctuationChanges.push({
          csvRow: index + 2,
          before: normalizedFact,
          after: fixedFact.text,
        });
      }

      return fixedFact.text;
    })
    .filter(Boolean);

  if (facts.length !== requiredIncludedFactCount) {
    throw new Error(
      `Expected exactly ${requiredIncludedFactCount} included facts but found ${facts.length}.`,
    );
  }

  return { facts, punctuationChanges, source };
}

async function generate() {
  const options = parseCliOptions();
  const autoOpen = isOptionEnabled(options.autoopen);
  const selectedFontPath = resolveFontPath(options);
  const { facts, punctuationChanges, source } = await loadFacts(options);
  const templateBytes = await fs.readFile(templatePdfPath);
  const templateDoc = await PDFDocument.load(templateBytes);

  const pageCount = templateDoc.getPageCount();
  if (pageCount < 2) {
    throw new Error('Template PDF must contain at least one front page and one back page.');
  }

  const frontTemplateIndex = 0;
  const backTemplateIndex = pageCount - 1;

  const outputDoc = await PDFDocument.create();
  outputDoc.registerFontkit(fontkit);
  const fontBytes = await fs.readFile(selectedFontPath);
  const measurementFont = await outputDoc.embedFont(fontBytes);

  for (const fact of facts) {
    const [page] = await outputDoc.copyPages(templateDoc, [frontTemplateIndex]);

    page.drawRectangle({
      x: textArea.x,
      y: textArea.y,
      width: textArea.width,
      height: textArea.height,
      color: rgb(1, 1, 1),
      borderWidth: 0,
    });

    const fit = fitText(fact, measurementFont, textArea.width, textArea.height);
    const firstLineY = textArea.y + (fit.lines.length - 1) * fit.lineHeight;

    fit.lines.forEach((line, index) => {
      page.drawText(line, {
        x: textArea.x,
        y: firstLineY - index * fit.lineHeight,
        size: fit.fontSize,
        font: measurementFont,
        color: rgb(0, 0, 0),
      });
    });

    outputDoc.addPage(page);
  }

  const [backPage] = await outputDoc.copyPages(templateDoc, [backTemplateIndex]);
  outputDoc.addPage(backPage);

  await fs.mkdir(outputDir, { recursive: true });
  const outputPath = path.join(outputDir, `cards-generated-${timestamp()}.pdf`);
  const outputBytes = await outputDoc.save();
  await fs.writeFile(outputPath, outputBytes);

  let punctuationLogPath = null;
  if (punctuationChanges.length > 0) {
    punctuationLogPath = await writePunctuationLog(punctuationChanges);
  }

  console.log(`Created ${outputPath}`);
  console.log(`Facts: ${facts.length}`);
  console.log(`Punctuation fixes: ${punctuationChanges.length}`);
  if (punctuationLogPath) {
    console.log(`Punctuation log: ${punctuationLogPath}`);
  }
  console.log(`Font: ${selectedFontPath}`);
  console.log(`CSV source: ${source}`);
  console.log('Final page count (facts + back):', facts.length + 1);

  if (autoOpen) {
    try {
      openFileWithDefaultApp(outputPath);
      console.log(`Opened PDF: ${outputPath}`);
    } catch (error) {
      console.warn(`Could not auto-open PDF: ${error.message || error}`);
    }
  }
}

generate().catch((error) => {
  console.error(error.message || error);
  process.exit(1);
});
