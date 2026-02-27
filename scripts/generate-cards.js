const fs = require('fs/promises');
const path = require('path');
const { parse } = require('csv-parse/sync');
const {
  PDFDocument,
  StandardFonts,
  PDFName,
  PDFDict,
  PDFHexString,
  drawText,
  degrees,
  rgb,
} = require('pdf-lib');

const workspaceRoot = path.resolve(__dirname, '..');
const templatePdfPath = path.join(workspaceRoot, 'resources', 'cards-template.pdf');
const factsCsvPath = path.join(workspaceRoot, 'resources', "That's Interesting But - Facts - Sheet1.csv");
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

async function loadFacts() {
  const csvRaw = await fs.readFile(factsCsvPath, 'utf8');
  const rows = parse(csvRaw, {
    columns: true,
    skip_empty_lines: true,
    relax_column_count: true,
    trim: true,
  });

  const facts = rows
    .map((row) => String(row.Fact || '').replace(/\s+/g, ' ').trim())
    .filter(Boolean);

  if (facts.length === 0) {
    throw new Error('No facts found in CSV file.');
  }

  return facts;
}

async function generate() {
  const facts = await loadFacts();
  const templateBytes = await fs.readFile(templatePdfPath);
  const templateDoc = await PDFDocument.load(templateBytes);

  const pageCount = templateDoc.getPageCount();
  if (pageCount < 2) {
    throw new Error('Template PDF must contain at least one front page and one back page.');
  }

  const frontTemplateIndex = 0;
  const backTemplateIndex = pageCount - 1;

  const frontResources = templateDoc.getPages()[frontTemplateIndex].node.Resources();
  const frontFontDict = frontResources.lookup(PDFName.of('Font'), PDFDict);
  const fontKeys = frontFontDict.keys().map((key) => key.decodeText());
  if (fontKeys.length === 0) {
    throw new Error('Could not locate any font resources on the front template page.');
  }
  const factFontResourceName = PDFName.of(
    fontKeys.includes('T1_0') ? 'T1_0' : fontKeys[0],
  );

  const outputDoc = await PDFDocument.create();
  const measurementFont = await outputDoc.embedFont(StandardFonts.HelveticaBold);

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
      page.pushOperators(
        ...drawText(PDFHexString.fromText(line), {
          x: textArea.x,
          y: firstLineY - index * fit.lineHeight,
          size: fit.fontSize,
          font: factFontResourceName,
          color: rgb(0, 0, 0),
          rotate: degrees(0),
          xSkew: degrees(0),
          ySkew: degrees(0),
        }),
      );
    });

    outputDoc.addPage(page);
  }

  const [backPage] = await outputDoc.copyPages(templateDoc, [backTemplateIndex]);
  outputDoc.addPage(backPage);

  await fs.mkdir(outputDir, { recursive: true });
  const outputPath = path.join(outputDir, `cards-generated-${timestamp()}.pdf`);
  const outputBytes = await outputDoc.save();
  await fs.writeFile(outputPath, outputBytes);

  console.log(`Created ${outputPath}`);
  console.log(`Facts: ${facts.length}`);
  console.log('Final page count (facts + back):', facts.length + 1);
}

generate().catch((error) => {
  console.error(error.message || error);
  process.exit(1);
});
