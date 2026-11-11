// ConvertFlow - Headless Local File Conversion Engine
import fs from 'fs';
import path from 'path';
import os from 'os';
import { execFile } from 'child_process';
import { fileURLToPath, pathToFileURL } from 'url';
import JSZip from 'jszip';
import { jsPDF } from 'jspdf';
import mammoth from 'mammoth';
import pptxgen from 'pptxgenjs';
import sharp from 'sharp';
import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf.mjs';

import { convertMedia } from '../server/converters/mediaConverter.js';
import { convertImage } from '../server/converters/imageConverter.js';
import { parseToNormalizedData, renderFromNormalizedData } from '../server/converters/dataConverter.js';
import { convertDocument } from '../server/converters/documentConverter.js';
import { convertServerPptx, parseServerPresentation } from '../server/converters/pptxConverter.js';
import { canHandleWithOffice, convertWithOffice } from '../server/converters/officeConverter.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT_DIR = path.resolve(__dirname, '..');
const PDF_RENDER_SCRIPT = path.resolve(ROOT_DIR, 'server', 'scripts', 'pdfRender.ps1');
const PYTHON_PDF_ENGINE = path.resolve(ROOT_DIR, 'server', 'scripts', 'pdfEngine.py');

// Configure pdfjs worker safely for Node.js
try {
  const workerPath = path.resolve(ROOT_DIR, 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.worker.mjs');
  if (fs.existsSync(workerPath)) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = pathToFileURL(workerPath).href;
  }
} catch {
  // Ignore worker config fallback
}

// Suppress noisy pdfjs font/standard warnings in CLI output
const PDFJS_VERBOSITY = pdfjsLib.VerbosityLevel?.ERRORS ?? 0;

/**
 * Generates a collision-free output file path in the target directory.
 */
export function getUniqueOutputPath(dir, baseName, ext) {
  const cleanExt = ext.replace(/^\./, '');
  let candidate = path.join(dir, `${baseName}.${cleanExt}`);
  let counter = 1;
  while (fs.existsSync(candidate)) {
    candidate = path.join(dir, `${baseName} (${counter}).${cleanExt}`);
    counter++;
  }
  return candidate;
}

/**
 * Generates a collision-free output folder path in the target directory.
 */
export function getUniqueFolderPath(dir, folderName) {
  let candidate = path.join(dir, folderName);
  let counter = 1;
  while (fs.existsSync(candidate)) {
    candidate = path.join(dir, `${folderName} (${counter})`);
    counter++;
  }
  return candidate;
}

/**
 * Creates a temporary path in OS temp directory.
 */
function createTempPath(ext) {
  const cleanExt = ext ? `.${ext.replace(/^\./, '')}` : '';
  return path.join(os.tmpdir(), `cf_${Date.now()}_${Math.random().toString(36).slice(2, 9)}${cleanExt}`);
}

function removeQuiet(targetPath) {
  if (!targetPath) return;
  try {
    if (fs.existsSync(targetPath)) {
      fs.rmSync(targetPath, { recursive: true, force: true });
    }
  } catch {
    // Ignore cleanup errors
  }
}

/**
 * Registers a Unicode TTF font (supporting Turkish ğüşıöçĞÜŞİÖÇ) on a jsPDF instance across Windows, macOS, and Linux.
 */
function applyUnicodeFont(doc, style = 'normal') {
  const candidates = process.platform === 'win32'
    ? [
        { reg: 'C:\\Windows\\Fonts\\arial.ttf', bold: 'C:\\Windows\\Fonts\\arialbd.ttf' },
        { reg: 'C:\\Windows\\Fonts\\segoeui.ttf', bold: 'C:\\Windows\\Fonts\\segoeuib.ttf' }
      ]
    : process.platform === 'darwin'
    ? [
        { reg: '/System/Library/Fonts/Supplemental/Arial.ttf', bold: '/System/Library/Fonts/Supplemental/Arial Bold.ttf' },
        { reg: '/Library/Fonts/Arial.ttf', bold: '/Library/Fonts/Arial Bold.ttf' }
      ]
    : [
        { reg: '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', bold: '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf' },
        { reg: '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf', bold: '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf' }
      ];

  const match = candidates.find((c) => fs.existsSync(c.reg));
  if (match) {
    try {
      const regularB64 = fs.readFileSync(match.reg).toString('base64');
      doc.addFileToVFS('unicode.ttf', regularB64);
      doc.addFont('unicode.ttf', 'AppUnicode', 'normal');

      const hasBold = match.bold && fs.existsSync(match.bold);
      if (hasBold) {
        const boldB64 = fs.readFileSync(match.bold).toString('base64');
        doc.addFileToVFS('unicode-bold.ttf', boldB64);
        doc.addFont('unicode-bold.ttf', 'AppUnicode', 'bold');
      }

      doc.setFont('AppUnicode', style === 'bold' && hasBold ? 'bold' : 'normal');
      return 'AppUnicode';
    } catch {
      // Fallback below
    }
  }
  doc.setFont('helvetica', style);
  return 'helvetica';
}

function escapeXml(unsafe) {
  return String(unsafe || '').replace(/[<>&'"]/g, (c) => {
    switch (c) {
      case '<': return '&lt;';
      case '>': return '&gt;';
      case '&': return '&amp;';
      case '\'': return '&apos;';
      case '"': return '&quot;';
      default: return c;
    }
  });
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Builds an OpenXML .docx Buffer from an array of paragraphs.
 */
async function buildDocxBuffer(paragraphs, docTitle = 'Belge') {
  const zip = new JSZip();

  let documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Title"/><w:spacing w:after="240"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="36"/><w:color w:val="0F172A"/></w:rPr><w:t>${escapeXml(docTitle)}</w:t></w:r>
    </w:p>`;

  const safeParas = paragraphs.length > 0 ? paragraphs : ['(Belgede okunabilir metin bulunamadı)'];

  for (const para of safeParas) {
    if (!para || para.startsWith('===') || para.startsWith('---')) continue;
    const isHeading = para.startsWith('# ') || para.startsWith('## ') || para.startsWith('### ');
    const text = isHeading ? para.replace(/^#+\s*/, '') : para;
    const szVal = para.startsWith('# ') ? '28' : para.startsWith('## ') ? '24' : '22';

    documentXml += `
    <w:p>
      <w:pPr><w:spacing w:after="160" w:line="280" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr>${isHeading ? '<w:b/>' : ''}<w:sz w:val="${szVal}"/><w:color w:val="1E293B"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>
    </w:p>`;
  }

  documentXml += `
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>`;

  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);

  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);

  zip.file('word/document.xml', documentXml);

  return zip.generateAsync({
    type: 'nodebuffer',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  });
}

/**
 * Extracts structured lines/paragraphs per page from a PDF using pdfjs-dist.
 */
async function extractPdfPagesText(inputPath) {
  const data = new Uint8Array(fs.readFileSync(inputPath));
  const pdfDoc = await pdfjsLib.getDocument({ data, verbosity: PDFJS_VERBOSITY }).promise;
  const numPages = pdfDoc.numPages;
  const pages = [];

  for (let pageNum = 1; pageNum <= numPages; pageNum++) {
    const page = await pdfDoc.getPage(pageNum);
    const textContent = await page.getTextContent();
    const lines = [];
    let lastY = null;
    let currentLine = '';

    for (const item of textContent.items) {
      const str = item.str ?? '';
      const y = item.transform ? Math.round(item.transform[5]) : null;
      if (lastY !== null && y !== null && Math.abs(y - lastY) > 8) {
        if (currentLine.trim()) lines.push(currentLine.trim());
        currentLine = str;
      } else {
        currentLine += (currentLine && str && !currentLine.endsWith(' ') && !str.startsWith(' ') ? ' ' : '') + str;
      }
      lastY = y;
    }
    if (currentLine.trim()) lines.push(currentLine.trim());
    pages.push({ pageNum, lines });
  }

  return { numPages, pages };
}

function resolvePythonBin() {
  const isWin = process.platform === 'win32';
  const relPy = isWin ? path.join('Scripts', 'python.exe') : path.join('bin', 'python');
  const candidates = [
    path.join(ROOT_DIR, '.venv', relPy),
    path.join(os.homedir(), '.venv', relPy)
  ];
  for (const c of candidates) {
    if (fs.existsSync(c)) return c;
  }
  return isWin ? 'python' : 'python3';
}

/**
 * Runs the Python PyMuPDF / pdf2docx engine.
 */
async function runPythonPdfEngine(mode, inputPath, outputPath) {
  if (!fs.existsSync(PYTHON_PDF_ENGINE)) {
    throw new Error(`Python PDF motoru bulunamadı: ${PYTHON_PDF_ENGINE}`);
  }

  const pythonBin = resolvePythonBin();

  return new Promise((resolve, reject) => {
    execFile(pythonBin, [PYTHON_PDF_ENGINE, mode, inputPath, outputPath], {
      timeout: 300000,
      windowsHide: true,
      maxBuffer: 10 * 1024 * 1024
    }, (err, stdout, stderr) => {
      if (err) {
        return reject(new Error(`Python PDF Motor Hatası: ${stderr || err.message}`));
      }
      resolve(stdout.trim());
    });
  });
}

/**
 * Renders all PDF pages to PNG files inside a temp folder (PyMuPDF primary, Windows WinRT fallback).
 */
async function renderPdfPagesToPngDir(inputPath, scale = 2.0) {
  const tempOutDir = createTempPath('');
  fs.mkdirSync(tempOutDir, { recursive: true });

  // Primary: PyMuPDF C engine (fast, crisp, handles all PDFs)
  try {
    await runPythonPdfEngine('png', inputPath, tempOutDir);
    const pyFiles = fs.readdirSync(tempOutDir)
      .filter((f) => f.toLowerCase().endsWith('.png'))
      .sort()
      .map((f) => path.join(tempOutDir, f));
    if (pyFiles.length > 0) {
      return { tempOutDir, pngFiles: pyFiles };
    }
  } catch {
    // Fallback to Windows WinRT PDF renderer below
  }

  // Fallback to Windows WinRT PDF renderer on Windows only
  if (process.platform !== 'win32' || !fs.existsSync(PDF_RENDER_SCRIPT)) {
    removeQuiet(tempOutDir);
    throw new Error('PDF sayfaları görsele dönüştürülemedi (PyMuPDF kurulu olduğundan emin olun).');
  }

  const args = [
    '-NoProfile',
    '-NonInteractive',
    '-ExecutionPolicy', 'Bypass',
    '-File', PDF_RENDER_SCRIPT,
    '-InputPath', inputPath,
    '-OutputDir', tempOutDir,
    '-Scale', String(scale)
  ];

  await new Promise((resolve, reject) => {
    execFile('powershell.exe', args, { timeout: 120000, windowsHide: true }, (err, stdout, stderr) => {
      if (err) {
        return reject(new Error(`PDF görsel işleme hatası: ${stderr || err.message}`));
      }
      resolve(stdout);
    });
  });

  const files = fs.readdirSync(tempOutDir)
    .filter((f) => f.toLowerCase().endsWith('.png'))
    .sort()
    .map((f) => path.join(tempOutDir, f));

  if (files.length === 0) {
    removeQuiet(tempOutDir);
    throw new Error('PDF sayfaları görsele dönüştürülemedi.');
  }

  return { tempOutDir, pngFiles: files };
}

/**
 * Renders clean multi-page vector PDF from plain text lines using Arial Unicode TTF.
 */
function renderLinesToPdfBuffer(lines, title = '') {
  const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
  const fontFamily = applyUnicodeFont(doc, 'normal');
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 48;
  const maxWidth = pageWidth - margin * 2;
  let cursorY = margin;

  if (title) {
    applyUnicodeFont(doc, 'bold');
    doc.setFontSize(15);
    doc.setTextColor(15, 23, 42);
    const titleLines = doc.splitTextToSize(title, maxWidth);
    doc.text(titleLines, margin, cursorY);
    cursorY += titleLines.length * 20 + 12;
    doc.setFont(fontFamily, 'normal');
  }

  doc.setFontSize(10.5);
  doc.setTextColor(30, 41, 59);

  for (const rawLine of lines) {
    const line = String(rawLine ?? '');
    if (line.trim() === '') {
      cursorY += 10;
      if (cursorY > pageHeight - margin) {
        doc.addPage();
        cursorY = margin;
      }
      continue;
    }
    const wrapped = doc.splitTextToSize(line, maxWidth);
    for (const wLine of wrapped) {
      if (cursorY + 15 > pageHeight - margin) {
        doc.addPage();
        cursorY = margin;
      }
      doc.text(wLine, margin, cursorY);
      cursorY += 14.5;
    }
  }

  return Buffer.from(doc.output('arraybuffer'));
}

/**
 * Converts HTML file to PDF using headless Edge/Chrome/Chromium across Windows, macOS, and Linux, or plain text fallback.
 */
async function convertHtmlToPdf(inputPath, outputPath) {
  const browserCandidates = process.platform === 'win32'
    ? [
        'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
        'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
        'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
      ]
    : process.platform === 'darwin'
    ? [
        '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
        '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
        '/Applications/Chromium.app/Contents/MacOS/Chromium'
      ]
    : [
        '/usr/bin/google-chrome',
        '/usr/bin/chromium',
        '/usr/bin/chromium-browser',
        '/usr/bin/microsoft-edge'
      ];
  const browserExe = browserCandidates.find((p) => fs.existsSync(p));

  if (browserExe) {
    try {
      const tempProfile = createTempPath('');
      fs.mkdirSync(tempProfile, { recursive: true });
      const fileUrl = pathToFileURL(inputPath).href;
      await new Promise((resolve, reject) => {
        execFile(browserExe, [
          '--headless=new',
          '--disable-gpu',
          '--no-first-run',
          '--no-default-browser-check',
          `--user-data-dir=${tempProfile}`,
          `--print-to-pdf=${outputPath}`,
          '--print-to-pdf-no-header',
          fileUrl
        ], { timeout: 45000, windowsHide: true }, (err) => {
          if (err) reject(err);
          else resolve();
        });
      });
      removeQuiet(tempProfile);
      if (fs.existsSync(outputPath) && fs.statSync(outputPath).size > 0) {
        return outputPath;
      }
    } catch {
      // Fallback to text extraction below
    }
  }

  const html = fs.readFileSync(inputPath, 'utf-8');
  const plainText = html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<\/(p|div|h[1-6]|li|tr|br)>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"');

  const buf = renderLinesToPdfBuffer(plainText.split(/\r?\n/));
  fs.writeFileSync(outputPath, buf);
  return outputPath;
}

/**
 * Builds a multi-resolution Windows .ico file (16x16, 32x32, 48x48, 256x256 PNG-in-ICO).
 */
async function convertToIco(inputPath, outputPath) {
  const sizes = [16, 32, 48, 256];
  const pngBuffers = await Promise.all(
    sizes.map((size) =>
      sharp(inputPath, { failOnError: false })
        .resize(size, size, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } })
        .png()
        .toBuffer()
    )
  );

  // ICONDIR (6 bytes) + N * ICONDIRENTRY (16 bytes each)
  const dirSize = 6 + sizes.length * 16;
  const header = Buffer.alloc(dirSize);
  header.writeUInt16LE(0, 0); // Reserved
  header.writeUInt16LE(1, 2); // Type: 1 = ICO
  header.writeUInt16LE(sizes.length, 4); // Image count

  let offset = dirSize;
  for (let i = 0; i < sizes.length; i++) {
    const size = sizes[i];
    const buf = pngBuffers[i];
    const entryOffset = 6 + i * 16;
    header.writeUInt8(size === 256 ? 0 : size, entryOffset + 0); // Width (0 = 256)
    header.writeUInt8(size === 256 ? 0 : size, entryOffset + 1); // Height (0 = 256)
    header.writeUInt8(0, entryOffset + 2); // Color palette
    header.writeUInt8(0, entryOffset + 3); // Reserved
    header.writeUInt16LE(1, entryOffset + 4); // Color planes
    header.writeUInt16LE(32, entryOffset + 6); // Bits per pixel
    header.writeUInt32LE(buf.length, entryOffset + 8); // Payload size
    header.writeUInt32LE(offset, entryOffset + 12); // Payload offset
    offset += buf.length;
  }

  fs.writeFileSync(outputPath, Buffer.concat([header, ...pngBuffers]));
  return outputPath;
}

/**
 * Main headless conversion entry point.
 * Returns { outputPath, isFolder, message }
 */
export async function convertFileLocal(inputPath, targetFormat) {
  if (!inputPath || !fs.existsSync(inputPath)) {
    throw new Error(`Kaynak dosya bulunamadı: ${inputPath}`);
  }

  const stat = fs.statSync(inputPath);
  if (!stat.isFile()) {
    throw new Error('Lütfen bir dosya seçin (klasör dönüştürme desteklenmiyor).');
  }

  const dir = path.dirname(inputPath);
  const fullExt = path.extname(inputPath);
  const baseName = path.basename(inputPath, fullExt);
  const src = fullExt.toLowerCase().replace(/^\./, '');
  const target = (targetFormat || '').toLowerCase().replace(/^\./, '').trim();

  if (!target) {
    throw new Error('Hedef format belirtilmedi.');
  }

  // ------------------------------------------------------------------
  // 1. ARCHIVE: ZIP Extract or Pack Any File into ZIP
  // ------------------------------------------------------------------
  if (src === 'zip' && target === 'extract') {
    const outFolder = getUniqueFolderPath(dir, baseName);
    fs.mkdirSync(outFolder, { recursive: true });

    const zipData = fs.readFileSync(inputPath);
    const zip = await JSZip.loadAsync(zipData);
    const resolvedRoot = path.resolve(outFolder);

    for (const [entryName, entry] of Object.entries(zip.files)) {
      const targetEntryPath = path.resolve(outFolder, entryName);
      if (!targetEntryPath.startsWith(resolvedRoot + path.sep) && targetEntryPath !== resolvedRoot) {
        continue; // Guard against Zip Slip
      }
      if (entry.dir) {
        fs.mkdirSync(targetEntryPath, { recursive: true });
      } else {
        fs.mkdirSync(path.dirname(targetEntryPath), { recursive: true });
        const content = await entry.async('nodebuffer');
        fs.writeFileSync(targetEntryPath, content);
      }
    }

    return {
      outputPath: outFolder,
      isFolder: true,
      message: `${path.basename(outFolder)} klasörüne çıkarıldı`
    };
  }

  if (target === 'zip') {
    const outPath = getUniqueOutputPath(dir, baseName, 'zip');
    const zip = new JSZip();
    zip.file(path.basename(inputPath), fs.readFileSync(inputPath));
    const buf = await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    fs.writeFileSync(outPath, buf);
    return {
      outputPath: outPath,
      isFolder: false,
      message: `${path.basename(outPath)} oluşturuldu`
    };
  }

  // ------------------------------------------------------------------
  // 2. PDF CONVERSIONS (Powered by PyMuPDF + pdf2docx Hybrid Engine)
  // ------------------------------------------------------------------
  if (src === 'pdf') {
    if (['txt', 'md', 'html'].includes(target)) {
      const outPath = getUniqueOutputPath(dir, baseName, target);
      await runPythonPdfEngine(target, inputPath, outPath);
      return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
    }

    if (target === 'docx') {
      const outPath = getUniqueOutputPath(dir, baseName, 'docx');
      const info = await runPythonPdfEngine('docx', inputPath, outPath);
      const isScanned = info.includes('scanned:');
      const suffix = isScanned ? ' (Taranmış PDF - tam sayfa aktarıldı)' : '';
      return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu${suffix}` };
    }

    if (['png', 'jpg', 'jpeg', 'webp'].includes(target)) {
      const { tempOutDir, pngFiles } = await renderPdfPagesToPngDir(inputPath, 2.0);
      try {
        if (pngFiles.length === 1) {
          const outPath = getUniqueOutputPath(dir, baseName, target === 'jpeg' ? 'jpg' : target);
          if (target === 'png') {
            fs.copyFileSync(pngFiles[0], outPath);
          } else {
            await convertImage(pngFiles[0], outPath, target, { quality: 92 });
          }
          return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
        }

        const outFolder = getUniqueFolderPath(dir, `${baseName}_sayfalar`);
        fs.mkdirSync(outFolder, { recursive: true });
        const ext = target === 'jpeg' ? 'jpg' : target;
        for (let i = 0; i < pngFiles.length; i++) {
          const name = path.basename(pngFiles[i], '.png') + `.${ext}`;
          const dest = path.join(outFolder, name);
          if (ext === 'png') {
            fs.copyFileSync(pngFiles[i], dest);
          } else {
            await convertImage(pngFiles[i], dest, ext, { quality: 92 });
          }
        }
        return {
          outputPath: outFolder,
          isFolder: true,
          message: `${pngFiles.length} sayfa ${path.basename(outFolder)} klasörüne kaydedildi`
        };
      } finally {
        removeQuiet(tempOutDir);
      }
    }

    if (target === 'pptx') {
      const outPath = getUniqueOutputPath(dir, baseName, 'pptx');
      const { tempOutDir, pngFiles } = await renderPdfPagesToPngDir(inputPath, 2.0);
      let pagesText = [];
      try {
        const parsed = await extractPdfPagesText(inputPath);
        pagesText = parsed.pages;
      } catch {
        pagesText = [];
      }

      try {
        const pptx = new pptxgen();
        pptx.layout = 'LAYOUT_16x9';
        pptx.title = baseName;
        pptx.author = 'ConvertFlow';

        for (let i = 0; i < pngFiles.length; i++) {
          const jpgBuf = await sharp(pngFiles[i]).flatten({ background: '#ffffff' }).jpeg({ quality: 90 }).toBuffer();
          const dataUri = `data:image/jpeg;base64,${jpgBuf.toString('base64')}`;
          const slide = pptx.addSlide();
          slide.addImage({ data: dataUri, x: 0, y: 0, w: '100%', h: '100%' });

          const noteLines = pagesText[i]?.lines || [];
          if (noteLines.length > 0) {
            slide.addNotes(noteLines.join(' ').substring(0, 2000));
          }
        }

        const pptxBuf = await pptx.write({ outputType: 'nodebuffer' });
        fs.writeFileSync(outPath, pptxBuf);
        return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
      } finally {
        removeQuiet(tempOutDir);
      }
    }
  }

  // ------------------------------------------------------------------
  // 3. WORD & TEXT & MARKDOWN & HTML DOCUMENTS
  // ------------------------------------------------------------------
  if (['docx', 'doc', 'rtf'].includes(src)) {
    // Direct Office COM for PDF, DOCX, TXT
    if (canHandleWithOffice(src, target)) {
      const outPath = getUniqueOutputPath(dir, baseName, target);
      try {
        await convertWithOffice(inputPath, outPath, target);
        if (fs.existsSync(outPath)) {
          return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
        }
      } catch (officeErr) {
        if (src !== 'docx') {
          throw officeErr;
        }
        // If src === 'docx', fall through to mammoth pure JS fallback
      }
    }

    // For .doc / .rtf -> html / md, convert via temp .docx using Word COM first
    let workingDocx = inputPath;
    let tempDocx = null;
    if (src !== 'docx') {
      tempDocx = createTempPath('docx');
      await convertWithOffice(inputPath, tempDocx, 'docx');
      workingDocx = tempDocx;
    }

    try {
      const outPath = getUniqueOutputPath(dir, baseName, target);
      const docxBuffer = fs.readFileSync(workingDocx);

      if (target === 'txt') {
        const { value } = await mammoth.extractRawText({ buffer: docxBuffer });
        fs.writeFileSync(outPath, '\uFEFF' + value, 'utf-8');
        return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
      }

      if (target === 'md') {
        let md = '';
        if (mammoth.convertToMarkdown) {
          const res = await mammoth.convertToMarkdown({ buffer: docxBuffer });
          md = res.value || '';
        }
        if (!md) {
          const res = await mammoth.extractRawText({ buffer: docxBuffer });
          md = res.value || '';
        }
        fs.writeFileSync(outPath, '\uFEFF' + md, 'utf-8');
        return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
      }

      if (target === 'html') {
        const { value: bodyHtml } = await mammoth.convertToHtml({ buffer: docxBuffer });
        const fullHtml = `<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="UTF-8">
  <title>${escapeHtml(baseName)}</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; max-width: 820px; margin: 2.5rem auto; padding: 0 1.5rem; line-height: 1.65; color: #1e293b; }
    table { border-collapse: collapse; width: 100%; margin: 1.5em 0; }
    th, td { border: 1px solid #cbd5e1; padding: 0.5em 0.75em; text-align: left; }
    th { background: #f8fafc; }
    img { max-width: 100%; height: auto; }
  </style>
</head>
<body>${bodyHtml}</body>
</html>`;
        fs.writeFileSync(outPath, fullHtml, 'utf-8');
        return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
      }

      if (target === 'pdf') {
        const { value: rawText } = await mammoth.extractRawText({ buffer: docxBuffer });
        const buf = renderLinesToPdfBuffer(rawText.split(/\r?\n/));
        fs.writeFileSync(outPath, buf);
        return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
      }
    } finally {
      removeQuiet(tempDocx);
    }
  }

  if (['txt', 'md', 'html', 'htm'].includes(src)) {
    const normSrc = src === 'htm' ? 'html' : src;
    const outPath = getUniqueOutputPath(dir, baseName, target);

    if (target === 'pdf') {
      if (normSrc === 'html') {
        await convertHtmlToPdf(inputPath, outPath);
      } else {
        const content = fs.readFileSync(inputPath, 'utf-8');
        const buf = renderLinesToPdfBuffer(content.split(/\r?\n/));
        fs.writeFileSync(outPath, buf);
      }
      return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
    }

    if (target === 'docx') {
      let text = fs.readFileSync(inputPath, 'utf-8');
      if (normSrc === 'html') {
        text = text
          .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
          .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
          .replace(/<\/(p|div|h[1-6]|li|tr|br)>/gi, '\n')
          .replace(/<[^>]+>/g, '')
          .replace(/&nbsp;/g, ' ')
          .replace(/&amp;/g, '&')
          .replace(/&lt;/g, '<')
          .replace(/&gt;/g, '>');
      }
      const paragraphs = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
      const buf = await buildDocxBuffer(paragraphs, baseName);
      fs.writeFileSync(outPath, buf);
      return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
    }

    if (['txt', 'md', 'html'].includes(target)) {
      await convertDocument(inputPath, outPath, normSrc, target);
      return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
    }
  }

  // ------------------------------------------------------------------
  // 4. PRESENTATIONS (PPTX / PPT)
  // ------------------------------------------------------------------
  if (['pptx', 'ppt'].includes(src)) {
    if (target === 'pdf') {
      const outPath = getUniqueOutputPath(dir, baseName, 'pdf');
      try {
        await convertWithOffice(inputPath, outPath, 'pdf');
        if (fs.existsSync(outPath)) {
          return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
        }
      } catch {
        const slides = await parseServerPresentation(inputPath);
        const lines = [];
        for (const s of slides) {
          lines.push(`Slayt ${s.index}: ${s.title}`);
          for (const b of s.bullets) lines.push(`  • ${b}`);
          lines.push('');
        }
        fs.writeFileSync(outPath, renderLinesToPdfBuffer(lines, baseName));
        return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
      }
    }

    if (['png', 'jpg', 'jpeg'].includes(target)) {
      const imgExt = target === 'jpeg' ? 'jpg' : target;
      const tempZip = createTempPath('zip');
      try {
        await convertWithOffice(inputPath, tempZip, imgExt);
        const zip = await JSZip.loadAsync(fs.readFileSync(tempZip));
        const entries = Object.entries(zip.files).filter(([, e]) => !e.dir).sort(([a], [b]) => a.localeCompare(b));

        if (entries.length === 1) {
          const outPath = getUniqueOutputPath(dir, baseName, imgExt);
          fs.writeFileSync(outPath, await entries[0][1].async('nodebuffer'));
          return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
        }

        const outFolder = getUniqueFolderPath(dir, `${baseName}_slaytlar`);
        fs.mkdirSync(outFolder, { recursive: true });
        for (const [name, entry] of entries) {
          fs.writeFileSync(path.join(outFolder, path.basename(name)), await entry.async('nodebuffer'));
        }
        return {
          outputPath: outFolder,
          isFolder: true,
          message: `${entries.length} slayt ${path.basename(outFolder)} klasörüne kaydedildi`
        };
      } finally {
        removeQuiet(tempZip);
      }
    }

    if (['html', 'txt', 'md'].includes(target)) {
      const outPath = getUniqueOutputPath(dir, baseName, target);
      await convertServerPptx(inputPath, outPath, target);
      return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
    }
  }

  // ------------------------------------------------------------------
  // 5. IMAGES (and GIF -> MP4)
  // ------------------------------------------------------------------
  const imageFormats = ['png', 'jpg', 'jpeg', 'webp', 'avif', 'tiff', 'tif', 'bmp', 'ico', 'heic', 'heif', 'svg', 'gif'];
  if (imageFormats.includes(src)) {
    if (src === 'gif' && target === 'mp4') {
      const outPath = getUniqueOutputPath(dir, baseName, 'mp4');
      await convertMedia(inputPath, outPath, 'mp4');
      return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
    }

    if (target === 'ico') {
      const outPath = getUniqueOutputPath(dir, baseName, 'ico');
      await convertToIco(inputPath, outPath);
      return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
    }

    if (target === 'pdf') {
      const outPath = getUniqueOutputPath(dir, baseName, 'pdf');
      const { data, info } = await sharp(inputPath, { failOnError: false })
        .rotate()
        .flatten({ background: '#ffffff' })
        .jpeg({ quality: 92 })
        .toBuffer({ resolveWithObject: true });

      const orientation = info.width > info.height ? 'landscape' : 'portrait';
      const doc = new jsPDF({ orientation, unit: 'px', format: [info.width, info.height] });
      doc.addImage(data.toString('base64'), 'JPEG', 0, 0, info.width, info.height);
      fs.writeFileSync(outPath, Buffer.from(doc.output('arraybuffer')));
      return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
    }

    if (imageFormats.includes(target)) {
      const cleanTarget = target === 'jpeg' ? 'jpg' : target;
      const outPath = getUniqueOutputPath(dir, baseName, cleanTarget);
      await convertImage(inputPath, outPath, cleanTarget, { quality: 90 });
      return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
    }
  }

  // ------------------------------------------------------------------
  // 6. VIDEO & AUDIO MEDIA
  // ------------------------------------------------------------------
  const mediaFormats = ['mp4', 'mkv', 'avi', 'mov', 'webm', 'flv', 'wmv', 'mp3', 'wav', 'ogg', 'aac', 'flac', 'm4a', 'wma', 'opus', 'gif'];
  if (mediaFormats.includes(src) && mediaFormats.includes(target)) {
    const outPath = getUniqueOutputPath(dir, baseName, target);
    await convertMedia(inputPath, outPath, target);
    return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
  }

  // ------------------------------------------------------------------
  // 7. DATA & TABLES
  // ------------------------------------------------------------------
  const dataFormats = ['xlsx', 'xls', 'csv', 'tsv', 'json', 'xml', 'yaml', 'yml', 'sql'];
  if (dataFormats.includes(src)) {
    if (['xlsx', 'xls'].includes(src) && target === 'pdf') {
      const outPath = getUniqueOutputPath(dir, baseName, 'pdf');
      await convertWithOffice(inputPath, outPath, 'pdf');
      return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
    }

    if (dataFormats.includes(target) || ['html', 'md'].includes(target)) {
      const outPath = getUniqueOutputPath(dir, baseName, target);
      const rows = await parseToNormalizedData(inputPath, src);
      await renderFromNormalizedData(rows, outPath, target);
      return { outputPath: outPath, isFolder: false, message: `${path.basename(outPath)} oluşturuldu` };
    }
  }

  throw new Error(`${src.toUpperCase()} -> ${target.toUpperCase()} dönüşümü desteklenmiyor.`);
}

