import JSZip from 'jszip';
import * as CFB from 'cfb';
import fs from 'fs';
import path from 'path';

/**
 * Server-side Universal Presentation Parser (Supports both .pptx and .ppt)
 */
export async function parseServerPresentation(inputPath) {
  const fileBuffer = fs.readFileSync(inputPath);
  const bytes = new Uint8Array(fileBuffer);
  const filename = path.basename(inputPath);

  const isZip = bytes[0] === 0x50 && bytes[1] === 0x4B;
  const isOle = bytes[0] === 0xD0 && bytes[1] === 0xCF && bytes[2] === 0x11 && bytes[3] === 0xE0;

  if (isOle) {
    return parsePptBinary(bytes, filename);
  }

  try {
    return await parsePptxZip(fileBuffer, filename);
  } catch (err) {
    return parsePptBinary(bytes, filename);
  }
}

async function parsePptxZip(fileBuffer, filename) {
  const zip = await JSZip.loadAsync(fileBuffer);
  const slideFilenames = Object.keys(zip.files)
    .filter(name => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
    .sort((a, b) => {
      const numA = parseInt(a.replace(/\D/g, ''), 10);
      const numB = parseInt(b.replace(/\D/g, ''), 10);
      return numA - numB;
    });

  const slides = [];
  for (let i = 0; i < slideFilenames.length; i++) {
    const sName = slideFilenames[i];
    const xml = await zip.files[sName].async('text');

    const paragraphs = [];
    const pMatches = xml.matchAll(/<a:p(?:>|\s[^>]*>)([\s\S]*?)<\/a:p>/gi);
    for (const pMatch of pMatches) {
      const tMatches = pMatch[1].matchAll(/<a:t(?:>|\s[^>]*>)([^<]*)<\/a:t>/gi);
      const parts = [];
      for (const t of tMatches) {
        if (t[1]) parts.push(t[1]);
      }
      const line = parts.join('').trim();
      if (line) paragraphs.push(line);
    }

    const title = paragraphs[0] || `Slayt ${i + 1}`;
    const bullets = paragraphs.slice(1);
    slides.push({ index: i + 1, title, bullets });
  }

  return slides;
}

function parsePptBinary(bytes, filename) {
  const cfb = CFB.read(bytes, { type: 'array' });
  const pptEntry = cfb.FileIndex.find(f => f.name.includes('PowerPoint Document'));
  if (!pptEntry || !pptEntry.content) {
    throw new Error('PPT dosyasında "PowerPoint Document" akışı bulunamadı.');
  }

  const stream = new Uint8Array(pptEntry.content);
  const slides = [];
  let currentSlideTexts = [];
  let slideIndex = 1;

  function scan(start, end) {
    let offset = start;
    while (offset + 8 <= end) {
      const recVer = stream[offset] & 0x0F;
      const recType = stream[offset + 2] | (stream[offset + 3] << 8);
      const recLen = stream[offset + 4] | (stream[offset + 5] << 8) | (stream[offset + 6] << 16) | (stream[offset + 7] << 24);

      if (offset + 8 + recLen > end) break;

      if (recType === 0x03EF) {
        if (currentSlideTexts.length > 0) {
          addSlide(currentSlideTexts);
          currentSlideTexts = [];
        }
      }

      if (recVer === 0x0F) {
        scan(offset + 8, offset + 8 + recLen);
        offset += 8 + recLen;
      } else {
        let str = null;
        if (recType === 0x0FA0 && recLen > 0) {
          str = Buffer.from(stream.subarray(offset + 8, offset + 8 + recLen)).toString('utf16le');
        } else if (recType === 0x0FA8 && recLen > 0) {
          str = Buffer.from(stream.subarray(offset + 8, offset + 8 + recLen)).toString('latin1');
        } else if (recType === 0x0FBA && recLen > 0) {
          str = Buffer.from(stream.subarray(offset + 8, offset + 8 + recLen)).toString('utf16le');
        }

        if (str) {
          const lines = str.split(/\r?\n|\r|\x0B|\x0C/).map(l => l.trim()).filter(l => l.length > 0 && !isNoise(l));
          for (const line of lines) {
            if (currentSlideTexts.length === 0 || currentSlideTexts[currentSlideTexts.length - 1] !== line) {
              currentSlideTexts.push(line);
            }
          }
        }

        offset += 8 + recLen;
      }
    }
  }

  function addSlide(texts) {
    const filtered = texts.filter(t => !isNoise(t));
    if (filtered.length === 0) return;
    const title = filtered[0];
    const bullets = filtered.slice(1);
    slides.push({ index: slideIndex++, title, bullets });
  }

  function isNoise(text) {
    const noise = [
      'Click to edit Master title style',
      'Click to edit Master text styles',
      'Second level', 'Third level', 'Fourth level', 'Fifth level',
      'Calibri', 'Arial', 'Times New Roman', 'Courier New', 'Symbol', 'Wingdings'
    ];
    if (noise.includes(text)) return true;
    if (/^___PPT\d+$/i.test(text)) return true;
    return text.length < 2;
  }

  scan(0, stream.length);

  if (currentSlideTexts.length > 0) {
    addSlide(currentSlideTexts);
  }

  return slides;
}

export async function convertServerPptx(inputPath, outputPath, targetFormat) {
  const slides = await parseServerPresentation(inputPath);
  const target = targetFormat.toLowerCase();

  if (target === 'txt') {
    let txt = `POWERPOINT SUNUM ÇIKTISI (Toplam ${slides.length} Slayt)\n\n`;
    for (const s of slides) {
      txt += `[SLAYT ${s.index}]: ${s.title}\n`;
      for (const b of s.bullets) {
        txt += `  • ${b}\n`;
      }
      txt += '\n';
    }
    fs.writeFileSync(outputPath, txt, 'utf-8');
    return outputPath;
  }

  if (target === 'md') {
    let md = `# Sunum Notları\n\n`;
    for (const s of slides) {
      md += `## Slayt ${s.index}: ${s.title}\n\n`;
      for (const b of s.bullets) {
        md += `- ${b}\n`;
      }
      md += '\n';
    }
    fs.writeFileSync(outputPath, md, 'utf-8');
    return outputPath;
  }

  if (target === 'html') {
    let html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Sunum</title><style>body{font-family:sans-serif;max-width:800px;margin:2rem auto;background:#0f172a;color:#f8fafc;padding:1rem;} .card{background:#1e293b;border-radius:12px;padding:1.5rem;margin-bottom:1.5rem;} h2{color:#a5b4fc;} li{margin-bottom:0.5rem;color:#cbd5e1;}</style></head><body><h1>Sunum İçeriği</h1>`;
    for (const s of slides) {
      html += `<div class="card"><h2>Slayt ${s.index}: ${escapeHtml(s.title)}</h2><ul>${s.bullets.map(b => `<li>${escapeHtml(b)}</li>`).join('')}</ul></div>`;
    }
    html += `</body></html>`;
    fs.writeFileSync(outputPath, html, 'utf-8');
    return outputPath;
  }

  throw new Error(`Sunucu motoru PPTX -> ${target.toUpperCase()} dönüşümünü desteklemiyor.`);
}

function escapeHtml(str) {
  return String(str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
