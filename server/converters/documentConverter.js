import fs from 'fs';
import path from 'path';
import { PDFDocument, rgb } from 'pdf-lib';

/**
 * Server-side Document and PDF manipulation engine
 */

export async function convertDocument(inputPath, outputPath, sourceExt, targetExt, options = {}) {
  const src = sourceExt.toLowerCase().replace('.', '');
  const target = targetExt.toLowerCase().replace('.', '');
  const content = fs.readFileSync(inputPath);

  // 1. Text / Markdown / HTML to PDF
  if (['txt', 'md', 'html'].includes(src) && target === 'pdf') {
    const pdfDoc = await PDFDocument.create();
    const page = pdfDoc.addPage([595.28, 841.89]); // A4
    const text = content.toString('utf-8');
    const lines = text.split(/\r?\n/).slice(0, 45); // First page preview/simple write
    
    let y = 800;
    for (const line of lines) {
      page.drawText(line.substring(0, 80), {
        x: 50,
        y,
        size: 10,
        color: rgb(0.1, 0.1, 0.1)
      });
      y -= 16;
      if (y < 50) break;
    }

    const pdfBytes = await pdfDoc.save();
    fs.writeFileSync(outputPath, pdfBytes);
    return outputPath;
  }

  // 2. Text / Markdown / HTML cross text conversions
  if (['txt', 'md', 'html'].includes(src) && ['txt', 'md', 'html'].includes(target)) {
    let str = content.toString('utf-8');
    if (src === 'html' && (target === 'txt' || target === 'md')) {
      str = str.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
               .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
               .replace(/<[^>]+>/g, '\n')
               .replace(/\n\s*\n/g, '\n\n');
    } else if ((src === 'txt' || src === 'md') && target === 'html') {
      str = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="font-family:sans-serif;white-space:pre-wrap;padding:20px;">${escapeHtml(str)}</body></html>`;
    }
    fs.writeFileSync(outputPath, str, 'utf-8');
    return outputPath;
  }

  throw new Error(`Dönüşüm ${src} -> ${target} henüz bu modülde desteklenmiyor.`);
}

function escapeHtml(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
