import fs from 'fs';
import os from 'os';
import path from 'path';
import sharp from 'sharp';
import { execFileSync } from 'child_process';
import { convertFileLocal } from '../context-menu/engine.js';

const testDir = path.join(os.tmpdir(), `cf_e2e_${Date.now()}`);
fs.mkdirSync(testDir, { recursive: true });

let passed = 0;
let failed = 0;

function check(cond, name) {
  if (cond) {
    console.log(`✅ [PASS] ${name}`);
    passed++;
  } else {
    console.error(`❌ [FAIL] ${name}`);
    failed++;
  }
}

async function run() {
  console.log(`\n=== ConvertFlow Sağ Tık Motoru Testi (${testDir}) ===\n`);

  // 1. TXT (Turkish) -> PDF -> DOCX, PPTX, PNG, JPG, WEBP, HTML, TXT, MD
  const txtFile = path.join(testDir, 'ornek_turkce.txt');
  fs.writeFileSync(txtFile, 'ConvertFlow Sağ Tık Testi\n\nMerhaba Dünya! Türkçe karakterler: ğüşıöç ĞÜŞİÖÇ.\nİkinci paragraf içeriği.', 'utf-8');

  const pdfRes = await convertFileLocal(txtFile, 'pdf');
  check(fs.existsSync(pdfRes.outputPath) && pdfRes.outputPath.endsWith('.pdf'), 'TXT -> PDF (Türkçe UTF-8)');

  const docxFromTxt = await convertFileLocal(txtFile, 'docx');
  check(fs.existsSync(docxFromTxt.outputPath) && docxFromTxt.outputPath.endsWith('.docx'), 'TXT -> DOCX');

  // Test collision-free naming when converting twice
  const pdfRes2 = await convertFileLocal(txtFile, 'pdf');
  check(path.basename(pdfRes2.outputPath) === 'ornek_turkce (1).pdf', 'Çakışmasız dosya adlandırma (ornek_turkce (1).pdf)');

  // PDF -> DOCX, PPTX, PNG, JPG, WEBP, HTML, TXT, MD
  const pdfPath = pdfRes.outputPath;
  for (const tgt of ['docx', 'pptx', 'png', 'jpg', 'webp', 'html', 'txt', 'md']) {
    const r = await convertFileLocal(pdfPath, tgt);
    const exists = fs.existsSync(r.outputPath);
    const nonEmpty = exists && (r.isFolder ? fs.readdirSync(r.outputPath).length > 0 : fs.statSync(r.outputPath).size > 0);
    check(nonEmpty, `PDF -> ${tgt.toUpperCase()} (${path.basename(r.outputPath)})`);
  }

  // DOCX -> HTML, TXT, MD, PDF
  const docxPath = docxFromTxt.outputPath;
  for (const tgt of ['html', 'txt', 'md', 'pdf']) {
    const r = await convertFileLocal(docxPath, tgt);
    check(fs.existsSync(r.outputPath) && fs.statSync(r.outputPath).size > 0, `DOCX -> ${tgt.toUpperCase()}`);
  }

  // 2. PNG -> JPG, WEBP, AVIF, ICO, PDF
  const pngFile = path.join(testDir, 'logo.png');
  await sharp({
    create: { width: 128, height: 128, channels: 4, background: { r: 99, g: 102, b: 241, alpha: 1 } }
  }).png().toFile(pngFile);

  for (const tgt of ['jpg', 'webp', 'avif', 'ico', 'pdf']) {
    const r = await convertFileLocal(pngFile, tgt);
    check(fs.existsSync(r.outputPath) && fs.statSync(r.outputPath).size > 0, `PNG -> ${tgt.toUpperCase()}`);
  }

  // 3. JSON -> XLSX, CSV, XML, YAML, SQL, HTML, MD -> XLSX back to JSON
  const jsonFile = path.join(testDir, 'veri.json');
  fs.writeFileSync(jsonFile, JSON.stringify([
    { id: 1, ad: 'Ahmet Yılmaz', sehir: 'İstanbul' },
    { id: 2, ad: 'Ayşe Kaya', sehir: 'Ankara' }
  ], null, 2), 'utf-8');

  let xlsxPath = null;
  for (const tgt of ['xlsx', 'csv', 'xml', 'yaml', 'sql', 'html', 'md']) {
    const r = await convertFileLocal(jsonFile, tgt);
    if (tgt === 'xlsx') xlsxPath = r.outputPath;
    check(fs.existsSync(r.outputPath) && fs.statSync(r.outputPath).size > 0, `JSON -> ${tgt.toUpperCase()}`);
  }

  const jsonFromXlsx = await convertFileLocal(xlsxPath, 'json');
  const parsedBack = JSON.parse(fs.readFileSync(jsonFromXlsx.outputPath, 'utf-8'));
  check(Array.isArray(parsedBack) && parsedBack.length === 2 && parsedBack[0].sehir === 'İstanbul', 'XLSX -> JSON gidiş-dönüş veri doğruluğu');

  // 4. ZIP pack & extract
  const zipRes = await convertFileLocal(txtFile, 'zip');
  check(fs.existsSync(zipRes.outputPath) && zipRes.outputPath.endsWith('.zip'), 'Dosya -> ZIP Arşivi');

  const extractRes = await convertFileLocal(zipRes.outputPath, 'extract');
  check(fs.existsSync(path.join(extractRes.outputPath, 'ornek_turkce.txt')), 'ZIP -> Klasöre Çıkart');

  // 5. Silent VBScript runner check (using --no-notify CLI via wscript? Let's invoke convert-cli directly with --no-notify and check output)
  const vbsTestTxt = path.join(testDir, 'vbs_test.txt');
  fs.writeFileSync(vbsTestTxt, 'Sessiz VBS başlatma testi', 'utf-8');
  execFileSync(process.execPath, [
    path.resolve('context-menu', 'convert-cli.js'),
    vbsTestTxt,
    'html',
    '--no-notify'
  ], { stdio: 'inherit' });
  check(fs.existsSync(path.join(testDir, 'vbs_test.html')), 'CLI runner (convert-cli.js) -> HTML');

  console.log(`\nSonuç: ${passed} başarılı, ${failed} başarısız.`);

  // Cleanup temp test folder
  fs.rmSync(testDir, { recursive: true, force: true });
  if (failed > 0) process.exit(1);
}

run().catch((err) => {
  console.error('Test çöktü:', err);
  fs.rmSync(testDir, { recursive: true, force: true });
  process.exit(1);
});

