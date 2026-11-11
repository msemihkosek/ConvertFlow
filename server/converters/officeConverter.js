import { execFile } from 'child_process';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const SCRIPT_PATH = path.resolve(__dirname, '..', 'scripts', 'officeConvert.ps1');

/**
 * Checks if a given extension and target format can be handled natively by Microsoft Office COM
 */
export function canHandleWithOffice(ext, targetFormat) {
  if (process.platform !== 'win32') return false;

  const source = (ext || '').toLowerCase().replace(/^\./, '');
  const target = (targetFormat || '').toLowerCase();

  const pptFormats = ['ppt', 'pptx', 'pot', 'potx', 'pps', 'ppsx'];
  const wordFormats = ['doc', 'docx', 'rtf', 'dot', 'dotx'];
  const excelFormats = ['xls', 'xlsx', 'xlsm', 'xlsb'];

  if (pptFormats.includes(source)) {
    return ['pdf', 'pptx', 'png', 'jpg', 'jpeg', 'zip'].includes(target);
  }

  if (wordFormats.includes(source)) {
    return ['pdf', 'docx', 'txt'].includes(target);
  }

  if (excelFormats.includes(source)) {
    return ['pdf', 'xlsx', 'csv'].includes(target);
  }

  return false;
}

/**
 * Converts a document using Microsoft Office COM via PowerShell
 * 
 * @param {string} inputPath 
 * @param {string} outputPath 
 * @param {string} targetFormat 
 * @returns {Promise<string>} outputPath
 */
export async function convertWithOffice(inputPath, outputPath, targetFormat) {
  if (!fs.existsSync(SCRIPT_PATH)) {
    throw new Error(`Office dönüştürme betiği bulunamadı: ${SCRIPT_PATH}`);
  }

  const args = [
    '-NoProfile',
    '-NonInteractive',
    '-ExecutionPolicy', 'Bypass',
    '-File', SCRIPT_PATH,
    '-InputPath', inputPath,
    '-OutputPath', outputPath,
    '-TargetFormat', targetFormat
  ];

  return new Promise((resolve, reject) => {
    execFile('powershell.exe', args, { timeout: 120000, windowsHide: true }, (err, stdout, stderr) => {
      if (err) {
        console.error('[ConvertFlow Office Engine Error]:', stderr || err.message);
        return reject(new Error(`Microsoft Office dönüştürme hatası: ${stderr || err.message}`));
      }

      if (!fs.existsSync(outputPath)) {
        return reject(new Error('Office dönüştürme tamamlandı ancak çıktı dosyası üretilemedi.'));
      }

      console.log(`[ConvertFlow Office Engine] Başarıyla tamamlandı: ${path.basename(outputPath)}`);
      resolve(outputPath);
    });
  });
}
