// ConvertFlow - CLI & Windows Context Menu Runner
import fs from 'fs';
import os from 'os';
import path from 'path';
import { execFileSync } from 'child_process';
import { fileURLToPath } from 'url';
import { convertFileLocal } from './engine.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const NOTIFY_SCRIPT = path.resolve(__dirname, 'notify.ps1');
const LOG_FILE = path.join(os.tmpdir(), 'convertflow-cli.log');

function appendLog(line) {
  try {
    const ts = new Date().toISOString();
    fs.appendFileSync(LOG_FILE, `[${ts}] ${line}\n`, 'utf-8');
  } catch {
    // Ignore log errors
  }
}

function sendToast(title, message, isError = false) {
  if (process.argv.includes('--no-notify') || !fs.existsSync(NOTIFY_SCRIPT)) return;
  try {
    const args = [
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy', 'Bypass',
      '-File', NOTIFY_SCRIPT,
      '-Title', title,
      '-Message', message
    ];
    if (isError) args.push('-IsError');
    execFileSync('powershell.exe', args, { timeout: 10000, windowsHide: true });
  } catch {
    // Ignore toast failures
  }
}

async function main() {
  const cleanArgs = process.argv.slice(2).filter((a) => !a.startsWith('--'));
  const inputPath = cleanArgs[0];
  const targetFormat = cleanArgs[1];

  if (!inputPath || !targetFormat) {
    console.error('Kullanım: node convert-cli.js "<dosya_yolu>" "<hedef_format>"');
    process.exit(1);
  }

  const fileName = path.basename(inputPath);
  appendLog(`START: "${inputPath}" -> ${targetFormat}`);

  // For larger files or heavy formats (PDF, Office, Video), show immediate start toast
  try {
    const ext = path.extname(inputPath).toLowerCase();
    const sizeBytes = fs.existsSync(inputPath) ? fs.statSync(inputPath).size : 0;
    const isHeavy = sizeBytes > 2 * 1024 * 1024 || ['.pdf', '.mp4', '.mkv', '.avi', '.mov', '.webm', '.docx', '.pptx', '.xlsx'].includes(ext);
    if (isHeavy) {
      sendToast('ConvertFlow • İşleniyor', `${fileName} → ${targetFormat.toUpperCase()} dönüştürülüyor...`);
    }
  } catch {
    // Ignore pre-toast errors
  }

  try {
    const result = await convertFileLocal(inputPath, targetFormat);
    appendLog(`OK: "${result.outputPath}" (${result.message})`);
    console.log(`✅ ${result.message} -> ${result.outputPath}`);
    sendToast('ConvertFlow • Dönüştürüldü', `${fileName} → ${result.message}`);
  } catch (err) {
    const errMsg = err?.message || String(err);
    appendLog(`ERROR: "${inputPath}" -> ${targetFormat}: ${err?.stack || errMsg}`);
    console.error(`❌ Hata: ${errMsg}`);
    sendToast('ConvertFlow • Hata', `${fileName}: ${errMsg}`, true);
    process.exit(1);
  }
}

main();

