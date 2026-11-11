// ConvertFlow - Universal Cross-Platform Setup Wizard (Windows, macOS, Linux)
import fs from 'fs';
import path from 'path';
import { execFileSync, spawnSync } from 'child_process';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT_DIR = path.resolve(__dirname, '..');
const VENV_DIR = path.join(ROOT_DIR, '.venv');

function findPython() {
  const candidates = process.platform === 'win32' ? ['python', 'py', 'python3'] : ['python3', 'python'];
  for (const cmd of candidates) {
    const res = spawnSync(cmd, ['--version'], { stdio: 'pipe', windowsHide: true });
    if (res.status === 0) return cmd;
  }
  return null;
}

function setupPythonVenv() {
  const isWin = process.platform === 'win32';
  const venvPy = isWin
    ? path.join(VENV_DIR, 'Scripts', 'python.exe')
    : path.join(VENV_DIR, 'bin', 'python');

  if (!fs.existsSync(venvPy)) {
    const sysPy = findPython();
    if (!sysPy) {
      console.warn('⚠️  Python bulunamadı. PDF -> DOCX/Görsel motoru için Python 3.10+ gereklidir.');
      return false;
    }
    console.log('🐍 Proje içi izole Python ortamı (.venv) oluşturuluyor...');
    execFileSync(sysPy, ['-m', 'venv', VENV_DIR], { stdio: 'inherit' });
  } else {
    console.log('🐍 Mevcut Python ortamı (.venv) bulundu.');
  }

  console.log('📦 Python PDF motoru bağımlılıkları kontrol ediliyor (pymupdf, pdf2docx, python-docx)...');
  const check = spawnSync(venvPy, ['-c', 'import pymupdf, pdf2docx, docx'], { stdio: 'pipe', windowsHide: true });
  if (check.status !== 0) {
    execFileSync(venvPy, ['-m', 'pip', 'install', '--upgrade', 'pymupdf', 'pdf2docx', 'python-docx'], {
      stdio: 'inherit'
    });
  } else {
    console.log('✅ Python PDF paketleri zaten hazır.');
  }
  return true;
}

async function main() {
  console.log('======================================================');
  console.log('   ConvertFlow - Çoklu Platform Kurulum Sihirbazı');
  console.log(`   Platform: ${process.platform} (${process.arch})`);
  console.log('======================================================\n');

  setupPythonVenv();

  console.log('\n🖱️  Sağ tık menüsü kuruluyor...');
  execFileSync(process.execPath, [path.join(__dirname, 'install-menu.js')], { stdio: 'inherit' });

  console.log('\n🎉 Kurulum tamamlandı! Artık dosyalarınıza sağ tıklayarak anında dönüştürebilirsiniz.');
}

main().catch((err) => {
  console.error('❌ Kurulum hatası:', err.message || err);
  process.exit(1);
});

