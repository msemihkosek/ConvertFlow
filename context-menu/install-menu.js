// ConvertFlow - Windows Right-Click Cascading Context Menu Installer
import fs from 'fs';
import os from 'os';
import path from 'path';
import { execFileSync } from 'child_process';
import { fileURLToPath } from 'url';
import sharp from 'sharp';
import { EXTENSION_TARGETS, getMenuTargetsForExtension } from './formats.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const RUN_HIDDEN_VBS = path.resolve(__dirname, 'run-hidden.vbs');
const ICON_PATH = path.resolve(__dirname, 'convertflow.ico');
const NODE_EXE = process.execPath;

/**
 * Generates a crisp ConvertFlow .ico icon using Sharp so the right-click menu looks native.
 */
async function ensureIconExists() {
  const svg = `
  <svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" viewBox="0 0 128 128">
    <defs>
      <linearGradient id="g" x1="0%" y1="0%" x2="100%" y2="100%">
        <stop offset="0%" stop-color="#6366f1"/>
        <stop offset="100%" stop-color="#a855f7"/>
      </linearGradient>
    </defs>
    <rect x="6" y="6" width="116" height="116" rx="28" fill="url(#g)"/>
    <path d="M40 48h36l-10-10a5 5 0 0 1 7-7l19 19a5 5 0 0 1 0 7L73 76a5 5 0 0 1-7-7l10-10H40a5 5 0 0 1 0-11z" fill="#ffffff"/>
    <path d="M88 80H52l10 10a5 5 0 0 1-7 7L36 78a5 5 0 0 1 0-7l19-19a5 5 0 0 1 7 7L52 69h36a5 5 0 0 1 0 11z" fill="#e0e7ff" opacity="0.92"/>
  </svg>`;

  const size = 64;
  const pngBuffer = await sharp(Buffer.from(svg))
    .resize(size, size)
    .png()
    .toBuffer();

  const header = Buffer.alloc(22);
  header.writeUInt16LE(0, 0);
  header.writeUInt16LE(1, 2);
  header.writeUInt16LE(1, 4);
  header.writeUInt8(size, 6);
  header.writeUInt8(size, 7);
  header.writeUInt8(0, 8);
  header.writeUInt8(0, 9);
  header.writeUInt16LE(1, 10);
  header.writeUInt16LE(32, 12);
  header.writeUInt32LE(pngBuffer.length, 14);
  header.writeUInt32LE(22, 18);

  fs.writeFileSync(ICON_PATH, Buffer.concat([header, pngBuffer]));
}

function psEscape(str) {
  return String(str).replace(/'/g, "''");
}

async function install() {
  if (process.platform !== 'win32') {
    console.error('Bu sağ tık menüsü kurulumu yalnızca Windows içindir.');
    process.exit(1);
  }

  if (!fs.existsSync(RUN_HIDDEN_VBS)) {
    throw new Error(`Gerekli başlatıcı bulunamadı: ${RUN_HIDDEN_VBS}`);
  }

  await ensureIconExists();

  const extensions = Object.keys(EXTENSION_TARGETS);
  const psLines = [
    '$ErrorActionPreference = "Stop"',
    '$hkcu = [Microsoft.Win32.Registry]::CurrentUser',
    'function Set-RegVal($subKey, $name, $val) {',
    '  $k = $hkcu.CreateSubKey($subKey, $true)',
    '  $k.SetValue($name, $val, [Microsoft.Win32.RegistryValueKind]::String)',
    '  $k.Close()',
    '}'
  ];

  let totalCommands = 0;

  for (const ext of extensions) {
    const menuClass = `ConvertFlow.Menu.${ext}`;
    const assocKey = `Software\\Classes\\SystemFileAssociations\\.${ext}\\shell\\ConvertFlow`;

    // Clear old submenu definition for clean re-install
    psLines.push(`$hkcu.DeleteSubKeyTree('Software\\Classes\\${psEscape(menuClass)}', $false)`);

    // Register top-level cascading entry on the extension
    psLines.push(`Set-RegVal '${psEscape(assocKey)}' 'MUIVerb' 'ConvertFlow ile Dönüştür'`);
    psLines.push(`Set-RegVal '${psEscape(assocKey)}' 'Icon' '${psEscape(ICON_PATH)}'`);
    psLines.push(`Set-RegVal '${psEscape(assocKey)}' 'ExtendedSubCommandsKey' '${psEscape(menuClass)}'`);

    const targets = getMenuTargetsForExtension(ext);
    targets.forEach((item, idx) => {
      const orderPrefix = String(idx + 1).padStart(2, '0');
      const itemKey = `Software\\Classes\\${menuClass}\\shell\\${orderPrefix}_${item.target}`;
      const cmdKey = `${itemKey}\\command`;
      const cmdValue = `wscript.exe //B //NoLogo "${RUN_HIDDEN_VBS}" "%1" "${item.target}" "${NODE_EXE}"`;

      psLines.push(`Set-RegVal '${psEscape(itemKey)}' 'MUIVerb' '${psEscape(item.label)}'`);
      psLines.push(`Set-RegVal '${psEscape(cmdKey)}' '' '${psEscape(cmdValue)}'`);
      totalCommands++;
    });
  }

  const tempPs1 = path.join(os.tmpdir(), `cf_install_menu_${Date.now()}.ps1`);
  // UTF-8 with BOM so PowerShell 5.1 reads Turkish characters cleanly
  fs.writeFileSync(tempPs1, '\uFEFF' + psLines.join('\r\n'), 'utf-8');

  try {
    execFileSync('powershell.exe', [
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy', 'Bypass',
      '-File', tempPs1
    ], { stdio: 'inherit', windowsHide: true });
  } finally {
    try { fs.unlinkSync(tempPs1); } catch {}
  }

  console.log(`\n✅ ConvertFlow sağ tık menüsü başarıyla kuruldu!`);
  console.log(`   • Kayıtlı dosya uzantısı : ${extensions.length} adet (.pdf, .docx, .png, .mp4, .xlsx, .zip vb.)`);
  console.log(`   • Toplam hızlı dönüşüm   : ${totalCommands} adet`);
  console.log(`   • Simge                  : ${ICON_PATH}`);
  console.log(`\n💡 Windows 11 ipucu: Dosyaya "Shift + Sağ Tık" yaparak veya "Sağ Tık -> Diğer seçenekleri göster" menüsünden "ConvertFlow ile Dönüştür" menüsüne hemen ulaşabilirsiniz.\n`);
}

install().catch((err) => {
  console.error('❌ Kurulum hatası:', err);
  process.exit(1);
});

