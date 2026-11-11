// ConvertFlow - Windows Right-Click Cascading Context Menu Uninstaller
import fs from 'fs';
import os from 'os';
import path from 'path';
import { execFileSync } from 'child_process';
import { EXTENSION_TARGETS } from './formats.js';

function psEscape(str) {
  return String(str).replace(/'/g, "''");
}

function uninstall() {
  if (process.platform !== 'win32') {
    console.error('Bu işlem yalnızca Windows içindir.');
    process.exit(1);
  }

  const extensions = Object.keys(EXTENSION_TARGETS);
  const psLines = [
    '$ErrorActionPreference = "SilentlyContinue"',
    '$hkcu = [Microsoft.Win32.Registry]::CurrentUser'
  ];

  for (const ext of extensions) {
    psLines.push(`$hkcu.DeleteSubKeyTree('Software\\Classes\\SystemFileAssociations\\.${psEscape(ext)}\\shell\\ConvertFlow', $false)`);
    psLines.push(`$hkcu.DeleteSubKeyTree('Software\\Classes\\ConvertFlow.Menu.${psEscape(ext)}', $false)`);
  }

  // Also clean any leftover ConvertFlow.Menu.* keys
  psLines.push(`
$classesKey = $hkcu.OpenSubKey('Software\\Classes', $true)
if ($classesKey) {
  foreach ($name in $classesKey.GetSubKeyNames()) {
    if ($name -like 'ConvertFlow.Menu.*') {
      $classesKey.DeleteSubKeyTree($name, $false)
    }
  }
  $classesKey.Close()
}
`);

  const tempPs1 = path.join(os.tmpdir(), `cf_uninstall_menu_${Date.now()}.ps1`);
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

  console.log('✅ ConvertFlow sağ tık menüsü Windows kayıt defterinden temiz bir şekilde kaldırıldı.');
}

uninstall();

