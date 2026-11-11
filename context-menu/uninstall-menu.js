// ConvertFlow - Cross-Platform Context Menu Uninstaller (Windows, macOS, Linux)
import fs from 'fs';
import os from 'os';
import path from 'path';
import { execFileSync } from 'child_process';
import { EXTENSION_TARGETS } from './formats.js';

function psEscape(str) {
  return String(str).replace(/'/g, "''");
}

function uninstallWindows() {
  const extensions = Object.keys(EXTENSION_TARGETS);
  const psLines = [
    '$ErrorActionPreference = "SilentlyContinue"',
    '$hkcu = [Microsoft.Win32.Registry]::CurrentUser'
  ];

  for (const ext of extensions) {
    psLines.push(`$hkcu.DeleteSubKeyTree('Software\\Classes\\SystemFileAssociations\\.${psEscape(ext)}\\shell\\ConvertFlow', $false)`);
    psLines.push(`$hkcu.DeleteSubKeyTree('Software\\Classes\\ConvertFlow.Menu.${psEscape(ext)}', $false)`);
  }

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

function uninstallMacOS() {
  const workflowDir = path.join(os.homedir(), 'Library', 'Services', 'ConvertFlow ile Dönüştür.workflow');
  if (fs.existsSync(workflowDir)) {
    fs.rmSync(workflowDir, { recursive: true, force: true });
  }
  try {
    execFileSync('/System/Library/CoreServices/pbs', ['-flush'], { stdio: 'ignore' });
  } catch {
    // Ignore
  }
  console.log('✅ ConvertFlow Quick Action macOS üzerinden kaldırıldı.');
}

function uninstallLinux() {
  const pathsToRemove = [
    path.join(os.homedir(), '.local', 'share', 'nautilus', 'scripts', 'ConvertFlow ile Dönüştür'),
    path.join(os.homedir(), '.local', 'share', 'nemo', 'scripts', 'ConvertFlow ile Dönüştür'),
    path.join(os.homedir(), '.local', 'share', 'caja', 'scripts', 'ConvertFlow ile Dönüştür'),
    path.join(os.homedir(), '.local', 'share', 'kio', 'servicemenus', 'convertflow.desktop'),
    path.join(os.homedir(), '.local', 'share', 'kservices5', 'ServiceMenus', 'convertflow.desktop')
  ];

  for (const p of pathsToRemove) {
    if (fs.existsSync(p)) {
      fs.rmSync(p, { force: true });
    }
  }
  console.log('✅ ConvertFlow sağ tık betikleri Linux üzerinden kaldırıldı.');
}

if (process.platform === 'win32') {
  uninstallWindows();
} else if (process.platform === 'darwin') {
  uninstallMacOS();
} else if (process.platform === 'linux') {
  uninstallLinux();
} else {
  console.error(`Desteklenmeyen işletim sistemi: ${process.platform}`);
  process.exit(1);
}
