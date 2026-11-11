// ConvertFlow - Optional Windows 11 Classic Context Menu Switcher
// Makes ConvertFlow (and all full right-click tools) appear immediately on the first right-click in Windows 11.
import { execFileSync } from 'child_process';

const mode = (process.argv[2] || '').toLowerCase();
const CLSID_KEY = 'HKCU\\Software\\Classes\\CLSID\\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\\InprocServer32';

function restartExplorer() {
  try {
    execFileSync('powershell.exe', [
      '-NoProfile', '-NonInteractive', '-Command',
      'Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue; Start-Process explorer.exe'
    ], { stdio: 'inherit', windowsHide: true });
  } catch (err) {
    console.warn('Explorer yeniden başlatılamadı, oturumu kapatıp açabilirsiniz:', err.message);
  }
}

if (mode === 'enable' || mode === 'classic') {
  execFileSync('reg.exe', ['add', CLSID_KEY, '/f', '/ve', '/d', ''], { stdio: 'inherit', windowsHide: true });
  restartExplorer();
  console.log('✅ Windows 11 klasik (doğrudan) sağ tık menüsü aktif edildi. Artık ilk sağ tıkta ConvertFlow görünecek!');
} else if (mode === 'disable' || mode === 'modern') {
  try {
    execFileSync('reg.exe', ['delete', 'HKCU\\Software\\Classes\\CLSID\\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}', '/f'], { stdio: 'inherit', windowsHide: true });
  } catch {
    // Already absent
  }
  restartExplorer();
  console.log('✅ Windows 11 varsayılan modern sağ tık menüsüne geri dönüldü.');
} else {
  console.log('Kullanım:');
  console.log('  node context-menu/win11-classic-menu.js enable   -> İlk sağ tıkta doğrudan menüyü açar');
  console.log('  node context-menu/win11-classic-menu.js disable  -> Varsayılan Windows 11 menüsüne döner');
}

