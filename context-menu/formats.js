// ConvertFlow - Windows Context Menu Format Definitions
// Maps each source file extension to its exact, supported right-click submenu targets.

export const TARGET_LABELS = {
  pdf: 'PDF Belgesine',
  docx: 'DOCX (Word)',
  pptx: 'PPTX (PowerPoint)',
  html: 'HTML (Web Sayfası)',
  txt: 'TXT (Düz Metin)',
  md: 'MD (Markdown)',
  png: 'PNG (Görsel)',
  jpg: 'JPG (Görsel)',
  webp: 'WEBP (Web Görsel)',
  avif: 'AVIF (Yeni Nesil Görsel)',
  ico: 'ICO (Simge)',
  tiff: 'TIFF (Baskı Görseli)',
  gif: 'GIF (Animasyon)',
  mp4: 'MP4 (Video)',
  mkv: 'MKV (Video)',
  webm: 'WEBM (Web Video)',
  avi: 'AVI (Video)',
  mov: 'MOV (QuickTime)',
  mp3: 'MP3 (Ses)',
  wav: 'WAV (Kayıpsız Ses)',
  aac: 'AAC (Ses)',
  flac: 'FLAC (Kayıpsız Ses)',
  ogg: 'OGG (Ses)',
  m4a: 'M4A (Apple Ses)',
  opus: 'OPUS (Ses)',
  xlsx: 'XLSX (Excel Tablosu)',
  csv: 'CSV (Tablo)',
  tsv: 'TSV (Sekmeli Tablo)',
  json: 'JSON (Veri)',
  xml: 'XML (Veri)',
  yaml: 'YAML (Veri)',
  sql: 'SQL (Veritabanı)',
  zip: 'ZIP (Arşivle)',
  extract: 'Klasöre Çıkart'
};

const IMAGE_TARGETS = ['png', 'jpg', 'webp', 'avif', 'ico', 'tiff', 'gif', 'pdf', 'zip'];
const VIDEO_TARGETS = ['mp4', 'mkv', 'webm', 'avi', 'mov', 'gif', 'mp3', 'wav', 'aac', 'flac', 'm4a', 'zip'];
const AUDIO_TARGETS = ['mp3', 'wav', 'aac', 'flac', 'ogg', 'm4a', 'opus', 'zip'];
const DATA_TARGETS = ['xlsx', 'csv', 'json', 'xml', 'yaml', 'sql', 'html', 'md', 'zip'];

export const EXTENSION_TARGETS = {
  // Documents
  pdf: ['docx', 'pptx', 'png', 'jpg', 'webp', 'html', 'txt', 'md', 'zip'],
  docx: ['pdf', 'html', 'txt', 'md', 'zip'],
  doc: ['pdf', 'docx', 'html', 'txt', 'zip'],
  rtf: ['pdf', 'docx', 'html', 'txt', 'zip'],
  pptx: ['pdf', 'png', 'jpg', 'html', 'txt', 'md', 'zip'],
  ppt: ['pdf', 'png', 'jpg', 'html', 'txt', 'md', 'zip'],
  txt: ['pdf', 'docx', 'html', 'md', 'zip'],
  md: ['pdf', 'docx', 'html', 'txt', 'zip'],
  html: ['pdf', 'docx', 'txt', 'md', 'zip'],
  htm: ['pdf', 'docx', 'txt', 'md', 'zip'],

  // Images
  png: IMAGE_TARGETS,
  jpg: IMAGE_TARGETS,
  jpeg: IMAGE_TARGETS,
  webp: IMAGE_TARGETS,
  avif: IMAGE_TARGETS,
  heic: IMAGE_TARGETS,
  heif: IMAGE_TARGETS,
  tiff: IMAGE_TARGETS,
  tif: IMAGE_TARGETS,
  bmp: IMAGE_TARGETS,
  ico: IMAGE_TARGETS,
  svg: IMAGE_TARGETS,
  gif: ['mp4', 'webp', 'png', 'jpg', 'pdf', 'zip'],

  // Videos
  mp4: VIDEO_TARGETS,
  mkv: VIDEO_TARGETS,
  avi: VIDEO_TARGETS,
  mov: VIDEO_TARGETS,
  webm: VIDEO_TARGETS,
  flv: VIDEO_TARGETS,
  wmv: VIDEO_TARGETS,

  // Audio
  mp3: AUDIO_TARGETS,
  wav: AUDIO_TARGETS,
  ogg: AUDIO_TARGETS,
  aac: AUDIO_TARGETS,
  flac: AUDIO_TARGETS,
  m4a: AUDIO_TARGETS,
  wma: AUDIO_TARGETS,
  opus: AUDIO_TARGETS,

  // Data
  xlsx: ['pdf', ...DATA_TARGETS],
  xls: ['pdf', ...DATA_TARGETS],
  csv: DATA_TARGETS,
  tsv: DATA_TARGETS,
  json: DATA_TARGETS,
  xml: DATA_TARGETS,
  yaml: DATA_TARGETS,
  yml: DATA_TARGETS,

  // Archive
  zip: ['extract']
};

const ALIAS_GROUPS = {
  jpg: ['jpg', 'jpeg'],
  jpeg: ['jpg', 'jpeg'],
  tiff: ['tiff', 'tif'],
  tif: ['tiff', 'tif'],
  yaml: ['yaml', 'yml'],
  yml: ['yaml', 'yml'],
  html: ['html', 'htm'],
  htm: ['html', 'htm']
};

export function getMenuTargetsForExtension(ext) {
  const clean = (ext || '').toLowerCase().replace(/^\./, '').trim();
  const list = EXTENSION_TARGETS[clean] || ['zip'];
  const excluded = ALIAS_GROUPS[clean] || [clean];
  return list
    .filter((t) => !excluded.includes(t))
    .map((target) => ({
      target,
      label: TARGET_LABELS[target] || target.toUpperCase()
    }));
}

