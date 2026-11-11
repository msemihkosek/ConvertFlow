import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import * as yaml from 'js-yaml';
import xml2js from 'xml2js';

/**
 * Universal Hub-and-Spoke Data Converter
 * Supports bidirectional conversions between:
 * JSON, CSV, TSV, XLSX, XLS, XML, YAML, SQL, Markdown Table, HTML Table
 */

// Step 1: Parse any input into a normalized array of objects (rows)
export async function parseToNormalizedData(inputPath, sourceExt) {
  const content = fs.readFileSync(inputPath);
  const ext = sourceExt.toLowerCase().replace('.', '');

  if (ext === 'json') {
    const parsed = JSON.parse(content.toString('utf-8'));
    return Array.isArray(parsed) ? parsed : [parsed];
  }

  if (ext === 'csv' || ext === 'tsv') {
    const delimiter = ext === 'tsv' ? '\t' : ',';
    const text = content.toString('utf-8');
    return parseDelimitedText(text, delimiter);
  }

  if (ext === 'yaml' || ext === 'yml') {
    const parsed = yaml.load(content.toString('utf-8'));
    return Array.isArray(parsed) ? parsed : [parsed];
  }

  if (ext === 'xml') {
    const parser = new xml2js.Parser({ explicitArray: false, mergeAttrs: true });
    const result = await parser.parseStringPromise(content.toString('utf-8'));
    // Find first array or wrap root object
    const rootKey = Object.keys(result)[0];
    const data = result[rootKey];
    if (Array.isArray(data)) return data;
    const innerKeys = Object.keys(data || {});
    for (const key of innerKeys) {
      if (Array.isArray(data[key])) return data[key];
    }
    return [data];
  }

  if (ext === 'xlsx' || ext === 'xls') {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(inputPath);
    const worksheet = workbook.worksheets[0];
    const rows = [];
    const headers = [];

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) {
        row.eachCell((cell, colNumber) => {
          headers[colNumber] = cell.value ? String(cell.value) : `Column_${colNumber}`;
        });
      } else {
        const rowObj = {};
        row.eachCell((cell, colNumber) => {
          const header = headers[colNumber] || `Column_${colNumber}`;
          rowObj[header] = cell.value !== undefined ? cell.value : '';
        });
        rows.push(rowObj);
      }
    });
    return rows;
  }

  throw new Error(`Desteklenmeyen veri kaynak formatı: ${ext}`);
}

// Step 2: Render normalized data into target format
export async function renderFromNormalizedData(rows, outputPath, targetExt, options = {}) {
  const ext = targetExt.toLowerCase().replace('.', '');

  if (!rows || rows.length === 0) {
    rows = [{ message: 'Boş veri' }];
  }

  // Collect all unique column headers
  const columns = Array.from(new Set(rows.flatMap(r => typeof r === 'object' && r !== null ? Object.keys(r) : [])));

  if (ext === 'json') {
    fs.writeFileSync(outputPath, JSON.stringify(rows, null, 2), 'utf-8');
    return outputPath;
  }

  if (ext === 'csv' || ext === 'tsv') {
    const delimiter = ext === 'tsv' ? '\t' : ',';
    let output = columns.map(escapeCsvValue).join(delimiter) + '\n';
    rows.forEach(row => {
      const line = columns.map(col => escapeCsvValue(row[col] !== undefined ? row[col] : '')).join(delimiter);
      output += line + '\n';
    });
    fs.writeFileSync(outputPath, output, 'utf-8');
    return outputPath;
  }

  if (ext === 'yaml' || ext === 'yml') {
    const yamlString = yaml.dump(rows, { indent: 2 });
    fs.writeFileSync(outputPath, yamlString, 'utf-8');
    return outputPath;
  }

  if (ext === 'xml') {
    const builder = new xml2js.Builder({ rootName: 'root' });
    const xml = builder.buildObject({ item: rows });
    fs.writeFileSync(outputPath, xml, 'utf-8');
    return outputPath;
  }

  if (ext === 'xlsx') {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Veri');
    sheet.columns = columns.map(c => ({ header: c, key: c, width: Math.max(c.length + 4, 15) }));
    
    // Style headers
    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } };

    rows.forEach(r => sheet.addRow(r));
    await workbook.xlsx.writeFile(outputPath);
    return outputPath;
  }

  if (ext === 'sql') {
    const tableName = options.tableName || 'tablo_verisi';
    let sql = `-- ConvertFlow Tarafından Oluşturulan SQL Tablosu\n`;
    sql += `CREATE TABLE IF NOT EXISTS ${tableName} (\n  ${columns.map(c => `\`${c}\` TEXT`).join(',\n  ')}\n);\n\n`;
    rows.forEach(row => {
      const vals = columns.map(c => {
        const val = row[c];
        if (val === null || val === undefined) return 'NULL';
        return `'${String(val).replace(/'/g, "''")}'`;
      }).join(', ');
      sql += `INSERT INTO ${tableName} (\`${columns.join('`, `')}\`) VALUES (${vals});\n`;
    });
    fs.writeFileSync(outputPath, sql, 'utf-8');
    return outputPath;
  }

  if (ext === 'html') {
    let html = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>table{border-collapse:collapse;width:100%;font-family:sans-serif;}th,td{border:1px solid #ddd;padding:8px;text-align:left;}th{background:#4f46e5;color:white;}tr:nth-child(even){background:#f9f9f9;}</style></head><body>`;
    html += `<table><thead><tr>${columns.map(c => `<th>${escapeHtml(c)}</th>`).join('')}</tr></thead><tbody>`;
    rows.forEach(row => {
      html += `<tr>${columns.map(c => `<td>${escapeHtml(row[c] !== undefined ? String(row[c]) : '')}</td>`).join('')}</tr>`;
    });
    html += `</tbody></table></body></html>`;
    fs.writeFileSync(outputPath, html, 'utf-8');
    return outputPath;
  }

  if (ext === 'md') {
    let md = `| ${columns.join(' | ')} |\n`;
    md += `| ${columns.map(() => '---').join(' | ')} |\n`;
    rows.forEach(row => {
      md += `| ${columns.map(c => String(row[c] !== undefined ? row[c] : '').replace(/\|/g, '\\|')).join(' | ')} |\n`;
    });
    fs.writeFileSync(outputPath, md, 'utf-8');
    return outputPath;
  }

  throw new Error(`Desteklenmeyen veri hedef formatı: ${ext}`);
}

// Helpers
function escapeCsvValue(val) {
  const str = String(val === undefined || val === null ? '' : val);
  if (str.includes(',') || str.includes('\n') || str.includes('"')) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  return str;
}

function escapeHtml(str) {
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function parseDelimitedText(text, delimiter) {
  const lines = text.split(/\r?\n/).filter(l => l.trim().length > 0);
  if (lines.length === 0) return [];
  const headers = lines[0].split(delimiter).map(h => h.trim().replace(/^"|"$/g, ''));
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const values = lines[i].split(delimiter).map(v => v.trim().replace(/^"|"$/g, ''));
    const obj = {};
    headers.forEach((h, idx) => {
      obj[h || `Col_${idx}`] = values[idx] !== undefined ? values[idx] : '';
    });
    rows.push(obj);
  }
  return rows;
}
