import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { fileURLToPath, pathToFileURL } from 'node:url';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';

const root = path.resolve(fileURLToPath(new URL('../data', import.meta.url)));
const templatesDir = path.join(root, 'templates');
const outputsDir = path.join(root, 'outputs');
const catalogFile = path.join(root, 'templates.json');
const execFileAsync = promisify(execFile);
const supportedExtensions = new Set(['.docx', '.xlsx', '.xls', '.pdf', '.txt', '.layout', '.json']);
const cleanName = name => String(name || 'template.docx').replace(/[^\p{L}\p{N}._ -]/gu, '_').slice(0, 120) || 'template.docx';
const extensionOf = name => path.extname(String(name || '')).toLowerCase();
const cleanLayoutName = name => String(name || '').replace(/^(?:🔵|🟡|🔴|🟢|⚪)\s*/u, '').trim();
const normalizeKey = value => cleanLayoutName(value).replace(/[\s_\-：:（）()]/g, '').toLowerCase();
const readCatalog = async () => { try { return JSON.parse(await fs.readFile(catalogFile, 'utf8')); } catch { return []; } };
const writeCatalog = value => fs.writeFile(catalogFile, JSON.stringify(value, null, 2), 'utf8');
const xmlEscape = value => String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const valueText = value => Array.isArray(value) ? value.map(valueText).join('、') : value && typeof value === 'object' ? (value.text || value.name || JSON.stringify(value)) : String(value ?? '');
const visibleXml = xml => xml.replace(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g, '$1').replace(/<[^>]+>/g, '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
const templateStats = (zip, extension) => {
  if (extension === '.xlsx') { const sheets = Object.keys(zip.files).filter(name => /^xl\/worksheets\/sheet\d+\.xml$/.test(name)); const xml = sheets.map(name => zip.file(name).asText()).join(''); return { worksheets: sheets.length, rows: (xml.match(/<row\b/g) || []).length, cells: (xml.match(/<c\b/g) || []).length, formulas: (xml.match(/<f\b/g) || []).length }; }
  if (extension === '.docx') { const xml = zip.file('word/document.xml')?.asText() || ''; return { worksheets: 1, rows: (xml.match(/<w:p\b/g) || []).length, cells: (xml.match(/<w:t\b/g) || []).length, formulas: 0 }; }
  return {};
};

const decodeLayout = bytes => {
  const raw = Buffer.from(bytes).toString('utf8').trim();
  let parsed;
  try { parsed = JSON.parse(raw); } catch { try { parsed = JSON.parse(Buffer.from(raw, 'base64').toString('utf8')); } catch { throw new Error('文件不是有效的飞书在线模板'); } }
  const content = typeof parsed.content === 'string' ? JSON.parse(parsed.content) : parsed.content || parsed;
  if (!content?.document?.pages?.length) throw new Error('文件不是有效的飞书在线模板');
  return { name: parsed.name || '', content };
};

const layoutFields = layout => {
  const found = [];
  const walk = value => {
    if (!value || typeof value !== 'object') return;
    if (value.type === 'variable' && Array.isArray(value.name)) for (const raw of value.name) {
      const name = String(raw).replace(/^(?:🔵|🟡|🔴|🟢|⚪)\s*/u, '').trim();
      if (name && name !== '#') found.push({ marker: name.startsWith('#') ? '#' : '', name: name.replace(/^#/, '') });
    }
    for (const child of Object.values(value)) walk(child);
  };
  walk(layout);
  const unique = new Map(); for (const field of found) if (!unique.has(field.name)) unique.set(field.name, field);
  return [...unique.values()];
};

const assertTemplate = (bytes, extension) => {
  if (bytes.length > 30 * 1024 * 1024) throw new Error('模板不能超过 30 MB');
  if (extension === '.docx') { const zip = new PizZip(bytes); if (!zip.file('word/document.xml')) throw new Error('文件不是有效的 DOCX 模板'); return zip; }
  if (extension === '.xlsx') { const zip = new PizZip(bytes); if (!Object.keys(zip.files).some(name => /^xl\/(workbook|worksheets\/sheet\d+)\.xml$/.test(name))) throw new Error('文件不是有效的 XLSX 模板'); return zip; }
  if (extension === '.txt' || extension === '.layout' || extension === '.json') { decodeLayout(bytes); return null; }
  if (!supportedExtensions.has(extension)) throw new Error('仅支持 DOCX、XLSX、XLS、PDF、飞书在线模板');
  if (!bytes.length) throw new Error('模板文件为空');
  return null;
};

export const extractFields = (zip, extension = '.docx') => {
  const names = extension === '.xlsx' ? Object.keys(zip.files).filter(name => /^xl\/(sharedStrings|worksheets\/sheet\d+)\.xml$/.test(name)) : Object.keys(zip.files).filter(name => /^word\/(document|header\d+|footer\d+)\.xml$/.test(name));
  const text = names.map(name => visibleXml(zip.file(name).asText())).join('');
  const delimiter = '{={{ }}=}'; const delimiterAt = text.indexOf(delimiter); const defaultText = delimiterAt < 0 ? text : text.slice(0, delimiterAt); const customText = delimiterAt < 0 ? '' : text.slice(delimiterAt + delimiter.length);
  const tags = [...defaultText.matchAll(/\{\s*(#!|[#/^!]?)\s*([^{}]+?)\s*\}/g), ...customText.matchAll(/\{\{\s*(#!|[#/^!]?)\s*([^{}]+?)\s*\}\}/g)];
  const fields = tags.map(match => { const marker = match[1] || ''; const raw = match[2].trim(); const name = raw.replace(/^!/, '').split(/\s*(?:==|!=|>=|<=|>|<|\|)\s*/)[0].trim(); return { marker: marker === '#!' || marker === '^' || /(?:==|!=|>=|<=|>|<)/.test(raw) ? '?' : marker, name }; }).filter(item => item.name && item.marker !== '/' && !item.name.startsWith('$') && !item.name.startsWith('%') && item.name !== '=' && !/[<>]/.test(item.name));
  const unique = new Map();
  for (const item of fields) if (!unique.has(item.name) || unique.get(item.name).marker === '/') unique.set(item.name, item);
  return [...unique.values()];
};

const withIndexes = value => Array.isArray(value) ? value.map((item, index) => item && typeof item === 'object' ? { ...Object.fromEntries(Object.entries(item).map(([key, child]) => [key, withIndexes(child)])), $index: index } : item) : value && typeof value === 'object' ? Object.fromEntries(Object.entries(value).map(([key, child]) => [key, withIndexes(child)])) : value;
const templateScope = (item, record) => {
  const source = {};
  for (const [key, value] of Object.entries(record || {})) { const indexed = withIndexes(value ?? ''); source[String(key)] = indexed; source[cleanLayoutName(key)] ??= indexed; }
  const normalized = Object.fromEntries(Object.entries(source).map(([key, value]) => [normalizeKey(key), value]));
  for (const field of item.fields || []) if (source[field.name] === undefined) source[field.name] = field.marker === '#' ? [] : normalized[normalizeKey(field.name)] ?? '';
  return source;
};

export async function listTemplates(scope = {}) {
  const catalog = await readCatalog(); let changed = false;
  for (const item of catalog) {
    const extension = item.extension || extensionOf(item.fileName);
    if (item.parserVersion === 2 && !item.fields?.some(field => /[<>]/.test(field.name)) && item.fields) continue;
    try { const bytes = await fs.readFile(path.join(templatesDir, item.fileName)); const zip = assertTemplate(bytes, extension); if (zip) { item.fields = extractFields(zip, extension); item.extension = extension; item.parserVersion = 2; changed = true; } else if (['.txt', '.layout', '.json'].includes(extension)) { item.fields = layoutFields(decodeLayout(bytes).content); item.extension = extension; item.parserVersion = 2; changed = true; } } catch {}
  }
  if (changed) await writeCatalog(catalog);
  const filtered = scope.tableId ? catalog.filter(item => item.tableId === scope.tableId && (!scope.baseId || item.baseId === scope.baseId)) : catalog;
  return filtered.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

export async function saveTemplate(originalName, bytes, binding = {}) {
  const name = cleanName(originalName); const extension = extensionOf(name);
  assertTemplate(bytes, extension);
  const id = crypto.randomUUID(); const fileName = `${id}${extension}`;
  await fs.mkdir(templatesDir, { recursive: true }); await fs.writeFile(path.join(templatesDir, fileName), bytes);
  const zip = extension === '.docx' || extension === '.xlsx' ? new PizZip(bytes) : null;
  const layout = ['.txt', '.layout', '.json'].includes(extension) ? decodeLayout(bytes).content : null;
  const item = { id, name, extension, fileName, fields: zip ? extractFields(zip, extension) : layout ? layoutFields(layout) : [], stats: zip ? templateStats(zip, extension) : layout ? { pages: layout.document.pages.length, blocks: JSON.stringify(layout).match(/"type":/g)?.length || 0 } : {}, parserVersion: 2, baseId: String(binding.baseId || ''), tableId: String(binding.tableId || ''), viewId: String(binding.viewId || ''), baseName: String(binding.baseName || ''), tableName: String(binding.tableName || ''), createdAt: new Date().toISOString() };
  const catalog = await readCatalog(); catalog.push(item); await writeCatalog(catalog); return item;
}

export async function removeTemplate(id) {
  const catalog = await readCatalog(); const item = catalog.find(entry => entry.id === id); if (!item) return false;
  await fs.rm(path.join(templatesDir, item.fileName), { force: true }); await writeCatalog(catalog.filter(entry => entry.id !== id)); return true;
}

export async function renameTemplate(id, requestedName) {
  const catalog = await readCatalog(); const item = catalog.find(entry => entry.id === id); if (!item) throw new Error('模板不存在');
  const extension = item.extension || extensionOf(item.fileName); const base = cleanName(requestedName).replace(/\.[^.]+$/, '').trim(); if (!base) throw new Error('模板名称不能为空');
  item.name = `${base}${extension}`; item.updatedAt = new Date().toISOString(); await writeCatalog(catalog); return item;
}

export async function previewTemplate(id) {
  const item = (await readCatalog()).find(entry => entry.id === id);
  if (!item) throw new Error('模板不存在');
  const extension = item.extension || extensionOf(item.fileName);
  const bytes = await fs.readFile(path.join(templatesDir, item.fileName));
  const layout = ['.txt', '.layout', '.json'].includes(extension) ? decodeLayout(bytes).content : null;
  if (layout) return { template: item, kind: 'layout', layout, fields: item.fields || layoutFields(layout), fileUrl: `/api/templates/${item.id}/file` };
  const zip = extension === '.docx' || extension === '.xlsx' ? assertTemplate(bytes, extension) : null;
  if (!zip) return { template: item, kind: extension === '.pdf' ? 'pdf' : 'binary', fileUrl: `/api/templates/${item.id}/file` };
  const names = extension === '.docx' ? Object.keys(zip.files).filter(name => /^word\/(document|header\d+|footer\d+)\.xml$/.test(name)) : Object.keys(zip.files).filter(name => /^xl\/worksheets\/sheet\d+\.xml$/.test(name));
  const content = names.map(name => visibleXml(zip.file(name).asText())).join('\n').replace(/\s+/g, ' ').trim().slice(0, 20000);
  return { template: item, kind: extension.slice(1), fields: item.fields || [], content, fileUrl: `/api/templates/${item.id}/file` };
}

const resolveTemplateValue = (scope, key) => { let value = scope; for (const part of String(key).trim().split(/\s*[./]\s*/)) value = matchingValue(value, part); return value; };
const conditionLiteral = raw => { const value = String(raw).trim(); if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) return value.slice(1, -1); if (/^-?\d+(?:\.\d+)?$/.test(value)) return Number(value); if (value === 'true') return true; if (value === 'false') return false; if (value === 'null') return null; return value; };
const templateParser = tag => ({ get(scope) { const index = String(tag).match(/^\$index\s*\+\s*(\d+)$/); if (index) return Number(scope?.$index || 0) + Number(index[1]); const condition = String(tag).match(/^(.+?)\s*(==|!=|>=|<=|>|<)\s*(.+)$/); if (!condition) return resolveTemplateValue(scope, tag); const left = resolveTemplateValue(scope, condition[1]); const right = conditionLiteral(condition[3]); if (condition[2] === '==') return String(left ?? '') === String(right ?? ''); if (condition[2] === '!=') return String(left ?? '') !== String(right ?? ''); if (condition[2] === '>') return Number(left) > Number(right); if (condition[2] === '<') return Number(left) < Number(right); if (condition[2] === '>=') return Number(left) >= Number(right); return Number(left) <= Number(right); } });
const renderDocx = (bytes, record) => { const zip = assertTemplate(bytes, '.docx'); for (const name of Object.keys(zip.files).filter(file => /^word\/.*\.xml$/.test(file))) zip.file(name, zip.file(name).asText().replace(/\{#!\s*([^{}]+?)\}/g, '{^$1}')); const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, parser: templateParser }); doc.render(record); return doc.getZip().generate({ type: 'nodebuffer' }); };
const xmlUnescape = value => String(value || '').replace(/&#x([0-9a-f]+);/gi, (_, code) => String.fromCodePoint(parseInt(code, 16))).replace(/&#(\d+);/g, (_, code) => String.fromCodePoint(Number(code))).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&');
const spreadsheetText = value => xmlUnescape([...String(value || '').matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(match => match[1]).join(''));
const renderTemplateText = (template, scope, loopName = '') => {
  let output = String(template || ''); const escapedLoop = loopName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  if (loopName) output = output.replace(new RegExp(`\\{#\\s*${escapedLoop}\\s*\\}`, 'g'), '').replace(new RegExp(`\\{/\\s*${escapedLoop}\\s*\\}`, 'g'), '');
  output = output.replace(/\{#(?!\!)([^{}]+)\}([\s\S]*?)\{\/\s*\1\s*\}/g, (_, condition, body) => templateParser(condition).get(scope) ? body : '').replace(/\{#([^{}]+)\}([\s\S]*?)\{\/\}/g, (_, condition, body) => templateParser(condition).get(scope) ? body : '').replace(/\{#!([^{}]+)\}([\s\S]*?)\{\/[^{}]*\}/g, (_, condition, body) => templateParser(condition).get(scope) ? '' : body);
  if (output.includes('{={{ }}=}')) return output.replace('{={{ }}=}', '').replace(/\{\{\s*([^{}]+?)\s*\}\}/g, (_, key) => valueText(templateParser(key).get(scope)));
  return output.replace(/\{\s*([^{}]+?)\s*\}/g, (_, key) => valueText(templateParser(key).get(scope)));
};
const renderXlsx = (bytes, record, item) => {
  const zip = assertTemplate(bytes, '.xlsx'); const scope = templateScope(item, record); const shared = [...(zip.file('xl/sharedStrings.xml')?.asText() || '').matchAll(/<si>([\s\S]*?)<\/si>/g)].map(match => spreadsheetText(match[1]));
  const cellSource = cell => { const sharedIndex = cell.match(/\bt="s"[^>]*>[\s\S]*?<v>(\d+)<\/v>/)?.[1]; if (sharedIndex !== undefined) return shared[Number(sharedIndex)] || ''; return spreadsheetText(cell); };
  const renderRow = (rowXml, sourceRow, targetRow, rowScope, loopName = '') => rowXml.replace(/(<row\b[^>]*\br=")\d+("[^>]*>)/, `$1${targetRow}$2`).replace(/<c\b([^>]*)>([\s\S]*?)<\/c>/g, (cell, attributes, inner) => { const nextAttributes = attributes.replace(/\br="([A-Z]+)\d+"/, `r="$1${targetRow}"`); const source = cellSource(cell); if (!source.includes('{')) { const delta = targetRow - sourceRow; const shifted = delta ? inner.replace(/(<f[^>]*>)([\s\S]*?)(<\/f>)/g, (_, open, formula, close) => `${open}${formula.replace(/(\$?[A-Z]{1,3})(\$?)(\d+)/g, (ref, column, fixed, row) => fixed ? ref : `${column}${Number(row) + delta}`)}${close}`) : inner; return `<c${nextAttributes}>${shifted}</c>`; } const rendered = renderTemplateText(source, rowScope, loopName); const inlineAttributes = nextAttributes.replace(/\s+t="[^"]*"/g, '') + ' t="inlineStr"'; return `<c${inlineAttributes}><is><t xml:space="preserve">${xmlEscape(rendered)}</t></is></c>`; });
  for (const name of Object.keys(zip.files).filter(file => /^xl\/worksheets\/sheet\d+\.xml$/.test(file))) {
    let shift = 0; const expansions = []; let xml = zip.file(name).asText();
    xml = xml.replace(/<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>/g, rowXml => { const sourceRow = Number(rowXml.match(/\br="(\d+)"/)?.[1]); const texts = [...rowXml.matchAll(/<c\b[^>]*>[\s\S]*?<\/c>/g)].map(match => cellSource(match[0])); const loop = [...texts.join('').matchAll(/\{#(?!\!)([^{}]+)\}/g)].map(match => match[1].trim()).find(name => Array.isArray(matchingValue(scope, name))); const rows = loop ? matchingValue(scope, loop) : null; const targetRow = sourceRow + shift; if (!rows) return renderRow(rowXml, sourceRow, targetRow, scope); expansions.push({ row: sourceRow, count: rows.length }); shift += rows.length - 1; return rows.map((child, index) => renderRow(rowXml, sourceRow, targetRow + index, { ...scope, ...(child && typeof child === 'object' ? child : {}), $index: index }, loop)).join(''); });
    const shiftRow = row => Number(row) + expansions.filter(item => item.row < Number(row)).reduce((total, item) => total + item.count - 1, 0); xml = xml.replace(/(<dimension\b[^>]*\bref="[A-Z]+\d+:[A-Z]+)(\d+)(")/, (_, start, row, end) => `${start}${shiftRow(row)}${end}`); zip.file(name, xml);
  }
  return zip.generate({ type: 'nodebuffer' });
};
const docxText = xml => xmlUnescape([...String(xml || '').matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map(match => match[1]).join(''));
const docxPreviewHtml = bytes => {
  const zip = assertTemplate(bytes, '.docx'); const xml = zip.file('word/document.xml').asText(); const body = xml.match(/<w:body[^>]*>([\s\S]*?)<w:sectPr[\s\S]*?<\/w:body>/)?.[1] || xml;
  const renderParagraph = value => `<p>${xmlEscape(docxText(value)).replace(/\n/g, '<br>')}</p>`;
  const renderTable = value => `<table>${[...value.matchAll(/<w:tr[\s\S]*?<\/w:tr>/g)].map(row => `<tr>${[...row[0].matchAll(/<w:tc[\s\S]*?<\/w:tc>/g)].map(cell => `<td>${[...cell[0].matchAll(/<w:p[\s\S]*?<\/w:p>/g)].map(paragraph => renderParagraph(paragraph[0])).join('')}</td>`).join('')}</tr>`).join('')}</table>`;
  const parts = []; let cursor = 0; const blocks = [...body.matchAll(/<w:tbl[\s\S]*?<\/w:tbl>|<w:p[\s\S]*?<\/w:p>/g)]; for (const block of blocks) { if (block.index < cursor) continue; parts.push(block[0].startsWith('<w:tbl') ? renderTable(block[0]) : renderParagraph(block[0])); cursor = block.index + block[0].length; }
  return `<!doctype html><html><head><meta charset="utf-8"><style>body{margin:0;padding:24px;background:#eef1f5;color:#111;font:14px/1.7 "Microsoft YaHei",sans-serif}.paper{max-width:820px;margin:auto;padding:34px 42px;background:#fff;box-shadow:0 2px 12px #0002}p{margin:0 0 9px;white-space:normal}table{width:100%;border-collapse:collapse;margin:12px 0}td{border:1px solid #667085;padding:6px 8px;vertical-align:top}br{line-height:1.7}</style></head><body><main class="paper">${parts.join('')}</main></body></html>`;
};
const xlsxPreviewHtml = bytes => {
  const zip = assertTemplate(bytes, '.xlsx'); const shared = [...(zip.file('xl/sharedStrings.xml')?.asText() || '').matchAll(/<si>([\s\S]*?)<\/si>/g)].map(match => spreadsheetText(match[1])); const sheet = zip.file('xl/worksheets/sheet1.xml')?.asText() || ''; const rows = [...sheet.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)].map(row => `<tr>${[...row[1].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)].map(cell => { const sharedIndex = cell[0].match(/\bt="s"[^>]*>[\s\S]*?<v>(\d+)<\/v>/)?.[1]; const value = sharedIndex === undefined ? spreadsheetText(cell[2]) : shared[Number(sharedIndex)] || ''; return `<td>${xmlEscape(value).replace(/\n/g, '<br>')}</td>`; }).join('')}</tr>`).join(''); return `<!doctype html><html><head><meta charset="utf-8"><style>body{margin:0;padding:24px;background:#eef1f5;font:13px/1.5 "Microsoft YaHei",sans-serif}.paper{max-width:1200px;margin:auto;padding:24px;background:#fff;box-shadow:0 2px 12px #0002}table{width:100%;border-collapse:collapse}td{border:1px solid #cbd5e1;padding:6px 8px;vertical-align:top;white-space:pre-wrap}</style></head><body><main class="paper"><table>${rows}</table></main></body></html>`;
};
export async function previewRecord(id, record) { const item = (await readCatalog()).find(entry => entry.id === id); if (!item) throw new Error('模板不存在'); const extension = item.extension || extensionOf(item.fileName); const bytes = await fs.readFile(path.join(templatesDir, item.fileName)); if (extension === '.docx') return { template: item, kind: 'html', html: docxPreviewHtml(renderDocx(bytes, templateScope(item, record))) }; if (extension === '.xlsx') return { template: item, kind: 'html', html: xlsxPreviewHtml(renderXlsx(bytes, record, item)) }; throw new Error('该模板格式暂不支持记录预览'); }
const matchingValue = (object, key) => { if (!object || typeof object !== 'object') return undefined; if (Object.hasOwn(object, key)) return object[key]; const normalized = normalizeKey(key); const match = Object.keys(object).find(candidate => normalizeKey(candidate) === normalized); return match === undefined ? undefined : object[match]; };
const layoutValue = (names, scope, rowContext) => {
  const path = (names || []).map(cleanLayoutName);
  if (path[0] === '#') return String((rowContext?.index ?? 0) + 1);
  if (rowContext && normalizeKey(path[0]) === normalizeKey(rowContext.root)) return valueText(matchingValue(rowContext.item, path.at(-1)));
  let value = matchingValue(scope, path[0]);
  for (const key of path.slice(1)) value = matchingValue(value, key);
  return valueText(value);
};
const layoutInline = (nodes, scope, rowContext) => (nodes || []).map(node => node.type === 'variable' ? xmlEscape(layoutValue(node.name, scope, rowContext)) : xmlEscape(node.text || '')).join('');
const renderLayoutHtml = (bytes, record) => {
  const layout = decodeLayout(bytes).content; const scope = templateScope({ fields: layoutFields(layout) }, record); const setting = layout.pageSetting || {};
  const renderParagraphs = (nodes, rowContext) => (nodes || []).map(node => `<p style="text-align:${node.align || 'left'};line-height:${node.lineHeight || layout.settings?.defaultLineHeight || 1.5};margin:0 0 ${node.marginBottom || 0};white-space:pre-wrap">${layoutInline(node.children, scope, rowContext)}</p>`).join('');
  const renderTable = table => {
    const dynamic = new Map((table.dynamicRows || []).map(item => [Number(item.rowIndex), cleanLayoutName(item.dataSource?.rootPath?.[0])]));
    const merges = table.merges || [];
    const rowHtml = (row, rowIndex, rowContext) => {
      const covered = new Set();
      for (const merge of merges) for (let rr = merge.rowIndex; rr < merge.rowIndex + merge.rowSpan; rr++) for (let cc = merge.colIndex; cc < merge.colIndex + merge.colSpan; cc++) if (rr !== merge.rowIndex || cc !== merge.colIndex) covered.add(`${rr}:${cc}`);
      const cells = (row.cells || []).map((cell, cellIndex) => {
        if (covered.has(`${rowIndex}:${cellIndex}`)) return '';
        const merge = merges.find(item => item.rowIndex === rowIndex && item.colIndex === cellIndex);
        const attrs = [merge?.colSpan > 1 ? `colspan="${merge.colSpan}"` : '', merge?.rowSpan > 1 ? `rowspan="${merge.rowSpan}"` : ''].filter(Boolean).join(' ');
        const html = `<td${attrs ? ` ${attrs}` : ''} style="background:${cell.background || ''};vertical-align:${cell.verticalAlign || 'top'}">${renderParagraphs(cell.content, rowContext)}</td>`;
        return html;
      }).join('');
      return `<tr>${cells}</tr>`;
    };
    const rows = (table.rows || []).flatMap((row, index) => { const rootName = dynamic.get(index); const rootValue = matchingValue(scope, rootName); const values = rootName && Array.isArray(rootValue) ? rootValue : null; return values?.length ? values.map((item, rowIndex) => rowHtml(row, index, { root: rootName, item, index: rowIndex })) : [rowHtml(row, index)]; }).join('');
    const columns = (table.columns || []).map(column => `<col style="width:${Number(column.width || 0)}%">`).join('');
    return `<table><colgroup>${columns}</colgroup><tbody>${rows}</tbody></table>`;
  };
  const renderBlock = block => Number(block.type) === 4 ? renderTable(block.table || {}) : `<div>${renderParagraphs(block.content)}</div>`;
  const pages = (layout.document.pages || []).map(page => `<section class="page">${(page.rows || []).map(row => `<div class="row">${(row.columns || []).map(column => `<div class="column" style="width:${Number(column.width || 100)}%">${(column.blocks || []).map(renderBlock).join('')}</div>`).join('')}</div>`).join('')}</section>`).join('');
  return Buffer.from(`<!doctype html><html><head><meta charset="utf-8"><style>@page{size:${Number(setting.width || 210)}mm ${Number(setting.height || 297)}mm;margin:0}*{box-sizing:border-box}html,body{margin:0;color:#111;font-family:"Microsoft YaHei","SimSun",sans-serif;font-size:${Number(layout.settings?.fontSize || 9)}pt}.page{width:${Number(setting.width || 210)}mm;min-height:${Number(setting.height || 297)}mm;padding:${Number(setting.paddingTop || 5)}mm ${Number(setting.paddingRight || 13.5)}mm ${Number(setting.paddingBottom || 5)}mm ${Number(setting.paddingLeft || 13.5)}mm;page-break-after:always;background:#fff}.page:last-child{page-break-after:auto}.row{display:flex;width:100%}.column{min-width:0}table{width:100%;border-collapse:collapse;table-layout:fixed}td{border:1px solid #222;padding:2px 3px;overflow-wrap:anywhere}tr{break-inside:avoid}p{min-height:1em}</style></head><body>${pages}</body></html>`, 'utf8');
};

const chromePath = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const htmlToPdf = async (html, outputPath) => {
  const htmlPath = `${outputPath}.html`; const profilePath = `${outputPath}-chrome`; await fs.writeFile(htmlPath, html);
  try { await execFileAsync(chromePath, ['--headless=new', '--no-sandbox', '--disable-gpu', '--disable-software-rasterizer', `--user-data-dir=${profilePath}`, '--no-pdf-header-footer', `--print-to-pdf=${outputPath}`, pathToFileURL(htmlPath).href], { timeout: 120000, windowsHide: true }); }
  finally { await fs.rm(htmlPath, { force: true }); await fs.rm(profilePath, { recursive: true, force: true }); }
  return fs.readFile(outputPath);
};

const layoutToWord = html => Buffer.from(`\ufeff${html.toString('utf8')}`, 'utf16le');
const layoutToXlsx = record => {
  const zip = new PizZip(); const entries = Object.entries(record || {}).filter(([, value]) => !Array.isArray(value));
  const rows = [entries.map(([name]) => name), entries.map(([, value]) => valueText(value))]; const cell = (value, column, row) => `<c r="${String.fromCharCode(65 + column)}${row}" t="inlineStr"><is><t>${xmlEscape(value)}</t></is></c>`;
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="数据" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows.map((row, index) => `<row r="${index + 1}">${row.map((value, column) => cell(value, column, index + 1)).join('')}</row>`).join('')}</sheetData></worksheet>`);
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
};

export async function renderTemplate(id, records, requestedFormat = '') {
  const item = (await listTemplates()).find(entry => entry.id === id); if (!item) throw new Error('模板不存在');
  if (!Array.isArray(records) || records.length === 0) throw new Error('至少需要一条记录');
  const bytes = await fs.readFile(path.join(templatesDir, item.fileName)); const extension = item.extension || extensionOf(item.fileName); const outputId = crypto.randomUUID(); await fs.mkdir(outputsDir, { recursive: true });
  const isLayout = ['.txt', '.layout', '.json'].includes(extension); const requested = String(requestedFormat || '').toLowerCase(); const defaultFormat = isLayout ? 'pdf' : extension.slice(1); const format = requested === 'word' ? 'doc' : requested || defaultFormat;
  if (!isLayout && format !== extension.slice(1)) throw new Error(`当前 ${extension.slice(1).toUpperCase()} 模板仅支持生成同格式文件`);
  if (isLayout && !['pdf', 'doc', 'xlsx'].includes(format)) throw new Error('在线模板支持生成 PDF、Word 或 Excel');
  const render = async (record, index) => {
    if (extension === '.docx') return renderDocx(bytes, templateScope(item, record));
    if (extension === '.xlsx') return renderXlsx(bytes, record, item);
    if (!isLayout) return bytes;
    const html = renderLayoutHtml(bytes, record);
    if (format === 'pdf') return htmlToPdf(html, path.join(outputsDir, `${outputId}-${index}.pdf`));
    if (format === 'doc') return layoutToWord(html);
    return layoutToXlsx(record);
  };
  const outputExtension = `.${format}`;
  if (records.length === 1) { const fileName = `${outputId}${outputExtension}`; const generated = await render(records[0], 0); await fs.writeFile(path.join(outputsDir, fileName), generated); if (format === 'pdf') await fs.rm(path.join(outputsDir, `${outputId}-0.pdf`), { force: true }); return { id: outputId, fileName, extension: format, name: cleanName(`${path.parse(item.name).name}-${new Date().toISOString().slice(0, 10)}${outputExtension}`) }; }
  const archive = new PizZip(); for (const [index, record] of records.entries()) { archive.file(`${String(index + 1).padStart(3, '0')}${outputExtension}`, await render(record, index)); if (format === 'pdf') await fs.rm(path.join(outputsDir, `${outputId}-${index}.pdf`), { force: true }); } const fileName = `${outputId}.zip`; await fs.writeFile(path.join(outputsDir, fileName), archive.generate({ type: 'nodebuffer', compression: 'DEFLATE' })); return { id: outputId, fileName, extension: 'zip', name: cleanName(`${path.parse(item.name).name}-${records.length}份.zip`) };
}

export const templateFile = (id, extension = '.docx') => path.join(templatesDir, `${id}${extension.startsWith('.') ? extension : `.${extension}`}`);
export const outputFile = (id, extension = 'docx') => path.join(outputsDir, `${id}.${extension}`);
