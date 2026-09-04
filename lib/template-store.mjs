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
const normalizeKey = value => cleanLayoutName(value).replace(/[\u2600-\u27BF\u{1F300}-\u{1FAFF}]/gu, '').replace(/[\s_\-：:（）()·]/g, '').toLowerCase();
const legacyPlaceholderName = value => String(value || '').replace(/[^A-Za-z0-9_\u3400-\u9fff]/g, '_');
const readCatalog = async () => { try { return JSON.parse(await fs.readFile(catalogFile, 'utf8')); } catch { return []; } };
const writeCatalog = value => fs.writeFile(catalogFile, JSON.stringify(value, null, 2), 'utf8');
const xmlEscape = value => String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const valueText = value => Array.isArray(value) ? value.map(valueText).join('、') : value && typeof value === 'object' ? (value.text || value.name || (value.value !== undefined ? valueText(value.value) : JSON.stringify(value))) : String(value ?? '');
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
const procurementAliases = { '__合同编号': ['合同编号'], '__供应商': ['供应商'], '__合同明细_采购明细': ['合同明细', '采购明细', '合同明细采购明细'], '__开票品名': ['开票品名'], 'SKU__': ['SKU'], '单价_含税___': ['单价（含税）', '单价(含税)', '含税单价'], '实收数量__': ['实收数量', '数量'], '总价_含税_': ['总价（含税）', '总价(含税)', '含税总价'], '__产品金额': ['产品金额'], '__产品运输费用': ['产品运输费用'], '__合同合计金额': ['合同合计金额'], '__采购合同': ['采购合同'], '__SKU种类': ['SKU种类'], '__SKU总数': ['SKU总数'] };
const fieldValue = (object, name) => {
  if (!object || typeof object !== 'object') return undefined;
  const names = [name, ...(procurementAliases[name] || [])];
  const entries = Object.entries(object).filter(([, value]) => value != null && value !== '');
  for (const candidate of names) {
    const exact = entries.find(([key]) => key === candidate) || entries.find(([key]) => legacyPlaceholderName(key) === candidate);
    if (exact) return exact[1];
  }
  for (const candidate of names) {
    const normalized = entries.find(([key]) => normalizeKey(key) === normalizeKey(candidate));
    if (normalized) return normalized[1];
  }
  return undefined;
};
const numberValue = value => { if (value && typeof value === 'object' && value.value !== undefined) return numberValue(value.value); const number = Number(String(value ?? '').replace(/[,，￥¥\s]/g, '')); return Number.isFinite(number) ? number : 0; };
const chineseCurrency = value => { const amount = Math.round(numberValue(value) * 100); const digits = '零壹贰叁肆伍陆柒捌玖'; const small = ['', '拾', '佰', '仟']; const large = ['', '万', '亿', '兆']; const groupText = group => { let out = ''; let zero = false; for (let i = 3; i >= 0; i--) { const digit = Math.floor(group / (10 ** i)) % 10; if (!digit) { if (out) zero = true; continue; } if (zero) out += '零'; out += digits[digit] + small[i]; zero = false; } return out; }; let out = ''; let missing = false; const integer = Math.floor(amount / 100); for (let i = 3; i >= 0; i--) { const group = Math.floor(integer / (10000 ** i)) % 10000; if (!group) { if (out) missing = true; continue; } if (out && (missing || group < 1000)) out += '零'; out += groupText(group) + large[i]; missing = false; } if (!out) out = '零'; const jiao = Math.floor(amount / 10) % 10; const fen = amount % 10; return `${numberValue(value) < 0 ? '负' : ''}${out}元${jiao ? `${digits[jiao]}角` : ''}${fen ? `${!jiao && integer ? '零' : ''}${digits[fen]}分` : (!jiao ? '整' : '')}`; };
const templateScope = (item, record) => {
  const source = {};
  for (const [key, value] of Object.entries(record || {})) { const indexed = withIndexes(value ?? ''); source[String(key)] = indexed; source[cleanLayoutName(key)] ??= indexed; }
  const loopName = Object.keys(source).find(key => /合同明细.*采购明细|采购明细.*合同明细/.test(normalizeKey(key))) || Object.keys(source).find(key => /采购明细|合同明细/.test(normalizeKey(key)) && Array.isArray(source[key]));
  const rows = loopName && Array.isArray(source[loopName]) ? source[loopName] : [];
  if (rows.length) { source['__合同明细_采购明细'] = rows; for (const row of rows) for (const name of ['__开票品名', 'SKU__', '单位', '单价_含税___', '实收数量__', '总价_含税_']) { const value = fieldValue(row, name); if (value !== undefined) row[name] = value; } source['__SKU总数'] = rows.reduce((sum, row) => sum + numberValue(fieldValue(row, '实收数量__')), 0); }
  const contractTotal = fieldValue(source, '__合同合计金额'); if (contractTotal !== undefined) source['合同合计金额_大写_'] = chineseCurrency(contractTotal);
  const productTotal = fieldValue(source, '__产品金额'); if (productTotal !== undefined) source['产品金额_大写_'] = chineseCurrency(productTotal);
  const freightTotal = fieldValue(source, '__产品运输费用'); if (freightTotal !== undefined) source['产品运输费用_大写_'] = chineseCurrency(freightTotal);
  source['__合同创建时间'] = `${new Date().getFullYear()}年${new Date().getMonth() + 1}月${new Date().getDate()}日`;
  const normalized = Object.fromEntries(Object.entries(source).map(([key, value]) => [normalizeKey(key), value]));
  for (const field of item.fields || []) if (source[field.name] === undefined) source[field.name] = field.marker === '#' ? [] : fieldValue(source, field.name) ?? normalized[normalizeKey(field.name)] ?? '';
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

export async function previewTemplate(id) {
  const item = (await readCatalog()).find(entry => entry.id === id);
  if (!item) throw new Error('模板不存在');
  const extension = item.extension || extensionOf(item.fileName);
  const bytes = await fs.readFile(path.join(templatesDir, item.fileName));
  const layout = ['.txt', '.layout', '.json'].includes(extension) ? decodeLayout(bytes).content : null;
  if (layout) return { template: item, kind: 'layout', layout, fields: item.fields || layoutFields(layout), fileUrl: `/api/templates/${item.id}/file` };
  const zip = extension === '.docx' || extension === '.xlsx' ? assertTemplate(bytes, extension) : null;
  if (!zip) return { template: item, kind: extension === '.pdf' ? 'pdf' : 'binary', fileUrl: `/api/templates/${item.id}/file` };
  if (extension === '.xlsx') return { template: item, kind: 'html', html: xlsxPreviewHtml(bytes), fields: item.fields || [], fileUrl: `/api/templates/${item.id}/file` };
  const names = extension === '.docx' ? Object.keys(zip.files).filter(name => /^word\/(document|header\d+|footer\d+)\.xml$/.test(name)) : Object.keys(zip.files).filter(name => /^xl\/worksheets\/sheet\d+\.xml$/.test(name));
  const content = names.map(name => visibleXml(zip.file(name).asText())).join('\n').replace(/\s+/g, ' ').trim().slice(0, 20000);
  return { template: item, kind: extension.slice(1), fields: item.fields || [], content, fileUrl: `/api/templates/${item.id}/file` };
}

const resolveTemplateValue = (scope, key) => { let value = scope; for (const part of String(key).trim().split(/\s*[./]\s*/)) value = matchingValue(value, part); return value; };
const conditionLiteral = raw => { const value = String(raw).trim(); if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) return value.slice(1, -1); if (/^-?\d+(?:\.\d+)?$/.test(value)) return Number(value); if (value === 'true') return true; if (value === 'false') return false; if (value === 'null') return null; return value; };
const templateParser = tag => ({ get(scope) { const index = String(tag).match(/^\$index\s*\+\s*(\d+)$/); if (index) return Number(scope?.$index || 0) + Number(index[1]); const condition = String(tag).match(/^(.+?)\s*(==|!=|>=|<=|>|<)\s*(.+)$/); if (!condition) { const value = resolveTemplateValue(scope, tag); return value && typeof value === 'object' && !Array.isArray(value) ? valueText(value) : value; } const rawLeft = resolveTemplateValue(scope, condition[1]); const left = rawLeft && typeof rawLeft === 'object' ? valueText(rawLeft) : rawLeft; const right = conditionLiteral(condition[3]); if (condition[2] === '==') return String(left ?? '') === String(right ?? ''); if (condition[2] === '!=') return String(left ?? '') !== String(right ?? ''); if (condition[2] === '>') return Number(left) > Number(right); if (condition[2] === '<') return Number(left) < Number(right); if (condition[2] === '>=') return Number(left) >= Number(right); return Number(left) <= Number(right); } });
const renderDocx = (bytes, record) => { const zip = assertTemplate(bytes, '.docx'); for (const name of Object.keys(zip.files).filter(file => /^word\/.*\.xml$/.test(file))) zip.file(name, zip.file(name).asText().replace(/\{#!\s*([^{}]+?)\}/g, '{^$1}')); const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, parser: templateParser }); doc.render(record); return doc.getZip().generate({ type: 'nodebuffer' }); };
const xmlUnescape = value => String(value || '').replace(/&#x([0-9a-f]+);/gi, (_, code) => String.fromCodePoint(parseInt(code, 16))).replace(/&#(\d+);/g, (_, code) => String.fromCodePoint(Number(code))).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&');
const normalizeSpreadsheetXml = value => { const xml = String(value || ''); const prefix = xml.match(/<([A-Za-z_][\w.-]*):(worksheet|sst|styleSheet|workbook)\b/)?.[1]; if (!prefix) return xml; const escaped = prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); return xml.replace(new RegExp(`<(/?)${escaped}:`, 'g'), '<$1').replace(new RegExp(`xmlns:${escaped}=`), 'xmlns='); };
const spreadsheetText = value => xmlUnescape([...String(value || '').matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(match => match[1]).join(''));
const renderTemplateText = (template, scope, loopName = '') => {
  let output = String(template || ''); const escapedLoop = loopName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  if (loopName) output = output.replace(new RegExp(`\\{#\\s*${escapedLoop}\\s*\\}`, 'g'), '').replace(new RegExp(`\\{/\\s*${escapedLoop}\\s*\\}`, 'g'), '');
  output = output.replace(/\{#(?!\!)([^{}]+)\}([\s\S]*?)\{\/\s*\1\s*\}/g, (_, condition, body) => templateParser(condition).get(scope) ? body : '').replace(/\{#([^{}]+)\}([\s\S]*?)\{\/\}/g, (_, condition, body) => templateParser(condition).get(scope) ? body : '').replace(/\{#!([^{}]+)\}([\s\S]*?)\{\/[^{}]*\}/g, (_, condition, body) => templateParser(condition).get(scope) ? '' : body);
  if (output.includes('{={{ }}=}')) return output.replace('{={{ }}=}', '').replace(/\{\{\s*([^{}]+?)\s*\}\}/g, (_, key) => valueText(templateParser(key).get(scope)));
  return output.replace(/\{\s*([^{}]+?)\s*\}/g, (_, key) => valueText(templateParser(key).get(scope)));
};
const renderXlsx = (bytes, record, item) => {
  const zip = assertTemplate(bytes, '.xlsx'); const scope = templateScope(item, record); const shared = [...normalizeSpreadsheetXml(zip.file('xl/sharedStrings.xml')?.asText() || '').matchAll(/<si>([\s\S]*?)<\/si>/g)].map(match => spreadsheetText(match[1]));
  const cellSource = cell => { const sharedIndex = cell.match(/\bt="s"[^>]*>[\s\S]*?<v>(\d+)<\/v>/)?.[1]; if (sharedIndex !== undefined) return shared[Number(sharedIndex)] || ''; const inline = spreadsheetText(cell); return inline || xmlUnescape(cell.match(/<v>([\s\S]*?)<\/v>/)?.[1] || ''); };
  const renderRow = (rowXml, sourceRow, targetRow, rowScope, loopName = '') => rowXml.replace(/(<row\b[^>]*\br=")\d+("[^>]*>)/, `$1${targetRow}$2`).replace(/<c\b([^>]*)>([\s\S]*?)<\/c>/g, (cell, attributes, inner) => { const nextAttributes = attributes.replace(/\br="([A-Z]+)\d+"/, `r="$1${targetRow}"`); const source = cellSource(cell); if (!source.includes('{')) { const delta = targetRow - sourceRow; const shifted = delta ? inner.replace(/(<f[^>]*>)([\s\S]*?)(<\/f>)/g, (_, open, formula, close) => `${open}${formula.replace(/(\$?[A-Z]{1,3})(\$?)(\d+)/g, (ref, column, fixed, row) => fixed ? ref : `${column}${Number(row) + delta}`)}${close}`) : inner; return `<c${nextAttributes}>${shifted}</c>`; } const rendered = renderTemplateText(source, rowScope, loopName); const inlineAttributes = nextAttributes.replace(/\s+t="[^"]*"/g, '') + ' t="inlineStr"'; return `<c${inlineAttributes}><is><t xml:space="preserve">${xmlEscape(rendered)}</t></is></c>`; });
  for (const name of Object.keys(zip.files).filter(file => /^xl\/worksheets\/sheet\d+\.xml$/.test(file))) {
    let shift = 0; const expansions = []; let xml = normalizeSpreadsheetXml(zip.file(name).asText());
    const rowMatches = [...xml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>/g)]; const rows = rowMatches.map(match => ({ xml: match[0], sourceRow: Number(match[1]), start: match.index, end: match.index + match[0].length })); const blocks = new Map();
    for (let index = 0; index < rows.length; index++) { const texts = [...rows[index].xml.matchAll(/<c\b[^>]*>[\s\S]*?<\/c>/g)].map(match => cellSource(match[0])).join(''); const loop = [...texts.matchAll(/\{#(?!\!)([^{}]+)\}/g)].map(match => match[1].trim()).find(loopName => Array.isArray(matchingValue(scope, loopName))); if (!loop) continue; const close = new RegExp(`\\{\\/\\s*${loop.replace(/[.*+?^${}()|[\]\\\\]/g, '\\\\$&')}\\s*\\}`, 'i'); const end = rows.findIndex((candidate, candidateIndex) => candidateIndex >= index && close.test([...candidate.xml.matchAll(/<c\b[^>]*>[\s\S]*?<\/c>/g)].map(match => cellSource(match[0])).join(''))); if (end >= index) blocks.set(index, { end, loop }); }
    let output = ''; let cursor = 0;
    for (let index = 0; index < rows.length; index++) { const current = rows[index]; output += xml.slice(cursor, current.start); const block = blocks.get(index); if (block) { const values = matchingValue(scope, block.loop) || []; const blockRows = rows.slice(index, block.end + 1); expansions.push({ row: current.sourceRow, count: Math.max(1, values.length) * blockRows.length }); for (const [valueIndex, child] of values.entries()) for (const [rowIndex, row] of blockRows.entries()) output += renderRow(row.xml, row.sourceRow, row.sourceRow + shift + valueIndex * blockRows.length + rowIndex, { ...scope, ...(child && typeof child === 'object' ? child : {}), $index: valueIndex }, block.loop); shift += values.length * blockRows.length - blockRows.length; cursor = blockRows.at(-1).end; index = block.end; continue; } const texts = [...current.xml.matchAll(/<c\b[^>]*>[\s\S]*?<\/c>/g)].map(match => cellSource(match[0])).join(''); const loop = [...texts.matchAll(/\{#(?!\!)([^{}]+)\}/g)].map(match => match[1].trim()).find(loopName => Array.isArray(matchingValue(scope, loopName))); const values = loop ? matchingValue(scope, loop) : null; const targetRow = current.sourceRow + shift; output += values ? values.map((child, valueIndex) => renderRow(current.xml, current.sourceRow, targetRow + valueIndex, { ...scope, ...(child && typeof child === 'object' ? child : {}), $index: valueIndex }, loop)).join('') : renderRow(current.xml, current.sourceRow, targetRow, scope); if (values) { expansions.push({ row: current.sourceRow, count: values.length }); shift += values.length - 1; } cursor = current.end; }
    xml = output + xml.slice(cursor);
    const shiftRow = row => Number(row) + expansions.filter(item => item.row < Number(row)).reduce((total, item) => total + item.count - 1, 0); xml = xml.replace(/(<dimension\b[^>]*\bref="[A-Z]+\d+:[A-Z]+)(\d+)(")/, (_, start, row, end) => `${start}${shiftRow(row)}${end}`); zip.file(name, xml);
  }
  return zip.generate({ type: 'nodebuffer' });
};
const xmlAttr = (source, name, fallback = '') => source.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1] ?? fallback;
const excelColumnIndex = letters => [...String(letters || '').toUpperCase()].reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0) - 1;
const excelColumnName = index => { let name = ''; for (let value = index + 1; value; value = Math.floor((value - 1) / 26)) name = String.fromCharCode(65 + (value - 1) % 26) + name; return name; };
const excelColor = tag => { const rgb = xmlAttr(tag, 'rgb'); if (rgb) return `#${rgb.slice(-6)}`; const indexed = Number(xmlAttr(tag, 'indexed', '-1')); return ({ 0:'#000000',1:'#ffffff',2:'#ff0000',3:'#00ff00',4:'#0000ff',5:'#ffff00',6:'#ff00ff',7:'#00ffff',8:'#000000',9:'#ffffff',22:'#c0c0c0',64:'#000000' })[indexed] || ''; };
const parseXlsxStyles = zip => {
  const xml = normalizeSpreadsheetXml(zip.file('xl/styles.xml')?.asText() || ''); const section = name => xml.match(new RegExp(`<${name}\\b[^>]*>([\\s\\S]*?)<\\/${name}>`))?.[1] || '';
  const fonts = [...section('fonts').matchAll(/<font\b[^>]*>([\s\S]*?)<\/font>/g)].map(match => { const body = match[1]; const color = excelColor(body.match(/<color\b[^>]*\/?>(?:<\/color>)?/)?.[0] || ''); return [body.includes('<b'), body.includes('<i'), body.includes('<u'), xmlAttr(body.match(/<sz\b[^>]*\/?>(?:<\/sz>)?/)?.[0] || '', 'val'), xmlAttr(body.match(/<name\b[^>]*\/?>(?:<\/name>)?/)?.[0] || '', 'val'), color]; });
  const fills = [...section('fills').matchAll(/<fill\b[^>]*>([\s\S]*?)<\/fill>/g)].map(match => { const pattern = xmlAttr(match[1].match(/<patternFill\b[^>]*/)?.[0] || '', 'patternType'); if (pattern !== 'solid') return ''; return excelColor(match[1].match(/<fgColor\b[^>]*\/?>(?:<\/fgColor>)?/)?.[0] || ''); });
  const borders = [...section('borders').matchAll(/<border\b[^>]*>([\s\S]*?)<\/border>/g)].map(match => ['top','right','bottom','left'].map(side => { const tag = match[1].match(new RegExp(`<${side}\\b[^>]*>[\\s\\S]*?<\\/${side}>|<${side}\\b[^>]*/>`))?.[0] || ''; const style = xmlAttr(tag, 'style'); if (!style) return ''; const color = excelColor(tag.match(/<color\b[^>]*\/?>(?:<\/color>)?/)?.[0] || '') || '#808080'; return `border-${side}:${['medium','thick','double'].includes(style) ? 2 : 1}px ${style === 'dashed' || style === 'dotted' ? style : style === 'double' ? 'double' : 'solid'} ${color}`; }).filter(Boolean).join(';'));
  return [...section('cellXfs').matchAll(/<xf\b([^>]*)(?:\/>|>([\s\S]*?)<\/xf>)/g)].map(match => { const attrs = match[1]; const body = match[2] || ''; const font = fonts[Number(xmlAttr(attrs, 'fontId', '0'))] || []; const fill = fills[Number(xmlAttr(attrs, 'fillId', '0'))] || ''; const border = borders[Number(xmlAttr(attrs, 'borderId', '0'))] || ''; const alignment = body.match(/<alignment\b[^>]*\/?>(?:<\/alignment>)?/)?.[0] || ''; const css = []; if (font[0]) css.push('font-weight:700'); if (font[1]) css.push('font-style:italic'); if (font[2]) css.push('text-decoration:underline'); if (font[3]) css.push(`font-size:${Number(font[3])}pt`); if (font[4]) css.push(`font-family:${font[4]}`); if (font[5]) css.push(`color:${font[5]}`); if (fill) css.push(`background:${fill}`); if (border) css.push(border); const horizontal = xmlAttr(alignment, 'horizontal'); const vertical = xmlAttr(alignment, 'vertical'); if (horizontal) css.push(`text-align:${horizontal === 'centerContinuous' ? 'center' : horizontal}`); if (vertical) css.push(`vertical-align:${vertical === 'center' ? 'middle' : vertical}`); if (xmlAttr(alignment, 'wrapText') === '1') css.push('white-space:pre-wrap'); return css.join(';'); });
};
const xlsxPreviewHtml = bytes => {
  const zip = assertTemplate(bytes, '.xlsx'); const shared = [...normalizeSpreadsheetXml(zip.file('xl/sharedStrings.xml')?.asText() || '').matchAll(/<si>([\s\S]*?)<\/si>/g)].map(match => spreadsheetText(match[1])); const styles = parseXlsxStyles(zip); const workbook = normalizeSpreadsheetXml(zip.file('xl/workbook.xml')?.asText() || ''); const rels = zip.file('xl/_rels/workbook.xml.rels')?.asText() || ''; const targets = new Map([...rels.matchAll(/<Relationship\b([^>]*)\/?>(?:<\/Relationship>)?/g)].map(match => [xmlAttr(match[1], 'Id'), xmlAttr(match[1], 'Target')]));
  const sheets = [...workbook.matchAll(/<sheet\b([^>]*)\/?>(?:<\/sheet>)?/g)].map((match, index) => { const target = targets.get(xmlAttr(match[1], 'r:id')) || `worksheets/sheet${index + 1}.xml`; return { name: xmlUnescape(xmlAttr(match[1], 'name', `工作表${index + 1}`)), path: target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\.\//, '')}` }; });
  const renderSheet = sheet => { const xml = normalizeSpreadsheetXml(zip.file(sheet.path)?.asText() || ''); const cells = new Map(); const rowHeights = new Map(); const hiddenRows = new Set(); let maxRow = 0; let maxCol = 0;
    for (const row of xml.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)) { const rowNumber = Number(xmlAttr(row[1], 'r', String(maxRow + 1))); maxRow = Math.max(maxRow, rowNumber); const height = Number(xmlAttr(row[1], 'ht')); if (height) rowHeights.set(rowNumber, height); if (xmlAttr(row[1], 'hidden') === '1') hiddenRows.add(rowNumber); for (const cell of row[2].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) { const ref = xmlAttr(cell[1], 'r'); const coordinate = ref.match(/^([A-Z]+)(\d+)$/i); if (!coordinate) continue; const column = excelColumnIndex(coordinate[1]); const rowIndex = Number(coordinate[2]); maxCol = Math.max(maxCol, column); maxRow = Math.max(maxRow, rowIndex); const type = xmlAttr(cell[1], 't'); const style = Number(xmlAttr(cell[1], 's', '0')); const raw = xmlUnescape(cell[2].match(/<v>([\s\S]*?)<\/v>/)?.[1] || ''); const formula = xmlUnescape(cell[2].match(/<f(?:\s[^>]*)?>([\s\S]*?)<\/f>/)?.[1] || ''); let value = raw; if (type === 's') value = shared[Number(raw)] || ''; else if (type === 'inlineStr') value = spreadsheetText(cell[2]); else if (type === 'b') value = raw === '1' ? 'TRUE' : 'FALSE'; else if (!raw && formula) value = `=${formula}`; cells.set(`${rowIndex}:${column}`, { value, style }); } }
    const widths = new Map(); const hiddenCols = new Set(); for (const col of xml.matchAll(/<col\b([^>]*)\/?>(?:<\/col>)?/g)) { const min = Number(xmlAttr(col[1], 'min', '1')) - 1; const max = Number(xmlAttr(col[1], 'max', String(min + 1))) - 1; for (let index = min; index <= max; index++) { const width = Number(xmlAttr(col[1], 'width')); if (width) widths.set(index, Math.max(24, Math.min(420, width * 7.4))); if (xmlAttr(col[1], 'hidden') === '1') hiddenCols.add(index); maxCol = Math.max(maxCol, index); } }
    const mergeStarts = new Map(); const mergeCovered = new Set(); for (const merge of xml.matchAll(/<mergeCell\b[^>]*\bref="([A-Z]+)(\d+):([A-Z]+)(\d+)"[^>]*\/?>(?:<\/mergeCell>)?/gi)) { const c1 = excelColumnIndex(merge[1]); const r1 = Number(merge[2]); const c2 = excelColumnIndex(merge[3]); const r2 = Number(merge[4]); mergeStarts.set(`${r1}:${c1}`, { colspan: c2 - c1 + 1, rowspan: r2 - r1 + 1 }); for (let row = r1; row <= r2; row++) for (let col = c1; col <= c2; col++) if (row !== r1 || col !== c1) mergeCovered.add(`${row}:${col}`); maxRow = Math.max(maxRow, r2); maxCol = Math.max(maxCol, c2); }
    const colgroup = `<col class="row-number-col">${Array.from({ length: maxCol + 1 }, (_, column) => `<col style="width:${Number((widths.get(column) || 88).toFixed(2))}px"${hiddenCols.has(column) ? ' class="hidden-col"' : ''}>`).join('')}`; const header = `<tr><th class="corner"></th>${Array.from({ length: maxCol + 1 }, (_, column) => `<th${hiddenCols.has(column) ? ' class="hidden-col"' : ''}>${excelColumnName(column)}</th>`).join('')}</tr>`;
    const body = Array.from({ length: maxRow }, (_, offset) => { const row = offset + 1; if (hiddenRows.has(row)) return ''; const cellsHtml = Array.from({ length: maxCol + 1 }, (_, column) => { if (hiddenCols.has(column) || mergeCovered.has(`${row}:${column}`)) return ''; const cell = cells.get(`${row}:${column}`) || { value: '', style: 0 }; const merge = mergeStarts.get(`${row}:${column}`); const attrs = `${merge?.colspan > 1 ? ` colspan="${merge.colspan}"` : ''}${merge?.rowspan > 1 ? ` rowspan="${merge.rowspan}"` : ''}`; return `<td${attrs} style="${styles[cell.style] || ''}" title="${xmlEscape(cell.value)}">${xmlEscape(cell.value).replace(/\r?\n/g, '<br>')}</td>`; }).join(''); return `<tr${rowHeights.has(row) ? ` style="height:${Math.max(18, rowHeights.get(row) * 1.34)}px"` : ''}><th class="row-number">${row}</th>${cellsHtml}</tr>`; }).join(''); return `<section class="sheet"><h2>${xmlEscape(sheet.name)}</h2><div class="sheet-scroll"><table><colgroup>${colgroup}</colgroup><thead>${header}</thead><tbody>${body}</tbody></table></div></section>`; };
  const rendered = (sheets.length ? sheets : [{ name:'工作表1', path:'xl/worksheets/sheet1.xml' }]).map(renderSheet).join(''); return `<!doctype html><html><head><meta charset="utf-8"><style>*{box-sizing:border-box}body{margin:0;padding:20px;background:#eef1f5;color:#202124;font:12px/1.35 -apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif}.sheet{margin:0 auto 22px;padding:14px;background:#fff;border-radius:8px;box-shadow:0 2px 12px #0002}.sheet h2{margin:0 0 10px;font-size:13px;font-weight:600}.sheet-scroll{max-width:100%;overflow:auto;border:1px solid #b9c0ca}table{border-collapse:collapse;table-layout:fixed;width:max-content;min-width:100%}th,td{height:24px;border:1px solid #d9dde3;padding:3px 5px;overflow:hidden;white-space:pre;text-overflow:ellipsis}thead th,.row-number{position:sticky;background:#f3f5f7;color:#596273;font-weight:400;text-align:center}.row-number{left:0;z-index:2;width:42px}.corner{left:0;z-index:3;width:42px}.row-number-col{width:42px}.hidden-col{display:none}td{background:#fff;vertical-align:bottom}</style></head><body>${rendered}</body></html>`;
};
export async function previewRecord(id, record) { const item = (await readCatalog()).find(entry => entry.id === id); if (!item) throw new Error('模板不存在'); const extension = item.extension || extensionOf(item.fileName); const bytes = await fs.readFile(path.join(templatesDir, item.fileName)); if (extension === '.docx') return { template: item, kind: 'docx', bytes: renderDocx(bytes, templateScope(item, record)) }; if (extension === '.xlsx') return { template: item, kind: 'html', html: xlsxPreviewHtml(renderXlsx(bytes, record, item)) }; throw new Error('该模板格式暂不支持记录预览'); }
const matchingValue = (object, key) => { if (!object || typeof object !== 'object') return undefined; if (Object.hasOwn(object, key)) return object[key]; const normalized = normalizeKey(key); const match = Object.keys(object).find(candidate => legacyPlaceholderName(candidate) === key) ?? Object.keys(object).find(candidate => normalizeKey(candidate) === normalized); return match === undefined ? undefined : object[match]; };
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
  const layout = decodeLayout(bytes).content; const scope = templateScope({ fields: layoutFields(layout) }, record); const setting = layout.pageSetting || {}; const adaptive = setting.mode === 'auto' || setting.pageSizeMode === 'auto';
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
  const pageSize = adaptive ? 'auto' : `${Number(setting.width || 210)}mm ${Number(setting.height || 297)}mm`; const pageWidth = adaptive ? '100%' : `${Number(setting.width || 210)}mm`; const pageHeight = adaptive ? 'auto' : `${Number(setting.height || 297)}mm`; return Buffer.from(`<!doctype html><html><head><meta charset="utf-8"><style>@page{size:${pageSize};margin:0}*{box-sizing:border-box}html,body{margin:0;color:#111;font-family:"Microsoft YaHei","SimSun",sans-serif;font-size:${Number(layout.settings?.fontSize || 9)}pt}.page{width:${pageWidth};min-height:${pageHeight};padding:${Number(setting.paddingTop || 5)}mm ${Number(setting.paddingRight || 13.5)}mm ${Number(setting.paddingBottom || 5)}mm ${Number(setting.paddingLeft || 13.5)}mm;page-break-after:always;background:#fff}.page:last-child{page-break-after:auto}.row{display:flex;width:100%}.column{min-width:0}table{width:100%;border-collapse:collapse;table-layout:fixed}td{border:1px solid #222;padding:2px 3px;overflow-wrap:anywhere}tr{break-inside:avoid}p{min-height:1em}</style></head><body>${pages}</body></html>`, 'utf8');
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
