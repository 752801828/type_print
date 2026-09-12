import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import PDFDocument from 'pdfkit';
import * as fontkitModule from 'fontkit';
const fontkit = fontkitModule.default || fontkitModule;

const root = path.resolve(fileURLToPath(new URL('../data', import.meta.url)));
const templatesDir = path.join(root, 'templates');
const outputsDir = path.join(root, 'outputs');
const catalogFile = path.join(root, 'templates.json');
const supportedExtensions = new Set(['.docx', '.xlsx', '.xls', '.pdf', '.txt', '.layout', '.json']);
const cleanName = name => String(name || 'template.docx').replace(/[^\p{L}\p{N}._ -]/gu, '_').slice(0, 120) || 'template.docx';
const extensionOf = name => path.extname(String(name || '')).toLowerCase();
const cleanLayoutName = name => String(name || '').replace(/^(?:🔵|🟡|🔴|🟢|⚪)\s*/u, '').trim();
const normalizeKey = value => cleanLayoutName(value).replace(/[\u2600-\u27BF\u{1F300}-\u{1FAFF}]/gu, '').replace(/[\s_\-：:（）()·]/g, '').toLowerCase();
const legacyPlaceholderName = value => String(value || '').replace(/[^A-Za-z0-9_\u3400-\u9fff]/g, '_');
const readCatalog = async () => { try { return JSON.parse(await fs.readFile(catalogFile, 'utf8')); } catch { return []; } };
const writeCatalog = value => fs.writeFile(catalogFile, JSON.stringify(value, null, 2), 'utf8');
const xmlEscape = value => String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const isTextFragment = value => value && typeof value === 'object' && (value.type === 'text' || Object.hasOwn(value, 'text'));
const cleanInvisible = value => String(value ?? '').replace(/[\u061C\u200B-\u200F\u202A-\u202E\u2060-\u206F\uFEFF]/g, '');
const valueText = value => Array.isArray(value) ? value.map(valueText).filter(Boolean).join(value.length && value.every(isTextFragment) ? '' : '、') : value && typeof value === 'object' ? (Object.hasOwn(value, 'text') ? cleanInvisible(value.text) : Object.hasOwn(value, 'name') ? cleanInvisible(value.name) : Object.hasOwn(value, 'value') ? valueText(value.value) : cleanInvisible(JSON.stringify(value))) : cleanInvisible(value);
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
const fieldValue = (object, name) => {
  if (!object || typeof object !== 'object') return undefined;
  const entries = Object.entries(object).filter(([, value]) => value != null && value !== '');
  const exact = entries.find(([key]) => key === name) || entries.find(([key]) => legacyPlaceholderName(key) === name);
  if (exact) return exact[1];
  const normalized = entries.find(([key]) => normalizeKey(key) === normalizeKey(name));
  if (normalized) return normalized[1];
  return undefined;
};
const numberValue = value => { if (value && typeof value === 'object' && value.value !== undefined) return numberValue(value.value); const number = Number(String(value ?? '').replace(/[,，￥¥\s]/g, '')); return Number.isFinite(number) ? number : 0; };
const numericScalar = value => { const raw = value && typeof value === 'object' && value.value !== undefined ? value.value : value; return typeof raw === 'number' && Number.isFinite(raw) ? raw : null; };
const chineseCurrency = value => { const amount = Math.round(numberValue(value) * 100); const digits = '零壹贰叁肆伍陆柒捌玖'; const small = ['', '拾', '佰', '仟']; const large = ['', '万', '亿', '兆']; const groupText = group => { let out = ''; let zero = false; for (let i = 3; i >= 0; i--) { const digit = Math.floor(group / (10 ** i)) % 10; if (!digit) { if (out) zero = true; continue; } if (zero) out += '零'; out += digits[digit] + small[i]; zero = false; } return out; }; let out = ''; let missing = false; const integer = Math.floor(amount / 100); for (let i = 3; i >= 0; i--) { const group = Math.floor(integer / (10000 ** i)) % 10000; if (!group) { if (out) missing = true; continue; } if (out && (missing || group < 1000)) out += '零'; out += groupText(group) + large[i]; missing = false; } if (!out) out = '零'; const jiao = Math.floor(amount / 10) % 10; const fen = amount % 10; return `${numberValue(value) < 0 ? '负' : ''}${out}元${jiao ? `${digits[jiao]}角` : ''}${fen ? `${!jiao && integer ? '零' : ''}${digits[fen]}分` : (!jiao ? '整' : '')}`; };
const templateScope = (item, record) => {
  const source = {};
  for (const [key, value] of Object.entries(record || {})) { const indexed = withIndexes(value ?? ''); source[String(key)] = indexed; source[cleanLayoutName(key)] ??= indexed; source[legacyPlaceholderName(key)] ??= indexed; }
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
    if (!Object.hasOwn(item, 'pinned')) { item.pinned = false; changed = true; }
    const extension = item.extension || extensionOf(item.fileName);
    if (!Object.hasOwn(item, 'autoUpload')) {
      item.autoUpload = false; item.outputFieldId = ''; item.outputReplace = false;
      if (['.txt', '.layout', '.json'].includes(extension)) try { const config = decodeLayout(await fs.readFile(path.join(templatesDir, item.fileName))).content.settings?.exportFile || {}; item.outputFieldId = String(config.targetFieldId || ''); item.outputReplace = Boolean(config.isReplace); item.autoUpload = Boolean(item.outputFieldId); } catch {}
      changed = true;
    }
    if (item.parserVersion === 2 && !item.fields?.some(field => /[<>]/.test(field.name)) && item.fields) continue;
    try { const bytes = await fs.readFile(path.join(templatesDir, item.fileName)); const zip = assertTemplate(bytes, extension); if (zip) { item.fields = extractFields(zip, extension); item.extension = extension; item.parserVersion = 2; changed = true; } else if (['.txt', '.layout', '.json'].includes(extension)) { item.fields = layoutFields(decodeLayout(bytes).content); item.extension = extension; item.parserVersion = 2; changed = true; } } catch {}
  }
  if (changed) await writeCatalog(catalog);
  const filtered = scope.tableId ? catalog.filter(item => item.tableId === scope.tableId && (!scope.baseId || item.baseId === scope.baseId)) : catalog;
  return filtered.sort((a, b) => Number(Boolean(b.pinned)) - Number(Boolean(a.pinned)) || b.createdAt.localeCompare(a.createdAt));
}

export async function saveTemplate(originalName, bytes, binding = {}) {
  const name = cleanName(originalName); const extension = extensionOf(name);
  assertTemplate(bytes, extension);
  const id = crypto.randomUUID(); const fileName = `${id}${extension}`;
  await fs.mkdir(templatesDir, { recursive: true }); await fs.writeFile(path.join(templatesDir, fileName), bytes);
  const zip = extension === '.docx' || extension === '.xlsx' ? new PizZip(bytes) : null;
  const layout = ['.txt', '.layout', '.json'].includes(extension) ? decodeLayout(bytes).content : null;
  const exportFile = layout?.settings?.exportFile || {}; const outputFieldId = String(binding.outputFieldId || exportFile.targetFieldId || '');
  const item = { id, name, extension, fileName, fields: zip ? extractFields(zip, extension) : layout ? layoutFields(layout) : [], stats: zip ? templateStats(zip, extension) : layout ? { pages: layout.document.pages.length, blocks: JSON.stringify(layout).match(/"type":/g)?.length || 0 } : {}, parserVersion: 2, pinned: false, baseId: String(binding.baseId || ''), tableId: String(binding.tableId || ''), viewId: String(binding.viewId || ''), baseName: String(binding.baseName || ''), tableName: String(binding.tableName || ''), outputNamePattern: String(binding.outputNamePattern || ''), autoUpload: Object.hasOwn(binding, 'autoUpload') ? Boolean(binding.autoUpload) : Boolean(outputFieldId), outputFieldId, outputReplace: Object.hasOwn(binding, 'outputReplace') ? Boolean(binding.outputReplace) : Boolean(exportFile.isReplace), createdAt: new Date().toISOString() };
  const catalog = await readCatalog(); catalog.push(item); await writeCatalog(catalog); return item;
}

export async function removeTemplate(id) {
  const catalog = await readCatalog(); const item = catalog.find(entry => entry.id === id); if (!item) return false;
  await fs.rm(path.join(templatesDir, item.fileName), { force: true }); await writeCatalog(catalog.filter(entry => entry.id !== id)); return true;
}

export async function updateTemplate(id, changes = {}) {
  const catalog = await readCatalog(); const item = catalog.find(entry => entry.id === id); if (!item) throw new Error('模板不存在');
  if (Object.hasOwn(changes, 'outputNamePattern')) item.outputNamePattern = String(changes.outputNamePattern || '').trim().slice(0, 120);
  if (Object.hasOwn(changes, 'autoUpload')) item.autoUpload = Boolean(changes.autoUpload);
  if (Object.hasOwn(changes, 'outputFieldId')) item.outputFieldId = String(changes.outputFieldId || '').trim().slice(0, 80);
  if (Object.hasOwn(changes, 'outputReplace')) item.outputReplace = Boolean(changes.outputReplace);
  if (Object.hasOwn(changes, 'pinned')) item.pinned = Boolean(changes.pinned);
  if (Object.hasOwn(changes, 'name')) { const value = cleanName(changes.name).trim(); if (!value) throw new Error('模板名称不能为空'); const extension = item.extension || extensionOf(item.fileName); const declared = extensionOf(value); item.name = cleanName(`${supportedExtensions.has(declared) ? path.basename(value, declared) : value}${extension}`); }
  await writeCatalog(catalog); return item;
}

export async function duplicateTemplate(id) {
  const item = (await readCatalog()).find(entry => entry.id === id); if (!item) throw new Error('模板不存在'); const extension = item.extension || extensionOf(item.fileName); const bytes = await fs.readFile(path.join(templatesDir, item.fileName));
  return saveTemplate(`${path.parse(item.name).name}-副本${extension}`, bytes, { baseId: item.baseId, tableId: item.tableId, viewId: item.viewId, baseName: item.baseName, tableName: item.tableName, outputNamePattern: item.outputNamePattern, autoUpload: item.autoUpload, outputFieldId: item.outputFieldId, outputReplace: item.outputReplace });
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

const resolveTemplateValue = (scope, key) => { let value = scope; let nestedArray = false; for (const part of String(key).trim().split(/\s*[./]\s*/)) { if (Array.isArray(value)) { nestedArray = true; value = value.map(item => matchingValue(item, part)).filter(item => item != null && item !== ''); } else value = matchingValue(value, part); } return nestedArray && Array.isArray(value) ? valueText(value) : value; };
const conditionLiteral = raw => { const value = String(raw).trim(); if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) return value.slice(1, -1); if (/^-?\d+(?:\.\d+)?$/.test(value)) return Number(value); if (value === 'true') return true; if (value === 'false') return false; if (value === 'null') return null; return value; };
const templateParser = tag => ({ get(scope) { const index = String(tag).match(/^\$index\s*\+\s*(\d+)$/); if (index) return Number(scope?.$index || 0) + Number(index[1]); const condition = String(tag).match(/^(.+?)\s*(==|!=|>=|<=|>|<)\s*(.+)$/); if (!condition) { const value = resolveTemplateValue(scope, tag); return value && typeof value === 'object' && !Array.isArray(value) ? valueText(value) : value; } const rawLeft = resolveTemplateValue(scope, condition[1]); const left = rawLeft && typeof rawLeft === 'object' ? valueText(rawLeft) : rawLeft; const right = conditionLiteral(condition[3]); if (condition[2] === '==') return String(left ?? '') === String(right ?? ''); if (condition[2] === '!=') return String(left ?? '') !== String(right ?? ''); if (condition[2] === '>') return Number(left) > Number(right); if (condition[2] === '<') return Number(left) < Number(right); if (condition[2] === '>=') return Number(left) >= Number(right); return Number(left) <= Number(right); } });
const renderDocx = (bytes, record) => { const zip = assertTemplate(bytes, '.docx'); for (const name of Object.keys(zip.files).filter(file => /^word\/.*\.xml$/.test(file))) zip.file(name, zip.file(name).asText().replace(/\{#!\s*([^{}]+?)\}/g, '{^$1}')); const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, parser: templateParser }); doc.render(record); return doc.getZip().generate({ type: 'nodebuffer' }); };
const xmlUnescape = value => String(value || '').replace(/&#x([0-9a-f]+);/gi, (_, code) => String.fromCodePoint(parseInt(code, 16))).replace(/&#(\d+);/g, (_, code) => String.fromCodePoint(Number(code))).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&');
const normalizeSpreadsheetXml = value => { const xml = String(value || ''); const prefix = xml.match(/<([A-Za-z_][\w.-]*):(worksheet|sst|styleSheet|workbook)\b/)?.[1]; if (!prefix) return xml; const escaped = prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); return xml.replace(new RegExp(`<(/?)${escaped}:`, 'g'), '<$1').replace(new RegExp(`xmlns:${escaped}=`), 'xmlns='); };
const spreadsheetCellPattern = /<c\b((?![^>]*\/>)[^>]*)>([\s\S]*?)<\/c>/g;
const spreadsheetText = value => xmlUnescape([...String(value || '').matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(match => match[1]).join(''));
const markSpreadsheetForRecalculation = zip => {
  const name = 'xl/workbook.xml'; let xml = zip.file(name)?.asText(); if (!xml) return;
  const settings = ' calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1" calcCompleted="0"';
  if (/<(?:[\w.-]+:)?calcPr\b/.test(xml)) xml = xml.replace(/<((?:[\w.-]+:)?calcPr)\b([^>]*?)(\/?)>/, (_, tag, attributes, closed) => `<${tag}${attributes.replace(/\s+(?:calcMode|fullCalcOnLoad|forceFullCalc|calcCompleted)="[^"]*"/g, '')}${settings}${closed}>`);
  else xml = xml.replace(/<\/((?:[\w.-]+:)?workbook)>/, (_, tag) => `<${tag.includes(':') ? `${tag.split(':')[0]}:` : ''}calcPr${settings}/></${tag}>`);
  zip.file(name, xml);
};
const expandedSpreadsheetRow = (row, expansions, edge = 'start') => {
  let shift = 0;
  for (const expansion of expansions) {
    const sourceEnd = expansion.start + expansion.sourceCount - 1;
    if (row < expansion.start) break;
    if (row <= sourceEnd) return expansion.start + shift + (edge === 'end' ? Math.max(0, expansion.targetCount - 1) : 0);
    shift += expansion.targetCount - expansion.sourceCount;
  }
  return row + shift;
};
const expandedSpreadsheetRef = (ref, expansions) => {
  const cells = String(ref || '').split(':');
  const move = (cell, edge) => cell.replace(/(\$?[A-Z]{1,3})(\$?)(\d+)$/i, (_, column, fixed, row) => `${column}${fixed}${expandedSpreadsheetRow(Number(row), expansions, edge)}`);
  const start = move(cells[0], 'start'), end = move(cells.at(-1), 'end');
  return start === end ? start : `${start}:${end}`;
};
const expandSpreadsheetMerges = (xml, expansions) => xml.replace(/<mergeCells\b[^>]*>([\s\S]*?)<\/mergeCells>/, (_, body) => {
  const refs = [];
  for (const match of body.matchAll(/<mergeCell\b[^>]*\bref="([^"]+)"[^>]*\/?>(?:<\/mergeCell>)?/g)) {
    const [from, to = from] = match[1].split(':'); const startRow = Number(from.match(/\d+$/)?.[0]); const endRow = Number(to.match(/\d+$/)?.[0]);
    const expansion = expansions.find(item => startRow >= item.start && endRow < item.start + item.sourceCount);
    if (!expansion || !expansion.sourceCount || !expansion.targetCount) { refs.push(expandedSpreadsheetRef(match[1], expansions)); continue; }
    const repeats = Math.max(1, Math.round(expansion.targetCount / expansion.sourceCount)); const shift = expandedSpreadsheetRow(expansion.start, expansions) - expansion.start;
    for (let repeat = 0; repeat < repeats; repeat++) refs.push(match[1].replace(/(\$?[A-Z]{1,3}\$?)(\d+)/g, (_, column, row) => `${column}${Number(row) + shift + repeat * expansion.sourceCount}`));
  }
  return refs.length ? `<mergeCells count="${refs.length}">${refs.map(ref => `<mergeCell ref="${ref}"/>`).join('')}</mergeCells>` : '';
});
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
  const renderRow = (rowXml, sourceRow, targetRow, rowScope, loopName = '') => rowXml.replace(/(<row\b[^>]*\br=")\d+("[^>]*>)/, `$1${targetRow}$2`).replace(spreadsheetCellPattern, (cell, attributes, inner) => { const nextAttributes = attributes.replace(/\br="([A-Z]+)\d+"/, `r="$1${targetRow}"`); const source = cellSource(cell); if (!source.includes('{')) { const delta = targetRow - sourceRow; const shifted = delta ? inner.replace(/(<f[^>]*>)([\s\S]*?)(<\/f>)/g, (_, open, formula, close) => `${open}${formula.replace(/(\$?[A-Z]{1,3})(\$?)(\d+)/g, (ref, column, fixed, row) => fixed ? ref : `${column}${Number(row) + delta}`)}${close}`) : inner; return `<c${nextAttributes}>${shifted}</c>`; } const rendered = renderTemplateText(source, rowScope, loopName); const escapedLoop = loopName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); const sourceWithoutLoop = loopName ? source.replace(new RegExp(`\\{#\\s*${escapedLoop}\\s*\\}|\\{/\\s*${escapedLoop}\\s*\\}`, 'gi'), '').trim() : source; const exact = sourceWithoutLoop.match(/^\s*\{\{?\s*([^{}]+?)\s*\}\}?\s*$/); const key = exact?.[1]?.trim() || ''; const scalar = exact && !/^\$index\b/i.test(key) ? numericScalar(templateParser(key).get(rowScope)) : null; if (scalar !== null && /^-?(?:\d+\.?\d*|\.\d+)$/.test(rendered.trim())) { const numericAttributes = nextAttributes.replace(/\s+t="[^"]*"/g, ''); return `<c${numericAttributes}><v>${xmlEscape(rendered.trim())}</v></c>`; } const inlineAttributes = nextAttributes.replace(/\s+t="[^"]*"/g, '') + ' t="inlineStr"'; return `<c${inlineAttributes}><is><t xml:space="preserve">${xmlEscape(rendered)}</t></is></c>`; });
  for (const name of Object.keys(zip.files).filter(file => /^xl\/worksheets\/sheet\d+\.xml$/.test(file))) {
    let shift = 0; const expansions = []; let xml = normalizeSpreadsheetXml(zip.file(name).asText());
    const rowMatches = [...xml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>/g)]; const rows = rowMatches.map(match => ({ xml: match[0], sourceRow: Number(match[1]), start: match.index, end: match.index + match[0].length })); const blocks = new Map();
    for (let index = 0; index < rows.length; index++) { const texts = [...rows[index].xml.matchAll(spreadsheetCellPattern)].map(match => cellSource(match[0])).join(''); const loop = [...texts.matchAll(/\{#(?!\!)([^{}]+)\}/g)].map(match => match[1].trim()).find(loopName => Array.isArray(matchingValue(scope, loopName))); if (!loop) continue; const close = new RegExp(`\\{\\/\\s*${loop.replace(/[.*+?^${}()|[\]\\\\]/g, '\\\\$&')}\\s*\\}`, 'i'); const end = rows.findIndex((candidate, candidateIndex) => candidateIndex >= index && close.test([...candidate.xml.matchAll(spreadsheetCellPattern)].map(match => cellSource(match[0])).join(''))); if (end >= index) blocks.set(index, { end, loop }); }
    let output = ''; let cursor = 0;
    for (let index = 0; index < rows.length; index++) { const current = rows[index]; output += xml.slice(cursor, current.start); const block = blocks.get(index); if (block) { const values = matchingValue(scope, block.loop) || []; const blockRows = rows.slice(index, block.end + 1); expansions.push({ start: current.sourceRow, sourceCount: blockRows.length, targetCount: values.length * blockRows.length }); for (const [valueIndex, child] of values.entries()) for (const [rowIndex, row] of blockRows.entries()) output += renderRow(row.xml, row.sourceRow, row.sourceRow + shift + valueIndex * blockRows.length + rowIndex, { ...scope, ...(child && typeof child === 'object' ? child : {}), $index: valueIndex }, block.loop); shift += values.length * blockRows.length - blockRows.length; cursor = blockRows.at(-1).end; index = block.end; continue; } const texts = [...current.xml.matchAll(spreadsheetCellPattern)].map(match => cellSource(match[0])).join(''); const loop = [...texts.matchAll(/\{#(?!\!)([^{}]+)\}/g)].map(match => match[1].trim()).find(loopName => Array.isArray(matchingValue(scope, loopName))); const values = loop ? matchingValue(scope, loop) : null; const targetRow = current.sourceRow + shift; output += values ? values.map((child, valueIndex) => renderRow(current.xml, current.sourceRow, targetRow + valueIndex, { ...scope, ...(child && typeof child === 'object' ? child : {}), $index: valueIndex }, loop)).join('') : renderRow(current.xml, current.sourceRow, targetRow, scope); if (values) { expansions.push({ start: current.sourceRow, sourceCount: 1, targetCount: values.length }); shift += values.length - 1; } cursor = current.end; }
    xml = output + xml.slice(cursor);
    xml = xml.replace(/(<dimension\b[^>]*\bref="[A-Z]+\d+:[A-Z]+)(\d+)(")/, (_, start, row, end) => `${start}${expandedSpreadsheetRow(Number(row), expansions, 'end')}${end}`);
    xml = xml.replace(/(<conditionalFormatting\b[^>]*\bsqref=")([^"]+)(")/g, (_, start, refs, end) => `${start}${refs.split(/\s+/).filter(Boolean).map(ref => expandedSpreadsheetRef(ref, expansions)).join(' ')}${end}`);
    xml = expandSpreadsheetMerges(xml, expansions); zip.file(name, xml);
  }
  markSpreadsheetForRecalculation(zip);
  evaluateXlsxFormulas(zip);
  return zip.generate({ type: 'nodebuffer' });
};
const xmlAttr = (source, name, fallback = '') => source.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1] ?? fallback;
const excelColumnIndex = letters => [...String(letters || '').toUpperCase()].reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0) - 1;
const excelColumnName = index => { let name = ''; for (let value = index + 1; value; value = Math.floor((value - 1) / 26)) name = String.fromCharCode(65 + (value - 1) % 26) + name; return name; };
const splitFormulaArgs = source => { const args = []; let start = 0; let depth = 0; let quote = ''; for (let index = 0; index < source.length; index++) { const char = source[index]; if (quote) { if (char === quote && source[index - 1] !== '\\') quote = ''; continue; } if (char === '"' || char === "'") { quote = char; continue; } if (char === '(') depth++; else if (char === ')') depth--; else if (char === ',' && depth === 0) { args.push(source.slice(start, index).trim()); start = index + 1; } } args.push(source.slice(start).trim()); return args; };
const parseSpreadsheetRef = value => { const match = String(value || '').trim().replace(/\$/g, '').match(/^([A-Z]{1,3})(\d+)?$/i); return match ? { column: excelColumnIndex(match[1]), row: match[2] ? Number(match[2]) : null } : null; };
const evaluateXlsxFormulas = zip => {
  const workbook = normalizeSpreadsheetXml(zip.file('xl/workbook.xml')?.asText() || ''); const rels = zip.file('xl/_rels/workbook.xml.rels')?.asText() || ''; const targets = new Map([...rels.matchAll(/<Relationship\b([^>]*)\/?>(?:<\/Relationship>)?/g)].map(match => [xmlAttr(match[1], 'Id'), xmlAttr(match[1], 'Target')]));
  const sheets = [...workbook.matchAll(/<sheet\b([^>]*)\/?>(?:<\/sheet>)?/g)].map((match, index) => { const target = targets.get(xmlAttr(match[1], 'r:id')) || `worksheets/sheet${index + 1}.xml`; return { name: xmlUnescape(xmlAttr(match[1], 'name', `工作表${index + 1}`)), path: target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\.\//, '')}` }; });
  if (!sheets.length) sheets.push({ name: '工作表1', path: 'xl/worksheets/sheet1.xml' });
  const shared = [...normalizeSpreadsheetXml(zip.file('xl/sharedStrings.xml')?.asText() || '').matchAll(/<si>([\s\S]*?)<\/si>/g)].map(match => spreadsheetText(match[1])); const data = new Map();
  for (const sheet of sheets) { const xml = normalizeSpreadsheetXml(zip.file(sheet.path)?.asText() || ''); const cells = new Map(); let maxRow = 1; for (const match of xml.matchAll(spreadsheetCellPattern)) { const ref = xmlAttr(match[1], 'r'); if (!ref) continue; maxRow = Math.max(maxRow, Number(ref.match(/\d+$/)?.[0] || 1)); cells.set(ref.toUpperCase(), { xml: match[0], attributes: match[1], inner: match[2] }); } data.set(sheet.name, { sheet, xml, cells, maxRow }); }
  const sheetByName = new Map([...data.keys()].map(name => [String(name).toLowerCase(), name])); const canonicalSheet = name => sheetByName.get(String(name || '').replace(/^'|'$/g, '').toLowerCase()) || name;
  const splitSheetRef = token => { const value = String(token || '').trim(); const index = value.lastIndexOf('!'); return index < 0 ? { sheet: null, ref: value } : { sheet: canonicalSheet(value.slice(0, index)), ref: value.slice(index + 1) }; };
  let formulaValue;
  const cellValue = (sheetName, ref, stack = new Set()) => { const entry = data.get(canonicalSheet(sheetName)); const cleanRef = String(ref || '').replace(/\$/g, '').toUpperCase(); const cell = entry?.cells.get(cleanRef); if (!cell) return undefined; const key = `${entry.sheet.name}!${cleanRef}`; if (stack.has(key)) return undefined; const formula = xmlUnescape(cell.inner.match(/<f(?:\s[^>]*)?>([\s\S]*?)<\/f>/)?.[1] || ''); if (formula) return formulaValue(formula, entry.sheet.name, new Set([...stack, key])); const type = xmlAttr(cell.attributes, 't'); const raw = xmlUnescape(cell.inner.match(/<v>([\s\S]*?)<\/v>/)?.[1] || ''); if (type === 's') return shared[Number(raw)] || ''; if (type === 'inlineStr') return spreadsheetText(cell.inner); if (type === 'b') return raw === '1'; if (raw === '') return ''; const number = Number(raw); return Number.isFinite(number) ? number : raw; };
  formulaValue = (source, currentSheet, stack) => { let formula = String(source || '').trim().replace(/^=/, '').trim(); while (formula.startsWith('(') && formula.endsWith(')')) formula = formula.slice(1, -1).trim(); const literal = formula.match(/^(?:"([\s\S]*)"|'([\s\S]*)')$/); if (literal) return literal[1] ?? literal[2]; if (/^-?(?:\d+\.?\d*|\.\d+)$/.test(formula)) return Number(formula); const ifError = formula.match(/^IFERROR\(([\s\S]*)\)$/i); if (ifError) { const [candidate, fallback] = splitFormulaArgs(ifError[1]); const value = formulaValue(candidate, currentSheet, stack); return value === undefined || value === null || String(value) === '#N/A' ? formulaValue(fallback, currentSheet, stack) : value; } const vlookup = formula.match(/^VLOOKUP\(([\s\S]*)\)$/i); if (vlookup) { const [lookupExpr, rangeExpr, columnExpr, exactExpr] = splitFormulaArgs(vlookup[1]); const lookup = formulaValue(lookupExpr, currentSheet, stack); const rangeRef = splitSheetRef(rangeExpr); const rangeSheet = data.get(canonicalSheet(rangeRef.sheet || currentSheet)); const parts = String(rangeRef.ref || '').replace(/\$/g, '').split(':'); const start = parseSpreadsheetRef(parts[0]); const end = parseSpreadsheetRef(parts[1] || parts[0]); if (!rangeSheet || !start || !end) return undefined; const firstRow = start.row || 1; const lastRow = end.row || rangeSheet.maxRow; const firstColumn = start.column; const lastColumn = end.column ?? firstColumn; const wantedColumn = firstColumn + Math.max(1, Number(formulaValue(columnExpr, currentSheet, stack) || 1)) - 1; const exact = exactExpr === undefined ? true : Number(formulaValue(exactExpr, currentSheet, stack)) === 0; let best; for (let row = firstRow; row <= lastRow; row++) { const candidate = cellValue(rangeSheet.sheet.name, `${excelColumnName(firstColumn)}${row}`, stack); if (String(candidate ?? '').trim() === String(lookup ?? '').trim()) { best = row; break; } if (!exact && candidate !== undefined && String(candidate) > String(lookup)) break; } return best === undefined || wantedColumn > lastColumn ? undefined : cellValue(rangeSheet.sheet.name, `${excelColumnName(wantedColumn)}${best}`, stack); } const ref = splitSheetRef(formula); const parsed = parseSpreadsheetRef(ref.ref); if (parsed?.row) return cellValue(ref.sheet || currentSheet, `${excelColumnName(parsed.column)}${parsed.row}`, stack); const comparison = formula.match(/^([\s\S]+?)(==|!=|>=|<=|>|<)([\s\S]+)$/); if (comparison) { const left = formulaValue(comparison[1], currentSheet, stack); const right = formulaValue(comparison[3], currentSheet, stack); return comparison[2] === '==' ? String(left) === String(right) : comparison[2] === '!=' ? String(left) !== String(right) : comparison[2] === '>' ? Number(left) > Number(right) : comparison[2] === '<' ? Number(left) < Number(right) : comparison[2] === '>=' ? Number(left) >= Number(right) : Number(left) <= Number(right); } return undefined; };
  const updateCell = (entry, cell) => { const formula = xmlUnescape(cell.inner.match(/<f(?:\s[^>]*)?>([\s\S]*?)<\/f>/)?.[1] || ''); if (!formula) return cell.xml; const ref = xmlAttr(cell.attributes, 'r'); const value = formulaValue(formula, entry.sheet.name, new Set()); if (value === undefined || value === null) return cell.xml; const isNumber = typeof value === 'number' && Number.isFinite(value); let attributes = cell.attributes.replace(/\s+t="[^"]*"/g, ''); if (!isNumber) attributes += ' t="str"'; const serialized = xmlEscape(String(value)); const inner = /<v>[\s\S]*?<\/v>/.test(cell.inner) ? cell.inner.replace(/<v>[\s\S]*?<\/v>/, `<v>${serialized}</v>`) : `${cell.inner}<v>${serialized}</v>`; return `<c${attributes}>${inner}</c>`; };
  for (const entry of data.values()) { let changed = false; const xml = entry.xml.replace(spreadsheetCellPattern, (cell, attributes, inner) => { const updated = updateCell(entry, { xml: cell, attributes, inner }); changed ||= updated !== cell; return updated; }); if (changed) zip.file(entry.sheet.path, xml); }
};
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
    for (const row of xml.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)) { const rowNumber = Number(xmlAttr(row[1], 'r', String(maxRow + 1))); maxRow = Math.max(maxRow, rowNumber); const height = Number(xmlAttr(row[1], 'ht')); if (height) rowHeights.set(rowNumber, height); if (xmlAttr(row[1], 'hidden') === '1') hiddenRows.add(rowNumber); for (const cell of row[2].matchAll(spreadsheetCellPattern)) { const ref = xmlAttr(cell[1], 'r'); const coordinate = ref.match(/^([A-Z]+)(\d+)$/i); if (!coordinate) continue; const column = excelColumnIndex(coordinate[1]); const rowIndex = Number(coordinate[2]); maxCol = Math.max(maxCol, column); maxRow = Math.max(maxRow, rowIndex); const type = xmlAttr(cell[1], 't'); const style = Number(xmlAttr(cell[1], 's', '0')); const raw = xmlUnescape(cell[2].match(/<v>([\s\S]*?)<\/v>/)?.[1] || ''); const formula = xmlUnescape(cell[2].match(/<f(?:\s[^>]*)?>([\s\S]*?)<\/f>/)?.[1] || ''); let value = raw; if (type === 's') value = shared[Number(raw)] || ''; else if (type === 'inlineStr') value = spreadsheetText(cell[2]); else if (type === 'b') value = raw === '1' ? 'TRUE' : 'FALSE'; else if (!raw && formula) value = `=${formula}`; cells.set(`${rowIndex}:${column}`, { value, style }); } }
    const widths = new Map(); const hiddenCols = new Set(); for (const col of xml.matchAll(/<col\b([^>]*)\/?>(?:<\/col>)?/g)) { const min = Number(xmlAttr(col[1], 'min', '1')) - 1; const max = Number(xmlAttr(col[1], 'max', String(min + 1))) - 1; for (let index = min; index <= max; index++) { const width = Number(xmlAttr(col[1], 'width')); if (width) widths.set(index, Math.max(24, Math.min(420, width * 7.4))); if (xmlAttr(col[1], 'hidden') === '1') hiddenCols.add(index); maxCol = Math.max(maxCol, index); } }
    const mergeStarts = new Map(); const mergeCovered = new Set(); for (const merge of xml.matchAll(/<mergeCell\b[^>]*\bref="([A-Z]+)(\d+):([A-Z]+)(\d+)"[^>]*\/?>(?:<\/mergeCell>)?/gi)) { const c1 = excelColumnIndex(merge[1]); const r1 = Number(merge[2]); const c2 = excelColumnIndex(merge[3]); const r2 = Number(merge[4]); mergeStarts.set(`${r1}:${c1}`, { colspan: c2 - c1 + 1, rowspan: r2 - r1 + 1 }); for (let row = r1; row <= r2; row++) for (let col = c1; col <= c2; col++) if (row !== r1 || col !== c1) mergeCovered.add(`${row}:${col}`); maxRow = Math.max(maxRow, r2); maxCol = Math.max(maxCol, c2); }
    const colgroup = `<col class="row-number-col">${Array.from({ length: maxCol + 1 }, (_, column) => `<col style="width:${Number((widths.get(column) || 88).toFixed(2))}px"${hiddenCols.has(column) ? ' class="hidden-col"' : ''}>`).join('')}`; const header = `<tr><th class="corner"></th>${Array.from({ length: maxCol + 1 }, (_, column) => `<th${hiddenCols.has(column) ? ' class="hidden-col"' : ''}>${excelColumnName(column)}</th>`).join('')}</tr>`;
    const body = Array.from({ length: maxRow }, (_, offset) => { const row = offset + 1; if (hiddenRows.has(row)) return ''; const cellsHtml = Array.from({ length: maxCol + 1 }, (_, column) => { if (hiddenCols.has(column) || mergeCovered.has(`${row}:${column}`)) return ''; const cell = cells.get(`${row}:${column}`) || { value: '', style: 0 }; const merge = mergeStarts.get(`${row}:${column}`); const attrs = `${merge?.colspan > 1 ? ` colspan="${merge.colspan}"` : ''}${merge?.rowspan > 1 ? ` rowspan="${merge.rowspan}"` : ''}`; return `<td${attrs} style="${styles[cell.style] || ''}" title="${xmlEscape(cell.value)}">${xmlEscape(cell.value).replace(/\r?\n/g, '<br>')}</td>`; }).join(''); return `<tr${rowHeights.has(row) ? ` style="height:${Math.max(18, rowHeights.get(row) * 1.34)}px"` : ''}><th class="row-number">${row}</th>${cellsHtml}</tr>`; }).join(''); return `<section class="sheet"><h2>${xmlEscape(sheet.name)}</h2><div class="sheet-scroll"><table><colgroup>${colgroup}</colgroup><thead>${header}</thead><tbody>${body}</tbody></table></div></section>`; };
  const rendered = (sheets.length ? sheets : [{ name:'工作表1', path:'xl/worksheets/sheet1.xml' }]).map(renderSheet).join(''); return `<!doctype html><html><head><meta charset="utf-8"><style>*{box-sizing:border-box}body{margin:0;padding:20px;background:#eef1f5;color:#202124;font:12px/1.35 -apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif}.sheet{margin:0 auto 22px;padding:14px;background:#fff;border-radius:8px;box-shadow:0 2px 12px #0002}.sheet h2{margin:0 0 10px;font-size:13px;font-weight:600}.sheet-scroll{max-width:100%;max-height:calc(100vh - 90px);overflow:scroll;border:1px solid #b9c0ca}table{border-collapse:collapse;table-layout:fixed;width:max-content;min-width:100%}th,td{height:24px;border:1px solid #d9dde3;padding:3px 5px;overflow:hidden;white-space:pre;text-overflow:ellipsis}thead th,.row-number{position:sticky;background:#f3f5f7;color:#596273;font-weight:400;text-align:center}.row-number{left:0;z-index:2;width:42px}.corner{left:0;z-index:3;width:42px}.row-number-col{width:42px}.hidden-col{display:none}td{background:#fff;vertical-align:bottom}</style></head><body>${rendered}</body></html>`;
};
export async function previewRecord(id, record) { const item = (await readCatalog()).find(entry => entry.id === id); if (!item) throw new Error('模板不存在'); const extension = item.extension || extensionOf(item.fileName); const bytes = await fs.readFile(path.join(templatesDir, item.fileName)); if (extension === '.docx') return { template: item, kind: 'docx', bytes: renderDocx(bytes, templateScope(item, record)) }; if (extension === '.xlsx') return { template: item, kind: 'html', html: xlsxPreviewHtml(renderXlsx(bytes, record, item)) }; if (['.txt', '.layout', '.json'].includes(extension)) return { template: item, kind: 'html', html: renderLayoutHtml(bytes, record).toString('utf8') }; throw new Error('该模板格式暂不支持记录预览'); }
const matchingValue = (object, key) => { if (!object || typeof object !== 'object') return undefined; if (Object.hasOwn(object, key)) return object[key]; const normalized = normalizeKey(key); const match = Object.keys(object).find(candidate => legacyPlaceholderName(candidate) === key) ?? Object.keys(object).find(candidate => normalizeKey(candidate) === normalized); return match === undefined ? undefined : object[match]; };
const layoutRawValue = (names, scope, rowContext) => {
  const path = (names || []).map(cleanLayoutName);
  if (path[0] === '#') return String((rowContext?.index ?? 0) + 1);
  if (rowContext && normalizeKey(path[0]) === normalizeKey(rowContext.root)) return matchingValue(rowContext.item, path.at(-1));
  let value = matchingValue(scope, path[0]);
  for (const key of path.slice(1)) value = matchingValue(value, key);
  return value;
};
const layoutValue = (names, scope, rowContext) => valueText(layoutRawValue(names, scope, rowContext));
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
  const pageSize = adaptive ? 'auto' : `${Number(setting.width || 210)}mm ${Number(setting.height || 297)}mm`; const pageWidth = adaptive ? '100%' : `${Number(setting.width || 210)}mm`; const pageHeight = adaptive ? 'auto' : `${Number(setting.height || 297)}mm`; return Buffer.from(`<!doctype html><html><head><meta charset="utf-8"><style>@page{size:${pageSize};margin:0}*{box-sizing:border-box}html,body{margin:0;color:#111;font-family:"Microsoft YaHei","SimSun",sans-serif;font-size:${Math.max(10, Number(layout.settings?.fontSize || 10))}pt}.page{width:${pageWidth};min-height:${pageHeight};padding:${Number(setting.paddingTop || 5)}mm ${Number(setting.paddingRight || 13.5)}mm ${Number(setting.paddingBottom || 5)}mm ${Number(setting.paddingLeft || 13.5)}mm;page-break-after:always;background:#fff}.page:last-child{page-break-after:auto}.row{display:flex;width:100%}.column{min-width:0}table{width:100%;border-collapse:collapse;table-layout:fixed}td{border:1px solid #222;padding:3px 4px;overflow-wrap:anywhere}tr{break-inside:avoid}p{min-height:1em}</style></head><body>${pages}</body></html>`, 'utf8');
};

const mmToPt = value => Number(value || 0) * 72 / 25.4;
const parsePt = value => { const match = String(value || '').match(/([\d.]+)\s*pt/i); return match ? Number(match[1]) : Number(value) || 9; };
const pdfFont = async () => {
  const preferred = /(?:cjk|chinese|noto|wqy|uming|simsun|simhei|yahei|deng)/i;
  const roots = ['/usr/share/fonts', '/usr/local/share/fonts', '/opt/fonts', 'C:\\Windows\\Fonts'];
  const candidates = [process.env.FEIYE_PDF_FONT, '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttf', '/usr/share/fonts/cjkuni-uming/uming.ttc', '/usr/share/fonts/truetype/arphic/uming.ttf', '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc', '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttf', 'C:\\Windows\\Fonts\\Deng.ttf', 'C:\\Windows\\Fonts\\simsunb.ttf', 'C:\\Windows\\Fonts\\simsun.ttc', 'C:\\Windows\\Fonts\\SimsunExtG.ttf'];
  for (const root of roots) { try { for (const entry of await fs.readdir(root, { withFileTypes: true, recursive: true })) { if (!entry.isFile() || !preferred.test(entry.name) || !/\.(?:ttf|otf|ttc)$/i.test(entry.name)) continue; candidates.push(path.join(entry.parentPath || root, entry.name)); } } catch {} }
  for (const candidate of [...new Set(candidates)]) if (candidate) try {
    await fs.access(candidate);
    const loaded = fontkit.openSync(candidate);
    if (typeof loaded.createSubset === 'function') return { path: candidate };
    const member = loaded.fonts?.find(font => typeof font.createSubset === 'function');
    if (member?.postscriptName) return { path: candidate, family: member.postscriptName };
  } catch {}
  return null;
};
const layoutPlain = (nodes, scope, rowContext) => (nodes || []).map(node => node.type === 'variable' ? layoutValue(node.name, scope, rowContext) : node.text || layoutPlain(node.children, scope, rowContext)).join('');
const layoutPdf = async (bytes, record, outputPath) => {
  const layout = decodeLayout(bytes).content;
  const scope = templateScope({ fields: layoutFields(layout) }, record);
  const setting = layout.pageSetting || {};
  const font = await pdfFont();
  if (!font) throw new Error('未找到中文 PDF 字体，请配置 FEIYE_PDF_FONT');
  const pageWidth = mmToPt(setting.width || 210), pageHeight = mmToPt(setting.height || 297);
  const top = mmToPt(setting.paddingTop ?? 5), bottom = pageHeight - mmToPt(setting.paddingBottom ?? 5);
  const doc = new PDFDocument({ autoFirstPage: false, margin: 0 });
  doc.font(font.path, font.family);
  const chunks = [];
  doc.on('data', chunk => chunks.push(chunk));
  const done = new Promise((resolve, reject) => { doc.once('end', () => resolve(Buffer.concat(chunks))); doc.once('error', reject); });
  const addPage = () => doc.addPage({ size: [pageWidth, pageHeight], margin: 0 });
  const gap = (value, size) => { const n = parseFloat(value); return Number.isFinite(n) ? /em$/.test(String(value)) ? n * size : /mm$/.test(String(value)) ? mmToPt(n) : n : 0; };
  const paragraphs = (nodes, width, context) => (nodes || []).map(p => {
    const size = Math.max(9, parsePt(p.children?.find(c => c.fontSize)?.fontSize || layout.settings?.fontSize || 10));
    const value = layoutPlain(p.children, scope, context);
    const options = { width, align: ['left','right','center','justify'].includes(p.align) ? p.align : 'left', lineGap: size * .25 };
    doc.fontSize(size);
    return { value, size, options, height: Math.max(size * 1.35, doc.heightOfString(value || ' ', options)), after: gap(p.marginBottom, size) };
  });
  const contentHeight = items => items.reduce((n, p) => n + p.height + p.after, 0);
  const paint = (items, x, y) => {
    for (const p of items) {
      doc.fontSize(p.size).fillColor('#111').text(p.value, x, y, { ...p.options, height: p.height + .5 });
      y += p.height + p.after;
    }
    return y;
  };
  const drawTable = (table, x, y, width) => {
    const weights = (table.columns || []).map(c => Number(c.width) || 1);
    const total = weights.reduce((a,b) => a+b, 0);
    if (!total) return y;
    const columns = weights.map(w => width * w / total);
    const dynamic = new Map((table.dynamicRows || []).map(r => [Number(r.rowIndex), cleanLayoutName(r.dataSource?.rootPath?.[0])]));
    const rows = (table.rows || []).flatMap((row, source) => {
      const root = dynamic.get(source), values = root && matchingValue(scope, root);
      return Array.isArray(values) ? values.map((item,index) => ({ row, source, context: {root,item,index} })) : [{row,source,context:null}];
    });
    const heights = rows.map(() => 18), cells = [];
    rows.forEach(({row,source,context}, r) => {
      (row.cells || []).forEach((cell,c) => {
        if (c >= columns.length) return;
        const merge = (table.merges || []).find(m => source >= m.rowIndex && source < m.rowIndex + m.rowSpan && c >= m.colIndex && c < m.colIndex + m.colSpan);
        if (merge && (merge.rowIndex !== source || merge.colIndex !== c)) return;
        const span = merge?.colSpan || 1;
        const lastSource = source + (merge?.rowSpan || 1) - 1;
        let end = r;
        if (merge?.rowSpan > 1) while (end + 1 < rows.length && rows[end + 1].source <= lastSource) end++;
        const w = columns.slice(c,c+span).reduce((a,b) => a+b, 0);
        const items = paragraphs(cell.content, Math.max(1,w-8), context);
        const needed = contentHeight(items) + 8;
        if (end === r) heights[r] = Math.max(heights[r], needed);
        cells.push({r,end,x:x+columns.slice(0,c).reduce((a,b)=>a+b,0),width:w,items,needed,background:cell.background});
      });
    });
    for (const cell of cells) {
      const available = heights.slice(cell.r,cell.end+1).reduce((a,b)=>a+b,0);
      if (available < cell.needed) heights[cell.end] += cell.needed - available;
    }
    for (let r=0;r<rows.length;) {
      let end=r;
      for (let i=r;i<=end;i++) for (const cell of cells) if (cell.r===i) end=Math.max(end,cell.end);
      const height=heights.slice(r,end+1).reduce((a,b)=>a+b,0);
      if (height > bottom-top) throw new Error('单个合并单元格内容超过一页，请拆分该模板单元格');
      if (y+height>bottom) {addPage();y=top;}
      for (const cell of cells.filter(c=>c.r>=r && c.r<=end)) {
        const cy=y+heights.slice(r,cell.r).reduce((a,b)=>a+b,0);
        const h=heights.slice(cell.r,cell.end+1).reduce((a,b)=>a+b,0);
        if (/^#[0-9a-f]{3,8}$/i.test(cell.background || '')) doc.rect(cell.x,cy,cell.width,h).fill(cell.background);
        doc.lineWidth(.5).strokeColor('#222').rect(cell.x,cy,cell.width,h).stroke();
        paint(cell.items,cell.x+4,cy+4);
      }
      y+=height;r=end+1;
    }
    return y;
  };
  for (const page of layout.document.pages || []) {
    addPage();
    let y=top;
    const left=mmToPt(setting.paddingLeft ?? 13.5), usable=pageWidth-left-mmToPt(setting.paddingRight ?? 13.5);
    for (const row of page.rows || []) {
      let x=left, rowBottom=y;
      for (const column of row.columns || []) {
        const width=usable*Number(column.width || 100)/100;
        let cy=y;
        for (const block of column.blocks || []) {
          cy+=gap(block.marginTop,9);
          if (Number(block.type)===4) cy=drawTable(block.table || {},x,cy,width);
          else for (const p of paragraphs(block.content,width,null)) {
            if(cy+p.height+p.after>bottom){addPage();cy=top;}
            cy=paint([p],x,cy);
          }
        }
        rowBottom=Math.max(rowBottom,cy);x+=width;
      }
      y=rowBottom;
    }
  }
  doc.end(); const output=await done; await fs.writeFile(outputPath,output); return output;
};
const wordText = value => String(value ?? '').split(/\r?\n/).map((part, index) => `${index ? '<w:br/>' : ''}<w:t xml:space="preserve">${xmlEscape(part)}</w:t>`).join('');
const emptyWordParagraph = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/></w:pPr></w:p>';
const wordRun = (node, scope, context, defaultFontSize) => {
  const value = node?.type === 'variable' ? layoutValue(node.name, scope, context) : node?.text || '';
  const format = node?.children?.find(child => child?.fontSize || child?.bold || child?.italic) || node;
  const size = Math.max(9, parsePt(format?.fontSize || defaultFontSize || 10));
  const properties = `<w:rFonts w:ascii="Microsoft YaHei" w:hAnsi="Microsoft YaHei" w:eastAsia="Microsoft YaHei"/><w:sz w:val="${Math.round(size * 2)}"/><w:szCs w:val="${Math.round(size * 2)}"/>${format?.bold ? '<w:b/>' : ''}${format?.italic ? '<w:i/>' : ''}`;
  return `<w:r><w:rPr>${properties}</w:rPr>${wordText(value)}</w:r>`;
};
const wordParagraph = (paragraph, scope, context, defaultFontSize, before = 0) => {
  const align = ['left', 'right', 'center', 'justify'].includes(paragraph?.align) ? paragraph.align : 'left';
  const size = Math.max(9, parsePt(paragraph?.children?.find(child => child?.fontSize)?.fontSize || defaultFontSize || 10));
  const line = Math.round((Number(paragraph?.lineHeight) || 1.5) * size * 20);
  const after = Math.round(Math.max(0, parseFloat(paragraph?.marginBottom) || 0) * 20);
  const runs = (paragraph?.children || []).map(node => wordRun(node, scope, context, defaultFontSize)).join('');
  return `<w:p><w:pPr><w:jc w:val="${align}"/><w:spacing w:line="${line}" w:lineRule="auto" w:before="${Math.max(0, Math.round(before * 20))}" w:after="${after}"/></w:pPr>${runs || '<w:r><w:t></w:t></w:r>'}</w:p>`;
};
const layoutToWord = (bytes, record) => {
  const layout = decodeLayout(bytes).content; const scope = templateScope({ fields: layoutFields(layout) }, record); const setting = layout.pageSetting || {};
  const defaultFontSize = layout.settings?.fontSize || 10; const firstColumns = layout.document.pages?.[0]?.rows?.[0]?.columns || []; const firstBlock = firstColumns[0]?.blocks?.[0]; const firstOffset = firstColumns.length === 1 ? Math.min(0, Number(firstBlock?.marginTop) || 0) : 0;
  const pageWidth = Math.round(Number(setting.width || 210) * 1440 / 25.4), pageHeight = Math.round(Number(setting.height || 297) * 1440 / 25.4);
  const left = Math.round(Number(setting.paddingLeft ?? 13.5) * 1440 / 25.4), right = Math.round(Number(setting.paddingRight ?? 13.5) * 1440 / 25.4);
  const contentWidth = Math.max(720, pageWidth - left - right);
  const paragraphs = (nodes, context, before = 0) => (nodes || []).map((paragraph, index) => wordParagraph(paragraph, scope, context, defaultFontSize, index ? 0 : before)).join('');
  const tableXml = table => {
    const weights = (table.columns || []).map(column => Number(column.width) || 1); const total = weights.reduce((sum, value) => sum + value, 0) || 1;
    const widths = weights.map(value => Math.max(120, Math.round(contentWidth * value / total)));
    const dynamic = new Map((table.dynamicRows || []).map(item => [Number(item.rowIndex), cleanLayoutName(item.dataSource?.rootPath?.[0])]));
    const rows = (table.rows || []).flatMap((row, source) => { const root = dynamic.get(source), values = root && matchingValue(scope, root); return Array.isArray(values) && values.length ? values.map((item, index) => ({ row, source, context: { root, item, index } })) : [{ row, source, context: null }]; });
    const rowXml = ({ row, source, context }) => {
      const cells = (row.cells || []).map((cell, column) => {
        const merge = (table.merges || []).find(item => source >= item.rowIndex && source < item.rowIndex + item.rowSpan && column >= item.colIndex && column < item.colIndex + item.colSpan);
        if (merge && column !== merge.colIndex) return '';
        const span = merge?.colSpan || 1, width = widths.slice(column, column + span).reduce((sum, value) => sum + value, 0);
        const vertical = merge?.rowSpan > 1 ? `<w:vMerge w:val="${source === merge.rowIndex ? 'restart' : 'continue'}"/>` : '';
        const content = merge && source !== merge.rowIndex ? emptyWordParagraph : paragraphs(cell.content, context);
        const shade = /^#[0-9a-f]{6}$/i.test(cell.background || '') ? `<w:shd w:fill="${cell.background.slice(1)}"/>` : '';
        return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ''}${vertical}${shade}<w:vAlign w:val="${cell.verticalAlign === 'middle' ? 'center' : cell.verticalAlign === 'bottom' ? 'bottom' : 'top'}"/></w:tcPr>${content || emptyWordParagraph}</w:tc>`;
      }).join('');
      return `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cells}</w:tr>`;
    };
    return `<w:tbl><w:tblPr><w:tblW w:w="${contentWidth}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="40" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="40" w:type="dxa"/></w:tblCellMar><w:tblBorders><w:top w:val="single" w:sz="4" w:color="222222"/><w:left w:val="single" w:sz="4" w:color="222222"/><w:bottom w:val="single" w:sz="4" w:color="222222"/><w:right w:val="single" w:sz="4" w:color="222222"/><w:insideH w:val="single" w:sz="4" w:color="222222"/><w:insideV w:val="single" w:sz="4" w:color="222222"/></w:tblBorders></w:tblPr><w:tblGrid>${widths.map(width => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>${rows.map(rowXml).join('')}</w:tbl>`;
  };
  const blocks = column => (column.blocks || []).map(block => Number(block.type) === 4 ? tableXml(block.table || {}) : paragraphs(block.content, null, Math.max(0, Number(block.marginTop) || 0))).join('');
  const pageRows = page => (page.rows || []).map(row => {
    const columns = row.columns || []; if (columns.length <= 1) return columns.map(blocks).join('');
    const widths = columns.map(column => Math.max(120, Math.round(contentWidth * Number(column.width || 100 / columns.length) / 100)));
    return `<w:tbl><w:tblPr><w:tblW w:w="${contentWidth}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders></w:tblPr><w:tblGrid>${widths.map(width => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid><w:tr>${columns.map((column, index) => `<w:tc><w:tcPr><w:tcW w:w="${widths[index]}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>${blocks(column) || '<w:p/>'}</w:tc>`).join('')}</w:tr></w:tbl>`;
  }).join('');
  const body = (layout.document.pages || []).map((page, index, pages) => `${pageRows(page)}${index < pages.length - 1 ? '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' : ''}`).join('');
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="${pageWidth}" w:h="${pageHeight}"${pageWidth > pageHeight ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="${Math.max(0, Math.round(Number(setting.paddingTop ?? 5) * 1440 / 25.4 + firstOffset * 20))}" w:right="${right}" w:bottom="${Math.round(Number(setting.paddingBottom ?? 5) * 1440 / 25.4)}" w:left="${left}" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  const zip = new PizZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', document);
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
};
const layoutToXlsx = (bytes, record) => {
  const layout = decodeLayout(bytes).content; const scope = templateScope({ fields: layoutFields(layout) }, record); const zip = new PizZip(); const rows = []; const merges = []; const pageBreaks = []; const setting = layout.pageSetting || {}; const defaultFontSize = Math.max(9, parsePt(layout.settings?.fontSize || 10));
  const columnNeed = column => Math.max(1, ...(column.blocks || []).filter(block => Number(block.type) === 4).map(block => block.table?.columns?.length || 1));
  const maxColumns = Math.max(1, ...(layout.document.pages || []).flatMap(page => (page.rows || []).map(row => (row.columns || []).reduce((sum, column) => sum + columnNeed(column), 0)))); const columnWidths = Array(maxColumns).fill(0);
  const columnSpans = columns => {
    const spans = columns.map(columnNeed), totalWidth = columns.reduce((sum, column) => sum + Number(column.width || 100 / Math.max(1, columns.length)), 0) || 100;
    while (spans.reduce((sum, value) => sum + value, 0) < maxColumns) {
      const index = spans.reduce((best, value, current) => maxColumns * Number(columns[current].width || 100 / columns.length) / totalWidth - value > maxColumns * Number(columns[best].width || 100 / columns.length) / totalWidth - spans[best] ? current : best, 0);
      spans[index]++;
    }
    return spans;
  };
  const fonts = [{ size: defaultFontSize, bold: false, color: '000000' }], fills = ['', 'gray125'], styles = [{ fontId: 0, fillId: 0, borderId: 0, align: '', vertical: 'top' }];
  const fontId = font => { const value = { size: Math.max(9, Number(font.size) || defaultFontSize), bold: Boolean(font.bold), color: String(font.color || '#000000').replace('#', '').toUpperCase() }; let index = fonts.findIndex(item => JSON.stringify(item) === JSON.stringify(value)); if (index < 0) { index = fonts.length; fonts.push(value); } return index; };
  const fillId = color => { const value = /^#[0-9a-f]{6}$/i.test(color || '') ? color.slice(1).toUpperCase() : ''; if (!value) return 0; let index = fills.indexOf(value); if (index < 0) { index = fills.length; fills.push(value); } return index; };
  const styleId = ({ font = {}, fill = '', border = false, align = '', vertical = 'top' }) => { const value = { fontId: fontId(font), fillId: fillId(fill), borderId: border ? 1 : 0, align: ['left','right','center','justify'].includes(align) ? align : '', vertical: vertical === 'middle' ? 'center' : vertical === 'bottom' ? 'bottom' : 'top' }; let index = styles.findIndex(item => JSON.stringify(item) === JSON.stringify(value)); if (index < 0) { index = styles.length; styles.push(value); } return index; };
  const paragraphStyle = (paragraph, cell = {}, border = false) => { const children = (paragraph?.children || []).flatMap(child => [child, ...(child.children || [])]); const sized = children.map(child => parsePt(child.fontSize)).filter(Boolean); const colored = children.find(child => /^#[0-9a-f]{6}$/i.test(child.color || '')); return styleId({ font: { size: sized.length ? Math.max(...sized) : defaultFontSize, bold: children.some(child => child.bold), color: colored?.color }, fill: cell.background, border, align: paragraph?.align, vertical: cell.verticalAlign }); };
  const ensureRow = index => { while (rows.length <= index) rows.push({ cells: new Map(), height: 15 }); return rows[index]; };
  const addCell = (row, column, value, style, numeric = false) => ensureRow(row).cells.set(column, { value: valueText(value).replace(/\r\n/g, '\n'), style, numeric: Boolean(numeric && /^-?(?:\d+\.?\d*|\.\d+)$/.test(valueText(value).trim())) });
  const addMerge = (rowStart, columnStart, rowEnd, columnEnd) => { if (rowEnd > rowStart || columnEnd > columnStart) merges.push({ rowStart, columnStart, rowEnd, columnEnd }); };
  const paragraphText = (paragraph, context) => layoutPlain(paragraph?.children, scope, context).replace(/\r\n/g, '\n');
  const visualLength = value => [...String(value || '')].reduce((length, character) => length + (/[^\x00-\xff]/.test(character) ? 2 : 1), 0);
  const wrappedLineCount = (value, width, size) => String(value || '').split('\n').reduce((count, line) => count + Math.max(1, Math.ceil(visualLength(line) / Math.max(1, width * 11 / Math.max(1, size)))), 0);
  const requiredRowHeight = (value, size, width) => Math.max(15, size * 1.5 * wrappedLineCount(value, width, size) + 3);
  const isNumericParagraph = (paragraph, context) => { const children = paragraph?.children || []; return children.length === 1 && children[0].type === 'variable' && !String(children[0].name?.[0] || '').startsWith('$index') && numericScalar(layoutRawValue(children[0].name, scope, context)) !== null; };
  const renderParagraphs = (paragraphs, startRow, startColumn, span) => { let row = startRow; for (const paragraph of paragraphs || []) { const value = paragraphText(paragraph, null); addCell(row, startColumn, value, paragraphStyle(paragraph), isNumericParagraph(paragraph, null)); addMerge(row, startColumn, row, startColumn + span - 1); const size = Math.max(defaultFontSize, ...(paragraph.children || []).map(child => parsePt(child.fontSize))); const width = columnWidths.slice(startColumn, startColumn + span).reduce((sum, item) => sum + item, 0); ensureRow(row).height = Math.max(ensureRow(row).height, requiredRowHeight(value, size, width)); row++; } return row; };
  const renderTable = (table, startRow, startColumn, span, regionWidth) => {
    const sourceRows = table.rows || []; const dynamic = new Map((table.dynamicRows || []).map(item => [Number(item.rowIndex), cleanLayoutName(item.dataSource?.rootPath?.[0])])); const entries = [];
    for (const [source, templateRow] of sourceRows.entries()) { const root = dynamic.get(source), values = root && matchingValue(scope, root); const contexts = Array.isArray(values) && values.length ? values.map((item, index) => ({ root, item, index })) : [null]; for (const context of contexts) entries.push({ source, templateRow, context }); }
    const tableColumns = Math.max(1, table.columns?.length || sourceRows[0]?.cells?.length || 1); const weights = Array.from({ length: tableColumns }, (_, index) => Number(table.columns?.[index]?.width) || 1); const totalWeight = weights.reduce((sum, value) => sum + value, 0) || 1;
    const heightNeeds = [];
    for (const [targetOffset, entry] of entries.entries()) {
      const rowIndex = startRow + targetOffset;
      for (let column = 0; column < Math.min(span, tableColumns); column++) {
        const sourceCell = entry.templateRow.cells?.[column] || {};
        addCell(rowIndex, startColumn + column, '', paragraphStyle(sourceCell.content?.[0], sourceCell, true));
      }
      for (const [column, cell] of (entry.templateRow.cells || []).entries()) {
        if (column >= span) break;
        const covering = (table.merges || []).find(merge => entry.source >= merge.rowIndex && entry.source < merge.rowIndex + merge.rowSpan && column >= merge.colIndex && column < merge.colIndex + merge.colSpan);
        if (covering && (entry.source !== covering.rowIndex || column !== covering.colIndex)) continue;
        const paragraphs = cell.content || []; const value = paragraphs.map(paragraph => paragraphText(paragraph, entry.context)).join('\n');
        addCell(rowIndex, startColumn + column, value, paragraphStyle(paragraphs[0], cell, true), paragraphs.length === 1 && isNumericParagraph(paragraphs[0], entry.context));
        const size = Math.max(defaultFontSize, ...paragraphs.flatMap(paragraph => (paragraph.children || []).map(child => parsePt(child.fontSize))));
        const columnSpan = Math.min(span - column, covering?.colSpan || 1);
        const width = columnWidths.slice(startColumn + column, startColumn + column + columnSpan).reduce((sum, item) => sum + item, 0);
        let endOffset = targetOffset;
        if (covering?.rowSpan > 1) while (endOffset + 1 < entries.length && entries[endOffset + 1].source < covering.rowIndex + covering.rowSpan) endOffset++;
        heightNeeds.push({ start: targetOffset, end: endOffset, height: Math.max(18, requiredRowHeight(value, size, width)) });
      }
    }
    for (const need of heightNeeds) {
      const available = Array.from({ length: need.end - need.start + 1 }, (_, index) => ensureRow(startRow + need.start + index).height).reduce((sum, height) => sum + height, 0);
      if (available < need.height) ensureRow(startRow + need.end).height += need.height - available;
    }
    for (const merge of table.merges || []) { const starts = entries.map((entry, index) => entry.source === merge.rowIndex ? index : -1).filter(index => index >= 0); for (const offset of starts) { let end = offset; while (end + 1 < entries.length && entries[end + 1].source < merge.rowIndex + merge.rowSpan) end++; addMerge(startRow + offset, startColumn + merge.colIndex, startRow + end, startColumn + Math.min(span, merge.colIndex + merge.colSpan) - 1); } }
    return startRow + entries.length;
  };
  const printableWidth = Math.max(40, (Number(setting.width || 210) - Number(setting.paddingLeft ?? 5) - Number(setting.paddingRight ?? 5)) / 1.85);
  const widthPriority = Array(maxColumns).fill(0);
  const setColumnWidths = (start, widths, priority) => widths.forEach((width, offset) => { const index = start + offset; if (priority > widthPriority[index]) { columnWidths[index] = width; widthPriority[index] = priority; } });
  for (const page of layout.document.pages || []) for (const pageRow of page.rows || []) {
    const columns = pageRow.columns || [], spans = columnSpans(columns), totalWidth = columns.reduce((sum, column) => sum + Number(column.width || 100 / Math.max(1, columns.length)), 0) || 100; let startColumn = 0;
    for (const [columnIndex, column] of columns.entries()) {
      const span = spans[columnIndex], regionWidth = printableWidth * Number(column.width || 100 / Math.max(1, columns.length)) / totalWidth;
      const table = (column.blocks || []).find(block => Number(block.type) === 4 && (block.table?.columns?.length || 0) === span)?.table;
      const weights = table?.columns?.map(item => Number(item.width) || 1), weightTotal = weights?.reduce((sum, value) => sum + value, 0) || 1;
      setColumnWidths(startColumn, weights ? weights.map(value => regionWidth * value / weightTotal) : Array(span).fill(regionWidth / span), weights ? 2 : 1);
      startColumn += span;
    }
  }
  for (const [pageIndex, page] of (layout.document.pages || []).entries()) { if (pageIndex && rows.length) pageBreaks.push(rows.length); for (const pageRow of page.rows || []) { const columns = pageRow.columns || []; const spans = columnSpans(columns); const totalWidth = columns.reduce((sum, column) => sum + Number(column.width || 100 / Math.max(1, columns.length)), 0) || 100; let startColumn = 0; const baseRow = rows.length; let rowEnd = baseRow; for (const [columnIndex, column] of columns.entries()) { const span = spans[columnIndex]; const regionWidth = printableWidth * Number(column.width || 100 / Math.max(1, columns.length)) / totalWidth; let cursor = baseRow; for (const block of column.blocks || []) { if (Number(block.type) === 4) cursor = renderTable(block.table || {}, cursor, startColumn, span, regionWidth); else cursor = renderParagraphs(block.content, cursor, startColumn, span); } rowEnd = Math.max(rowEnd, cursor); startColumn += span; } while (rows.length < rowEnd) ensureRow(rows.length); } }
  if (!rows.length) ensureRow(0);
  const cell = (entry, column, row) => entry.numeric ? `<c r="${excelColumnName(column)}${row}"${entry.style ? ` s="${entry.style}"` : ''}><v>${xmlEscape(entry.value)}</v></c>` : `<c r="${excelColumnName(column)}${row}"${entry.style ? ` s="${entry.style}"` : ''} t="inlineStr"><is><t xml:space="preserve">${xmlEscape(entry.value)}</t></is></c>`;
  const mergeXml = merges.map(item => `<mergeCell ref="${excelColumnName(item.columnStart)}${item.rowStart + 1}:${excelColumnName(item.columnEnd)}${item.rowEnd + 1}"/>`).join('');
  const sheetRows = rows.map((row, index) => `<row r="${index + 1}" ht="${Number(row.height.toFixed(2))}" customHeight="1">${[...row.cells.entries()].sort(([left], [right]) => left - right).map(([column, entry]) => cell(entry, column, index + 1)).join('')}</row>`).join('');
  const fontXml = fonts.map(font => `<font>${font.bold ? '<b/>' : ''}<sz val="${font.size}"/><color rgb="FF${font.color}"/><name val="Microsoft YaHei"/><family val="2"/></font>`).join(''); const fillXml = fills.map((fill, index) => index === 0 ? '<fill><patternFill patternType="none"/></fill>' : index === 1 ? '<fill><patternFill patternType="gray125"/></fill>' : `<fill><patternFill patternType="solid"><fgColor rgb="FF${fill}"/><bgColor indexed="64"/></patternFill></fill>`).join(''); const xfXml = styles.map(style => `<xf numFmtId="0" fontId="${style.fontId}" fillId="${style.fillId}" borderId="${style.borderId}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment${style.align ? ` horizontal="${style.align}"` : ''} vertical="${style.vertical}" wrapText="1"/></xf>`).join('');
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="数据" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  zip.file('xl/styles.xml', `<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="${fonts.length}">${fontXml}</fonts><fills count="${fills.length}">${fillXml}</fills><borders count="2"><border/><border><left style="thin"><color rgb="FF222222"/></left><right style="thin"><color rgb="FF222222"/></right><top style="thin"><color rgb="FF222222"/></top><bottom style="thin"><color rgb="FF222222"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${styles.length}">${xfXml}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`);
  const cols = columnWidths.map((width, index) => `<col min="${index + 1}" max="${index + 1}" width="${Number(Math.max(3, Math.min(60, width)).toFixed(2))}" customWidth="1"/>`).join(''); const breaks = pageBreaks.length ? `<rowBreaks count="${pageBreaks.length}" manualBreakCount="${pageBreaks.length}">${pageBreaks.map(row => `<brk id="${row}" min="0" max="16383" man="1"/>`).join('')}</rowBreaks>` : '';
  zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:${excelColumnName(Math.max(0, maxColumns - 1))}${Math.max(1, rows.length)}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols>${cols}</cols><sheetData>${sheetRows}</sheetData>${mergeXml ? `<mergeCells count="${merges.length}">${mergeXml}</mergeCells>` : ''}${breaks}<pageMargins left="${Number(setting.paddingLeft ?? 5) / 25.4}" right="${Number(setting.paddingRight ?? 5) / 25.4}" top="${Number(setting.paddingTop ?? 5) / 25.4}" bottom="${Number(setting.paddingBottom ?? 5) / 25.4}" header="0" footer="0"/><pageSetup paperSize="9" orientation="${Number(setting.width || 210) > Number(setting.height || 297) ? 'landscape' : 'portrait'}" fitToWidth="1" fitToHeight="1"/></worksheet>`);
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
};

export async function renderTemplate(id, records, requestedFormat = '', keepFiles = false) {
  const item = (await listTemplates()).find(entry => entry.id === id); if (!item) throw new Error('模板不存在');
  if (!Array.isArray(records) || records.length === 0) throw new Error('至少需要一条记录');
  const bytes = await fs.readFile(path.join(templatesDir, item.fileName)); const extension = item.extension || extensionOf(item.fileName); const outputId = crypto.randomUUID(); await fs.mkdir(outputsDir, { recursive: true });
  const isLayout = ['.txt', '.layout', '.json'].includes(extension); const requested = String(requestedFormat || '').toLowerCase(); const defaultFormat = isLayout ? 'pdf' : extension.slice(1); const format = requested === 'word' || requested === 'doc' ? 'docx' : requested || defaultFormat;
  if (!isLayout && format !== extension.slice(1)) throw new Error(`当前 ${extension.slice(1).toUpperCase()} 模板仅支持生成同格式文件`);
  if (isLayout && !['pdf', 'docx', 'xlsx'].includes(format)) throw new Error('在线模板支持生成 PDF、Word 或 Excel');
  const render = async (record, index) => {
    if (extension === '.docx') return renderDocx(bytes, templateScope(item, record));
    if (extension === '.xlsx') return renderXlsx(bytes, record, item);
    if (!isLayout) return bytes;
    if (format === 'pdf') return layoutPdf(bytes, record, path.join(outputsDir, `${outputId}-${index}.pdf`));
    if (format === 'docx') return layoutToWord(bytes, record);
    return layoutToXlsx(bytes, record);
  };
  const outputExtension = `.${format}`;
  const outputBaseName = record => {
    const rendered = item.outputNamePattern ? renderTemplateText(item.outputNamePattern, templateScope(item, record)).trim() : '';
    return rendered ? cleanName(rendered).replace(/\.(?:docx?|xlsx?|pdf|zip)$/i, '') : `${path.parse(item.name).name}-${new Date().toISOString().slice(0, 10)}`;
  };
  if (records.length === 1) { const fileName = `${outputId}${outputExtension}`; const name = `${outputBaseName(records[0])}${outputExtension}`; const generated = await render(records[0], 0); await fs.writeFile(path.join(outputsDir, fileName), generated); if (format === 'pdf') await fs.rm(path.join(outputsDir, `${outputId}-0.pdf`), { force: true }); return { id: outputId, fileName, extension: format, name, files: [{ id: outputId, extension: format, name, recordIndex: 0 }] }; }
  const archive = new PizZip(); const files = []; for (const [index, record] of records.entries()) { const generated = await render(record, index); const id = crypto.randomUUID(); const name = `${outputBaseName(record)}${outputExtension}`; if (keepFiles) { await fs.writeFile(path.join(outputsDir, `${id}${outputExtension}`), generated); files.push({ id, extension: format, name, recordIndex: index }); } archive.file(`${String(index + 1).padStart(3, '0')}-${name}`, generated); if (format === 'pdf') await fs.rm(path.join(outputsDir, `${outputId}-${index}.pdf`), { force: true }); } const fileName = `${outputId}.zip`; await fs.writeFile(path.join(outputsDir, fileName), archive.generate({ type: 'nodebuffer', compression: 'DEFLATE' })); return { id: outputId, fileName, extension: 'zip', name: `${outputBaseName(records[0])}-${records.length}份.zip`, files };
}

export const templateFile = (id, extension = '.docx') => path.join(templatesDir, `${id}${extension.startsWith('.') ? extension : `.${extension}`}`);
export const outputFile = (id, extension = 'docx') => path.join(outputsDir, `${id}.${extension}`);
