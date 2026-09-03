import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs/promises';
import PizZip from 'pizzip';
import { saveTemplate, renameTemplate, renderTemplate, previewTemplate, removeTemplate, outputFile, listTemplates } from '../lib/template-store.mjs';

test('imports and exports XLSX, XLS and PDF formats', async () => {
  const xlsx = new PizZip(); xlsx.file('xl/workbook.xml', '<workbook/>'); xlsx.file('xl/worksheets/sheet1.xml', '<sheet><c>{编号}</c></sheet>');
  const cases = [['sample.xlsx', xlsx.generate({ type: 'nodebuffer' })], ['sample.xls', Buffer.from('xls')], ['sample.pdf', Buffer.from('%PDF-1.4')]];
  const saved = []; const outputs = [];
  try {
    for (const [name, bytes] of cases) { const item = await saveTemplate(name, bytes, { baseId: 'base', tableId: 'table' }); saved.push(item); const output = await renderTemplate(item.id, [{ 编号: 'A-1' }]); outputs.push(output); assert.equal(output.extension, item.extension.slice(1)); }
    assert.equal((await listTemplates({ baseId: 'base', tableId: 'table' })).filter(item => saved.some(savedItem => savedItem.id === item.id)).length, 3);
    assert.equal((await listTemplates({ baseId: 'other', tableId: 'table' })).filter(item => saved.some(savedItem => savedItem.id === item.id)).length, 0);
  } finally { for (const output of outputs) await fs.rm(outputFile(output.id, output.extension), { force: true }); for (const item of saved) await removeTemplate(item.id); }
});

test('imports Feishu online layout export and preserves structured preview', async () => {
  const source = Buffer.from(JSON.stringify({ name: '在线模板', content: JSON.stringify({ document: { pages: [{ rows: [{ columns: [{ width: 100, blocks: [{ type: 1, content: [{ type: 'paragraph', children: [{ text: '客户：' }, { type: 'variable', name: ['🔵客户名称'] }] }] }] }] }] }] } }) })).toString('base64');
  const item = await saveTemplate('online.txt', Buffer.from(source), { baseId: 'layout-base', tableId: 'layout-table' });
  const outputs = [];
  try {
    const preview = await previewTemplate(item.id); assert.equal(preview.kind, 'layout'); assert.equal(preview.layout.document.pages.length, 1); assert.deepEqual(preview.fields, [{ marker: '', name: '客户名称' }]);
    const renamed = await renameTemplate(item.id, '客户模板'); assert.equal(renamed.name, '客户模板.txt');
    const record = { '🔵客户名称': '测试客户', '🔵合同明细': [{ '🔴SKU': 'SKU-彩色', '🔴单价': '12.30' }] };
    const pdf = await renderTemplate(item.id, [record]); outputs.push(pdf); assert.equal(pdf.extension, 'pdf'); assert.match((await fs.readFile(outputFile(pdf.id, 'pdf'))).subarray(0, 4).toString(), /%PDF/);
    const word = await renderTemplate(item.id, [record], 'word'); outputs.push(word); assert.equal(word.extension, 'doc'); assert.match((await fs.readFile(outputFile(word.id, 'doc'))).toString('utf16le'), /测试客户/);
    const excel = await renderTemplate(item.id, [record], 'xlsx'); outputs.push(excel); assert.equal(excel.extension, 'xlsx'); assert.match(new PizZip(await fs.readFile(outputFile(excel.id, 'xlsx'))).file('xl/worksheets/sheet1.xml').asText(), /测试客户/);
  } finally { for (const output of outputs) await fs.rm(outputFile(output.id, output.extension), { force: true }); await removeTemplate(item.id); }
});

test('expands XLSX shared-string loop rows', async () => {
  const zip = new PizZip();
  zip.file('xl/workbook.xml', '<workbook/>');
  zip.file('xl/sharedStrings.xml', '<?xml version="1.0"?><sst><si><t>{#明细}{$index+1}</t></si><si><t>{品名}</t></si><si><t>{/明细}</t></si><si><t>尾部</t></si></sst>');
  zip.file('xl/worksheets/sheet1.xml', '<?xml version="1.0"?><worksheet><dimension ref="A1:C2"/><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row><row r="2"><c r="A2" t="s"><v>3</v></c></row></sheetData></worksheet>');
  const item = await saveTemplate('loop.xlsx', zip.generate({ type: 'nodebuffer' })); let output;
  try {
    assert.ok(item.fields.some(field => field.name === '明细' && field.marker === '#'));
    output = await renderTemplate(item.id, [{ 明细: [{ 品名: '产品A' }, { 品名: '产品B' }] }]);
    const sheet = new PizZip(await fs.readFile(outputFile(output.id, 'xlsx'))).file('xl/worksheets/sheet1.xml').asText();
    assert.match(sheet, /r="A1"[^>]*t="inlineStr"/); assert.match(sheet, />1<\/t>/); assert.match(sheet, /产品A/);
    assert.match(sheet, /r="A2"[^>]*t="inlineStr"/); assert.match(sheet, />2<\/t>/); assert.match(sheet, /产品B/);
    assert.match(sheet, /<row r="3"><c r="A3"/); assert.match(sheet, /dimension ref="A1:C3"/);
  } finally { if (output) await fs.rm(outputFile(output.id, output.extension), { force: true }); await removeTemplate(item.id); }
});

test('renders filled Office record previews as HTML', async () => {
  const docx = new PizZip();
  docx.file('[Content_Types].xml', `<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Default Extension='xml' ContentType='application/xml'/><Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/></Types>`);
  docx.file('_rels/.rels', `<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships>`);
  docx.file('word/document.xml', `<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body><w:p><w:r><w:t>客户：{客户}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  const docxItem = await saveTemplate('preview.docx', docx.generate({ type: 'nodebuffer' }));
  const xlsx = new PizZip(); xlsx.file('xl/workbook.xml', '<workbook/>'); xlsx.file('xl/worksheets/sheet1.xml', `<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>{编号}</t></is></c></row></sheetData></worksheet>`);
  const xlsxItem = await saveTemplate('preview.xlsx', xlsx.generate({ type: 'nodebuffer' }));
  try { const docxPreview = await (await import('../lib/template-store.mjs')).previewRecord(docxItem.id, { 客户: '张三' }); assert.equal(docxPreview.kind, 'html'); assert.match(docxPreview.html, /张三/); const xlsxPreview = await (await import('../lib/template-store.mjs')).previewRecord(xlsxItem.id, { 编号: 'A-100' }); assert.equal(xlsxPreview.kind, 'html'); assert.match(xlsxPreview.html, /A-100/); } finally { await removeTemplate(docxItem.id); await removeTemplate(xlsxItem.id); }
});
