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

test('fills namespace-prefixed XLSX string cells stored in v nodes', async () => {
  const zip = new PizZip();
  zip.file('xl/workbook.xml', '<workbook/>');
  zip.file('xl/worksheets/sheet1.xml', '<?xml version="1.0"?><x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData><x:row r="1"><x:c r="A1" t="str"><x:v>订单号：{订单号}</x:v></x:c></x:row><x:row r="2"><x:c r="A2" t="str"><x:v>{#订单明细}{$index+1}</x:v></x:c><x:c r="B2" t="str"><x:v>{产品名称}{/订单明细}</x:v></x:c></x:row></x:sheetData></x:worksheet>');
  const item = await saveTemplate('string-cell.xlsx', zip.generate({ type: 'nodebuffer' })); let output;
  try {
    output = await renderTemplate(item.id, [{ 订单号: 'DD2026001', 订单明细: [{ 产品名称: '无线键盘' }, { 产品名称: '商务显示器' }] }]);
    const sheet = new PizZip(await fs.readFile(outputFile(output.id, 'xlsx'))).file('xl/worksheets/sheet1.xml').asText();
    assert.match(sheet, /订单号：DD2026001/); assert.match(sheet, /无线键盘/); assert.match(sheet, /商务显示器/);
  } finally { if (output) await fs.rm(outputFile(output.id, output.extension), { force: true }); await removeTemplate(item.id); }
});

test('expands XLSX loops spanning multiple template rows', async () => {
  const zip = new PizZip();
  zip.file('xl/workbook.xml', '<workbook/>');
  zip.file('xl/worksheets/sheet1.xml', '<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" t="str"><v>{#采购明细}</v></c></row><row r="2"><c r="A2" t="str"><v>品名：{产品名称}</v></c></row><row r="3"><c r="A3" t="str"><v>数量：{数量}</v></c></row><row r="4"><c r="A4" t="str"><v>{/采购明细}</v></c></row><row r="5"><c r="A5" t="str"><v>结束</v></c></row></sheetData></worksheet>');
  const item = await saveTemplate('multi-row-loop.xlsx', zip.generate({ type: 'nodebuffer' })); let output;
  try {
    output = await renderTemplate(item.id, [{ 采购明细: [{ 产品名称: '钢笔', 数量: 3 }, { 产品名称: '文件夹', 数量: 5 }] }]);
    const sheet = new PizZip(await fs.readFile(outputFile(output.id, 'xlsx'))).file('xl/worksheets/sheet1.xml').asText();
    assert.match(sheet, /品名：钢笔/); assert.match(sheet, /数量：3/); assert.match(sheet, /品名：文件夹/); assert.match(sheet, /数量：5/); assert.match(sheet, /结束/);
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

test('applies procurement template special mappings', async () => {
  const docx = new PizZip();
  docx.file('[Content_Types].xml', `<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Default Extension='xml' ContentType='application/xml'/><Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/></Types>`);
  docx.file('_rels/.rels', `<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships>`);
  docx.file('word/document.xml', `<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body><w:p><w:r><w:t>{__SKU总数}|{合同合计金额_大写_}|{__合同创建时间}</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>{#__合同明细_采购明细}{__开票品名}</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>{SKU__}|{单价_含税___}|{实收数量__}|{总价_含税_}{/__合同明细_采购明细}</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`);
  const item = await saveTemplate('采购合同（含税）-采购明细.docx', docx.generate({ type: 'nodebuffer' })); let output;
  try {
    output = await renderTemplate(item.id, [{ 合同合计金额: 12800, '合同明细-采购明细': [{ 开票品名: '无线键盘', SKU: 'KB-002', '单价(含税)': 240, 实收数量: 20, '总价(含税)': 4800 }, { 开票品名: '显示器', SKU: 'MN-003', '单价(含税)': 800, 实收数量: 10, '总价(含税)': 8000 }] }]);
    const xml = new PizZip(await fs.readFile(outputFile(output.id, 'docx'))).file('word/document.xml').asText();
    assert.match(xml, /KB-002/); assert.match(xml, /MN-003/); assert.match(xml, />30\|/); assert.match(xml, /壹万贰仟捌佰元整/); assert.match(xml, new RegExp(`${new Date().getFullYear()}年`));
  } finally { if (output) await fs.rm(outputFile(output.id, output.extension), { force: true }); await removeTemplate(item.id); }
});

test('renders numeric object values in procurement detail rows', async () => {
  const docx = new PizZip();
  docx.file('[Content_Types].xml', `<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Default Extension='xml' ContentType='application/xml'/><Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/></Types>`);
  docx.file('_rels/.rels', `<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships>`);
  docx.file('word/document.xml', `<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body><w:tbl><w:tr><w:tc><w:p><w:r><w:t>{#__合同明细_采购明细}{单价_含税___}{/__合同明细_采购明细}</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`);
  const item = await saveTemplate('采购合同含税.docx', docx.generate({ type: 'nodebuffer' })); let output;
  try { output = await renderTemplate(item.id, [{ '合同明细-采购明细': [{ '单价_含税___': '', '单价(含税)💻': { value: 240 }, '实收数量💻': 2 }] }]); const xml = new PizZip(await fs.readFile(outputFile(output.id, 'docx'))).file('word/document.xml').asText(); assert.match(xml, /240/); } finally { if (output) await fs.rm(outputFile(output.id, output.extension), { force: true }); await removeTemplate(item.id); }
});


test('renders XLSX template preview with worksheet grid and merged cells', async () => {
  const xlsx = new PizZip();
  xlsx.file('xl/workbook.xml', `<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="报价单" r:id="rId1"/></sheets></workbook>`);
  xlsx.file('xl/_rels/workbook.xml.rels', `<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`);
  xlsx.file('xl/styles.xml', `<styleSheet><fonts count="1"><font><b/><sz val="12"/></font></fonts><fills count="1"><fill><patternFill patternType="solid"><fgColor rgb="FFF2CC"/></patternFill></fill></fills><borders count="1"><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border></borders><cellXfs count="1"><xf fontId="0" fillId="0" borderId="0"><alignment horizontal="center"/></xf></cellXfs></styleSheet>`);
  xlsx.file('xl/worksheets/sheet1.xml', `<worksheet><cols><col min="1" max="1" width="18"/><col min="2" max="2" width="24"/></cols><sheetData><row r="1" ht="28"><c r="A1" s="0" t="inlineStr"><is><t>报价单标题</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>客户</t></is></c><c r="B2" t="inlineStr"><is><t>{客户}</t></is></c></row></sheetData><mergeCells><mergeCell ref="A1:B1"/></mergeCells></worksheet>`);
  const item = await saveTemplate('styled-preview.xlsx', xlsx.generate({ type: 'nodebuffer' }));
  try { const preview = await previewTemplate(item.id); assert.equal(preview.kind, 'html'); assert.match(preview.html, /报价单/); assert.match(preview.html, /colspan="2"/); assert.match(preview.html, /width:133\.2px/); assert.match(preview.html, /background:#F2CC|background:#fff2cc/i); assert.match(preview.html, />A<\/th>/); assert.match(preview.html, /row-number/); } finally { await removeTemplate(item.id); }
});
