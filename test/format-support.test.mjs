import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs/promises';
import PizZip from 'pizzip';
import { saveTemplate, renderTemplate, previewTemplate, removeTemplate, outputFile, listTemplates } from '../lib/template-store.mjs';

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

test('XLSX accepts original field names and legacy underscore placeholders together', async () => {
  const xlsx = new PizZip(); xlsx.file('xl/workbook.xml', '<workbook/>'); xlsx.file('xl/worksheets/sheet1.xml', '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>{🔵发票发货方信息}|{发票发货方信息}|{__发票发货方信息}</t></is></c></row></sheetData></worksheet>');
  const item = await saveTemplate('raw-field.xlsx', xlsx.generate({ type: 'nodebuffer' })); let output;
  try { output = await renderTemplate(item.id, [{ '🔴发票发货方信息': '广州公司' }]); const sheet = new PizZip(await fs.readFile(outputFile(output.id, 'xlsx'))).file('xl/worksheets/sheet1.xml').asText(); assert.match(sheet, /广州公司\|广州公司\|广州公司/); }
  finally { if (output) await fs.rm(outputFile(output.id, output.extension), { force: true }); await removeTemplate(item.id); }
});

test('rich text fragments keep their source formatting without inserted separators', async () => {
  const xlsx = new PizZip(); xlsx.file('xl/workbook.xml', '<workbook/>'); xlsx.file('xl/worksheets/sheet1.xml', '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>{联系信息}</t></is></c></row></sheetData></worksheet>');
  const item = await saveTemplate('rich-text.xlsx', xlsx.generate({ type: 'nodebuffer' })); let output;
  try { output = await renderTemplate(item.id, [{ 联系信息: [{ type: 'text', text: 'Attn(联系人):Leon LI\n' }, { type: 'text', text: 'TEL(电话):\u202A+86 13660195555\u202C\n' }, { type: 'text', text: '' }] }]); const sheet = new PizZip(await fs.readFile(outputFile(output.id, output.extension))).file('xl/worksheets/sheet1.xml').asText(); assert.match(sheet, /Attn\(联系人\):Leon LI\nTEL\(电话\):\+86 13660195555\n/); assert.doesNotMatch(sheet, /、/); assert.doesNotMatch(sheet, /202A|202C/); }
  finally { if (output) await fs.rm(outputFile(output.id, output.extension), { force: true }); await removeTemplate(item.id); }
});

test('imports Feishu online layout export and preserves structured preview', async () => {
  const source = Buffer.from(JSON.stringify({ name: '在线模板', content: JSON.stringify({ settings: { fontSize: 9 }, pageSetting: { paddingTop: 5 }, document: { pages: [{ rows: [{ columns: [{ width: 100, blocks: [{ type: 1, marginTop: -50, content: [{ type: 'paragraph', children: [{ text: '客户：' }, { type: 'variable', name: ['🔵客户名称'] }] }] }] }] }] }] } }) })).toString('base64');
  const item = await saveTemplate('online.txt', Buffer.from(source), { baseId: 'layout-base', tableId: 'layout-table' });
  const outputs = [];
  try {
    const preview = await previewTemplate(item.id); assert.equal(preview.kind, 'layout'); assert.equal(preview.layout.document.pages.length, 1); assert.deepEqual(preview.fields, [{ marker: '', name: '客户名称' }]);
    const record = { '🔵客户名称': '测试客户', '🔵合同明细': [{ '🔴SKU': 'SKU-彩色', '🔴单价': '12.30' }] };
    const recordPreview = await (await import('../lib/template-store.mjs')).previewRecord(item.id, record); assert.equal(recordPreview.kind, 'html'); assert.match(recordPreview.html, /测试客户/);
    const pdf = await renderTemplate(item.id, [record]); outputs.push(pdf); assert.equal(pdf.extension, 'pdf'); assert.match((await fs.readFile(outputFile(pdf.id, 'pdf'))).subarray(0, 4).toString(), /%PDF/);
    const word = await renderTemplate(item.id, [record], 'word'); outputs.push(word); assert.equal(word.extension, 'docx'); const wordXml = new PizZip(await fs.readFile(outputFile(word.id, 'docx'))).file('word/document.xml').asText(); assert.match(wordXml, /测试客户/); assert.match(wordXml, /w:sectPr/); assert.match(wordXml, /w:top="0"/); assert.match(wordXml, /w:sz w:val="18"/);
    const excel = await renderTemplate(item.id, [record], 'xlsx'); outputs.push(excel); assert.equal(excel.extension, 'xlsx'); const excelXml = new PizZip(await fs.readFile(outputFile(excel.id, 'xlsx'))).file('xl/worksheets/sheet1.xml').asText(); assert.match(excelXml, /测试客户/); assert.match(excelXml, /<sheetData>/);
  } finally { for (const output of outputs) await fs.rm(outputFile(output.id, output.extension), { force: true }); await removeTemplate(item.id); }
});

test('XLSX numeric placeholders remain numeric cells for spreadsheet selection totals', async () => {
  const xlsx = new PizZip(); xlsx.file('xl/workbook.xml', '<workbook/>'); xlsx.file('xl/worksheets/sheet1.xml', '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>{数量}</t></is></c><c r="B1" t="inlineStr"><is><t>数量：{数量}</t></is></c></row></sheetData></worksheet>');
  const item = await saveTemplate('numeric-cells.xlsx', xlsx.generate({ type: 'nodebuffer' })); let output;
  try { output = await renderTemplate(item.id, [{ 数量: 24 }]); const sheet = new PizZip(await fs.readFile(outputFile(output.id, 'xlsx'))).file('xl/worksheets/sheet1.xml').asText(); assert.match(sheet, /<c r="A1"[^>]*><v>24<\/v><\/c>/); assert.match(sheet, /<c r="B1"[^>]*t="inlineStr"[^>]*>.*数量：24/); }
  finally { if (output) await fs.rm(outputFile(output.id, output.extension), { force: true }); await removeTemplate(item.id); }
});

test('online layout XLSX export keeps the layout structure and PDF uses a PDFKit font path', async () => {
  const source = Buffer.from(JSON.stringify({ content: { pageSetting: { width: 210, height: 297, paddingLeft: 5, paddingRight: 5 }, settings: { fontSize: 9 }, document: { pages: [{ rows: [{ columns: [{ width: 100, blocks: [{ type: 4, table: { columns: [{ width: 1 }, { width: 2 }], merges: [{ rowIndex: 0, colIndex: 0, rowSpan: 1, colSpan: 2 }], rows: [{ cells: [{ background: '#90c3fc', content: [{ type: 'paragraph', align: 'center', children: [{ text: '标题', bold: true, fontSize: '12pt' }] }] }, { content: [{ type: 'paragraph', children: [{ type: 'variable', name: ['客户'] }] }] }] }] } }] }] }] }] } } })).toString('base64');
  const item = await saveTemplate('layout-export.txt', Buffer.from(source)); let output;
  try { output = await renderTemplate(item.id, [{ 客户: '客户A' }], 'xlsx'); const generated = new PizZip(await fs.readFile(outputFile(output.id, 'xlsx'))); const xml = generated.file('xl/worksheets/sheet1.xml').asText(); const styles = generated.file('xl/styles.xml').asText(); assert.match(xml, /标题/); assert.match(xml, /<mergeCell ref="A1:B1"/); assert.match(xml, /showGridLines="0"/); assert.match(xml, /fitToWidth="1"/); assert.match(xml, /<cols>/); assert.match(styles, /FF90C3FC/); assert.match(styles, /<b\/>/); assert.match(styles, /<sz val="12"\/>/); }
  finally { if (output) await fs.rm(outputFile(output.id, output.extension), { force: true }); await removeTemplate(item.id); }
});

test('online layout XLSX keeps narrow columns, wrapped text visible and merged borders complete', async () => {
  const cell = children => ({ content: [{ type: 'paragraph', children }] });
  const source = Buffer.from(JSON.stringify({ content: { settings: { fontSize: 9 }, pageSetting: { width: 210, paddingLeft: 13.5, paddingRight: 13.5 }, document: { pages: [{ rows: [{ columns: [{ width: 100, blocks: [{ type: 4, table: {
    columns: [{ width: 5.71 }, { width: 25.36 }, { width: 6.28 }, { width: 13.57 }, { width: 10.5 }, { width: 13.97 }, { width: 24.61 }],
    rows: [
      { cells: [cell([{ type: 'variable', name: ['长文本'] }]), ...Array.from({ length: 6 }, () => cell([]))] },
      { cells: [cell([{ text: '完整边框' }]), ...Array.from({ length: 6 }, () => cell([]))] }
    ],
    merges: [{ rowIndex: 1, colIndex: 0, rowSpan: 1, colSpan: 7 }]
  } }] }] }] }] } } })).toString('base64');
  const item = await saveTemplate('layout-wrapping.txt', Buffer.from(source)); let output;
  try {
    output = await renderTemplate(item.id, [{ 长文本: '这是需要自动换行显示的长文本，不能被下一行遮挡。This text must remain visible.' }], 'xlsx');
    const generated = new PizZip(await fs.readFile(outputFile(output.id, 'xlsx'))); const xml = generated.file('xl/worksheets/sheet1.xml').asText();
    const firstWidth = Number(xml.match(/<col min="1" max="1" width="([\d.]+)"/)?.[1]); const rowHeight = Number(xml.match(/<row r="1" ht="([\d.]+)"/)?.[1]);
    assert.ok(firstWidth < 8, `首列宽度应保留模板比例，实际为 ${firstWidth}`);
    assert.ok(rowHeight > 30, `长文本行高应自动增大，实际为 ${rowHeight}`);
    assert.match(xml, /<mergeCell ref="A2:G2"/);
    for (const column of ['A', 'B', 'C', 'D', 'E', 'F', 'G']) assert.match(xml, new RegExp(`<c r="${column}2"[^>]* s="[1-9]`));
  } finally { if (output) await fs.rm(outputFile(output.id, output.extension), { force: true }); await removeTemplate(item.id); }
});

test('online layout XLSX spans full-width paragraphs across a nearby table grid', async () => {
  const cell = text => ({ content: [{ type: 'paragraph', children: [{ text }] }] });
  const paragraph = text => ({ type: 1, content: [{ type: 'paragraph', children: [{ text }] }] });
  const table = { columns: Array.from({ length: 7 }, (_, index) => ({ width: index === 1 ? 20 : 10 })), rows: [{ cells: Array.from({ length: 7 }, (_, index) => cell(`列${index + 1}`)) }] };
  const source = Buffer.from(JSON.stringify({ content: { settings: { fontSize: 9 }, pageSetting: { width: 210, paddingLeft: 6, paddingRight: 6 }, document: { pages: [{ rows: [
    { columns: [{ width: 100, blocks: [paragraph('整页宽度合同条款，不能只写入第一列')] }] },
    { columns: [{ width: 50, blocks: [paragraph('甲方信息')] }, { width: 50, blocks: [paragraph('乙方信息包含较长的公司名称、纳税人识别号、联系地址、联系电话、开户行和银行账号，最终列宽变窄后也必须完整显示')] }] },
    { columns: [{ width: 100, blocks: [{ type: 4, table }] }] }
  ] }] } } })).toString('base64');
  const item = await saveTemplate('mixed-layout.txt', Buffer.from(source)); let output;
  try {
    output = await renderTemplate(item.id, [{}], 'xlsx');
    const xml = new PizZip(await fs.readFile(outputFile(output.id, 'xlsx'))).file('xl/worksheets/sheet1.xml').asText();
    assert.match(xml, /<mergeCell ref="A1:G1"/);
    assert.match(xml, /<mergeCell ref="A2:D2"/);
    assert.match(xml, /<mergeCell ref="E2:G2"/);
    const widths = [...xml.matchAll(/<col min="\d+" max="\d+" width="([\d.]+)"/g)].map(match => Number(match[1]));
    const secondRowHeight = Number(xml.match(/<row r="2" ht="([\d.]+)"/)?.[1]);
    assert.equal(widths.length, 7);
    assert.ok(Math.max(...widths) < 30, `单列不应被整页段落撑宽：${widths.join(', ')}`);
    assert.ok(secondRowHeight > 30, `行高必须按最终列宽计算：${secondRowHeight}`);
  } finally { if (output) await fs.rm(outputFile(output.id, output.extension), { force: true }); await removeTemplate(item.id); }
});

test('online layout Word export preserves columns, tables and all field values in a real DOCX', async () => {
  const cell = children => ({ content: [{ type: 'paragraph', children }] });
  const source = Buffer.from(JSON.stringify({ content: { settings: { fontSize: 9 }, pageSetting: { width: 210, height: 297, paddingTop: 5, paddingLeft: 12, paddingRight: 12 }, document: { pages: [{ rows: [{ columns: [{ width: 50, blocks: [{ type: 1, marginTop: -50, content: [{ type: 'paragraph', children: [{ text: '左侧：' }, { type: 'variable', name: ['客户'] }] }] }] }, { width: 50, blocks: [{ type: 4, table: { columns: [{ width: 1 }, { width: 2 }], rows: [{ cells: [cell([{ text: '编号' }]), cell([{ type: 'variable', name: ['明细', 'SKU'] }])] }], dynamicRows: [{ rowIndex: 0, dataSource: { rootPath: ['明细'] } }] } }] }] }] }] } } })).toString('base64');
  const item = await saveTemplate('layout-word.txt', Buffer.from(source)); let output;
  try {
    output = await renderTemplate(item.id, [{ 客户: '广州客户', 明细: [{ SKU: 'SKU-001' }, { SKU: 'SKU-002' }] }], 'word');
    assert.equal(output.extension, 'docx');
    const zip = new PizZip(await fs.readFile(outputFile(output.id, 'docx'))); const xml = zip.file('word/document.xml').asText();
    assert.match(xml, /广州客户/); assert.match(xml, /SKU-001/); assert.match(xml, /SKU-002/); assert.ok((xml.match(/<w:tbl>/g) || []).length >= 2); assert.match(xml, /w:pgSz/); assert.match(xml, /w:sz w:val="18"/);
  } finally { if (output) await fs.rm(outputFile(output.id, output.extension), { force: true }); await removeTemplate(item.id); }
});

test('PDF font discovery accepts TTC fonts and scans common font directories', async () => {
  const source = await fs.readFile(new URL('../lib/template-store.mjs', import.meta.url), 'utf8');
  assert.match(source, /\.ttc/);
  assert.match(source, /readdir\(root, \{ withFileTypes: true, recursive: true \}\)/);
  assert.match(source, /member\?\.postscriptName/);
  assert.match(source, /doc\.font\(font\.path, font\.family\)/);
});

test('paginates repeated PDF rows independently of horizontal merged cells', async () => {
  const cell = children => ({ content: [{ type: 'paragraph', children }] });
  const table = {
    columns: [{ width: 1 }, { width: 1 }],
    rows: [
      { cells: [cell([{ text: 'Merged heading' }]), cell([])] },
      { cells: [cell([{ type: 'variable', name: ['明细', 'SKU'] }]), cell([{ text: '12' }])] }
    ],
    merges: [{ rowIndex: 0, colIndex: 0, rowSpan: 1, colSpan: 2 }],
    dynamicRows: [{ rowIndex: 1, dataSource: { rootPath: ['明细'] } }]
  };
  const source = { name: 'PDF pagination', content: { document: { pages: [{ rows: [{ columns: [{ width: 100, blocks: [{ type: 4, table }] }] }] }] } } };
  const item = await saveTemplate('pagination.layout', Buffer.from(JSON.stringify(source)));
  let output;
  try {
    output = await renderTemplate(item.id, [{ 明细: Array.from({ length: 60 }, (_, i) => ({ SKU: `ROW-${i}` })) }], 'pdf');
    const pdf = (await fs.readFile(outputFile(output.id, 'pdf'))).toString('latin1');
    assert.equal((pdf.match(/\/Type \/Page\b/g) || []).length, 2);
  } finally {
    if (output) await fs.rm(outputFile(output.id, 'pdf'), { force: true });
    await removeTemplate(item.id);
  }
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

test('extends conditional formatting and moves merged footer rows after XLSX loops', async () => {
  const zip = new PizZip();
  zip.file('xl/workbook.xml', '<workbook/>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1" t="str"><v>{#明细}{值}{/明细}</v></c></row><row r="2"><c r="A2" t="str"><v>合计</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A2:B2"/></mergeCells><conditionalFormatting sqref="A1:B1"><cfRule type="expression"><formula>A1&gt;0</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="A2:B2"><cfRule type="expression"><formula>A2&lt;&gt;&quot;&quot;</formula></cfRule></conditionalFormatting></worksheet>');
  const item = await saveTemplate('conditional-loop.xlsx', zip.generate({ type: 'nodebuffer' })); let output;
  try {
    output = await renderTemplate(item.id, [{ 明细: [{ 值: 1 }, { 值: 2 }, { 值: 3 }] }]);
    const sheet = new PizZip(await fs.readFile(outputFile(output.id, 'xlsx'))).file('xl/worksheets/sheet1.xml').asText();
    assert.match(sheet, /sqref="A1:B3"/); assert.match(sheet, /sqref="A4:B4"/); assert.match(sheet, /mergeCell ref="A4:B4"/); assert.match(sheet, /dimension ref="A1:B4"/); assert.match(sheet, /<c r="A1"[^>]*><v>1<\/v><\/c>/);
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

test('keeps self-closing XLSX cells separate from following template cells', async () => {
  const zip = new PizZip();
  zip.file('xl/workbook.xml', '<workbook/>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet><sheetData><row r="1"><c r="A1" s="1"/><c r="B1" s="1" t="inlineStr"><is><t>{编号}</t></is></c></row></sheetData></worksheet>');
  const item = await saveTemplate('self-closing.xlsx', zip.generate({ type: 'nodebuffer' })); let output;
  try {
    output = await renderTemplate(item.id, [{ 编号: 'A-100' }]);
    const xml = new PizZip(await fs.readFile(outputFile(output.id, 'xlsx'))).file('xl/worksheets/sheet1.xml').asText();
    assert.match(xml, /<c r="A1" s="1"\/>/); assert.match(xml, /<c r="B1" s="1" t="inlineStr">/); assert.doesNotMatch(xml, /<c[^>]*\/\s+t=/);
  } finally { if (output) await fs.rm(outputFile(output.id, output.extension), { force: true }); await removeTemplate(item.id); }
});

test('marks XLSX formulas for a full recalculation when opened', async () => {
  const zip = new PizZip();
  zip.file('xl/workbook.xml', '<workbook><calcPr calcId="191029" concurrentCalc="0"/></workbook>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>{仓库}</t></is></c><c r="B1"><f>VLOOKUP(A1,地址库!A:B,2,0)</f><v>#N/A</v></c></row></sheetData></worksheet>');
  const item = await saveTemplate('formula.xlsx', zip.generate({ type: 'nodebuffer' })); let output;
  try {
    output = await renderTemplate(item.id, [{ 仓库: 'GYR2' }]); const generated = new PizZip(await fs.readFile(outputFile(output.id, 'xlsx')));
    const workbook = generated.file('xl/workbook.xml').asText(); const sheet = generated.file('xl/worksheets/sheet1.xml').asText();
    assert.match(workbook, /calcMode="auto"/); assert.match(workbook, /fullCalcOnLoad="1"/); assert.match(workbook, /forceFullCalc="1"/); assert.match(workbook, /calcCompleted="0"/);
    assert.match(sheet, /<f>VLOOKUP\(A1,地址库!A:B,2,0\)<\/f>/); assert.match(sheet, />GYR2<\/t>/);
  } finally { if (output) await fs.rm(outputFile(output.id, output.extension), { force: true }); await removeTemplate(item.id); }
});

test('fills cached VLOOKUP results so formula cells map in previews and downloads', async () => {
  const zip = new PizZip();
  zip.file('xl/workbook.xml', '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="发票" sheetId="1" r:id="rId1"/><sheet name="地址表" sheetId="2" r:id="rId2"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>{仓库}</t></is></c><c r="B1"><f>VLOOKUP(A1,地址表!D:I,6,0)</f><v></v></c></row></sheetData></worksheet>');
  zip.file('xl/worksheets/sheet2.xml', '<worksheet><sheetData><row r="1"><c r="D1" t="inlineStr"><is><t>GYR2</t></is></c><c r="I1" t="inlineStr"><is><t>Room 401, Guangzhou</t></is></c></row></sheetData></worksheet>');
  const item = await saveTemplate('formula-vlookup.xlsx', zip.generate({ type: 'nodebuffer' })); let output;
  try {
    output = await renderTemplate(item.id, [{ 仓库: 'GYR2' }]);
    const generated = new PizZip(await fs.readFile(outputFile(output.id, 'xlsx')));
    const sheet = generated.file('xl/worksheets/sheet1.xml').asText();
    assert.match(sheet, /t="str"[^>]*>.*<v>Room 401, Guangzhou<\/v>/);
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

test('returns a filled DOCX for faithful record preview and HTML for XLSX', async () => {
  const docx = new PizZip();
  docx.file('[Content_Types].xml', `<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Default Extension='xml' ContentType='application/xml'/><Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/></Types>`);
  docx.file('_rels/.rels', `<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships>`);
  docx.file('word/document.xml', `<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body><w:p><w:r><w:t>客户：{客户}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  const docxItem = await saveTemplate('preview.docx', docx.generate({ type: 'nodebuffer' }));
  const xlsx = new PizZip(); xlsx.file('xl/workbook.xml', '<workbook/>'); xlsx.file('xl/worksheets/sheet1.xml', `<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>{编号}</t></is></c></row></sheetData></worksheet>`);
  const xlsxItem = await saveTemplate('preview.xlsx', xlsx.generate({ type: 'nodebuffer' }));
  try { const docxPreview = await (await import('../lib/template-store.mjs')).previewRecord(docxItem.id, { 客户: '张三' }); assert.equal(docxPreview.kind, 'docx'); assert.match(new PizZip(docxPreview.bytes).file('word/document.xml').asText(), /张三/); const xlsxPreview = await (await import('../lib/template-store.mjs')).previewRecord(xlsxItem.id, { 编号: 'A-100' }); assert.equal(xlsxPreview.kind, 'html'); assert.match(xlsxPreview.html, /A-100/); } finally { await removeTemplate(docxItem.id); await removeTemplate(xlsxItem.id); }
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
    try { const preview = await previewTemplate(item.id); assert.equal(preview.kind, 'html'); assert.match(preview.html, /报价单/); assert.match(preview.html, /colspan="2"/); assert.match(preview.html, /width:133\.2px/); assert.match(preview.html, /background:#F2CC|background:#fff2cc/i); assert.match(preview.html, />A<\/th>/); assert.match(preview.html, /row-number/); assert.match(preview.html, /max-height:calc\(100vh - 90px\);overflow:scroll/); } finally { await removeTemplate(item.id); }
});
