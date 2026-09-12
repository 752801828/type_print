import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import vm from 'node:vm';
import test from 'node:test';
import PizZip from 'pizzip';
import { saveTemplate, renderTemplate, outputFile, removeTemplate } from '../lib/template-store.mjs';

const source = await fs.readFile(new URL('../public/app.js', import.meta.url), 'utf8');
function frontend() {
  const context = vm.createContext({ location: { pathname: '/feishu', hostname: 'localhost' } });
  vm.runInContext(source.slice(0, source.indexOf("$('createTemplate').onclick")) +
    '\nthis.api = { state, linkedSchemaCache, legacyPlaceholderName, findTemplateField, readLinkedRows, recordScope, activeRecordFields, attachmentCellValue, templateFieldDiagnostics, readSchema, readField, readRawField, variableFields, text, rememberCurrentUserName };', context);
  return context.api;
}

test('legacy encoding preserves every underscore and takes priority over similar fields', () => {
  const api = frontend();
  assert.equal(api.legacyPlaceholderName('单价(含税)💻'), '单价_含税___');
  assert.equal(api.legacyPlaceholderName('实收数量💻'), '实收数量__');
  assert.equal(api.legacyPlaceholderName('A / B+%'), 'A___B__');
  const fields = [{ name: '单价' }, { name: '单价(含税)' }, { name: '单价(含税)💻' }];
  assert.equal(api.findTemplateField(fields, '单价_含税___').name, '单价(含税)💻');
  assert.equal(api.findTemplateField([{ name: '数量' }, { name: '实收数量💻' }], '实收数量__').name, '实收数量💻');
  assert.equal(api.findTemplateField([{ id: 'short', name: '合同明细' }, { id: 'linked', name: '合同明细-采购明细' }], '__合同明细_采购明细').id, 'linked');
  assert.equal(api.findTemplateField([{ name: '合同明细' }], '__合同明细_采购明细'), undefined);
  assert.equal(api.findTemplateField([{ name: '数量' }], '实收数量__'), undefined);
  api.state.fields = [{ id: 'invoice', name: '🔴发票发货方信息' }]; api.state.selectedTemplate = { fields: [] };
  const scope = api.recordScope({ fields: { invoice: '广州公司' }, loops: {} });
  assert.equal(scope['🔴发票发货方信息'], '广州公司'); assert.equal(scope['发票发货方信息'], '广州公司'); assert.equal(scope['__发票发货方信息'], '广州公司');
  assert.equal(api.text({ type: 'text', text: '' }), '');
  assert.equal(api.text([{ type: 'text', text: '' }, { type: 'text', text: '有效内容' }]), '有效内容');
  assert.equal(api.text([{ type: 'text', text: 'Attn(联系人):Leon LI\n' }, { type: 'text', text: 'TEL(电话):+86 13660195555\n' }]), 'Attn(联系人):Leon LI\nTEL(电话):+86 13660195555\n');
  assert.equal(api.text({ type: 'text', text: '\u202A+59172776633\u202C' }), '+59172776633');
});

test('linked row to print payload uses real price and quantity, never derives a missing price', async () => {
  const api = frontend();
  api.state.selectedTemplate = { fields: [
    { name: '__合同明细_采购明细', marker: '#' },
    ...['单价_含税___', '实收数量__', '总价_含税_'].map(name => ({ name, marker: '' }))
  ] };
  const schema = [
    ['other', '单价'], ['price', '单价(含税)💻'],
    ['otherQty', '数量'], ['qty', '实收数量💻'], ['total', '总价(含税)']
  ];
  const records = {
    rec1: { other: 999, price: { value: 240 }, otherQty: 99, qty: 2, total: 480 },
    rec2: { other: 999, price: null, otherQty: 99, qty: 2, total: 480 },
    rec3: { other: 999, price: 0, otherQty: 99, qty: 0, total: 0 }
  };
  const table = {
    getFieldList: async () => schema.map(([id, name]) => ({
      id, getMeta: async () => ({ id, name }), getValue: async () => null, getCellString: async () => ''
    })),
    getRecordById: async id => ({ fields: records[id] }),
    getCellValue: async (id, recordId) => records[recordId][id]
  };
  const rows = await api.readLinkedRows({ tableId: 'tbl1', recordIds: Object.keys(records) },
    { base: { getTable: async () => table } });
  const payload = api.recordScope({ fields: {}, loops: { '合同明细-采购明细': rows } });
  const details = payload['__合同明细_采购明细'];
  assert.equal(details[0]['单价_含税___'], '240');
  assert.equal(details[0]['实收数量__'], '2');
  assert.equal(details[1]['单价_含税___'], '');
  assert.equal(details[2]['单价_含税___'], '0');
  assert.equal(details[2]['实收数量__'], '0');
});

test('bulk metadata and record values avoid per-cell bridge calls including blanks and zeros', async () => {
  const api = frontend(); let calls = 0;
  const fields = await api.readSchema({ getFieldMetaList: async () => { calls++; return Array.from({length:100},(_,id)=>({id:String(id),name:`字段${id}`,property:{}})); } });
  assert.equal(calls,1); assert.equal(fields.length,100);
  for (const value of [null, '', 0, false, [], {recordIds:['rec1'],tableId:'linked'}]) {
    const table = {getCellValue:()=>{throw new Error('must not re-read stored cells');}};
    const record = {fields:{'0':value}};
    assert.deepEqual(await api.readField(fields[0],'record',record,table),value ?? '');
    assert.deepEqual(await api.readRawField(fields[0],'record',record,table),value ?? '');
  }
});

test('display formatting reuses the field API while preserving numeric decimals', async () => {
  const api = frontend(); let fieldApiCalls = 0; let displayCalls = 0;
  const [field] = await api.readSchema({
    getFieldMetaList: async () => [{ id: 'price', name: '单价', type: 2 }],
    getFieldById: async () => { fieldApiCalls++; return { getCellString: async () => { displayCalls++; return '151.20'; } }; },
    getCellValue: async () => 151.199999999999
  });
  assert.equal(await api.readField(field, 'record-1', { fields: { price: 151.199999999999 } }, {}), '151.20');
  assert.equal(await api.readField(field, 'record-2', { fields: { price: 151.199999999999 } }, {}), '151.20');
  assert.equal(fieldApiCalls, 1);
  assert.equal(displayCalls, 2);
});

test('formatted numeric fields and their underscore aliases stay numeric in XLSX', async () => {
  const api = frontend();
  api.state.fields = [{ id: 'total', name: '总数量💻', type: 20 }, { id: 'code', name: '编号', type: 1 }];
  api.state.selectedTemplate = { fields: [{ name: '总数量__', marker: '' }, { name: '编号', marker: '' }] };
  const payload = api.recordScope({ fields: { total: '1,234', code: '001234' }, loops: {} });
  assert.equal(payload['总数量__'].__numeric, true); assert.equal(payload['总数量__'].value, 1234); assert.equal(payload['总数量__'].text, '1,234');
  assert.equal(payload['编号'], '001234');
  const zip = new PizZip(); zip.file('xl/workbook.xml', '<workbook/>'); zip.file('xl/worksheets/sheet1.xml', '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>{总数量__}</t></is></c><c r="B1" t="inlineStr"><is><t>{编号}</t></is></c></row></sheetData></worksheet>');
  const item = await saveTemplate('formatted-number.xlsx', zip.generate({ type: 'nodebuffer' })); let output;
  try { output = await renderTemplate(item.id, [payload]); const xml = new PizZip(await fs.readFile(outputFile(output.id, 'xlsx'))).file('xl/worksheets/sheet1.xml').asText(); assert.match(xml, /<c r="A1"><v>1234<\/v><\/c>/); assert.match(xml, /<c r="B1" t="inlineStr">[\s\S]*001234/); }
  finally { await removeTemplate(item.id); if (output) await fs.rm(outputFile(output.id, 'xlsx'), { force: true }); }
});

test('user name detection accepts nested Feishu person values', () => {
  const api = frontend(); api.state.userId = 'ou_current'; api.state.fields = [{ id: 'creator', name: '创建人', type: 1003 }];
  api.rememberCurrentUserName({ creator: { value: [{ open_id: 'ou_current', display_name: '李雷' }] } });
  assert.equal(api.state.userName, '李雷');
});

test('record loading keeps only the label and template fields', () => {
  const api = frontend();
  api.state.fields = [{ id: 'label', name: '采购合同' }, { id: 'needed', name: '供应商' }, ...Array.from({ length: 50 }, (_, index) => ({ id: `unused-${index}`, name: `无关字段${index}` }))];
  api.state.viewFields = [api.state.fields[0]];
  api.state.selectedTemplate = { fields: [{ marker: '', name: '__供应商' }] };
  assert.deepEqual([...api.activeRecordFields()].map(field => field.id), ['label', 'needed']);
  api.state.fields.push({ id: 'attachment', name: '生成文件', type: 17 }); api.state.selectedTemplate.autoUpload = true; api.state.selectedTemplate.outputFieldId = 'attachment';
  assert.deepEqual([...api.activeRecordFields()].map(field => field.id), ['label', 'needed', 'attachment']);
});

test('attachment upload appends by default and replaces only when requested', () => {
  const api = frontend(); const old = { name: '旧文件.pdf', token: 'old' }; const file = { name: '新文件.pdf', size: 12, type: 'application/pdf', lastModified: 123 };
  assert.deepEqual([...api.attachmentCellValue([old], file, 'new')].map(item => item.token), ['old', 'new']);
  assert.deepEqual([...api.attachmentCellValue([old], file, 'new', true)].map(item => item.token), ['new']);
});

test('record loading includes filename fields and formats Feishu dates', async () => {
  const api = frontend();
  api.state.fields = [{ id: 'label', name: '采购合同' }, { id: 'payee', name: '收款人' }, { id: 'date', name: '合同创建日期', type: 5, api: { getCellString: async () => '2026/09/08' } }, { id: 'formula', name: '收货地址', type: 19, api: { getCellString: async () => 'Room 401, Guangzhou' } }];
  api.state.viewFields = [api.state.fields[0]];
  api.state.selectedTemplate = { fields: [], outputNamePattern: '{收款人}-{合同创建日期}-{收货地址}' };
  assert.deepEqual([...api.activeRecordFields()].map(field => field.id), ['label', 'payee', 'date', 'formula']);
  assert.equal(await api.readField(api.state.fields[2], 'record', { fields: { date: 1788796800000 } }, {}), '2026/09/08');
  assert.equal(await api.readField(api.state.fields[3], 'record', { fields: { formula: null } }, {}), 'Room 401, Guangzhou');
  assert.equal(await api.readField({ id: 'money', name: '金额', type: 99003 }, 'record', { fields: { money: '￥1,234.50' } }, {}), '1,234.50');
  assert.equal(await api.readField({ id: 'price', name: '单价', type: 2, api: { getCellString: async () => '151.20' } }, 'record', { fields: { price: 151.199999999999 } }, {}), '151.20');
  assert.equal(await api.readField({ id: 'money', name: '金额', type: 99003, api: { getCellString: async () => '￥151.20' } }, 'record', { fields: { money: 151.199999999999 } }, {}), '151.20');
});

test('record loading includes the relation used by a linked filename field', () => {
  const api = frontend();
  api.state.fields = [{ id: 'label', name: '采购合同' }, { id: 'relation', name: '合同明细', relationTableId: 'linked' }];
  api.state.viewFields = [api.state.fields[0]];
  api.state.selectedTemplate = { fields: [], outputNamePattern: '{合同明细.供应商}' };
  assert.deepEqual([...api.activeRecordFields()].map(field => field.id), ['label', 'relation']);
  api.linkedSchemaCache.set('linked', { tableName: '采购明细表', fields: [{ id: 'supplier', name: '供应商' }] });
  assert.ok(api.variableFields().some(field => field.name === '合同明细.供应商' && field.label === '{合同明细.供应商}'));
});

test('field diagnostics accept matching fields from the linked detail table', () => {
  const api = frontend();
  api.state.fields = [{ id: 'details', name: '合同明细-采购明细' }];
  api.state.selectedTemplate = { fields: [
    { marker: '#', name: '__合同明细_采购明细' },
    ...['__开票品名', '单位', '单位', '单价_含税___', '总价_含税_'].map(name => ({ marker: '', name }))
  ] };
  api.linkedSchemaCache.set('details-table', { fields: [
    { id: 'name', name: '开票品名💻' }, { id: 'unit', name: '单位' },
    { id: 'price', name: '单价(含税)💻' }, { id: 'total', name: '总价(含税)💻' }
  ] });
  const diagnostics = [...api.templateFieldDiagnostics()];
  assert.equal(diagnostics.filter(item => !item.matched).length, 0);
  assert.equal(diagnostics.filter(item => item.name === '单位').length, 1);
  assert.equal(diagnostics.find(item => item.name === '单价_含税___').fieldId, 'price');
});

test('DOCX renderer matches encoded names inside loops and keeps absent price blank', async () => {
  const zip = new PizZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>{#__合同明细_采购明细}[{单价_含税___}|{实收数量__}|{A___B__}]{/__合同明细_采购明细}</w:t></w:r></w:p></w:body></w:document>');
  const template = await saveTemplate('encoded-fields.docx', zip.generate({ type: 'nodebuffer' }));
  let output;
  try {
    output = await renderTemplate(template.id, [{ '合同明细-采购明细': [
      { '单价(含税)': 999, '单价(含税)💻': 240, 数量: 99, '实收数量💻': 2, 'A / B+%': '匹配成功' },
      { '单价_含税___': '', '实收数量💻': 2, '总价(含税)': 480, 'A / B+%': '缺失' },
      { '单价(含税)💻': 0, '实收数量💻': 0, 'A / B+%': '零值' }
    ] }]);
    const xml = new PizZip(await fs.readFile(outputFile(output.id, 'docx'))).file('word/document.xml').asText();
    assert.ok(xml.includes('[240|2|匹配成功]'));
    assert.ok(xml.includes('[|2|缺失]'));
    assert.ok(xml.includes('[0|0|零值]'));
  } finally {
    if (output) await fs.rm(outputFile(output.id, output.extension), { force: true });
    await removeTemplate(template.id);
  }
});
