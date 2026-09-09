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
    '\nthis.api = { state, linkedSchemaCache, legacyPlaceholderName, findTemplateField, readLinkedRows, recordScope, activeRecordFields, templateFieldDiagnostics, readSchema, readField, readRawField, variableFields, text };', context);
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

test('record loading keeps only the label and template fields', () => {
  const api = frontend();
  api.state.fields = [{ id: 'label', name: '采购合同' }, { id: 'needed', name: '供应商' }, ...Array.from({ length: 50 }, (_, index) => ({ id: `unused-${index}`, name: `无关字段${index}` }))];
  api.state.viewFields = [api.state.fields[0]];
  api.state.selectedTemplate = { fields: [{ marker: '', name: '__供应商' }] };
  assert.deepEqual([...api.activeRecordFields()].map(field => field.id), ['label', 'needed']);
});

test('record loading includes filename fields and formats Feishu dates', async () => {
  const api = frontend();
  api.state.fields = [{ id: 'label', name: '采购合同' }, { id: 'payee', name: '收款人' }, { id: 'date', name: '合同创建日期', type: 5, api: { getCellString: async () => '2026/09/08' } }];
  api.state.viewFields = [api.state.fields[0]];
  api.state.selectedTemplate = { fields: [], outputNamePattern: '{收款人}-{合同创建日期}' };
  assert.deepEqual([...api.activeRecordFields()].map(field => field.id), ['label', 'payee', 'date']);
  assert.equal(await api.readField(api.state.fields[2], 'record', { fields: { date: 1788796800000 } }, {}), '2026/09/08');
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
    ...['__开票品名', '单位', '单价_含税___', '总价_含税_'].map(name => ({ marker: '', name }))
  ] };
  api.linkedSchemaCache.set('details-table', { fields: [
    { id: 'name', name: '开票品名💻' }, { id: 'unit', name: '单位' },
    { id: 'price', name: '单价(含税)💻' }, { id: 'total', name: '总价(含税)💻' }
  ] });
  const diagnostics = [...api.templateFieldDiagnostics()];
  assert.equal(diagnostics.filter(item => !item.matched).length, 0);
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
