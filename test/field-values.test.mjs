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
    '\nthis.api = { state, legacyPlaceholderName, findTemplateField, readLinkedRows, recordScope };', context);
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
