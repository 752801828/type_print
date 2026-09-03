import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs/promises';
import PizZip from 'pizzip';
import { saveTemplate, renderTemplate, removeTemplate, outputFile } from '../lib/template-store.mjs';

test('renders one DOCX and selected records as ZIP', async () => {
  const zip = new PizZip();
  zip.file('[Content_Types].xml', `<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Default Extension='xml' ContentType='application/xml'/><Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/></Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships>`);
  zip.file('word/document.xml', `<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body><w:p><w:r><w:t>{客户}</w:t></w:r></w:p><w:p><w:r><w:t>{#items}{$index+1}:{name}{/items}</w:t></w:r></w:p><w:p><w:r><w:t>{#状态=="通过"}条件成立{/}</w:t></w:r></w:p><w:p><w:r><w:t>{#!是否新人}旧客户{/是否新人}</w:t></w:r></w:p><w:p><w:r><w:t>{={{ }}=} {{客户}}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  const item = await saveTemplate('smoke.docx', zip.generate({ type: 'nodebuffer' }));
  assert.ok(item.fields.some(field => field.name === '状态' && field.marker === '?')); assert.ok(item.fields.every(field => field.name !== '=' && !field.name.includes('{')));
  const outputs = [];
  try {
    const one = await renderTemplate(item.id, [{ 客户: '甲', 状态: '通过', 是否新人: false, items: [{ name: '子项' }] }]);
    const many = await renderTemplate(item.id, [{ 客户: '甲', items: [{ name: '子项' }] }, { 客户: '乙', items: [] }]);
    assert.equal(one.extension, 'docx'); assert.equal(many.extension, 'zip');
    const renderedXml = new PizZip(await fs.readFile(outputFile(one.id, one.extension))).file('word/document.xml').asText();
    assert.match(renderedXml, /甲/); assert.match(renderedXml, /1:子项/); assert.match(renderedXml, /条件成立/); assert.match(renderedXml, /旧客户/); assert.doesNotMatch(renderedXml, /undefined/);
    outputs.push(one, many);
  } finally { for (const output of outputs) await fs.rm(outputFile(output.id, output.extension), { force: true }); await removeTemplate(item.id); }
});
