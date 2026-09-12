import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs/promises';
import PizZip from 'pizzip';
import { extractFields, saveTemplate, updateTemplate, duplicateTemplate, removeTemplate, templateFile, listTemplates } from '../lib/template-store.mjs';

test('extracts unique DOCX variables and loop markers', () => {
  const zip = new PizZip();
  zip.file('word/document.xml', '<w:document><w:r><w:t>{客户</w:t></w:r><w:r><w:t>名称}</w:t></w:r><w:t>{#items}{品名}{/items}</w:t><w:t>{客户名称}</w:t></w:document>');
  assert.deepEqual(extractFields(zip), [{ marker: '', name: '客户名称' }, { marker: '#', name: 'items' }, { marker: '', name: '品名' }]);
});

test('renames and duplicates a bound template with its settings', async () => {
  const zip = new PizZip(); zip.file('word/document.xml', '<w:document><w:t>{客户}</w:t></w:document>');
  const original = await saveTemplate('原模板.docx', zip.generate({ type: 'nodebuffer' }), { baseId: 'base', tableId: 'table', outputNamePattern: '{客户}-合同', autoUpload: true, outputFieldId: 'attachment', outputReplace: true }); let copy;
  try {
    const renamed = await updateTemplate(original.id, { name: '新名称.pdf', pinned: true }); assert.equal(renamed.name, '新名称.docx'); assert.equal(renamed.outputNamePattern, '{客户}-合同'); assert.equal(renamed.outputFieldId, 'attachment'); assert.equal(renamed.pinned, true); assert.equal((await listTemplates({ baseId: 'base', tableId: 'table' }))[0].id, original.id);
    copy = await duplicateTemplate(original.id); assert.equal(copy.name, '新名称-副本.docx'); assert.equal(copy.tableId, 'table'); assert.equal(copy.outputNamePattern, '{客户}-合同'); assert.equal(copy.autoUpload, true); assert.equal(copy.outputFieldId, 'attachment'); assert.equal(copy.outputReplace, true); assert.equal(copy.pinned, false); assert.deepEqual(await fs.readFile(templateFile(copy.id, '.docx')), await fs.readFile(templateFile(original.id, '.docx')));
  } finally { if (copy) await removeTemplate(copy.id); await removeTemplate(original.id); }
});

test('imports the online template attachment target', async () => {
  const source = Buffer.from(JSON.stringify({ content: { settings: { exportFile: { targetFieldId: 'fldAttachment', isReplace: false } }, document: { pages: [{ rows: [] }] } } }));
  const item = await saveTemplate('online-upload.txt', source);
  try { assert.equal(item.autoUpload, true); assert.equal(item.outputFieldId, 'fldAttachment'); assert.equal(item.outputReplace, false); }
  finally { await removeTemplate(item.id); }
});
