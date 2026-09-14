import test from 'node:test';
import assert from 'node:assert/strict';
import PizZip from 'pizzip';
import { saveTemplate, removeTemplate, renderFeishuDocx, mergeTemplateDocx } from '../lib/template-store.mjs';

test('online templates fill and merge records with filename rules for Feishu import', async () => {
  const template = { content: { document: { pages: [{ rows: [{ columns: [{ width: 100, blocks: [{ type: 1, content: [{ children: [{ type: 'variable', name: ['客户'] }] }] }] }] }] }] } } };
  const item = await saveTemplate('云文档.layout', Buffer.from(JSON.stringify(template)), { outputNamePattern: '{客户}-订单' });
  try {
    const result = await renderFeishuDocx(item.id, [{ 客户: '客户甲' }, { 客户: '客户乙' }]);
    const zip = new PizZip(result.bytes), xml = zip.file('word/document.xml').asText();
    assert.match(xml, /客户甲/); assert.match(xml, /客户乙/); assert.match(xml, /w:br w:type="page"/);
    assert.equal(result.name, '客户甲-订单-2条合并'); assert.equal((xml.match(/<w:sectPr>/g) || []).length, 1);
    await assert.rejects(renderFeishuDocx(item.id, []), /1–100/);
  } finally { await removeTemplate(item.id); }
});

test('merger refuses mismatched header or relationship packages rather than dropping content', () => {
  const doc = header => { const zip = new PizZip(); zip.file('word/document.xml', '<w:document><w:body><w:p/><w:sectPr/></w:body></w:document>'); zip.file('word/header1.xml', header); return zip.generate({ type: 'nodebuffer' }); };
  assert.throws(() => mergeTemplateDocx([doc('customer A'), doc('customer B')]), /页眉/);
  assert.equal(new PizZip(mergeTemplateDocx([doc('fixed'), doc('fixed')])).file('word/header1.xml').asText(), 'fixed');
});
