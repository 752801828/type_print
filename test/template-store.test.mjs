import assert from 'node:assert/strict';
import test from 'node:test';
import PizZip from 'pizzip';
import { extractFields } from '../lib/template-store.mjs';

test('extracts unique DOCX variables and loop markers', () => {
  const zip = new PizZip();
  zip.file('word/document.xml', '<w:document><w:r><w:t>{客户</w:t></w:r><w:r><w:t>名称}</w:t></w:r><w:t>{#items}{品名}{/items}</w:t><w:t>{客户名称}</w:t></w:document>');
  assert.deepEqual(extractFields(zip), [{ marker: '', name: '客户名称' }, { marker: '#', name: 'items' }, { marker: '', name: '品名' }]);
});
