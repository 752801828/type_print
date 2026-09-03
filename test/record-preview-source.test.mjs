import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs/promises';

test('record panel opens a full online-layout render instead of a field table', async () => {
  const [html, source] = await Promise.all([fs.readFile(new URL('../public/index.html', import.meta.url), 'utf8'), fs.readFile(new URL('../public/app.js', import.meta.url), 'utf8')]);
  assert.doesNotMatch(html, /id="fieldList"/);
  assert.match(html, /id="recordPreviewList"/);
  assert.match(source, /renderLayoutPreview\(data\.layout, recordScope\(record\)\)/);
  assert.match(source, /block\.table\.dynamicRows/);
  assert.match(html, /id="outputFormat"/);
  assert.match(source, /await openPreview\(result\.template\.id\)/);
  assert.match(source, /data-rename-id/);
  assert.match(source, /data-delete-id/);
  assert.match(source, /method: 'DELETE'/);
  assert.match(html, /id="variableGuideDialog"/);
  assert.match(source, /openVariableGuide\('fields'\)/);
  assert.match(source, /\{#!是否新人\}/);
  assert.match(source, /\{=\{\{ \}\}=\}/);
});
