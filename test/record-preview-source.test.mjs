import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs/promises';

test('record panel opens a full online-layout render instead of a field table', async () => {
  const [html, source] = await Promise.all([fs.readFile(new URL('../public/index.html', import.meta.url), 'utf8'), fs.readFile(new URL('../public/app.js', import.meta.url), 'utf8')]);
  assert.doesNotMatch(html, /id="fieldList"/);
  assert.match(html, /id="recordPreviewList"/);
  assert.match(source, /record-preview/);
  assert.match(source, /block\.table\.dynamicRows/);
  assert.match(html, /id="outputFormat"/);
  assert.match(source, /await openPreview\(result\.template\.id\)/);
  assert.doesNotMatch(source, /data-rename-id|renameTemplateItem/);
  assert.match(source, /data-delete-id/);
  assert.match(source, /method: 'DELETE'/);
  assert.doesNotMatch(html, /variableGuideDialog|dataSourceVariables|variableQuickRef|scanTemplate|batchDialog/);
  assert.match(source, /record-preview/);
  assert.match(source, /srcdoc = data\.html/);
  assert.match(html, /id="previewDocx"/);
  assert.match(html, /vendor\/docx-preview\.min\.js/);
  assert.match(source, /window\.docx\.renderAsync/);
  assert.match(source, /\[\['word','Word'\],\['xlsx','Excel（XLSX）'\],\['pdf','PDF'\]\]/);
  assert.match(source, /controller\.abort\(\), 45000/);
  assert.doesNotMatch(source, /docxPreviewHtml/);
});
