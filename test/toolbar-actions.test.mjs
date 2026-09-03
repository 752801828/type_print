import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs/promises';

test('keeps template import separate from file generation', async () => {
  const html = await fs.readFile(new URL('../public/index.html', import.meta.url), 'utf8');
  const source = await fs.readFile(new URL('../public/app.js', import.meta.url), 'utf8');
  assert.match(html, /id="createTemplate">＋　导入模板/);
  assert.match(html, /id="addTemplate"[^>]*>＋<\/button>/);
  assert.doesNotMatch(html, />更多<|>导出</);
  assert.match(source, /\$\('createTemplate'\)\.onclick = \$\('addTemplate'\)\.onclick = \(\) => \$\('templateFile'\)\.click\(\)/);
  assert.match(source, /\$\('generateTop'\)\.onclick = \(\) => state\.selectedTemplate && state\.records\.length \? \$\('generate'\)\.click\(\) : toast/);
});
