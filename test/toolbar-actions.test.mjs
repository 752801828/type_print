import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs/promises';

test('keeps template import separate from file generation', async () => {
  const html = await fs.readFile(new URL('../public/index.html', import.meta.url), 'utf8');
  const source = await fs.readFile(new URL('../public/app.js', import.meta.url), 'utf8');
  assert.match(html, /id="createTemplate">＋　导入模板/);
  assert.doesNotMatch(html, />更多<|>导出|重命名|性能扫描|模板设置|查看数据源变量|变量指令速查|变量使用方法|进入批量模式/);
  assert.match(source, /\$\('createTemplate'\)\.onclick = \(\) => \$\('templateFile'\)\.click\(\)/);
  assert.match(source, /\$\('generateTop'\)\.onclick = \(\) => state\.selectedTemplate && state\.records\.length \? \$\('generate'\)\.click\(\) : toast/);
  assert.doesNotMatch(source, /renameTemplate|openVariableGuide|renderVariableGuide|openBatchDialog/);
});
