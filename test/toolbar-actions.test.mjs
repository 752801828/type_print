import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs/promises';
import vm from 'node:vm';

test('keeps template import separate from file generation', async () => {
  const html = await fs.readFile(new URL('../public/index.html', import.meta.url), 'utf8');
  const source = await fs.readFile(new URL('../public/app.js', import.meta.url), 'utf8');
  assert.match(html, /id="createTemplate">＋　导入模板/);
  assert.ok(html.indexOf('id="createTemplate"') < html.indexOf('id="categoryList"'));
  assert.doesNotMatch(html, />更多<|>导出|重命名|性能扫描|模板设置|查看数据源变量|变量指令速查|变量使用方法|进入批量模式/);
  assert.match(source, /\$\('createTemplate'\)\.onclick = \(\) => \$\('templateFile'\)\.click\(\)/);
  assert.match(source, /\$\('generateTop'\)\.onclick = \(\) => state\.selectedTemplate && state\.records\.length \? \$\('generate'\)\.click\(\) : toast/);
  assert.doesNotMatch(source, /renameTemplate|openVariableGuide|renderVariableGuide|openBatchDialog/);
});

test('import selects the template without opening a preview window', async () => {
  const source = await fs.readFile(new URL('../public/app.js', import.meta.url), 'utf8');
  const template = { id: 'imported', name: '测试模板.docx' };
  const state = { context: { tableId: 'table' }, templates: [] };
  let draws = 0, reads = 0;
  const context = vm.createContext({ state, appUrl: path => path,
    fetch: async () => ({ ok: true, json: async () => ({ template }) }),
    drawTemplates: () => draws++, readCurrentRecord: async () => reads++, toast: async () => {},
    openPreview: () => { throw new Error('Import must not open a preview'); }
  });
  vm.runInContext(source.slice(source.indexOf('async function uploadTemplate('), source.indexOf("$('createTemplate').onclick")), context);
  await context.uploadTemplate({ name: 'test.docx', arrayBuffer: async () => new ArrayBuffer(0) });
  assert.equal(state.selectedTemplate, template);
  assert.equal(state.templates[0], template);
  assert.equal(draws, 1);
  assert.equal(reads, 1);
});
