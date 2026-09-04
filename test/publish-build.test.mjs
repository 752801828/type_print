import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import test from 'node:test';

test('contains publishable dist output with relative assets', async () => {
  const root = new URL('../', import.meta.url);
  const pkg = JSON.parse(await fs.readFile(new URL('package.json', root), 'utf8'));
  const html = await fs.readFile(new URL('dist/index.html', root), 'utf8');
  const app = await fs.readFile(new URL('dist/app.js', root), 'utf8');
  assert.equal(pkg.output, 'dist');
  assert.equal(pkg.scripts.build, 'node build.mjs');
  assert.match(html, /href="\.\/styles\.css"/);
  assert.match(html, /src="\.\/app\.js"/);
  assert.doesNotMatch(html, /(?:href|src)="\/feishu\//);
  assert.match(app, /https:\/\/gzwy\.online\/feishu/);
  await fs.access(new URL('dist/vendor/lark-base/index.mjs', root));
});
