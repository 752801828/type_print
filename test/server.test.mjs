import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs';
import path from 'node:path';

test('server is independent and fixed to the new port', () => {
  const source = fs.readFileSync(path.resolve('server.mjs'), 'utf8');
  assert.match(source, /process\.env\.PORT \|\| 4318/);
  assert.doesNotMatch(source, /template-print-demo|印序|4177/);
});
