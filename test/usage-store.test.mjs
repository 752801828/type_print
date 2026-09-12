import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs/promises';
import { recordUsage, usageFile } from '../lib/usage-store.mjs';

test('usage tracking hashes users and stores no record values', async () => {
  let previous;
  try { previous = await fs.readFile(usageFile); } catch {}
  try {
    await fs.rm(usageFile, { force: true });
    const event = await recordUsage({ userId: 'ou_private_user', baseId: 'base', baseName: '订单', tableId: 'table', tableName: '发票', format: 'XLSX', records: 3, secret: 'cell value' });
    assert.notEqual(event.user, 'ou_private_user');
    assert.equal(event.user.length, 16);
    assert.equal(event.records, 3);
    const stored = await fs.readFile(usageFile, 'utf8');
    assert.doesNotMatch(stored, /ou_private_user|cell value/);
    assert.match(stored, /"tableName":"发票"/);
  } finally {
    if (previous) await fs.writeFile(usageFile, previous); else await fs.rm(usageFile, { force: true });
  }
});
