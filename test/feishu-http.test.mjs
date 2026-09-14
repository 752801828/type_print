import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import net from 'node:net';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { saveTemplate, removeTemplate } from '../lib/template-store.mjs';

test('HTTP OAuth to merged import flow, duplicate requests and cross-user isolation', { timeout: 30000 }, async () => {
  const socket = net.createServer(); socket.listen(0, '127.0.0.1'); await once(socket, 'listening'); const port = socket.address().port; await new Promise(resolve => socket.close(resolve));
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'feiye-http-'));
  const mock = new URL('./fixtures/mock-feishu.mjs', import.meta.url).href;
  const child = spawn(process.execPath, ['--import', mock, 'server.mjs'], { cwd: process.cwd(), env: { ...process.env, HOST: '127.0.0.1', PORT: String(port), FEIYE_FEISHU_APP_ID: 'test', FEIYE_FEISHU_APP_SECRET: 'test', FEIYE_FEISHU_OAUTH_REDIRECT_URI: 'https://example.com/feishu/api/oauth/feishu/callback', FEIYE_FEISHU_OAUTH_DATA_DIR: directory }, stdio: ['ignore', 'pipe', 'pipe'] });
  const base = `http://127.0.0.1:${port}/feishu`;
  const api = (route, data, session = '') => fetch(`${base}${route}`, { method: data === undefined ? 'GET' : 'POST', headers: { 'content-type': 'application/json', ...(session ? { authorization: `Bearer ${session}` } : {}) }, body: data === undefined ? undefined : JSON.stringify(data) });
  let item;
  try {
    await once(child.stdout, 'data');
    assert.equal((await api('/api/health')).status, 200);
    assert.equal((await api('/api/feishu-docs', {})).status, 401);
    const authorize = async code => {
      const verifier = crypto.randomBytes(32).toString('base64url');
      const challenge = crypto.createHash('sha256').update(verifier).digest('hex');
      const attempt = await (await api('/api/oauth/feishu/start', { challenge })).json();
      const redirect = await fetch(`${base}/api/oauth/feishu/authorize?state=${attempt.state}`, { redirect: 'manual' });
      assert.equal(redirect.status, 302);
      const cookie = redirect.headers.get('set-cookie').split(';')[0];
      assert.equal((await fetch(`${base}/api/oauth/feishu/callback?state=${attempt.state}&code=${code}`)).status, 400);
      const callback = await fetch(`${base}/api/oauth/feishu/callback?state=${attempt.state}&code=${code}`, { headers: { cookie } });
      assert.equal(callback.status, 200); assert.match(await callback.text(), /授权完成/);
      const login = await (await api('/api/oauth/feishu/status', { state: attempt.state, verifier })).json();
      assert.ok(login.session); return login.session;
    };
    const userA = await authorize('user-a'), userB = await authorize('user-b');
    const layout = { content: { document: { pages: [{ rows: [{ columns: [{ width: 100, blocks: [{ type: 1, content: [{ children: [{ type: 'variable', name: ['客户'] }] }] }] }] }] }] } } };
    item = await saveTemplate('OAuth集成测试.layout', Buffer.from(JSON.stringify(layout)), { baseId: 'test-base', tableId: 'test-table' });
    const input = { templateId: item.id, records: [{ 客户: '甲' }, { 客户: '乙' }], baseId: 'test-base', tableId: 'test-table', requestId: crypto.randomUUID() };
    const response = await api('/api/feishu-docs', input, userA); assert.equal(response.status, 202);
    const job = await response.json();
    assert.equal((await api(`/api/feishu-docs/${job.id}`, undefined, userB)).status, 403);
    const again = await (await api('/api/feishu-docs', input, userA)).json(); assert.equal(again.id, job.id);
    let result;
    for (let i = 0; i < 50; i++) {
      result = await (await api(`/api/feishu-docs/${job.id}`, undefined, userA)).json();
      if (result.stage === 'complete') break;
      await new Promise(resolve => setTimeout(resolve, 50));
    }
    assert.equal(result.stage, 'complete'); assert.equal(result.url, 'https://example.feishu.cn/docx/document'); assert.equal(result.ownerName, '用户a');
    const file = await fs.readFile(path.join(directory, 'jobs', `${job.id}.json`), 'utf8');
    assert.doesNotMatch(file, /test-refresh|test-user-a|"客户"/);
    assert.equal((await api('/api/oauth/feishu/disconnect', {}, userA)).status, 200);
    assert.equal((await api(`/api/feishu-docs/${job.id}`, undefined, userA)).status, 401);
  } finally {
    child.kill(); if (child.exitCode === null) await once(child, 'exit');
    if (item) await removeTemplate(item.id);
    await fs.rm(directory, { recursive: true, force: true });
  }
});
