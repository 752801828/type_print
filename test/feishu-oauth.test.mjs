import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { createFeishuOAuth, uploadFeishuDocument } from '../lib/feishu-oauth.mjs';

const env = { FEIYE_FEISHU_APP_ID: 'test-app', FEIYE_FEISHU_APP_SECRET: 'test-secret', FEIYE_FEISHU_OAUTH_REDIRECT_URI: 'https://example.com/feishu/api/oauth/feishu/callback' };
const hash = value => crypto.createHash('sha256').update(value).digest('hex');
const reply = value => ({ ok: true, status: 200, json: async () => ({ code: 0, ...value }) });
async function fixture(fn) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'feiye-oauth-'));
  let clock = Date.now(), refreshes = 0;
  const calls = [];
  const request = async (url, options) => {
    calls.push({ url, options });
    if (url.endsWith('/user_info')) return reply({ data: { open_id: 'ou_real', tenant_key: 'tenant-a', name: '授权用户' } });
    if (url.endsWith('/oauth/token')) {
      const input = JSON.parse(options.body);
      if (input.grant_type === 'refresh_token') { refreshes++; await new Promise(resolve => setTimeout(resolve, 10)); }
      return reply({ access_token: `access-${refreshes}`, expires_in: 7200, refresh_token: `refresh-${refreshes}`, refresh_token_expires_in: 604800 });
    }
    if (url.endsWith('/medias/upload_all')) return reply({ data: { file_token: 'file-upload' } });
    if (url.endsWith('/import_tasks')) return reply({ data: { ticket: 'ticket-1' } });
    throw new Error(`Unexpected endpoint: ${url}`);
  };
  const options = { env, directory, request, now: () => clock };
  const oauth = createFeishuOAuth(options);
  const authorize = async () => {
    const verifier = crypto.randomBytes(32).toString('base64url');
    const attempt = oauth.start(hash(verifier));
    const authorization = oauth.authorize(attempt.state);
    const cookie = authorization.cookie.split(';')[0];
    await oauth.callback(attempt.state, 'valid-code', cookie);
    const result = await oauth.claim(attempt.state, verifier);
    return { ...result, attempt, verifier, cookie };
  };
  try { await fn({ oauth, options, directory, calls, authorize, advance: ms => clock += ms, refreshes: () => refreshes }); }
  finally { await fs.rm(directory, { recursive: true, force: true }); }
}

test('OAuth state is short-lived, one-use and bound to popup cookie plus private claim verifier', async () => fixture(async ({ oauth, authorize, advance }) => {
  const verifier = crypto.randomBytes(32).toString('base64url');
  const attempt = oauth.start(hash(verifier));
  const authorization = oauth.authorize(attempt.state);
  assert.match(authorization.cookie, /HttpOnly; Secure; SameSite=Lax/);
  assert.match(authorization.url, /^https:\/\/accounts\.feishu\.cn\//);
  assert.match(decodeURIComponent(authorization.url), /docs:document:import/);
  assert.match(decodeURIComponent(authorization.url), /docs:document\.media:upload/);
  assert.doesNotMatch(authorization.url, /test-secret/);
  await assert.rejects(oauth.callback(attempt.state, 'x', 'bad-cookie'), /校验失败/);
  await assert.rejects(oauth.claim(attempt.state, 'x'.repeat(43)), /校验失败/);
  advance(11 * 60000);
  await assert.rejects(oauth.claim(attempt.state, verifier), /过期/);
  const successful = await authorize();
  await assert.rejects(oauth.callback(successful.attempt.state, 'x', successful.cookie), /过期/);
  await assert.rejects(oauth.claim(successful.attempt.state, successful.verifier), /过期/);
}));

test('OAuth credentials are encrypted at rest, survive restart, and cannot be fetched with a user ID', async () => fixture(async ({ oauth, options, directory, authorize }) => {
  const login = await authorize();
  assert.equal(login.user.name, '授权用户');
  assert.equal(login.access_token, undefined);
  const file = await fs.readFile(path.join(directory, 'tokens.enc'));
  for (const value of ['access-0', 'refresh-0', 'ou_real', login.session]) assert.equal(file.includes(Buffer.from(value)), false);
  assert.equal(await createFeishuOAuth(options).token(login.session), 'access-0');
  await assert.rejects(oauth.identity('ou_real'), /授权/);
  await assert.rejects(oauth.identity(crypto.randomBytes(32).toString('base64url')), /过期/);
}));

test('concurrent refresh uses one request and stores the rotated token and actual expiry', async () => fixture(async ({ oauth, authorize, advance, refreshes, calls }) => {
  const login = await authorize();
  advance(7200 * 1000);
  assert.deepEqual(await Promise.all([oauth.token(login.session), oauth.token(login.session), oauth.token(login.session)]), ['access-1', 'access-1', 'access-1']);
  assert.equal(refreshes(), 1);
  advance(7200 * 1000); await oauth.token(login.session);
  assert.equal(JSON.parse(calls.filter(call => call.url.endsWith('/oauth/token')).at(-1).options.body).refresh_token, 'refresh-1');
  advance(8 * 86400000);
  await assert.rejects(oauth.token(login.session), /过期/);
}));

test('disconnect removes server credentials and invalidates sessions', async () => fixture(async ({ oauth, authorize }) => {
  const first = await authorize(), second = await authorize();
  await oauth.disconnect(first.session);
  await assert.rejects(oauth.token(second.session), /过期/);
}));

test('import uses user token, temporary media and empty mount key for personal root', async () => fixture(async ({ oauth, authorize, calls }) => {
  const { session } = await authorize();
  assert.equal(await uploadFeishuDocument(oauth, session, Buffer.from('test docx'), '发票'), 'ticket-1');
  const upload = calls.find(call => call.url.endsWith('/medias/upload_all'));
  assert.equal(upload.options.headers.authorization, 'Bearer access-0');
  assert.equal(upload.options.form, undefined);
  assert.equal(upload.options.body.get('parent_type'), 'ccm_import_open');
  assert.deepEqual(JSON.parse(upload.options.body.get('extra')), { obj_type: 'docx', file_extension: 'docx' });
  assert.deepEqual(JSON.parse(calls.at(-1).options.body).point, { mount_type: 1, mount_key: '' });
  await assert.rejects(uploadFeishuDocument(oauth, session, Buffer.alloc(20 * 1024 * 1024 + 1), 'big'), /20MB/);
}));

test('missing server configuration fails clearly and provider errors do not expose secrets', async () => {
  assert.throws(() => createFeishuOAuth({ env: {} }).start('a'.repeat(64)), /尚未配置/);
  const oauth = createFeishuOAuth({ env, request: async () => ({ ok: false, status: 400, json: async () => ({ error: 'invalid_grant', error_description: 'SECRET_ACCESS_TOKEN' }) }) });
  await assert.rejects(oauth.api('/test'), error => error.status === 401 && !error.message.includes('SECRET'));
});
